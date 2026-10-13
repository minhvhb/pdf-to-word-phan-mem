import streamlit as st

# ==========================================
# CẤU HÌNH TRANG CHỦ ĐẠO (BẮT BUỘC PHẢI Ở DÒNG ĐẦU TIÊN)
# ==========================================
st.set_page_config(page_title="Hệ thống Công cụ Văn phòng tại NMTLBT Craven A", page_icon="⚙️", layout="centered", initial_sidebar_state="expanded")

import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth
import os
import re
import time
import zipfile
import pandas as pd
from io import BytesIO
import json
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from PIL import Image
from pypdf import PdfReader, PdfWriter, Transformation
import difflib
import fitz

# Ẩn nút gập thanh bên mặc định của Streamlit
st.markdown(
    """
    <style>
        [data-testid="collapsedControl"] { display: none; }
        [data-testid="stSidebarCollapseButton"] { display: none; }
    </style>
    """,
    unsafe_allow_html=True
)

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

def clear_file():
    st.session_state.uploader_key += 1

# ==========================================
# 0. HẠ TẦNG GỌI GEMINI: NHIỀU API KEY XOAY VÒNG + TỰ ĐỔI MODEL FLASH
# ==========================================
import threading
import hashlib
from collections import OrderedDict
from datetime import datetime, timedelta

# Hạn mức tính RIÊNG cho từng model -> dùng cả chuỗi Flash/Flash-Lite để cộng dồn lượt/ngày của 1 key.
# Model nào key không có hạn mức hoặc không tồn tại sẽ tự bị bỏ qua. Ghi đè trong Secrets bằng GEMINI_MODELS.
GEMINI_DEFAULT_MODELS = ["gemini-3.7-flash", "gemini-3.6-flash", "gemini-3.5-flash", "gemini-3.5-flash-lite",
                         "gemini-3.1-flash-lite", "gemini-2.5-flash", "gemini-2.5-flash-lite"]


def _read_gemini_keys():
    """Đọc danh sách key từ Secrets: GEMINI_API_KEYS (danh sách hoặc chuỗi) hoặc GEMINI_API_KEY (1 key)."""
    raw = None
    try:
        raw = st.secrets.get("GEMINI_API_KEYS", None)
        if raw is None:
            raw = st.secrets.get("GEMINI_API_KEY", None)
    except Exception:
        raw = None
    if raw is None:
        raw = os.environ.get("GEMINI_API_KEYS") or os.environ.get("GEMINI_API_KEY")
    if raw is None:
        return []
    items = re.split(r"[,\s;]+", raw) if isinstance(raw, str) else [str(x) for x in list(raw)]
    keys = []
    for k in items:
        k = k.strip().strip("\"'")
        if k and k not in keys:
            keys.append(k)
    return keys


def _next_quota_reset_ts():
    """Hạn mức theo ngày của Gemini API reset lúc 0h giờ Thái Bình Dương (Pacific)."""
    try:
        from zoneinfo import ZoneInfo
        n = datetime.now(ZoneInfo("America/Los_Angeles"))
        return (n + timedelta(days=1)).replace(hour=0, minute=0, second=5, microsecond=0).timestamp()
    except Exception:
        return time.time() + 3 * 3600


def _retry_delay(msg):
    """Đọc thời gian chờ Google gợi ý (vd 'Please retry in 20.5s' / retryDelay: '20s'). None nếu không có."""
    m = re.search(r"retry(?:Delay)?['\"]?\s*(?:in|:)\s*['\"]?([\d.]+)\s*s", msg, re.I)
    try:
        return float(m.group(1)) if m else None
    except Exception:
        return None


def _payload_hash(contents, config):
    h = hashlib.sha256()
    items = contents if isinstance(contents, (list, tuple)) else [contents]
    for it in items:
        if isinstance(it, str):
            h.update(it.encode("utf-8", "ignore"))
        else:
            inl = getattr(it, "inline_data", None)
            if inl is not None and getattr(inl, "data", None) is not None:
                h.update(str(getattr(inl, "mime_type", "")).encode())
                h.update(inl.data)
            else:
                h.update(repr(it).encode("utf-8", "ignore"))
        h.update(b"\x00")
    h.update(repr(config).encode("utf-8", "ignore"))
    return h.hexdigest()


def _fmt_ts(ts):
    try:
        from zoneinfo import ZoneInfo
        return datetime.fromtimestamp(ts, ZoneInfo("Asia/Ho_Chi_Minh")).strftime("%H:%M %d/%m")
    except Exception:
        return datetime.fromtimestamp(ts).strftime("%H:%M %d/%m")


class GeminiKeyPool:
    """
    Dùng chung cho mọi người dùng/phiên của ứng dụng. Với mỗi lần gọi:
      for model (ưu tiên cao -> thấp): for key (lần lượt): thử; lỗi hạn mức -> khoá tạm (key, model) rồi sang key kế.
    Hết key ở model này mới hạ xuống model flash kế tiếp. Lỗi tạm thời ngắn (<=30s) thì chờ rồi thử lại.
    """

    def __init__(self, keys, models, mode="sequential", max_concurrent=3, cache_ttl=6 * 3600, cache_size=40):
        self.keys, self.models, self.mode = list(keys), list(models), mode
        self.lock = threading.Lock()
        self.blocks = {}                      # (key|None, model|None) -> (đến_lúc, lý_do)
        self.ok_calls = {k: 0 for k in self.keys}
        self.rr = 0
        self._clients = {}
        self.last_err = {}                    # key -> mô tả lỗi gần nhất
        self.model_calls = {}                 # model -> số lượt thành công
        self.cache = OrderedDict()            # hash -> (thời điểm, response): tránh gọi lại cùng file + cùng yêu cầu
        self.cache_ttl, self.cache_size, self.cache_hits = cache_ttl, cache_size, 0
        self.sem = threading.Semaphore(max(1, int(max_concurrent)))   # giới hạn số cuộc gọi đồng thời (chống vọt RPM)

    # ---- client
    def _client(self, key):
        c = self._clients.get(key)
        if c is None:
            c = genai.Client(api_key=key)      # key AI Studio (AIza... hoặc AQ....) đều gọi Gemini API thường
            self._clients[key] = c
        return c

    # ---- khoá tạm
    def _block(self, key, model, secs, reason):
        with self.lock:
            self.blocks[(key, model)] = (time.time() + secs, reason)

    def _blocked_until(self, key, model, now):
        with self.lock:
            ends = [self.blocks[p][0] for p in ((key, None), (key, model), (None, model)) if p in self.blocks]
        ends = [e for e in ends if e > now]
        return max(ends) if ends else 0

    def _ordered_keys(self):
        if self.mode == "round_robin" and len(self.keys) > 1:
            with self.lock:
                s = self.rr % len(self.keys)
                self.rr += 1
            return self.keys[s:] + self.keys[:s]
        return list(self.keys)

    def _classify(self, key, model, e):
        msg = str(e)
        low = msg.lower()
        code = getattr(e, "code", None)
        if not isinstance(code, int):
            m = re.match(r"\s*(\d{3})\b", msg)
            code = int(m.group(1)) if m else None
        if any(x in low for x in ("api_key_invalid", "api key not valid", "permission_denied", "leaked", "suspended",
                                  "api key expired")) or code in (401, 403):
            short = re.sub(r"\s+", " ", msg)[:170]
            self._block(key, None, 600, "Bị từ chối truy cập (401/403): " + short)
            return "retry"
        if code == 429 or "resource_exhausted" in low:
            if "limit: 0" in low:
                self._block(key, model, 6 * 3600, "Model không có hạn mức (miễn phí) với key này")
            elif "perday" in low.replace(" ", "") or "per day" in low:
                self._block(key, model, max(60, _next_quota_reset_ts() - time.time()), "Hết hạn mức ngày")
            else:
                d = _retry_delay(msg)
                self._block(key, model, min(max(d + 1, 2), 65) if d else 65, "Quá giới hạn/phút")
            return "retry"
        if code in (500, 502, 503, 504) or any(x in low for x in ("unavailable", "overloaded", "high demand",
                                                                  "deadline")):
            self._block(None, model, 20, "Model quá tải")
            return "retry"
        if code == 404 or "not_found" in low or "is not found" in low:
            self._block(key, model, 3600, "Model không tồn tại với key này")
            return "retry"
        return "fatal"                         # lỗi do dữ liệu (file sai định dạng, quá lớn...) -> báo ngay

    def generate(self, contents, config=None):
        ck = _payload_hash(contents, config)
        now0 = time.time()
        with self.lock:
            hit = self.cache.get(ck)
            if hit and now0 - hit[0] < self.cache_ttl:
                self.cache.move_to_end(ck)
                self.cache_hits += 1
                return hit[1]
        resp = self._generate_uncached(contents, config)
        with self.lock:
            self.cache[ck] = (time.time(), resp)
            while len(self.cache) > self.cache_size:
                self.cache.popitem(last=False)
        return resp

    def _generate_uncached(self, contents, config=None):
        last_err = None
        for rnd in range(3):
            soonest = None
            for model in self.models:
                for key in self._ordered_keys():
                    until = self._blocked_until(key, model, time.time())
                    if until:
                        w = until - time.time()
                        soonest = w if soonest is None else min(soonest, w)
                        continue
                    try:
                        kw = {"model": model, "contents": contents}
                        if config is not None:
                            kw["config"] = config
                        with self.sem:
                            resp = self._client(key).models.generate_content(**kw)
                        with self.lock:
                            self.ok_calls[key] = self.ok_calls.get(key, 0) + 1
                            self.model_calls[model] = self.model_calls.get(model, 0) + 1
                        return resp
                    except Exception as e:
                        last_err = e
                        self.last_err[key] = re.sub(r"\s+", " ", str(e))[:200]
                        if self._classify(key, model, e) == "fatal":
                            raise
            soonest = self._min_wait()
            if soonest is not None and soonest <= 30 and rnd < 2:
                time.sleep(soonest + 0.5)       # chờ khoá ngắn (giới hạn/phút) rồi thử lại
                continue
            break
        raise RuntimeError(self._final_message(last_err))

    def _final_message(self, last_err):
        now = time.time()
        auth_keys, lines = 0, []
        with self.lock:
            for k in self.keys:
                b = self.blocks.get((k, None))
                if b and b[0] > now and b[1].startswith("Bị từ chối"):
                    auth_keys += 1
                lines.append(f"• Key …{k[-4:]}: {self.last_err.get(k, 'chưa có lỗi')[:150]}")
        if auth_keys == len(self.keys):
            head = "Không có API key nào được Google chấp nhận (bị từ chối truy cập 401/403) — đây KHÔNG phải lỗi hết hạn mức."
        else:
            nxt = self.next_available()
            head = "Các API key/model Gemini đang hết hạn mức hoặc quá tải." + (
                f" Dự kiến có lại khoảng {_fmt_ts(nxt)}." if nxt else "")
        return head + " Chi tiết từng key:\n" + "\n".join(lines)

    def reset(self):
        with self.lock:
            self.blocks.clear()
            self.last_err.clear()

    def _min_wait(self):
        """Số giây ngắn nhất để có ít nhất 1 tổ hợp (key, model) dùng lại được."""
        now = time.time()
        best = None
        for model in self.models:
            for key in self.keys:
                u = self._blocked_until(key, model, now)
                w = max(0.0, u - now) if u else 0.0
                best = w if best is None else min(best, w)
        return best

    def next_available(self):
        now = time.time()
        best = None
        for model in self.models:
            for key in self.keys:
                u = self._blocked_until(key, model, now)
                if not u:
                    return None
                best = u if best is None else min(best, u)
        return best

    def snapshot(self):
        now = time.time()
        with self.lock:
            items = [(p, v) for p, v in self.blocks.items() if v[0] > now]
            calls = dict(self.ok_calls)
        out = []
        for k in self.keys:
            bl = [(p[1] or "toàn bộ", v[0], v[1]) for p, v in items if p[0] == k]
            out.append({"key": "…" + k[-4:], "ok": calls.get(k, 0), "blocks": bl, "err": self.last_err.get(k, "")})
        mod = [(p[1], v[0], v[1]) for p, v in items if p[0] is None]
        return out, mod


@st.cache_resource(show_spinner=False)
def _build_gemini_pool(keys, models, mode, max_conc):
    return GeminiKeyPool(keys, models, mode, max_concurrent=max_conc)


def get_gemini_pool():
    keys = _read_gemini_keys()
    if not keys:
        return None
    try:
        models = list(st.secrets.get("GEMINI_MODELS", GEMINI_DEFAULT_MODELS))
    except Exception:
        models = list(GEMINI_DEFAULT_MODELS)
    try:
        mode = str(st.secrets.get("GEMINI_KEY_MODE", "sequential")).lower()
    except Exception:
        mode = "sequential"
    try:
        max_conc = int(st.secrets.get("GEMINI_MAX_CONCURRENT", 3))
    except Exception:
        max_conc = 3
    return _build_gemini_pool(tuple(keys), tuple(models), mode, max_conc)


def require_gemini_key():
    """Dùng ở đầu mỗi công cụ cần AI: dừng trang nếu chưa cấu hình key, trả về key đầu tiên (chỉ để khởi tạo client)."""
    keys = _read_gemini_keys()
    if not keys:
        st.error("⚠️ Hệ thống chưa được cấu hình API Key (GEMINI_API_KEYS). Vui lòng liên hệ Quản trị viên!")
        st.stop()
    return keys[0]


def call_gemini_smart(client, contents_payload, config=None):
    """
    Giữ nguyên tên/cách gọi cũ nhưng nay đi qua kho nhiều key: tự xoay key khi hết hạn mức,
    tự hạ xuống model flash kế tiếp, chờ khi chỉ bị giới hạn theo phút. (Tham số client được bỏ qua.)
    """
    pool = get_gemini_pool()
    if pool is None:
        raise RuntimeError("Chưa cấu hình GEMINI_API_KEYS trong Secrets.")
    return pool.generate(contents_payload, config)


def render_gemini_status():
    """Bảng theo dõi key cho Quản trị viên (hiện ở thanh bên)."""
    pool = get_gemini_pool()
    with st.sidebar.expander("🔑 Trạng thái API Key (Admin)"):
        if pool is None:
            st.error("Chưa có key nào trong Secrets.")
            return
        st.caption(f"{len(pool.keys)} key • Model ưu tiên: {' → '.join(pool.models)} • Chế độ: {pool.mode}")
        keys, models_blocked = pool.snapshot()
        for k in keys:
            st.markdown(f"**{k['key']}** — {'🟢 sẵn sàng' if not k['blocks'] else '🟠 có khoá tạm'} • đã dùng {k['ok']} lượt")
            for scope, until, why in k["blocks"]:
                st.caption(f"• {scope}: {why}, đến {_fmt_ts(until)}")
            if k["err"]:
                st.caption(f"Lỗi gần nhất: {k['err']}")
        for m, until, why in models_blocked:
            st.caption(f"Model {m}: {why}, đến {_fmt_ts(until)}")
        if st.button("🔄 Xoá khoá tạm & thử lại", key="gemini_pool_reset"):
            pool.reset()
            st.rerun()
        if pool.model_calls:
            st.caption("Lượt thành công theo model: " + ", ".join(f"{m}: {n}" for m, n in pool.model_calls.items()))
        st.caption(f"Tiết kiệm nhờ bộ nhớ đệm (cùng file + cùng yêu cầu): {pool.cache_hits} lượt")
        st.caption("Số liệu chỉ tính từ lúc ứng dụng khởi động lại gần nhất.")

# ==========================================
# 1. ĐỌC DỮ LIỆU TÀI KHOẢN VÀ ĐĂNG NHẬP
# ==========================================
with open('config.yaml', 'r', encoding='utf-8') as file:
    config = yaml.load(file, Loader=SafeLoader)

stauth.Hasher.hash_passwords(config['credentials'])

authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days']
)

# ==========================================
# 2. GIAO DIỆN TRANG ĐĂNG NHẬP
# ==========================================
if st.session_state.get("authentication_status") != True:
    st.markdown("""
        <style>
            .stApp {
                background-image: url("https://images.unsplash.com/photo-1477414348463-c0eb7f1359b6?q=80&w=2070&auto=format&fit=crop");
                background-size: cover;
                background-position: center;
                background-attachment: fixed;
            }
            [data-testid="stForm"] {
                background-color: rgba(255, 255, 255, 0.95) !important;
                padding: 40px !important;
                border-radius: 15px !important;
                box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3) !important;
                width: 700px !important; 
                max-width: 95vw !important;
                margin: 10vh auto 0 auto !important;
                border: 1px solid rgba(255, 255, 255, 0.5) !important;
            }
            [data-testid="stForm"] h2 {
                text-align: center !important;
                color: #003366 !important;
                font-weight: 900 !important;
                font-size: 24px !important; 
                white-space: nowrap !important;
                margin-bottom: 25px !important;
                padding-bottom: 0px !important;
                width: 100% !important;
            }
            [data-testid="stForm"] [data-testid="stVerticalBlock"] > div {
                width: 100% !important;
                max-width: 100% !important;
            }
            [data-testid="stForm"] .stTextInput {
                width: 100% !important;
                min-width: 100% !important;
            }
            [data-testid="stForm"] div[data-baseweb="input"] {
                background-color: #ffffff !important;
                border: 2px solid #000000 !important;
                border-radius: 6px !important;
                width: 100% !important;
            }
            [data-testid="stForm"] div[data-baseweb="input"]:focus-within {
                border-color: #003366 !important; 
            }
            [data-testid="stForm"] input {
                background-color: transparent !important;
                color: #000 !important;
                font-size: 15px !important;
            }
            [data-testid="stForm"] .stTextInput label p {
                font-size: 14px !important;
                font-weight: bold !important;
                color: #333 !important;
            }
            [data-testid="stForm"] div[data-testid="stFormSubmitButton"] {
                width: 100% !important;
                max-width: 100% !important;
                display: block !important;
            }
            [data-testid="stForm"] div[data-testid="stFormSubmitButton"] button {
                width: 100% !important;
                min-width: 100% !important;
                background-color: #003366 !important;
                color: white !important;
                font-size: 16px !important;
                letter-spacing: 1px !important;
                border-radius: 8px !important;
                padding: 12px 0 !important;
                border: none !important;
                margin-top: 15px !important;
                display: block !important;
                transition: all 0.3s ease-in-out !important;
            }
            [data-testid="stForm"] div[data-testid="stFormSubmitButton"] button p {
                font-weight: 900 !important;
                margin: 0 !important;
            }
            [data-testid="stForm"] div[data-testid="stFormSubmitButton"] button:hover {
                background-color: #001f3f !important;
                box-shadow: 0px 4px 15px rgba(0,0,0,0.4) !important;
            }
            div[data-testid="stVerticalBlock"] > div:has(div[data-testid="stMarkdownContainer"]) {
                display: flex;
                justify-content: center;
                margin-top: 15px;
            }
        </style>
    """, unsafe_allow_html=True)

authenticator.login(
    location='main',
    fields={
        'Form name': 'CÔNG CỤ VĂN PHÒNG NMTLBT CRAVEN A',
        'Username': 'Tên đăng nhập',
        'Password': 'Mật khẩu',
        'Login': 'ĐĂNG NHẬP'
    }
)

if st.session_state["authentication_status"] == False:
    st.error('Tên đăng nhập hoặc mật khẩu không đúng!')
    st.stop()
elif st.session_state["authentication_status"] == None:
    st.warning('Vui lòng đăng nhập để sử dụng công cụ')
    st.stop()

# ==========================================
# 3. GIAO DIỆN THANH BÊN (SIDEBAR) TỐI ƯU
# ==========================================
st.markdown("""
    <style>
        [data-testid="stSidebarUserContent"] { 
            padding-top: 1rem !important; 
            padding-bottom: 0rem !important; 
            padding-left: 1rem !important; 
            padding-right: 1rem !important;
        }
        [data-testid="stSidebarUserContent"] div[role="radiogroup"] > label { 
            margin-bottom: -6px !important; 
        }
        [data-testid="stSidebarUserContent"] hr { 
            margin-top: 5px; 
            margin-bottom: 5px; 
        }
        .custom-alert {
            background-color: #f8d7da;
            color: #842029;
            padding: 10px;
            border-radius: 5px;
            font-size: 13.5px;
            line-height: 1.4;
            border: 1px solid #f5c2c7;
        }
    </style>
""", unsafe_allow_html=True)

col1, col2, col3 = st.sidebar.columns([1, 2, 1])
with col2:
    authenticator.logout('Đăng xuất', 'main')

st.sidebar.markdown(f"<h5 style='text-align: center; margin-top: 5px; margin-bottom: 0px;'>Chào mừng {st.session_state['name']}!</h5>", unsafe_allow_html=True)
st.sidebar.markdown("---")

st.sidebar.markdown("### 📌 Menu Công Cụ")
# ==========================================
# APP 12: PDF SANG WORD - OFFLINE (GIỮ NGUYÊN BỐ CỤC)
# ==========================================
# Engine A: pdf2docx (thư viện chuyên dụng, văn bản trôi theo dòng, dễ chỉnh sửa)
# Engine B: dựng lại từ lớp chữ của PDF (tự viết): đúng khổ giấy, lề, cỡ chữ, khoảng cách dòng, bảng, cột,
#           ảnh/logo/con dấu đặt tại toạ độ tuyệt đối. Không dùng AI, không cần mạng.
import io as _io
import re as _re
import math as _math
import unicodedata as _ud
import statistics as _stat
from collections import Counter as _Counter

P2W_DESC = 0.215          # tỉ lệ phần "đuôi" dưới đường cơ sở của dòng (Word/LibreOffice giữ nguyên khi dòng cố định)

_P2W_FONT_MAP = [
    ("arialnarrow", "Arial Narrow"), ("timesnewroman", "Times New Roman"), ("liberationserif", "Times New Roman"),
    ("dejavuserif", "Times New Roman"), ("times", "Times New Roman"), ("liberationsans", "Arial"),
    ("dejavusans", "Arial"), ("arial", "Arial"), ("helvetica", "Arial"), ("liberationmono", "Courier New"),
    ("couriernew", "Courier New"), ("courier", "Courier New"), ("calibri", "Calibri"), ("cambria", "Cambria"),
    ("tahoma", "Tahoma"), ("verdana", "Verdana"), ("segoeui", "Segoe UI"), ("georgia", "Georgia"),
    ("garamond", "Garamond"), ("palatino", "Palatino Linotype"), ("bookantiqua", "Book Antiqua"),
    ("trebuchet", "Trebuchet MS"), ("centurygothic", "Century Gothic"), ("consolas", "Consolas"),
    ("candara", "Candara"), ("roboto", "Roboto"), ("opensans", "Open Sans"), ("symbol", "Symbol"),
    ("wingdings", "Wingdings"),
]


# ------------------------------------------------------------------ tiện ích
def _p2w_font(fontname):
    n = _re.sub(r"^[A-Z]{6}\+", "", fontname or "")
    low = n.lower()
    bold = bool(_re.search(r"bold|black|heavy|semibold|demibold", low))
    ital = bool(_re.search(r"italic|oblique|ital\b|-it\b", low))
    key = _re.sub(r"[^a-z]", "", _re.split(r"[-,]", low)[0])
    for k, v in _P2W_FONT_MAP:
        if k in key:
            return v, bold, ital
    base = _re.split(r"[-,]", n)[0]
    base = _re.sub(r"(PSMT|MT|PS)$", "", base)
    base = _re.sub(r"(?<=[a-z])(?=[A-Z])", " ", base).strip()
    return (base or "Times New Roman"), bold, ital


def _p2w_color(c):
    """màu chữ của pdfplumber -> 'RRGGBB' hoặc None (đen/không rõ)."""
    if c is None or isinstance(c, str):
        return None
    try:
        t = tuple(float(x) for x in (c if isinstance(c, (list, tuple)) else (c,)))
    except Exception:
        return None
    if len(t) == 1:
        r = g = b = t[0]
    elif len(t) == 3:
        r, g, b = t
    elif len(t) == 4:
        cc, m, y, k = t
        r, g, b = (1 - cc) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k)
    else:
        return None
    if max(r, g, b) < 0.16:
        return None
    return "%02X%02X%02X" % tuple(int(round(max(0, min(1, v)) * 255)) for v in (r, g, b))


def _p2w_is_white(c):
    if c is None:
        return True
    hx = _p2w_color(c) if not isinstance(c, str) else None
    if hx is None:
        try:
            t = tuple(float(x) for x in (c if isinstance(c, (list, tuple)) else (c,)))
            return max(t) < 0.16 and False
        except Exception:
            return True
    return min(int(hx[i:i + 2], 16) for i in (0, 2, 4)) >= 245


def _p2w_fix_char(t, font=""):
    mm = _re.match(r"^\(cid:(\d+)\)$", t)
    if mm:
        n = int(mm.group(1))
        if "wingdings" in font.lower() or "symbol" in font.lower():
            return chr(0xF000 + n) if n < 256 else "•"
        return "•" if n in (149, 151, 183) else ("\u00a0" if n < 33 else "•")
    if len(t) > 1 or 0xFB00 <= ord(t[0]) <= 0xFB06:
        return _ud.normalize("NFKC", t)
    return t


def _p2w_cluster(vals, tol):
    vals = sorted(vals)
    out, cur = [], [vals[0]]
    for v in vals[1:]:
        if v - cur[-1] <= tol:
            cur.append(v)
        else:
            out.append(sum(cur) / len(cur))
            cur = [v]
    out.append(sum(cur) / len(cur))
    return out


def _p2w_idx(vals, v):
    return min(range(len(vals)), key=lambda i: abs(vals[i] - v))


# ------------------------------------------------------------------ 1) chữ -> các đoạn dòng
def _p2w_segments(page, tables=()):
    ph = float(page.height)
    chars = []
    for c in page.chars:
        t = c.get("text", "")
        if not t or t.isspace() or not c.get("upright", True):
            continue
        m = c.get("matrix")
        base = ph - m[5] if m else c["bottom"] - 0.217 * c["size"]
        if not (c["top"] - 1 <= base <= c["bottom"] + 1):
            base = c["bottom"] - 0.217 * c["size"]
        chars.append({"t": _p2w_fix_char(t, c.get("fontname", "")), "x0": c["x0"], "x1": c["x1"], "base": base, "size": float(c["size"]),
                      "font": c.get("fontname", ""), "col": c.get("non_stroking_color")})
    if not chars:
        return []
    chars.sort(key=lambda c: (c["base"], c["x0"]))
    lines = []
    for c in chars:
        if lines and abs(c["base"] - lines[-1]["mb"]) <= 0.42 * max(c["size"], lines[-1]["ms"]):
            ln = lines[-1]
            ln["cs"].append(c)
            ln["mb"] = sum(x["base"] for x in ln["cs"]) / len(ln["cs"])
            ln["ms"] = max(ln["ms"], c["size"])
        else:
            lines.append({"cs": [c], "mb": c["base"], "ms": c["size"]})
    segs = []
    for ln in lines:
        cs = sorted(ln["cs"], key=lambda c: c["x0"])
        dd = []                                             # bỏ ký tự vẽ chồng (chữ đậm giả)
        for c in cs:
            if dd and c["t"] == dd[-1]["t"] and abs(c["x0"] - dd[-1]["x0"]) < 0.25 * c["size"] and abs(c["base"] - dd[-1]["base"]) < 1.0:
                continue
            dd.append(c)
        cs = dd
        sz_dom = _Counter(round(c["size"] * 2) / 2 for c in cs).most_common(1)[0][0]
        groups, cur = [], [cs[0]]
        for c in cs[1:]:
            xcut = c["x0"] - cur[-1]["x1"] > 0.3 * c["size"] and any(t["y0"] - 2 <= c["base"] <= t["y1"] + 2 and any(cur[-1]["x1"] <= xb + 1.5 and c["x0"] >= xb - 1.5 for xb in t["xs"][1:-1])
                       for t in tables)
            if c["x0"] - cur[-1]["x1"] > 2.0 * sz_dom or xcut:
                groups.append(cur)
                cur = [c]
            else:
                cur.append(c)
        groups.append(cur)
        for g in groups:
            segs.append(_p2w_make_seg(g))
    return segs


def _p2w_make_seg(g):
    sz_dom = _Counter(round(c["size"] * 2) / 2 for c in g).most_common(1)[0][0]
    base_dom = _stat.median([c["base"] for c in g if abs(c["size"] - sz_dom) < 1] or [g[0]["base"]])
    runs, prev = [], None
    for c in g:
        gap = (c["x0"] - prev["x1"]) if prev else 0
        if prev and gap > 0.15 * c["size"]:
            n = 1 if gap < 0.9 * c["size"] else min(8, max(1, round(gap / (0.27 * c["size"]))))
            sp = " " * n
        else:
            sp = ""
        font, b, i = _p2w_font(c["font"])
        if "wingdings" in c["font"].lower():
            font = "Wingdings"
        elif c["font"].lower().endswith("symbol") or "+symbol" in c["font"].lower():
            font = "Symbol"
        sup = 0
        if c["size"] <= 0.8 * sz_dom and (base_dom - c["base"]) >= 0.2 * c["size"]:
            sup = 1
        elif c["size"] <= 0.8 * sz_dom and (c["base"] - base_dom) >= 0.2 * c["size"]:
            sup = -1
        key = (font, round(c["size"] * 2) / 2, b, i, _p2w_color(c["col"]), sup)
        if runs and runs[-1]["key"] == key:
            runs[-1]["t"] += sp + c["t"]
            runs[-1]["x1"] = c["x1"]
        else:
            if runs and sp:
                runs[-1]["t"] += sp
            runs.append({"key": key, "t": c["t"], "x0": c["x0"], "x1": c["x1"]})
        prev = c
    out_runs = []
    for r in runs:
        f, s, b, i, col, sup = r["key"]
        out_runs.append({"t": r["t"], "font": f, "size": s, "b": b, "i": i, "u": False, "c": col, "sup": sup,
                         "x0": r["x0"], "x1": r["x1"]})
    x0, x1 = g[0]["x0"], max(c["x1"] for c in g)
    return {"kind": "seg", "x0": x0, "x1": x1, "base": base_dom, "size": sz_dom, "runs": out_runs,
            "y0": base_dom - 0.78 * sz_dom, "y1": base_dom + 0.22 * sz_dom}


# ------------------------------------------------------------------ 2) bảng có kẻ
def _p2w_find_tables(page, segs):
    pw, ph = float(page.width), float(page.height)
    res = []
    try:
        tabs = page.find_tables()
    except Exception:
        return res
    for t in tabs:
        cells = [tuple(c) for c in (t.cells or [])]
        if not cells:
            continue
        x0, top, x1, bot = t.bbox
        if len(cells) == 1 and (x1 - x0) * (bot - top) > 0.5 * pw * ph:
            continue                                         # khung viền cả trang, không phải bảng
        xs = _p2w_cluster([v for c in cells for v in (c[0], c[2])], 2.0)
        ys = _p2w_cluster([v for c in cells for v in (c[1], c[3])], 2.0)
        occ, spans = {}, []
        for c in cells:
            c0, c1 = _p2w_idx(xs, c[0]), _p2w_idx(xs, c[2])
            r0, r1 = _p2w_idx(ys, c[1]), _p2w_idx(ys, c[3])
            if c1 <= c0 or r1 <= r0:
                continue
            if any((r, cc) in occ for r in range(r0, r1) for cc in range(c0, c1)):
                continue
            for r in range(r0, r1):
                for cc in range(c0, c1):
                    occ[(r, cc)] = True
            spans.append({"r0": r0, "r1": r1, "c0": c0, "c1": c1})
        for r in range(len(ys) - 1):                         # ô thiếu -> ô trống 1x1
            for cc in range(len(xs) - 1):
                if (r, cc) not in occ:
                    spans.append({"r0": r, "r1": r + 1, "c0": cc, "c1": cc + 1})
        res.append({"kind": "tbl", "xs": xs, "ys": ys, "cells": spans, "x0": xs[0], "x1": xs[-1], "y0": ys[0],
                    "y1": ys[-1]})
    return res


def _p2w_assign_to_tables(segs, tables):
    rest = []
    for s in segs:
        cx, cy = (s["x0"] + s["x1"]) / 2, s["base"] - 0.3 * s["size"]
        placed = False
        for t in tables:
            if t["x0"] - 2 <= cx <= t["x1"] + 2 and t["y0"] - 2 <= cy <= t["y1"] + 2:
                for c in t["cells"]:
                    if t["xs"][c["c0"]] - 1 <= cx <= t["xs"][c["c1"]] + 1 and t["ys"][c["r0"]] - 1 <= cy <= t["ys"][c["r1"]] + 1:
                        c.setdefault("items", []).append(s)
                        placed = True
                        break
                if placed:
                    break
        if not placed:
            rest.append(s)
    return rest


def _p2w_cell_fills(page, tables):
    rects = []
    for r in page.rects:
        if r.get("fill") and not _p2w_is_white(r.get("non_stroking_color")):
            hx = _p2w_color(r.get("non_stroking_color"))
            rects.append((r["x0"], r["top"], r["x1"], r["bottom"], hx))
    for t in tables:
        for c in t["cells"]:
            cx0, cx1, cy0, cy1 = t["xs"][c["c0"]], t["xs"][c["c1"]], t["ys"][c["r0"]], t["ys"][c["r1"]]
            area = max((cx1 - cx0) * (cy1 - cy0), 1)
            for (a, b, d, e, hx) in rects:
                ix = max(0, min(cx1, d) - max(cx0, a))
                iy = max(0, min(cy1, e) - max(cy0, b))
                if ix * iy / area >= 0.8 and hx:
                    c["fill"] = hx
                    break


# ------------------------------------------------------------------ 3) ảnh + đồ họa vector -> khối nổi
def _p2w_graphics(page, segs, tables, prefer_text, rpt):
    pw, ph = float(page.width), float(page.height)
    boxes = []
    for im in page.images:
        w, h = im["x1"] - im["x0"], im["bottom"] - im["top"]
        if w < 2 or h < 2:
            continue
        cov = (w * h) / (pw * ph)
        boxes.append({"box": [im["x0"], im["top"], im["x1"], im["bottom"]], "img": True, "fill": False, "cov": cov})
    page_image = any(b["cov"] >= 0.85 for b in boxes)
    # đường kẻ mảnh nằm sát dưới chữ => gạch chân
    thin = []
    for ln in list(page.lines) + list(page.rects):
        w, h = ln["x1"] - ln["x0"], ln["bottom"] - ln["top"]
        if h < 2.0 and w > 4:
            thin.append(ln)
    used = set()
    for ti, ln in enumerate(thin):
        for s in segs:
            if abs(ln["top"] - (s["base"] + 0.06 * s["size"])) <= 0.12 * s["size"] + 0.8 and ln["x0"] >= s["x0"] - 2 \
                    and ln["x1"] <= s["x1"] + 2 and (ln["x1"] - ln["x0"]) >= 2.5 * s["size"]:
                for r in s["runs"]:
                    ov = min(r["x1"], ln["x1"]) - max(r["x0"], ln["x0"])
                    if ov > 0.5 * (r["x1"] - r["x0"]):
                        r["u"] = True
                used.add(id(ln))
                break

    def in_table(b):
        cx, cy = (b[0] + b[2]) / 2, (b[1] + b[3]) / 2
        return any(t["x0"] - 3 <= cx <= t["x1"] + 3 and t["y0"] - 3 <= cy <= t["y1"] + 3 for t in tables)

    vec = []
    for kind, lst in (("l", page.lines), ("c", page.curves), ("r", page.rects)):
        for g in lst:
            if id(g) in used:
                continue
            b = [g["x0"], g["top"], g["x1"], g["bottom"]]
            w, h = b[2] - b[0], b[3] - b[1]
            if w < 0.3 and h < 0.3:
                continue
            if kind == "r":
                filled = bool(g.get("fill")) and not _p2w_is_white(g.get("non_stroking_color"))
                if not (g.get("stroke") or filled) or w * h > 0.5 * pw * ph:
                    continue
            else:
                filled = bool(g.get("fill")) and not _p2w_is_white(g.get("non_stroking_color"))
            if in_table(b):
                continue
            vec.append({"box": b, "fill": filled})
    # gộp các nét vẽ gần nhau thành một cụm
    n = len(vec)
    parent = list(range(n))

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a
    for i in range(n):
        for j in range(i + 1, n):
            a, b = vec[i]["box"], vec[j]["box"]
            if a[0] - 2 <= b[2] and b[0] - 2 <= a[2] and a[1] - 2 <= b[3] and b[1] - 2 <= a[3]:
                parent[find(i)] = find(j)
    cl = {}
    for i in range(n):
        cl.setdefault(find(i), []).append(vec[i])
    for items in cl.values():
        box = [min(v["box"][0] for v in items), min(v["box"][1] for v in items),
               max(v["box"][2] for v in items), max(v["box"][3] for v in items)]
        boxes.append({"box": box, "img": False, "fill": any(v["fill"] for v in items), "cov": 0})
    # chữ xoay -> coi như hình
    rot = [c for c in page.chars if not c.get("upright", True) and c.get("text", "").strip()]
    if rot:
        boxes.append({"box": [min(c["x0"] for c in rot), min(c["top"] for c in rot), max(c["x1"] for c in rot),
                              max(c["bottom"] for c in rot)], "img": True, "fill": True, "cov": 0})
        rpt.append("Có chữ xoay ngang/dọc: được giữ dạng hình ảnh.")
    return boxes, page_image


def _p2w_blank(img, boxes_px):
    import numpy as np
    a = np.asarray(img).copy()
    H, W = a.shape[:2]
    for (x0, y0, x1, y1) in boxes_px:
        x0, y0, x1, y1 = max(0, x0 - 1), max(0, y0 - 1), min(W, x1 + 1), min(H, y1 + 1)
        if x1 - x0 < 2 or y1 - y0 < 2:
            continue
        X0, Y0, X1, Y1 = max(0, x0 - 4), max(0, y0 - 4), min(W, x1 + 4), min(H, y1 + 4)
        ring = np.concatenate([a[Y0:y0, X0:X1].reshape(-1, a.shape[2]), a[y1:Y1, X0:X1].reshape(-1, a.shape[2]),
                               a[y0:y1, X0:x0].reshape(-1, a.shape[2]), a[y0:y1, x1:X1].reshape(-1, a.shape[2])])
        col = np.median(ring, axis=0) if len(ring) else np.array([255] * a.shape[2])
        a[y0:y1, x0:x1] = col
    from PIL import Image
    return Image.fromarray(a)


def _p2w_crop(full, box, pw, ph, blank_boxes=None, transparent=False):
    from PIL import Image
    import numpy as np
    W, H = full.size
    sx, sy = W / pw, H / ph
    x0, y0 = int(max(0, math_floor(box[0] * sx) - 1)), int(max(0, math_floor(box[1] * sy) - 1))
    x1, y1 = int(min(W, _math.ceil(box[2] * sx) + 1)), int(min(H, _math.ceil(box[3] * sy) + 1))
    if x1 - x0 < 2 or y1 - y0 < 2:
        return None
    crop = full.crop((x0, y0, x1, y1)).convert("RGB")
    if blank_boxes:
        px = []
        for b in blank_boxes:
            bx0, by0, bx1, by1 = int(b[0] * sx) - x0, int(b[1] * sy) - y0, int(b[2] * sx) - x0, int(b[3] * sy) - y0
            if bx1 > 0 and by1 > 0 and bx0 < crop.size[0] and by0 < crop.size[1]:
                px.append((bx0, by0, bx1, by1))
        if px:
            crop = _p2w_blank(crop, px)
    out = _io.BytesIO()
    if transparent:
        arr = np.asarray(crop).astype("float32")
        alpha = 255.0 - arr.min(axis=2)
        a = np.clip(alpha, 1, 255)[..., None]
        col = np.clip((arr - (255.0 - a)) * 255.0 / a, 0, 255)
        rgba = np.dstack([col, alpha]).astype("uint8")
        Image.fromarray(rgba, "RGBA").save(out, format="PNG")
    elif crop.size[0] * crop.size[1] > 300000:
        crop.save(out, format="JPEG", quality=92)
    else:
        crop.save(out, format="PNG")
    return out.getvalue()


def math_floor(v):
    return int(_math.floor(v))


# ------------------------------------------------------------------ 4) XY-cut: chia cột / chia khối
def _p2w_gaps(intervals):
    iv = sorted(intervals)
    gaps, cur_end = [], iv[0][1]
    for a, b in iv[1:]:
        if a > cur_end:
            gaps.append((a - cur_end, (a + cur_end) / 2.0))
        cur_end = max(cur_end, b)
    return gaps


def _p2w_leaf(items):
    nodes, cur = [], []
    for it in sorted(items, key=lambda i: (i["y0"], i["x0"])):
        if it["kind"] == "tbl":
            if cur:
                nodes.append(("lines", cur))
                cur = []
            nodes.append(("tbl", it))
        else:
            cur.append(it)
    if cur:
        nodes.append(("lines", cur))
    return _p2w_merge_lines(nodes)


def _p2w_merge_lines(nodes):
    out = []
    for n in nodes:
        if n[0] == "lines" and out and out[-1][0] == "lines":
            out[-1] = ("lines", out[-1][1] + n[1])
        else:
            out.append(n)
    return out


def _p2w_xycut(items):
    if not items:
        return []
    if len(items) == 1:
        return _p2w_leaf(items)
    sizes = [i["size"] for i in items if i["kind"] == "seg"] or [12.0]
    ms = _stat.median(sizes)
    vg = _p2w_gaps([(i["x0"], i["x1"]) for i in items])
    hg = _p2w_gaps([(i["y0"], i["y1"]) for i in items])
    vmin, hmin = max(14.0, 1.6 * ms), 0.25 * ms
    vbest = max([g for g, _ in vg], default=0)
    hbest = max([g for g, _ in hg], default=0)
    if vbest >= vmin and vbest >= hbest:
        cuts = sorted(p for g, p in vg if g >= vmin)
        groups = [[] for _ in range(len(cuts) + 1)]
        for it in items:
            cx = (it["x0"] + it["x1"]) / 2
            groups[sum(1 for c in cuts if cx > c)].append(it)
        groups = [g for g in groups if g]
        if len(groups) >= 2:
            cols = [{"items": g, "x0": min(i["x0"] for i in g), "x1": max(i["x1"] for i in g),
                     "nodes": _p2w_xycut(g)} for g in groups]
            return [("cols", cols)]
    if hbest >= hmin:
        pos = max(hg)[1]
        top = [i for i in items if (i["y0"] + i["y1"]) / 2 < pos]
        bot = [i for i in items if (i["y0"] + i["y1"]) / 2 >= pos]
        if top and bot:
            return _p2w_merge_lines(_p2w_xycut(top) + _p2w_xycut(bot))
    return _p2w_leaf(items)


# ------------------------------------------------------------------ 5) gom dòng thành đoạn văn
def _p2w_pair_kind(a, b, regL):
    ca, cb = (a["x0"] + a["x1"]) / 2, (b["x0"] + b["x1"]) / 2
    dx0, dx1 = abs(a["x0"] - b["x0"]), abs(a["x1"] - b["x1"])
    if dx0 <= 3:
        return "left"
    if abs(ca - cb) <= 3:
        return "center"
    if dx1 <= 3:
        return "indent" if abs(b["x0"] - regL) <= 4 or abs(a["x0"] - regL) <= 4 else "right"
    d = a["x0"] - b["x0"]
    if abs(d) <= 5 * max(a["size"], b["size"]) and (abs(b["x0"] - regL) <= 4 or abs(a["x0"] - regL) <= 4):
        return "indent"
    return None


def _p2w_form_paras(segs, regL, regR):
    segs = sorted(segs, key=lambda s: (s["base"], s["x0"]))
    paras, cur, kind = [], [], None
    for s in segs:
        if not cur:
            cur, kind = [s], None
            continue
        p = cur[-1]
        gap = s["base"] - p["base"]
        ok = abs(s["size"] - p["size"]) <= 1.0
        if ok:
            if len(cur) == 1:
                ok = 0.85 * p["size"] <= gap <= 1.75 * p["size"]
                k = _p2w_pair_kind(p, s, regL) if ok else None
                ok = k is not None
            else:
                ref = (cur[-1]["base"] - cur[0]["base"]) / (len(cur) - 1)
                ok = abs(gap - ref) <= 0.15 * ref + 1.0
                if ok:
                    k = kind
                    if k == "left" or k == "indent":
                        ok = abs(s["x0"] - cur[-1]["x0"]) <= 3 if len(cur) > 2 or kind == "left" else \
                            abs(s["x0"] - cur[1]["x0"]) <= 3
                        if kind == "indent" and len(cur) == 2 and abs(s["x0"] - cur[1]["x0"]) > 3:
                            ok = False
                    elif k == "center":
                        ok = abs((s["x0"] + s["x1"]) / 2 - (cur[-1]["x0"] + cur[-1]["x1"]) / 2) <= 3
                    elif k == "right":
                        ok = abs(s["x1"] - cur[-1]["x1"]) <= 3
        if ok:
            if len(cur) == 1:
                kind = k
            cur.append(s)
        else:
            paras.append(cur)
            cur, kind = [s], None
    if cur:
        paras.append(cur)
    return [_p2w_para_info(p, regL, regR) for p in paras]


def _p2w_para_info(segs, regL, regR):
    n = len(segs)
    W = max(regR - regL, 1.0)
    x0s, x1s = [s["x0"] for s in segs], [s["x1"] for s in segs]
    cs = [(a + b) / 2 for a, b in zip(x0s, x1s)]
    rc = (regL + regR) / 2
    fs0 = segs[0]["size"]
    align = "left"
    if n == 1:
        if x0s[0] - regL <= max(4.0, 0.02 * W):
            align = "left"
        elif abs(cs[0] - rc) <= max(6.0, 0.03 * W):
            align = "center"
        elif regR - x1s[0] <= max(4.0, 0.015 * W):
            align = "right"
    else:
        ref0 = _stat.median(x0s[1:]) if n > 1 else x0s[0]
        if max(cs) - min(cs) <= 3 and (max(x0s) - min(x0s)) > 4:
            align = "center"
        elif max(x0s[1:]) - min(x0s[1:]) <= 3.5:
            rights = x1s[:-1]
            if len(rights) >= 1 and max(rights) - min(rights) <= 3 and max(rights) >= regR - 0.06 * W and \
                    abs(x0s[0] - ref0) <= 5 * fs0:
                align = "justify"
        elif max(x1s) - min(x1s) <= 3:
            align = "right"
    pitches = [segs[i + 1]["base"] - segs[i]["base"] for i in range(n - 1)]
    pitch = _stat.median(pitches) if pitches else 1.2 * fs0
    pitch = max(pitch, 0.9 * fs0)
    ref = _stat.median(x0s[1:]) if n > 1 else x0s[0]
    return {"segs": segs, "align": align, "pitch": pitch, "fs0": fs0, "base0": segs[0]["base"], "n": n,
            "ref": ref, "fi": (x0s[0] - ref) if n > 1 else 0.0, "x0": min(x0s), "x1": max(x1s),
            "cmean": _stat.mean(cs), "rights": x1s}


def _p2w_para_top(p):
    return p["base0"] - p["pitch"] + P2W_DESC * p["fs0"]


def _p2w_nodes_top(nodes, regL, regR):
    tops = []
    for n in nodes:
        if n[0] == "lines":
            tops += [_p2w_para_top(p) for p in _p2w_form_paras(n[1], regL, regR)]
        elif n[0] == "cols":
            tops += [_p2w_nodes_top(c["nodes"], c["x0"], c["x1"]) for c in n[1]]
        else:
            tops.append(n[1]["y0"])
    return min(tops) if tops else 0.0


def _p2w_nodes_bottom(nodes):
    bs = []
    for n in nodes:
        if n[0] == "lines":
            bs += [s["y1"] for s in n[1]]
        elif n[0] == "cols":
            bs += [_p2w_nodes_bottom(c["nodes"]) for c in n[1]]
        else:
            bs.append(n[1]["y1"])
    return max(bs) if bs else 0.0


# ------------------------------------------------------------------ 6) thao tác Word mức thấp
def _p2w_xml():
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import qn, nsdecls
    return OxmlElement, parse_xml, qn, nsdecls


def _p2w_mark_size(p, half_points):
    OxmlElement, _, qn, _ = _p2w_xml()
    pPr = p._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        sect = pPr.find(qn("w:sectPr"))
        if sect is not None:
            sect.addprevious(rPr)
        else:
            pPr.append(rPr)
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), str(half_points))


def _p2w_tiny(p, h=1.0):
    from docx.shared import Pt
    from docx.enum.text import WD_LINE_SPACING
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(max(1.0, min(h, 1500.0)))
    _p2w_mark_size(p, 2)


def _p2w_run(run, name, size, b, i, u, color, sup):
    from docx.shared import Pt, RGBColor
    OxmlElement, _, qn, _ = _p2w_xml()
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), name)
    run.font.size = Pt(max(1.0, round(size * 2) / 2))
    run.bold, run.italic, run.underline = bool(b), bool(i), bool(u)
    if color:
        run.font.color.rgb = RGBColor(int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16))
    if sup == 1:
        run.font.superscript = True
    elif sup == -1:
        run.font.subscript = True


def _p2w_add_text(par, runs):
    for r in runs:
        for k, part in enumerate(r["t"].split("\n")):
            if k > 0:
                par.add_run().add_break()
            if part == "":
                continue
            run = par.add_run(part)
            _p2w_run(run, r["font"], r["size"], r["b"], r["i"], r["u"], r["c"], r["sup"])


def _p2w_page_setup(sec, w, h, ml, mr, mt, mb):
    from docx.shared import Pt, Mm
    from docx.enum.section import WD_ORIENT
    if abs(w - 595.28) < 8 and abs(h - 841.89) < 8:
        sec.page_width, sec.page_height = Mm(210), Mm(297)
    elif abs(w - 841.89) < 8 and abs(h - 595.28) < 8:
        sec.page_width, sec.page_height = Mm(297), Mm(210)
    else:
        sec.page_width, sec.page_height = Pt(w), Pt(h)
    sec.orientation = WD_ORIENT.LANDSCAPE if w > h else WD_ORIENT.PORTRAIT
    sec.left_margin, sec.right_margin = Pt(ml), Pt(mr)
    sec.top_margin, sec.bottom_margin = Pt(mt), Pt(mb)
    sec.header_distance = Pt(0)
    sec.footer_distance = Pt(0)


def _p2w_float(run, data, x, y, w, h, behind=True):
    from docx.shared import Pt
    OxmlElement, parse_xml, qn, nsdecls = _p2w_xml()
    shape = run.add_picture(_io.BytesIO(data), width=Pt(max(w, 1)), height=Pt(max(h, 1)))
    inline = shape._inline
    graphic = inline.find(qn("a:graphic"))
    docPr = inline.find(qn("wp:docPr"))
    anchor = parse_xml(
        '<wp:anchor %s distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" '
        'behindDoc="%d" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:posOffset>%d</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="page"><wp:posOffset>%d</wp:posOffset></wp:positionV>'
        '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        '<wp:docPr id="%s" name="%s"/><wp:cNvGraphicFramePr/></wp:anchor>'
        % (nsdecls("wp", "a"), 1 if behind else 0, int(Pt(x)), int(Pt(y)), int(Pt(w)), int(Pt(h)),
           docPr.get("id"), docPr.get("name")))
    anchor.append(graphic)
    drawing = inline.getparent()
    drawing.remove(inline)
    drawing.append(anchor)


def _p2w_tbl_props(tbl, ind_pt, width_pt, borders, mar_lr, mar_tb):
    OxmlElement, _, qn, _ = _p2w_xml()
    tblPr = tbl._tbl.tblPr
    for tag in ("w:tblW", "w:tblInd", "w:tblLayout", "w:tblCellMar", "w:jc"):
        el = tblPr.find(qn(tag))
        if el is not None:
            tblPr.remove(el)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(int(width_pt * 20)))
    tblW.set(qn("w:type"), "dxa")
    ind = OxmlElement("w:tblInd")
    ind.set(qn("w:w"), str(int(ind_pt * 20)))
    ind.set(qn("w:type"), "dxa")
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    mar = OxmlElement("w:tblCellMar")
    for side, v in (("top", mar_tb), ("left", mar_lr), ("bottom", mar_tb), ("right", mar_lr)):
        e = OxmlElement("w:" + side)
        e.set(qn("w:w"), str(int(v * 20)))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    look = tblPr.find(qn("w:tblLook"))
    for el in (tblW, ind, lay, mar):
        if look is not None:
            look.addprevious(el)
        else:
            tblPr.append(el)
    if borders:
        tbl.style = "Table Grid"


def _p2w_shade(cell, hexcolor):
    OxmlElement, _, qn, _ = _p2w_xml()
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def _p2w_compat15(doc):
    OxmlElement, _, qn, _ = _p2w_xml()
    st = doc.settings.element
    compat = st.find(qn("w:compat"))
    if compat is None:
        compat = OxmlElement("w:compat")
        st.append(compat)
    for cs in compat.findall(qn("w:compatSetting")):
        if cs.get(qn("w:name")) == "compatibilityMode":
            cs.set(qn("w:val"), "15")
            return
    cs = OxmlElement("w:compatSetting")
    cs.set(qn("w:name"), "compatibilityMode")
    cs.set(qn("w:uri"), "http://schemas.microsoft.com/office/word")
    cs.set(qn("w:val"), "15")
    compat.append(cs)


# ------------------------------------------------------------------ 7) phát (emit) nội dung vào Word
class _P2WCtx:
    def __init__(self, doc, cell, regL, regR, flow_y, ml, reflow):
        self.doc, self.cell, self.regL, self.regR, self.flow_y, self.ml, self.reflow = doc, cell, regL, regR, flow_y, ml, reflow
        self.fresh = cell is not None
        self.emitted = 0

    def new_par(self):
        self.emitted += 1
        if self.cell is not None:
            if self.fresh:
                self.fresh = False
                return self.cell.paragraphs[0]
            return self.cell.add_paragraph()
        return self.doc.add_paragraph()


def _p2w_emit_para(ctx, p):
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    AL = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT,
          "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
    segs, al = p["segs"], p["align"]
    top = _p2w_para_top(p)
    par = ctx.new_par()
    pf = par.paragraph_format
    pf.space_before = Pt(max(0.0, top - ctx.flow_y))
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(p["pitch"])
    pf.alignment = AL[al]
    regL, regR = ctx.regL, ctx.regR
    ext = max(14.0, 0.12 * (regR - regL))
    if al == "center":
        half = max((p["x1"] - p["x0"]) / 2.0 + ext, 1.0)
        left, right = p["cmean"] - half - regL, regR - (p["cmean"] + half)
        li, ri, fi = max(0.0, left), max(0.0, right), 0.0
    elif al == "right":
        li, ri, fi = max(0.0, p["x0"] - ext - regL), max(0.0, regR - max(p["rights"])), 0.0
    elif al == "justify":
        rr = max(p["rights"][:-1]) if p["n"] > 1 else p["x1"]
        li, ri, fi = max(0.0, p["ref"] - regL), max(0.0, regR - rr), p["fi"]
    else:
        li, ri, fi = max(0.0, p["ref"] - regL), max(-ext, regR - (p["x1"] + ext)), p["fi"]
    pf.left_indent = Pt(li)
    pf.right_indent = Pt(ri)
    for sg in segs:
        for tx in sg.get("tabs", []):
            pf.tab_stops.add_tab_stop(Pt(max(0.0, tx - regL)))
    if abs(fi) > 1.5:
        pf.first_line_indent = Pt(max(-li, fi))
    # chữ: mỗi dòng PDF = 1 dòng Word (ngắt mềm) -> giữ đúng điểm xuống dòng; reflow=True thì nối các dòng "đầy"
    runs = []
    for k, s in enumerate(segs):
        sr = [dict(r) for r in s["runs"]]
        if k < len(segs) - 1:
            soft = ctx.reflow and al in ("left", "justify") and s["x1"] >= regR - 0.05 * (regR - regL)
            sr[-1]["t"] = sr[-1]["t"] + (" " if soft else "\n")
        runs += sr
    _p2w_add_text(par, runs)
    ctx.flow_y = max(ctx.flow_y, top) + p["n"] * p["pitch"]


def _p2w_emit_nodes(ctx, nodes):
    for n in nodes:
        if n[0] == "lines":
            for p in _p2w_form_paras(n[1], ctx.regL, ctx.regR):
                _p2w_emit_para(ctx, p)
        elif n[0] == "cols":
            _p2w_emit_cols(ctx, n[1])
        else:
            _p2w_emit_table(ctx, n[1])


def _p2w_spacer(ctx, gap):
    from docx.shared import Pt
    sp = ctx.new_par()
    _p2w_tiny(sp, max(gap, 1.0))
    ctx.flow_y += max(gap, 1.0)


def _p2w_empty_cell(cell):
    for p in cell.paragraphs:
        if not p.text.strip() and not p.runs:
            _p2w_tiny(p, 1.0)


def _p2w_emit_cols(ctx, cols):
    from docx.shared import Pt
    regL, regR = ctx.regL, ctx.regR
    top = min(_p2w_nodes_top(c["nodes"], c["x0"], c["x1"]) for c in cols)
    bot = max(_p2w_nodes_bottom(c["nodes"]) for c in cols)
    bounds = [min(c["x0"] for c in cols)]
    for a, b in zip(cols, cols[1:]):
        bounds.append((a["x1"] + b["x0"]) / 2.0)
    bounds.append(max(c["x1"] for c in cols) + 2.0)
    _p2w_spacer(ctx, top - ctx.flow_y)
    tbl = (ctx.cell.add_table(1, len(cols)) if ctx.cell is not None else ctx.doc.add_table(rows=1, cols=len(cols)))
    tbl.autofit = False
    _p2w_tbl_props(tbl, bounds[0] - ctx.ml if ctx.cell is None else max(0.0, bounds[0] - ctx.regL), bounds[-1] - bounds[0],
                   False, 0.0, 0.0)
    tbl.rows[0].height = Pt(max(bot - top, 2.0))
    from docx.enum.table import WD_ROW_HEIGHT_RULE
    tbl.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    for i, c in enumerate(cols):
        w = bounds[i + 1] - bounds[i]
        tbl.columns[i].width = Pt(w)
        cell = tbl.cell(0, i)
        cell.width = Pt(w)
        sub = _P2WCtx(ctx.doc, cell, bounds[i], bounds[i + 1], top, ctx.ml, ctx.reflow)
        _p2w_emit_nodes(sub, c["nodes"])
        if sub.emitted == 0:
            _p2w_empty_cell(cell)
    ctx.flow_y = max(ctx.flow_y, bot)


def _p2w_emit_table(ctx, t):
    from docx.shared import Pt
    from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_ALIGN_VERTICAL
    xs, ys = t["xs"], t["ys"]
    nr, nc = len(ys) - 1, len(xs) - 1
    MAR_LR, MAR_TB = 3.5, 0.75
    _p2w_spacer(ctx, t["y0"] - ctx.flow_y)
    tbl = ctx.doc.add_table(rows=nr, cols=nc) if ctx.cell is None else ctx.cell.add_table(nr, nc)
    tbl.autofit = False
    _p2w_tbl_props(tbl, (xs[0] - ctx.ml) if ctx.cell is None else max(0.0, xs[0] - ctx.regL), xs[-1] - xs[0], True,
                   MAR_LR, MAR_TB)
    for i in range(nc):
        tbl.columns[i].width = Pt(xs[i + 1] - xs[i])
    for r in range(nr):
        tbl.rows[r].height = Pt(max(ys[r + 1] - ys[r], 2.0))
        tbl.rows[r].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    jobs = []
    for c in t["cells"]:
        a = tbl.cell(c["r0"], c["c0"])
        if c["r1"] - c["r0"] > 1 or c["c1"] - c["c0"] > 1:
            a = a.merge(tbl.cell(c["r1"] - 1, c["c1"] - 1))
            while len(a.paragraphs) > 1 and not a.paragraphs[-1].text.strip():
                a._tc.remove(a.paragraphs[-1]._p)
        jobs.append((a, c))
    for cell, c in jobs:
        x0c, x1c, y0c, y1c = xs[c["c0"]], xs[c["c1"]], ys[c["r0"]], ys[c["r1"]]
        cell.width = Pt(x1c - x0c)
        if c.get("fill"):
            _p2w_shade(cell, c["fill"])
        items = c.get("items") or []
        regL, regR = x0c + MAR_LR, x1c - MAR_LR
        nodes = _p2w_xycut(items)
        start = y0c + MAR_TB
        if nodes:
            ctop, cbot = _p2w_nodes_top(nodes, regL, regR), _p2w_nodes_bottom(nodes)
            gt, gb = ctop - y0c, y1c - cbot
            if gt > 1.5 and abs(gt - gb) <= max(2.0, 0.15 * (y1c - y0c)):
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                start = ctop
        sub = _P2WCtx(ctx.doc, cell, regL, regR, start, ctx.ml, ctx.reflow)
        _p2w_emit_nodes(sub, nodes)
        if sub.emitted == 0:
            _p2w_empty_cell(cell)
    ctx.flow_y = max(ctx.flow_y, t["y1"])


# ------------------------------------------------------------------ 8) mô hình trang + dựng file
def _p2w_merge_icon_rows(segs, boxes):
    """Hình nhỏ nằm giữa dòng chữ (icon phím, vòng tròn...) -> các đoạn chữ hai bên được nối thành 1 dòng bằng tab."""
    icons = [b["box"] for b in boxes if b["img"] and b["cov"] < 0.05 and (b["box"][2] - b["box"][0]) <= 90
             and (b["box"][3] - b["box"][1]) <= 70]
    if not icons or not segs:
        return segs
    segs = list(segs)
    for ib in sorted(icons, key=lambda b: -b[0]):
        icx, icy = (ib[0] + ib[2]) / 2, (ib[1] + ib[3]) / 2
        cand = [s for s in segs if s.get("kind") == "seg" and ib[1] - 3 <= s["base"] <= ib[3] + 3 + 0.3 * s["size"]]
        if not cand:
            continue
        # chọn dòng gần tâm icon nhất, có chữ sát bên trái hoặc bên phải icon
        rows = {}
        for s in cand:
            rows.setdefault(round(s["base"] / 3.0), []).append(s)
        best = None
        for k, row in rows.items():
            lft = [s for s in row if s["x1"] <= ib[0] + 6 and ib[0] - s["x1"] <= 5 * s["size"]]
            rgt = [s for s in row if s["x0"] >= ib[2] - 6 and s["x0"] - ib[2] <= 5 * s["size"]]
            if not (lft or rgt):
                continue
            by = sum(s["base"] for s in row) / len(row)
            d = abs(by - icy)
            if best is None or d < best[0]:
                best = (d, lft, rgt)
        if not best:
            continue
        _, lft, rgt = best
        A = max(lft, key=lambda s: s["x1"]) if lft else None
        B = min(rgt, key=lambda s: s["x0"]) if rgt else None
        if A is not None and B is not None and A is not B:
            tab = dict(A["runs"][-1])
            tab.update({"t": "\t", "u": False, "x0": A["x1"], "x1": B["x0"]})
            A["runs"] = A["runs"] + [tab] + B["runs"]
            A["tabs"] = A.get("tabs", []) + [B["x0"]] + B.get("tabs", [])
            A["x1"] = B["x1"]
            A["size"] = max(A["size"], B["size"])
            segs.remove(B)
    return segs


def p2w_page_model(page, full_img, opts, rpt, pno):
    pw, ph = float(page.width), float(page.height)
    tables = _p2w_find_tables(page, None)
    segs = _p2w_segments(page, tables)
    boxes, page_image = _p2w_graphics(page, segs, tables, opts.get("prefer_text", False), rpt)
    # trang quét (ảnh phủ kín trang) hoặc trang chỉ có hình
    if page_image and not opts.get("prefer_text", False):
        data = _p2w_crop(full_img, [0, 0, pw, ph], pw, ph)
        if segs:
            rpt.append(f"Trang {pno}: là trang quét có lớp chữ ẩn -> giữ nguyên dạng ảnh để bố cục chính xác 100% "
                       f"(chữ không sửa được).")
        else:
            rpt.append(f"Trang {pno}: là trang quét (không có lớp chữ) -> giữ dạng ảnh. Muốn chữ sửa được hãy dùng "
                       f"app 1 (AI/OCR).")
        return {"w": pw, "h": ph, "mode": "image", "floats": [(data, 0.0, 0.0, pw, ph)], "nodes": []}
    if not segs and not tables and not boxes:
        return {"w": pw, "h": ph, "mode": "blank", "floats": [], "nodes": []}
    if page_image:
        boxes = [b for b in boxes if b["cov"] < 0.85]
    segs = _p2w_merge_icon_rows(segs, boxes)
    segs = _p2w_assign_to_tables(segs, tables)
    _p2w_cell_fills(page, tables)
    floats = []
    def _wb(s):
        return [[r.get("x0", s["x0"]), s["base"] - 0.85 * s["size"], r.get("x1", s["x1"]), s["base"] + 0.25 * s["size"]]
                for r in s["runs"] if r["t"].strip()]
    word_boxes = [w for s in segs for w in _wb(s)]
    for t in tables:
        for c in t["cells"]:
            for s in c.get("items", []):
                word_boxes += _wb(s)
    for b in boxes:
        x0, y0, x1, y1 = b["box"]
        if b["img"]:
            inside = [w for w in word_boxes if w[0] < x1 and w[2] > x0 and w[1] < y1 and w[3] > y0
                      and (min(w[2], x1) - max(w[0], x0)) > 0.5 and (min(w[3], y1) - max(w[1], y0)) > 0.3]
            data = _p2w_crop(full_img, b["box"], pw, ph, blank_boxes=inside or None)
        else:
            inside = [w for w in word_boxes if w[0] < x1 and w[2] > x0 and w[1] < y1 and w[3] > y0]
            data = _p2w_crop(full_img, [x0 - 1, y0 - 1, x1 + 1, y1 + 1], pw, ph, blank_boxes=inside,
                             transparent=not b["fill"])
            x0, y0, x1, y1 = x0 - 1, y0 - 1, x1 + 1, y1 + 1
        if data:
            floats.append((data, x0, y0, x1 - x0, y1 - y0))
    items = segs + tables
    nodes = _p2w_xycut(items) if items else []
    return {"w": pw, "h": ph, "mode": "flow", "floats": floats, "nodes": nodes, "items": items}


def p2w_build_docx(models, opts):
    from docx import Document
    from docx.shared import Pt
    from docx.enum.section import WD_SECTION
    OxmlElement, _, qn, _ = _p2w_xml()
    doc = Document()
    _p2w_compat15(doc)
    nm = doc.styles["Normal"]
    nm.font.name = "Times New Roman"
    nm.font.size = Pt(11)
    rf = nm.element.get_or_add_rPr().find(qn("w:rFonts"))
    if rf is not None:
        rf.set(qn("w:eastAsia"), "Times New Roman")
    npf = nm.paragraph_format
    npf.space_before = npf.space_after = Pt(0)
    npf.line_spacing = 1.0
    n_pages = len(models)
    for pi, m in enumerate(models):
        sec = doc.sections[-1]
        w, h = m["w"], m["h"]
        if m["mode"] == "flow":
            its = m["items"]
            ml = max(0.0, min(i["x0"] for i in its))
            mr = max(0.0, w - max(i["x1"] for i in its))
            ml, mr = min(ml, 0.45 * w), min(mr, 0.45 * w)
            tops = [(_p2w_nodes_top([n], i0, i1)) for n in m["nodes"]
                    for i0, i1 in [(ml, w - mr)]] or [36.0]
            mt = max(0.0, min(tops) - 1.0)
            mb = max(0.0, min(70.0, h - max(i["y1"] for i in its) - 14.0))
            _p2w_page_setup(sec, w, h, ml, mr, mt, mb)
        else:
            ml = mr = mt = mb = 0.0
            _p2w_page_setup(sec, w, h, 0, 0, 0, 0)
        host = doc.add_paragraph()
        _p2w_tiny(host, 1.0)
        for data, x, y, fw, fh in m["floats"]:
            _p2w_float(host.add_run(), data, x, y, fw, fh, behind=True)
        if m["mode"] == "flow":
            ctx = _P2WCtx(doc, None, ml, w - mr, mt + 1.0, ml, opts.get("reflow", False))
            _p2w_emit_nodes(ctx, m["nodes"])
        if pi < n_pages - 1:
            doc.add_section(WD_SECTION.NEW_PAGE)
            _p2w_tiny(doc.paragraphs[-1], 1.0)
        else:
            _p2w_tiny(doc.add_paragraph(), 1.0)
    out = _io.BytesIO()
    doc.save(out)
    return out.getvalue()


def p2w_engine_b(pdf_bytes, opts, progress=None):
    import pdfplumber
    import pypdfium2 as pdfium
    rpt, models = [], []
    pdf = pdfplumber.open(_io.BytesIO(pdf_bytes))
    pdfi = pdfium.PdfDocument(pdf_bytes)
    n = len(pdf.pages)
    for i, page in enumerate(pdf.pages):
        full = pdfi[i].render(scale=opts.get("dpi", 200) / 72.0).to_pil().convert("RGB")
        models.append(p2w_page_model(page, full, opts, rpt, i + 1))
        if progress:
            progress((i + 1) / n, f"Đang phân tích trang {i + 1}/{n}...")
    pdf.close()
    return p2w_build_docx(models, opts), rpt


# ------------------------------------------------------------------ Engine A + kiểm tra độ phủ chữ
def p2w_engine_a(pdf_bytes):
    import tempfile
    import os
    from pdf2docx import Converter
    with tempfile.TemporaryDirectory() as td:
        src, dst = os.path.join(td, "in.pdf"), os.path.join(td, "out.docx")
        with open(src, "wb") as f:
            f.write(pdf_bytes)
        cv = Converter(src)
        try:
            cv.convert(dst)
        finally:
            cv.close()
        with open(dst, "rb") as f:
            return f.read()


def _p2w_words(text):
    text = _re.sub(r"\(cid:\d+\)", " ", text)
    text = "".join(ch for ch in text if not (0xE000 <= ord(ch) <= 0xF8FF))
    return _re.findall(r"\w+", _ud.normalize("NFC", text).lower())


def _p2w_tokens(text):
    text = _re.sub(r"\(cid:\d+\)", " ", text)
    text = "".join(ch for ch in text if not (0xE000 <= ord(ch) <= 0xF8FF))
    return [(_ud.normalize("NFC", w).lower(), w) for w in _re.findall(r"\w+", text)]


def p2w_text_diff(pdf_bytes, docx_bytes):
    """So khớp từng từ PDF <-> Word. Trả về dict: cov, n, miss, extra, missing[{word,page,ctx}], extra[{word,ctx}]."""
    import pdfplumber
    from docx import Document
    OxmlElement, _, qn, _ = _p2w_xml()
    pt = []                                            # (norm, orig, page)
    with pdfplumber.open(_io.BytesIO(pdf_bytes)) as pdf:
        for pi, p in enumerate(pdf.pages):
            pt += [(n, o, pi + 1) for n, o in _p2w_tokens(p.extract_text() or "")]
    d = Document(_io.BytesIO(docx_bytes))
    parts = []
    for el in d.element.body.iter():
        if el.tag == qn("w:t"):
            parts.append(el.text or "")
        elif el.tag in (qn("w:br"), qn("w:tab"), qn("w:p"), qn("w:tc")):
            parts.append(" ")
    dt = _p2w_tokens("".join(parts))
    a, b = _Counter(n for n, _, _ in pt), _Counter(n for n, _ in dt)
    miss_c, extra_c = a - b, b - a
    missing, extra = [], []
    for k, (n, o, pg) in enumerate(pt):
        if miss_c.get(n, 0) > 0:
            miss_c[n] -= 1
            missing.append({"word": o, "page": pg, "ctx": " ".join(x[1] for x in pt[max(0, k - 4):k + 5])})
    for k, (n, o) in enumerate(dt):
        if extra_c.get(n, 0) > 0:
            extra_c[n] -= 1
            extra.append({"word": o, "ctx": " ".join(x[1] for x in dt[max(0, k - 4):k + 5])})
    total = max(len(pt), 1)
    return {"cov": (len(pt) - len(missing)) / total, "n": len(pt), "miss": len(missing), "extra": len(extra),
            "missing": missing, "extra_list": extra}


def p2w_text_coverage(pdf_bytes, docx_bytes):
    r = p2w_text_diff(pdf_bytes, docx_bytes)
    return r["cov"], r["n"], r["miss"], r["extra"]



# =====================================================================================
# APP 12: PDF -> WORD (OFFLINE)
# =====================================================================================
def app_pdf_to_word_offline():
    st.header("📝 12. PDF sang Word (Offline)")
    st.caption("Chạy hoàn toàn trên máy chủ, không gửi dữ liệu cho AI. Giữ nguyên vị trí chữ, bảng, hình ảnh, con dấu.")
    up = st.file_uploader("Tải lên file PDF:", type=["pdf"], key=f"app12_{st.session_state.uploader_key}")
    c1, c2 = st.columns(2)
    with c1:
        engine = st.radio("Bộ máy chuyển đổi:", ["🅱️ Giữ đúng vị trí (khuyên dùng)", "🅰️ pdf2docx (dễ chỉnh sửa)", "⚖️ Cả hai (tải ZIP để so sánh)"], key="p2w_engine")
    with c2:
        reflow = st.checkbox("Chữ tự xuống dòng (dễ sửa)", value=False, key="p2w_reflow",
                             help="TẮT (khuyên dùng): mỗi dòng PDF = 1 dòng Word, giống PDF nhất nhưng khi sửa chữ có thể phải xuống dòng thủ công. "
                                  "BẬT: Word tự dồn chữ, sửa thoải mái hơn nhưng điểm ngắt dòng có thể lệch so với PDF.")
        prefer_text = st.checkbox("PDF scan có lớp chữ: lấy chữ gõ được", value=False, key="p2w_pt",
                                  help="Chỉ bật khi PDF là bản scan nhưng đã được OCR (bôi đen chọn được chữ) và bạn muốn sửa chữ. "
                                       "TẮT: trang scan giữ nguyên là ảnh nên giống bản gốc 100% nhưng không sửa được chữ.")
        dpi = st.select_slider("Độ nét hình ảnh (DPI):", options=[150, 200, 300], value=200, key="p2w_dpi",
                               help="150: file nhẹ. 200 (khuyên dùng): cân bằng. 300: hình/con dấu nét hơn để in, file nặng và chạy chậm hơn.")
    if not up:
        st.info("💡 PDF tạo từ Word/Excel cho kết quả ~100%. PDF scan (ảnh) sẽ được giữ nguyên dạng ảnh; muốn chữ sửa được hãy dùng App 1 (AI online).")
        return
    if st.button("🚀 Chuyển sang Word", type="primary"):
        data = up.getvalue()
        base = os.path.splitext(up.name)[0]
        prog = st.progress(0.0, "Đang bắt đầu...")
        outs = {}
        try:
            opts = {"dpi": int(dpi), "reflow": bool(reflow), "prefer_text": bool(prefer_text)}
            if not engine.startswith("🅰️"):
                db, rpt = p2w_engine_b(data, opts, lambda f, t: prog.progress(min(f, 0.99), t))
                outs["B"] = (db, rpt)
            if not engine.startswith("🅱️"):
                prog.progress(0.9, "Đang chạy pdf2docx...")
                try:
                    outs["A"] = (p2w_engine_a(data), [])
                except Exception as e:
                    st.warning(f"pdf2docx không chạy được ({e}). " + ("" if "B" in outs else "Chuyển sang bộ máy B."))
                    if "B" not in outs:
                        db, rpt = p2w_engine_b(data, opts, None)
                        outs["B"] = (db, rpt)
            prog.progress(1.0, "Hoàn tất!")
        except Exception as e:
            prog.empty()
            st.error(f"Lỗi chuyển đổi: {e}")
            return
        st.session_state["p2w_result"] = {"base": base, "outs": outs, "pdf": data}
    res = st.session_state.get("p2w_result")
    if not res:
        return
    base, outs = res["base"], res["outs"]
    for k, (d, rpt) in outs.items():
        name = "Giữ đúng vị trí" if k == "B" else "pdf2docx"
        try:
            df = p2w_text_diff(res["pdf"], d)
            if df["miss"] == 0 and df["extra"] == 0:
                st.success(f"✅ {name}: khớp 100% chữ ({df['n']} từ).")
            else:
                st.warning(f"⚠️ {name}: khớp {df['cov']*100:.2f}% chữ ({df['n'] - df['miss']}/{df['n']} từ). "
                           f"Thiếu {df['miss']} từ (có trong PDF nhưng không thấy trong Word), thừa {df['extra']} từ (có trong Word nhưng không có trong PDF).")
                with st.expander("🔎 Xem các từ cần kiểm tra trong file Word (Ctrl+F để tìm)", expanded=True):
                    for m_ in df["missing"][:30]:
                        st.markdown(f"- **Thiếu** `{m_['word']}` — PDF trang {m_['page']}: …{m_['ctx']}…")
                    for e_ in df["extra_list"][:30]:
                        st.markdown(f"- **Thừa** `{e_['word']}` — trong Word: …{e_['ctx']}…")
        except Exception:
            st.success(f"✅ {name}: đã chuyển xong.")
        for w in rpt[:10]:
            st.warning(w if isinstance(w, str) else str(w))
    if len(outs) == 1:
        k, (d, _) = next(iter(outs.items()))
        st.download_button("📥 Tải file Word (.docx)", d, file_name=f"{base}.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        zb = BytesIO()
        with zipfile.ZipFile(zb, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr(f"{base}_GiuViTri.docx", outs["B"][0])
            z.writestr(f"{base}_pdf2docx.docx", outs["A"][0])
        st.download_button("📥 Tải ZIP (2 bản Word)", zb.getvalue(), file_name=f"{base}_2ban.zip", mime="application/zip")
    st.button("🗑️ Xóa & làm file mới", on_click=clear_file, key="p2w_clear")
    st.caption("⚠️ Font lạ chưa cài trên máy sẽ được Word thay bằng font gần nhất; PDF dùng font VNI/.Vn cũ có thể sai chữ khi trích xuất.")


app_mode = st.sidebar.radio(
    "Vui lòng chọn ứng dụng:",
    [
        "📄 1. PDF sang Word (Online)", 
        "🖨️ 2. Chuyển PDF về khổ A4", 
        "📊 3. PDF/Ảnh sang Excel", 
        "🔍 4. So sánh Văn bản / Hợp đồng",
        "✂️ 5. Cắt & Ghép PDF",
        "💻 6. Chuyên gia Công thức & VBA",
        "🗜️ 7. Nén & Đóng Gói ZIP Gửi Mail",
        "🔄 8. Xoay Trang PDF",
        "🛡️ 9. Xóa Siêu Dữ Liệu Bảo Mật",
        "✒️ 10. Tách Nền Chữ Ký",
        "🔳 11. Tạo Mã QR",
        "📝 12. PDF sang Word (Offline)"
    ],
    label_visibility="collapsed"
)
st.sidebar.markdown("---") 

st.sidebar.markdown("""
    <div class="custom-alert">
        <b>⚠️ NGUYÊN TẮC SỬ DỤNG:</b><br>
        <b>- Bảo mật:</b> KHÔNG tải lên tài liệu MẬT, TỐI MẬT, dữ liệu tài chính chưa công khai.<br>
        <b>- Tối ưu:</b> Chỉ tải file PDF <b>dưới 30 trang/lần</b>.
    </div>
""", unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.markdown("""
    <div style="text-align: center; font-size: 12px; color: #666; margin-top: 10px; line-height: 1.5;">
        <b>BỘ CÔNG CỤ VĂN PHÒNG</b><br>
        Phiên bản: 1.0.0<br>
        Phát triển: Vũ Huỳnh Bình Minh<br>
        Nhân viên nghiệp vụ - Phòng Nghiệp vụ<br>
        © 2026 NMTLBT Craven "A"
    </div>
""", unsafe_allow_html=True)


if st.session_state.get("username") == "admin":
    render_gemini_status()

# ==========================================
# 4. CÁC HÀM XỬ LÝ CHỨC NĂNG CỦA ỨNG DỤNG
# ==========================================
def parse_and_add_runs(paragraph, text):
    parts_bold = re.split(r'\*\*(.*?)\*\*', text)
    for i, p_bold in enumerate(parts_bold):
        is_bold = (i % 2 == 1)
        parts_italic = re.split(r'\*(.*?)\*', p_bold)
        for j, p_italic in enumerate(parts_italic):
            is_italic = (j % 2 == 1)
            parts_underline = re.split(r'<u>(.*?)</u>', p_italic)
            for k, p_underline in enumerate(parts_underline):
                is_underline = (k % 2 == 1)
                if p_underline:
                    run = paragraph.add_run(p_underline)
                    run.bold = is_bold
                    run.italic = is_italic
                    run.underline = is_underline

def clean_tags_and_align(text, paragraph, default_align):
    align = default_align
    text_clean = text.strip()
    
    if '[CENTER]' in text_clean:
        align = WD_ALIGN_PARAGRAPH.CENTER
        text_clean = text_clean.replace('[CENTER]', '').strip()
    if '[RIGHT]' in text_clean:
        align = WD_ALIGN_PARAGRAPH.RIGHT
        text_clean = text_clean.replace('[RIGHT]', '').strip()
    if '[LEFT]' in text_clean:
        align = WD_ALIGN_PARAGRAPH.LEFT
        text_clean = text_clean.replace('[LEFT]', '').strip()
        
    paragraph.alignment = align
    return text_clean

def app_pdf_to_word():
    api_key_input = require_gemini_key()

    st.title("📄 Ứng dụng Chuyển đổi PDF & Ảnh sang Word")
    
    uploaded_file = st.file_uploader("Tải lên file ảnh (JPG, PNG) hoặc PDF:", type=["jpg", "jpeg", "png", "pdf"], key=f"app1_{st.session_state.uploader_key}")

    if uploaded_file is not None:
        st.success(f"Đã tải lên file: **{uploaded_file.name}**")
        
        if st.button("🚀 Bắt đầu Chuyển đổi", type="primary"):
            with st.spinner("🤖 Đang phân tích lề và vẽ lại bảng biểu phức tạp..."):
                try:
                    temp_input_path = f"temp_{uploaded_file.name}"
                    with open(temp_input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    client = genai.Client(api_key=api_key_input)
                    with open(temp_input_path, "rb") as f:
                        file_bytes = f.read()
                    mime_type = "application/pdf" if uploaded_file.name.endswith(".pdf") else "image/jpeg"
                    
                    prompt = """
                    NHIỆM VỤ OCR - BẮT BUỘC TUÂN THỦ NGHIÊM NGẶT CÁC QUY TẮC SAU:

                    1. THỂ THỨC VĂN BẢN (QUỐC HIỆU & CHỮ KÝ):
                       - Phần Quốc hiệu trên cùng (hoặc các khối chữ ký ngang nhau): BẮT BUỘC dùng [HEADER_TABLE] (bảng tàng hình 2 cột) để chia tỷ lệ.
                       - Không dùng bảng cho Footer.

                    2. CANH LỀ ĐOẠN VĂN:
                       - Ghi [CENTER] ở đầu dòng cần canh giữa.
                       - Ghi [RIGHT] ở đầu dòng lệch phải.

                    3. ĐỊNH DẠNG CHỮ: 
                       - In đậm -> bọc trong ** (VD: **THÔNG BÁO**). 
                       - In nghiêng -> bọc trong * (VD: *Nơi nhận:*).
                       - Gạch chân -> bọc trong <u> và </u>.

                    4. BẢNG BIỂU PHỨC TẠP (QUAN TRỌNG NHẤT):
                       - Bắt buộc vẽ bằng Markdown (|...|).
                       - TẤT CẢ các dòng trong cùng 1 bảng PHẢI CÓ SỐ LƯỢNG CỘT BẰNG NHAU (cùng số lượng ký tự `|`).
                       - XỬ LÝ Ô GỘP (MERGED CELLS) TRONG BẢN GỐC:
                         + Gộp ngang (kéo dài sang phải): Ghi chữ vào ô trái cùng, các ô bị gộp phía sau điền CHÍNH XÁC chữ: `[MERGE_LEFT]`
                         + Gộp dọc (kéo dài xuống dưới): Ghi chữ vào ô trên cùng, các ô bị gộp phía dưới điền CHÍNH XÁC chữ: `[MERGE_UP]`
                       - Ví dụ bảng có tiêu đề gộp:
                         | STT | TÊN VĂN BẢN | TÌNH TRẠNG HỒ SƠ | [MERGE_LEFT] | [MERGE_LEFT] | [MERGE_LEFT] |
                         | [MERGE_UP] | [MERGE_UP] | Có | [MERGE_LEFT] | [MERGE_LEFT] | Không có |
                         | [MERGE_UP] | [MERGE_UP] | Bản chính | Bản sao | Bản photo | [MERGE_UP] |
                       - TUYỆT ĐỐI không để ô trống làm lệch số lượng cột. Dùng `<br>` để xuống dòng trong ô.

                    5. KHÔNG sinh ra mã phân trang (Start of Page). KHÔNG dùng mã LaTeX toán học.
                    """

                    response = call_gemini_smart(
                        client, 
                        [types.Part.from_bytes(data=file_bytes, mime_type=mime_type), prompt]
                    )
                    
                    is_landscape = False
                    try:
                        if uploaded_file.name.lower().endswith('.pdf'):
                            pdf_reader = PdfReader(BytesIO(file_bytes))
                            first_page = pdf_reader.pages[0]
                            w = float(first_page.mediabox.width)
                            h = float(first_page.mediabox.height)
                            is_landscape = w > h
                        else:
                            img = Image.open(BytesIO(file_bytes))
                            is_landscape = img.width > img.height
                    except Exception as e:
                        is_landscape = False
                    
                    doc = Document()
                    section = doc.sections[0]
                    
                    if is_landscape:
                        section.orientation = WD_ORIENT.LANDSCAPE
                        section.page_width = Mm(297)
                        section.page_height = Mm(210)
                    else:
                        section.orientation = WD_ORIENT.PORTRAIT
                        section.page_width = Mm(210)
                        section.page_height = Mm(297)

                    section.top_margin = Mm(20)
                    section.bottom_margin = Mm(20)
                    section.left_margin = Mm(30)
                    section.right_margin = Mm(20)

                    style = doc.styles['Normal']
                    font = style.font
                    font.name = 'Times New Roman'
                    font.size = Pt(13)
                    
                    def build_docx_table(doc_obj, buffer, is_header_table=False):
                        if not buffer: return
                        num_cols = max(len(row) for row in buffer)
                        
                        normalized_buffer = []
                        for row in buffer:
                            new_row = list(row)
                            while len(new_row) < num_cols:
                                new_row.append('')
                            normalized_buffer.append(new_row)
                            
                        current_table = doc_obj.add_table(rows=len(normalized_buffer), cols=num_cols)
                        
                        if not is_header_table:
                            current_table.style = 'Table Grid'
                            current_table.autofit = True
                        else:
                            current_table.autofit = False

                        for row_idx, row_data in enumerate(normalized_buffer):
                            row_cells = current_table.rows[row_idx].cells
                            
                            if is_header_table and num_cols == 2:
                                row_cells[0].width = Cm(6.0)
                                row_cells[1].width = Cm(10.0)

                            for col_idx, cell_data in enumerate(row_data):
                                cell_text = cell_data.strip()
                                if cell_text not in ['[MERGE_LEFT]', '[MERGE_UP]']:
                                    cell = row_cells[col_idx]
                                    cell.text = ""
                                    cell_lines = cell_text.split('<br>')
                                    for idx, c_line in enumerate(cell_lines):
                                        p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
                                        p.paragraph_format.space_after = Pt(0)
                                        
                                        default_al = WD_ALIGN_PARAGRAPH.CENTER if (is_header_table or row_idx == 0) else WD_ALIGN_PARAGRAPH.LEFT
                                        clean_text = clean_tags_and_align(c_line.strip(), p, default_al)
                                        parse_and_add_runs(p, clean_text)
                                        
                        for row_idx in range(len(normalized_buffer)):
                            for col_idx in range(num_cols):
                                cell_text = normalized_buffer[row_idx][col_idx].strip()
                                try:
                                    if cell_text == '[MERGE_LEFT]' and col_idx > 0:
                                        current_table.cell(row_idx, col_idx - 1).merge(current_table.cell(row_idx, col_idx))
                                    elif cell_text == '[MERGE_UP]' and row_idx > 0:
                                        current_table.cell(row_idx - 1, col_idx).merge(current_table.cell(row_idx, col_idx))
                                except Exception:
                                    pass

                    response_text = response.text
                    table_buffer = []
                    is_next_table_header = False

                    for line in response_text.split('\n'):
                        line_stripped = re.sub(r'<td[^>]*>', '', line.strip())
                        line_stripped = re.sub(r'</td>', '', line_stripped)
                        
                        line_stripped = re.sub(r'==\s*(Start|End)\s+of\s+Page.*?==', '', line_stripped, flags=re.IGNORECASE).strip()
                        line_stripped = line_stripped.replace('$\ge$', '≥').replace('$\rightarrow$', '→').replace('$\le$', '≤').replace('$\Rightarrow$', '⇒').replace('$\leftarrow$', '←')

                        if not line_stripped or line_stripped.startswith("```"):
                            continue

                        if line_stripped == '[HEADER_TABLE]':
                            is_next_table_header = True
                            continue

                        if line_stripped.startswith('|') and line_stripped.endswith('|'):
                            if re.match(r'^[\s\|\-:]+$', line_stripped):
                                continue
                            
                            cells_data = [cell.strip() for cell in line_stripped.split('|')][1:-1]
                            table_buffer.append(cells_data)
                        else:
                            if table_buffer:
                                if all(cell == '' for cell in table_buffer[0]):
                                    table_buffer.pop(0)
                                build_docx_table(doc, table_buffer, is_header_table=is_next_table_header)
                                table_buffer = []
                                is_next_table_header = False
                            
                            if line_stripped.startswith('---'):
                                continue
                                
                            if line_stripped:
                                p = doc.add_paragraph()
                                p.paragraph_format.space_after = Pt(6)
                                clean_line = clean_tags_and_align(line_stripped, p, WD_ALIGN_PARAGRAPH.JUSTIFY)
                                parse_and_add_runs(p, clean_line)

                    if table_buffer:
                        if all(cell == '' for cell in table_buffer[0]):
                            table_buffer.pop(0)
                        build_docx_table(doc, table_buffer, is_header_table=is_next_table_header)

                    output_docx_path = "ket_qua.docx"
                    doc.save(output_docx_path)

                    st.success("🎉 Chuyển đổi thành công! Bảng biểu phức tạp đã được gộp ô chính xác.")

                    with open(output_docx_path, "rb") as file_download:
                        st.download_button(
                            label="📥 Tải xuống file Word (.docx)",
                            data=file_download,
                            file_name=f"Converted_{uploaded_file.name.split('.')[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            on_click=clear_file
                        )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi: {e}")

def app_number_2():
    st.title("🖨️ Chuyển PDF về khổ A4")
    st.markdown("Xóa bỏ mọi khung ẩn của bản vẽ cũ, ép lại chính xác thành khổ A4 tiêu chuẩn.")

    uploaded_pdf = st.file_uploader("Tải lên bản vẽ PDF cần xử lý:", type=["pdf"], key=f"app2_{st.session_state.uploader_key}")

    if uploaded_pdf is not None:
        st.success(f"Đã tải lên file: **{uploaded_pdf.name}**")
        
        if st.button("📏 Chuyển thành A4", type="primary"):
            with st.spinner("Đang truy quét và ghi đè các khung viền ẩn..."):
                try:
                    reader = PdfReader(uploaded_pdf)
                    writer = PdfWriter()

                    A4_W = 595.276
                    A4_H = 841.890

                    for page in reader.pages:
                        orig_w = float(page.mediabox.width)
                        orig_h = float(page.mediabox.height)

                        is_landscape = orig_w > orig_h
                        target_w = A4_H if is_landscape else A4_W
                        target_h = A4_W if is_landscape else A4_H

                        scale_w = target_w / orig_w
                        scale_h = target_h / orig_h
                        scale_factor = min(scale_w, scale_h)

                        scaled_w = orig_w * scale_factor
                        scaled_h = orig_h * scale_factor
                        tx = (target_w - scaled_w) / 2.0
                        ty = (target_h - scaled_h) / 2.0

                        op = Transformation().scale(sx=scale_factor, sy=scale_factor).translate(tx=tx, ty=ty)
                        page.add_transformation(op)

                        page.mediabox.lower_left = (0, 0)
                        page.mediabox.upper_right = (target_w, target_h)
                        page.cropbox.lower_left = (0, 0)
                        page.cropbox.upper_right = (target_w, target_h)
                        
                        if "/BleedBox" in page:
                            page.bleedbox.lower_left = (0, 0)
                            page.bleedbox.upper_right = (target_w, target_h)
                        if "/TrimBox" in page:
                            page.trimbox.lower_left = (0, 0)
                            page.trimbox.upper_right = (target_w, target_h)
                        if "/ArtBox" in page:
                            page.artbox.lower_left = (0, 0)
                            page.artbox.upper_right = (target_w, target_h)

                        writer.add_page(page)

                    output_path = f"A4_Chuan_{uploaded_pdf.name}"
                    with open(output_path, "wb") as f:
                        writer.write(f)

                    st.success("🎉 Xử lý thành công! Toàn bộ khung hình đã được chuyển thành A4.")

                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 Tải xuống bản vẽ A4 chuẩn (.pdf)",
                            data=f,
                            file_name=output_path,
                            mime="application/pdf",
                            on_click=clear_file
                        )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi: {e}")

def parse_rich_text(text_val, font_name="Times New Roman", size=12):
    text_str = str(text_val) if text_val is not None else ""
    parts = re.split(r'\*\*(.*?)\*\*', text_str)

    if len(parts) == 1:
        return text_str 

    rt = CellRichText()
    font_normal = InlineFont(rFont=font_name, sz=size)
    font_bold = InlineFont(rFont=font_name, sz=size, b=True)

    for i, p in enumerate(parts):
        if not p: continue
        if i % 2 == 1: 
            rt.append(TextBlock(font=font_bold, text=p))
        else: 
            rt.append(TextBlock(font=font_normal, text=p))
    return rt

def app_number_3():
    api_key_input = require_gemini_key()

    st.title("📊 Bóc tách PDF/Ảnh sang Excel (Chuẩn A4 & Giữ Định Dạng)")
    st.markdown("Trích xuất và tự động định dạng giống PDF gốc, ép sẵn khổ in **A4 (Ngang/Dọc tự động)**.")

    uploaded_excel_file = st.file_uploader("Tải lên tài liệu (Ảnh hoặc PDF):", type=["jpg", "jpeg", "png", "pdf"], key=f"app3_{st.session_state.uploader_key}")

    if uploaded_excel_file is not None:
        st.success(f"Đã tải lên file: **{uploaded_excel_file.name}**")
        
        if st.button("🚀 Trích xuất ra Excel", type="primary"):
            with st.spinner("🤖 Vui lòng đợi..."):
                try:
                    temp_input_path = f"temp_excel_{uploaded_excel_file.name}"
                    with open(temp_input_path, "wb") as f:
                        f.write(uploaded_excel_file.getbuffer())

                    client = genai.Client(api_key=api_key_input)
                    with open(temp_input_path, "rb") as f:
                        file_bytes = f.read()
                    mime_type = "application/pdf" if uploaded_excel_file.name.endswith(".pdf") else "image/jpeg"
                    
                    prompt = """
                    Bạn là một chuyên gia số hóa tài liệu. Hãy đọc kỹ tài liệu và trả về kết quả DUY NHẤT dưới dạng chuỗi JSON hợp lệ. Không trả lời thêm.
                    
                    LƯU Ý QUAN TRỌNG VỀ ĐỊNH DẠNG:
                    Hãy phân tích và BỌC CÁC CHỮ ĐƯỢC IN ĐẬM trong bản gốc bằng dấu sao kép (**). 
                    Ví dụ: "- **Thời gian:** 9 giờ ngày 15 tháng 7 năm 2026"
                    
                    Cấu trúc JSON BẮT BUỘC:
                    {
                      "title": "Dòng tiêu đề trên cùng",
                      "info_lines": [
                        "Dòng thông tin 1",
                        "Dòng thông tin 2"
                      ],
                      "headers": ["Cột 1", "Cột 2", "Cột 3"],
                      "rows": [
                        ["Dữ liệu 1", "Dữ liệu 2", "Dữ liệu 3"],
                        ["Dữ liệu 1", "Dữ liệu 2", "Dữ liệu 3"]
                      ]
                    }
                    """

                    response = call_gemini_smart(
                        client, 
                        [types.Part.from_bytes(data=file_bytes, mime_type=mime_type), prompt]
                    )
                    
                    is_landscape = False
                    try:
                        if uploaded_excel_file.name.lower().endswith('.pdf'):
                            pdf_reader = PdfReader(BytesIO(file_bytes))
                            first_page = pdf_reader.pages[0]
                            w = float(first_page.mediabox.width)
                            h = float(first_page.mediabox.height)
                            is_landscape = w > h
                        else:
                            img = Image.open(BytesIO(file_bytes))
                            is_landscape = img.width > img.height
                    except Exception as e:
                        is_landscape = False

                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"):
                        raw_text = raw_text[7:]
                    if raw_text.startswith("```"):
                        raw_text = raw_text[3:]
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3]
                    raw_text = raw_text.strip()
                    
                    data = json.loads(raw_text)

                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Danh_Sach"

                    ws.page_setup.paperSize = ws.PAPERSIZE_A4
                    if is_landscape:
                        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                    else:
                        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
                    ws.print_options.horizontalCentered = True 
                    
                    font_title = Font(name="Times New Roman", size=14, bold=True)
                    font_bold = Font(name="Times New Roman", size=12, bold=True)
                    font_normal = Font(name="Times New Roman", size=12)
                    align_center = Alignment(horizontal="center", vertical="center")
                    align_left = Alignment(horizontal="left", vertical="center")
                    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                    current_row = 1
                    total_cols = len(data.get("headers", [1,2,3,4,5]))

                    title = data.get("title", "")
                    if title:
                        clean_title = str(title).replace('**', '').upper()
                        cell = ws.cell(row=current_row, column=1, value=clean_title)
                        cell.font = font_title
                        cell.alignment = align_center
                        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=total_cols)
                        current_row += 1

                    for info in data.get("info_lines", []):
                        cell = ws.cell(row=current_row, column=1)
                        rich_val = parse_rich_text(info)
                        
                        cell.value = rich_val
                        if isinstance(rich_val, str): 
                            cell.font = font_normal
                            
                        cell.alignment = align_left
                        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=total_cols)
                        current_row += 1

                    current_row += 1 

                    headers = data.get("headers", [])
                    for col_idx, header in enumerate(headers, 1):
                        clean_header = str(header).replace('**', '')
                        cell = ws.cell(row=current_row, column=col_idx, value=clean_header)
                        cell.font = font_bold
                        cell.alignment = align_center
                        cell.border = thin_border
                    current_row += 1

                    for row_data in data.get("rows", []):
                        for col_idx, val in enumerate(row_data, 1):
                            cell = ws.cell(row=current_row, column=col_idx)
                            rich_val = parse_rich_text(val)
                            
                            cell.value = rich_val
                            if isinstance(rich_val, str):
                                cell.font = font_normal
                                
                            cell.border = thin_border
                            
                            if col_idx == 1 or col_idx == total_cols:
                                cell.alignment = align_center
                            else:
                                cell.alignment = align_left
                        current_row += 1

                    ws.column_dimensions['A'].width = 8   
                    ws.column_dimensions['B'].width = 25  
                    ws.column_dimensions['C'].width = 30  
                    ws.column_dimensions['D'].width = 25  
                    ws.column_dimensions['E'].width = 15  

                    output = BytesIO()
                    wb.save(output)
                    processed_data = output.getvalue()

                    st.success("🎉 Đã xuất bảng Excel thành công! Khổ giấy in đã được set tự động xoay ngang/dọc.")

                    st.download_button(
                        label="📥 Tải xuống Excel Chuẩn Format (.xlsx)",
                        data=processed_data,
                        file_name=f"Excel_Chuan_{uploaded_excel_file.name.split('.')[0]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        on_click=clear_file
                    )

                except Exception as e:
                    st.error(f"Đã xảy ra lỗi hệ thống: {e}")

def extract_text_from_file(uploaded_file, client):
    text = ""
    file_ext = uploaded_file.name.split('.')[-1].lower()
    
    if file_ext == "docx":
        doc = Document(uploaded_file)
        text = "\n".join([para.text for para in doc.paragraphs])
    
    elif file_ext == "pdf":
        try:
            reader = PdfReader(uploaded_file)
            extracted = "\n".join([p.extract_text() or "" for p in reader.pages]).strip()
            if len(extracted) > 50: 
                uploaded_file.seek(0) 
                return extracted
        except Exception:
            pass
            
        uploaded_file.seek(0) 
        file_bytes = uploaded_file.getbuffer()
        prompt = "Hãy trích xuất TOÀN BỘ nội dung văn bản trong tài liệu này một cách chính xác nhất. Trả về plain text."
        response = call_gemini_smart(
            client, 
            [types.Part.from_bytes(data=file_bytes, mime_type="application/pdf"), prompt]
        )
        text = response.text
        
    elif file_ext in ["png", "jpg", "jpeg"]:
        file_bytes = uploaded_file.getbuffer()
        prompt = "Hãy trích xuất TOÀN BỘ nội dung văn bản trong tài liệu này một cách chính xác nhất. Trả về plain text."
        response = call_gemini_smart(
            client, 
            [types.Part.from_bytes(data=file_bytes, mime_type=f"image/{file_ext}"), prompt]
        )
        text = response.text
    
    return text

def app_document_compare():
    api_key_input = require_gemini_key()

    st.title("🔍 So sánh Văn bản / Hợp đồng")
    st.markdown("Đối soát độ lệch chữ giữa 2 phiên bản tài liệu. Hỗ trợ đối chiếu chéo **Word vs Word**, **PDF scan vs Word**.")

    col1, col2 = st.columns(2)
    with col1:
        st.info("📄 BẢN GỐC (V1)")
        file_v1 = st.file_uploader("Tải lên bản gốc (PDF/Word/Ảnh):", key=f"v1_{st.session_state.uploader_key}")
    with col2:
        st.warning("📝 BẢN ĐỐI TÁC GỬI LẠI (V2)")
        file_v2 = st.file_uploader("Tải lên bản chỉnh sửa (PDF/Word/Ảnh):", key=f"v2_{st.session_state.uploader_key}")

    if file_v1 and file_v2:
        if st.button("🔎 Soi Sự Khác Biệt", type="primary"):
            with st.spinner("🤖 Đang quét tài liệu và so sánh từng ký tự..."):
                try:
                    client = genai.Client(api_key=api_key_input)
                    
                    st.toast("Đang đọc và giải mã Bản Gốc (V1)...")
                    text_v1 = extract_text_from_file(file_v1, client)
                    
                    st.toast("Đang đọc và giải mã Bản Đối Tác (V2)...")
                    text_v2 = extract_text_from_file(file_v2, client)

                    words_v1 = text_v1.splitlines() 
                    words_v2 = text_v2.splitlines()
                    
                    words_v1_flat = [word for line in words_v1 for word in line.split(" ") if word.strip()]
                    words_v2_flat = [word for line in words_v2 for word in line.split(" ") if word.strip()]

                    matcher = difflib.SequenceMatcher(None, words_v1_flat, words_v2_flat)
                    
                    edits = []
                    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                        if tag != 'equal':
                            edits.append({
                                'tag': tag,
                                'old': " ".join(words_v1_flat[i1:i2]),
                                'new': " ".join(words_v2_flat[j1:j2])
                            })
                    
                    total_edits = len(edits)

                    st.toast("AI đang đánh giá rủi ro các điểm thay đổi...")
                    summary_prompt = f"""
                    Tôi có 2 phiên bản hợp đồng. Phát hiện tổng cộng {total_edits} lần thay đổi.
                    Bản gốc: {text_v1[:3000]}... 
                    Bản sửa: {text_v2[:3000]}...
                    
                    Hãy viết một đoạn văn ngắn gọn đánh giá rủi ro pháp lý/ngữ nghĩa của các điểm khác biệt này.
                    BẮT BUỘC TUÂN THỦ: 
                    1. TUYỆT ĐỐI KHÔNG dùng ký tự đặc biệt như dấu sao (*), thăng (#), gạch ngang (-), hay bảng biểu (|). 
                    2. Trả lời bằng văn bản thuần túy (plain text) để đưa trực tiếp vào Word.
                    """
                    summary_response = call_gemini_smart(
                        client, 
                        [summary_prompt]
                    )
                    ai_text = summary_response.text.replace('*', '').replace('#', '').replace('`', '').strip()

                    doc_report = Document()
                    
                    style = doc_report.styles['Normal']
                    style.font.name = 'Times New Roman'
                    style.font.size = Pt(13)

                    p_title = doc_report.add_heading(level=0)
                    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_title = p_title.add_run('BÁO CÁO ĐỐI CHIẾU TÀI LIỆU (V1 vs V2)')
                    r_title.font.name = 'Times New Roman'
                    r_title.font.size = Pt(14)
                    
                    p_h1 = doc_report.add_heading(level=1)
                    r_h1 = p_h1.add_run('I. Bảng Tổng kết chi tiết các lần chỉnh sửa')
                    r_h1.font.name = 'Times New Roman'
                    r_h1.font.size = Pt(12)

                    p_total = doc_report.add_paragraph()
                    run_total = p_total.add_run(f"Hệ thống ghi nhận tổng cộng {total_edits} lần chỉnh sửa từ đối tác.")
                    run_total.bold = True
                    run_total.font.color.rgb = RGBColor(255, 0, 0)
                    
                    if total_edits > 0:
                        table = doc_report.add_table(rows=1, cols=4)
                        table.style = 'Table Grid'
                        hdr_cells = table.rows[0].cells
                        hdr_cells[0].text = 'STT'
                        hdr_cells[1].text = 'Thao tác'
                        hdr_cells[2].text = 'Bản gốc (V1)'
                        hdr_cells[3].text = 'Đối tác sửa (V2)'
                        
                        for cell in hdr_cells:
                            for p in cell.paragraphs:
                                for r in p.runs:
                                    r.font.name = 'Times New Roman'
                                    r.bold = True
                        
                        for idx, edit in enumerate(edits, 1):
                            row_cells = table.add_row().cells
                            row_cells[0].text = str(idx)
                            
                            if edit['tag'] == 'delete':
                                row_cells[1].text = 'Xóa bỏ'
                                row_cells[2].text = edit['old']
                                row_cells[3].text = "(Đã xóa)"
                            elif edit['tag'] == 'insert':
                                row_cells[1].text = 'Thêm mới'
                                row_cells[2].text = "(Không có)"
                                row_cells[3].text = edit['new']
                            elif edit['tag'] == 'replace':
                                row_cells[1].text = 'Thay thế'
                                row_cells[2].text = edit['old']
                                row_cells[3].text = edit['new']

                    p_h2 = doc_report.add_heading(level=1)
                    r_h2 = p_h2.add_run('II. Nhận định rủi ro tổng quan (AI)')
                    r_h2.font.name = 'Times New Roman'
                    r_h2.font.size = Pt(12)
                    
                    doc_report.add_paragraph(ai_text)

                    p_h3 = doc_report.add_heading(level=1)
                    r_h3 = p_h3.add_run('III. Chi tiết văn bản (Kính lúp bôi màu)')
                    r_h3.font.name = 'Times New Roman'
                    r_h3.font.size = Pt(12)
                    
                    legend = doc_report.add_paragraph()
                    run_del = legend.add_run("Chữ màu đỏ có gạch ngang: Bị đối tác xóa bỏ\n")
                    run_del.font.color.rgb = RGBColor(255, 0, 0)
                    run_del.font.strike = True
                    run_add = legend.add_run("Chữ màu xanh có gạch chân: Được đối tác thêm mới")
                    run_add.font.color.rgb = RGBColor(0, 0, 255)
                    run_add.font.underline = True

                    p_diff = doc_report.add_paragraph()
                    
                    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                        if tag == 'equal':
                            p_diff.add_run(" " + " ".join(words_v1_flat[i1:i2]))
                        elif tag == 'delete':
                            run = p_diff.add_run(" " + " ".join(words_v1_flat[i1:i2]))
                            run.font.color.rgb = RGBColor(255, 0, 0)
                            run.font.strike = True
                        elif tag == 'insert':
                            run = p_diff.add_run(" " + " ".join(words_v2_flat[j1:j2]))
                            run.font.color.rgb = RGBColor(0, 0, 255)
                            run.font.underline = True
                        elif tag == 'replace':
                            run_old = p_diff.add_run(" " + " ".join(words_v1_flat[i1:i2]))
                            run_old.font.color.rgb = RGBColor(255, 0, 0)
                            run_old.font.strike = True
                            
                            run_new = p_diff.add_run(" " + " ".join(words_v2_flat[j1:j2]))
                            run_new.font.color.rgb = RGBColor(0, 0, 255)
                            run_new.font.underline = True

                    output_report_path = "Bao_Cao_So_Sanh.docx"
                    doc_report.save(output_report_path)

                    st.success("🎉 So sánh hoàn tất! Báo cáo đã được định dạng chuẩn hành chính.")

                    with open(output_report_path, "rb") as file_download:
                        st.download_button(
                            label="📥 Tải xuống Báo cáo chi tiết (.docx)",
                            data=file_download,
                            file_name="Bao_Cao_So_Sanh_V1_vs_V2.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            on_click=clear_file
                        )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi: {e}")

# ==========================================
# APP 5: CẮT, TRÍCH XUẤT TỔNG HỢP & GHÉP PDF
# ==========================================
def app_pdf_split_merge():
    import fitz
    from PIL import Image
    from io import BytesIO
    import zipfile

    st.title("✂️ Cắt, Trích Xuất & Ghép Nối PDF")
    st.markdown("Công cụ thao tác PDF toàn diện: Cắt trang trực quan qua ảnh xem trước, rã file hàng loạt (Burst), và ghép nối với khả năng tùy chỉnh thứ tự file linh hoạt.")

    # CSS căn giữa checkbox: căn ở KHUNG CHA (stElementContainer), không chỉ ở bên trong checkbox
    st.markdown("""
        <style>
        /* Khung cha chứa checkbox: căn giữa */
        div[data-testid="stElementContainer"]:has(div[data-testid="stCheckbox"]),
        div.element-container:has(div[data-testid="stCheckbox"]) {
            display: flex !important;
            justify-content: center !important;
            width: 100% !important;
        }
        /* Checkbox chỉ rộng vừa nội dung để được căn giữa */
        div[data-testid="stCheckbox"] {
            width: auto !important;
        }
        div[data-testid="stCheckbox"] label {
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
        }
        /* In đậm nhãn của checkbox */
        div[data-testid="stCheckbox"] label p {
            font-weight: bold !important;
        }
        </style>
    """, unsafe_allow_html=True)

    tab_split, tab_merge = st.tabs(["✂️ Cắt Trang & Tách File (Visual Split)", "📑 Ghép Nhiều File PDF (Reorder Merge)"])

    # ----------------------------------------------------
    # TAB 1: CẮT TRANG VÀ TÁCH FILE TRỰC QUAN (GỘP ĐA FILE)
    # ----------------------------------------------------
    with tab_split:
        st.subheader("1. Chọn các file PDF cần trích xuất")
        uploaded_split_files = st.file_uploader(
            "Tải lên một hoặc nhiều file PDF (Hệ thống sẽ trải tất cả các trang ra màn hình):",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"app5_split_{st.session_state.uploader_key}"
        )

        if uploaded_split_files:
            try:
                # Gộp ngầm tất cả các file vào 1 document trong bộ nhớ để trải ra màn hình
                combined_doc = fitz.open()
                for f in uploaded_split_files:
                    temp_doc = fitz.open("pdf", f.getvalue())
                    combined_doc.insert_pdf(temp_doc)
                
                total_pages = len(combined_doc)
                st.info(f"📁 Đã tải lên **{len(uploaded_split_files)} file**. Trải dài tổng cộng: **{total_pages} trang**.")

                # Tạo key lưu trạng thái dựa trên danh sách file để không bị loạn khi đổi file
                sig = "|".join([f.name for f in uploaded_split_files])
                state_check_key = f"split_selected_{sig}_{st.session_state.uploader_key}"
                
                if state_check_key not in st.session_state:
                    # MẶC ĐỊNH LÀ FALSE (Không chọn trang nào cả)
                    st.session_state[state_check_key] = [False] * total_pages
                    for i in range(total_pages):
                        st.session_state[f"chk_page_{i}_{state_check_key}"] = False

                st.markdown("---")
                col_view1, col_view2 = st.columns([2, 1])
                with col_view1:
                    st.markdown("### 👁️ Danh sách trang tổng hợp")
                    st.caption("Tick chọn những trang bạn muốn giữ lại để trích xuất thành 1 file duy nhất.")
                with col_view2:
                    grid_cols = st.slider("🔍 Cỡ ảnh (Số trang / hàng):", min_value=1, max_value=4, value=2, help="Kéo về 1 hoặc 2 để phóng to trang giấy.")

                zoom_mat = 0.8 if grid_cols <= 2 else 0.4
                mat = fitz.Matrix(zoom_mat, zoom_mat)
                cols = st.columns(grid_cols)

                for idx in range(total_pages):
                    c_idx = idx % grid_cols
                    page = combined_doc[idx]
                    with cols[c_idx]:
                        pix = page.get_pixmap(matrix=mat)
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        st.image(img, use_container_width=True)

                        st.markdown(f"<div style='text-align: center; color: red; font-size: 24px; font-weight: 900; margin-top: 5px; margin-bottom: 0px;'>Trang {idx + 1}</div>", unsafe_allow_html=True)

                        st.session_state[state_check_key][idx] = st.checkbox(
                            "Chọn trang này",
                            key=f"chk_page_{idx}_{state_check_key}"
                        )
                        st.markdown("<br>", unsafe_allow_html=True)

                st.markdown("---")
                col_act1, col_act2 = st.columns(2)

                with col_act1:
                    st.markdown("#### 🎯 Trích xuất trang đã chọn")
                    selected_indices = [i for i, val in enumerate(st.session_state[state_check_key]) if val]
                    st.caption(f"Đang chọn **{len(selected_indices)}/{total_pages}** trang.")

                    if st.button("🚀 Xuất 1 PDF Đã Chọn", type="primary", disabled=(len(selected_indices) == 0)):
                        with st.spinner("Đang trích xuất trang đã chọn..."):
                            out_doc = fitz.open()
                            out_doc.insert_pdf(combined_doc, from_page=0, to_page=total_pages - 1)
                            out_doc.select(selected_indices)
                            out_stream = BytesIO()
                            out_doc.save(out_stream, garbage=3, deflate=True)
                            
                            st.success(f"🎉 Đã trích xuất thành công {len(selected_indices)} trang thành 1 file duy nhất!")
                            st.download_button(
                                label="📥 Tải xuống PDF Đã Cắt",
                                data=out_stream.getvalue(),
                                file_name="Tai_Lieu_Trich_Xuat_Tong_Hop.pdf",
                                mime="application/pdf",
                                on_click=clear_file
                            )

                with col_act2:
                    st.markdown("#### ⚡ Tách rời toàn bộ trang (Burst)")
                    st.caption("Rã từng trang thành các file PDF riêng biệt (ZIP).")
                    if st.button("📦 Rã Từng Trang Sang ZIP"):
                        with st.spinner(f"Đang bóc tách {total_pages} trang thành các file riêng..."):
                            zip_buffer = BytesIO()
                            with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                                for i in range(total_pages):
                                    single_page_doc = fitz.open()
                                    single_page_doc.insert_pdf(combined_doc, from_page=i, to_page=i)
                                    p_stream = BytesIO()
                                    single_page_doc.save(p_stream, garbage=3, deflate=True)
                                    zf.writestr(f"Trang_{i+1:03d}.pdf", p_stream.getvalue())

                            st.success("🎉 Đã rã toàn bộ trang và đóng gói ZIP thành công!")
                            st.download_button(
                                label="📥 Tải xuống File ZIP",
                                data=zip_buffer.getvalue(),
                                file_name="Tach_Roi_Tong_Hop.zip",
                                mime="application/zip",
                                on_click=clear_file
                            )

            except Exception as e:
                st.error(f"Lỗi khi xử lý file PDF: {e}")

    # ----------------------------------------------------
    # TAB 2: GHÉP NHIỀU FILE VỚI TÍNH NĂNG ĐỔI THỨ TỰ (REORDER)
    # ----------------------------------------------------
    with tab_merge:
        st.subheader("2. Tải lên Khay chứa file PDF")
        uploaded_merge_files = st.file_uploader(
            "Tải lên các file PDF (Có thể tải lên nhiều lần):",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"app5_merge_{st.session_state.uploader_key}"
        )

        if uploaded_merge_files:
            merge_state_key = f"merge_order_list_{st.session_state.uploader_key}"
            sig_key = f"sig_{merge_state_key}"

            current_sig = "|".join(sorted([f.name for f in uploaded_merge_files]))
            if sig_key not in st.session_state or st.session_state[sig_key] != current_sig:
                st.session_state[merge_state_key] = list(uploaded_merge_files)
                st.session_state[sig_key] = current_sig

            st.markdown("### 📋 DANH SÁCH FILE CHỐT ĐỂ GHÉP")
            st.warning("Các file bị xóa ở danh sách dưới đây sẽ KHÔNG được ghép, mặc dù file đó vẫn còn nằm ở Khay tải lên phía trên.")
            st.info(f"Đang có **{len(st.session_state[merge_state_key])} file** sẵn sàng. Bấm mũi tên để đổi thứ tự.")

            for i, f in enumerate(st.session_state[merge_state_key]):
                col_name, col_up, col_down, col_del = st.columns([6, 1.2, 1.2, 1.2])
                with col_name:
                    file_mb = len(f.getvalue()) / (1024 * 1024)
                    st.markdown(f"**Vị trí {i+1}:** 📄 `{f.name}` *({file_mb:.2f} MB)*")
                with col_up:
                    if st.button("⬆️ Lên", key=f"up_{i}_{merge_state_key}", disabled=(i == 0)):
                        st.session_state[merge_state_key][i - 1], st.session_state[merge_state_key][i] = (
                            st.session_state[merge_state_key][i],
                            st.session_state[merge_state_key][i - 1]
                        )
                        st.rerun()
                with col_down:
                    if st.button("⬇️ Xuống", key=f"down_{i}_{merge_state_key}", disabled=(i == len(st.session_state[merge_state_key]) - 1)):
                        st.session_state[merge_state_key][i + 1], st.session_state[merge_state_key][i] = (
                            st.session_state[merge_state_key][i],
                            st.session_state[merge_state_key][i + 1]
                        )
                        st.rerun()
                with col_del:
                    if st.button("🗑️ Xóa khỏi DS", key=f"del_{i}_{merge_state_key}"):
                        st.session_state[merge_state_key].pop(i)
                        st.rerun()

            st.markdown("---")
            merge_name = st.text_input("Tên file PDF sau khi ghép nối:", value="Tai_Lieu_Ghep_Tong_Hop.pdf")
            if not merge_name.lower().endswith(".pdf"):
                merge_name += ".pdf"

            if st.button("🚀 Bắt đầu Ghép Nối", type="primary", disabled=(len(st.session_state[merge_state_key]) == 0)):
                with st.spinner("Đang tiến hành ghép nối các file PDF theo thứ tự chỉ định..."):
                    try:
                        merged_doc = fitz.open()
                        for file_obj in st.session_state[merge_state_key]:
                            sub_doc = fitz.open("pdf", file_obj.getvalue())
                            merged_doc.insert_pdf(sub_doc)

                        out_merged = BytesIO()
                        merged_doc.save(out_merged, garbage=3, deflate=True)

                        total_pages_merged = len(merged_doc)
                        merged_size_mb = len(out_merged.getvalue()) / (1024 * 1024)

                        st.success(f"🎉 Ghép thành công **{len(st.session_state[merge_state_key])} file** thành 1 file duy nhất! (Tổng cộng **{total_pages_merged} trang**, nặng **{merged_size_mb:.2f} MB**)")
                        st.download_button(
                            label=f"📥 Tải xuống {merge_name}",
                            data=out_merged.getvalue(),
                            file_name=merge_name,
                            mime="application/pdf",
                            on_click=clear_file
                        )
                    except Exception as e:
                        st.error(f"Lỗi khi ghép các file PDF: {e}")
                        # ==========================================

def app_excel_expert():
    api_key_input = require_gemini_key()

    st.title("💻 Chuyên gia Công thức & VBA Excel")
    st.markdown("Tải file lên, mô tả bài toán và nhận ngay Công thức hoặc Mã VBA chuẩn xác.")

    uploaded_file = st.file_uploader("Tải lên bảng dữ liệu (Excel/CSV) làm mẫu:", type=["csv", "xlsx", "xls"], key=f"app6_{st.session_state.uploader_key}")
    
    if uploaded_file:
        st.success(f"Đã nạp thành công: **{uploaded_file.name}**")
        
        context_str = ""
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, nrows=3)
                context_str = f"Tên file: {uploaded_file.name}\nDữ liệu mẫu (Tiêu đề và 3 dòng đầu):\n{df.to_csv(index=False)}"
            else:
                xls = pd.ExcelFile(uploaded_file)
                sheet_names = xls.sheet_names
                df = pd.read_excel(uploaded_file, sheet_name=sheet_names[0], nrows=3)
                context_str = f"Các sheets có trong file: {', '.join(sheet_names)}\nDữ liệu mẫu của sheet đầu tiên (Tiêu đề và 3 dòng đầu):\n{df.to_csv(index=False)}"
        except Exception as e:
            st.error(f"Lỗi đọc file mẫu: {e}")
            return

        user_query = st.text_area("Bạn muốn xử lý dữ liệu này như thế nào? (Ví dụ: Tính tổng cột Thành tiền nếu cột Trạng thái là Đã giao)", height=100)

        if st.button("🚀 Tạo Công thức / Mã VBA", type="primary"):
            if not user_query.strip():
                st.warning("Vui lòng nhập yêu cầu của bạn!")
                return
                
            with st.spinner("🤖 AI đang tính toán phương án tối ưu nhất..."):
                client = genai.Client(api_key=api_key_input)
                
                prompt = f"""
                Bạn là Chuyên gia Cố vấn Excel và VBA.
                
                ĐÂY LÀ CẤU TRÚC FILE CỦA NGƯỜI DÙNG:
                {context_str}
                
                YÊU CẦU CỦA NGƯỜI DÙNG:
                "{user_query}"
                
                QUY TẮC PHẢN HỒI (BẮT BUỘC TUÂN THỦ):
                1. Quy tắc Ưu tiên (Formula First): Luôn cố gắng giải quyết bằng Công thức Excel tiêu chuẩn trước. Nếu công thức ngắn và tối ưu, hãy trả về công thức.
                2. Đo lường độ phức tạp (Chuyển cấp VBA): Nếu bài toán quá khó, cần vòng lặp, cần thao tác qua nhiều sheet, hoặc công thức quá nặng -> BẮT BUỘC dùng mã VBA.
                3. Luật VBA: Nếu dùng VBA, TRƯỚC KHI in mã code, BẮT BUỘC phải in ra dòng cảnh báo này: "Yêu cầu này không thể giải quyết hiệu quả bằng công thức. Hệ thống đề xuất dùng đoạn mã VBA dưới đây để xử lý tự động và nhẹ máy hơn:".
                4. Hướng dẫn VBA: NẾU có mã VBA, phải in ra 4 bước hướng dẫn sau ở cuối cùng:
                   - Cách dùng VBA:
                   1. Nhấn Alt + F11 để mở cửa sổ VBA.
                   2. Chọn Insert > Module và dán đoạn mã này vào.
                   3. Tại file Excel, vào tab Insert > Shapes > Vẽ một hình chữ nhật.
                   4. Click chuột phải vào hình vừa vẽ > Chọn Assign Macro... > Chọn tên Macro vừa dán > OK. Bấm vào hình để chạy!
                5. Vấn đề Dấu phân cách: Trong công thức, sử dụng dấu chấm phẩy (;) để phân cách các hàm (chuẩn máy tính Việt Nam).
                6. Tính tương thích: Ưu tiên dùng các hàm phổ biến (INDEX, MATCH, IF, SUMIFS...).
                7. TUYỆT ĐỐI xuất code trong khối lệnh Markdown (ví dụ ```excel hoặc ```vba) để giao diện hiển thị nút Copy.
                8. Trả lời cực kỳ ngắn gọn, đi thẳng vào vấn đề, không giải thích dài dòng lan man.
                """
                
                try:
                    response = call_gemini_smart(
                        client, 
                        [prompt]
                    )
                    
                    st.info("💡 Kết quả từ AI (Nhấp vào biểu tượng ở góc phải khung code để sao chép):")
                    st.markdown(response.text)
                    
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi AI: {e}")

# ==========================================
# APP 7: NÉN DUNG LƯỢNG & ĐÓNG GÓI ZIP GỬI MAIL (CỐT LÕI PYMUPDF - TỐI ƯU MÀU SẮC)
# ==========================================
def compress_single_image(img_bytes, quality=50, is_color=True):
    try:
        img = Image.open(BytesIO(img_bytes))
        
        # Xử lý hệ màu theo lựa chọn của người dùng
        if not is_color:
            img = img.convert("L")
        elif img.mode != "RGB":
            img = img.convert("RGB")
            
        # Thu nhỏ khung hình nếu ảnh quá lớn
        max_dim = int(800 + (quality - 30) * (1600 - 800) / (90 - 30))
        if max(img.width, img.height) > max_dim:
            ratio = max_dim / float(max(img.width, img.height))
            new_size = (int(img.width * ratio), int(img.height * ratio))
            img = img.resize(new_size, Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS)
                
        out = BytesIO()
        img.save(out, format="JPEG", quality=quality, optimize=True)
        return out.getvalue()
    except Exception:
        return img_bytes

def compress_pdf_extreme(pdf_bytes, quality=50, is_color=True):
    """Sử dụng PyMuPDF cào bằng trang PDF, tối ưu hóa DPI và giữ nguyên hệ màu RGB"""
    try:
        doc = fitz.open("pdf", pdf_bytes)
        new_pdf = fitz.open()
        
        # Ép DPI từ 100 (siêu nhẹ) đến 200 (sắc nét) dựa trên thanh trượt
        dpi = int(100 + (quality - 30) * (200 - 100) / (90 - 30))
        
        for page in doc:
            zoom = dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            
            # Khởi tạo bản vẽ với RGB (Màu) hoặc GRAY (Đen trắng)
            c_space = fitz.csRGB if is_color else fitz.csGRAY
            pix = page.get_pixmap(matrix=mat, colorspace=c_space)
            
            # Ép nén bằng thuật toán phân tích hạt JPEG
            img_bytes = pix.tobytes("jpeg", jpg_quality=quality)
            
            # Gói lại vào trang PDF mới tinh (loại bỏ hoàn toàn rác mã hóa của máy scan)
            imgdoc = fitz.open("jpeg", img_bytes)
            pdfbytes = imgdoc.convert_to_pdf()
            imgpdf = fitz.open("pdf", pdfbytes)
            new_pdf.insert_pdf(imgpdf)
            
        out_stream = BytesIO()
        new_pdf.save(out_stream, garbage=4, deflate=True)
        return out_stream.getvalue()
    except Exception as e:
        return pdf_bytes

def app_compress_and_zip():
    st.title("🗜️ Nén Dung Lượng & Đóng Gói ZIP Gửi Mail")
    st.markdown("Tối ưu hóa các file PDF/Ảnh scan nặng, tự động đóng gói toàn bộ vào 1 file **.ZIP** để gửi email công ty không bị vượt trần 25MB.")

    uploaded_files = st.file_uploader(
        "Tải lên một hoặc nhiều file cần nén và đóng gói (PDF, JPG, PNG, DOCX, XLSX...):",
        accept_multiple_files=True,
        key=f"app7_{st.session_state.uploader_key}"
    )

    if uploaded_files:
        total_orig_bytes = sum(len(f.getvalue()) for f in uploaded_files)
        total_orig_mb = total_orig_bytes / (1024 * 1024)
        st.info(f"📁 Đã chọn **{len(uploaded_files)} file** | Tổng dung lượng gốc: **{total_orig_mb:.2f} MB**")

        # Giao diện điều khiển mới chia làm 2 cột
        col1, col2 = st.columns(2)
        with col1:
            che_do_mau = st.radio(
                "Chế độ màu sắc:",
                ["🌈 Giữ nguyên màu gốc", "🔲 Chuyển Đen Trắng"],
                help="Chế độ Giữ Màu vẫn ép nén rất tốt. Chỉ dùng Đen Trắng khi bạn cần file đạt mức siêu nhẹ gọn."
            )
        with col2:
            quality = st.slider(
                "Mức chất lượng nén (Mặc định 50):",
                min_value=30,
                max_value=90,
                value=50,
                step=5,
                help="Điều chỉnh thanh trượt càng thấp, cả độ phân giải và chất lượng ảnh sẽ càng được tối ưu nhỏ lại."
            )
        
        # Bổ sung khung hướng dẫn sử dụng chi tiết
        st.markdown("""
        > 💡 **Bạn hoàn toàn làm chủ độ nét thông qua Thanh trượt chất lượng:**
        > - **Kéo lên mức 70 - 90:** Dành cho các hợp đồng quan trọng, ảnh chụp linh kiện máy móc cần độ nét cao để nghiệm thu (chấp nhận file ZIP đầu ra sẽ nặng hơn một chút).
        > - **Để mức 40 - 50:** Tối ưu hóa tối đa để đính kèm qua email nội bộ bị giới hạn 25MB.
        """)

        # Biến boolean kiểm tra chế độ màu
        is_color = "🌈" in che_do_mau

        zip_filename = st.text_input("Tên file ZIP xuất ra:", value="Tai_Lieu_Gui_Mail.zip")
        if not zip_filename.lower().endswith(".zip"):
            zip_filename += ".zip"

        if st.button("🚀 Bắt đầu Nén & Đóng Gói ZIP", type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            zip_buffer = BytesIO()
            report_rows = []

            # Sử dụng thuật toán ZipDeflated cấp độ nén cao nhất
            with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
                for idx, file in enumerate(uploaded_files):
                    fname = file.name
                    fext = fname.split('.')[-1].lower()
                    raw_bytes = file.getvalue()
                    orig_len = len(raw_bytes)

                    status_text.text(f"Đang xử lý ({idx+1}/{len(uploaded_files)}): {fname}...")

                    # Phân luồng nén
                    if fext == "pdf":
                        processed_bytes = compress_pdf_extreme(raw_bytes, quality=quality, is_color=is_color)
                    elif fext in ["jpg", "jpeg", "png"]:
                        processed_bytes = compress_single_image(raw_bytes, quality=quality, is_color=is_color)
                        # Đồng nhất định dạng ảnh đầu ra thành JPG để nhẹ gọn nhất
                        if not fname.lower().endswith(".jpg") and not fname.lower().endswith(".jpeg"):
                            fname = f"{fname.rsplit('.', 1)[0]}.jpg"
                    else:
                        # Giữ nguyên bản gốc đối với các file Word, Excel
                        processed_bytes = raw_bytes

                    new_len = len(processed_bytes)
                    zf.writestr(fname, processed_bytes)

                    saved_pct = ((orig_len - new_len) / orig_len * 100) if orig_len > 0 else 0
                    report_rows.append({
                        "Tên file": fname,
                        "Gốc (MB)": f"{orig_len / (1024*1024):.2f}",
                        "Sau nén (MB)": f"{new_len / (1024*1024):.2f}",
                        "Tiết kiệm": f"{saved_pct:.1f}%"
                    })

                    progress_bar.progress((idx + 1) / len(uploaded_files))

            final_zip_bytes = zip_buffer.getvalue()
            final_zip_mb = len(final_zip_bytes) / (1024 * 1024)
            total_saved_pct = ((total_orig_bytes - len(final_zip_bytes)) / total_orig_bytes * 100) if total_orig_bytes > 0 else 0

            status_text.empty()
            progress_bar.empty()

            # Hiển thị số liệu trực quan
            col1, col2, col3 = st.columns(3)
            col1.metric("Tổng dung lượng gốc", f"{total_orig_mb:.2f} MB")
            col2.metric("File ZIP hoàn thành", f"{final_zip_mb:.2f} MB")
            col3.metric("Tiết kiệm được", f"{total_saved_pct:.1f}%")

            if final_zip_mb <= 25.0:
                st.success(f"🎉 Siêu nén thành công! File ZIP nặng **{final_zip_mb:.2f} MB**, hoàn toàn đủ điều kiện gửi email công ty.")
            else:
                st.warning(f"⚠️ File ZIP hiện tại nặng **{final_zip_mb:.2f} MB** (vẫn vượt 25MB). Hãy thử giảm mức chất lượng xuống 30-40% hoặc chọn Đen trắng.")

            st.dataframe(pd.DataFrame(report_rows), use_container_width=True)

            st.download_button(
                label=f"📥 Tải xuống {zip_filename}",
                data=final_zip_bytes,
                file_name=zip_filename,
                mime="application/zip",
                on_click=clear_file
            )
# ==========================================
# ==========================================
# APP 8: XOAY TÙY Ý TỪNG TRANG & TỰ ĐỘNG GHÉP FILE
# ==========================================
def app_rotate_pdf():
    import fitz
    from io import BytesIO
    from PIL import Image

    st.title("🔄 Xoay Trang PDF & Ghép File")
    st.markdown("Tải lên nhiều file PDF, điều chỉnh góc xoay cho từng trang bị ngược. Hệ thống sẽ tự động ghép tất cả lại thành **1 file PDF duy nhất**.")

    uploaded_files = st.file_uploader(
        "Tải lên các file PDF cần kiểm tra, xoay và ghép:", 
        type=["pdf"], 
        accept_multiple_files=True, 
        key=f"app8_{st.session_state.uploader_key}"
    )

    if uploaded_files:
        st.info(f"📁 Đã chọn **{len(uploaded_files)} file**. Các file sẽ được ghép nối theo đúng thứ tự tải lên.")
        
        # Lặp qua từng file và tạo khung mở rộng riêng
        for file_idx, uploaded_file in enumerate(uploaded_files):
            pdf_bytes = uploaded_file.getvalue()
            
            try:
                doc = fitz.open("pdf", pdf_bytes)
                num_pages = len(doc)
                
                with st.expander(f"📄 File {file_idx + 1}: {uploaded_file.name} ({num_pages} trang)", expanded=True):
                    # Quản lý bộ nhớ trạng thái xoay riêng cho từng file
                    state_key = f"rotations_{uploaded_file.name}_{st.session_state.uploader_key}"
                    if state_key not in st.session_state:
                        st.session_state[state_key] = {i: 0 for i in range(num_pages)}

                    cols = st.columns(3)
                    
                    for i in range(num_pages):
                        col_idx = i % 3
                        page = doc[i]
                        
                        with cols[col_idx]:
                            st.markdown(f"<div style='text-align: center; color: #003366; margin-top: 15px;'><b>Trang {i+1}</b></div>", unsafe_allow_html=True)
                            
                            rot_val = st.selectbox(
                                "Hướng xoay:",
                                options=[0, 90, 180, 270],
                                index=[0, 90, 180, 270].index(st.session_state[state_key][i]),
                                format_func=lambda x: "✅ Giữ nguyên" if x == 0 else "↪️ Phải 90°" if x == 90 else "🔃 Ngược 180°" if x == 180 else "↩️ Trái 90°",
                                key=f"select_rot_{uploaded_file.name}_{i}_{state_key}",
                                label_visibility="collapsed"
                            )
                            st.session_state[state_key][i] = rot_val
                            
                            current_rot = page.rotation
                            page.set_rotation((current_rot + rot_val) % 360)
                            
                            mat = fitz.Matrix(0.3, 0.3)
                            pix = page.get_pixmap(matrix=mat)
                            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                            
                            st.image(img, use_container_width=True)
                            st.markdown("<br>", unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Lỗi khi đọc file {uploaded_file.name}: {e}")

        st.markdown("---")
        
        if st.button("🚀 Lưu & Tự Động Ghép File", type="primary"):
            with st.spinner("Đang áp dụng góc xoay và ghép tất cả thành 1 file PDF duy nhất..."):
                try:
                    # Khởi tạo một file PDF tổng (rỗng)
                    merged_doc = fitz.open()
                    
                    for file in uploaded_files:
                        out_doc = fitz.open("pdf", file.getvalue())
                        state_key = f"rotations_{file.name}_{st.session_state.uploader_key}"
                        
                        if state_key in st.session_state:
                            for i in range(len(out_doc)):
                                p = out_doc[i]
                                p.set_rotation((p.rotation + st.session_state[state_key][i]) % 360)
                        
                        # Chèn file PDF đã xoay vào file tổng
                        merged_doc.insert_pdf(out_doc)
                    
                    out_stream = BytesIO()
                    merged_doc.save(out_stream, garbage=3, deflate=True)
                    
                    st.success("🎉 Đã xoay và ghép xong tất cả các file thành 1 file duy nhất!")
                    st.download_button(
                        label="📥 Tải xuống PDF Đã Ghép",
                        data=out_stream.getvalue(),
                        file_name="Tai_Lieu_Da_Xoay_Va_Ghep.pdf",
                        mime="application/pdf",
                        on_click=clear_file
                    )
                except Exception as e:
                    st.error(f"Có lỗi trong quá trình ghép file: {e}")
# ==========================================
# ==========================================
# APP 9: XÓA SIÊU DỮ LIỆU (METADATA CLEANER) BẢO MẬT TÀI LIỆU
# ==========================================
def clean_pdf_metadata(pdf_bytes):
    import fitz
    from io import BytesIO
    doc = fitz.open("pdf", pdf_bytes)
    # Gán trắng toàn bộ từ điển Metadata ẩn
    doc.set_metadata({})
    out = BytesIO()
    doc.save(out, garbage=4, deflate=True)
    return out.getvalue()

def clean_docx_metadata(docx_bytes):
    from docx import Document
    from io import BytesIO
    doc = Document(BytesIO(docx_bytes))
    props = doc.core_properties
    props.author = ""
    props.last_modified_by = ""
    props.comments = ""
    props.title = ""
    props.subject = ""
    props.category = ""
    props.keywords = ""
    out = BytesIO()
    doc.save(out)
    return out.getvalue()

def clean_xlsx_metadata(xlsx_bytes):
    import openpyxl
    from io import BytesIO
    wb = openpyxl.load_workbook(BytesIO(xlsx_bytes))
    props = wb.properties
    props.creator = ""
    props.lastModifiedBy = ""
    props.title = ""
    props.subject = ""
    props.description = ""
    props.keywords = ""
    props.category = ""
    props.company = ""
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

def app_metadata_cleaner():
    import zipfile
    from io import BytesIO

    st.title("🛡️ Xóa Siêu Dữ Liệu (Metadata) Bảo Mật Tài Liệu")
    st.markdown("Tự động xóa sạch tên tác giả, người chỉnh sửa cuối cùng, thông tin máy in, phần mềm và lịch sử ẩn bên trong file **PDF, Word (.docx), Excel (.xlsx)** trước khi gửi ra ngoài.")

    uploaded_files = st.file_uploader(
        "Tải lên các file cần tẩy sạch dữ liệu ẩn (Hỗ trợ PDF, DOCX, XLSX):",
        type=["pdf", "docx", "xlsx"],
        accept_multiple_files=True,
        key=f"app9_{st.session_state.uploader_key}"
    )

    if uploaded_files:
        st.info(f"📁 Đã chọn **{len(uploaded_files)} file** sẵn sàng làm sạch dữ liệu bảo mật.")

        if st.button("🚀 Bắt đầu Tẩy Sạch Dữ Liệu Ẩn", type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            # Trường hợp 1: Tải lên 1 file -> Trả về đúng định dạng gốc
            if len(uploaded_files) == 1:
                file = uploaded_files[0]
                fname = file.name
                fext = fname.split('.')[-1].lower()
                raw_bytes = file.getvalue()

                status_text.text(f"Đang làm sạch file: {fname}...")
                try:
                    if fext == "pdf":
                        cleaned_bytes = clean_pdf_metadata(raw_bytes)
                        mime_type = "application/pdf"
                    elif fext == "docx":
                        cleaned_bytes = clean_docx_metadata(raw_bytes)
                        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    elif fext == "xlsx":
                        cleaned_bytes = clean_xlsx_metadata(raw_bytes)
                        mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    else:
                        cleaned_bytes = raw_bytes
                        mime_type = "application/octet-stream"

                    progress_bar.progress(100)
                    status_text.empty()
                    st.success(f"🎉 Đã tẩy sạch toàn bộ metadata của file **{fname}**!")

                    st.download_button(
                        label=f"📥 Tải xuống {fname} (Bản sạch)",
                        data=cleaned_bytes,
                        file_name=f"Cleaned_{fname}",
                        mime=mime_type,
                        on_click=clear_file
                    )
                except Exception as e:
                    st.error(f"Lỗi khi xử lý file: {e}")

            # Trường hợp 2: Tải lên nhiều file -> Làm sạch và đóng gói ZIP
            else:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for idx, file in enumerate(uploaded_files):
                        fname = file.name
                        fext = fname.split('.')[-1].lower()
                        raw_bytes = file.getvalue()

                        status_text.text(f"Đang xử lý ({idx+1}/{len(uploaded_files)}): {fname}...")
                        try:
                            if fext == "pdf":
                                cleaned_bytes = clean_pdf_metadata(raw_bytes)
                            elif fext == "docx":
                                cleaned_bytes = clean_docx_metadata(raw_bytes)
                            elif fext == "xlsx":
                                cleaned_bytes = clean_xlsx_metadata(raw_bytes)
                            else:
                                cleaned_bytes = raw_bytes

                            zf.writestr(f"Cleaned_{fname}", cleaned_bytes)
                        except Exception:
                            zf.writestr(f"Loi_{fname}", raw_bytes)

                        progress_bar.progress((idx + 1) / len(uploaded_files))

                progress_bar.progress(100)
                status_text.empty()
                st.success("🎉 Đã làm sạch toàn bộ metadata và đóng gói thành công!")

                st.download_button(
                    label="📥 Tải xuống File ZIP đã làm sạch",
                    data=zip_buffer.getvalue(),
                    file_name="Tai_Lieu_Da_Lam_Sach.zip",
                    mime="application/zip",
                    on_click=clear_file
                )
# ==========================================
# ==========================================
# APP 10: TÁCH NỀN CHỮ KÝ (OFFLINE BẰNG OPENCV & NUMPY)
# ==========================================
def app_signature_extractor():
    import numpy as np
    import cv2
    from PIL import Image, ImageOps
    from io import BytesIO

    st.title("✍️ Tách Nền Chữ Ký (Giữ Nguyên Màu Mực)")
    st.markdown("Chụp ảnh chữ ký trên giấy trắng, hệ thống (chạy offline 100%) sẽ lọc bỏ nền giấy, lọc đốm bụi và xuất ra file PNG trong suốt hoàn hảo với màu mực thật.")

    MAX_SIDE = 3500

    def load_bgr_from_bytes(img_bytes):
        im = ImageOps.exif_transpose(Image.open(BytesIO(img_bytes))).convert("RGB")
        arr = np.asarray(im)[:, :, ::-1].copy()
        h, w = arr.shape[:2]
        m = max(h, w)
        if m > MAX_SIDE:
            s = MAX_SIDE / m
            arr = cv2.resize(arr, (int(w * s), int(h * s)), interpolation=cv2.INTER_AREA)
        return arr

    def estimate_background(bgr):
        h, w = bgr.shape[:2]
        s = 800.0 / max(h, w) if max(h, w) > 800 else 1.0
        small = cv2.resize(bgr, None, fx=s, fy=s, interpolation=cv2.INTER_AREA) if s < 1 else bgr
        ks = max(9, int(max(small.shape[:2]) * 0.06)) | 1
        kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (ks, ks))
        bg = cv2.morphologyEx(small, cv2.MORPH_CLOSE, kernel)
        bg = cv2.GaussianBlur(bg, (0, 0), ks / 4.0)
        return cv2.resize(bg, (w, h), interpolation=cv2.INTER_LINEAR)

    def auto_threshold(ink):
        vals = ink[ink > 20]
        if vals.size < 200:
            return 40.0
        v8 = np.clip(vals, 0, 255).astype(np.uint8).reshape(-1, 1)
        t, _ = cv2.threshold(v8, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        return float(np.clip(t, 30, 140))

    def process(bgr, sens=1.0, speck=1.0, crop=True, flat_color=False):
        h, w = bgr.shape[:2]
        bg = estimate_background(bgr)
        norm = np.clip(bgr.astype(np.float32) / np.maximum(bg.astype(np.float32), 1.0), 0.0, 1.0)
        ink = (1.0 - norm.min(axis=2)) * 255.0
        ink = cv2.GaussianBlur(ink, (0, 0), 0.7)

        t = auto_threshold(ink) / max(sens, 0.1)
        lo, hi = t * 0.45, t * 1.15
        a = np.clip((ink - lo) / (hi - lo), 0.0, 1.0)
        a = a * a * (3.0 - 2.0 * a)

        if speck > 0:
            mask = (a > 0.25).astype(np.uint8)
            n, lab, stats, cent = cv2.connectedComponentsWithStats(mask, connectivity=8)
            if n > 1:
                areas = stats[1:, cv2.CC_STAT_AREA].astype(np.float64)
                min_area = max(8.0, 5e-6 * h * w * speck)
                big = np.zeros(n, dtype=bool)
                big[1:] = areas >= max(areas.max() * 0.02, 4 * min_area)
                if not big.any():
                    big[1 + int(areas.argmax())] = True
                sc = 0.25
                sb = cv2.resize(big[lab].astype(np.uint8), None, fx=sc, fy=sc, interpolation=cv2.INTER_NEAREST)
                r = max(3, int(0.06 * max(h, w) * sc))
                zone = cv2.dilate(sb, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (2 * r + 1, 2 * r + 1)))
                zone = cv2.resize(zone, (w, h), interpolation=cv2.INTER_NEAREST).astype(bool)
                cx = np.clip(cent[:, 0].astype(int), 0, w - 1)
                cy = np.clip(cent[:, 1].astype(int), 0, h - 1)
                ar = np.zeros(n)
                ar[1:] = areas
                keep = big | (zone[cy, cx] & (ar >= max(4.0, 0.9 * min_area))) | (ar >= 4 * min_area)
                keep[0] = False
                kept = cv2.dilate(keep[lab].astype(np.uint8), np.ones((3, 3), np.uint8), iterations=2)
                a = a * kept

        if a.max() < 0.3:
            return None

        core = a > 0.9
        if core.sum() < 30:
            core = a > 0.5
        wgt = core.astype(np.float32)
        sig = max(2.0, min(h, w) / 400.0)
        num = cv2.GaussianBlur(bgr.astype(np.float32) * wgt[..., None], (0, 0), sig)
        den = cv2.GaussianBlur(wgt, (0, 0), sig)[..., None]
        spread = num / np.maximum(den, 1e-4)
        typical = np.median(bgr[core], axis=0).astype(np.float32)
        
        if flat_color:
            col = np.broadcast_to(typical, bgr.shape).copy()
        else:
            col = np.where(core[..., None], bgr.astype(np.float32), np.where(den > 0.02, spread, typical))

        rgba = np.dstack([col[..., ::-1], a * 255.0])
        rgba = np.clip(rgba + 0.5, 0, 255).astype(np.uint8)

        if crop:
            ys, xs = np.where(a > 0.05)
            pad = max(10, int(0.03 * max(h, w)))
            y0, y1 = max(0, ys.min() - pad), min(h, ys.max() + pad + 1)
            x0, x1 = max(0, xs.min() - pad), min(w, xs.max() + pad + 1)
            rgba = rgba[y0:y1, x0:x1]
        return rgba

    def composite(rgba, mode="caro"):
        h, w = rgba.shape[:2]
        if mode == "trang":
            bg = np.full((h, w, 3), 255.0, np.float32)
        elif mode == "xanh":
            bg = np.zeros((h, w, 3), np.float32)
            bg[...] = (60, 200, 90)
        else:
            yy, xx = np.mgrid[0:h, 0:w]
            tile = ((yy // 12 + xx // 12) % 2)[..., None]
            bg = np.where(tile == 0, 238.0, 205.0) * np.ones((h, w, 3), np.float32)
        al = rgba[..., 3:4].astype(np.float32) / 255.0
        out = rgba[..., :3].astype(np.float32) * al + bg * (1 - al)
        return Image.fromarray(np.clip(out, 0, 255).astype(np.uint8), "RGB")

    # --- GIAO DIỆN STREAMLIT CHO APP 11 ---
    col1, col2 = st.columns([1, 2])
    
    with col1:
        uploaded_sig = st.file_uploader("Tải lên ảnh chữ ký (JPG, PNG)", type=["jpg", "jpeg", "png"])
        
        st.markdown("**⚙️ Tùy chỉnh (Không bắt buộc)**")
        sens = st.slider("Độ nhạy nét mực (Tăng nếu bị đứt nét)", min_value=0.4, max_value=2.5, value=1.0, step=0.1)
        speck = st.slider("Mức lọc đốm bụi rác (0 = Tắt)", min_value=0.0, max_value=8.0, value=1.0, step=0.5)
        crop_opt = st.checkbox("Tự cắt sát chữ ký", value=True)
        flat_opt = st.checkbox("Làm đều 1 màu mực (thay vì màu thật)", value=False)
        bg_opt = st.radio("Nền xem thử:", ["Ô caro", "Trắng", "Xanh lá (soi viền thừa)"], index=0)

    with col2:
        if uploaded_sig:
            st.image(uploaded_sig, caption="📸 Ảnh gốc", use_container_width=True)
            
            with st.spinner("Đang trích xuất nét mực..."):
                try:
                    bgr_img = load_bgr_from_bytes(uploaded_sig.getvalue())
                    rgba_result = process(bgr_img, sens=sens, speck=speck, crop=crop_opt, flat_color=flat_opt)
                    
                    if rgba_result is None:
                        st.error("❌ Không thấy nét mực! Vui lòng thử tăng 'Độ nhạy nét mực' hoặc chọn ảnh rõ nét hơn.")
                    else:
                        # Map chế độ nền sang mode chuẩn của hàm composite
                        bg_mode_map = {"Ô caro": "caro", "Trắng": "trang", "Xanh lá (soi viền thừa)": "xanh"}
                        preview_img = composite(rgba_result, mode=bg_mode_map[bg_opt])
                        
                        st.image(preview_img, caption="✨ Kết quả (Đã tách nền trong suốt)", use_container_width=True)
                        
                        # Chuyển RGBA numpy array sang PNG Bytes để tải xuống
                        final_img = Image.fromarray(rgba_result, "RGBA")
                        out_stream = BytesIO()
                        final_img.save(out_stream, format="PNG", optimize=True)
                        
                        st.download_button(
                            label="📥 Tải xuống Chữ ký (.PNG trong suốt)",
                            data=out_stream.getvalue(),
                            file_name=f"Chu_Ky_Trong_Suot_{uploaded_sig.name.split('.')[0]}.png",
                            mime="image/png",
                            type="primary",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"Lỗi khi xử lý ảnh: {e}")
# ==========================================
# APP 11: TẠO MÃ QR ĐA NGUỒN (OFFLINE)
# ==========================================
QR_BANKS = {
    "Vietcombank": "970436", "VietinBank": "970415", "BIDV": "970418", "Agribank": "970405",
    "Techcombank": "970407", "MB Bank": "970422", "ACB": "970416", "VPBank": "970432",
    "Sacombank": "970403", "TPBank": "970423", "VIB": "970441", "HDBank": "970437",
    "SHB": "970443", "OCB": "970448", "MSB": "970426", "SeABank": "970440",
}


def qr_crc16(s):
    """CRC16-CCITT (poly 0x1021, init 0xFFFF) dùng cho chuẩn EMVCo/VietQR."""
    crc = 0xFFFF
    for b in s.encode("utf-8"):
        crc ^= b << 8
        for _ in range(8):
            crc = ((crc << 1) ^ 0x1021) if crc & 0x8000 else (crc << 1)
            crc &= 0xFFFF
    return f"{crc:04X}"


def _tlv(tag, val):
    return f"{tag}{len(val):02d}{val}"


def qr_vietqr_payload(bin_code, account, amount=0, note=""):
    acc = _tlv("00", bin_code) + _tlv("01", account)
    merchant = _tlv("00", "A000000727") + _tlv("01", acc) + _tlv("02", "QRIBFTTA")
    s = _tlv("00", "01") + _tlv("01", "12" if amount else "11") + _tlv("38", merchant) + _tlv("53", "704")
    if amount:
        s += _tlv("54", str(int(amount)))
    s += _tlv("58", "VN")
    if note.strip():
        s += _tlv("62", _tlv("08", note.strip()[:25]))
    s += "6304"
    return s + qr_crc16(s)


def _esc_wifi(s):
    for ch in ('\\', ';', ',', ':', '"'):
        s = s.replace(ch, '\\' + ch)
    return s


def _esc_vc(s):
    return str(s).replace('\\', '\\\\').replace(';', '\\;').replace(',', '\\,').replace('\n', '\\n')


def qr_matrix(data, ecc="M"):
    """Trả về ma trận bool (không có viền). Ưu tiên thư viện qrcode, dự phòng OpenCV."""
    try:
        import qrcode
        lv = {"L": qrcode.constants.ERROR_CORRECT_L, "M": qrcode.constants.ERROR_CORRECT_M,
              "Q": qrcode.constants.ERROR_CORRECT_Q, "H": qrcode.constants.ERROR_CORRECT_H}[ecc]
        q = qrcode.QRCode(version=None, error_correction=lv, border=0)
        q.add_data(data)
        q.make(fit=True)
        return [[bool(c) for c in row] for row in q.get_matrix()]
    except ImportError:
        import cv2
        import numpy as np
        p = cv2.QRCodeEncoder_Params()
        p.correction_level = {"L": cv2.QRCodeEncoder_CORRECT_LEVEL_L, "M": cv2.QRCodeEncoder_CORRECT_LEVEL_M,
                              "Q": cv2.QRCodeEncoder_CORRECT_LEVEL_Q, "H": cv2.QRCodeEncoder_CORRECT_LEVEL_H}[ecc]
        img = cv2.QRCodeEncoder_create(p).encode(data)
        if img is None:
            raise ValueError("Nội dung quá dài để tạo mã QR.")
        m = img < 128
        ys, xs = np.where(m)
        m = m[ys.min():ys.max() + 1, xs.min():xs.max() + 1]      # cắt sát vùng mã (OpenCV chừa viền 2 module)
        # OpenCV có thể phóng mỗi module thành nhiều điểm ảnh -> tìm bước và thu về 1 điểm/module
        step = 0
        for arr in (m, m.T):
            runs = []
            for row in arr:
                cnt = 1
                for i in range(1, len(row)):
                    if row[i] == row[i - 1]:
                        cnt += 1
                    else:
                        runs.append(cnt)
                        cnt = 1
                runs.append(cnt)
            step = int(np.gcd.reduce(runs)) if step == 0 else int(np.gcd(step, np.gcd.reduce(runs)))
        step = max(step, 1)
        n = m.shape[1] // step
        return [[bool(m[r * step + step // 2, c * step + step // 2]) for c in range(n)] for r in range(n)]


def _hex_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _in_finder(r, c, n):
    return (r < 7 and c < 7) or (r < 7 and c >= n - 7) or (r >= n - 7 and c < 7)


def qr_render(data, ecc="M", box=20, border=4, fg="#000000", fg2=None, bg="#FFFFFF", transparent=False,
              shape="square", logo_bytes=None, caption=""):
    from PIL import Image, ImageDraw, ImageFont
    from io import BytesIO
    if logo_bytes and ecc != "H":
        ecc = "H"
    mat = qr_matrix(data, ecc)
    n = len(mat)
    S = (n + 2 * border) * box
    mask = Image.new("L", (S, S), 0)
    d = ImageDraw.Draw(mask)
    off = border * box
    rad = int(box * 0.5)
    for r in range(n):
        for c in range(n):
            if not mat[r][c]:
                continue
            x, y = off + c * box, off + r * box
            if _in_finder(r, c, n):
                if shape == "square":
                    d.rectangle([x, y, x + box - 1, y + box - 1], fill=255)
                else:   # mắt định vị: bo nhẹ nhưng vẫn liền khối để dễ quét
                    d.rounded_rectangle([x, y, x + box - 1, y + box - 1], radius=int(box * 0.3), fill=255)
                    for dr, dc in ((0, 1), (1, 0), (0, -1), (-1, 0)):
                        rr, cc = r + dr, c + dc
                        if 0 <= rr < n and 0 <= cc < n and mat[rr][cc] and _in_finder(rr, cc, n):
                            d.rectangle([x + (box // 2 if dc == 1 else 0), y + (box // 2 if dr == 1 else 0),
                                         x + box - 1 - (box // 2 if dc == -1 else 0),
                                         y + box - 1 - (box // 2 if dr == -1 else 0)], fill=255)
            elif shape == "dots":
                pad = int(box * 0.08)
                d.ellipse([x + pad, y + pad, x + box - 1 - pad, y + box - 1 - pad], fill=255)
            elif shape == "rounded":
                d.rounded_rectangle([x, y, x + box - 1, y + box - 1], radius=rad, fill=255)
            else:
                d.rectangle([x, y, x + box - 1, y + box - 1], fill=255)

    c1 = _hex_rgb(fg)
    if fg2 and fg2.lower() != fg.lower():
        c2 = _hex_rgb(fg2)
        import numpy as np
        t = (np.add.outer(np.arange(S), np.arange(S)) / (2.0 * (S - 1)))[..., None]
        arr = (np.array(c1) * (1 - t) + np.array(c2) * t).astype("uint8")
        ink = Image.fromarray(arr, "RGB")
    else:
        ink = Image.new("RGB", (S, S), c1)

    out = Image.new("RGBA", (S, S), (0, 0, 0, 0) if transparent else _hex_rgb(bg) + (255,))
    out.paste(ink, (0, 0), mask)

    if logo_bytes:
        lg = Image.open(BytesIO(logo_bytes)).convert("RGBA")
        pad = max(3, box // 4)
        side = max(8, int(n * box * 0.19) - 2 * pad)      # cả khung nền logo <= ~19% cạnh mã (~3,6% diện tích)
        lg.thumbnail((side, side), Image.LANCZOS)
        plate = Image.new("RGBA", (lg.width + 2 * pad, lg.height + 2 * pad), (0, 0, 0, 0))
        ImageDraw.Draw(plate).rounded_rectangle([0, 0, plate.width - 1, plate.height - 1], radius=pad,
                                                fill=_hex_rgb(bg) + (255,))
        px, py = (S - plate.width) // 2, (S - plate.height) // 2
        out.alpha_composite(plate, (px, py))
        out.alpha_composite(lg, (px + pad, py + pad))

    if caption.strip():
        fsize = max(14, box * 2)
        font = None
        for p in ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "DejaVuSans-Bold.ttf",
                  "C:/Windows/Fonts/arialbd.ttf", "arialbd.ttf"):
            try:
                font = ImageFont.truetype(p, fsize)
                break
            except Exception:
                pass
        if font is None:
            try:
                font = ImageFont.load_default(size=fsize)
            except Exception:
                font = ImageFont.load_default()
        lines = []
        for para in caption.strip().split("\n"):
            cur = ""
            for wd in para.split(" "):
                t = (cur + " " + wd).strip()
                if ImageDraw.Draw(out).textlength(t, font=font) <= S * 0.94 or not cur:
                    cur = t
                else:
                    lines.append(cur)
                    cur = wd
            lines.append(cur)
        lh = int(fsize * 1.3)
        full = Image.new("RGBA", (S, S + lh * len(lines) + box), (0, 0, 0, 0) if transparent else _hex_rgb(bg) + (255,))
        full.alpha_composite(out, (0, 0))
        dd = ImageDraw.Draw(full)
        for i, ln in enumerate(lines):
            w = dd.textlength(ln, font=font)
            dd.text(((S - w) / 2, S - box * 1.2 + i * lh), ln, font=font, fill=c1 + (255,))
        out = full
    return out


def qr_svg(data, ecc="M", border=4, fg="#000000", bg="#FFFFFF", transparent=False, shape="square"):
    mat = qr_matrix(data, ecc)
    n = len(mat)
    T = n + 2 * border
    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {T} {T}" shape-rendering="crispEdges">']
    if not transparent:
        parts.append(f'<rect width="{T}" height="{T}" fill="{bg}"/>')
    parts.append(f'<g fill="{fg}">')
    for r in range(n):
        for c in range(n):
            if mat[r][c]:
                x, y = c + border, r + border
                if shape == "dots" and not _in_finder(r, c, n):
                    parts.append(f'<circle cx="{x + .5}" cy="{y + .5}" r=".46"/>')
                elif shape != "square":
                    parts.append(f'<rect x="{x}" y="{y}" width="1" height="1" rx=".3"/>')
                else:
                    parts.append(f'<rect x="{x}" y="{y}" width="1" height="1"/>')
    parts.append('</g></svg>')
    return "\n".join(parts).encode("utf-8")


def qr_to_png_bytes(img):
    from io import BytesIO
    b = BytesIO()
    img.save(b, format="PNG", optimize=True)
    return b.getvalue()


def app_qr_generator():
    import zipfile
    from io import BytesIO
    from datetime import datetime as _dt

    st.title("🔳 Tạo Mã QR Đa Năng")
    st.markdown("Tạo mã QR từ **đường link, Wi-Fi, danh thiếp, chuyển khoản ngân hàng (VietQR), email, SMS, bản đồ, lịch họp...** "
                "hoặc **tạo hàng loạt từ file Excel**. Chạy ngay trên máy chủ, không gửi nội dung cho AI.")

    tab_one, tab_batch = st.tabs(["🔳 Tạo từng mã", "📚 Tạo hàng loạt từ Excel/CSV"])

    # ---------------- cấu hình giao diện chung ----------------
    def style_panel(prefix):
        with st.expander("🎨 Tùy chỉnh giao diện mã QR", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                fg = st.color_picker("Màu mã", "#000000", key=f"{prefix}_fg")
                grad = st.checkbox("Chuyển màu (gradient)", key=f"{prefix}_grad")
                fg2 = st.color_picker("Màu thứ hai", "#003366", key=f"{prefix}_fg2") if grad else None
                shape_lbl = st.radio("Kiểu điểm ảnh", ["Vuông", "Bo tròn", "Chấm tròn"], horizontal=True, key=f"{prefix}_shape")
            with c2:
                bg = st.color_picker("Màu nền", "#FFFFFF", key=f"{prefix}_bg")
                transparent = st.checkbox("Nền trong suốt (PNG)", key=f"{prefix}_tr")
                ecc = st.selectbox("Mức sửa lỗi", ["L (7%)", "M (15%)", "Q (25%)", "H (30%)"], index=1, key=f"{prefix}_ecc")
                box = st.slider("Kích thước (độ lớn mỗi ô)", 8, 40, 20, key=f"{prefix}_box")
            logo = st.file_uploader("Logo ở giữa (không bắt buộc, tự nâng mức sửa lỗi lên H)", type=["png", "jpg", "jpeg"], key=f"{prefix}_logo")
            caption = st.text_input("Chữ dưới mã QR (không bắt buộc)", key=f"{prefix}_cap")
        shape = {"Vuông": "square", "Bo tròn": "rounded", "Chấm tròn": "dots"}[shape_lbl]
        low = fg2 if grad else None
        # độ tương phản: mã sáng hơn nền thì máy quét thường không đọc được
        def lum(h):
            r, g, b = _hex_rgb(h)
            return 0.299 * r + 0.587 * g + 0.114 * b
        if not transparent and lum(fg) >= lum(bg):
            st.warning("⚠️ Màu mã đang sáng hơn hoặc bằng màu nền, nhiều máy quét sẽ không đọc được. Nên dùng mã TỐI trên nền SÁNG.")
        if not transparent and abs(lum(fg) - lum(bg)) < 90:
            st.warning("⚠️ Độ tương phản giữa mã và nền thấp, hãy chọn màu đậm hơn.")
        return dict(fg=fg, fg2=low, bg=bg, transparent=transparent, shape=shape, ecc=ecc[0], box=box,
                    logo=logo.getvalue() if logo else None, caption=caption)

    # ---------------- TAB 1 ----------------
    with tab_one:
        kinds = ["🔗 Đường link / Văn bản", "📶 Wi-Fi", "👤 Danh thiếp (vCard)", "🏦 Chuyển khoản VietQR",
                 "✉️ Email", "💬 Tin nhắn SMS", "📞 Số điện thoại", "📍 Vị trí bản đồ", "📅 Sự kiện / Lịch họp"]
        kind = st.selectbox("Bạn muốn mã QR chứa nội dung gì?", kinds)
        data = ""
        err = None
        if kind == kinds[0]:
            data = st.text_area("Nhập đường link hoặc nội dung văn bản:", height=100, placeholder="https://...")
            if data.strip().lower().startswith("www."):
                data = "https://" + data.strip()
        elif kind == kinds[1]:
            c1, c2 = st.columns(2)
            ssid = c1.text_input("Tên Wi-Fi (SSID)")
            pw = c2.text_input("Mật khẩu", type="password")
            sec = c1.selectbox("Bảo mật", ["WPA/WPA2/WPA3", "WEP", "Không mật khẩu"])
            hid = c2.checkbox("Wi-Fi ẩn")
            if ssid:
                t = {"WPA/WPA2/WPA3": "WPA", "WEP": "WEP", "Không mật khẩu": "nopass"}[sec]
                data = f"WIFI:T:{t};S:{_esc_wifi(ssid)};" + (f"P:{_esc_wifi(pw)};" if t != "nopass" else "") + ("H:true;" if hid else "") + ";"
        elif kind == kinds[2]:
            c1, c2 = st.columns(2)
            name = c1.text_input("Họ và tên")
            org = c2.text_input("Đơn vị / Công ty")
            title = c1.text_input("Chức vụ")
            phone = c2.text_input("Điện thoại")
            email = c1.text_input("Email")
            web = c2.text_input("Website")
            addr = st.text_input("Địa chỉ")
            if name:
                last, _, first = name.strip().rpartition(" ")
                lines = ["BEGIN:VCARD", "VERSION:3.0", f"N:{_esc_vc(last)};{_esc_vc(first or name)};;;", f"FN:{_esc_vc(name)}"]
                for k, v in (("ORG", org), ("TITLE", title), ("TEL;TYPE=CELL", phone), ("EMAIL", email), ("URL", web)):
                    if v:
                        lines.append(f"{k}:{_esc_vc(v)}")
                if addr:
                    lines.append(f"ADR:;;{_esc_vc(addr)};;;;")
                lines.append("END:VCARD")
                data = "\n".join(lines)
        elif kind == kinds[3]:
            c1, c2 = st.columns(2)
            bank = c1.selectbox("Ngân hàng", list(QR_BANKS.keys()) + ["Khác (nhập mã BIN)"])
            bin_code = QR_BANKS.get(bank) or c2.text_input("Mã BIN ngân hàng (6 số)")
            acc = c1.text_input("Số tài khoản")
            holder = c2.text_input("Tên chủ tài khoản (hiển thị dưới mã)")
            amount = c1.number_input("Số tiền (0 = để người chuyển tự nhập)", min_value=0, step=10000, value=0)
            note = c2.text_input("Nội dung chuyển khoản (tối đa 25 ký tự)", max_chars=25)
            if acc and bin_code and bin_code.isdigit():
                data = qr_vietqr_payload(bin_code, acc.strip(), amount, note)
            st.caption("Mã theo chuẩn VietQR, quét được bằng app ngân hàng. Hãy chuyển thử 1 khoản nhỏ để kiểm tra trước khi gửi cho nhiều người.")
        elif kind == kinds[4]:
            to = st.text_input("Gửi đến (email)")
            sub = st.text_input("Tiêu đề")
            body = st.text_area("Nội dung", height=80)
            from urllib.parse import quote
            if to:
                data = f"mailto:{to}?subject={quote(sub)}&body={quote(body)}"
        elif kind == kinds[5]:
            ph = st.text_input("Số điện thoại nhận")
            msg = st.text_area("Nội dung tin nhắn", height=80)
            if ph:
                data = f"SMSTO:{ph}:{msg}"
        elif kind == kinds[6]:
            ph = st.text_input("Số điện thoại")
            if ph:
                data = f"tel:{ph}"
        elif kind == kinds[7]:
            mode = st.radio("Nhập theo:", ["Tọa độ", "Link Google Maps"], horizontal=True)
            if mode == "Tọa độ":
                c1, c2 = st.columns(2)
                lat = c1.text_input("Vĩ độ (VD: 10.7769)")
                lon = c2.text_input("Kinh độ (VD: 106.7009)")
                if lat and lon:
                    data = f"geo:{lat.strip()},{lon.strip()}"
            else:
                data = st.text_input("Dán link Google Maps")
        elif kind == kinds[8]:
            c1, c2 = st.columns(2)
            title = c1.text_input("Tên sự kiện")
            loc = c2.text_input("Địa điểm")
            d1 = c1.date_input("Ngày bắt đầu")
            t1 = c1.time_input("Giờ bắt đầu", value=_dt(2000, 1, 1, 8, 0).time())
            d2 = c2.date_input("Ngày kết thúc", value=d1)
            t2 = c2.time_input("Giờ kết thúc", value=_dt(2000, 1, 1, 9, 0).time())
            desc = st.text_area("Mô tả", height=70)
            if title:
                f = lambda d, t: _dt.combine(d, t).strftime("%Y%m%dT%H%M%S")
                data = ("BEGIN:VCALENDAR\nVERSION:2.0\nBEGIN:VEVENT\n" f"SUMMARY:{_esc_vc(title)}\n"
                        f"DTSTART:{f(d1, t1)}\nDTEND:{f(d2, t2)}\n" + (f"LOCATION:{_esc_vc(loc)}\n" if loc else "")
                        + (f"DESCRIPTION:{_esc_vc(desc)}\n" if desc else "") + "END:VEVENT\nEND:VCALENDAR")

        opts = style_panel("one")
        if kind == kinds[3] and not opts["caption"] and locals().get("holder"):
            opts["caption"] = holder.upper()

        st.markdown("---")
        if not data.strip():
            st.info("👆 Hãy nhập thông tin ở trên, mã QR sẽ hiện ra tại đây.")
        else:
            try:
                img = qr_render(data, opts["ecc"], opts["box"], 4, opts["fg"], opts["fg2"], opts["bg"],
                                opts["transparent"], opts["shape"], opts["logo"], opts["caption"])
                cl, cr = st.columns([1, 1])
                with cl:
                    st.image(img, caption=f"Kích thước: {img.width}×{img.height}px", use_container_width=True)
                with cr:
                    st.success("✅ Đã tạo mã QR. Nên dùng điện thoại quét thử trước khi gửi.")
                    if len(data) > 1200:
                        st.warning("Nội dung khá dài nên mã QR sẽ rất dày, khó quét. Hãy rút gọn nội dung hoặc dùng link ngắn.")
                    fname = "QR_" + _dt.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button("📥 Tải PNG", qr_to_png_bytes(img), f"{fname}.png", "image/png", type="primary", use_container_width=True)
                    st.download_button("📥 Tải SVG (in ấn, phóng to không vỡ)",
                                       qr_svg(data, opts["ecc"], 4, opts["fg"], opts["bg"], opts["transparent"], opts["shape"]),
                                       f"{fname}.svg", "image/svg+xml", use_container_width=True)
                    st.caption("SVG chỉ giữ 1 màu, không kèm logo và chữ.")
            except Exception as e:
                st.error(f"Không tạo được mã QR: {e}")

    # ---------------- TAB 2 ----------------
    with tab_batch:
        st.markdown("Tải file **Excel/CSV**, chọn cột chứa nội dung (link, mã sản phẩm, số tài sản...). Mỗi dòng sẽ ra **1 mã QR**, đóng gói thành file ZIP.")
        up = st.file_uploader("Tải lên file Excel (.xlsx) hoặc CSV:", type=["xlsx", "csv"], key=f"app11_{st.session_state.uploader_key}")
        if up:
            try:
                df = pd.read_csv(up, dtype=str) if up.name.lower().endswith(".csv") else pd.read_excel(up, dtype=str)
                df = df.fillna("")
                st.dataframe(df.head(8), use_container_width=True)
                st.caption(f"Tổng cộng {len(df)} dòng.")
                cols = list(df.columns)
                c1, c2 = st.columns(2)
                col_data = c1.selectbox("Cột chứa NỘI DUNG mã QR:", cols)
                col_name = c2.selectbox("Cột dùng đặt TÊN file (và chữ dưới mã):", ["(Số thứ tự)"] + cols)
                with_cap = st.checkbox("In chữ ở cột tên dưới mỗi mã QR", value=True)
                opts_b = style_panel("bat")
                if len(df) > 500:
                    st.warning("Trên 500 dòng sẽ xử lý khá lâu, nên chia nhỏ file.")
                if st.button("🚀 Tạo hàng loạt", type="primary"):
                    bar = st.progress(0.0)
                    zbuf, used, fails = BytesIO(), set(), []
                    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for i, row in df.reset_index(drop=True).iterrows():
                            val = str(row[col_data]).strip()
                            if not val:
                                fails.append(f"Dòng {i + 2}: nội dung trống")
                                continue
                            label = str(row[col_name]).strip() if col_name != "(Số thứ tự)" else f"{i + 1:03d}"
                            safe = re.sub(r'[\\/:*?"<>|\n\r]+', "_", label)[:60] or f"{i + 1:03d}"
                            base, k = safe, 1
                            while safe in used:
                                k += 1
                                safe = f"{base}_{k}"
                            used.add(safe)
                            try:
                                im = qr_render(val, opts_b["ecc"], opts_b["box"], 4, opts_b["fg"], opts_b["fg2"], opts_b["bg"],
                                               opts_b["transparent"], opts_b["shape"], opts_b["logo"],
                                               label if (with_cap and col_name != "(Số thứ tự)") else "")
                                zf.writestr(f"{safe}.png", qr_to_png_bytes(im))
                            except Exception as e:
                                fails.append(f"Dòng {i + 2}: {str(e)[:60]}")
                            bar.progress((i + 1) / len(df))
                    bar.empty()
                    st.success(f"🎉 Đã tạo {len(used) - 0} mã QR.")
                    for f_ in fails[:20]:
                        st.warning(f_)
                    st.download_button("📥 Tải ZIP các mã QR", zbuf.getvalue(), "QR_HangLoat.zip", "application/zip", type="primary")
            except Exception as e:
                st.error(f"Lỗi đọc file: {e}")

# ==========================================                        
# 5. KÍCH HOẠT ỨNG DỤNG THEO LỰA CHỌN MENU
# ==========================================
if "1. PDF sang Word" in app_mode:
    app_pdf_to_word()
elif "2. Chuyển PDF" in app_mode:
    app_number_2()
elif "3. PDF/Ảnh sang Excel" in app_mode:
    app_number_3()
elif "4. So sánh Văn bản" in app_mode:
    app_document_compare()
elif "5. Cắt & Ghép PDF" in app_mode:
    app_pdf_split_merge()
elif "6. Chuyên gia Công thức" in app_mode:
    app_excel_expert()
elif "7. Nén & Đóng Gói" in app_mode:
    app_compress_and_zip()
elif "8. Xoay Trang PDF" in app_mode:
    app_rotate_pdf()
elif "9. Xóa Siêu Dữ Liệu" in app_mode:
    app_metadata_cleaner()
elif "10. Tách Nền Chữ Ký" in app_mode:
    app_signature_extractor()
elif "11. Tạo Mã QR" in app_mode:
    app_qr_generator()
elif "12. PDF sang Word" in app_mode:
    app_pdf_to_word_offline()
