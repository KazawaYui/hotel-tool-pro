"""Mẫu file và ảnh nền nhúng sẵn dạng base64."""

import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from copy import copy
import xlrd, datetime, io, zipfile, base64, os, json
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.colors import white, black
import unicodedata as _ud, re as _re

# Các file .b64 nằm ở GỐC dự án, không nằm trong package — một chỗ duy nhất
# để sửa nếu sau này dọn asset vào hotel/. Dùng đường dẫn tính từ __file__
# chứ không phải thư mục hiện hành, để app chạy đúng dù khởi động từ đâu.
ASSET_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@st.cache_resource
def _load_app_icon():
    try:
        from PIL import Image
        p = os.path.join(ASSET_DIR, 'icon.b64')
        with open(p, 'r') as f:
            raw = base64.b64decode(f.read())
        img = Image.open(io.BytesIO(raw))
        img.load()  # ép tải đầy đủ ảnh ngay tại đây — tránh lỗi lazy-load khi
                    # st.cache_resource dùng lại ảnh này ở ngữ cảnh/luồng khác
        return img
    except Exception:
        return "🌸"


# ── Load embedded templates ──────────────────────────────────────────────
@st.cache_resource
def load_template(name):
    path = os.path.join(ASSET_DIR, f'tmpl_{name}.b64')
    with open(path, 'r') as f:
        return base64.b64decode(f.read())

# ── Ảnh nền: phục vụ từ static/ thay vì nhúng base64 vào CSS ─────────────
# Trước đây mỗi ảnh được nhúng thẳng dạng data: URI vào khối CSS, có ảnh
# lặp tới 3 lần → ~565 KB CSS. Inline CSS KHÔNG BAO GIỜ được trình duyệt
# cache nên toàn bộ chỗ đó phải tải lại mỗi lần mở trang. Ảnh phục vụ qua
# URL thì có ETag: lần sau trình duyệt hỏi 304 và không tải lại gì cả.
#
# Dùng đường dẫn TƯƠNG ĐỐI ('app/static/…' chứ không phải '/app/static/…')
# để app vẫn đúng khi deploy dưới một đường dẫn con.
STATIC_DIR = os.path.join(ASSET_DIR, 'static')

def _bg_url(name):
    """URL ảnh nền trong static/. Trả '' nếu thiếu file — CSS gặp url("")
    thì bỏ qua, nền lùi về màu nền phẳng chứ không vỡ layout."""
    return f'app/static/{name}.jpg' if os.path.exists(
        os.path.join(STATIC_DIR, f'{name}.jpg')) else ''

def _dark_bg_url():
    """Ảnh nền chế độ tối (mèo con ngủ)."""
    return _bg_url('bg_dark')

def _light_bg_url():
    """Ảnh nền chế độ sáng (mèo con chui trong túi giấy)."""
    return _bg_url('bg_light')

def _season_bg_url(key, theme):
    """Ảnh nền màn chào theo mùa/dịp lễ. Mùa nào CHƯA có ảnh riêng thì dùng
    lại ảnh nền sáng/tối sẵn có — không bao giờ để trống."""
    return _bg_url(f'bg_{key}') or (_dark_bg_url() if theme == 'dark'
                                    else _light_bg_url())


# Ten duoc module nay so huu — liet ke tuong minh de `import *` lay
# duoc ca helper co gach duoi dau.
__all__ = [
    'ASSET_DIR', 'STATIC_DIR', '_bg_url', '_dark_bg_url', '_light_bg_url',
    '_load_app_icon', '_season_bg_url', 'load_template'
]
