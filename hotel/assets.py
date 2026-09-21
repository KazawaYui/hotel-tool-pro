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

@st.cache_resource
def _dark_bg_data_uri():
    """Ảnh nền chế độ tối (mèo con ngủ) — nhúng thẳng base64 vào CSS, cùng
    kiểu với load_template() ở trên, không cần hosting/URL ngoài."""
    path = os.path.join(ASSET_DIR, 'bg_dark.b64')
    with open(path, 'r') as f:
        return 'data:image/jpeg;base64,' + f.read().strip()

@st.cache_resource
def _light_bg_data_uri():
    """Ảnh nền chế độ sáng (mèo con chui trong túi giấy) — cùng cơ chế với
    _dark_bg_data_uri() ở trên."""
    path = os.path.join(ASSET_DIR, 'bg_light.b64')
    with open(path, 'r') as f:
        return 'data:image/jpeg;base64,' + f.read().strip()

@st.cache_resource
def _season_bg_data_uri(key, theme):
    """Ảnh nền màn chào theo mùa/dịp lễ (bg_<key>.b64). Mùa nào CHƯA có ảnh
    riêng thì dùng lại ảnh nền sáng/tối sẵn có — không bao giờ để trống."""
    path = os.path.join(ASSET_DIR, f'bg_{key}.b64')
    if os.path.exists(path):
        try:
            with open(path, 'r') as f:
                return 'data:image/jpeg;base64,' + f.read().strip()
        except Exception:
            pass
    return _dark_bg_data_uri() if theme == 'dark' else _light_bg_data_uri()


# Ten duoc module nay so huu — liet ke tuong minh de `import *` lay
# duoc ca helper co gach duoi dau.
__all__ = [
    '_dark_bg_data_uri', '_light_bg_data_uri', '_load_app_icon', '_season_bg_data_uri',
    'load_template'
]
