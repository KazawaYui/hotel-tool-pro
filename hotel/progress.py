"""Tiến độ ca làm việc lưu trên đĩa server, sống sót qua tải lại trang."""

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

from .common import now_vn, today_vn

# ── Tiến độ ca làm việc — lưu trên đĩa server để SỐNG SÓT qua việc tải lại
# trang (F5) trong ngày. LƯU Ý: file này KHÔNG bền vững qua các lần deploy lại
# app (Streamlit Cloud xóa filesystem mỗi lần deploy) — chỉ chống việc mất dữ
# liệu do reload trang trong 1 ngày làm việc, không thay thế backup lâu dài.
# Chỉ lưu tiến độ + số liệu TỔNG HỢP (không tên/hộ chiếu khách) cho các công cụ
# xử lý dữ liệu khách; RIÊNG Sổ giao ca lưu đầy đủ nội dung vì đó chính là mục
# đích của sổ giao ca (thông tin cần truyền lại nguyên vẹn cho ca sau).
# Gốc dự án, KHÔNG phải dirname(__file__): bám theo __file__ sẽ ghi vào
# hotel/data/ — vừa lệch chỗ so với dữ liệu ca đang có, vừa lọt khỏi mục
# data/ trong .gitignore nên tiến độ ca có thể bị commit nhầm lên kho.
DATA_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')

def _progress_path(date_iso=None):
    os.makedirs(DATA_DIR, exist_ok=True)
    return os.path.join(DATA_DIR, f'progress_{date_iso or today_vn().isoformat()}.json')

def _default_progress():
    return {'date': today_vn().isoformat(), 'tasks': {}, 'handover_entries': []}

def _load_progress():
    p = _progress_path()
    if os.path.exists(p):
        try:
            with open(p, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if data.get('date') == today_vn().isoformat():
                return data
        except Exception:
            pass
    return _default_progress()

def _yesterday_total():
    """Tổng khách HÔM QUA đọc từ file tiến độ ngày hôm trước — dùng để so sánh
    trên thẻ tổng quan. Không có file (hoặc hôm qua chưa chạy công cụ) thì trả
    None và thẻ sẽ không hiện dòng so sánh, KHÔNG suy đoán số."""
    try:
        p = _progress_path((today_vn() - datetime.timedelta(days=1)).isoformat())
        if os.path.exists(p):
            with open(p, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return ((data.get('tasks', {}) or {}).get('daily', {}) or {}).get('summary', {}).get('total')
    except Exception:
        pass
    return None

def _atomic_write_json(path, data):
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def _progress_update(mutate_fn):
    """Đọc file tiến độ MỚI NHẤT từ đĩa (không phải bản trong session_state) rồi
    mới sửa và ghi lại — giảm rủi ro 2 tab/phiên trong ngày ghi đè mất dữ liệu
    của nhau. Đồng bộ luôn vào session_state để hiển thị ngay trong lượt chạy
    hiện tại."""
    state = _load_progress()
    mutate_fn(state)
    state['last_updated'] = now_vn().strftime('%H:%M:%S')
    try:
        _atomic_write_json(_progress_path(), state)
    except Exception:
        pass  # đĩa lỗi/không ghi được không được làm crash app — tính năng chỉ là tiện ích
    st.session_state.progress = state
    return state


# Ten duoc module nay so huu — liet ke tuong minh de `import *` lay
# duoc ca helper co gach duoi dau.
__all__ = [
    'DATA_DIR', '_atomic_write_json', '_default_progress', '_load_progress',
    '_progress_path', '_progress_update', '_yesterday_total'
]
