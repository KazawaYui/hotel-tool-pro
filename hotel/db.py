"""Sổ giao ca lưu trên Supabase/Postgres, bền vững qua mọi lần deploy."""

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


# ── Sổ giao ca — lưu trữ đám mây (Supabase/Postgres), BỀN VỮNG qua mọi lần
# deploy lại (khác với data/progress_*.json ở trên chỉ sống qua 1 ngày). Cần
# cấu hình st.secrets["connections"]["supabase_db"]["url"] — xem hướng dẫn
# trong secrets.toml.example. NẾU CHƯA CẤU HÌNH: mọi hàm db_* trả về None/rỗng
# một cách an toàn, màn Sổ giao ca tự động dùng lại lưu tạm trên đĩa (không
# bền vững qua deploy) — app KHÔNG bao giờ crash vì thiếu Supabase.
try:
    from sqlalchemy import text as _sql_text
except Exception:
    _sql_text = None

def _redact_db_error(e):
    """Ẩn mật khẩu (nếu lỡ lọt vào chuỗi kết nối trong thông báo lỗi của
    SQLAlchemy) trước khi hiển thị cho người dùng — tránh lộ secret lên UI."""
    import re as _re_local
    return _re_local.sub(r'://([^:/@\s]+):([^@\s]+)@', r'://\1:***@', str(e))

def _get_db():
    if _sql_text is None:
        st.session_state['_db_last_error'] = ("Thiếu thư viện SQLAlchemy/psycopg2-binary trong môi trường chạy — "
                                               "vào Manage app → Reboot app để cài lại requirements.txt.")
        return None
    try:
        return st.connection("supabase_db", type="sql")
    except Exception as e:
        st.session_state['_db_last_error'] = _redact_db_error(e)
        return None

@st.cache_resource(show_spinner=False)
def _db_schema_ready():
    """Tạo bảng shift_handover nếu chưa có — chỉ chạy 1 lần mỗi phiên server
    (cache_resource), không phải mỗi lần rerun."""
    conn = _get_db()
    if conn is None:
        return False
    try:
        with conn.session as s:
            s.execute(_sql_text("""
                CREATE TABLE IF NOT EXISTS shift_handover (
                    id BIGSERIAL PRIMARY KEY,
                    shift_date DATE NOT NULL,
                    entry_time TEXT NOT NULL,
                    category TEXT NOT NULL,
                    room TEXT,
                    note TEXT NOT NULL,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
            """))
            s.execute(_sql_text(
                "CREATE INDEX IF NOT EXISTS idx_shift_handover_date ON shift_handover(shift_date)"))
            s.commit()
        return True
    except Exception as e:
        st.session_state['_db_last_error'] = _redact_db_error(e)
        return False

def db_available():
    return _get_db() is not None and _db_schema_ready()

def db_add_entry(shift_date, entry_time, category, room, note):
    conn = _get_db()
    with conn.session as s:
        s.execute(_sql_text("""INSERT INTO shift_handover (shift_date, entry_time, category, room, note)
                               VALUES (:d, :t, :c, :r, :n)"""),
                  {'d': shift_date, 't': entry_time, 'c': category, 'r': room, 'n': note})
        s.commit()

def db_delete_entry(entry_id):
    conn = _get_db()
    with conn.session as s:
        s.execute(_sql_text("DELETE FROM shift_handover WHERE id = :id"), {'id': int(entry_id)})
        s.commit()

def db_load_entries(shift_date):
    """Trả về DataFrame [id, entry_time, category, room, note] cho 1 ngày, mới nhất trước."""
    conn = _get_db()
    return conn.query(
        "SELECT id, entry_time, category, room, note FROM shift_handover "
        "WHERE shift_date = :d ORDER BY entry_time DESC, id DESC",
        params={'d': shift_date}, ttl=0)

def db_load_dates(limit=180):
    """Danh sách các ngày đã có ghi chú (mới nhất trước) — phục vụ ô chọn ngày xem lại lịch sử."""
    conn = _get_db()
    df = conn.query(
        "SELECT DISTINCT shift_date FROM shift_handover ORDER BY shift_date DESC LIMIT :lim",
        params={'lim': limit}, ttl=0)
    return list(df['shift_date']) if not df.empty else []

def compute_day_summary(df_entries):
    """Tổng hợp tự động: đếm ghi chú theo phân loại + danh sách phòng được nhắc tới."""
    if df_entries is None or df_entries.empty:
        return {'total': 0, 'by_category': {}, 'rooms': []}
    by_cat = df_entries['category'].value_counts().to_dict()
    rooms = sorted(set(str(r).strip() for r in df_entries['room'].dropna() if str(r).strip()))
    return {'total': len(df_entries), 'by_category': by_cat, 'rooms': rooms}

# Load app icon (favicon) từ icon.b64

# Ten duoc module nay so huu — liet ke tuong minh de `import *` lay
# duoc ca helper co gach duoi dau.
__all__ = [
    '_db_schema_ready', '_get_db', '_redact_db_error', 'compute_day_summary',
    'db_add_entry', 'db_available', 'db_delete_entry', 'db_load_dates',
    'db_load_entries'
]
