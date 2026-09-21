"""Đối chiếu Smile với trang quản lý lưu trú người nước ngoài."""

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

from .common import _norm_pp, _norm_room

def reconcile(smile_bytes, luutru_bytes, today):
    """Đối chiếu file Smile (inhouse) với file trang quản lý lưu trú.
    today: pd.Timestamp ngày xuất file (hôm nay)."""
    # ── Đọc Smile ──
    df1 = pd.read_excel(io.BytesIO(smile_bytes), header=0)
    smile = df1[['Passport #','NAT','Rm#','Arrival','Last Name','First Name']].copy()
    smile = smile.dropna(subset=['Passport #'])
    smile['pp'] = smile['Passport #'].apply(_norm_pp)
    smile['Arrival'] = pd.to_datetime(smile['Arrival'], errors='coerce')
    smile['name'] = (smile['Last Name'].astype(str).str.strip() + ' ' +
                     smile['First Name'].astype(str).str.strip())
    smile['room'] = smile['Rm#'].apply(_norm_room)
    # Cột Departure (dò tên linh hoạt phòng khi khác tên)
    _dep_col = next((c for c in df1.columns if 'depart' in str(c).lower()), None)
    smile['Departure'] = pd.to_datetime(df1[_dep_col], errors='coerce') if _dep_col else pd.NaT
    # Lọc: bỏ VNM + bỏ arrival = hôm nay + bỏ departure = hôm nay (khách trả phòng)
    smile_f = smile[(smile['NAT'] != 'VNM') &
                    (smile['Arrival'].dt.date != today.date()) &
                    (smile['Departure'].dt.date != today.date())].copy()

    # ── Đọc Lưu trú ──
    df2 = pd.read_excel(io.BytesIO(luutru_bytes), header=9)
    df2 = df2.dropna(subset=['Họ tên'])
    df2['pp'] = df2['Số hộ chiếu'].apply(_norm_pp)
    # Cột ngày đi dự kiến: file mới đổi tên thành "Thời gian dự kiến tạm trú tại CSLT".
    # Dò linh hoạt để chạy được cả file mẫu mới lẫn cũ.
    _ddk_col = next((c for c in df2.columns
                     if 'dự kiến' in str(c).lower() and 'tạm trú' in str(c).lower()), None)
    if _ddk_col is None:
        _ddk_col = next((c for c in df2.columns if 'đi dự kiến' in str(c).lower()), None)
    df2['ddk'] = pd.to_datetime(df2[_ddk_col], format='%d/%m/%Y', errors='coerce') if _ddk_col else pd.NaT
    df2['room'] = df2['Số phòng'].apply(_norm_room)
    # Lọc: bỏ ngày đi dự kiến = hôm nay
    luutru_f = df2[df2['ddk'].dt.date != today.date()].copy()

    smile_pp = set(smile_f['pp'])
    luutru_pp = set(luutru_f['pp'])

    # Đối chiếu người
    chua_dk = smile_f[~smile_f['pp'].isin(luutru_pp)][['name','pp','NAT','room','Arrival']].copy()
    chua_dk.columns = ['Họ tên','Số hộ chiếu','Quốc tịch','Số phòng','Ngày đến']
    chua_dk['Ngày đến'] = chua_dk['Ngày đến'].dt.strftime('%d/%m/%Y')

    thua = luutru_f[~luutru_f['pp'].isin(smile_pp)].copy()
    # Giai đoạn lưu trú: Ngày đến → Ngày đi dự kiến
    _dd = pd.to_datetime(thua['Ngày đến '], format='%d/%m/%Y', errors='coerce') if 'Ngày đến ' in thua else pd.Series([pd.NaT]*len(thua))
    thua['_luutru'] = (_dd.dt.strftime('%d/%m/%Y').fillna('?') + ' → ' +
                       thua['ddk'].dt.strftime('%d/%m/%Y').fillna('?'))
    _qt = thua['QT'] if 'QT' in thua else ''
    thua = pd.DataFrame({
        'Họ tên': thua['Họ tên'].values,
        'Số hộ chiếu': thua['pp'].values,
        'Quốc tịch': _qt.values if hasattr(_qt,'values') else _qt,
        'Số phòng': thua['room'].values,
        'Giai đoạn lưu trú': thua['_luutru'].values,
    })

    dup = luutru_f[luutru_f['pp'].duplicated(keep=False) & (luutru_f['pp']!='')]
    dup = dup[['Họ tên','pp','room']].sort_values('pp').copy()
    dup.columns = ['Họ tên','Số hộ chiếu','Số phòng']

    # Đối chiếu phòng
    smile_rooms = set(r for r in smile_f['room'] if r)
    luutru_rooms = set(r for r in luutru_f['room'] if r)
    def _sortkey(x): return (len(x), x)
    room_chua = sorted(smile_rooms - luutru_rooms, key=_sortkey)
    room_thua = sorted(luutru_rooms - smile_rooms, key=_sortkey)

    return {
        'smile_total': len(smile), 'smile_filtered': len(smile_f),
        'luutru_total': len(df2), 'luutru_filtered': len(luutru_f),
        'chua_dk': chua_dk, 'thua': thua, 'dup': dup,
        'room_chua': room_chua, 'room_thua': room_thua,
        'room_match': len(smile_rooms & luutru_rooms),
        'smile_rooms': len(smile_rooms), 'luutru_rooms': len(luutru_rooms),
    }




# Ten duoc module nay so huu — liet ke tuong minh de `import *` lay
# duoc ca helper co gach duoi dau.
__all__ = [
    'reconcile'
]
