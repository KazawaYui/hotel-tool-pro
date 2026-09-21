"""Tiện ích dùng chung: giờ Việt Nam, đọc ô Excel, chuẩn hoá phòng/tên/hộ chiếu/ngày.

Gom về đây vì được dùng xuyên suốt mọi khâu nghiệp vụ — để rải rác ở
module chuyên biệt sẽ tạo phụ thuộc chéo vô nghĩa."""

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


VN_TZ = datetime.timezone(datetime.timedelta(hours=7))
def now_vn():
    return datetime.datetime.now(VN_TZ)
def today_vn():
    return now_vn().date()

def _strip_accents(s):
    """Bỏ dấu tiếng Việt, giữ lại đúng chữ cái gốc.

    Đ/đ KHÔNG phải chữ D kèm dấu phụ mà là ký tự riêng (U+0110/U+0111) nên
    NFD không tách ra được. Nếu chỉ NFD rồi lọc [^a-z0-9], chữ đ không lọt
    qua bộ lọc và bị XOÁ HẲN: "Đà Nẵng" → "a nang", "Ấn Độ" → "ano". Phải
    đổi tay Đ→D trước khi chuẩn hoá."""
    s = str(s or '').replace('Đ', 'D').replace('đ', 'd')
    s = _ud.normalize('NFD', s)
    return ''.join(c for c in s if _ud.category(c) != 'Mn')

# ── Helpers ───────────────────────────────────────────────────────────────
def fmt(v):
    if v is None or str(v) in ('NaT','nan',''): return ''
    if hasattr(v,'strftime'): return v.strftime('%d/%m/%Y')
    return str(v).strip()[:10]

def make_code(prefix, ns):
    if not ns: return prefix
    p = ns.replace('-','/').split('/')
    return f"{prefix}{p[0].zfill(2)}{p[1].zfill(2)}{p[2][-2:]}" if len(p)==3 else prefix

def snapshot_styles(cells):
    """Chụp định dạng của một dòng ô mẫu thành các đối tượng style THẬT.

    cell.font trả về StyleProxy — không hash được nên openpyxl từ chối đưa
    thẳng vào bảng style dùng chung. copy() ở đây vừa gỡ lớp proxy vừa tách
    khỏi template gốc, và chỉ chạy MỘT LẦN cho cả file thay vì mỗi ô."""
    return [(copy(c.font), copy(c.fill), copy(c.border), copy(c.alignment),
             c.number_format) for c in cells]


def apply_style(cell, style):
    """Dán định dạng đã chụp sẵn vào ô.

    Gán thẳng, không copy lại: openpyxl lưu style vào bảng dùng chung của
    workbook rồi cho ô giữ chỉ số, nên nhiều ô trỏ cùng một đối tượng là
    cách dùng bình thường. Copy từng ô chỉ tạo hàng nghìn đối tượng y hệt
    nhau rồi vứt đi — với file 124 khách là gần 12.000 lệnh copy thừa."""
    f, fl, b, al, nf = style
    if f: cell.font = f
    if fl: cell.fill = fl
    if b: cell.border = b
    if al: cell.alignment = al
    cell.number_format = nf


def cp(src, dst):
    """Chép định dạng giữa hai ô. Dùng cho vài ô lẻ; điền hàng loạt thì dùng
    snapshot_styles() + apply_style() để không copy lại từng ô."""
    for a in ('font', 'fill', 'border', 'alignment'):
        v = getattr(src, a)
        if v:
            setattr(dst, a, copy(v))
    dst.number_format = src.number_format

def serial2date(s):
    if not s: return None
    try: return datetime.datetime(1899,12,30)+datetime.timedelta(days=int(s))
    except: return None

def wb_to_bytes(wb):
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def _gv(row, *names):
    """Lấy giá trị đầu tiên khác rỗng theo danh sách tên cột (chịu biến thể tên cột)."""
    for n in names:
        if n in row.index:
            v = row[n]
            if pd.notna(v) and str(v).strip() != '':
                return v
    return None

def _fmt_room(v):
    """Số phòng/số điện thoại đọc từ Excel đôi khi ra dạng số thực (103.0)
    → trả về '103'. Dùng `v or ''` sẽ SAI với NaN (NaN truthy trong Python)
    → lọt chuỗi "nan" ra file; phải chặn None/NaN riêng trước."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    s = str(v).strip()
    if s.endswith('.0') and s[:-2].isdigit():
        s = s[:-2]
    return s

def _to_date(v):
    """Đọc ngày linh hoạt (datetime / chuỗi dd/mm/yyyy) → datetime.date hoặc None."""
    if v is None:
        return None
    if hasattr(v, 'year') and not isinstance(v, str):
        try:
            return datetime.date(v.year, v.month, v.day)
        except Exception:
            return None
    t = pd.to_datetime(str(v).strip(), dayfirst=True, errors='coerce')
    return None if pd.isna(t) else t.date()


def _norm_name(s):
    """Chuẩn hóa tên để khớp giữa 2 file: bỏ dấu, hoa thường, gộp khoảng trắng.
    Cả hai phía đều chuẩn hoá lúc chạy bằng hàm này nên luôn khớp nhau."""
    s = _strip_accents(s).lower().strip()
    return _re.sub(r'\s+', ' ', _re.sub(r'[^a-z0-9 ]', '', s)).strip()


def _fix_date(v):
    """Chuẩn hóa ngày về pd.Timestamp đúng nghĩa dd/mm.

    File Smile export bị lỗi: ngày dạng dd/mm với ngày ≤ 12 (vd '7/8/2026' = 7 tháng 8)
    bị Excel hiểu nhầm kiểu Mỹ mm/dd → lưu thành datetime(month=7, day=8) kèm format
    mm-dd-yy. Khi đọc lại, cần HOÁN month↔day để khôi phục: datetime(y,7,8) → 7 tháng 8.
    Ngày ≥ 13 thì Excel không nhầm được nên giữ dạng text dd/mm bình thường.
    """
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    # datetime từ Excel → đã bị đảo month/day, khôi phục bằng cách hoán lại
    if hasattr(v, 'year') and not isinstance(v, str):
        try:
            return pd.Timestamp(year=v.year, month=v.day, day=v.month)
        except Exception:
            return pd.Timestamp(v)   # day>12: không đảo được, giữ nguyên
    s = str(v).strip()
    if '/' in s:
        p = s.split('/')
        if len(p) == 3:
            dd, mm, yy = p
            if len(yy) == 2: yy = '20' + yy
            try:
                return pd.Timestamp(year=int(yy), month=int(mm), day=int(dd))
            except Exception:
                return None
    try:
        return pd.to_datetime(s, dayfirst=True)
    except Exception:
        return None

def _fix_departure_swap(ni_str, nd_str):
    """Nếu NGÀY ĐI (dd/mm/yyyy) đọc ra TRƯỚC NGÀY ĐẾN — thường do lỗi đảo
    dd/mm khi cả 2 số ≤12 — thử hoán dd↔mm; dùng bản đã hoán nếu nó không còn
    trước ngày đến nữa. Không đụng các trường hợp khác (giữ nguyên nếu không
    chắc chắn). Cùng heuristic đã kiểm chứng ở ARR/XNC Converter."""
    if not ni_str or not nd_str:
        return ni_str
    try:
        dep = datetime.datetime.strptime(ni_str, '%d/%m/%Y')
        arr = datetime.datetime.strptime(nd_str, '%d/%m/%Y')
    except Exception:
        return ni_str
    if dep >= arr:
        return ni_str
    m = _re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', ni_str)
    if not m:
        return ni_str
    dd, mm, yy = int(m.group(1)), int(m.group(2)), m.group(3)
    if dd > 12 or mm > 12 or dd == mm:
        return ni_str
    swapped = f"{mm:02d}/{dd:02d}/{yy}"
    try:
        sw = datetime.datetime.strptime(swapped, '%d/%m/%Y')
    except Exception:
        return ni_str
    return swapped if sw >= arr else ni_str


def _norm_pp(p):
    """Chuẩn hóa số hộ chiếu."""
    if pd.isna(p): return ''
    s = str(p).strip().upper().replace(' ','')
    if s.endswith('.0'): s = s[:-2]
    return s

def _norm_room(r):
    if pd.isna(r): return ''
    s = str(r).strip()
    if s.endswith('.0'): s = s[:-2]
    return s.upper()  # 12a05 và 12A05 là một phòng


# Ten duoc module nay so huu — liet ke tuong minh de `import *` lay
# duoc ca helper co gach duoi dau.
__all__ = [
    'apply_style', 'snapshot_styles', 'VN_TZ', '_fix_date', '_fix_departure_swap', '_fmt_room', '_gv', '_norm_name',
    '_norm_pp', '_norm_room', '_strip_accents', '_to_date', 'cp', 'fmt', 'make_code',
    'now_vn', 'serial2date', 'today_vn', 'wb_to_bytes'
]
