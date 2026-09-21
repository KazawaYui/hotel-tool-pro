"""Regcard PDF và file ARR."""

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

from .common import _fix_date
from .assets import ASSET_DIR
from .lookups import _norm_nat

# ── Regcard PDF builder ───────────────────────────────────────────────────
# Dùng ASSET_DIR (gốc dự án) chứ KHÔNG phải dirname(__file__): các file .b64
# nằm ở gốc, không nằm trong package — bám theo __file__ sẽ tìm nhầm vào
# hotel/ và làm hỏng toàn bộ công cụ Regcard.
@st.cache_resource
def load_regcard_template():
    path = os.path.join(ASSET_DIR, 'tmpl_regcard.b64')
    with open(path, 'r') as f:
        return base64.b64decode(f.read())

def load_group_template():
    path = os.path.join(ASSET_DIR, 'tmpl_group.b64')
    with open(path, 'r') as f:
        return base64.b64decode(f.read())

def _grp_date(d):
    """Ngày cho regcard group: dd/mm/yyyy (như mẫu 24/07/2026)."""
    if pd.isna(d): return ''
    if hasattr(d, 'strftime'):
        return f"{d.day:02d}/{d.month:02d}/{d.year}"
    s = str(d).strip()
    if '/' in s:
        p = s.split('/')
        if len(p) == 3:
            dd, mm, yy = p
            if len(yy) == 2: yy = '20' + yy
            return f"{int(dd):02d}/{int(mm):02d}/{yy}"
        return s
    # Chuỗi ISO yyyy-mm-dd
    try:
        t = pd.to_datetime(s)
        return f"{t.day:02d}/{t.month:02d}/{t.year}"
    except Exception:
        return s

def _rc_clean_name(n):
    if pd.isna(n): return ''
    return str(n).strip().rstrip(',').strip()

def _rc_conf(c):
    if pd.isna(c): return ''
    return str(int(c)) if isinstance(c,(int,float)) else str(c)

def _rc_date(d):
    if pd.isna(d): return ''
    if hasattr(d,'strftime') and not isinstance(d, str):
        return f"{d.day:02d}/{d.month:02d}/{d.year}"   # dd/mm/yyyy
    s=str(d).strip()
    if '/' in s:
        p=s.split('/')
        if len(p)==3:
            dd,mm,yy=p
            if len(yy)==2: yy='20'+yy
            return f"{int(dd):02d}/{int(mm):02d}/{yy}"   # pad 7/8 → 07/08
        return s
    try:
        t=pd.to_datetime(s)                            # ISO yyyy-mm-dd
        return f"{t.day:02d}/{t.month:02d}/{t.year}"
    except Exception:
        return s

def _rc_nights(arr, dep):
    """Số đêm = Departure - Arrival. Xử lý cả datetime lẫn chuỗi dd/mm/yyyy."""
    def _to_ts(v):
        if pd.isna(v): return None
        if hasattr(v, 'year') and not isinstance(v, str):
            return pd.Timestamp(v)          # datetime/Timestamp — giữ nguyên
        s = str(v).strip()
        if '/' in s:                        # chuỗi dd/mm/yyyy hoặc dd/mm/yy
            p = s.split('/')
            if len(p) == 3:
                dd, mm, yy = p
                if len(yy) == 2: yy = '20' + yy
                try:
                    return pd.Timestamp(year=int(yy), month=int(mm), day=int(dd))
                except Exception:
                    return None
        try:
            return pd.to_datetime(s)        # ISO yyyy-mm-dd hoặc dạng khác
        except Exception:
            return None
    a, d = _to_ts(arr), _to_ts(dep)
    if a is None or d is None:
        return ''
    n = (d - a).days
    return str(n) if n >= 0 else ''

# Đường kẻ DỌC của bảng trên mẫu regcard đoàn, dạng (x, y_trên, y_dưới) — đo
# bằng cách render mẫu ở 1px/1điểm rồi dò cột pixel đen theo từng hàng. Phải
# ghi theo hàng vì các hàng chia cột khác nhau: hàng Arrival/Departure cắt ở
# x=291, hai hàng dưới cắt ở x=155 và 357.
# Ô che dữ liệu mẫu phải nằm GỌN GIỮA hai đường — lấn ra là xoá mất khung.
_GROUP_TABLE_LINES = (
    (31, 160, 345),    # mép trái, suốt bảng
    (553, 160, 345),   # mép phải, suốt bảng
    (291, 160, 209),   # chỉ hàng Arrival / Departure
    (155, 209, 345),   # hai hàng dưới
    (357, 209, 345),   # hai hàng dưới
)

# Mẫu PDF có sẵn dữ liệu ví dụ in cứng (đoàn 21099 / TBA / VIRGO TRAVEL /
# 25 phòng / 52 khách / 08-09.08.2026). Không che thì chữ mới vẽ ĐÈ LÊN chữ
# cũ, phiếu in ra chồng hai lớp số đọc không ra — regcard cá nhân đã có bước
# che này từ đầu, regcard đoàn thì bị bỏ sót.
# Toạ độ (x0, trên, x1, dưới), nới tới sát mép ô để chứa được giá trị dài
# hơn bản mẫu.
_GROUP_BLANK = [
    (156.3, 166, 290, 184),   # Arrival Date      (ô 155 → 292)
    (433, 166, 551, 184),     # Departure Date    (ô 292 → 553)
    (33, 246, 153, 266),      # Group Code        (ô  31 → 155)
    (157, 246, 355, 267),     # Group Name        (ô 155 → 357)
    (359, 246, 551, 267),     # Travel Agent      (ô 357 → 553)
    (33, 312, 153, 333),      # No of rooms       (ô  31 → 155)
    (157, 310, 355, 330),     # No of pax         (ô 155 → 357)
]


def build_group_regcard(grp_df, tmpl_bytes):
    """Vẽ 1 Registration Card for Group từ các dòng cùng 1 mã Group.
    Trả về trang PDF đã merge. Bảng Kind of rooms để trống (điền tay)."""
    H = 841.0
    FONT = "Times-Roman"; SIZE = 11
    first = grp_df.iloc[0]

    # Group Code = giá trị cột Group
    gcode = first.get('Group')
    if pd.notna(gcode):
        gcode = str(int(gcode)) if isinstance(gcode, (int, float)) else str(gcode)
    else:
        gcode = ''

    # Số phòng: đếm phòng unique, loại phòng ảo 9xxx
    rooms = set()
    for r in grp_df['Rm'].dropna():
        s = str(r).strip()
        if s.endswith('.0'): s = s[:-2]
        if s and not _re.fullmatch(r'9\d{3}', s):
            rooms.add(s.upper())
    n_rooms = len(rooms)

    # Số pax = tổng Adt + Chl + Enf
    def _sum(col):
        return int(grp_df[col].fillna(0).sum()) if col in grp_df.columns else 0
    n_pax = _sum('Adt') + _sum('Chl') + _sum('Enf')

    data = {
        'arrival':   (155.6, 181.2, _grp_date(first.get('Arrival'))),
        'departure': (436.9, 181.2, _grp_date(first.get('Departure'))),
        'gcode':     (48.2,  261.6, gcode),
        'gname':     (165.7, 263.9, _rc_clean_name(first.get('Name'))),
        'agent':     (364.4, 263.9, str(first.get('Company')) if pd.notna(first.get('Company')) else ''),
        'nrooms':    (50.9,  330.6, str(n_rooms)),
        'npax':      (187.2, 326.2, str(n_pax)),
    }


    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(595, 841))
    c.setFillColor(white)
    for x0, top, x1, bot in _GROUP_BLANK:
        c.rect(x0, H - bot, (x1 - x0), (bot - top), fill=1, stroke=0)
    c.setFillColor(black)
    c.setFont(FONT, SIZE)
    for key, (x, bottom, val) in data.items():
        if not val:
            continue
        # thu nhỏ nếu tên đoàn / hãng quá dài
        maxw = {'gname': 185, 'agent': 175}.get(key)
        fs = SIZE
        if maxw:
            while fs > 7 and c.stringWidth(val, FONT, fs) > maxw:
                fs -= 0.3
        c.setFont(FONT, fs)
        c.drawString(x, H - bottom, val)
        c.setFont(FONT, SIZE)
    c.save(); buf.seek(0)

    base = PdfReader(io.BytesIO(tmpl_bytes))
    overlay = PdfReader(buf)
    page = base.pages[0]
    page.merge_page(overlay.pages[0])
    return page



def build_regcards(xlsx_bytes, only_main=True):
    """Tạo PDF regcard hàng loạt, gộp theo Conf# (đoàn nhiều phòng → 1 regcard,
    các số phòng gộp chung vào ô Room No)."""
    df = pd.read_excel(io.BytesIO(xlsx_bytes))
    # Chuẩn hóa ngày (sửa lỗi Excel đảo dd/mm↔mm/dd với ngày ≤12 từ Smile export)
    for _dc in ('Arrival', 'Departure'):
        if _dc in df.columns:
            df[_dc] = df[_dc].map(_fix_date)
    H = 841.0
    FONT = "Times-Roman"; SIZE = 9.8
    # Baseline chính xác (bottom) đo từ dữ liệu mẫu gốc — chữ trùng khít 100%
    POS = {
        'name':(125.50,109.60),'conf':(526.75,108.52),'arrival':(119.90,144.22),
        'departure':(329.90,144.22),'nights':(500.30,144.22),'type':(113.60,179.92),
        'rm':(360.60,179.92),'company':(221.40,216.32),
        'special':(137.40,288.40),
    }
    # Ô che dữ liệu cũ — vừa khít vùng chữ, không lấn đường kẻ bảng
    BLANK = [
        (125,99,250,110.5),(526,98,568,109.5),(119,133.5,168,145.2),
        (329,133.5,378,145.2),(500,133.5,512,145.2),(113,169,145,180.9),
        (360,169,410,180.9),(221,205.5,320,217.3),
        (135,277,575,289.4),   # che dòng "AI Lunch ( EUR )..." in sẵn trên template
    ]

    # ── Gộp theo Conf# ──
    # Điền mã Conf# xuống các dòng trống (khách đi cùng booking),
    # rồi gộp TẤT CẢ dòng cùng 1 mã Conf# thành 1 regcard, gộp mọi số phòng.
    def _clean_room(r):
        s = str(r).strip()
        if s.endswith('.0'): s = s[:-2]
        return s

    df = df.copy()
    df['_conf_ff'] = df['Conf#'].ffill()

    # ── Tách các booking ĐOÀN (có mã Group) để dùng mẫu Registration Card for Group ──
    # KHÔNG ffill cột Group — chỉ dòng có sẵn mã Group mới thuộc đoàn.
    # Các dòng cùng Conf# trong đoàn: điền Group xuống theo từng Conf# nếu dòng đầu có Group.
    if 'Group' in df.columns:
        # Điền mã Group xuống các dòng cùng Conf# (đoàn nhiều phòng, chỉ dòng đầu có Group)
        df['_group_ff'] = df.groupby('_conf_ff')['Group'].transform(
            lambda s: s.ffill().bfill() if s.notna().any() else s)
    else:
        df['_group_ff'] = pd.NA
    has_group = df['_group_ff'].notna()
    df_group = df[has_group].copy()

    writer = PdfWriter()
    count = 0

    tmpl_bytes = load_regcard_template()
    grp_tmpl = None  # nạp lười khi gặp đoàn đầu tiên
    _rendered_groups = set()  # tránh vẽ trùng 1 đoàn khi trải nhiều Conf#

    # Duyệt theo ĐÚNG THỨ TỰ xuất hiện trong file (gộp theo Conf#).
    # Nhóm nào có mã Group → vẽ 1 trang Registration Card for Group (gộp cả đoàn);
    # nhóm thường → regcard thường như cũ.
    for conf_val, grp in df.groupby('_conf_ff', sort=False):
        main_rows = grp[grp['Conf#'].notna()]
        if len(main_rows) == 0:
            continue
        main = main_rows.iloc[0]

        # ── Nếu booking này thuộc ĐOÀN → dùng mẫu group ──
        gid = grp['_group_ff'].dropna().iloc[0] if grp['_group_ff'].notna().any() else None
        if gid is not None:
            if gid in _rendered_groups:
                continue  # đoàn đã vẽ ở lần gặp trước
            _rendered_groups.add(gid)
            if grp_tmpl is None:
                grp_tmpl = load_group_template()
            gdf = df_group[df_group['_group_ff'] == gid]
            page = build_group_regcard(gdf, grp_tmpl)
            writer.add_page(page)
            count += 1
            continue

        # ── Booking thường → regcard thường ──
        rooms = []
        for r in grp['Rm']:
            if pd.notna(r):
                rs = _clean_room(r)
                # Bỏ phòng ảo đầu 9 dạng 9000-9999 (9002, 9005, 9010, 9040... — posting master)
                if rs and rs not in rooms and not _re.fullmatch(r'9\d{3}', rs):
                    rooms.append(rs)
        # Gom mã Specials của cả nhóm (để bắt CN/EB dù nằm ở dòng nào)
        spec_codes = set()
        if 'Specials' in grp.columns:
            for sv in grp['Specials'].dropna():
                for code in str(sv).split(','):
                    code = code.strip().upper()
                    if code:
                        spec_codes.add(code)
        g = {'main': main, 'rooms': rooms, 'specials': spec_codes}
        row = g['main']
        name = _rc_clean_name(row.get('Name'))
        if not name:
            continue
        company = str(row.get('Company')) if pd.notna(row.get('Company')) else ''
        # ── Ô SPECIAL REQUEST ──
        # Mặc định: AI Lunch, AI Dinner, Minibar set up.
        # CELERIS → thêm ( EUR ) sau Lunch & Dinner.
        # Specials có CN → thêm Connecting Room; có EB → thêm Extra Bed.
        if _norm_nat(company).startswith('celeris'):
            sr = 'AI Lunch ( EUR ), AI Dinner ( EUR ), Minibar set up'
        else:
            sr = 'AI Lunch, AI Dinner, Minibar set up'
        _spec = g.get('specials', set())
        if 'CN' in _spec:
            sr += ', Connecting Room'
        if 'EB' in _spec:
            sr += ', Extra Bed'

        data = {
            'name': name,
            'conf': _rc_conf(row.get('Conf#')),
            'arrival': _rc_date(row.get('Arrival')),
            'departure': _rc_date(row.get('Departure')),
            'nights': _rc_nights(row.get('Arrival'), row.get('Departure')),
            'type': str(row.get('Type')) if pd.notna(row.get('Type')) else '',
            'rm': ', '.join(g['rooms']),   # gộp các số phòng
            'company': company,
            'special': sr,
        }
        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(595,841))
        c.setFillColor(white)
        for x0,top,x1,bot in BLANK:
            c.rect(x0, H-bot, (x1-x0), (bot-top), fill=1, stroke=0)
        c.setFillColor(black)
        c.setFont(FONT, SIZE)
        MAXW = {'name': 300, 'company': 200}  # ô 1 dòng: thu nhỏ nếu tràn
        for key,(x,bottom) in POS.items():
            val = data[key]
            if not val:
                continue
            if key == 'rm':
                # Ô số phòng: 1 phòng → vị trí chuẩn như mẫu gốc.
                # Nhiều phòng → căn giữa CHÍNH XÁC giữa 2 nhãn (đo từ 2 ảnh crop độc lập,
                # sai lệch <1pt): nhãn "ROOM NO./(Số phòng)" kết thúc ~261pt,
                # nhãn "ROOM RATE/(Giá phòng)" bắt đầu ~441pt.
                # → Vùng vẽ 266–436pt, TÂM = 351pt, rộng 170pt
                #   (đủ ~8 phòng/dòng ở cỡ chữ nguyên vẹn 9.8pt → 24 phòng/3 dòng).
                rooms = [s.strip() for s in val.split(',') if s.strip()]
                RM_X0, RM_X1 = 266.0, 436.0
                RM_CX = (RM_X0 + RM_X1) / 2        # = 351.0
                RM_MAXW = RM_X1 - RM_X0            # = 170pt mỗi dòng
                def _wrap(items, fs):
                    lines=[]; cur=''
                    for it in items:
                        test = (cur + ', ' + it) if cur else it
                        if c.stringWidth(test, FONT, fs) <= RM_MAXW:
                            cur = test
                        else:
                            if cur: lines.append(cur)
                            cur = it
                    if cur: lines.append(cur)
                    return lines
                if len(rooms) <= 1:
                    c.drawString(x, H - bottom, val)
                else:
                    fs = SIZE
                    lines = _wrap(rooms, fs)
                    if len(lines) <= 3:
                        gap = fs + 1.6             # tới ~24 phòng: cỡ chữ 9.8pt nguyên vẹn
                    else:
                        gap = fs + 0.6             # 4 dòng sát nhau, vẫn nguyên cỡ chữ (~32 phòng)
                        while len(lines) > 4 and fs > 7:
                            fs -= 0.3              # chống tràn tuyệt đối cho case phi thực tế
                            lines = _wrap(rooms, fs)
                    c.setFont(FONT, fs)
                    for i, ln in enumerate(lines):
                        c.drawCentredString(RM_CX, H - bottom - i*gap, ln)
                    c.setFont(FONT, SIZE)
            else:
                maxw = MAXW.get(key)
                if key == 'special':
                    # Ô special: vẽ từ x=137.4, giới hạn mép phải bảng ~573 → rộng ~436pt.
                    # Thu nhỏ nhẹ nếu quá dài (hiếm), giữ nguyên baseline dòng in sẵn.
                    SP_MAXW = 573 - x
                    fs = SIZE
                    while fs > 7 and c.stringWidth(val, FONT, fs) > SP_MAXW:
                        fs -= 0.2
                    c.setFont(FONT, fs)
                    c.drawString(x, H - bottom, val)
                    c.setFont(FONT, SIZE)
                elif maxw:
                    fs = SIZE
                    while fs > 6 and c.stringWidth(val, FONT, fs) > maxw:
                        fs -= 0.3
                    c.setFont(FONT, fs)
                    c.drawString(x, H-bottom, val)
                    c.setFont(FONT, SIZE)
                else:
                    c.drawString(x, H-bottom, val)
        c.save(); buf.seek(0)
        base = PdfReader(io.BytesIO(tmpl_bytes))
        overlay = PdfReader(buf)
        page = base.pages[0]
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
        count += 1
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue(), count

# ── Đối chiếu Smile vs Trang lưu trú người nước ngoài ─────────────────────

# Ten duoc module nay so huu — liet ke tuong minh de `import *` lay
# duoc ca helper co gach duoi dau.
# ── Tạo file ARR từ file Arrival (Book) Smile ──────────────────────────────
def build_arr(book_bytes):
    """Tạo file ARR ĐÚNG định dạng của ARR Converter gốc (tool HTML riêng, không
    phải file mẫu in cũ):
    - 6 cột: Conf# / Arrival / Departure / Company / Notice / [số phòng] — đọc
      cột nguồn theo TÊN (Conf#, Folio#, Type, Arrival, Departure, Company,
      Notice), không theo vị trí cố định như bản cũ.
    - Font Patrick Hand toàn bộ; Conf# cỡ 50 đậm nền cam nhạt; số phòng cỡ 50;
      các ô còn lại cỡ 20. Dòng dữ liệu cao 120, dòng header cao 142.5.
    - Số phòng = số dòng Folio# hợp lệ trùng Conf# (bỏ dòng Type='**' - dummy).
    - Dòng phụ chèn ngay sau booking tương ứng, gộp A:F, nền màu theo loại:
      CÀ THẺ (cam) · THU TIỀN (xanh lá) · XEM LẠI BU (vàng) · FOC LATE C/O (xanh
      dương, tự đọc giờ trong Notice nếu có, vd "FOC LATE C/O 18:00").
    - Nhận diện nghiệp vụ đầy đủ (OTA, từ khóa CÀ THẺ/THU TIỀN/FOC/XEM LẠI BU)
      y hệt bộ từ khóa của ARR Converter gốc.
    """
    import re as _re2
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    df = pd.read_excel(io.BytesIO(book_bytes), header=None)
    hdr = None
    for i in range(min(5, len(df))):
        if any(str(v).strip() == 'Conf#' for v in df.iloc[i] if pd.notna(v)):
            hdr = i; break
    if hdr is None:
        raise ValueError("Không tìm thấy dòng header chứa \"Conf#\" trong file. Kiểm tra lại file Arrival Smile.")

    headers = [str(v).strip() if pd.notna(v) else '' for v in df.iloc[hdr]]
    col = {}
    for i, h in enumerate(headers):
        if h and h not in col:
            col[h] = i
    data = df.iloc[hdr + 1:].reset_index(drop=True)

    def C(name):
        return col.get(name, -1)

    conf_c, folio_c, type_c = C('Conf#'), C('Folio#'), C('Type')
    arr_c, dep_c, comp_c, notice_c = C('Arrival'), C('Departure'), C('Company'), C('Notice')
    if -1 in (conf_c, folio_c, arr_c, comp_c):
        raise ValueError("File thiếu cột bắt buộc (Conf#, Folio#, Arrival, Company). Kiểm tra lại file.")

    # Sửa lỗi Excel đảo dd/mm↔mm/dd với ngày ≤12 từ Smile export (dùng chung hàm _fix_date)
    def _arr_fmt_date(v):
        fixed = _fix_date(v)
        if fixed is not None:
            return fixed.strftime('%d/%m/%y')
        return str(v).strip() if isinstance(v, str) and v.strip() else ''

    # ── Đếm số phòng (= số dòng Folio# hợp lệ) theo từng Conf#, bỏ dòng Type='**' (dummy) ──
    room_counts = {}
    dummy_count = 0
    for _, row in data.iterrows():
        conf = row.iloc[conf_c] if conf_c >= 0 else None
        folio = row.iloc[folio_c] if folio_c >= 0 else None
        if pd.notna(conf) and pd.notna(folio):
            typ = row.iloc[type_c] if type_c >= 0 else None
            if str(typ).strip() == '**':
                dummy_count += 1
                continue
            room_counts[conf] = room_counts.get(conf, 0) + 1

    # ── Danh sách booking theo ĐÚNG thứ tự xuất hiện, mỗi Conf# 1 dòng ──
    seen = set()
    ordered = []
    for _, row in data.iterrows():
        conf = row.iloc[conf_c] if conf_c >= 0 else None
        folio = row.iloc[folio_c] if folio_c >= 0 else None
        typ = row.iloc[type_c] if type_c >= 0 else None
        arr = row.iloc[arr_c] if arr_c >= 0 else None
        comp = row.iloc[comp_c] if comp_c >= 0 else None
        dep = row.iloc[dep_c] if dep_c >= 0 else None
        notice = row.iloc[notice_c] if notice_c >= 0 else None
        if pd.isna(conf) or pd.isna(folio) or pd.isna(arr) or pd.isna(comp):
            continue
        if str(typ).strip() == '**':
            continue
        if conf in seen:
            continue
        if not room_counts.get(conf):
            continue
        seen.add(conf)
        ordered.append({
            'type': 'bk', 'conf': conf,
            'arrival': _arr_fmt_date(arr),
            'departure': _arr_fmt_date(dep) if pd.notna(dep) else '',
            'company': str(comp).strip(),
            'notice': str(notice).strip() if pd.notna(notice) else '',
            'rooms': room_counts[conf],
        })
    if not ordered:
        raise ValueError("File không có dữ liệu booking hợp lệ nào.")

    # ── Bộ nhận diện nghiệp vụ — y hệt ARR Converter gốc ──
    ARR_OTA = ['EXPEDIA','BOOKING','AGODA','TRIP.COM','CTRIP','AIRBNB','TRAVELOKA',
        'KLOOK','KAYAK','PRICELINE','HOTELS.COM','ORBITZ','TRIVAGO','MAKEMYTRIP',
        'LASTMINUTE','HOSTELWORLD','WOTIF','HOTWIRE','VRBO','HOMEAWAY','IVIVU',
        'MYTOUR','LUXSTAY','VNTRIP','GOTADI','TRAVELPORT','SKYSCANNER','BESTPRICE',
        'LATEROOMS','EASYJET','RYANAIR','JETSTAR','HOTELBEDS','TOURICO','GETAROOM']
    ARR_CA_THE = ['TACC','CC UPON','CHARGE CC','CHARGE CARD','CREDIT CARD','DEBIT CARD',
        'CC AUTH','AUTH CC','AUTHORIZE CC','AUTHORIZE CARD','CC ON ARRIVAL',
        'BILL TO CC','SWIPE CC','SWIPE CARD','PRE-AUTH','PREAUTH','PREPAID CC',
        'CHARGE ON CC','CARD ON ARRIVAL','CC AT CI','CC AT CHECK','PAY BY CARD',
        'CARD PAYMENT','TC UPON','TC ON ARRIVAL','TAKE CC','TAKE CARD']
    ARR_THU_TIEN = ['PAY UPON','PAY ON ARRIVAL','CASH ON ARRIVAL','CASH UPON','COLLECT CASH',
        'COLLECT PAYMENT','COLLECT ON ARRIVAL','PAYMENT ON ARRIVAL','CASH PAYMENT',
        'CASH AT CHECK','CASH AT CI','DUE ON ARRIVAL','PAYABLE ON ARRIVAL',
        'PAY AT CI','PAY AT CHECK','CASH DUE','OUTSTANDING','BALANCE DUE',
        'PAYMENT DUE','COLLECT AT CI','COLLECT AT CHECK',
        'RC UPON','UPON C/I','UPON CI','UPON CHECK-IN','UPON CHECKIN',
        'GOA UPON','ROH UPON','COLLECT UPON']
    ARR_XEM_LAI = ['CASH UPON','CASH ON ARRIVAL','CASH AT CI','CASH PAYMENT',
        'PAY UPON','PAY AT CI','COLLECT CASH','CASH DUE']
    ARR_FOC_LCO = ['FOC LATE CHECK','FOC LATE CHECKOUT','FOC LATE C/O','FOC LCO',
        'LCO FOC','LATE CHECK OUT FOC','LATE CHECKOUT FOC','LATE C/O FOC',
        'COMP LATE CHECK','COMP LATE CHECKOUT','COMP LCO',
        'COMPLIMENTARY LATE CHECK','COMPLIMENTARY LCO',
        'GRATIS LATE CHECK','FREE LATE CHECK']

    def _pay_type(notice, company):
        n = str(notice or '').upper()
        co = str(company or '').upper()
        is_ota = any(o in co for o in ARR_OTA)
        is_ca_the = any(k in n for k in ARR_CA_THE)
        is_thu_tien = any(k in n for k in ARR_THU_TIEN)
        is_foc_lco = any(k in n for k in ARR_FOC_LCO)
        has_foc = 'FOC' in n or 'COMP' in n or 'COMPLIMENTARY' in n
        has_lco = 'LATE CHECK' in n or ' LCO' in n or 'LATE C/O' in n or 'LATE CHECKOUT' in n
        if is_foc_lco or (has_foc and has_lco):
            return 'foc_lco'
        if is_ota and any(k in n for k in ARR_XEM_LAI):
            return 'xem_lai_bu'
        if is_ca_the:
            return 'ca_the'
        if is_thu_tien:
            return 'thu_tien'
        return 'none'

    result = []
    for i, bk in enumerate(ordered):
        result.append(bk)
        if i >= len(ordered) - 1:
            continue
        pt = _pay_type(bk['notice'], bk['company'])
        if pt == 'ca_the':
            result.append({'type': 'sep', 'conf': 'CÀ THẺ'})
        elif pt == 'thu_tien':
            result.append({'type': 'sep', 'conf': 'THU TIỀN'})
        elif pt == 'xem_lai_bu':
            result.append({'type': 'sep', 'conf': 'XEM LẠI BU'})
        elif pt == 'foc_lco':
            m = _re2.search(r'\b(\d{1,2}[:Hh]\d{2})\b', bk['notice'])
            time_str = (' ' + m.group(1).upper()) if m else ''
            result.append({'type': 'sep', 'conf': 'FOC LATE C/O' + time_str})

    # ── Xuất Excel đúng định dạng ARR Converter gốc ──
    wb = Workbook(); ws = wb.active; ws.title = 'Sheet1'
    for i, w in enumerate([39.4, 15.1, 16.0, 21.9, 50.0, 10.3], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    thin = Side(style='thin')
    border_all = Border(top=thin, bottom=thin, left=thin, right=thin)
    center_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)
    conf_fill = PatternFill('solid', fgColor='FDEADA')
    fill_colors = {'CÀ THẺ': 'FDEADA', 'THU TIỀN': 'D4F4E8', 'XEM LẠI BU': 'FFF8DC'}

    ws.row_dimensions[1].height = 142.5
    for i, h in enumerate(['Conf#', 'Arrival', 'Departure', 'Company', 'Notice', None], 1):
        cell = ws.cell(1, i)
        cell.value = h
        cell.font = Font(name='Patrick Hand', size=50 if h == 'Conf#' else 20)
        cell.alignment = center_wrap
        cell.border = border_all

    r = 2
    for item in result:
        ws.row_dimensions[r].height = 120.0
        if item['type'] == 'sep':
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
            label = item.get('conf') or 'CÀ THẺ'
            color = fill_colors.get(label) or ('D6EAF8' if label.startswith('FOC LATE C/O') else 'FDEADA')
            cell = ws.cell(r, 1)
            cell.value = label
            cell.font = Font(name='Patrick Hand', size=50)
            cell.alignment = center_wrap
            cell.fill = PatternFill('solid', fgColor=color)
            cell.border = border_all
        else:
            vals = [item['conf'], item['arrival'], item['departure'], item['company'], item['notice'], item['rooms']]
            for ci, v in enumerate(vals, 1):
                cell = ws.cell(r, ci)
                cell.value = v
                cell.font = Font(name='Patrick Hand', size=50 if ci in (1, 6) else 20, bold=(ci == 1))
                cell.alignment = center_wrap
                cell.border = border_all
                if ci == 1:
                    cell.fill = conf_fill
        r += 1

    bookings = [x for x in result if x['type'] == 'bk']
    stats = {
        'bookings': len(bookings),
        'rooms': sum(b['rooms'] for b in bookings),
        'ota': sum(1 for b in bookings if any(o in b['company'].upper() for o in ARR_OTA)),
        'dummy': dummy_count,
        'ca_the': sum(1 for x in result if x['type'] == 'sep' and x['conf'] == 'CÀ THẺ'),
        'thu_tien': sum(1 for x in result if x['type'] == 'sep' and x['conf'] == 'THU TIỀN'),
        'xem_lai_bu': sum(1 for x in result if x['type'] == 'sep' and x['conf'] == 'XEM LẠI BU'),
        'foc_lco': sum(1 for x in result if x['type'] == 'sep' and x['conf'].startswith('FOC LATE C/O')),
    }
    return wb, stats


__all__ = [
    '_GROUP_BLANK', '_GROUP_TABLE_LINES', 'build_arr', '_grp_date', '_rc_clean_name', '_rc_conf', '_rc_date', '_rc_nights',
    'build_group_regcard', 'build_regcards', 'load_group_template',
    'load_regcard_template'
]
