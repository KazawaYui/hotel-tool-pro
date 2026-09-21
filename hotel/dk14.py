"""Sinh sổ ĐK14 bằng cách chèn thẳng XML vào mẫu, giữ nguyên 100% định dạng."""

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

from .assets import load_template
from .lookups import NAT_DK14

def _dk_map_gender(g):
    """Chuẩn hóa giới tính về 'Nam'/'Nữ' — chấp nhận cả chữ cái đơn (M/F, kể
    cả không hoa) lẫn tiếng Việt đầy đủ (Nam/Nữ), khác bản cũ chỉ nhận đúng
    ký tự 'M'/'F' (sai hoàn toàn nếu nguồn ghi 'Nam'/'Nữ' như mọi cột GIỚI
    TÍNH khác trong app — khi đó ngày sinh sẽ không vào được cột nào)."""
    s = str(g or '').strip().lower()
    if s in ('nam', 'm', 'male'):
        return 'Nam'
    if s in ('nữ', 'nu', 'f', 'female'):
        return 'Nữ'
    return str(g or '').strip()

def _dk_is_dummy(name, room):
    """Nhận diện dòng dummy/test — không phải khách lưu trú thật, không đưa
    vào sổ ĐK14 chính thức nộp công an."""
    n = str(name or '').strip().lower()
    if not n:
        return True, 'Tên trống'
    if 'dummy' in n:
        return True, 'Tên là dummy'
    if 'pending' in n:
        return True, 'Tên là pending'
    if 'water sport' in n:
        return True, 'Không phải khách lưu trú'
    if _re.fullmatch(r'[A-Za-z\s]+\d{4,}', str(name or '').strip()):
        return True, 'Tên chứa mã đặt phòng'
    digits = _re.sub(r'\D', '', str(room or ''))
    if digits and int(digits) >= 9000:
        return True, f'Phòng {room} ≥ 9000 (phòng ảo/posting master)'
    return False, ''

# ── ĐK14: chuẩn hoá giá trị đọc từ file nguồn ─────────────────────────────
# xlrd/openpyxl trả ô ngày về SỐ SERIAL Excel (không phải datetime) khi đọc
# thô — cố tình đọc thô vì mọi thư viện đọc sẵn kiểu ngày đều có nguy cơ lệch
# múi giờ; công thức (serial - 25569) * 86400 giây quy đổi thẳng nên không lệch.
def _dk_cell_str(v):
    """Ô số nguyên (số giấy tờ, số phòng...) bị đọc thành float — bỏ đuôi '.0'
    để không ghi '123456789.0' vào sổ nộp công an."""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def _dk_to_date(v):
    """Mọi kiểu ô ngày (serial Excel, datetime, chuỗi dd/mm/yyyy) → date."""
    if v is None or v == '':
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, (int, float)):
        try:
            return (datetime.datetime(1970, 1, 1)
                    + datetime.timedelta(seconds=round((float(v) - 25569) * 86400))).date()
        except Exception:
            return None
    m = _re.match(r'^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$', str(v).strip())
    if m:
        try:
            return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None

def _dk_fmt_date(v):
    d = _dk_to_date(v)
    return d.strftime('%d/%m/%Y') if d else ''

def _dk_is_invalid_id(v):
    """Mã giấy tờ CHƯA CẤP (không phải số thật) — xoá khỏi sổ. GKS/GBL được xử
    lý riêng ở nơi gọi vì đó là dữ liệu hợp lệ (giấy khai sinh / giấy bảo lãnh)."""
    s = str(v or '').strip().upper()
    return len(s) <= 4 and bool(_re.fullmatch(r'[A-Z]{2,4}', s)) and s in ('GKS', 'GBL', 'GKA', 'GBS')

def _dk_detect_columns(rows):
    """Tự dò dòng tiêu đề (trong 5 dòng đầu) và vị trí từng cột theo TÊN tiêu
    đề, thay vì bám chỉ số cột cứng — file IH đổi thứ tự/thêm cột vẫn chạy đúng."""
    header_row, headers = -1, None
    for i in range(min(5, len(rows))):
        r = rows[i] or []
        has_name = any(c is not None and _re.match(r'^(họ và tên|name|full name)', str(c).strip(), _re.I) for c in r)
        has_dob = any(c is not None and _re.search(r'(ngày sinh|date of birth|ngày tháng năm sinh|dob)', str(c), _re.I) for c in r)
        if has_name or has_dob:
            header_row, headers = i, r
            break
    if header_row == -1:
        header_row, headers = 0, (rows[0] if rows else [])
    cols = {'headerRow': header_row, 'dataStart': header_row + 1}
    for idx, cell in enumerate(headers or []):
        if cell is None:
            continue
        s = str(cell).lower().strip()
        if not s:
            continue
        if s in ('stt', 'no', 'no.', 'số tt') or s.startswith('stt '):
            cols['stt'] = idx
        elif _re.search(r'họ và tên (khách|kh)', s) or s in ('họ và tên', 'name', 'full name', 'guest name'):
            cols['name'] = idx
        elif _re.search(r'(ngày sinh|date of birth|ngày tháng năm sinh|dob)', s) and 'thông báo' not in s:
            cols['dob'] = idx
        elif _re.fullmatch(r'(giới tính|gender|sex)', s) or s.startswith('giới tính') or s.startswith('gender'):
            cols['gender'] = idx
        elif 'quốc tịch' in s:
            cols['nationality'] = idx
        elif 'quốc gia' in s and 'nationality' not in cols:
            cols['country'] = idx
        elif _re.search(r'country.*residence|residence.*country|nationality', s):
            cols['nationality'] = idx
        elif 'country' in s and 'nationality' not in cols:
            cols['country'] = idx
        elif _re.match(r'^(số giấy tờ|passport|cccd|cmnd|id card)', s) or _re.search(r'(passport|id card)', s):
            cols['id'] = idx
        elif s.startswith('loại giấy tờ'):
            cols['docType'] = idx
        elif s.startswith('tên giấy tờ'):
            cols['docName'] = idx
        elif _re.search(r'(số điện thoại|phone|tel|mobile)', s):
            cols['phone'] = idx
        elif 'loại cư trú' in s:
            cols['cuTru'] = idx
        elif _re.match(r'^(tỉnh|tp|province|city)', s) or s.startswith('tỉnh/'):
            cols['tinh'] = idx
        elif _re.match(r'^(quận|huyện|district)', s) or s.startswith('quận/'):
            cols['quan'] = idx
        elif _re.match(r'^(phường|xã|ward|commune)', s) or s.startswith('phường/'):
            cols['phuong'] = idx
        elif _re.search(r'(địa chỉ chi tiết|address line|^address$)', s):
            cols['addressDetail'] = idx
        elif s.startswith('địa chỉ') and 'addressDetail' not in cols:
            cols['addressDetail'] = idx
        elif _re.search(r'(arrival|ngày đến|thời gian.*đến|check.in)', s) or s.startswith('đến'):
            cols['arrival'] = idx
        elif _re.search(r'(depature|departure|ngày đi|thời gian.*đi|check.out)', s) or s.startswith('đi'):
            cols['departure'] = idx
        elif _re.search(r'(số buồng|số phòng|^phòng|room\s*#|^room$|room number)', s) or 'phòng/khoa' in s:
            cols['room'] = idx
        elif _re.search(r'(người thông báo|notifier|reporter)', s):
            cols['notifier'] = idx
    return cols, (headers or [])

def _dk_addr_from_parts(cu_tru, tinh, quan, phuong, detail, iso):
    """File IH không có cột địa chỉ gộp thì ghép từ các cột rời. CHỈ ghép cho
    khách Việt Nam — khách nước ngoài để trống theo đúng mẫu ĐK14."""
    if iso != 'VNM':
        return '   '
    cu_tru = str(cu_tru or '').strip()
    segs = [str(p or '').strip() for p in (detail, phuong, quan, tinh)]
    segs = [p for p in segs if p and p not in ('null', 'undefined', 'nan')]
    addr = ', '.join(segs)
    if not addr and not cu_tru:
        return '   '
    if cu_tru:
        return f'{cu_tru}: {addr}' if addr else cu_tru
    return addr

def _dk_read_rows(file_bytes):
    """Đọc file nguồn IH ở dạng THÔ (ô ngày giữ nguyên số serial). Nhận cả .xls
    (xlrd) lẫn .xlsx (openpyxl) — nhận diện bằng chữ ký file, không theo đuôi."""
    if file_bytes[:2] == b'PK':
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        return rows
    wb = xlrd.open_workbook(file_contents=file_bytes)
    ws = wb.sheet_by_index(0)
    return [[ws.cell_value(r, c) for c in range(ws.ncols)] for r in range(ws.nrows)]

def _dk_transform(rows, default_notifier, default_checkin):
    """IH → các dòng dữ liệu ĐK14, kèm nhật ký từng dòng đã sửa/bỏ qua."""
    cols, headers = _dk_detect_columns(rows)
    issues, skipped, data = [], [], []
    missing = [k for k in ('name', 'dob') if k not in cols]
    if missing:
        issues.append(('warn', cols['headerRow'] + 1, '(tiêu đề)',
                       'Không tìm thấy cột: ' + ', '.join(missing)))

    def get(row, key):
        i = cols.get(key)
        return row[i] if i is not None and i < len(row) else None

    seen_ids, stt = set(), 0
    for i in range(cols['dataStart'], len(rows)):
        row = rows[i] or []
        if all(c is None or str(c).strip() == '' for c in row):
            continue
        raw_name = get(row, 'name')
        name_str = str(raw_name or '')
        room = _dk_cell_str(get(row, 'room'))
        is_dummy, reason = _dk_is_dummy(name_str, room)
        if is_dummy:
            skipped.append((name_str.strip() or '(trống)', reason))
            issues.append(('skip', i + 1, name_str.strip() or '(trống)', f'Bỏ qua — {reason}'))
            continue
        stt += 1
        note = []
        name = name_str.strip()
        if name_str != name:
            note.append('Trim khoảng trắng tên')
        gender = _dk_map_gender(get(row, 'gender'))
        dob = _dk_fmt_date(get(row, 'dob'))
        if not dob:
            note.append('Thiếu ngày sinh')
        if gender not in ('Nam', 'Nữ'):
            note.append('Thiếu/không xác định giới tính')
        iso = _dk_cell_str(get(row, 'nationality') or get(row, 'country')).upper()
        nat = NAT_DK14.get(iso, iso)
        if iso and iso not in NAT_DK14:
            note.append(f'Mã quốc tịch "{iso}" không có trong danh mục')

        id_num = _dk_cell_str(get(row, 'id'))
        id_up = id_num.upper()
        if id_up in ('GKS', 'GBL'):
            # Giấy khai sinh (GKS) / giấy bảo lãnh (GBL) — dữ liệu HỢP LỆ (khách
            # dùng thay CCCD/hộ chiếu). Giữ nguyên, không xoá, không tính trùng
            # vì nhiều khách có thể cùng dùng chung 1 mã.
            pass
        elif _dk_is_invalid_id(id_num):
            note.append(f'Số giấy tờ "{id_num}" chưa cấp → xóa')
            id_num = ''
        elif not id_num:
            note.append('Thiếu số giấy tờ')
        elif id_num in seen_ids:
            note.append(f'Số giấy tờ "{id_num}" trùng dòng trước')
        else:
            seen_ids.add(id_num)

        if 'addressDetail' in cols:
            addr = _dk_cell_str(get(row, 'addressDetail')).strip() or '   '
        else:
            addr = _dk_addr_from_parts(get(row, 'cuTru'), get(row, 'tinh'), get(row, 'quan'),
                                       get(row, 'phuong'), get(row, 'addressDetail'), iso)

        arr = _dk_to_date(get(row, 'arrival'))
        if arr is None:
            arr = default_checkin
            if 'arrival' in cols:
                note.append('Thiếu ngày đến → dùng ngày mặc định')
        dep = _dk_to_date(get(row, 'departure'))
        if dep is None:
            note.append('Thiếu ngày đi')

        notifier = _dk_cell_str(get(row, 'notifier')).strip() or default_notifier

        if note:
            kind = 'fix' if any('→' in x or 'Trim' in x for x in note) else 'warn'
            issues.append((kind, i + 1, name, ' · '.join(note)))
        data.append([stt, name,
                     dob if gender == 'Nam' else '', dob if gender == 'Nữ' else '',
                     nat, id_num, addr, arr, dep, room, notifier, '', ''])
    return data, issues, skipped, cols, headers

# ── ĐK14: ghi file bằng cách CHÈN THẲNG XML vào mẫu ──────────────────────
# Chỉ thay phần <sheetData> từ dòng 18 trở đi, giữ nguyên byte mọi thứ còn lại
# của file mẫu (kiểu ô, viền, font, gộp ô, khổ giấy, hình vẽ, cấu hình in...).
# Cách này giữ được nhiều hơn so với dựng lại workbook bằng openpyxl — openpyxl
# không giữ hình vẽ và cấu hình máy in.
_DK_COLS = 'ABCDEFGHIJKLM'
_DK_DEFAULT_STYLES = {'A': '8', 'B': '8', 'C': '19', 'D': '19', 'E': '8', 'F': '13',
                      'G': '8', 'H': '8', 'I': '8', 'J': '8', 'K': '8', 'L': '8', 'M': '8'}
_DK_DEFAULT_ROW_ATTRS = 's="4" customFormat="1" ht="14.25" x14ac:dyDescent="0.15"'

def _dk_esc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;'))

def _dk_template_styles(xml):
    """Lấy kiểu ô/thuộc tính dòng từ chính các dòng dữ liệu mẫu (18–25) để dòng
    sinh ra trông y hệt mẫu; thiếu thì dùng bộ mặc định của mẫu TT30/2026."""
    col_styles, row_attrs = {}, ''
    for test_row in range(18, 26):
        m = _re.search(r'<row([^>]*\br="%d"[^>]*)>(.*?)</row>' % test_row, xml, _re.S)
        if not m:
            continue
        if not row_attrs:
            attrs = [f'{k}="{v}"' for k, v in _re.findall(r'(\w+(?::\w+)?)="([^"]*)"', m.group(1))
                     if k not in ('r', 'spans')]
            row_attrs = ' '.join(attrs)
        for col, cattr in _re.findall(r'<c\s+r="([A-Z]+)%d"([^>/]*)' % test_row, m.group(2)):
            sm = _re.search(r'\ss="(\d+)"', cattr)
            if sm and col not in col_styles:
                col_styles[col] = sm.group(1)
    for col, style in _DK_DEFAULT_STYLES.items():
        col_styles.setdefault(col, style)
    return col_styles, (row_attrs or _DK_DEFAULT_ROW_ATTRS)

def _dk_row_xml(row_num, values, col_styles, row_attrs):
    out = [f'<row r="{row_num}" spans="1:13" {row_attrs}>']
    for ci, col in enumerate(_DK_COLS):
        val = values[ci] if ci < len(values) else ''
        ref = f'{col}{row_num}'
        s_attr = f' s="{col_styles[col]}"' if col_styles.get(col) else ''
        if val is None or val == '':
            out.append(f'<c r="{ref}"{s_attr}/>')
        elif isinstance(val, (datetime.date, datetime.datetime)):
            # Ghi ngày dạng CHỮ "dd/mm/yyyy" thay vì số serial: số serial chỉ hiện
            # đúng khi ô có định dạng ngày, mà mẫu không đảm bảo điều đó.
            out.append(f'<c r="{ref}"{s_attr} t="inlineStr"><is><t>'
                       f'{val.strftime("%d/%m/%Y")}</t></is></c>')
        elif isinstance(val, int) and not isinstance(val, bool):
            out.append(f'<c r="{ref}"{s_attr}><v>{val}</v></c>')
        else:
            sv = str(val)
            t_attr = ' xml:space="preserve"' if sv != sv.strip() else ''
            out.append(f'<c r="{ref}"{s_attr} t="inlineStr"><is><t{t_attr}>'
                       f'{_dk_esc(sv)}</t></is></c>')
    out.append('</row>')
    return ''.join(out)

def _dk_set_cell_text(xml, ref, text):
    """Ghi đè 1 ô của mẫu bằng chuỗi (kể cả ô đang trỏ tới sharedStrings)."""
    esc = _dk_esc(text)
    t_attr = ' xml:space="preserve"' if text != text.strip() else ''
    def _sub(m):
        attrs = _re.sub(r'\s+t="[^"]*"', '', m.group(1))
        return f'<c r="{ref}"{attrs} t="inlineStr"><is><t{t_attr}>{esc}</t></is></c>'
    return _re.sub(r'<c\s+r="%s"([^>/]*?)(?:/>|>.*?</c>)' % ref, _sub, xml, flags=_re.S)

def build_dk14(xls_bytes, default_notifier='', default_checkin=None):
    """IH → file ĐK14 .xlsx. Trả (bytes file, số khách, danh sách bỏ qua,
    nhật ký dòng, thông tin cột đã dò được)."""
    rows = _dk_read_rows(xls_bytes)
    data, issues, skipped, cols, headers = _dk_transform(rows, default_notifier, default_checkin)

    src = io.BytesIO(load_template('dk14'))
    out = io.BytesIO()
    sheet_path = 'xl/worksheets/sheet1.xml'
    with zipfile.ZipFile(src) as zin:
        if sheet_path not in zin.namelist():
            raise ValueError('Mẫu ĐK14 không có xl/worksheets/sheet1.xml')
        xml = zin.read(sheet_path).decode('utf-8')

        col_styles, row_attrs = _dk_template_styles(xml)
        deps = [r[8] for r in data if isinstance(r[8], datetime.date)]
        xml = _dk_set_cell_text(xml, 'J11',
                                'Từ ngày: ' + (default_checkin.strftime('%d/%m/%Y') if default_checkin else ''))
        xml = _dk_set_cell_text(xml, 'J12',
                                'Đến ngày: ' + (max(deps).strftime('%d/%m/%Y') if deps else ''))

        new_rows = ''.join(_dk_row_xml(18 + i, r, col_styles, row_attrs) for i, r in enumerate(data))
        end = xml.index('</sheetData>')
        start = -1
        for m in _re.finditer(r'<row\s+r="(\d+)"[^>]*>', xml):
            if int(m.group(1)) >= 18:
                start = m.start()
                break
        xml = (xml[:end] if start == -1 else xml[:start]) + new_rows + xml[end:]

        last_row = max(17 + len(data), 18)
        xml = _re.sub(r'<dimension\s+ref="[^"]*"\s*/>', f'<dimension ref="A1:M{last_row}"/>', xml)

        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                zout.writestr(item, xml.encode('utf-8') if item.filename == sheet_path
                              else zin.read(item.filename))
    return out.getvalue(), len(data), skipped, issues, (cols, headers)



# Ten duoc module nay so huu — liet ke tuong minh de `import *` lay
# duoc ca helper co gach duoi dau.
__all__ = [
    '_DK_COLS', '_DK_DEFAULT_ROW_ATTRS', '_DK_DEFAULT_STYLES', '_dk_addr_from_parts',
    '_dk_cell_str', '_dk_detect_columns', '_dk_esc', '_dk_fmt_date', '_dk_is_dummy',
    '_dk_is_invalid_id', '_dk_map_gender', '_dk_read_rows', '_dk_row_xml',
    '_dk_set_cell_text', '_dk_template_styles', '_dk_to_date', '_dk_transform',
    'build_dk14'
]
