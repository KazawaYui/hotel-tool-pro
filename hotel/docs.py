"""Sinh hồ sơ KBTT (khách nước ngoài) và VNM (khách Việt) nộp công an."""

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

from .common import apply_style, snapshot_styles, _fix_departure_swap, _fmt_room, _norm_name, _norm_pp, cp, fmt, make_code
from .assets import load_template
from .lookups import (
    LOAI_GIAY, _norm_nat, lookup_nat_kbtt, lookup_province_vnm, lookup_ward_vnm,
)

def parse_visa_file(visa_bytes):
    """Đọc file thô visa/quản lý người nước ngoài → dict {"by_pp": {...}, "by_name": {...}}.
    Hỗ trợ 2 định dạng cột:
    1) File "Trang quản lý người nước ngoài" thật: 'HỌ TÊN' (tên đầy đủ 1 cột),
       'SỐ HỘ CHIẾU', và 'THỜI HẠN ĐƯỢC PHÉP TẠM TRÚ TẠI VIỆT NAM'.
    2) File cũ dạng Last Name / First Name / Visa date.
    Khớp theo SỐ HỘ CHIẾU trước (chính xác nhất, không sợ trùng tên/sai thứ tự),
    tên đã chuẩn hóa dùng làm dự phòng khi không có/không khớp số hộ chiếu.
    Tự động BỎ QUA khách Việt Nam (nếu file có cột quốc tịch) — công dân VN
    không có "thời hạn tạm trú tại Việt Nam" nên không cần đưa vào visa_map."""
    df = pd.read_excel(io.BytesIO(visa_bytes))
    def _find(*names):
        for n in names:
            for c in df.columns:
                if _norm_nat(c) == _norm_nat(n):
                    return c
        return None
    c_full  = _find('HỌ TÊN', 'HO TEN', 'Họ và tên', 'Full Name', 'FullName')
    c_last  = _find('Last Name', 'LastName', 'Họ')
    c_first = _find('First Name', 'FirstName', 'Tên')
    c_pp    = _find('SỐ HỘ CHIẾU', 'So Ho Chieu', 'Passport', 'Passport Number', 'PassportNo')
    c_nat   = _find('QUỐC TỊCH', 'MÃ QUỐC TỊCH', 'Nationality', 'LOẠI KHÁCH')
    c_date  = _find('Visa date', 'Visadate', 'Thời hạn tạm trú',
                     'THỜI HẠN ĐƯỢC PHÉP TẠM TRÚ TẠI VIỆT NAM',
                     'Thoi han duoc phep tam tru tai Viet Nam', 'Tam tru den')
    if c_date is None:
        raise ValueError("File visa không có cột ngày visa/thời hạn tạm trú. Vui lòng kiểm tra lại file.")

    def _is_vn(val):
        """Nhận diện khách Việt Nam qua cột quốc tịch/loại khách (nếu có)."""
        s = _norm_nat(val)
        return s in ('vnm', 'viet nam', 'vietnam') or s.startswith('vnm') or s == 'viet nam'

    by_pp = {}
    by_name = {}
    n_skipped_vn = 0
    for _, r in df.iterrows():
        # Bỏ qua khách Việt Nam nếu nhận diện được qua cột quốc tịch/loại khách
        if c_nat and pd.notna(r[c_nat]) and _is_vn(r[c_nat]):
            n_skipped_vn += 1
            continue
        d = r[c_date]
        if pd.isna(d):
            continue
        # Cột ngày đọc THẲNG (month=tháng thật, day=ngày thật) — KHÔNG đảo như
        # cột Departure. Đã kiểm chứng: có ngày 25/26/31 nên day chính là ngày thật.
        if hasattr(d, 'year') and not isinstance(d, str):
            dstr = f"{d.day:02d}/{d.month:02d}/{d.year}"
        else:
            s = str(d).strip()
            p = s.split('/')
            if len(p) == 3:
                dd, mm, yy = p
                if len(yy) == 2: yy = '20' + yy
                try:
                    dstr = f"{int(dd):02d}/{int(mm):02d}/{yy}"
                except Exception:
                    continue
            else:
                try:
                    t = pd.to_datetime(s, dayfirst=True)
                    dstr = f"{t.day:02d}/{t.month:02d}/{t.year}"
                except Exception:
                    continue

        # Khớp theo số hộ chiếu — ưu tiên, chính xác nhất
        if c_pp and pd.notna(r[c_pp]):
            pp_key = _norm_pp(r[c_pp])
            if pp_key:
                by_pp.setdefault(pp_key, dstr)

        # Khớp theo tên — dự phòng khi không có/không khớp số hộ chiếu
        if c_full and pd.notna(r[c_full]):
            key = _norm_name(str(r[c_full]))
            if key:
                by_name.setdefault(key, dstr)
        else:
            ln = str(r[c_last]).strip() if c_last and pd.notna(r[c_last]) else ''
            fn = str(r[c_first]).strip() if c_first and pd.notna(r[c_first]) else ''
            # Lưu cả 2 thứ tự để khớp linh hoạt dù file NNN đảo Họ/Tên
            for combo in ((ln + ' ' + fn), (fn + ' ' + ln)):
                key = _norm_name(combo)
                if key:
                    by_name.setdefault(key, dstr)
    return {"by_pp": by_pp, "by_name": by_name, "skipped_vn": n_skipped_vn}

def build_kbtt(df_intl, visa_map=None):
    """Điền mẫu KBTT. Dòng 3 là dòng "[TEST] SAMPLE" BẮT BUỘC giữ nguyên (không
    bị ghi đè) — dữ liệu khách thật được điền bắt đầu từ dòng 4 trở xuống.
    Cột L 'THỜI HẠN ĐƯỢC PHÉP TẠM TRÚ TẠI VIỆT NAM' CHỈ điền khi có visa_map
    từ file Visa rời do lễ tân chủ động upload (khớp theo SỐ HỘ CHIẾU trước,
    tên là dự phòng). KHÔNG tự đọc cột 'TẠM TRÚ'/'TẠM TRÚ ĐẾN NGÀY' có sẵn
    trong file dữ liệu khách nữa — cột đó thường do PMS tự điền mặc định
    (vd trùng ngày đi), không phải visa đã xác nhận thật; tin nhầm có thể
    khai sai thời hạn tạm trú lên hồ sơ chính thức nộp công an.
    Trả về (wb, danh_sách_tên_không_khớp, nguồn_visa, invalid_ids)."""
    visa_map = visa_map or {}
    # Tương thích ngược: nếu visa_map là dict phẳng {tên: ngày} kiểu cũ, coi như by_name
    if isinstance(visa_map, dict) and ("by_pp" in visa_map or "by_name" in visa_map):
        by_pp = visa_map.get("by_pp", {})
        by_name = visa_map.get("by_name", {})
    else:
        by_pp = {}
        by_name = visa_map
    unmatched = []
    invalid_ids = []  # khách có SỐ GIẤY TỜ chỉ là mã tạm nội bộ (GKS/GBL...), đã bị để trống
    wb = load_workbook(io.BytesIO(load_template('kbtt')))
    ws = wb['KBTT']
    # Cấu trúc mẫu: dòng 1 = ô merge A1:L1 (tiêu đề + chú ý đỏ), dòng 2 = header,
    # dòng 3 = "[TEST] SAMPLE" BẮT BUỘC giữ nguyên, dữ liệu khách thật từ dòng 4.
    ref = [ws.cell(3,c) for c in range(1,13)]   # dùng style dòng 3 làm mẫu định dạng cho các dòng khách
    ref_styles = snapshot_styles(ref)           # chụp 1 lần, dán lại cho mọi ô
    n = len(df_intl)
    # Xóa dòng dữ liệu thừa (nếu có), luôn giữ tối thiểu tới dòng 3 (dòng TEST)
    last_data_row = 3 + n
    if ws.max_row > last_data_row:
        ws.delete_rows(last_data_row + 1, ws.max_row - last_data_row)
    for i,(_,row) in enumerate(df_intl.iterrows(),1):
        er=i+3   # dữ liệu khách thật bắt đầu dòng 4 (dòng 3 là TEST, giữ nguyên)
        ht=str(row.get('HỌ TÊN ',row.get('HỌ TÊN',''))).strip()
        ns=fmt(row['NGÀY SINH']); nd=fmt(row['NGÀY ĐẾN']); ni=fmt(row.get('NGÀY ÐI',row.get('NGÀY ĐI','')))
        ni=_fix_departure_swap(ni, nd)
        gt='M - Nam' if str(row.get('GIỚI TÍNH','')).strip()=='Nam' else 'F - Nữ'
        qt=lookup_nat_kbtt(row.get('QUỐC TỊCH',''))
        sh=str(row.get('SỐ GIẤY TỜ','')).strip(); sp=str(row.get('SỐ PHÒNG','')).strip()
        # SỐ GIẤY TỜ chỉ là mã tạm nội bộ (GKS/GBL/GKA/GBS = trẻ em dùng giấy
        # khai sinh/giấy bảo lãnh, chưa có hộ chiếu thật) — KHÔNG được nộp lên
        # hồ sơ KBTT như số hộ chiếu thật, để trống + cảnh báo thay vì điền sai.
        if sh.upper() in ('GKS','GBL','GKA','GBS'):
            invalid_ids.append(ht); sh=''
        # Cột L (12) — THỜI HẠN ĐƯỢC PHÉP TẠM TRÚ TẠI VIỆT NAM — CHỈ điền từ
        # file Visa rời (upload thủ công), khớp theo SỐ HỘ CHIẾU trước, tên là dự phòng
        vd = ''
        if by_pp or by_name:
            vd = by_pp.get(_norm_pp(sh), '') or by_name.get(_norm_name(ht), '')
            if not vd:
                unmatched.append(ht)
        vals=[i,ht,ns,'D - Ngày',gt,qt,sh,sp,nd,ni,ni,vd]
        for ci,val in enumerate(vals,1):
            cell=ws.cell(er,ci); cell.value=val if isinstance(val,int) else str(val)
            apply_style(cell, ref_styles[ci-1])
    # Bảo toàn ô merge tiêu đề + chiều cao dòng 1 (phòng khi delete_rows làm xê dịch)
    if 'A1:L1' not in [str(m) for m in ws.merged_cells.ranges]:
        try: ws.merge_cells('A1:L1')
        except Exception: pass
    ws.row_dimensions[1].height = 41.1
    # Khôi phục rich text ô A1: tiêu đề (đen) + dòng chú ý (ĐỎ) — openpyxl làm mất khi save
    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
        from openpyxl.styles.colors import Color
        from openpyxl.styles import Alignment
        a1 = ws.cell(1,1)
        a1.value = CellRichText(
            TextBlock(InlineFont(rFont='Times New Roman', sz=16, b=True),
                      'DANH SÁCH HỒ SƠ KBTT\r\n'),
            TextBlock(InlineFont(rFont='Times New Roman', sz=16, b=True, color=Color(rgb='FFFF0000')),
                      '(*Lưu ý: Người khai báo chịu trách nhiệm trước pháp luật về các nội dung thông tin cung cấp)')
        )
        a1.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    except Exception:
        pass
    # Cập nhật vùng Table1 cho khớp số dòng thực (header + dòng TEST + dữ liệu khách)
    if 'Table1' in ws.tables:
        ws.tables['Table1'].ref = f"A2:L{3 + n}"
    return wb, unmatched, ('file' if (by_pp or by_name) else None), invalid_ids

def build_vnm(df_vn):
    wb = load_workbook(io.BytesIO(load_template('vnm')))
    wsn = next((s for s in wb.sheetnames if 'KHACH' in s or 'DS' in s), wb.sheetnames[0])
    ws = wb[wsn]
    ref = [ws.cell(5,c) for c in range(1,ws.max_column+1)]
    ref_styles = snapshot_styles(ref)           # chụp 1 lần, dán lại cho mọi ô
    for r in range(ws.max_row,4,-1): ws.delete_rows(r)
    gks_cnt=0; gbl_cnt=0
    ward_unmatched=[]  # (tên khách, phường/xã gốc) không tự tra được mã — giữ raw, cần lễ tân kiểm tra
    _CUTRU_MAP={'thuongtru':'1 - Thường trú','tamtru':'2 - Tạm trú'}
    for i,(_,row) in enumerate(df_vn.iterrows(),1):
        er=i+4
        ht=str(row.get('HỌ TÊN ',row.get('HỌ TÊN',''))).strip()
        ns=fmt(row['NGÀY SINH']); nd=fmt(row['NGÀY ĐẾN']); ni=fmt(row.get('NGÀY ÐI',row.get('NGÀY ĐI','')))
        ni=_fix_departure_swap(ni, nd)
        gt='F - Nữ' if str(row.get('GIỚI TÍNH','')).strip()=='Nữ' else 'M - Nam'
        sg_raw=str(row.get('SỐ GIẤY TỜ','')).strip()
        lg_raw=str(row.get('LOẠI GIẤY TỜ','')).strip()
        is_gks='GKS' in sg_raw.upper(); is_gbl='GBL' in sg_raw.upper()
        ten_giay=''
        if is_gks:
            # Mã tạm nội bộ (chưa có giấy khai sinh thật cấp số) — theo đúng
            # danh mục DANH_MUC của mẫu, KHÔNG dùng "5 - Giấy khai sinh" (mã
            # đó dành cho giấy khai sinh CÓ số thật), ghi "9 - Giấy Tờ Khác"
            # + nêu rõ loại trong TÊN GIẤY TỜ để tránh khai sai giấy tờ.
            sg=make_code('GKS',ns); lg='9 - Giấy Tờ Khác'; ten_giay='giấy khai sinh'; gks_cnt+=1
        elif is_gbl:
            sg=make_code('GBL',ns); lg='9 - Giấy Tờ Khác'; ten_giay='giấy bảo lãnh'; gbl_cnt+=1
        elif sg_raw and sg_raw[0].isalpha():
            # Số giấy tờ bắt đầu bằng chữ cái (vd: P02628567) → là hộ chiếu
            sg=sg_raw; lg='4 - Hộ chiếu'
        else:
            sg=sg_raw; lg=LOAI_GIAY.get(lg_raw,lg_raw)
        # Dùng _fmt_room (không phải str(x).strip() thẳng) vì ô trống trong
        # Excel đọc qua pandas ra NaN (float) — str(nan)='nan' vẫn là chuỗi
        # KHÔNG rỗng nên lọt qua điều kiện "if tinh_raw" và bị tra cứu/in ra
        # chữ "nan" thẳng vào file VNM, dù thực chất khách đó không có dữ liệu.
        tinh_raw=_fmt_room(row.get('TP/TỈNH',''))
        tinh=lookup_province_vnm(tinh_raw) if tinh_raw else ''
        phuong_raw=_fmt_room(row.get('PHƯỜNG/XÃ',''))
        phuong, ward_ok, inferred_prov = lookup_ward_vnm(phuong_raw, tinh or None)
        if phuong_raw and not ward_ok:
            ward_unmatched.append((ht, phuong_raw))
        if inferred_prov and not tinh:
            tinh = inferred_prov
        dc=str(row.get('ÐỊA CHỈ',row.get('ĐỊA CHỈ',''))).strip()
        sp=str(row.get('SỐ PHÒNG','')).strip()
        dt=_fmt_room(row.get('SỐ ĐIỆN THOẠI',''))  # phòng vệ thêm nếu cột vẫn lọt qua dạng số (mất số 0 đầu)
        cutru=_CUTRU_MAP.get(_norm_nat(row.get('THƯỜNG TRÚ / TẠM TRÚ','')),'1 - Thường trú')
        vals=[i,ht,ns,gt,'VNM - Viet Nam',lg,ten_giay,sg,dt,cutru,tinh,phuong,dc,nd,ni,sp,'1 - Du lịch','','']
        for ci,val in enumerate(vals,1):
            cell=ws.cell(er,ci); cell.value=val if isinstance(val,int) else str(val)
            if ci<=len(ref_styles): apply_style(cell, ref_styles[ci-1])
    return wb, gks_cnt, gbl_cnt, ward_unmatched


# Ten duoc module nay so huu — liet ke tuong minh de `import *` lay
# duoc ca helper co gach duoi dau.
__all__ = [
    'build_kbtt', 'build_vnm', 'parse_visa_file'
]
