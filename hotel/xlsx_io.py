"""Đọc/ghi file QLLT: quy đổi tỷ giá, chia giá phòng connecting, tách theo loại khách."""

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

from .common import _fmt_room, wb_to_bytes
from .assets import load_template
from .lookups import _norm_nat

# ── Cặp phòng connecting — tự động chia đều ĐƠN GIÁ khi 1 phòng bị bỏ trống ──
# Danh sách cố định theo sơ đồ tầng khách sạn (nguồn: C_p_CNT.xlsx — 13 tầng
# 5→17, mỗi tầng 6 cặp = 78 cặp; toà nhà không có tầng 13). Đây là kết cấu vật
# lý của toà nhà, gần như không đổi — để hằng số thay vì nạp file mỗi lần, dễ
# đối chiếu/sửa khi khách sạn cải tạo.
# CHỈ áp dụng cho các cặp có TRONG danh sách này — cặp phòng nào phát sinh
# ngoài danh sách (VD: tầng chưa được liệt kê) thì GIỮ NGUYÊN giá, không đoán.
# Mỗi dòng dưới đây = 1 dòng trong file nguồn; 2 dòng liên tiếp là 1 cặp
# connecting (đúng theo các ô đã gộp A5:A6, A7:A8… trong file).
#            F5   F6   F7   F8   F9   F10   F11   F12   F12A     F14   F15   F16   F17
_CNT_FLOOR_ROWS = [
    (538, 638, 738, 838, 934, 1034, 1134, 1230, '12A30', 1428, 1528, 1628, 1728),
    (540, 640, 740, 840, 936, 1036, 1136, 1232, '12A32', 1430, 1530, 1630, 1730),
    (542, 642, 742, 842, 938, 1038, 1138, 1234, '12A34', 1432, 1532, 1632, 1732),
    (544, 644, 744, 844, 940, 1040, 1140, 1236, '12A36', 1434, 1534, 1634, 1734),
    (546, 646, 746, 846, 942, 1042, 1142, 1238, '12A38', 1436, 1536, 1636, 1736),
    (548, 648, 748, 848, 944, 1044, 1144, 1240, '12A40', 1438, 1538, 1638, 1738),
    (541, 641, 741, 841, 937, 1037, 1137, 1233, '12A33', 1433, 1529, 1629, 1729),
    (543, 643, 743, 843, 939, 1039, 1139, 1235, '12A35', 1435, 1531, 1631, 1731),
    (545, 645, 745, 845, 941, 1041, 1141, 1237, '12A37', 1437, 1533, 1633, 1733),
    (547, 647, 747, 847, 943, 1043, 1143, 1239, '12A39', 1439, 1535, 1635, 1735),
    (549, 649, 749, 849, 945, 1045, 1145, 1241, '12A41', 1440, 1537, 1637, 1737),
    (550, 650, 750, 850, 946, 1046, 1146, 1242, '12A42', 1441, 1539, 1639, 1739),
]
# So khớp bằng chữ HOA: PMS xuất số phòng chữ thường ("12a30") còn sơ đồ tầng
# ghi hoa ("12A30") — không chuẩn hoá thì cặp tầng 12A không bao giờ khớp.
CONNECTING_ROOM_PAIRS = []
for _fi in range(len(_CNT_FLOOR_ROWS[0])):  # duyệt hết số cột tầng, không cắt bớt
    for _pi in range(0, len(_CNT_FLOOR_ROWS), 2):  # 2 dòng liên tiếp = 1 cặp
        CONNECTING_ROOM_PAIRS.append(
            (_fmt_room(_CNT_FLOOR_ROWS[_pi][_fi]).upper(),
             _fmt_room(_CNT_FLOOR_ROWS[_pi + 1][_fi]).upper()))

def split_connecting_room_prices(xlsx_bytes):
    """Tự động chia đều ĐƠN GIÁ cho cặp phòng connecting khi CHUNG MÃ CHECKIN
    (cùng 1 lượt đặt) mà chỉ 1 phòng được ghi giá, phòng kia bằng 0 — PMS
    thường gộp giá cả 2 phòng connecting vào 1 phòng khi đặt chung.

    Quy tắc CHỈ áp dụng khi cả 2 điều kiện đúng:
    1. Cặp (roomA, roomB) có TRONG CONNECTING_ROOM_PAIRS (không đoán ngoài
       danh sách — cặp phát sinh ngoài ý muốn thì GIỮ NGUYÊN giá).
    2. Cả 2 phòng CHUNG 1 MÃ CHECKIN, và đúng 1 phòng có tổng giá > 0, phòng
       còn lại tổng giá = 0.

    Tổng lẻ (không chia hết 2) thì phòng SỐ NHỎ HƠN nhận phần làm tròn XUỐNG,
    phòng SỐ LỚN HƠN nhận phần làm tròn LÊN — cộng lại đúng bằng tổng ban đầu
    (đối chiếu đúng theo file QLLT mẫu đã tách tay: VD 8.313.043 → 4.156.521 +
    4.156.522). Trả về (bytes file đã sửa, danh sách bản ghi đã chia để hiển
    thị minh bạch cho người dùng kiểm tra lại)."""
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active
    headers = [c.value for c in ws[1]]
    headers_norm = [_norm_nat(h) if h else None for h in headers]

    def _col(name):
        target = _norm_nat(name)
        return next((i + 1 for i, hn in enumerate(headers_norm) if hn == target), None)

    ci_checkin, ci_room, ci_gia = _col('MÃ CHECKIN'), _col('SỐ PHÒNG'), _col('ĐƠN GIÁ')
    if not (ci_checkin and ci_room and ci_gia):
        return xlsx_bytes, []  # thiếu cột cần thiết — không đụng vào file

    # Gom danh sách dòng theo (mã checkin, số phòng) để tính tổng giá mỗi phòng
    # và biết CHÍNH XÁC dòng nào đang giữ giá trị > 0 cần sửa.
    groups = {}  # (checkin, room) -> list các dòng (row index)
    for r in range(2, ws.max_row + 1):
        checkin = ws.cell(r, ci_checkin).value
        room = _fmt_room(ws.cell(r, ci_room).value).upper()
        if checkin is None or str(checkin).strip() == '' or not room:
            continue
        checkin = str(checkin).strip()
        groups.setdefault((checkin, room), []).append(r)

    rooms_by_checkin = {}
    for (checkin, room) in groups:
        rooms_by_checkin.setdefault(checkin, set()).add(room)

    report = []
    for checkin, rooms in rooms_by_checkin.items():
        for room_a, room_b in CONNECTING_ROOM_PAIRS:
            if room_a not in rooms or room_b not in rooms:
                continue
            rows_a = groups[(checkin, room_a)]
            rows_b = groups[(checkin, room_b)]

            def _total(rows):
                s = 0
                for r in rows:
                    v = ws.cell(r, ci_gia).value
                    if isinstance(v, (int, float)) and not pd.isna(v):
                        s += v
                return s

            total_a, total_b = _total(rows_a), _total(rows_b)
            if not ((total_a > 0) != (total_b > 0)):
                continue  # cả 2 đều có giá, hoặc cả 2 đều 0 — không đụng vào
            nonzero_room, nonzero_rows, nonzero_total = (
                (room_a, rows_a, total_a) if total_a > 0 else (room_b, rows_b, total_b))
            zero_room, zero_rows = (room_b, rows_b) if nonzero_room == room_a else (room_a, rows_a)

            lower, higher = sorted([room_a, room_b])
            half_floor = int(nonzero_total) // 2
            half_ceil = int(nonzero_total) - half_floor
            new_val = {lower: half_floor, higher: half_ceil}

            # Dòng đang giữ giá trị > 0 → sửa thành phần chia của chính phòng đó.
            # Dòng ĐẦU TIÊN của phòng đang 0 → điền phần chia còn lại. Các dòng
            # khác (khách đi cùng phòng) giữ nguyên 0 — không đụng.
            for r in nonzero_rows:
                v = ws.cell(r, ci_gia).value
                if isinstance(v, (int, float)) and v == nonzero_total:
                    ws.cell(r, ci_gia).value = new_val[nonzero_room]
                    break
            ws.cell(zero_rows[0], ci_gia).value = new_val[zero_room]

            report.append({'checkin': checkin, 'room_a': room_a, 'room_b': room_b,
                           'total': int(nonzero_total),
                           'split_a': new_val[room_a], 'split_b': new_val[room_b]})

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), report

# ── Processing ────────────────────────────────────────────────────────────
def process_xlsx(xlsx_bytes, rate):
    """Điền dữ liệu file đầu vào lên FILE MẪU customer (QLLT) — giữ nguyên 100%
    template: sheet 'customer' + sheet 'Danh-muc', định dạng header, style ô dữ liệu.
    Map cột theo tên (không phân biệt hoa thường / dấu / khoảng trắng), quy đổi
    ĐƠN GIÁ (USD → VND, tô vàng ô đã đổi), đánh lại STT."""
    src_wb = load_workbook(io.BytesIO(xlsx_bytes))
    src_ws = src_wb.active

    # Map header nguồn (chuẩn hóa) → chỉ số cột nguồn
    src_map = {}
    for c in src_ws[1]:
        if c.value is not None and str(c.value).strip():
            src_map.setdefault(_norm_nat(c.value), c.column)

    # Nạp template QLLT (customer + Danh-muc), dòng 2 là mẫu định dạng ô dữ liệu
    wb = load_workbook(io.BytesIO(load_template('customer')))
    ws = wb['customer']
    n_cols = ws.max_column
    ref = [ws.cell(2, ci) for ci in range(1, n_cols + 1)]
    ref_styles = [(copy(c.font), copy(c.fill), copy(c.border), copy(c.alignment), c.number_format) for c in ref]

    # Với mỗi cột template, tìm cột nguồn tương ứng theo tên
    headers = [ws.cell(1, ci).value for ci in range(1, n_cols + 1)]
    headers_norm = [_norm_nat(h) if h else None for h in headers]
    col_src = [src_map.get(hn) if hn else None for hn in headers_norm]
    # Khớp tên tuyệt đối thất bại với 1 số cột nếu file nguồn đặt tên rút gọn
    # hơn mẫu QLLT (vd nguồn "CỬA KHẨU" ↔ mẫu "CỬA KHẨU NHẬP CẢNH", nguồn
    # "TẠM TRÚ" ↔ mẫu "TẠM TRÚ ĐẾN NGÀY") → mất trắng dữ liệu cột đó dù có
    # trong file nguồn. Với cột còn thiếu, thử khớp NỚI LỎNG: 1 trong 2 tên là
    # phần đầu của tên còn lại (đủ dài để tránh khớp nhầm với tên ngắn/mơ hồ),
    # mỗi cột nguồn chỉ được dùng khớp 1 lần để tránh nhầm giữa 2 cột đích.
    _used_src_cols = {c for c in col_src if c}
    for i, hn in enumerate(headers_norm):
        if col_src[i] or not hn:
            continue
        for src_key, src_col in src_map.items():
            if src_col in _used_src_cols:
                continue
            if len(src_key) >= 4 and (hn.startswith(src_key) or src_key.startswith(hn)):
                col_src[i] = src_col
                _used_src_cols.add(src_col)
                break
    don_gia_idx = next((i + 1 for i, h in enumerate(headers) if h and _norm_nat(h) == _norm_nat('ĐƠN GIÁ')), None)

    ws.delete_rows(2)  # bỏ dòng mẫu

    conv = 0
    er = 1
    for row in src_ws.iter_rows(min_row=2, max_row=src_ws.max_row):
        if all(c.value is None or str(c.value).strip() == '' for c in row):
            continue
        er += 1
        for ci in range(1, n_cols + 1):
            cell = ws.cell(er, ci)
            f, fl, b, a, nf = ref_styles[ci - 1]
            # Gán thẳng, không copy lại từng ô — xem giải thích ở common.cp().
            # ref_styles đã là bản sao tạo MỘT LẦN từ dòng mẫu, đủ để không
            # đụng vào style gốc của template.
            cell.font = f; cell.fill = fl; cell.border = b
            cell.alignment = a; cell.number_format = nf
            sc = col_src[ci - 1]
            if sc is not None and row[sc - 1].value is not None:
                cell.value = row[sc - 1].value
        ws.cell(er, 1).value = er - 1  # STT đánh lại
        if don_gia_idx:
            dg = ws.cell(er, don_gia_idx)
            if dg.value and isinstance(dg.value, (int, float)) and 0 < dg.value < 1000:
                dg.value = round(dg.value * rate)
                conv += 1
    return wb, conv

def split_wb(wb, loai):
    wb2 = load_workbook(io.BytesIO(wb_to_bytes(wb)))
    ws2 = wb2.active
    lc = next(c.column for c in ws2[1] if c.value=='LOẠI KHÁCH')
    dels = [row[0].row for row in ws2.iter_rows(min_row=2,max_row=ws2.max_row)
            if row[lc-1].value != loai]
    for r in reversed(dels): ws2.delete_rows(r)
    for i, row in enumerate(ws2.iter_rows(min_row=2,max_row=ws2.max_row),1): row[0].value=i
    return wb2


# Ten duoc module nay so huu — liet ke tuong minh de `import *` lay
# duoc ca helper co gach duoi dau.
__all__ = [
    'CONNECTING_ROOM_PAIRS', '_CNT_FLOOR_ROWS', 'process_xlsx',
    'split_connecting_room_prices', 'split_wb'
]
