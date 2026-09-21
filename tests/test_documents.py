"""Sinh file thật nộp công an: KBTT (khách nước ngoài) và VNM (khách Việt).

Kiểm tra trên đúng SHEET DỮ LIỆU, không phải sheet tra cứu đi kèm mẫu —
tìm chuỗi trên cả workbook sẽ khớp nhầm vào bảng danh mục và test hoá ra vô
nghĩa (đã dính đúng bẫy này khi viết).
"""
import datetime
import io

import pandas as pd
import pytest
from openpyxl import load_workbook


def _khach_quoc_te(**kw):
    base = {'HỌ TÊN': 'JOHN SMITH', 'SỐ PHÒNG': '742', 'LOẠI KHÁCH': 'Quốc tế',
            'SỐ GIẤY TỜ': 'C1234567', 'NGÀY SINH': datetime.date(1990, 1, 1),
            'GIỚI TÍNH': 'Nam', 'QUỐC TỊCH': 'United States of America',
            'NGÀY ĐẾN': '01/09/2026', 'NGÀY ĐI': '05/09/2026'}
    base.update(kw)
    return pd.DataFrame([base])


def _khach_viet(**kw):
    base = {'HỌ TÊN': 'NGUYEN VAN A', 'SỐ PHÒNG': '538', 'LOẠI KHÁCH': 'Việt Nam',
            'SỐ GIẤY TỜ': '012345678901', 'NGÀY SINH': datetime.date(1990, 1, 1),
            'GIỚI TÍNH': 'Nam', 'QUỐC TỊCH': 'Việt Nam',
            'TP/TỈNH': 'Khánh Hòa', 'PHƯỜNG/XÃ': 'Vĩnh Hải',
            'NGÀY ĐẾN': '01/09/2026', 'NGÀY ĐI': '05/09/2026'}
    base.update(kw)
    return pd.DataFrame([base])


def _kbtt(app, df, visa_map=None):
    """→ (sheet KBTT, danh sách dòng khách thật dạng chuỗi)."""
    wb = app.build_kbtt(df, visa_map)[0]
    ws = wb['KBTT']
    rows = [' | '.join(str(c.value or '') for c in ws[r])
            for r in range(4, ws.max_row + 1)]      # dòng 1-2 tiêu đề, dòng 3 mẫu
    return wb, ws, rows


def _vnm(app, df):
    wb = app.build_vnm(df)
    wb = wb[0] if isinstance(wb, tuple) else wb
    ws = wb['DS_KHACH_VIET_NAM_LUU_TRU']
    rows = [' | '.join(str(c.value or '') for c in ws[r])
            for r in range(5, ws.max_row + 1)]      # dòng 4 là [EXAMPLE]
    return wb, ws, rows


# ── KBTT ──────────────────────────────────────────────────────────────────
def test_kbtt_xuat_file_xlsx_hop_le(app):
    wb, _ws, _rows = _kbtt(app, _khach_quoc_te())
    assert app.wb_to_bytes(wb)[:2] == b'PK'


def test_kbtt_giu_nguyen_dong_mau(app):
    """Dòng 3 '[TEST] SAMPLE' của mẫu BẮT BUỘC còn nguyên, khách thật điền từ
    dòng 4 — ghi đè lên dòng 3 là hỏng mẫu nộp."""
    _wb, ws, _rows = _kbtt(app, _khach_quoc_te())
    assert '[TEST] SAMPLE' in ' '.join(str(c.value or '') for c in ws[3])


def test_kbtt_dien_ten_va_ho_chieu(app):
    _wb, _ws, rows = _kbtt(app, _khach_quoc_te(
        **{'HỌ TÊN': 'JANE DOE', 'SỐ GIẤY TỜ': 'X9876543'}))
    assert len(rows) == 1
    assert 'JANE DOE' in rows[0] and 'X9876543' in rows[0]


def test_kbtt_khach_my_co_ma_usa(app):
    """'United States of America' trước đây không tra được → hồ sơ trống mã."""
    _wb, _ws, rows = _kbtt(app, _khach_quoc_te(
        **{'QUỐC TỊCH': 'United States of America'}))
    assert 'USA - United States of America' in rows[0]


def test_kbtt_khach_uc_khong_bi_ghi_thanh_duc(app):
    """Lỗi thật: khoá 'uc' bị 'Đức' chiếm → khách Úc khai thành khách Đức."""
    _wb, _ws, rows = _kbtt(app, _khach_quoc_te(**{'QUỐC TỊCH': 'Úc'}))
    assert 'AUS - Australia' in rows[0]
    assert 'Germany' not in rows[0]


def test_kbtt_khach_anh_ra_gbr(app):
    _wb, _ws, rows = _kbtt(app, _khach_quoc_te(**{'QUỐC TỊCH': 'United Kingdom'}))
    assert 'GBR - United Kingdom' in rows[0]
    assert 'GBD' not in rows[0]


@pytest.mark.parametrize('qt,ma', [
    ('Philippines', 'PHL'), ('Myanmar', 'MMR'), ('Nga', 'RUS'),
    ('Séc', 'CZE'), ('Đài Loan', 'CHN'), ('Đức', 'D - Germany'),
])
def test_kbtt_cac_quoc_tich_da_sua(app, qt, ma):
    _wb, _ws, rows = _kbtt(app, _khach_quoc_te(**{'QUỐC TỊCH': qt}))
    assert ma in rows[0], f'{qt} -> {rows[0]}'


def test_kbtt_khong_co_file_visa_thi_de_trong_han_tam_tru(app):
    """Cột 'TẠM TRÚ' do PMS tự điền KHÔNG đáng tin — chỉ điền khi lễ tân chủ
    động upload file Visa, tránh khai sai hạn tạm trú lên hồ sơ."""
    _wb, _ws, rows = _kbtt(app, _khach_quoc_te(**{'TẠM TRÚ': '30/12/2026'}))
    assert '30/12/2026' not in rows[0]


def test_kbtt_co_file_visa_thi_dien_theo_ho_chieu(app):
    _wb, _ws, rows = _kbtt(app, _khach_quoc_te(),
                           {'by_pp': {'C1234567': '30/12/2026'}, 'by_name': {}})
    assert '30/12/2026' in rows[0]


def test_kbtt_danh_sach_rong_van_chay(app):
    wb = app.build_kbtt(pd.DataFrame(columns=['HỌ TÊN', 'SỐ PHÒNG', 'LOẠI KHÁCH']))[0]
    assert app.wb_to_bytes(wb)[:2] == b'PK'


def test_kbtt_nhieu_khach_danh_so_thu_tu(app):
    df = pd.concat([_khach_quoc_te(**{'HỌ TÊN': f'GUEST {i}'}) for i in range(1, 4)],
                   ignore_index=True)
    _wb, _ws, rows = _kbtt(app, df)
    assert len(rows) == 3
    for i, r in enumerate(rows, 1):
        assert f'GUEST {i}' in r


# ── VNM ───────────────────────────────────────────────────────────────────
def test_vnm_xuat_file_xlsx_hop_le(app):
    wb, _ws, _rows = _vnm(app, _khach_viet())
    assert app.wb_to_bytes(wb)[:2] == b'PK'


def test_vnm_giu_nguyen_dong_example(app):
    _wb, ws, _rows = _vnm(app, _khach_viet())
    assert '[EXAMPLE]' in ' '.join(str(c.value or '') for c in ws[4])


def test_vnm_dien_ma_tinh_va_phuong(app):
    _wb, _ws, rows = _vnm(app, _khach_viet())
    assert len(rows) == 1
    assert '511 - Khánh Hòa' in rows[0]
    assert 'Vĩnh' in rows[0]


def test_vnm_dia_danh_co_chu_D(app):
    """'Đà Nẵng' từng bị chuẩn hoá thành 'a nang' → không khớp bảng."""
    _wb, _ws, rows = _vnm(app, _khach_viet(
        **{'TP/TỈNH': 'Đà Nẵng', 'PHƯỜNG/XÃ': ''}))
    assert '501 - TP. Đà Nẵng' in rows[0]


def test_vnm_dien_ten_va_giay_to(app):
    _wb, _ws, rows = _vnm(app, _khach_viet(
        **{'HỌ TÊN': 'TRAN THI B', 'SỐ GIẤY TỜ': '098765432109'}))
    assert 'TRAN THI B' in rows[0] and '098765432109' in rows[0]


def test_vnm_danh_sach_rong_van_chay(app):
    wb = app.build_vnm(pd.DataFrame(columns=['HỌ TÊN', 'SỐ PHÒNG']))
    wb = wb[0] if isinstance(wb, tuple) else wb
    assert app.wb_to_bytes(wb)[:2] == b'PK'


# ── Tách khách quốc tế / Việt Nam ─────────────────────────────────────────
def test_split_wb_tach_dung_loai_khach(app):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = 'customer'
    ws.append(['STT', 'HỌ TÊN', 'LOẠI KHÁCH', 'SỐ PHÒNG'])
    ws.append([1, 'ALPHA', 'Quốc tế', '742'])
    ws.append([2, 'BRAVO', 'Việt Nam', '538'])
    ws.append([3, 'CHARLIE', 'Quốc tế', '744'])
    raw = app.wb_to_bytes(wb)

    def noi_dung(loai):
        w = app.split_wb(load_workbook(io.BytesIO(raw)), loai)
        s = w[w.sheetnames[0]]
        return ' '.join(str(c.value or '') for row in s.iter_rows() for c in row)

    qt, vn = noi_dung('Quốc tế'), noi_dung('Việt Nam')
    assert 'ALPHA' in qt and 'CHARLIE' in qt and 'BRAVO' not in qt
    assert 'BRAVO' in vn and 'ALPHA' not in vn
