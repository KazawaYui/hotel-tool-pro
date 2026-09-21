"""ĐK14 — sổ đăng ký lưu trú nộp công an.

Sai ở đây có hai chiều nguy hiểm khác nhau:
- Bỏ sót khách thật  -> thiếu người trong sổ nộp công an.
- Để lọt dòng rác    -> sổ có người không tồn tại.
"""
import datetime

import pytest


# ── Lọc dòng không phải khách lưu trú thật ────────────────────────────────
@pytest.mark.parametrize('ten', [
    'DUMMY', 'Dummy Guest', 'DUMMY GUEST 1', 'dummy',
    'PENDING', 'Pending Arrival', 'pending arrival',
    'Water Sport', 'WATER SPORT',
])
def test_loai_dong_rac_theo_ten(app, ten):
    la_rac, ly_do = app._dk_is_dummy(ten, '742')
    assert la_rac, f'{ten!r} phải bị loại'
    assert ly_do


def test_loai_ten_trong(app):
    assert app._dk_is_dummy('', '742')[0]
    assert app._dk_is_dummy(None, '742')[0]
    assert app._dk_is_dummy('   ', '742')[0]


@pytest.mark.parametrize('phong,loai', [
    ('9000', True), ('9001', True), ('9999', True),      # phòng ảo/posting master
    ('8999', False), ('742', False), ('12A30', False),
    ('G1201', False), ('1739', False), ('', False),
])
def test_loai_theo_phong_ao(app, phong, loai):
    assert app._dk_is_dummy('NGUYEN VAN A', phong)[0] is loai


@pytest.mark.parametrize('ten', [
    'NGUYEN VAN A 1234',
    'Nguyễn Văn A 1234',          # có dấu — từng lọt qua bộ lọc
    'Trần Thị Bích 5678',
    "O'BRIEN JOHN 1234",          # dấu nháy — từng lọt
    'JOHN-PAUL SMITH 1234',       # gạch nối — từng lọt
    'LÊ VĂN C 99999',
])
def test_loai_ten_dinh_ma_dat_phong(app, ten):
    """Bộ lọc chỉ nhận [A-Za-z\\s] nên tên có dấu tiếng Việt, dấu nháy hay
    gạch nối đều lọt qua — cùng một cái tên, chỉ khác cách gõ dấu, lại cho
    kết quả khác nhau."""
    assert app._dk_is_dummy(ten, '742')[0], f'{ten!r} phải bị loại'


@pytest.mark.parametrize('ten', [
    'NGUYEN VAN A', 'Nguyễn Văn An', 'Trần Thị Bích',
    "O'BRIEN JOHN", 'JOHN-PAUL SMITH', 'LE VAN C 123',   # 3 số: chưa đủ thành mã
    'JOHN SMITH III',
])
def test_giu_lai_khach_that(app, ten):
    """Chiều nguy hiểm hơn: loại nhầm khách thật -> thiếu người trong sổ."""
    assert not app._dk_is_dummy(ten, '742')[0], f'{ten!r} là khách thật'


# ── Mã giấy tờ ────────────────────────────────────────────────────────────
@pytest.mark.parametrize('ma', ['GKS', 'GBL', 'GKA', 'GBS', 'gks', 'gbl'])
def test_ma_giay_to_chua_cap(app, ma):
    assert app._dk_is_invalid_id(ma)


@pytest.mark.parametrize('ma', [
    'C1234567', '012345678901', '', 'AB', 'ABCDE', 'GK', '123',
])
def test_ma_giay_to_that_khong_bi_coi_la_chua_cap(app, ma):
    assert not app._dk_is_invalid_id(ma)


# ── Đọc ngày từ mọi kiểu ô Excel ──────────────────────────────────────────
def test_serial_excel_khop_voi_serial2date(app):
    """Hai đường đọc ngày trong app phải cho cùng kết quả, nếu lệch thì cùng
    một khách sẽ có ngày khác nhau giữa các file xuất ra."""
    for s in (1, 40000, 45000, 45800):
        assert app.serial2date(s).date() == app._dk_to_date(s)


@pytest.mark.parametrize('v', ['rác', '32/01/2026', '00/00/0000', 'null', None, ''])
def test_ngay_khong_doc_duoc_thi_tra_none(app, v):
    """Không được đoán bừa ngày — thà trống để lễ tân tự điền."""
    assert app._dk_to_date(v) is None
    assert app._dk_fmt_date(v) == ''


def test_dinh_dang_ngay_dd_mm_yyyy(app):
    assert app._dk_fmt_date(datetime.date(2026, 3, 5)) == '05/03/2026'
    assert app._dk_fmt_date('5/3/2026') == '05/03/2026'


# ── Ghép địa chỉ thường trú ───────────────────────────────────────────────
def test_dia_chi_chi_cho_khach_viet(app):
    assert app._dk_addr_from_parts('', 'Khánh Hòa', '', 'Vĩnh Hải',
                                   '12 Trần Phú', 'USA').strip() == ''
    assert app._dk_addr_from_parts('', 'Khánh Hòa', '', 'Vĩnh Hải',
                                   '12 Trần Phú', 'VNM').strip() != ''


@pytest.mark.parametrize('rac', ['null', 'undefined', 'nan'])
def test_dia_chi_bo_gia_tri_rac_cua_pms(app, rac):
    got = app._dk_addr_from_parts('', 'Khánh Hòa', rac, 'Vĩnh Hải', '12 Trần Phú', 'VNM')
    assert rac not in got


def test_dia_chi_thu_tu_tu_chi_tiet_den_tinh(app):
    got = app._dk_addr_from_parts('', 'Khánh Hòa', 'Nha Trang', 'Vĩnh Hải',
                                  '12 Trần Phú', 'VNM')
    assert got.index('12 Trần Phú') < got.index('Vĩnh Hải') < got.index('Khánh Hòa')


# ── Giới tính ─────────────────────────────────────────────────────────────
def test_gioi_tinh_nhan_moi_cach_ghi(app):
    for v in ('Nam', 'nam', 'NAM', 'M', 'm', 'male', 'Male'):
        assert app._dk_map_gender(v) == 'Nam', v
    for v in ('Nữ', 'nữ', 'nu', 'NU', 'F', 'f', 'female'):
        assert app._dk_map_gender(v) == 'Nữ', v


def test_gioi_tinh_la_thi_giu_nguyen_de_le_tan_tu_sua(app):
    assert app._dk_map_gender('Other') == 'Other'
    assert app._dk_map_gender('') == ''


# ── Ô số bị Excel đọc thành float ─────────────────────────────────────────
def test_so_giay_to_khong_co_duoi_cham_khong(app):
    """'123456789.0' lọt vào sổ nộp công an là hỏng hồ sơ."""
    assert app._dk_cell_str(123456789.0) == '123456789'
    assert app._dk_cell_str(12345678901234.0) == '12345678901234'
    assert '.0' not in app._dk_cell_str(742.0)
