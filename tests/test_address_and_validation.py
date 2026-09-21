"""Địa danh (VNM), nhận diện dòng rác (ĐK14) và kiểm tra chất lượng dữ liệu."""
import datetime

import pandas as pd
import pytest


# ── Tra tỉnh/thành cho mẫu VNM ────────────────────────────────────────────
@pytest.mark.parametrize('raw,ma', [
    ('Hà Nội', '101'),
    ('Đà Nẵng', '501'),
    ('Khánh Hòa', '511'),
    ('Đồng Nai', '713'),
    ('Đắk Lắk', '605'),
    ('Điện Biên', '302'),
])
def test_tra_tinh(app, raw, ma):
    got = app.lookup_province_vnm(raw)
    assert got.startswith(ma + ' - '), f'{raw} -> {got}'


def test_tra_tinh_chiu_duoc_tien_to(app):
    """Người dùng gõ kèm 'Tỉnh'/'TP.' vẫn phải khớp."""
    assert app.lookup_province_vnm('Tỉnh Khánh Hòa') == app.lookup_province_vnm('Khánh Hòa')
    assert app.lookup_province_vnm('TP. Đà Nẵng') == app.lookup_province_vnm('Đà Nẵng')
    assert app.lookup_province_vnm('  đà nẵng  ') == app.lookup_province_vnm('Đà Nẵng')


def test_tra_tinh_khong_doan_bua(app):
    """Không khớp thì trả nguyên văn để lễ tân tự sửa, không đoán tỉnh khác."""
    assert app.lookup_province_vnm('Tỉnh Không Tồn Tại') == 'Tỉnh Không Tồn Tại'
    assert app.lookup_province_vnm('') == ''


def test_tra_phuong_xa(app):
    val, khop, _suy = app.lookup_ward_vnm('Vĩnh Hải', app.lookup_province_vnm('Khánh Hòa'))
    assert khop and ' - ' in val


def test_tra_phuong_mo_ho_thi_tra_nguyen_van(app):
    val, khop, _ = app.lookup_ward_vnm('Phường Không Có Thật')
    assert not khop and val == 'Phường Không Có Thật'


def test_chu_D_khong_bi_nuot_trong_ten_phuong(app):
    """'Đông'/'Đức' từng bị chuẩn hoá thành 'ong'/'uc' — mất hẳn chữ Đ."""
    assert app._norm_addr('Xã Cam Hải Đông') == 'cam hai dong'
    assert app._norm_addr('Phường Đức Thắng') == 'duc thang'


# ── ĐK14: nhận diện dòng không phải khách thật ────────────────────────────
@pytest.mark.parametrize('ten,phong', [
    ('', '101'),
    ('DUMMY GUEST', '101'),
    ('Pending Arrival', '102'),
    ('Water Sport', '103'),
    ('NGUYEN VAN A 12345', '104'),      # tên dính mã đặt phòng
    ('Khách thật', '9001'),             # phòng ảo ≥ 9000
    ('Khách thật', '9999'),
])
def test_dong_rac_bi_loai(app, ten, phong):
    la_rac, ly_do = app._dk_is_dummy(ten, phong)
    assert la_rac and ly_do


@pytest.mark.parametrize('ten,phong', [
    ('NGUYEN VAN A', '742'),
    ('John Smith', '12A30'),
    ('Trần Thị B', '1428'),
    ('Khách thật', '8999'),             # ngay dưới ngưỡng phòng ảo
])
def test_khach_that_khong_bi_loai(app, ten, phong):
    la_rac, _ = app._dk_is_dummy(ten, phong)
    assert not la_rac


# ── ĐK14: mã giấy tờ chưa cấp ─────────────────────────────────────────────
@pytest.mark.parametrize('ma,invalid', [
    ('GKS', True), ('GBL', True), ('GKA', True), ('GBS', True),
    ('gks', True),
    ('C1234567', False), ('123456789', False), ('', False), ('AB', False),
])
def test_ma_giay_to_chua_cap(app, ma, invalid):
    assert app._dk_is_invalid_id(ma) is invalid


# ── ĐK14: ghép địa chỉ từ các mảnh ────────────────────────────────────────
def test_ghep_dia_chi_bo_mieng_rac(app):
    """'null'/'undefined'/'nan' từ PMS không được lọt vào sổ nộp công an."""
    got = app._dk_addr_from_parts('', 'Khánh Hòa', 'null', 'Vĩnh Hải', '12 Trần Phú', 'VNM')
    assert 'null' not in got and 'undefined' not in got
    assert 'Vĩnh Hải' in got and 'Khánh Hòa' in got


def test_ghep_dia_chi_chi_danh_cho_khach_viet(app):
    """Khách nước ngoài để trống địa chỉ thường trú theo đúng mẫu ĐK14."""
    assert app._dk_addr_from_parts('', 'Khánh Hòa', '', 'Vĩnh Hải',
                                   '12 Trần Phú', 'USA').strip() == ''


def test_ghep_dia_chi_co_noi_cu_tru(app):
    got = app._dk_addr_from_parts('Thường trú', 'Khánh Hòa', '', 'Vĩnh Hải', '12 Trần Phú', 'VNM')
    assert got.startswith('Thường trú:')


def test_ghep_dia_chi_rong(app):
    assert app._dk_addr_from_parts('', '', '', '', '', 'VNM').strip() == ''


# ── Kiểm tra chất lượng dữ liệu trước khi nộp ─────────────────────────────
def _df(**kw):
    base = {'HỌ TÊN': 'NGUYEN VAN A', 'SỐ PHÒNG': '742', 'LOẠI KHÁCH': 'Quốc tế',
            'SỐ GIẤY TỜ': 'C1234567', 'NGÀY SINH': datetime.date(1990, 1, 1),
            'GIỚI TÍNH': 'Nam', 'QUỐC TỊCH': 'United States of America',
            'NGÀY ĐẾN': '01/09/2026', 'NGÀY ĐI': '05/09/2026'}
    base.update(kw)
    return pd.DataFrame([base])


def test_khach_du_thong_tin_thi_khong_bao_loi(app):
    assert len(app.validate_guests(_df())) == 0


def test_thieu_ho_ten_la_loi_do(app):
    iss = app.validate_guests(_df(**{'HỌ TÊN': ''}))
    assert '🔴' in set(iss['Mức độ'])


def test_khach_quoc_te_thieu_ho_chieu_la_loi_do(app):
    iss = app.validate_guests(_df(**{'SỐ GIẤY TỜ': ''}))
    assert '🔴' in set(iss['Mức độ'])


def test_khach_viet_thieu_giay_to_chi_la_canh_bao(app):
    iss = app.validate_guests(_df(**{'LOẠI KHÁCH': 'Việt Nam', 'SỐ GIẤY TỜ': ''}))
    assert '🔴' not in set(iss['Mức độ'])


def test_ngay_di_truoc_ngay_den_la_loi_do(app):
    iss = app.validate_guests(_df(**{'NGÀY ĐẾN': '05/09/2026', 'NGÀY ĐI': '01/09/2026'}))
    assert any('TRƯỚC ngày đến' in v for v in iss['Vấn đề'])
    assert '🔴' in set(iss['Mức độ'])


def test_quoc_tich_khong_co_ma_bi_canh_bao(app):
    """3 nước còn thiếu trong bảng phải hiện cảnh báo cho lễ tân tự điền."""
    iss = app.validate_guests(_df(**{'QUỐC TỊCH': 'Campuchia'}))
    assert any('chưa có mã' in v for v in iss['Vấn đề'])


def test_quoc_tich_co_ma_khong_bi_canh_bao(app):
    for qt in ('Úc', 'United Kingdom', 'Philippines', 'Nga'):
        iss = app.validate_guests(_df(**{'QUỐC TỊCH': qt}))
        assert not any('chưa có mã' in v for v in iss['Vấn đề']), qt
