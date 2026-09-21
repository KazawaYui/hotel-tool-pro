"""Tra mã quốc tịch cho hồ sơ KBTT nộp công an.

Sai mã ở đây nguy hiểm hơn thiếu mã: thiếu thì app hiện cảnh báo "Quốc tịch
chưa có mã" cho lễ tân tự điền, còn sai thì hồ sơ vẫn trông hợp lệ và trôi
thẳng lên công an.
"""
import pytest


# ── Lỗi từng tra ra SAI NƯỚC ──────────────────────────────────────────────
def test_uc_khong_duoc_ra_duc(app):
    """Khoá 'uc' vốn sinh từ 'Đức' theo cách chuẩn hoá cũ nuốt chữ đ, chiếm
    mất chỗ của 'Úc' → khách Úc bị khai thành khách Đức, im lặng."""
    assert app.lookup_nat_kbtt('Úc') == 'AUS - Australia'
    assert app.lookup_nat_kbtt('Đức') == 'D - Germany'


def test_united_kingdom_ra_gbr_khong_phai_gbd(app):
    """'United Kingdom' từng trỏ sang GBD = British Territories Citizen."""
    assert app.lookup_nat_kbtt('United Kingdom') == 'GBR - United Kingdom'
    assert (app.lookup_nat_kbtt('United Kingdom British Territories Citizen')
            == 'GBD - United Kingdom British Territories Citizen')
    assert app.lookup_nat_kbtt('Anh') == 'GBR - United Kingdom'


# ── Tên tiếng Anh đầy đủ do PMS xuất ra ───────────────────────────────────
@pytest.mark.parametrize('ten,expect', [
    ('United States of America', 'USA - United States of America'),
    ('Philippines', 'PHL - Philippines'),
    ('Myanmar', 'MMR - Myanmar'),
    ('Syrian Arab Republic', 'SYR - Syrian Arab Republic'),
    ('Libyan Arab Jamahiriya', 'LBY - Libyan Arab Jamahiriya'),
    ('Korea Democratic Peoples Republic of', 'PRK - Korea Democratic Peoples Republic of'),
    ('Iran Ilasmic Republic of', 'IRN - Iran Ilasmic Republic of'),
    ('El Salvado', 'SLV - El Salvado'),
    ('Holy See (Vatican City State )', 'VAT - Holy See (Vatican City State )'),
    ('Australia', 'AUS - Australia'),
    ('Germany', 'D - Germany'),
])
def test_ten_tieng_anh_day_du(app, ten, expect):
    assert app.lookup_nat_kbtt(ten) == expect


# ── Tên tiếng Việt thông dụng ─────────────────────────────────────────────
@pytest.mark.parametrize('ten,expect', [
    ('Nga', 'RUS - Russia'),
    ('Séc', 'CZE - Czech Republic'),
    ('Đông Timor', 'TLS - Timor Leste'),
    ('Đài Loan', 'CHN - China'),          # theo đúng quy ước sẵn có của bảng
    ('Trung Quốc', 'CHN - China'),
    ('Mỹ', 'USA - United States of America'),
    ('Ấn Độ', 'IND - India'),
    ('Đan Mạch', 'DNK - Denmark'),
    ('Hàn Quốc', 'KOR - Korea (South)'),
    ('Nhật Bản', 'JPN - Japan'),
    ('Thái Lan', 'THA - Thailand'),
    ('Pháp', 'FRA - France'),
])
def test_ten_tieng_viet(app, ten, expect):
    assert app.lookup_nat_kbtt(ten) == expect


# ── Bất biến: MỌI mục trong bảng phải tra được bằng chính tên của nó ──────
def test_moi_muc_tra_duoc_bang_ten_hien_thi(app):
    """Chặn kiểu lỗi 'United States of America' không ra mã: bảng có mục
    nhưng không tra được bằng đúng tên hiển thị của mục đó."""
    sai = []
    for val in set(app.NAT_NORM.values()):
        if ' - ' not in val:
            continue
        ten = val.split(' - ', 1)[1].strip()
        got = app.lookup_nat_kbtt(ten)
        if got != val:
            sai.append((ten, val, got))
    assert not sai, f"{len(sai)} mục không tra được bằng tên của chính nó: {sai[:5]}"


def test_khong_co_khoa_chet(app):
    """Mọi khoá phải tự chuẩn hoá về chính nó (bằng 1 trong 2 cách chuẩn hoá),
    nếu không thì mục đó vĩnh viễn không ai tra tới."""
    chet = [k for k in app.NAT_NORM
            if app._norm_nat(k) != k and app._norm_nat_legacy(k) != k]
    assert not chet, f"khoá không bao giờ khớp được: {chet[:10]}"


# ── Không đoán bừa khi thiếu dữ liệu ──────────────────────────────────────
@pytest.mark.parametrize('ten', ['Campuchia', 'Lào', 'Nam Phi'])
def test_nuoc_chua_co_trong_bang_thi_giu_nguyen(app, ten):
    """3 nước này thật sự KHÔNG có trong bảng KBTT. Phải trả nguyên văn để
    validate_guests() bật cảnh báo cho lễ tân tự điền — tuyệt đối không tự
    đoán sang mã nước khác."""
    assert app.lookup_nat_kbtt(ten) == ten


def test_giu_nguyen_dang_ma_co_san(app):
    assert app.lookup_nat_kbtt('USA - United States of America') == 'USA - United States of America'


def test_rong(app):
    assert app.lookup_nat_kbtt('') == ''
    assert app.lookup_nat_kbtt(None) == ''
