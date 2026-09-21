"""Lịch âm Việt Nam + chọn mùa/dịp lễ cho màn chào.

Thuật toán Hồ Ngọc Đức. Các mốc dưới đây đối chiếu với lịch công bố, không
phải tự thuật toán sinh ra rồi lấy làm chuẩn.
"""
import datetime

import pytest


# ── Mùng 1 Tết các năm (dương lịch) ───────────────────────────────────────
@pytest.mark.parametrize('duong,am', [
    (datetime.date(2020, 1, 25), (1, 1, 2020)),
    (datetime.date(2021, 2, 12), (1, 1, 2021)),
    (datetime.date(2022, 2, 1),  (1, 1, 2022)),
    (datetime.date(2023, 1, 22), (1, 1, 2023)),
    (datetime.date(2024, 2, 10), (1, 1, 2024)),
    (datetime.date(2025, 1, 29), (1, 1, 2025)),
    (datetime.date(2026, 2, 17), (1, 1, 2026)),
    (datetime.date(2027, 2, 6),  (1, 1, 2027)),
    (datetime.date(2028, 1, 26), (1, 1, 2028)),
    (datetime.date(2029, 2, 13), (1, 1, 2029)),
    (datetime.date(2030, 2, 2),  (1, 1, 2030)),
])
def test_mung_1_tet(app, duong, am):
    d, m, y, _leap = app.solar_to_lunar(duong)
    assert (d, m, y) == am


# ── Rằm tháng 8 (Trung Thu) ───────────────────────────────────────────────
@pytest.mark.parametrize('duong', [
    datetime.date(2024, 9, 17),
    datetime.date(2025, 10, 6),
    datetime.date(2026, 9, 25),
])
def test_ram_thang_tam(app, duong):
    d, m, _y, leap = app.solar_to_lunar(duong)
    assert (d, m, leap) == (15, 8, 0)


# ── Chọn mùa / dịp lễ ─────────────────────────────────────────────────────
def test_tet_bao_gom_ca_truoc_va_sau_giao_thua(app):
    """Mùng 1→7 tháng Giêng, và 28→30 tháng Chạp."""
    assert app._season_key(datetime.date(2026, 2, 17)) == 'tet'   # mùng 1
    assert app._season_key(datetime.date(2026, 2, 23)) == 'tet'   # mùng 7
    assert app._season_key(datetime.date(2026, 2, 15)) == 'tet'   # 29 Chạp


def test_trung_thu(app):
    assert app._season_key(datetime.date(2026, 9, 25)) == 'trungthu'   # rằm
    assert app._season_key(datetime.date(2026, 9, 24)) == 'trungthu'   # 14
    assert app._season_key(datetime.date(2026, 9, 26)) == 'trungthu'   # 16


def test_giang_sinh(app):
    assert app._season_key(datetime.date(2026, 12, 24)) == 'noel'
    assert app._season_key(datetime.date(2026, 12, 25)) == 'noel'
    assert app._season_key(datetime.date(2026, 12, 20)) == 'noel'
    assert app._season_key(datetime.date(2026, 12, 27)) != 'noel'


@pytest.mark.parametrize('d,mua', [
    (datetime.date(2026, 3, 10), 'spring'),
    (datetime.date(2026, 6, 10), 'summer'),
    (datetime.date(2026, 9, 10), 'autumn'),
    (datetime.date(2026, 11, 10), 'winter'),
])
def test_bon_mua_thuong(app, d, mua):
    assert app._season_key(d) == mua


def test_moi_mua_deu_co_nhan(app):
    for key in ('spring', 'summer', 'autumn', 'winter', 'tet', 'trungthu', 'noel'):
        assert key in app.SEASON_LABEL and app.SEASON_LABEL[key].strip()


@pytest.mark.parametrize('gio,key', [
    (5, 'dawn'), (7, 'dawn'), (8, 'day'), (15, 'day'),
    (16, 'dusk'), (18, 'dusk'), (19, 'night'), (23, 'night'), (4, 'night'),
])
def test_khung_gio(app, gio, key):
    assert app._hour_key(gio) == key


def test_moi_mua_lay_duoc_anh_nen(app):
    """Mùa nào chưa có ảnh riêng phải rơi về ảnh mặc định, không được lỗi."""
    for key in ('spring', 'summer', 'autumn', 'winter', 'tet', 'trungthu', 'noel'):
        for theme in ('dark', 'light'):
            uri = app._season_bg_data_uri(key, theme)
            assert uri.startswith('data:image/'), f"{key}/{theme} không ra ảnh"
