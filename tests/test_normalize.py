"""Chuẩn hoá dữ liệu đọc từ Excel: số phòng, hộ chiếu, tên, địa danh, ngày."""
import datetime

import pytest


# ── Số phòng ──────────────────────────────────────────────────────────────
@pytest.mark.parametrize('raw,expect', [
    ('742', '742'),
    (742, '742'),
    (742.0, '742'),          # Excel đọc số nguyên thành float
    ('742.0', '742'),
    ('  742  ', '742'),
    ('12A05', '12A05'),
    (None, ''),
    ('', ''),
])
def test_fmt_room(app, raw, expect):
    assert app._fmt_room(raw) == expect


def test_fmt_room_khong_tra_chuoi_nan(app):
    """NaN là truthy trong Python — `v or ''` sẽ lọt chuỗi 'nan' ra file nộp
    công an. Phải chặn riêng."""
    import pandas as pd
    assert app._fmt_room(float('nan')) == ''
    assert app._fmt_room(pd.NA if hasattr(pd, 'NA') else float('nan')) in ('', '<NA>')


@pytest.mark.parametrize('raw,expect', [
    ('12a05', '12A05'),      # hoa/thường là cùng một phòng
    ('12A05', '12A05'),
    ('742.0', '742'),
    (742.0, '742'),
    ('  g1201 ', 'G1201'),
])
def test_norm_room_chuan_hoa_hoa_thuong(app, raw, expect):
    assert app._norm_room(raw) == expect


# ── Hộ chiếu ──────────────────────────────────────────────────────────────
@pytest.mark.parametrize('raw,expect', [
    ('c1234567', 'C1234567'),
    ('C 1234 567', 'C1234567'),
    ('123456789.0', '123456789'),
    (123456789, '123456789'),
    ('  x9 ', 'X9'),
])
def test_norm_pp(app, raw, expect):
    assert app._norm_pp(raw) == expect


def test_norm_pp_nan(app):
    assert app._norm_pp(float('nan')) == ''


# ── Địa danh ──────────────────────────────────────────────────────────────
@pytest.mark.parametrize('raw,expect', [
    ('Thành phố Hà Nội', 'ha noi'),
    ('TP. Hồ Chí Minh', 'ho chi minh'),
    ('Tỉnh Khánh Hòa', 'khanh hoa'),
    ('Phường Vĩnh Hải', 'vinh hai'),
    ('Xã Cam Hải Đông', 'cam hai dong'),
    ('Thị trấn Diên Khánh', 'dien khanh'),
    ('  Đà   Nẵng  ', 'da nang'),
])
def test_norm_addr_bo_dau_va_tien_to(app, raw, expect):
    assert app._norm_addr(raw) == expect


def test_norm_addr_chuoi_rong(app):
    assert app._norm_addr(None) == ''
    assert app._norm_addr('') == ''


# ── Ngày tháng ────────────────────────────────────────────────────────────
@pytest.mark.parametrize('raw,expect', [
    ('25/12/2026', datetime.date(2026, 12, 25)),
    ('01/02/2026', datetime.date(2026, 2, 1)),      # dd/mm chứ không phải mm/dd
    (datetime.date(2026, 3, 15), datetime.date(2026, 3, 15)),
    (datetime.datetime(2026, 3, 15, 10, 30), datetime.date(2026, 3, 15)),
    (None, None),
    ('không phải ngày', None),
])
def test_to_date(app, raw, expect):
    assert app._to_date(raw) == expect


def test_serial2date(app):
    """Serial Excel: 1 = 31/12/1899, 45000 = 04/07/2023."""
    assert app.serial2date(1).date() == datetime.date(1899, 12, 31)
    assert app.serial2date(45000).date() == datetime.date(2023, 3, 15)
    assert app.serial2date(None) is None
    assert app.serial2date(0) is None            # 0 là falsy → None


@pytest.mark.parametrize('raw,expect', [
    (45000, datetime.date(2023, 3, 15)),                     # serial Excel
    (datetime.date(2026, 5, 20), datetime.date(2026, 5, 20)),
    (datetime.datetime(2026, 5, 20, 8, 0), datetime.date(2026, 5, 20)),
    ('20/05/2026', datetime.date(2026, 5, 20)),
    ('20-05-2026', datetime.date(2026, 5, 20)),
    ('', None),
    (None, None),
    ('rác', None),
    ('32/01/2026', None),                                    # ngày không tồn tại
])
def test_dk_to_date(app, raw, expect):
    assert app._dk_to_date(raw) == expect


def test_dk_fmt_date(app):
    assert app._dk_fmt_date('20/05/2026') == '20/05/2026'
    assert app._dk_fmt_date(None) == ''
    assert app._dk_fmt_date('rác') == ''


# ── _fix_date: Smile export đảo mm/dd khi cả hai số ≤ 12 ──────────────────
def test_fix_date_chuoi_giu_nguyen_dd_mm(app):
    assert app._fix_date('7/8/2026') == __import__('pandas').Timestamp(2026, 8, 7)
    assert app._fix_date('25/12/2026') == __import__('pandas').Timestamp(2026, 12, 25)


def test_fix_date_datetime_duoc_hoan_lai(app):
    """Excel lưu '7 tháng 8' thành datetime(month=7, day=8) → phải hoán lại."""
    import pandas as pd
    got = app._fix_date(datetime.datetime(2026, 7, 8))
    assert got == pd.Timestamp(2026, 8, 7)


def test_fix_date_datetime_ngay_lon_hon_12_giu_nguyen(app):
    """day=25 > 12 nên Excel không nhầm được — giữ nguyên, không hoán."""
    import pandas as pd
    got = app._fix_date(datetime.datetime(2026, 7, 25))
    assert got == pd.Timestamp(2026, 7, 25)


def test_fix_date_nam_2_chu_so(app):
    import pandas as pd
    assert app._fix_date('7/8/26') == pd.Timestamp(2026, 8, 7)


def test_fix_date_rong(app):
    assert app._fix_date(None) is None


# ── _fix_departure_swap: sửa ngày đi trước ngày đến ───────────────────────
def test_departure_swap_sua_khi_di_truoc_den(app):
    """Đi 05/08 trước đến 07/08 → hoán thành 08/05? Không — hoán dd/mm của
    ngày đi: 05/08 → 08/05 vẫn trước. Dùng ca thực tế: đến 07/08, đi 09/08
    bị ghi nhầm 08/09 → không cần sửa. Ca cần sửa: đến 09/08, đi 08/09 sai."""
    # đi 03/09 (3 tháng 9) trước đến 05/09 → hoán thành 09/03 vẫn trước → giữ nguyên
    assert app._fix_departure_swap('03/09/2026', '05/09/2026') == '03/09/2026'
    # đi 09/03 (9 tháng 3) trước đến 05/09 → hoán thành 03/09 vẫn trước → giữ nguyên
    assert app._fix_departure_swap('09/03/2026', '05/09/2026') == '09/03/2026'
    # ca hoán CÓ tác dụng: đến 05/09, đi ghi 06/09 → không sai, giữ nguyên
    assert app._fix_departure_swap('06/09/2026', '05/09/2026') == '06/09/2026'


def test_departure_swap_co_tac_dung(app):
    """Đến 10/09/2026, đi bị ghi 09/10 (9 tháng 10) → đọc ra 09/10 SAU đến,
    không sửa. Ca sửa thật: đến 09/10/2026, đi ghi 10/09 → 10/09 trước
    09/10 → hoán thành 09/10 = đúng bằng ngày đến → nhận."""
    assert app._fix_departure_swap('10/09/2026', '09/10/2026') == '09/10/2026'


def test_departure_swap_khong_dung_khi_khong_chac(app):
    # ngày > 12 thì không thể là lỗi đảo → giữ nguyên
    assert app._fix_departure_swap('25/01/2026', '30/01/2026') == '25/01/2026'
    # dd == mm thì hoán vô nghĩa → giữ nguyên
    assert app._fix_departure_swap('05/05/2026', '10/05/2026') == '05/05/2026'
    # đã đúng thứ tự → không đụng
    assert app._fix_departure_swap('15/09/2026', '10/09/2026') == '15/09/2026'
    # thiếu dữ liệu → giữ nguyên
    assert app._fix_departure_swap('', '10/09/2026') == ''
    assert app._fix_departure_swap('10/09/2026', '') == '10/09/2026'
    # chuỗi rác → giữ nguyên
    assert app._fix_departure_swap('rác', '10/09/2026') == 'rác'


# ── Giới tính ĐK14 ────────────────────────────────────────────────────────
@pytest.mark.parametrize('raw,expect', [
    ('Nam', 'Nam'), ('nam', 'Nam'), ('M', 'Nam'), ('m', 'Nam'), ('male', 'Nam'),
    ('Nữ', 'Nữ'), ('nu', 'Nữ'), ('F', 'Nữ'), ('f', 'Nữ'), ('female', 'Nữ'),
])
def test_dk_map_gender(app, raw, expect):
    assert app._dk_map_gender(raw) == expect


def test_dk_map_gender_la_thi_giu_nguyen(app):
    assert app._dk_map_gender('Other') == 'Other'
    assert app._dk_map_gender(None) == ''


# ── Ô số bị đọc thành float ───────────────────────────────────────────────
@pytest.mark.parametrize('raw,expect', [
    (123456789.0, '123456789'),
    (123456789, '123456789'),
    ('123456789', '123456789'),
    (1.5, '1.5'),
    (None, ''),
    ('  abc  ', 'abc'),
])
def test_dk_cell_str(app, raw, expect):
    assert app._dk_cell_str(raw) == expect
