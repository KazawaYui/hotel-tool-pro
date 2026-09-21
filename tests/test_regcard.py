"""Regcard (phiếu đăng ký khách) và số đêm lưu trú.

Regcard in ra đưa khách ký nên sai tên/ngày là khách nhìn thấy ngay.
"""
import datetime

import pandas as pd
import pytest


# ── Tên khách ─────────────────────────────────────────────────────────────
@pytest.mark.parametrize('raw,expect', [
    ('SMITH, JOHN', 'SMITH, JOHN'),
    ('SMITH,', 'SMITH'),                 # dấu phẩy thừa ở cuối
    ('  SMITH JOHN  ', 'SMITH JOHN'),
    ('SMITH JOHN, ', 'SMITH JOHN'),
    ('Nguyễn Văn A', 'Nguyễn Văn A'),    # giữ nguyên dấu — đây là tên in ra
])
def test_don_ten_khach(app, raw, expect):
    assert app._rc_clean_name(raw) == expect


def test_ten_rong(app):
    assert app._rc_clean_name(None) == ''
    assert app._rc_clean_name(float('nan')) == ''


# ── Mã đặt phòng (Conf#) ──────────────────────────────────────────────────
@pytest.mark.parametrize('raw,expect', [
    (2021649, '2021649'),
    (2021649.0, '2021649'),              # Excel đọc thành float
    ('2021649', '2021649'),
    ('ABC123', 'ABC123'),
])
def test_ma_dat_phong_khong_co_duoi_cham_khong(app, raw, expect):
    assert app._rc_conf(raw) == expect


def test_ma_dat_phong_rong(app):
    assert app._rc_conf(None) == ''


# ── Ngày trên regcard: luôn dd/mm/yyyy có đệm số 0 ────────────────────────
@pytest.mark.parametrize('raw,expect', [
    (datetime.date(2026, 8, 7), '07/08/2026'),
    (datetime.datetime(2026, 8, 7, 14, 0), '07/08/2026'),
    ('7/8/2026', '07/08/2026'),          # phải đệm 0
    ('07/08/2026', '07/08/2026'),
    ('7/8/26', '07/08/2026'),            # năm 2 chữ số
    (pd.Timestamp('2026-12-25'), '25/12/2026'),
])
def test_ngay_regcard(app, raw, expect):
    assert app._rc_date(raw) == expect


@pytest.mark.parametrize('raw,expect', [
    (datetime.date(2026, 7, 24), '24/07/2026'),
    ('24/07/2026', '24/07/2026'),
    ('5/6/2026', '05/06/2026'),
])
def test_ngay_regcard_doan(app, raw, expect):
    assert app._grp_date(raw) == expect


def test_ngay_rong(app):
    assert app._rc_date(None) == ''
    assert app._grp_date(None) == ''


def test_ngay_regcard_khong_hieu_kieu_my(app):
    """'7/8/2026' trên file Smile là 7 tháng 8, không phải 8 tháng 7."""
    assert app._rc_date('7/8/2026') == '07/08/2026'


# ── Số đêm lưu trú ────────────────────────────────────────────────────────
@pytest.mark.parametrize('den,di,dem', [
    ('01/09/2026', '05/09/2026', '4'),
    (datetime.date(2026, 9, 1), datetime.date(2026, 9, 5), '4'),
    ('01/09/2026', '02/09/2026', '1'),
    ('31/08/2026', '01/09/2026', '1'),          # qua tháng
    ('31/12/2026', '01/01/2027', '1'),          # qua năm
    ('01/09/2026', '01/09/2026', '0'),          # day-use
])
def test_so_dem(app, den, di, dem):
    """Trả về CHUỖI vì giá trị này điền thẳng vào ô trên regcard."""
    assert app._rc_nights(den, di) == dem


def test_so_dem_thieu_du_lieu_thi_khong_doan(app):
    assert app._rc_nights(None, '05/09/2026') == ''
    assert app._rc_nights('01/09/2026', None) == ''
    assert app._rc_nights('rác', '05/09/2026') == ''


def test_so_dem_di_truoc_den_thi_de_trong(app):
    """Số đêm âm là vô nghĩa — để trống cho lễ tân tự kiểm tra."""
    assert app._rc_nights('05/09/2026', '01/09/2026') == ''


# ── build_arr: file ARR từ file Arrival của Smile ─────────────────────────
def _book(rows, headers=None):
    import io
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(list(headers or ['Conf#', 'Folio#', 'Type', 'Last Name', 'First Name',
                               'Rm#', 'Arrival', 'Departure', 'Company', 'Notice']))
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_build_arr_upload_nham_file_thi_bao_ro_rang(app):
    """Lễ tân upload nhầm file phải nhận được thông báo nói rõ sai ở đâu và
    cần file nào, không phải một traceback khó hiểu."""
    sai = _book([('X',)], headers=['Tên'])
    with pytest.raises(ValueError, match='Conf#'):
        app.build_arr(sai)


def test_build_arr_file_khong_co_booking_thi_bao_ro_rang(app):
    with pytest.raises(ValueError, match='không có dữ liệu booking'):
        app.build_arr(_book([]))


# ── Regcard đoàn: che dữ liệu ví dụ in sẵn trên mẫu ──────────────────────
def test_moi_o_dien_deu_co_o_che(app):
    """Mẫu PDF có sẵn dữ liệu ví dụ in cứng (đoàn 21099 / VIRGO TRAVEL /
    25 phòng / 52 khách). Không che thì chữ mới vẽ ĐÈ lên chữ cũ và phiếu in
    ra chồng hai lớp — đây là lỗi đã xảy ra thật. Mọi ô có điền giá trị đều
    phải nằm trong một ô che."""
    import pandas as pd
    import datetime
    # vị trí (x, bottom) của từng ô điền, lấy đúng từ build_group_regcard
    vi_tri = {
        'Arrival': (155.6, 181.2), 'Departure': (436.9, 181.2),
        'Group Code': (48.2, 261.6), 'Group Name': (165.7, 263.9),
        'Travel Agent': (364.4, 263.9), 'No of rooms': (50.9, 330.6),
        'No of pax': (187.2, 326.2),
    }
    for ten, (x, bottom) in vi_tri.items():
        trong = [(a, t, b, d) for a, t, b, d in app._GROUP_BLANK
                 if a <= x + 1 <= b and t <= bottom - 3 <= d]
        assert trong, f'ô "{ten}" tại ({x},{bottom}) không có ô che nào phủ'


def test_o_che_khong_lan_vao_khung_bang(app):
    """Ô che lấn qua đường kẻ là xoá mất khung bảng trên phiếu in — đã dính
    đúng lỗi này khi kéo ô che tới x=570 trong khi mép phải bảng ở x=553.

    Chỉ xét đường kẻ CÓ GIAO nhau về chiều dọc với ô che: các hàng của bảng
    chia cột khác nhau nên đường ở hàng trên không liên quan hàng dưới."""
    for x0, top, bot_x1, bot in app._GROUP_BLANK:
        x1 = bot_x1
        cat = [(vx, vy0, vy1) for vx, vy0, vy1 in app._GROUP_TABLE_LINES
               if x0 < vx < x1 and vy0 < bot and top < vy1]
        assert not cat, f'ô che ({x0},{top},{x1},{bot}) cắt qua đường kẻ {cat}'


def test_regcard_doan_tao_duoc_trang_pdf(app):
    import pandas as pd
    import datetime
    grp = pd.DataFrame([
        {'Group': 880123, 'Name': 'SAMPLE TOUR GROUP', 'Company': 'SAMPLE TRAVEL CO.',
         'Rm': r, 'Arrival': datetime.date(2026, 9, 19),
         'Departure': datetime.date(2026, 9, 23), 'Adt': 2, 'Chl': 0, 'Enf': 0}
        for r in ('742', '744', '746')
    ])
    page = app.build_group_regcard(grp, app.load_group_template())
    assert page is not None


def test_regcard_doan_bo_phong_ao_khi_dem(app):
    """Phòng ảo 9xxx (posting master) không tính vào số phòng của đoàn."""
    import pandas as pd
    import datetime
    grp = pd.DataFrame([
        {'Group': 1, 'Name': 'G', 'Company': 'C', 'Rm': r,
         'Arrival': datetime.date(2026, 9, 19), 'Departure': datetime.date(2026, 9, 23),
         'Adt': 1, 'Chl': 0, 'Enf': 0}
        for r in ('742', '744', '9001')
    ])
    page = app.build_group_regcard(grp, app.load_group_template())
    assert page is not None      # 2 phòng thật, không phải 3
