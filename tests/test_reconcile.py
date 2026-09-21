"""Đối chiếu Smile với trang quản lý lưu trú, và đối chiếu hệ thống phòng.

Sai ở đây làm lễ tân đi tìm nhầm: báo khách "chưa đăng ký" trong khi đã
đăng ký rồi, hoặc bỏ sót khách thật sự chưa đăng ký.
"""
import datetime
import io

import pandas as pd
import pytest
from openpyxl import Workbook


def _xlsx(headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(list(headers))
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


HOM_NAY = pd.Timestamp('2026-09-21')


def _smile(rows):
    """File inhouse Smile tối giản."""
    return _xlsx(['Rm#', 'Last Name', 'First Name', 'NAT', 'Passport',
                  'Arrival', 'Departure'], rows)


# ── Chuẩn hoá số hộ chiếu: hai file gõ khác nhau vẫn phải khớp ───────────
@pytest.mark.parametrize('a,b', [
    ('C1234567', 'c1234567'),
    ('C1234567', ' C1234567 '),
    ('C1234567', 'C123 4567'),
    ('123456789', '123456789.0'),
])
def test_ho_chieu_gõ_khac_nhau_van_khop(app, a, b):
    assert app._norm_pp(a) == app._norm_pp(b)


def test_ho_chieu_khac_nhau_thi_khong_khop(app):
    assert app._norm_pp('C1234567') != app._norm_pp('C1234568')


# ── Chuẩn hoá số phòng ────────────────────────────────────────────────────
@pytest.mark.parametrize('a,b', [
    ('12A05', '12a05'), ('742', '742.0'), ('G1201', 'g1201'),
])
def test_phong_gõ_khac_nhau_van_la_mot(app, a, b):
    assert app._norm_room(a) == app._norm_room(b)


# ── Đối chiếu hệ thống phòng ──────────────────────────────────────────────
def test_phong_inhouse_thieu_trong_file_bi_bao_do(app):
    smile = _smile([
        ('742', 'SMITH', 'JOHN', 'USA', 'C1', '2026-09-19', '2026-09-25'),
        ('744', 'DOE', 'JANE', 'USA', 'C2', '2026-09-19', '2026-09-25'),
    ])
    he_thong = _xlsx(['Số phòng'], [('742',)])          # thiếu 744
    r = app.reconcile_rooms(smile, he_thong, HOM_NAY)
    assert '744' in r['room_chua']
    assert '742' not in r['room_chua']


def test_phong_thua_trong_file(app):
    smile = _smile([('742', 'SMITH', 'JOHN', 'USA', 'C1', '2026-09-19', '2026-09-25')])
    he_thong = _xlsx(['Số phòng'], [('742',), ('999',)])
    r = app.reconcile_rooms(smile, he_thong, HOM_NAY)
    assert '999' in r['room_thua']


def test_phong_trung_trong_file(app):
    smile = _smile([('742', 'SMITH', 'JOHN', 'USA', 'C1', '2026-09-19', '2026-09-25')])
    he_thong = _xlsx(['Số phòng'], [('742',), ('742',)])
    r = app.reconcile_rooms(smile, he_thong, HOM_NAY)
    assert '742' in r['sys_dup']


def test_phong_ao_9xxx_bi_loai_khoi_doi_chieu(app):
    """Phòng ảo 9000-9999 (posting master) không phải phòng ở thật."""
    smile = _smile([
        ('742', 'SMITH', 'JOHN', 'USA', 'C1', '2026-09-19', '2026-09-25'),
        ('9001', 'MASTER', 'POSTING', '', '', '2026-09-19', '2026-09-25'),
    ])
    he_thong = _xlsx(['Số phòng'], [('742',)])
    r = app.reconcile_rooms(smile, he_thong, HOM_NAY)
    assert '9001' not in r['room_chua']
    assert r['room_chua'] == []


def test_khach_den_hom_nay_khong_tinh_la_inhouse(app):
    """Khách Arrival = hôm nay chưa nhận phòng xong, chưa cần đăng ký."""
    smile = _smile([
        ('742', 'SMITH', 'JOHN', 'USA', 'C1', '2026-09-21', '2026-09-25'),
        ('744', 'DOE', 'JANE', 'USA', 'C2', '2026-09-19', '2026-09-25'),
    ])
    he_thong = _xlsx(['Số phòng'], [('538',)])   # phòng bất kỳ: file rỗng là lỗi riêng
    r = app.reconcile_rooms(smile, he_thong, HOM_NAY)
    assert '742' not in r['room_chua'], 'khách đến hôm nay chưa tính là inhouse'
    assert '744' in r['room_chua']


def test_khach_tra_phong_hom_nay_khong_tinh(app):
    smile = _smile([('742', 'SMITH', 'JOHN', 'USA', 'C1', '2026-09-19', '2026-09-21')])
    he_thong = _xlsx(['Số phòng'], [('538',)])
    r = app.reconcile_rooms(smile, he_thong, HOM_NAY)
    assert '742' not in r['room_chua']


def test_phong_hoa_thuong_van_khop(app):
    smile = _smile([('12a05', 'SMITH', 'JOHN', 'USA', 'C1', '2026-09-19', '2026-09-25')])
    he_thong = _xlsx(['Số phòng'], [('12A05',)])
    r = app.reconcile_rooms(smile, he_thong, HOM_NAY)
    assert r['room_chua'] == [] and r['room_thua'] == []


def test_o_tieu_de_trong_file_phong_khong_bi_coi_la_phong(app):
    smile = _smile([('742', 'SMITH', 'JOHN', 'USA', 'C1', '2026-09-19', '2026-09-25')])
    he_thong = _xlsx(['Số phòng'], [('742',), ('Room',), ('STT',)])
    r = app.reconcile_rooms(smile, he_thong, HOM_NAY)
    assert r['room_thua'] == []


def test_file_phong_rong_thi_bao_loi_ro_rang(app):
    smile = _smile([('742', 'SMITH', 'JOHN', 'USA', 'C1', '2026-09-19', '2026-09-25')])
    with pytest.raises(ValueError, match='không có dữ liệu'):
        app.reconcile_rooms(smile, _xlsx(['Số phòng'], []), HOM_NAY)


def test_file_smile_thieu_cot_phong_thi_bao_loi_ro_rang(app):
    sai = _xlsx(['Tên'], [('X',)])
    with pytest.raises(ValueError, match="Rm#"):
        app.reconcile_rooms(sai, _xlsx(['Số phòng'], [('742',)]), HOM_NAY)


def test_dem_khop_dung(app):
    smile = _smile([
        ('742', 'A', 'A', 'USA', 'C1', '2026-09-19', '2026-09-25'),
        ('744', 'B', 'B', 'USA', 'C2', '2026-09-19', '2026-09-25'),
    ])
    he_thong = _xlsx(['Số phòng'], [('742',), ('744',)])
    r = app.reconcile_rooms(smile, he_thong, HOM_NAY)
    assert r['room_match'] == 2
    assert r['room_chua'] == [] and r['room_thua'] == []
