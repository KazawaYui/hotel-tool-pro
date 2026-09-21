"""Mọi file tài nguyên phải nạp được từ bất kỳ thư mục chạy nào.

Loại lỗi này đã xảy ra thật: khi tách logic sang package hotel/, các hàm nạp
mẫu bám theo dirname(__file__) bắt đầu tìm vào hotel/ thay vì gốc dự án.
load_template() được sửa nhưng load_regcard_template() và
load_group_template() bị bỏ sót — công cụ Regcard + ARR hỏng hoàn toàn mà
bộ test cũ không hề báo, vì không test nào GỌI tới chúng.
"""
import os

import pytest


MAU = ['customer', 'kbtt', 'vnm', 'dk14']


@pytest.mark.parametrize('ten', MAU)
def test_nap_duoc_mau_xlsx(app, ten):
    data = app.load_template(ten)
    assert data[:2] == b'PK', f'mẫu {ten} không phải file xlsx hợp lệ'
    assert len(data) > 1000


def test_nap_duoc_mau_regcard(app):
    """Công cụ Regcard từng hỏng hoàn toàn vì đường dẫn này."""
    data = app.load_regcard_template()
    assert data[:4] == b'%PDF', 'mẫu regcard không phải PDF'
    assert len(data) > 1000


def test_nap_duoc_mau_regcard_doan(app):
    data = app.load_group_template()
    assert data[:4] == b'%PDF', 'mẫu regcard đoàn không phải PDF'
    assert len(data) > 1000


def test_nap_duoc_icon(app):
    assert app._load_app_icon() is not None


def test_moi_ham_nap_mau_dung_duoc_tu_thu_muc_khac(app, tmp_path, monkeypatch):
    """Chạy app từ thư mục khác vẫn phải nạp được tài nguyên — đường dẫn
    phải tính từ vị trí file mã nguồn, không phải thư mục hiện hành."""
    monkeypatch.chdir(tmp_path)
    for ten in MAU:
        app.load_template.clear()
        assert app.load_template(ten)[:2] == b'PK', ten
    app.load_regcard_template.clear()
    assert app.load_regcard_template()[:4] == b'%PDF'
    assert app.load_group_template()[:4] == b'%PDF'


def test_thu_muc_tien_do_nam_o_goc_du_an(app):
    """Tiến độ ca phải ghi vào data/ ở gốc — hotel/data/ vừa lệch chỗ so với
    dữ liệu đang có, vừa lọt khỏi .gitignore nên có thể bị commit nhầm."""
    import hotel.progress as progress
    goc = os.path.dirname(os.path.dirname(os.path.abspath(progress.__file__)))
    assert progress.DATA_DIR == os.path.join(goc, 'data'), progress.DATA_DIR
    assert not progress.DATA_DIR.replace(os.sep, '/').endswith('hotel/data')


def test_du_anh_nen_trong_static(app):
    """CSS trỏ tới app/static/bg_*.jpg — thiếu file thì nền biến mất lặng lẽ."""
    for key in ('dark', 'light', 'spring', 'autumn', 'winter', 'tet', 'noel'):
        p = os.path.join(app.STATIC_DIR, f'bg_{key}.jpg')
        assert os.path.exists(p), f'thiếu ảnh nền {p}'
        assert os.path.getsize(p) > 1000
