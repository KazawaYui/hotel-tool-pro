"""Nạp phần LOGIC của app để test — không dựng UI Streamlit.

Thiết kế có chủ đích: thử import module thật trước, không có mới cắt app.py.
Nhờ vậy CÙNG một bộ test chạy được cả TRƯỚC và SAU khi tách module — chính
nó là bằng chứng việc tách module không làm đổi hành vi.
"""
import os
import sys
import types

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

_cached = None


def load_app():
    """Trả về namespace chứa toàn bộ hàm logic của app."""
    global _cached
    if _cached is not None:
        return _cached

    # ── Đường 1: đã tách module thì import thẳng (nhanh, đúng chuẩn) ──
    try:
        import hotel_core  # noqa: F401
        _cached = hotel_core
        return _cached
    except ImportError:
        pass

    # ── Đường 2: còn là file đơn — cắt phần trước mốc "# ── UI ──" rồi exec.
    # Phần sau mốc gọi st.sidebar/st.markdown ở mức module nên không exec được
    # ngoài runtime Streamlit.
    src_path = os.path.join(ROOT, 'app.py')
    with open(src_path, encoding='utf-8') as f:
        src = f.read()
    marker = '# ── UI ─'
    idx = src.index(marker)
    mod = types.ModuleType('app_logic')
    mod.__file__ = src_path
    cwd = os.getcwd()
    os.chdir(ROOT)          # load_template() đọc file .b64 theo đường dẫn tương đối
    try:
        exec(compile(src[:idx], 'app.py', 'exec'), mod.__dict__)
    finally:
        os.chdir(cwd)
    _cached = mod
    return _cached
