"""Gom toàn bộ logic nghiệp vụ về một chỗ.

`app.py` và bộ test chỉ cần `from hotel_core import *`.

Thứ tự import đi từ tầng thấp lên cao và không có phụ thuộc vòng:

    common                       ← không phụ thuộc gì
    lunar · progress · db · assets
    lookups
    reconcile · regcard · xlsx_io · dk14 · docs · reports

Phải gộp bằng __all__ chứ không dùng `from hotel.x import *` trơn: rất
nhiều helper của app có tên bắt đầu bằng gạch dưới (_fmt_room, _norm_addr,
_dk_to_date...) mà `import *` mặc định bỏ qua chúng.
"""
from hotel import (
    common, lunar, progress, db, assets, lookups,
    reconcile, regcard, xlsx_io, dk14, docs, reports,
)

_MODULES = (
    common, lunar, progress, db, assets, lookups,
    reconcile, regcard, xlsx_io, dk14, docs, reports,
)

for _m in _MODULES:
    globals().update({_n: getattr(_m, _n) for _n in _m.__all__})

__all__ = sorted({_n for _m in _MODULES for _n in _m.__all__})
