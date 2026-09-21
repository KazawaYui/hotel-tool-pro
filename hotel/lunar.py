"""Âm lịch Việt Nam (thuật toán Hồ Ngọc Đức) + mùa/khung giờ cho màn chào."""

import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from copy import copy
import xlrd, datetime, io, zipfile, base64, os, json
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.colors import white, black
import unicodedata as _ud, re as _re

from .common import now_vn, today_vn

# ── Chế độ giao diện (sáng/tối) — chọn tay hoặc tự động theo giờ Việt Nam.
# Mốc 6h-18h coi là ban ngày (giao diện sáng), ngoài khoảng đó là ban đêm
# (giao diện tối) — có thể chỉnh 2 số này nếu muốn đổi mốc giờ.
THEME_DAY_START_HOUR = 6
THEME_NIGHT_START_HOUR = 18

def _compute_effective_theme():
    mode = st.session_state.get('theme_mode', 'auto')
    if mode in ('light', 'dark'):
        return mode
    h = now_vn().hour
    return 'light' if THEME_DAY_START_HOUR <= h < THEME_NIGHT_START_HOUR else 'dark'

# ── Âm lịch Việt Nam ──────────────────────────────────────────────────────
# Dùng để biết hôm nay có phải Tết hay Trung Thu — 2 mốc này theo âm lịch nên
# KHÔNG thể suy ra từ ngày dương. Thuật toán thiên văn tiêu chuẩn (Hồ Ngọc Đức),
# múi giờ +7; tính thẳng nên không cần bảng tra ngày lễ phải cập nhật hằng năm.
def _jd_from_date(dd, mm, yy):
    a = (14 - mm) // 12
    y, m = yy + 4800 - a, mm + 12 * a - 3
    jd = dd + (153 * m + 2) // 5 + 365 * y + y // 4 - y // 100 + y // 400 - 32045
    if jd < 2299161:
        jd = dd + (153 * m + 2) // 5 + 365 * y + y // 4 - 32083
    return jd

def _new_moon_jd(k):
    import math
    T = k / 1236.85
    T2, dr = T * T, math.pi / 180
    T3 = T2 * T
    jd1 = 2415020.75933 + 29.53058868 * k + 0.0001178 * T2 - 0.000000155 * T3
    jd1 += 0.00033 * math.sin((166.56 + 132.87 * T - 0.009173 * T2) * dr)
    M = 359.2242 + 29.10535608 * k - 0.0000333 * T2 - 0.00000347 * T3
    Mpr = 306.0253 + 385.81691806 * k + 0.0107306 * T2 + 0.00001236 * T3
    F = 21.2964 + 390.67050646 * k - 0.0016528 * T2 - 0.00000239 * T3
    c1 = (0.1734 - 0.000393 * T) * math.sin(M * dr) + 0.0021 * math.sin(2 * dr * M)
    c1 -= 0.4068 * math.sin(Mpr * dr) + 0.0161 * math.sin(dr * 2 * Mpr)
    c1 -= 0.0004 * math.sin(dr * 3 * Mpr)
    c1 += 0.0104 * math.sin(dr * 2 * F) - 0.0051 * math.sin(dr * (M + Mpr))
    c1 -= 0.0074 * math.sin(dr * (M - Mpr)) + 0.0004 * math.sin(dr * (2 * F + M))
    c1 -= 0.0004 * math.sin(dr * (2 * F - M)) - 0.0006 * math.sin(dr * (2 * F + Mpr))
    c1 += 0.0010 * math.sin(dr * (2 * F - Mpr)) + 0.0005 * math.sin(dr * (2 * Mpr + M))
    if T < -11:
        dt = (0.001 + 0.000839 * T + 0.0002261 * T2
              - 0.00000845 * T3 - 0.000000081 * T * T3)
    else:
        dt = -0.000278 + 0.000265 * T + 0.000262 * T2
    return jd1 + c1 - dt

def _new_moon_day(k, tz=7):
    return int(_new_moon_jd(k) + 0.5 + tz / 24)

def _sun_longitude_deg6(jdn, tz=7):
    import math
    t = (jdn - 0.5 - tz / 24 - 2451545.0) / 36525
    t2, dr = t * t, math.pi / 180
    M = 357.52910 + 35999.05030 * t - 0.0001559 * t2 - 0.00000048 * t * t2
    L0 = 280.46645 + 36000.76983 * t + 0.0003032 * t2
    dl = (1.914600 - 0.004817 * t - 0.000014 * t2) * math.sin(dr * M)
    dl += (0.019993 - 0.000101 * t) * math.sin(dr * 2 * M) + 0.000290 * math.sin(dr * 3 * M)
    L = (L0 + dl) * dr
    L -= math.pi * 2 * int(L / (math.pi * 2))
    return int(L / math.pi * 6)

def _lunar_month11(yy, tz=7):
    off = _jd_from_date(31, 12, yy) - 2415021
    k = int(off / 29.530588853)
    nm = _new_moon_day(k, tz)
    return _new_moon_day(k - 1, tz) if _sun_longitude_deg6(nm, tz) >= 9 else nm

def _leap_month_offset(a11, tz=7):
    k = int((a11 - 2415021.076998695) / 29.530588853 + 0.5)
    i = 1
    arc = _sun_longitude_deg6(_new_moon_day(k + i, tz), tz)
    while True:
        last, i = arc, i + 1
        arc = _sun_longitude_deg6(_new_moon_day(k + i, tz), tz)
        if arc == last or i >= 14:
            break
    return i - 1

def solar_to_lunar(d, tz=7):
    """Ngày dương (datetime.date) -> (ngày, tháng, năm, có_nhuận) âm lịch."""
    day_number = _jd_from_date(d.day, d.month, d.year)
    k = int((day_number - 2415021.076998695) / 29.530588853)
    month_start = _new_moon_day(k + 1, tz)
    if month_start > day_number:
        month_start = _new_moon_day(k, tz)
    a11 = _lunar_month11(d.year, tz)
    b11 = a11
    if a11 >= month_start:
        lunar_year = d.year
        a11 = _lunar_month11(d.year - 1, tz)
    else:
        lunar_year = d.year + 1
        b11 = _lunar_month11(d.year + 1, tz)
    lunar_day = day_number - month_start + 1
    diff = int((month_start - a11) / 29)
    lunar_leap, lunar_month = 0, diff + 11
    if b11 - a11 > 365:
        leap_off = _leap_month_offset(a11, tz)
        if diff >= leap_off:
            lunar_month = diff + 10
            if diff == leap_off:
                lunar_leap = 1
    if lunar_month > 12:
        lunar_month -= 12
    if lunar_month >= 11 and diff < 4:
        lunar_year -= 1
    return lunar_day, lunar_month, lunar_year, lunar_leap

# ── Mùa / dịp lễ của màn chào ─────────────────────────────────────────────
# Ngày lễ (theo âm lịch hoặc ngày cố định) được ưu tiên ĐÈ LÊN mùa.
SEASON_LABEL = {
    'spring':   '🌸 TIẾT XUÂN',   'summer': '🌊 TIẾT HẠ',
    'autumn':   '🍂 TIẾT THU',    'winter': '🌧️ TIẾT ĐÔNG',
    'tet':      '🎋 TẾT NGUYÊN ĐÁN',
    'trungthu': '🏮 TRUNG THU',   'noel':   '🎄 GIÁNG SINH',
}

def _season_key(d=None):
    d = d or today_vn()
    if d.month == 12 and 20 <= d.day <= 26:
        return 'noel'
    try:
        ld, lm, _, leap = solar_to_lunar(d)
        if not leap:
            if lm == 1 and ld <= 7:          # mùng 1 -> mùng 7 Tết
                return 'tet'
            if lm == 12 and ld >= 28:        # 28, 29, 30 Tết
                return 'tet'
            if lm == 8 and 14 <= ld <= 16:   # rằm tháng 8
                return 'trungthu'
    except Exception:
        pass                                  # lỗi lịch không được làm hỏng màn chào
    return ('spring' if d.month <= 3 else 'summer' if d.month <= 6
            else 'autumn' if d.month <= 9 else 'winter')

def _hour_key(h=None):
    h = now_vn().hour if h is None else h
    return ('dawn' if 5 <= h < 8 else 'day' if 8 <= h < 16
            else 'dusk' if 16 <= h < 19 else 'night')


# Ten duoc module nay so huu — liet ke tuong minh de `import *` lay
# duoc ca helper co gach duoi dau.
__all__ = [
    'SEASON_LABEL', 'THEME_DAY_START_HOUR', 'THEME_NIGHT_START_HOUR',
    '_compute_effective_theme', '_hour_key', '_jd_from_date', '_leap_month_offset',
    '_lunar_month11', '_new_moon_day', '_new_moon_jd', '_season_key',
    '_sun_longitude_deg6', 'solar_to_lunar'
]
