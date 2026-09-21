"""Hotel Tool Pro — giao diện Streamlit cho lễ tân Tân Hotel.

Toàn bộ logic nghiệp vụ nằm trong package `hotel/` (xem hotel_core.py).
File này chỉ còn phần dựng giao diện: CSS/JS, sidebar, và 7 màn công cụ.
"""
import streamlit as st
import pandas as pd
import datetime, io, zipfile, re as _re
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

from hotel_core import *   # noqa: F401,F403 — toàn bộ logic nghiệp vụ

# Phải là lệnh Streamlit ĐẦU TIÊN của app, trước mọi st.* khác.
st.set_page_config(page_title="Tân Hotel", page_icon=_load_app_icon(), layout="wide")

# ── UI ────────────────────────────────────────────────────────────────────

# Giao diện Neumorphism/soft UI — đơn sắc sáng, đổ bóng nổi-chìm
if not st.session_state.get("_app_scripts_injected"):
    st.session_state["_app_scripts_injected"] = True
    # Toàn bộ lớp giao diện Neumorphism (CSS + hoa anh đào nền + hiệu ứng chữ
    # + màn khởi động) gộp trong 1 lệnh st.iframe duy nhất, tiêm đúng 1 lần
    # mỗi phiên — mọi animation chỉ dùng transform/opacity (chạy trên GPU
    # compositor), không backdrop-filter, không chạy lại khi rerun.
    _boot_script = """
<script>
(function(){
    var doc = window.parent.document;
    // Đặt theme NGAY từ khối này để màn chào không bị chớp sáng khi đang ở
    // chế độ tối (khối gán data-theme mỗi lần rerun chạy sau khối này)
    doc.documentElement.setAttribute('data-theme', '__GREET_THEME__');
    if (doc.getElementById('main-app-style')) return;
    var css = doc.createElement('style');
    css.id = 'main-app-style';
    css.textContent = `
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header[data-testid="stHeader"] {background: transparent;}

    /* ── Hệ màu (light mặc định) — khai báo ở html để lớp popover/dropdown
       render NGOÀI .stApp cũng kế thừa được, và khai báo lại ở .stApp vì
       khai báo trực tiếp trên phần tử luôn thắng giá trị kế thừa. ── */
    html, .stApp {
        --ease: cubic-bezier(0.4, 0, 0.2, 1);
        --bg: #f6f7fa; --surf: #ffffff; --surf2: #fafbfc;
        --line: #e7e9ef; --line2: #f0f2f6; --chip: #f2f4f8;
        --tx: #0f172a; --tx2: #5b6478; --tx3: #98a1b3;
        --acc: #4f46e5; --acc2: #7c3aed;
        --ok: #0d9668; --warn: #d97706; --err: #dc2626;
        --sh: 0 1px 2px rgba(15,23,42,.05), 0 8px 24px rgba(15,23,42,.05);
        --sh-lg: 0 2px 6px rgba(15,23,42,.06), 0 14px 36px rgba(15,23,42,.09);
        --r-lg: 16px; --r-md: 11px; --r-sm: 9px; --r-pill: 999px;
    }
    html[data-theme="dark"], html[data-theme="dark"] .stApp {
        --bg: #12151d; --surf: #1a1f2b; --surf2: #161b25;
        --line: #262d3b; --line2: #212734; --chip: #242b39;
        --tx: #eef2fa; --tx2: #a3adc2; --tx3: #727d95;
        --acc: #8b93f8; --acc2: #a78bfa;
        --ok: #34d399; --warn: #fbbf24; --err: #f87171;
        --sh: 0 1px 2px rgba(0,0,0,.3), 0 10px 30px rgba(0,0,0,.32);
        --sh-lg: 0 2px 8px rgba(0,0,0,.35), 0 18px 44px rgba(0,0,0,.42);
    }
    .stApp {
        background-color: var(--bg);
        color: var(--tx);
        background-image: linear-gradient(rgba(246,247,250,.80), rgba(246,247,250,.93)),
                           url("__LIGHT_BG_DATA_URI__");
        background-size: cover; background-position: center center;
        background-repeat: no-repeat; background-attachment: fixed;
    }
    html[data-theme="dark"] .stApp {
        background-image: linear-gradient(rgba(18,21,29,.78), rgba(18,21,29,.92)),
                           url("__DARK_BG_DATA_URI__");
    }

    /* ── Sidebar ── */
    section[data-testid="stSidebar"] {
        background: var(--surf); border-right: 1px solid var(--line); box-shadow: none;
    }
    section[data-testid="stSidebar"] .stButton button {
        background: transparent; border: 1px solid transparent;
        border-radius: var(--r-sm);
        color: var(--tx2); font-weight: 560; text-align: left;
        justify-content: flex-start; padding: 0.5rem 0.6rem;
        box-shadow: none; transition: background 0.12s var(--ease), color 0.12s var(--ease);
    }
    section[data-testid="stSidebar"] .stButton button:hover {
        background: var(--chip); color: var(--tx);
    }
    section[data-testid="stSidebar"] .stButton button[kind="primary"] {
        background: linear-gradient(135deg, rgba(99,91,255,.12), rgba(167,139,250,.10)) !important;
        color: var(--tx) !important; font-weight: 700;
        border-color: rgba(120,110,250,.28) !important; box-shadow: none !important;
    }
    /* Icon mỗi mục nav nằm trong ô chip; mục đang mở thì chip đổi sang gradient */
    section[data-testid="stSidebar"] .stButton button [data-testid="stIconMaterial"] {
        background: var(--chip); border-radius: 7px; padding: 4px;
        width: 23px; height: 23px; display: inline-flex; align-items: center;
        justify-content: center; font-size: 15px !important; margin-right: 2px;
    }
    section[data-testid="stSidebar"] .stButton button[kind="primary"] [data-testid="stIconMaterial"] {
        background: linear-gradient(135deg, var(--acc), var(--acc2)); color: #fff;
    }
    /* Badge số bên phải mục nav (số điền động ở khối style tiêm mỗi lần rerun) */
    section[data-testid="stSidebar"] .stButton button::after {
        margin-left: auto; font-size: 0.62rem; font-weight: 800; color: var(--tx2);
        background: var(--chip); padding: 1px 7px; border-radius: var(--r-pill);
    }
    /* Khung nền cho cụm nút gạt giao diện (hàng cột duy nhất trong sidebar) */
    section[data-testid="stSidebar"] div[data-testid="stHorizontalBlock"] {
        background: var(--chip); border-radius: 10px; padding: 3px; gap: 3px !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stHorizontalBlock"] .stButton button {
        justify-content: center; text-align: center; padding: 0.36rem 0.2rem;
        font-size: 0.72rem; font-weight: 680; border: none;
    }
    section[data-testid="stSidebar"] div[data-testid="stHorizontalBlock"] .stButton button[kind="primary"] {
        background: var(--surf) !important; color: var(--tx) !important;
        border-color: transparent !important; box-shadow: 0 1px 3px rgba(15,23,42,.16) !important;
    }
    .sb-brand {padding: 0.1rem 0.3rem 0.9rem; border-bottom: 1px solid var(--line2); margin-bottom: 0.5rem;}
    .sb-mascot {background: linear-gradient(140deg,#ff8fb1,#c084fc); border-radius: 9px;
        padding: 4px; box-shadow: 0 4px 12px rgba(200,120,220,.35);}
    .sb-brand-title {color: var(--tx); font-weight: 730; font-size: 0.92rem; letter-spacing: -.015em;}
    .sb-mascot {display:inline-block; width:22px; height:22px; vertical-align:-5px; margin-right:5px;
        animation: flowerSway 3.4s ease-in-out infinite;}
    .sb-mascot svg {width:100%; height:100%;}
    .sb-brand-sub {color: var(--tx3); font-size: 0.7rem; margin-top: 1px;}
    .sb-section {color: var(--tx3); font-size: 0.6rem; font-weight: 800;
        text-transform: uppercase; letter-spacing: 0.1em; padding: 0.8rem 0.55rem 0.3rem;}
    .sb-user {display:flex; align-items:center; gap:9px; padding-top:0.8rem;
        margin-top:0.5rem; border-top: 1px solid var(--line2);}
    .sb-av {width:28px; height:28px; border-radius:50%; flex-shrink:0; color:#fff;
        background: linear-gradient(135deg,#818cf8,#c084fc);
        display:flex; align-items:center; justify-content:center; font-size:0.72rem; font-weight:760;}
    .sb-un {font-size:0.76rem; font-weight:680; color: var(--tx);}
    .sb-ur {font-size:0.66rem; color: var(--tx3);}
    .sb-status {display:flex; align-items:center; gap:7px; color: var(--tx3);
        font-size:0.68rem; padding-top:0.55rem;}
    .sb-dot {position:relative; width:7px; height:7px; border-radius:50%; background:#34d399; flex-shrink:0;}
    .sb-dot::after {content:""; position:absolute; inset:0; border-radius:50%; background:#34d399;
        animation: pulse 2.2s var(--ease) infinite;}
    @keyframes pulse {
        0%   {transform: scale(1);   opacity: 0.5;}
        70%  {transform: scale(2.6); opacity: 0;}
        100% {transform: scale(2.6); opacity: 0;}
    }
    @keyframes flowerSway {
        0%, 100% {transform: rotate(-8deg);}
        50%      {transform: rotate(8deg);}
    }

    /* ── Nội dung hiện dần so le sau khi màn chào tan ── */
    @keyframes tanReveal {
        from {opacity: 0; transform: translateY(10px);}
        to   {opacity: 1; transform: none;}
    }
    .tan-rv {opacity: 0; animation: tanReveal 0.42s var(--ease) forwards;}

    /* ── Chuyển màn mượt khi đổi công cụ (sidebar) — TRƯỚC đây đổi công cụ là
       cắt cứng, không có hiệu ứng gì, khác hẳn màn chào mượt mà lúc mở web.
       Lớp .tan-page-in được JS thêm vào mỗi khi menu thực sự đổi (xem khối
       script ở cuối app.py) — chỉ 1 lần mờ+trượt nhẹ, không so le từng khối
       như màn chào vì lặp lại mỗi cú click sẽ thấy chậm/rườm. */
    @keyframes tanPageIn {
        from {opacity: 0; transform: translateY(7px);}
        to   {opacity: 1; transform: none;}
    }
    .tan-page-in {animation: tanPageIn 0.26s var(--ease);}
    @media (prefers-reduced-motion: reduce) {
        .tan-page-in {animation: none !important;}
    }

    /* ── Top bar: breadcrumb + trạng thái ── */
    .tan-topbar {
        display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
        font-size: 0.82rem; color: var(--tx); padding: 0.15rem 0 0.15rem;
    }
    /* Thanh top bar: nền liền khối + viền dưới, tràn ra sát 2 mép khung nội dung */
    /* Khung nội dung có padding 96px trên / 80px hai bên — kéo âm đúng bằng đó
       để thanh chạy sát mép, cộng width bù lại phần margin âm hai bên. */
    .st-key-tan_topbar {
        background: var(--surf); border-bottom: 1px solid var(--line);
        margin: -4.6rem -1.6rem 1.1rem !important; width: calc(100% + 3.2rem) !important;
        max-width: none !important; padding: 0.5rem 1.6rem !important;
    }
    /* Nút "Xuất báo cáo" trên thanh là nút phụ (nút chính là Xử lý hàng ngày) */
    .st-key-tb_report_dl button, .st-key-tb_report_off button {
        background: var(--surf) !important; color: var(--tx) !important;
        border: 1px solid var(--line) !important; box-shadow: none !important;
        font-weight: 640;
    }
    /* Thanh progress mảnh trong thẻ số liệu */
    .kpi-bar {height: 3px; border-radius: 3px; background: var(--chip); margin-top: .7rem; overflow: hidden;}
    .kpi-bar i {display: block; height: 100%; border-radius: 3px; background: var(--acc);}
    .kpi-bar i.err {background: var(--err);}
    .tan-crumb {color: var(--tx3);}
    .tan-sep {color: var(--tx3); opacity: .6;}
    .tan-topbar-r {margin-left: auto; display: flex; align-items: center; gap: 7px;}
    .tan-h1 {font-size: 1.35rem; font-weight: 790; letter-spacing: -.03em; color: var(--tx);}
    .tan-h1sub {font-size: 0.78rem; color: var(--tx3); margin: 2px 0 0.7rem;}

    /* ── Thẻ số liệu nhỏ (bento) ── */
    .kpi-lab {font-size: 0.76rem; color: var(--tx2); font-weight: 640;
        display: flex; align-items: center; gap: 8px;}
    .kpi-ic {width: 26px; height: 26px; border-radius: 8px; display: inline-flex;
        align-items: center; justify-content: center; font-size: 0.78rem; background: var(--chip);}
    .kpi-ic.err {background: rgba(239,68,68,.14);}
    .kpi-ic.ok {background: rgba(16,185,129,.14);}
    .kpi-ic.acc {background: rgba(99,91,240,.14);}
    .kpi-val {font-size: 1.85rem; font-weight: 800; letter-spacing: -.035em;
        line-height: 1; margin-top: 1.5rem; color: var(--tx); font-variant-numeric: tabular-nums;}
    .kpi-val.err {color: var(--err);}
    .kpi-sub {font-size: 0.7rem; color: var(--tx3); font-weight: 620; margin-top: .45rem;}

    /* ── Panel (thẻ có tiêu đề + danh sách dòng) ── */
    .pan-h {display: flex; align-items: center; gap: 8px; padding-bottom: .55rem;
        margin-bottom: .5rem; border-bottom: 1px solid var(--line2);}
    .pan-t {font-size: 0.85rem; font-weight: 730; letter-spacing: -.015em; color: var(--tx);}
    .pan-row {display: flex; align-items: center; gap: 8px; flex-wrap: wrap; font-size: 0.82rem;}
    .pan-task {display: flex; align-items: center; gap: 9px; padding: .62rem .1rem;
        font-size: 0.82rem; border-bottom: 1px solid var(--line2);}
    .pan-task:last-of-type {border-bottom: 0;}
    .pan-time {margin-left: auto; font-size: 0.7rem; color: var(--tx3); font-weight: 650;
        font-variant-numeric: tabular-nums;}
    .pan-rt {color: var(--tx); font-weight: 560;}
    .pan-mut {font-size: 0.7rem; color: var(--tx3); font-weight: 620;}
    .pan-line {display: flex; align-items: center; justify-content: space-between;
        gap: 10px; padding: .48rem 0; font-size: 0.8rem; color: var(--tx2);
        border-bottom: 1px solid var(--line2);}
    .pan-line:last-child {border-bottom: 0;}
    .pan-v {font-weight: 750; color: var(--tx); font-variant-numeric: tabular-nums;}
    .pan-empty {font-size: 0.78rem; color: var(--tx3); line-height: 1.55; padding: .3rem 0 .7rem;}
    .pan-foot {margin-top: auto; padding-top: .35rem;}
    /* Hộp cảnh báo trong panel (đúng kiểu ô nhắc việc của bản dựng) */
    .pan-al {display: flex; gap: 9px; padding: .6rem .65rem; border-radius: 11px;
        margin-bottom: .45rem; border: 1px solid;}
    .pan-al.r {background: rgba(239,68,68,.10); border-color: rgba(239,68,68,.26);}
    .pan-al.a {background: rgba(245,158,11,.10); border-color: rgba(245,158,11,.26);}
    .pan-al-ic {font-size: 0.85rem; line-height: 1.2;}
    .pan-al-t {font-size: 0.78rem; font-weight: 680; color: var(--tx); line-height: 1.35;}
    .pan-al-m {font-size: 0.68rem; color: var(--tx3); margin-top: 2px; line-height: 1.35;}

    /* ── Nhãn mục ── */
    .section-label {
        font-size: 0.63rem; font-weight: 800; color: var(--tx3);
        text-transform: uppercase; letter-spacing: 0.09em; margin-bottom: 0.6rem;
    }

    /* ── Thẻ nội dung (st.container(border=True)) ──
       Bản Streamlit này KHÔNG còn testid "stVerticalBlockBorderWrapper";
       khối có viền là stVerticalBlock mang thêm thuộc tính overflow (các khối
       thường không có). Giữ cả selector cũ cho bản Streamlit đời trước. */
    div[data-testid="stVerticalBlock"][overflow],
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background: var(--surf); border: 1px solid var(--line) !important;
        border-radius: 14px !important; box-shadow: var(--sh);
        transition: box-shadow 0.15s var(--ease);
    }
    div[data-testid="stVerticalBlock"][overflow]:hover,
    div[data-testid="stVerticalBlockBorderWrapper"]:hover {box-shadow: var(--sh-lg);}

    /* ── Thẻ HERO (ảnh mèo làm nền, phủ tối dần để chữ luôn nổi) ── */
    .tan-hero {
        position: relative; overflow: hidden; color: #fff;
        border-radius: var(--r-lg); padding: 1.15rem 1.35rem;
        min-height: 200px; display: flex; flex-direction: column; justify-content: center;
        background-image: linear-gradient(100deg, rgba(22,18,48,.95) 0%, rgba(34,26,66,.74) 46%, rgba(48,36,84,.26) 100%),
                           url("__LIGHT_BG_DATA_URI__");
        background-size: cover, cover; background-position: center, center right;
        box-shadow: 0 12px 32px rgba(30,25,70,.28);
    }
    html[data-theme="dark"] .tan-hero {
        background-image: linear-gradient(100deg, rgba(16,14,32,.95) 0%, rgba(26,22,50,.74) 46%, rgba(40,32,72,.24) 100%),
                           url("__DARK_BG_DATA_URI__");
        box-shadow: 0 14px 36px rgba(0,0,0,.45);
    }
    .tan-hero-lab {font-size: 0.74rem; font-weight: 640; opacity: .88;}
    .tan-hero-val {font-size: 2.55rem; font-weight: 820; letter-spacing: -.04em; line-height: 1; margin-top: .35rem;}
    .tan-hero-sub {font-size: 0.75rem; opacity: .85; font-weight: 600; margin-top: .45rem;}
    .tan-hero-split {display: flex; gap: 1.6rem; margin-top: .9rem; padding-top: .75rem;
        border-top: 1px solid rgba(255,255,255,.22); flex-wrap: wrap;}
    .tan-hero-k {font-size: 0.68rem; opacity: .82; font-weight: 620;}
    .tan-hero-v {font-size: 1.1rem; font-weight: 790; margin-top: 1px; letter-spacing: -.02em;}

    /* ── Chip trạng thái ── */
    .tan-chip {display:inline-block; font-size:0.64rem; font-weight:740; padding:2px 8px;
        border-radius: var(--r-pill); background: var(--chip); color: var(--tx2);}
    .tan-chip.ok {background: rgba(16,185,129,.14); color: var(--ok);}
    .tan-chip.warn {background: rgba(245,158,11,.15); color: var(--warn);}
    .tan-chip.err {background: rgba(239,68,68,.13); color: var(--err);}

    /* ── Ô nhập ── */
    [data-testid="stTextInputRootElement"],
    [data-testid="stTextAreaRootElement"],
    [data-testid="stNumberInputContainer"] {
        background: var(--surf2) !important; border-radius: var(--r-md) !important;
    }
    .stTextInput input, .stNumberInput input, .stTextArea textarea {
        background: var(--surf2) !important; border: 1px solid var(--line) !important;
        border-radius: var(--r-md) !important; color: var(--tx) !important;
        box-shadow: none !important; transition: border-color 0.12s var(--ease), box-shadow 0.12s var(--ease);
    }
    .stTextInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus {
        border-color: var(--acc) !important;
        box-shadow: 0 0 0 3px rgba(99,91,240,.16) !important;
    }
    div[data-testid="stNumberInput"] button {border: none; background: transparent; color: var(--tx2) !important;}
    [data-testid="stNumberInputStepDown"], [data-testid="stNumberInputStepUp"],
    button[aria-label="Open"], button[aria-label="Show password"], button[aria-label="Hide password"] {
        color: var(--tx2) !important;
    }
    div[data-testid="stSelectbox"] [role="group"] {
        background: var(--surf2) !important; border-radius: var(--r-md) !important;
    }
    div[data-testid="stSelectbox"] input[role="combobox"] {color: var(--tx) !important;}
    div:has(> [role="listbox"]) {background: var(--surf) !important;}
    [role="option"] {color: var(--tx) !important;}
    [role="option"][aria-selected="true"], [role="option"]:hover {background: var(--chip) !important;}
    [data-testid="stWidgetLabel"] p {color: var(--tx2) !important; font-weight: 600; font-size: 0.8rem;}

    /* ── st.date_input: khung ô nhập của BaseWeb hardcode nền sáng
       (rgb(237,240,245)), không kế thừa theme — vẫn trắng dù đã bật tối. ── */
    [data-testid="stDateInput"] [data-baseweb="input"] {
        background: var(--surf2) !important; border: 1px solid var(--line) !important;
        border-radius: var(--r-md) !important;
    }
    [data-testid="stDateInputField"] {
        background: transparent !important; color: var(--tx) !important;
        caret-color: var(--tx) !important;
    }
    [data-testid="stDateInputField"]::placeholder {color: var(--tx3) !important;}
    /* Lịch chọn ngày hiện trong portal NGOÀI .stApp — không kế thừa biến màu
       của .stApp nên khai báo riêng ở đây (giống khối html[data-theme] khác).
       Khung popover bọc ngoài lịch cũng hardcode nền sáng — sửa luôn. */
    [data-baseweb="popover"] {background: var(--surf) !important;}
    div[data-baseweb="calendar"] {background: var(--surf) !important; color: var(--tx) !important;}
    /* Thanh "‹ Tháng Năm ›" và hàng tên thứ (Su Mo Tu...) là 2 div thường,
       không có data-baseweb riêng — cũng bị hardcode nền sáng như trên. */
    div[data-baseweb="calendar"] div:has(> button[aria-label="Previous month."]),
    div[data-baseweb="calendar"] div[role="presentation"] {
        background: var(--surf) !important; color: var(--tx) !important;
    }
    div[data-baseweb="calendar"] [role="presentation"] {color: var(--tx2) !important;}
    div[data-baseweb="calendar"] button {color: var(--tx) !important; background: transparent !important;}
    div[data-baseweb="calendar"] button[aria-disabled="true"] {color: var(--tx3) !important;}
    /* Ô trống đầu/cuối lưới (ngày thuộc tháng trước/sau) là gridcell RỖNG
       (không có div con chứa số) — tự vẽ nền bằng ::after màu sáng hardcode.
       Chỉ nhắm đúng ô rỗng (:empty) — ô có số (kể cả ô đang chọn, cũng vẽ
       vòng tròn tô màu bằng ::after) phải giữ nguyên, không bị vạ lây. */
    div[data-baseweb="calendar"] [role="gridcell"]:empty::after {background: transparent !important;}

    /* ── st.code(): nền pre/code hardcode sáng, không kế thừa theme ── */
    [data-testid="stCode"] pre {background: var(--surf2) !important; border: 1px solid var(--line) !important;}
    [data-testid="stCode"] code {color: var(--tx) !important;}
    /* Nút "Copy to clipboard" nổi góc trên code block: khung bọc nút cũng
       hardcode nền sáng (nút tự thân trong suốt nên trước đó không thấy). */
    div:has(> [data-testid="stElementToolbarButton"]) {background: var(--surf2) !important;}
    [data-testid="stElementToolbarButton"] svg {color: var(--tx2) !important;}
    div[data-baseweb="calendar"] div[aria-selected="true"] button {
        background: var(--acc) !important; color: #fff !important;
    }

    /* ── Tải file ── */
    div[data-testid="stFileUploader"] {
        background: var(--surf2); border: 1px dashed var(--line); border-radius: var(--r-md);
        padding: 0.5rem; box-shadow: none;
        transition: border-color 0.12s var(--ease);
    }
    div[data-testid="stFileUploader"]:hover {border-color: var(--acc);}
    div[data-testid="stFileUploader"] section {background: transparent; border: none;}
    div[data-testid="stFileUploader"] button {
        background: var(--surf) !important; color: var(--tx) !important;
        border: 1px solid var(--line) !important; box-shadow: none !important;
        border-radius: var(--r-sm) !important; font-weight: 640;
    }

    /* ── Nút ── */
    div[data-testid="stMainBlockContainer"] .stButton button, .stDownloadButton button {
        border-radius: var(--r-md); font-weight: 660; border: 1px solid var(--line);
        background: var(--surf); color: var(--tx);
        box-shadow: none;
        transition: background 0.12s var(--ease), border-color 0.12s var(--ease), transform 0.1s var(--ease);
    }
    div[data-testid="stMainBlockContainer"] .stButton button:hover, .stDownloadButton button:hover {
        background: var(--chip); border-color: var(--line);
    }
    div[data-testid="stMainBlockContainer"] .stButton button:active, .stDownloadButton button:active {
        transform: scale(0.985);
    }
    div[data-testid="stMainBlockContainer"] .stButton button[kind="primary"], .stDownloadButton button {
        background: linear-gradient(135deg, var(--acc), var(--acc2));
        color: #ffffff; border: none;
        box-shadow: 0 6px 18px rgba(99,91,240,.32);
    }
    div[data-testid="stMainBlockContainer"] .stButton button[kind="primary"]:hover, .stDownloadButton button:hover {
        filter: brightness(1.06); background: linear-gradient(135deg, var(--acc), var(--acc2));
    }

    /* ── Metric ── */
    div[data-testid="stMetric"] {
        background: var(--surf); border: 1px solid var(--line); border-radius: 13px;
        padding: 0.85rem 0.95rem 0.75rem; box-shadow: var(--sh);
        transition: box-shadow 0.15s var(--ease);
    }
    div[data-testid="stMetric"]:hover {box-shadow: var(--sh-lg);}
    div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] * {
        font-weight: 780 !important; color: var(--tx) !important; letter-spacing: -.03em;
        font-variant-numeric: tabular-nums;
    }
    div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] * {
        font-size: 0.76rem !important; color: var(--tx2) !important; font-weight: 620 !important;
    }

    /* ── Cảnh báo ── */
    div[data-testid="stAlert"] {border-radius: var(--r-md); border: 1px solid transparent;}

    /* ── Form / expander / bảng ── */
    div[data-testid="stForm"] {
        background: var(--surf); border: 1px solid var(--line);
        border-radius: 14px; padding: 1rem 1.1rem; box-shadow: var(--sh);
    }
    div[data-testid="stExpander"] {background: var(--surf); border-radius: var(--r-md);}
    /* Streamlit tô nền xám sáng cho summary khi expander ĐANG MỞ — ở chế độ tối
       thành chữ trắng trên nền sáng, không đọc được. Ép trong suốt cả 2 trạng thái. */
    div[data-testid="stExpander"] summary,
    div[data-testid="stExpander"] details[open] > summary {
        color: var(--tx) !important; background: transparent !important;
    }
    div[data-testid="stDataFrame"] {
        border-radius: var(--r-md); overflow: hidden; border: 1px solid var(--line);
    }

    /* ── Chữ nền tối: ép màu cho phần Streamlit tự đặt màu tĩnh ── */
    html[data-theme="dark"] .stApp p, html[data-theme="dark"] .stApp span,
    html[data-theme="dark"] .stApp label, html[data-theme="dark"] .stApp li,
    html[data-theme="dark"] .stApp h1, html[data-theme="dark"] .stApp h2,
    html[data-theme="dark"] .stApp h3, html[data-theme="dark"] .stApp h4,
    html[data-theme="dark"] [data-testid="stMarkdownContainer"] {color: var(--tx);}
    html[data-theme="dark"] [data-testid="stCaptionContainer"],
    html[data-theme="dark"] [data-testid="stCaptionContainer"] * {color: var(--tx3) !important;}
    html[data-theme="dark"] [data-testid="stCheckbox"] label,
    html[data-theme="dark"] [data-testid="stRadio"] label {color: var(--tx) !important;}
    /* Thẻ hero luôn nền tối nên chữ bên trong luôn trắng ở cả 2 chế độ */
    .tan-hero, .tan-hero * {color: #fff !important;}

    /* ── Bố cục ── */
    div[data-testid="stMainBlockContainer"] {max-width: 1320px; padding: 4.6rem 1.6rem 2.5rem;}
    section[data-testid="stSidebar"] {width: 252px !important; min-width: 252px !important;}
    /* Khe giữa các thẻ = 13px như bản dựng (trừ cụm nút giao diện ở sidebar) */
    div[data-testid="stMain"] div[data-testid="stHorizontalBlock"] {gap: 0.82rem;}
    /* Padding trong thẻ có viền: gọn lại cho khớp bản dựng */
    div[data-testid="stVerticalBlock"][overflow] {
        padding: 0.85rem 1rem !important; display: flex; flex-direction: column;
    }
    /* Link "Xem tất cả" ở đầu panel: nút nhưng nhìn như link, đúng bản dựng */
    .st-key-dash_seeall button {
        background: transparent !important; border: none !important; box-shadow: none !important;
        color: var(--acc) !important; font-size: 0.72rem !important; font-weight: 700 !important;
        padding: 0 !important; justify-content: flex-end !important; min-height: 0 !important;
    }
    .st-key-dash_seeall button:hover {background: transparent !important; text-decoration: underline;}
    /* Phần tử cuối trong 2 panel này ghim xuống đáy thẻ như bản dựng */
    .st-key-dash_tasks > div[data-testid="stElementContainer"]:last-child,
    .st-key-dash_ho > div[data-testid="stElementContainer"]:last-child {margin-top: auto;}
    div[data-testid="stElementContainer"]:has(iframe[height="1"]) {display: none;}
    div[data-testid="stMainBlockContainer"] hr {border-color: var(--line);}

    .stButton button:focus-visible, .stDownloadButton button:focus-visible {
        outline: 3px solid rgba(99,91,240,.4); outline-offset: 2px;
    }

    @media (max-width: 768px) {
        div[data-testid="stMainBlockContainer"] {padding-left: 1rem; padding-right: 1rem;}
        .st-key-tan_topbar {margin: -2rem -1rem 0.9rem; width: calc(100% + 2rem); padding: 0.45rem 1rem;}
        .tan-hero-val {font-size: 2rem;}
        .tan-hero-split {gap: 1.1rem;}
        div[data-testid="stMetric"] {padding: 0.6rem 0.7rem 0.55rem;}
    }

    @media (prefers-reduced-motion: reduce) {
        .sb-mascot, .sb-dot::after, #bg-sakura-layer .petal,
        #boot-splash .bs-pt, #boot-splash .bs-bg, #boot-splash .bs-moon,
        #boot-splash .bs-blink {
            animation: none !important;
        }
        .tan-rv, #boot-splash .bs-card, #boot-splash .bs-corner,
        #boot-splash .bs-strip, #boot-splash .bs-enter, #boot-splash .bs-hint2 {
            animation: none !important; opacity: 1 !important; transform: none !important;
        }
        #boot-splash .bs-card { transform: translateY(-50%) !important; }
        .stApp *, #boot-splash {transition-duration: 0.01ms !important;}
    }
    `;
    doc.head.appendChild(css);

    // ── Hoa anh đào rơi liên tục ở nền — thuần transform/opacity, vô hại hiệu năng ──
    if (!doc.getElementById('bg-sakura-layer')) {
        var css2 = doc.createElement('style');
        css2.id = 'bg-sakura-style';
        css2.textContent = `
          #bg-sakura-layer { position: fixed; inset: 0; pointer-events: none; overflow: hidden; z-index: 0; }
          #bg-sakura-layer .petal {
              position: absolute; top: -20px; opacity: 0.55; will-change: transform;
              animation-name: bgSakuraFall; animation-timing-function: linear; animation-iteration-count: infinite;
          }
          @keyframes bgSakuraFall {
              0%   { transform: translate(0,0) rotate(0deg); }
              100% { transform: translate(var(--drift), 112vh) rotate(360deg); }
          }
        `;
        doc.head.appendChild(css2);
        var layer = doc.createElement('div');
        layer.id = 'bg-sakura-layer';
        var colors = ['#f6a8c9', '#f293bc', '#f9c1d9'];
        for (var i = 0; i < 10; i++) {
            var p = doc.createElement('div');
            p.className = 'petal';
            var size = 8 + Math.random()*6;
            var dur = 11 + Math.random()*8;
            p.style.left = (Math.random()*100) + 'vw';
            p.style.width = size + 'px';
            p.style.height = size + 'px';
            p.style.borderRadius = '0 60% 0 60%';
            p.style.background = 'radial-gradient(circle at 30% 30%, #fff, ' + colors[i % 3] + ' 70%)';
            p.style.setProperty('--drift', (Math.random()*160 - 80) + 'px');
            p.style.animationDuration = dur + 's';
            p.style.animationDelay = (-Math.random()*dur) + 's';
            layer.appendChild(p);
        }
        doc.body.insertBefore(layer, doc.body.firstChild);
    }

    // ── Màn chào ca trực (bản phối B+C) ──────────────────────────────────
    // Ảnh nền tràn màn hình đổi theo MÙA và DỊP LỄ, ám sắc trời đổi theo GIỜ,
    // hạt rơi theo mùa; lời chào nằm trong thẻ kính bên trái, đồng hồ góc phải,
    // thanh số liệu ở đáy. Sau 2 giây thì DỪNG chờ người dùng bấm Enter (hoặc
    // chạm màn hình) mới vào app — giống màn khoá Windows. Chỉ chào lần đầu
    // mỗi phiên tab; refresh trong cùng tab không phải chào lại.
    var _bsSeen = false;
    try { _bsSeen = window.parent.sessionStorage.getItem('tanBootSplashSeen') === '1'; } catch (e) {}
    if (!_bsSeen && !doc.getElementById('boot-splash')) {
        try { window.parent.sessionStorage.setItem('tanBootSplashSeen', '1'); } catch (e) {}
        var css3 = doc.createElement('style');
        css3.id = 'boot-splash-style';
        css3.textContent = `
          #boot-splash {
            position: fixed; inset: 0; z-index: 999999; overflow: hidden;
            background: #07090f; cursor: pointer; user-select: none;
            transition: opacity 0.5s ease;
            font-family: -apple-system, "Segoe UI", Roboto, "Helvetica Neue", sans-serif;
          }
          #boot-splash.bs-hide { opacity: 0; pointer-events: none; }

          /* Lớp 1 — ảnh nền theo mùa, phóng rất chậm cho đỡ tĩnh */
          #boot-splash .bs-bg {
            position: absolute; inset: 0; background-size: cover; background-position: center 55%;
            background-image: url('__GREET_PHOTO__');
            animation: bsKen 34s ease-in-out infinite alternate;
          }
          @keyframes bsKen {
            from { transform: scale(1.02) translate(0.6%, 0); }
            to   { transform: scale(1.09) translate(-1.2%, -1.2%); }
          }
          /* Lớp 2 — ám sắc trời theo giờ (soft-light nên ảnh giữ nguyên chi tiết) */
          #boot-splash .bs-hour { position: absolute; inset: 0; mix-blend-mode: soft-light; }
          #boot-splash[data-h="dawn"]  .bs-hour { background: linear-gradient(180deg,#2a3d6b,#ff9d5c 68%,#ffd9a8); }
          #boot-splash[data-h="day"]   .bs-hour { background: linear-gradient(180deg,#4e97dc,#cfe6f7); }
          #boot-splash[data-h="dusk"]  .bs-hour { background: linear-gradient(180deg,#3a1c58,#e0724a 62%,#ffb066); }
          #boot-splash[data-h="night"] .bs-hour { background: linear-gradient(180deg,#070c1e,#1c2352 68%,#2c2f60); }
          /* Lớp 3 — ám sắc mùa */
          #boot-splash .bs-season { position: absolute; inset: 0; mix-blend-mode: overlay; opacity: .5; }
          #boot-splash[data-s="spring"]   .bs-season { background: linear-gradient(120deg,#ffb3d9,#c9a8ff); }
          #boot-splash[data-s="summer"]   .bs-season { background: linear-gradient(120deg,#7fd4ff,#b9f0e0); }
          #boot-splash[data-s="autumn"]   .bs-season { background: linear-gradient(120deg,#e09a4a,#c96a2b); }
          #boot-splash[data-s="winter"]   .bs-season { background: linear-gradient(120deg,#8fb8e0,#cfe2f5); }
          #boot-splash[data-s="tet"]      .bs-season { background: linear-gradient(120deg,#e03a3a,#ffc93a); opacity: .34; }
          #boot-splash[data-s="trungthu"] .bs-season { background: linear-gradient(120deg,#ffb347,#ff7a45); opacity: .40; }
          #boot-splash[data-s="noel"]     .bs-season { background: linear-gradient(120deg,#4a8fd4,#e05a5a); opacity: .42; }
          /* Lớp 4 — scrim giữ chữ luôn đọc được trên mọi ảnh */
          #boot-splash .bs-scrim {
            position: absolute; inset: 0;
            background: linear-gradient(96deg,rgba(6,8,14,.76) 0%,rgba(6,8,14,.40) 40%,rgba(6,8,14,.06) 66%,rgba(6,8,14,.34) 100%),
                        linear-gradient(0deg,rgba(6,8,14,.88) 0%,rgba(6,8,14,.18) 34%,transparent 52%);
          }
          /* Lớp 5 — trăng rằm, chỉ dịp Trung Thu */
          #boot-splash .bs-moon {
            position: absolute; z-index: 3; right: 31%; top: 15%; display: none;
            width: 132px; height: 132px; border-radius: 50%;
            background: radial-gradient(circle at 60% 36%, #fffdf0 58%, #f2e2b8);
            animation: bsMoon 6s ease-in-out infinite;
          }
          #boot-splash[data-s="trungthu"] .bs-moon { display: block; }
          @keyframes bsMoon {
            0%,100% { box-shadow: 0 0 110px 38px rgba(255,225,150,.36); }
            50%     { box-shadow: 0 0 140px 54px rgba(255,225,150,.52); }
          }
          /* Lớp 6 — hạt rơi theo mùa */
          #boot-splash .bs-pt { position: absolute; z-index: 2; will-change: transform; pointer-events: none; }
          @keyframes bsFall { to { transform: translate(var(--dx), 118vh) rotate(700deg); } }
          @keyframes bsRise { to { transform: translate(var(--dx), -118vh) rotate(24deg); } }

          /* Thẻ kính chứa lời chào */
          #boot-splash .bs-card {
            position: absolute; left: 80px; top: calc(50% - 44px); transform: translateY(-50%); z-index: 6;
            width: 498px; max-width: calc(100vw - 48px); padding: 32px 34px 28px; border-radius: 24px;
            background: rgba(255,255,255,.10); backdrop-filter: blur(26px) saturate(150%);
            border: 1px solid rgba(255,255,255,.20);
            box-shadow: 0 30px 90px rgba(0,0,0,.5), inset 0 1px 0 rgba(255,255,255,.24);
            animation: bsCard .75s cubic-bezier(.34,1.32,.64,1) both;
          }
          @keyframes bsCard {
            from { opacity: 0; transform: translateY(-50%) translateX(-30px); }
            to   { opacity: 1; transform: translateY(-50%); }
          }
          #boot-splash .bs-top { display: flex; align-items: center; gap: 12px; }
          #boot-splash .bs-logo {
            width: 42px; height: 42px; border-radius: 14px; flex-shrink: 0;
            background: linear-gradient(140deg,#ff8fb1,#a78bfa);
            display: flex; align-items: center; justify-content: center;
            box-shadow: 0 10px 24px rgba(190,110,220,.45);
          }
          #boot-splash .bs-logo svg { width: 27px; height: 27px; }
          #boot-splash .bs-bn { font-size: .9rem; font-weight: 760; color: #fff; }
          #boot-splash .bs-bs { font-size: .7rem; color: rgba(255,255,255,.58); }
          #boot-splash .bs-kick {
            margin-left: auto; font-size: .68rem; font-weight: 750; letter-spacing: .06em; color: #fff;
            padding: 5px 11px; border-radius: 999px; white-space: nowrap;
            background: rgba(255,255,255,.14); border: 1px solid rgba(255,255,255,.2);
          }
          #boot-splash .bs-hi {
            margin-top: 22px; font-size: 2.05rem; font-weight: 840; letter-spacing: -.04em;
            color: #fff; line-height: 1.1;
          }
          #boot-splash .bs-hi em {
            font-style: normal; background: linear-gradient(100deg,#ffc9de,#c4b5fd 55%,#93c5fd);
            -webkit-background-clip: text; background-clip: text; color: transparent;
          }
          #boot-splash .bs-sub { margin-top: 9px; font-size: .815rem; font-weight: 600; color: rgba(255,255,255,.7); }
          #boot-splash .bs-enter {
            margin-top: 26px; display: flex; align-items: center; gap: 9px;
            font-size: .82rem; font-weight: 600; color: rgba(255,255,255,.78);
            opacity: 0; animation: bsFade .45s ease 2s forwards;
          }
          #boot-splash .bs-enter kbd {
            font: inherit; font-weight: 800; color: #fff; background: rgba(255,255,255,.2);
            border: 1px solid rgba(255,255,255,.3); border-radius: 8px; padding: 4px 12px;
          }
          #boot-splash .bs-hint2 {
            margin-top: 9px; font-size: .72rem; color: rgba(255,255,255,.6);
            opacity: 0; animation: bsFade .45s ease 2.3s forwards;
          }
          @keyframes bsFade { to { opacity: 1; } }
          #boot-splash .bs-blink { animation: bsBlink 1.6s ease-in-out infinite; }
          @keyframes bsBlink { 0%,100% { opacity: .4; } 50% { opacity: 1; } }

          /* Đồng hồ + ngày, góc phải trên */
          #boot-splash .bs-corner {
            position: absolute; right: 74px; top: 60px; z-index: 6; text-align: right;
            animation: bsRight .8s cubic-bezier(.34,1.3,.64,1) .18s both;
          }
          @keyframes bsRight { from { opacity: 0; transform: translateX(26px); } to { opacity: 1; transform: none; } }
          #boot-splash .bs-clock {
            font-size: 3.9rem; font-weight: 250; color: #fff; letter-spacing: -.04em; line-height: 1;
            font-variant-numeric: tabular-nums; text-shadow: 0 6px 40px rgba(0,0,0,.5);
          }
          #boot-splash .bs-date { margin-top: 7px; font-size: .85rem; font-weight: 600; color: rgba(255,255,255,.7); }

          /* Thanh số liệu ở đáy */
          #boot-splash .bs-strip {
            position: absolute; left: 0; right: 0; bottom: 0; z-index: 6; height: 84px;
            display: flex; align-items: center; padding: 0 80px;
            background: linear-gradient(0deg,rgba(6,8,14,.72),rgba(6,8,14,.14));
            border-top: 1px solid rgba(255,255,255,.13); backdrop-filter: blur(14px);
            animation: bsUp .6s ease .45s both;
          }
          @keyframes bsUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: none; } }
          #boot-splash .bs-cell { flex: 1; display: flex; flex-direction: column; gap: 3px; position: relative; }
          #boot-splash .bs-cell + .bs-cell { padding-left: 32px; }
          #boot-splash .bs-cell + .bs-cell::before {
            content: ""; position: absolute; left: 0; top: 5px; bottom: 5px; width: 1px;
            background: rgba(255,255,255,.14);
          }
          #boot-splash .bs-k { font-size: .7rem; font-weight: 640; letter-spacing: .05em; color: rgba(255,255,255,.55); }
          #boot-splash .bs-v {
            font-size: 1.36rem; font-weight: 800; color: #fff;
            font-variant-numeric: tabular-nums; letter-spacing: -.02em;
          }
          #boot-splash .bs-v.warn { color: #ffd07a; }
          #boot-splash .bs-v.mut { color: rgba(255,255,255,.4); font-weight: 600; }

          /* Màn hẹp (điện thoại): bỏ đồng hồ, thẻ kính tràn ngang, thanh số liệu gọn lại */
          @media (max-width: 820px) {
            #boot-splash .bs-corner { display: none; }
            #boot-splash .bs-card { left: 24px; right: 24px; width: auto; padding: 24px 22px 22px; }
            #boot-splash .bs-hi { font-size: 1.6rem; }
            #boot-splash .bs-strip { padding: 0 22px; height: 74px; }
            #boot-splash .bs-cell:nth-child(n+4) { display: none; }
            #boot-splash .bs-cell + .bs-cell { padding-left: 18px; }
            #boot-splash .bs-k {
                font-size: .61rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
            }
            #boot-splash .bs-v { font-size: 1.15rem; }
          }
        `;
        doc.head.appendChild(css3);

        var el = doc.createElement('div');
        el.id = 'boot-splash';
        el.setAttribute('data-h', '__GREET_HOUR__');
        el.setAttribute('data-s', '__GREET_SEASON__');
        el.innerHTML =
            '<div class="bs-bg"></div><div class="bs-hour"></div><div class="bs-season"></div>' +
            '<div class="bs-scrim"></div><div class="bs-moon"></div>' +
            '<div class="bs-card">' +
              '<div class="bs-top">' +
                '<div class="bs-logo"><svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg"><g>' +
                  '<path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#fff"/>' +
                  '<path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#ffe8f1" transform="rotate(72 50 50)"/>' +
                  '<path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#fff" transform="rotate(144 50 50)"/>' +
                  '<path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#ffe8f1" transform="rotate(216 50 50)"/>' +
                  '<path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#fff" transform="rotate(288 50 50)"/>' +
                  '<circle cx="50" cy="50" r="7" fill="#ffd6e6"/>' +
                '</g></svg></div>' +
                '<div><div class="bs-bn">Tân Hotel</div><div class="bs-bs">Front Office toolkit</div></div>' +
                '<span class="bs-kick">__GREET_SEASON_LABEL__</span>' +
              '</div>' +
              '<div class="bs-hi">Chào <em>__GREET_SHIFT__</em>, Tân __GREET_EMOJI__</div>' +
              '<div class="bs-sub">__GREET_SUB__</div>' +
              '<div class="bs-enter"><span class="bs-blink">&#9654;</span> Nhấn <kbd>Enter</kbd> để vào</div>' +
              '<div class="bs-hint2">hoặc chạm/bấm chuột vào màn hình</div>' +
            '</div>' +
            '<div class="bs-corner">' +
              '<div class="bs-clock">__GREET_CLOCK__</div>' +
              '<div class="bs-date">__GREET_DATE__</div>' +
            '</div>' +
            '<div class="bs-strip">__GREET_CELLS__</div>';
        doc.body.appendChild(el);

        // ── Hạt rơi theo mùa — thuần transform/opacity, chạy trên GPU ──
        var BS_PT = {
          spring:   {n: 26, col: ['#f6a8c9','#f293bc','#f9c1d9'], r: '0 60% 0 60%', sz: [10,18],  dur: [5,9],   a: 'bsFall'},
          summer:   {n: 32, col: ['#fff4c0','#ffffff','#bff0ff'], r: '50%',         sz: [3,8],    dur: [6,11],  a: 'bsRise'},
          autumn:   {n: 24, col: ['#e0913f','#c96a2b','#e8b45a'], r: '0 70% 0 70%', sz: [11,19],  dur: [6,10],  a: 'bsFall'},
          winter:   {n: 70, col: ['rgba(200,225,250,.85)'],       r: '2px',         sz: [1.6,2.8],dur: [1.0,1.9],a: 'bsFall', tall: 9},
          tet:      {n: 26, col: ['#ffd24a','#ffc107','#ffe27a'], r: '0 60% 0 60%', sz: [10,17],  dur: [5,9],   a: 'bsFall'},
          trungthu: {n: 13, col: ['#ff9f43','#ffb86b','#ff7f50'], r: '34% 34% 42% 42%', sz: [16,26], dur: [11,18], a: 'bsRise', glow: true},
          noel:     {n: 80, col: ['#ffffff','#eaf4ff'],           r: '50%',         sz: [3,7],    dur: [6,12],  a: 'bsFall'}
        }['__GREET_SEASON__'];
        if (BS_PT) {
            for (var j = 0; j < BS_PT.n; j++) {
                var pt = doc.createElement('div');
                pt.className = 'bs-pt';
                var sz = BS_PT.sz[0] + Math.random() * (BS_PT.sz[1] - BS_PT.sz[0]);
                var du = BS_PT.dur[0] + Math.random() * (BS_PT.dur[1] - BS_PT.dur[0]);
                var col = BS_PT.col[j % BS_PT.col.length];
                var fill = (BS_PT.r === '50%')
                    ? 'radial-gradient(circle,' + col + ',transparent 72%)'
                    : 'radial-gradient(circle at 30% 30%,#fff,' + col + ' 72%)';
                if (BS_PT.glow) fill = 'radial-gradient(circle at 50% 38%,#fff2c0,' + col + ' 72%)';
                pt.style.cssText = (BS_PT.a === 'bsRise' ? 'top:112vh;' : 'top:-42px;') +
                    'left:' + (Math.random() * 100) + 'vw;' +
                    'width:' + sz + 'px;height:' + (BS_PT.tall ? sz * BS_PT.tall : sz * (BS_PT.glow ? 1.3 : 1)) + 'px;' +
                    'background:' + fill + ';border-radius:' + BS_PT.r + ';' +
                    (BS_PT.glow ? 'box-shadow:0 0 18px 4px rgba(255,170,80,.55);' : '') +
                    '--dx:' + (Math.random() * 180 - 90) + 'px;' +
                    'animation:' + BS_PT.a + ' ' + du + 's linear ' + (-Math.random() * du) + 's infinite';
                el.appendChild(pt);
            }
        }

        // ── Nội dung app hiện dần so le sau khi màn chào tan ──
        function bsRevealApp() {
            var seq = [];
            var sb = doc.querySelector('section[data-testid="stSidebar"]');
            if (sb) seq.push(sb);
            var blk = doc.querySelector('div[data-testid="stMainBlockContainer"] > div[data-testid="stVerticalBlock"]');
            if (blk && blk.children.length) {
                Array.prototype.forEach.call(blk.children, function(c){ seq.push(c); });
            } else {
                var m = doc.querySelector('[data-testid="stMain"]');
                if (m) seq.push(m);
            }
            seq.forEach(function(node, i){
                node.style.animationDelay = (i * 0.06) + 's';
                node.classList.add('tan-rv');
            });
            // Gỡ lớp hiệu ứng sau khi chạy xong để không ảnh hưởng các lần rerun sau
            window.parent.setTimeout(function(){
                seq.forEach(function(node){
                    node.classList.remove('tan-rv');
                    node.style.animationDelay = '';
                });
            }, 2200);
        }

        // ── Cổng Enter: chỉ mở sau 2 giây, bấm sớm hơn không có tác dụng ──
        var bsArmed = false, bsDone = false;
        window.parent.setTimeout(function(){ bsArmed = true; }, 2000);

        function bsDismiss() {
            if (bsDone || !bsArmed) return;
            bsDone = true;
            doc.removeEventListener('keydown', bsKey, true);
            document.removeEventListener('keydown', bsKey, true);
            el.classList.add('bs-hide');
            bsRevealApp();
            window.parent.setTimeout(function(){
                if (el.parentNode) el.parentNode.removeChild(el);
            }, 600);
        }
        function bsKey(ev) {
            if (ev.key === 'Enter' || ev.key === ' ' || ev.key === 'Spacebar' || ev.key === 'Escape') {
                ev.preventDefault();
                bsDismiss();
            }
        }
        // Bắt phím ở CẢ tài liệu cha lẫn iframe này — tuỳ nơi con trỏ đang focus
        doc.addEventListener('keydown', bsKey, true);
        document.addEventListener('keydown', bsKey, true);
        el.addEventListener('click', bsDismiss);
        el.addEventListener('touchstart', bsDismiss, {passive: true});
        try { window.parent.focus(); } catch (e) {}
    }
})();
</script>
"""
    _boot_script = _boot_script.replace('__DARK_BG_DATA_URI__', _dark_bg_data_uri())
    _boot_script = _boot_script.replace('__LIGHT_BG_DATA_URI__', _light_bg_data_uri())

    # ── Nội dung màn chào: lấy từ SỐ LIỆU THẬT đã lưu trong ngày. Chưa chạy
    # công cụ nào thì ô số liệu hiện dấu "—", KHÔNG bịa con số. ──
    _g_now = now_vn()
    _g_theme = _compute_effective_theme()
    _g_season, _g_hourkey = _season_key(), _hour_key()
    _g_thu = ['Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy', 'Chủ Nhật'][_g_now.weekday()]
    if THEME_DAY_START_HOUR <= _g_now.hour < 14:
        _g_shift, _g_hours, _g_emo = 'ca sáng', '06:00 – 14:00', '🌅'
    elif 14 <= _g_now.hour < 22:
        _g_shift, _g_hours, _g_emo = 'ca chiều', '14:00 – 22:00', '☀️' if _g_hourkey == 'day' else '🌇'
    else:
        _g_shift, _g_hours, _g_emo = 'ca đêm', '22:00 – 06:00', '🌙'

    _g_tasks = (_load_progress().get('tasks') or {})
    _g_sum = ((_g_tasks.get('daily') or {}).get('summary') or {})
    _g_todo = sum(1 for _k in ('daily', 'regcard', 'recon_person', 'recon_room')
                  if not (_g_tasks.get(_k) or {}).get('done'))
    _g_todo_txt = (f'còn {_g_todo} việc chưa xong' if _g_todo
                   else 'đã xong các việc trong ngày')

    def _g_cell(label, value, cls=''):
        _mut = ' mut' if value is None else (f' {cls}' if cls else '')
        return (f'<div class="bs-cell"><div class="bs-k">{label}</div>'
                f'<div class="bs-v{_mut}">{"—" if value is None else value}</div></div>')

    _g_cells = (_g_cell('KHÁCH LƯU TRÚ', _g_sum.get('total'))
                + _g_cell('QUỐC TẾ', _g_sum.get('intl'))
                + _g_cell('VIỆT NAM', _g_sum.get('vn'))
                + _g_cell('CHECK-IN HÔM NAY', _g_sum.get('checkin_n'))
                + _g_cell('VIỆC CHƯA XONG', _g_todo, 'warn' if _g_todo else ''))

    for _ph, _val in (
        ('__GREET_PHOTO__', _season_bg_data_uri(_g_season, _g_theme)),
        ('__GREET_HOUR__', _g_hourkey),
        ('__GREET_SEASON_LABEL__', SEASON_LABEL[_g_season]),
        ('__GREET_SEASON__', _g_season),
        ('__GREET_SHIFT__', _g_shift),
        ('__GREET_EMOJI__', _g_emo),
        ('__GREET_SUB__', f'{_g_thu}, {_g_now.strftime("%d/%m/%Y")} · {_g_hours} · {_g_todo_txt}'),
        ('__GREET_CLOCK__', _g_now.strftime('%H:%M')),
        ('__GREET_DATE__', f'{_g_thu}, {_g_now.strftime("%d/%m/%Y")}'),
        ('__GREET_CELLS__', _g_cells),
        ('__GREET_THEME__', _g_theme),
    ):
        _boot_script = _boot_script.replace(_ph, _val)
    st.iframe(_boot_script, height=1)

# ── Chế độ giao diện: áp dụng data-theme lên <html> mỗi lần rerun (khác khối
# CSS phía trên chỉ tiêm 1 lần/phiên) — vì hiệu lực có thể đổi giữa các lần
# rerun khi ở chế độ "Tự động" (qua giờ) hoặc khi người dùng đổi lựa chọn.
if 'theme_mode' not in st.session_state:
    st.session_state.theme_mode = 'auto'
_effective_theme = _compute_effective_theme()
# Badge số bên phải mục nav (số khách đã xử lý / số ghi chú sổ giao ca): số thay
# đổi mỗi lần rerun nên tiêm qua CSS ::after ở đây, không nhét vào khối CSS tĩnh.
_bd = st.session_state.get('daily_results') or {}
_bd_n = _bd.get('total') or (st.session_state.get('progress', {})
                             .get('tasks', {}).get('daily', {}).get('summary', {}) or {}).get('total')
try:
    _ho_n = len(db_load_entries(today_vn())) if db_available() else len(
        (st.session_state.get('handover') or {}).get('entries', []))
except Exception:
    _ho_n = 0
_badge_css = ''
if _bd_n:
    _badge_css += f'.st-key-nav_daily button::after{{content:"{_bd_n}";}}'
if _ho_n:
    _badge_css += f'.st-key-nav_handover button::after{{content:"{_ho_n}";}}'
st.iframe(f"""
<script>
(function(){{
  var doc = window.parent.document;
  doc.documentElement.setAttribute('data-theme', '{_effective_theme}');
  var s = doc.getElementById('tan-badge-style');
  if (!s) {{ s = doc.createElement('style'); s.id = 'tan-badge-style'; doc.head.appendChild(s); }}
  s.textContent = {_badge_css!r};
}})();
</script>
""", height=1)



# Menu selection (session state)
if "menu" not in st.session_state:
    st.session_state.menu = "dashboard"

MENU_LABELS = {
    'dashboard': 'Tổng quan ca trực', 'daily': 'Xử lý hàng ngày', 'regcard': 'Regcard + ARR',
    'handover': 'Sổ giao ca', 'recon': 'Đối chiếu (cổng mật khẩu)',
    'recon_person': 'Đối chiếu người nước ngoài', 'recon_room': 'Đối chiếu hệ thống phòng',
}

# Nạp tiến độ đã lưu trên đĩa cho HÔM NAY — chạy 1 lần khi phiên bắt đầu, và
# tự làm mới nếu phiên mở vắt qua nửa đêm (ngày mới → tiến độ trống lại).
if st.session_state.get('progress_date') != today_vn().isoformat():
    st.session_state.progress_date = today_vn().isoformat()
    st.session_state.progress = _load_progress()
    st.session_state.handover = {'entries': list(st.session_state.progress.get('handover_entries', []))}

def set_theme_mode(mode):
    st.session_state.theme_mode = mode

def go_menu(name):
    st.session_state.menu = name

# ── Sidebar điều hướng ────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('''
    <div class="sb-brand">
        <div class="sb-brand-title"><span class="sb-mascot">
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
                <g>
                    <path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#ffc2dd"/>
                    <path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#ffb3d1" transform="rotate(72 50 50)"/>
                    <path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#ffc2dd" transform="rotate(144 50 50)"/>
                    <path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#ffb3d1" transform="rotate(216 50 50)"/>
                    <path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#ffc2dd" transform="rotate(288 50 50)"/>
                    <circle cx="50" cy="50" r="7" fill="#fff6ee"/>
                    <circle cx="47" cy="47" r="1.3" fill="#ffcf6b"/>
                    <circle cx="53" cy="47" r="1.3" fill="#ffcf6b"/>
                    <circle cx="50" cy="52.5" r="1.3" fill="#ffcf6b"/>
                </g>
            </svg>
        </span> Tân Hotel</div>
        <div class="sb-brand-sub">Front Office toolkit</div>
    </div>
    ''', unsafe_allow_html=True)

    st.markdown('<div class="sb-section">Công cụ</div>', unsafe_allow_html=True)
    st.button("Tổng quan ca trực", key="nav_dashboard", use_container_width=True,
              icon=":material/dashboard:",
              type="primary" if st.session_state.menu == "dashboard" else "secondary",
              on_click=go_menu, args=("dashboard",))
    st.button("Xử lý hàng ngày", key="nav_daily", use_container_width=True,
              icon=":material/checklist:",
              type="primary" if st.session_state.menu == "daily" else "secondary",
              on_click=go_menu, args=("daily",))
    st.button("Regcard + ARR", key="nav_regcard", use_container_width=True,
              icon=":material/print:",
              type="primary" if st.session_state.menu == "regcard" else "secondary",
              on_click=go_menu, args=("regcard",))
    st.button("Sổ giao ca", key="nav_handover", use_container_width=True,
              icon=":material/handshake:",
              type="primary" if st.session_state.menu == "handover" else "secondary",
              on_click=go_menu, args=("handover",))

    st.markdown('<div class="sb-section">Đối chiếu</div>', unsafe_allow_html=True)
    if st.session_state.get("recon_ok"):
        st.button("Người nước ngoài", key="nav_recon_person", use_container_width=True,
                  icon=":material/public:",
                  type="primary" if st.session_state.menu == "recon_person" else "secondary",
                  on_click=go_menu, args=("recon_person",))
        st.button("Hệ thống phòng", key="nav_recon_room", use_container_width=True,
                  icon=":material/door_front:",
                  type="primary" if st.session_state.menu == "recon_room" else "secondary",
                  on_click=go_menu, args=("recon_room",))
    else:
        st.button("Đối chiếu lưu trú", key="nav_recon", use_container_width=True,
                  icon=":material/lock:",
                  type="primary" if st.session_state.menu == "recon" else "secondary",
                  on_click=go_menu, args=("recon",))

    st.markdown('<div class="sb-section">Giao diện</div>', unsafe_allow_html=True)
    # Nút gạt 3 trạng thái (segmented) thay cho ô chọn — nhìn gọn và bấm 1 lần
    _tcols = st.columns(3, gap="small")
    for _tci, (_tk, _tlbl) in enumerate([('auto', 'Tự động'), ('light', 'Sáng'), ('dark', 'Tối')]):
        _tcols[_tci].button(_tlbl, key=f"theme_btn_{_tk}", use_container_width=True,
                            type="primary" if st.session_state.get('theme_mode') == _tk else "secondary",
                            on_click=set_theme_mode, args=(_tk,))

    _sb_shift = ('Ca sáng · 06:00–14:00' if 6 <= now_vn().hour < 14
                 else 'Ca chiều · 14:00–22:00' if 14 <= now_vn().hour < 22 else 'Ca đêm · 22:00–06:00')
    _sb_store = 'Supabase đã kết nối' if db_available() else 'Lưu tạm trên máy chủ'
    st.markdown(
        f'<div class="sb-user"><div class="sb-av">T</div>'
        f'<div><div class="sb-un">Tân</div><div class="sb-ur">{_sb_shift}</div></div></div>'
        f'<div class="sb-status"><span class="sb-dot"></span>{_sb_store}</div>',
        unsafe_allow_html=True)

# ── Thanh top bar: breadcrumb + trạng thái + hành động (thay banner "Welcome") ──
_page_name = MENU_LABELS.get(st.session_state.menu, 'Tổng quan ca trực')
_tb_now = now_vn()
_tb_chip = ('<span class="tan-chip ok">● Supabase đã kết nối</span>' if db_available()
            else '<span class="tan-chip">● Lưu tạm trên máy chủ</span>')
_tb_d = st.session_state.get('daily_results')
_tb_rc = st.session_state.get('rc_results')
_tb_rp = st.session_state.get('recon_results')
_tb_rr = st.session_state.get('reconr_results')
# Báo cáo ngày chỉ dựng khi đang ở Tổng quan VÀ đã có ít nhất 1 công cụ chạy
# xong — tránh dựng workbook thừa ở mọi trang, mọi lần rerun.
_tb_report = bool(_tb_d or _tb_rc or _tb_rp or _tb_rr) and st.session_state.menu == "dashboard"

with st.container(key="tan_topbar"):
    _tbl, _tbm, _tbr = st.columns([4.6, 1.25, 1.35], vertical_alignment="center")
    _tbl.markdown(
        f'<div class="tan-topbar"><span class="tan-crumb">Tân Hotel</span>'
        f'<span class="tan-sep">/</span><b>{_page_name}</b>'
        f'<span class="tan-topbar-r">{_tb_chip}'
        f'<span class="tan-chip">📅 {_tb_now.strftime("%d/%m/%Y")}</span></span></div>',
        unsafe_allow_html=True)
    if _tb_report:
        _tb_rp_date = (_tb_d or {}).get('date_str') or today_vn().strftime('%d_%m')
        _tb_wb = build_daily_report(_tb_rp_date.replace('_', '/'),
                                    _tb_d if _tb_d and _tb_d.get('has_xlsx') else None,
                                    (_tb_rc or {}).get('arr_stats'), _tb_rp, _tb_rr)
        _tbm.download_button("⇩ Xuất báo cáo", wb_to_bytes(_tb_wb),
                             file_name=f"bao_cao_ngay_{_tb_rp_date}.xlsx",
                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             use_container_width=True, key="tb_report_dl")
    else:
        _tbm.button("⇩ Xuất báo cáo", key="tb_report_off", use_container_width=True, disabled=True,
                    help="Mở Tổng quan ca trực và chạy ít nhất một công cụ để xuất báo cáo ngày")
    _tbr.button("⚡ Xử lý hàng ngày", key="tb_go_daily", use_container_width=True,
                type="primary", on_click=go_menu, args=("daily",))

# ── Tạo file ARR từ file Arrival (Book) Smile ──────────────────────────────
def build_arr(book_bytes):
    """Tạo file ARR ĐÚNG định dạng của ARR Converter gốc (tool HTML riêng, không
    phải file mẫu in cũ):
    - 6 cột: Conf# / Arrival / Departure / Company / Notice / [số phòng] — đọc
      cột nguồn theo TÊN (Conf#, Folio#, Type, Arrival, Departure, Company,
      Notice), không theo vị trí cố định như bản cũ.
    - Font Patrick Hand toàn bộ; Conf# cỡ 50 đậm nền cam nhạt; số phòng cỡ 50;
      các ô còn lại cỡ 20. Dòng dữ liệu cao 120, dòng header cao 142.5.
    - Số phòng = số dòng Folio# hợp lệ trùng Conf# (bỏ dòng Type='**' - dummy).
    - Dòng phụ chèn ngay sau booking tương ứng, gộp A:F, nền màu theo loại:
      CÀ THẺ (cam) · THU TIỀN (xanh lá) · XEM LẠI BU (vàng) · FOC LATE C/O (xanh
      dương, tự đọc giờ trong Notice nếu có, vd "FOC LATE C/O 18:00").
    - Nhận diện nghiệp vụ đầy đủ (OTA, từ khóa CÀ THẺ/THU TIỀN/FOC/XEM LẠI BU)
      y hệt bộ từ khóa của ARR Converter gốc.
    """
    import re as _re2
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    df = pd.read_excel(io.BytesIO(book_bytes), header=None)
    hdr = None
    for i in range(min(5, len(df))):
        if any(str(v).strip() == 'Conf#' for v in df.iloc[i] if pd.notna(v)):
            hdr = i; break
    if hdr is None:
        raise ValueError("Không tìm thấy dòng header chứa \"Conf#\" trong file. Kiểm tra lại file Arrival Smile.")

    headers = [str(v).strip() if pd.notna(v) else '' for v in df.iloc[hdr]]
    col = {}
    for i, h in enumerate(headers):
        if h and h not in col:
            col[h] = i
    data = df.iloc[hdr + 1:].reset_index(drop=True)

    def C(name):
        return col.get(name, -1)

    conf_c, folio_c, type_c = C('Conf#'), C('Folio#'), C('Type')
    arr_c, dep_c, comp_c, notice_c = C('Arrival'), C('Departure'), C('Company'), C('Notice')
    if -1 in (conf_c, folio_c, arr_c, comp_c):
        raise ValueError("File thiếu cột bắt buộc (Conf#, Folio#, Arrival, Company). Kiểm tra lại file.")

    # Sửa lỗi Excel đảo dd/mm↔mm/dd với ngày ≤12 từ Smile export (dùng chung hàm _fix_date)
    def _arr_fmt_date(v):
        fixed = _fix_date(v)
        if fixed is not None:
            return fixed.strftime('%d/%m/%y')
        return str(v).strip() if isinstance(v, str) and v.strip() else ''

    # ── Đếm số phòng (= số dòng Folio# hợp lệ) theo từng Conf#, bỏ dòng Type='**' (dummy) ──
    room_counts = {}
    dummy_count = 0
    for _, row in data.iterrows():
        conf = row.iloc[conf_c] if conf_c >= 0 else None
        folio = row.iloc[folio_c] if folio_c >= 0 else None
        if pd.notna(conf) and pd.notna(folio):
            typ = row.iloc[type_c] if type_c >= 0 else None
            if str(typ).strip() == '**':
                dummy_count += 1
                continue
            room_counts[conf] = room_counts.get(conf, 0) + 1

    # ── Danh sách booking theo ĐÚNG thứ tự xuất hiện, mỗi Conf# 1 dòng ──
    seen = set()
    ordered = []
    for _, row in data.iterrows():
        conf = row.iloc[conf_c] if conf_c >= 0 else None
        folio = row.iloc[folio_c] if folio_c >= 0 else None
        typ = row.iloc[type_c] if type_c >= 0 else None
        arr = row.iloc[arr_c] if arr_c >= 0 else None
        comp = row.iloc[comp_c] if comp_c >= 0 else None
        dep = row.iloc[dep_c] if dep_c >= 0 else None
        notice = row.iloc[notice_c] if notice_c >= 0 else None
        if pd.isna(conf) or pd.isna(folio) or pd.isna(arr) or pd.isna(comp):
            continue
        if str(typ).strip() == '**':
            continue
        if conf in seen:
            continue
        if not room_counts.get(conf):
            continue
        seen.add(conf)
        ordered.append({
            'type': 'bk', 'conf': conf,
            'arrival': _arr_fmt_date(arr),
            'departure': _arr_fmt_date(dep) if pd.notna(dep) else '',
            'company': str(comp).strip(),
            'notice': str(notice).strip() if pd.notna(notice) else '',
            'rooms': room_counts[conf],
        })
    if not ordered:
        raise ValueError("File không có dữ liệu booking hợp lệ nào.")

    # ── Bộ nhận diện nghiệp vụ — y hệt ARR Converter gốc ──
    ARR_OTA = ['EXPEDIA','BOOKING','AGODA','TRIP.COM','CTRIP','AIRBNB','TRAVELOKA',
        'KLOOK','KAYAK','PRICELINE','HOTELS.COM','ORBITZ','TRIVAGO','MAKEMYTRIP',
        'LASTMINUTE','HOSTELWORLD','WOTIF','HOTWIRE','VRBO','HOMEAWAY','IVIVU',
        'MYTOUR','LUXSTAY','VNTRIP','GOTADI','TRAVELPORT','SKYSCANNER','BESTPRICE',
        'LATEROOMS','EASYJET','RYANAIR','JETSTAR','HOTELBEDS','TOURICO','GETAROOM']
    ARR_CA_THE = ['TACC','CC UPON','CHARGE CC','CHARGE CARD','CREDIT CARD','DEBIT CARD',
        'CC AUTH','AUTH CC','AUTHORIZE CC','AUTHORIZE CARD','CC ON ARRIVAL',
        'BILL TO CC','SWIPE CC','SWIPE CARD','PRE-AUTH','PREAUTH','PREPAID CC',
        'CHARGE ON CC','CARD ON ARRIVAL','CC AT CI','CC AT CHECK','PAY BY CARD',
        'CARD PAYMENT','TC UPON','TC ON ARRIVAL','TAKE CC','TAKE CARD']
    ARR_THU_TIEN = ['PAY UPON','PAY ON ARRIVAL','CASH ON ARRIVAL','CASH UPON','COLLECT CASH',
        'COLLECT PAYMENT','COLLECT ON ARRIVAL','PAYMENT ON ARRIVAL','CASH PAYMENT',
        'CASH AT CHECK','CASH AT CI','DUE ON ARRIVAL','PAYABLE ON ARRIVAL',
        'PAY AT CI','PAY AT CHECK','CASH DUE','OUTSTANDING','BALANCE DUE',
        'PAYMENT DUE','COLLECT AT CI','COLLECT AT CHECK',
        'RC UPON','UPON C/I','UPON CI','UPON CHECK-IN','UPON CHECKIN',
        'GOA UPON','ROH UPON','COLLECT UPON']
    ARR_XEM_LAI = ['CASH UPON','CASH ON ARRIVAL','CASH AT CI','CASH PAYMENT',
        'PAY UPON','PAY AT CI','COLLECT CASH','CASH DUE']
    ARR_FOC_LCO = ['FOC LATE CHECK','FOC LATE CHECKOUT','FOC LATE C/O','FOC LCO',
        'LCO FOC','LATE CHECK OUT FOC','LATE CHECKOUT FOC','LATE C/O FOC',
        'COMP LATE CHECK','COMP LATE CHECKOUT','COMP LCO',
        'COMPLIMENTARY LATE CHECK','COMPLIMENTARY LCO',
        'GRATIS LATE CHECK','FREE LATE CHECK']

    def _pay_type(notice, company):
        n = str(notice or '').upper()
        co = str(company or '').upper()
        is_ota = any(o in co for o in ARR_OTA)
        is_ca_the = any(k in n for k in ARR_CA_THE)
        is_thu_tien = any(k in n for k in ARR_THU_TIEN)
        is_foc_lco = any(k in n for k in ARR_FOC_LCO)
        has_foc = 'FOC' in n or 'COMP' in n or 'COMPLIMENTARY' in n
        has_lco = 'LATE CHECK' in n or ' LCO' in n or 'LATE C/O' in n or 'LATE CHECKOUT' in n
        if is_foc_lco or (has_foc and has_lco):
            return 'foc_lco'
        if is_ota and any(k in n for k in ARR_XEM_LAI):
            return 'xem_lai_bu'
        if is_ca_the:
            return 'ca_the'
        if is_thu_tien:
            return 'thu_tien'
        return 'none'

    result = []
    for i, bk in enumerate(ordered):
        result.append(bk)
        if i >= len(ordered) - 1:
            continue
        pt = _pay_type(bk['notice'], bk['company'])
        if pt == 'ca_the':
            result.append({'type': 'sep', 'conf': 'CÀ THẺ'})
        elif pt == 'thu_tien':
            result.append({'type': 'sep', 'conf': 'THU TIỀN'})
        elif pt == 'xem_lai_bu':
            result.append({'type': 'sep', 'conf': 'XEM LẠI BU'})
        elif pt == 'foc_lco':
            m = _re2.search(r'\b(\d{1,2}[:Hh]\d{2})\b', bk['notice'])
            time_str = (' ' + m.group(1).upper()) if m else ''
            result.append({'type': 'sep', 'conf': 'FOC LATE C/O' + time_str})

    # ── Xuất Excel đúng định dạng ARR Converter gốc ──
    wb = Workbook(); ws = wb.active; ws.title = 'Sheet1'
    for i, w in enumerate([39.4, 15.1, 16.0, 21.9, 50.0, 10.3], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    thin = Side(style='thin')
    border_all = Border(top=thin, bottom=thin, left=thin, right=thin)
    center_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)
    conf_fill = PatternFill('solid', fgColor='FDEADA')
    fill_colors = {'CÀ THẺ': 'FDEADA', 'THU TIỀN': 'D4F4E8', 'XEM LẠI BU': 'FFF8DC'}

    ws.row_dimensions[1].height = 142.5
    for i, h in enumerate(['Conf#', 'Arrival', 'Departure', 'Company', 'Notice', None], 1):
        cell = ws.cell(1, i)
        cell.value = h
        cell.font = Font(name='Patrick Hand', size=50 if h == 'Conf#' else 20)
        cell.alignment = center_wrap
        cell.border = border_all

    r = 2
    for item in result:
        ws.row_dimensions[r].height = 120.0
        if item['type'] == 'sep':
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
            label = item.get('conf') or 'CÀ THẺ'
            color = fill_colors.get(label) or ('D6EAF8' if label.startswith('FOC LATE C/O') else 'FDEADA')
            cell = ws.cell(r, 1)
            cell.value = label
            cell.font = Font(name='Patrick Hand', size=50)
            cell.alignment = center_wrap
            cell.fill = PatternFill('solid', fgColor=color)
            cell.border = border_all
        else:
            vals = [item['conf'], item['arrival'], item['departure'], item['company'], item['notice'], item['rooms']]
            for ci, v in enumerate(vals, 1):
                cell = ws.cell(r, ci)
                cell.value = v
                cell.font = Font(name='Patrick Hand', size=50 if ci in (1, 6) else 20, bold=(ci == 1))
                cell.alignment = center_wrap
                cell.border = border_all
                if ci == 1:
                    cell.fill = conf_fill
        r += 1

    bookings = [x for x in result if x['type'] == 'bk']
    stats = {
        'bookings': len(bookings),
        'rooms': sum(b['rooms'] for b in bookings),
        'ota': sum(1 for b in bookings if any(o in b['company'].upper() for o in ARR_OTA)),
        'dummy': dummy_count,
        'ca_the': sum(1 for x in result if x['type'] == 'sep' and x['conf'] == 'CÀ THẺ'),
        'thu_tien': sum(1 for x in result if x['type'] == 'sep' and x['conf'] == 'THU TIỀN'),
        'xem_lai_bu': sum(1 for x in result if x['type'] == 'sep' and x['conf'] == 'XEM LẠI BU'),
        'foc_lco': sum(1 for x in result if x['type'] == 'sep' and x['conf'].startswith('FOC LATE C/O')),
    }
    return wb, stats


# ── Dashboard ca trực ─────────────────────────────────────────────────────
@st.fragment
def _render_dashboard():
    _now = now_vn()
    _thu = ['Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy', 'Chủ Nhật'][_now.weekday()]
    _shift = ('Ca sáng' if 6 <= _now.hour < 14 else 'Ca chiều' if 14 <= _now.hour < 22 else 'Ca đêm')

    _d = st.session_state.get('daily_results')
    _rc = st.session_state.get('rc_results')
    _re_p = st.session_state.get('recon_results')
    _re_r = st.session_state.get('reconr_results')
    # Tiến độ đã lưu trên đĩa cho HÔM NAY — còn nguyên dù tải lại trang hoặc
    # công cụ đó được chạy ở 1 phiên/tab khác trước đó trong ngày.
    _ptasks = st.session_state.get('progress', {}).get('tasks', {})
    _p_daily = _ptasks.get('daily', {})
    _p_regcard = _ptasks.get('regcard', {})
    _p_rp = _ptasks.get('recon_person', {})
    _p_rr = _ptasks.get('recon_room', {})

    st.markdown(f'<div class="tan-h1">Tổng quan ca trực</div>'
                f'<div class="tan-h1sub">{_thu}, {_now.strftime("%d/%m/%Y")} · {_shift} · '
                f'cập nhật {_now.strftime("%H:%M")}</div>', unsafe_allow_html=True)

    # Số liệu khách: ưu tiên kết quả phiên hiện tại, dự phòng bản đã lưu trong ngày
    if _d and _d.get('has_xlsx'):
        _iss_d = _d.get('issues')
        _n_red = int((_iss_d['Mức độ'] == '🔴').sum()) if _iss_d is not None and len(_iss_d) else 0
        _n_yel = int((_iss_d['Mức độ'] == '🟡').sum()) if _iss_d is not None and len(_iss_d) else 0
        _hero = {'total': _d['total'], 'intl': _d['intl'], 'vn': _d['vn'], 'red': _n_red, 'yellow': _n_yel,
                 'rooms': _d.get('rooms_cnt'), 'cin': _d.get('checkin_n'), 'cout': _d.get('checkout_n'),
                 'visa_watch': len(_d.get('visa_watch') or []),
                 'unknown_nats': len(_d.get('unknown_nats') or []),
                 'invalid_ids': len(_d.get('kbtt_invalid_ids') or []), 'stale': None}
    elif _p_daily.get('summary'):
        _ps = _p_daily['summary']
        _hero = {'total': _ps.get('total'), 'intl': _ps.get('intl'), 'vn': _ps.get('vn'),
                 'red': _ps.get('red_issues') or 0, 'yellow': _ps.get('yellow_issues') or 0,
                 'rooms': _ps.get('rooms_cnt'), 'cin': _ps.get('checkin_n'), 'cout': _ps.get('checkout_n'),
                 'visa_watch': _ps.get('visa_watch_count') or 0,
                 'unknown_nats': 0, 'invalid_ids': 0, 'stale': _p_daily.get('time')}
    else:
        _hero = None

    # ── HÀNG 1: thẻ HERO (ảnh mèo) + 2 thẻ số liệu ──
    _r1a, _r1b, _r1c = st.columns([2, 1, 1], gap="small")
    with _r1a:
        if _hero:
            _yday = _yesterday_total()
            _cmp = ''
            if _yday:
                _delta = (_hero['total'] or 0) - _yday
                _cmp = (f'▲ {_delta} khách so với hôm qua · ' if _delta > 0 else
                        f'▼ {abs(_delta)} khách so với hôm qua · ' if _delta < 0 else
                        'bằng hôm qua · ')
            _room_txt = f"{_hero['rooms']} phòng có khách · " if _hero.get('rooms') else ''
            _stale_txt = (f"số liệu lúc {_hero['stale']}" if _hero['stale'] else "cập nhật trong phiên này")
            st.markdown(
                '<div class="tan-hero">'
                '<div class="tan-hero-lab">🛏️ Tổng khách lưu trú hôm nay</div>'
                f'<div class="tan-hero-val">{_hero["total"]}</div>'
                f'<div class="tan-hero-sub">{_cmp}{_room_txt}{_stale_txt}</div>'
                '<div class="tan-hero-split">'
                f'<div><div class="tan-hero-k">🌍 Quốc tế</div><div class="tan-hero-v">{_hero["intl"]}</div></div>'
                f'<div><div class="tan-hero-k">🇻🇳 Việt Nam</div><div class="tan-hero-v">{_hero["vn"]}</div></div>'
                f'<div><div class="tan-hero-k">🔑 Check-in</div><div class="tan-hero-v">'
                f'{_hero["cin"] if _hero.get("cin") is not None else "—"}</div></div>'
                f'<div><div class="tan-hero-k">🚪 Check-out</div><div class="tan-hero-v">'
                f'{_hero["cout"] if _hero.get("cout") is not None else "—"}</div></div>'
                '</div></div>', unsafe_allow_html=True)
        else:
            st.markdown(
                '<div class="tan-hero">'
                '<div class="tan-hero-lab">👋 Chào ca trực</div>'
                '<div class="tan-hero-val" style="font-size:1.45rem">Chưa có số liệu hôm nay</div>'
                '<div class="tan-hero-sub">Bắt đầu bằng <b>Xử lý hàng ngày</b> hoặc '
                '<b>Regcard + ARR</b> ở sidebar — số liệu sẽ tự lên đây.</div>'
                '<div class="tan-hero-split">'
                '<div><div class="tan-hero-k">🌍 Quốc tế</div><div class="tan-hero-v">—</div></div>'
                '<div><div class="tan-hero-k">🇻🇳 Việt Nam</div><div class="tan-hero-v">—</div></div>'
                '<div><div class="tan-hero-k">🔑 Check-in</div><div class="tan-hero-v">—</div></div>'
                '<div><div class="tan-hero-k">🚪 Check-out</div><div class="tan-hero-v">—</div></div>'
                '</div></div>', unsafe_allow_html=True)

    with _r1b:
        with st.container(border=True, height=200):
            _red_n = _hero['red'] if _hero else None
            _cls = 'err' if _red_n else 'ok'
            _red_pct = round(100 * _red_n / _hero['total']) if (_hero and _hero.get('total') and _red_n) else 0
            st.markdown(
                f'<div class="kpi-lab"><span class="kpi-ic {_cls}">⚠️</span>Lỗi dữ liệu</div>'
                f'<div class="kpi-val {"err" if _red_n else ""}">{_red_n if _hero else "—"}</div>'
                f'<div class="kpi-sub">{"cần sửa trước khi nộp" if _red_n else ("dữ liệu sạch" if _hero else "chưa có dữ liệu")}</div>'
                f'<div class="kpi-bar"><i class="err" style="width:{min(_red_pct, 100)}%"></i></div>',
                unsafe_allow_html=True)

    with _r1c:
        with st.container(border=True, height=200):
            _rate = st.session_state.get('rate_input') or (_d or {}).get('rate')
            st.markdown(
                '<div class="kpi-lab"><span class="kpi-ic acc">💱</span>Tỷ giá VCB</div>'
                f'<div class="kpi-val" style="font-size:1.5rem">{f"{_rate:,.0f}" if _rate else "—"}</div>'
                f'<div class="kpi-sub">{"USD → VNĐ (chuyển khoản)" if _rate else "lấy tỷ giá ở Xử lý hàng ngày"}</div>'
                f'<div class="kpi-bar"><i style="width:{100 if _rate else 0}%"></i></div>',
                unsafe_allow_html=True)

    st.write("")

    # ── HÀNG 2: tiến độ công việc + cảnh báo + bàn giao gần nhất ──
    _TICK_ON = ('<span style="display:inline-block;width:15px;height:15px;border-radius:5px;'
                'background:linear-gradient(135deg,#34d399,#059669);color:#fff;font-size:9px;'
                'line-height:15px;text-align:center;vertical-align:-2px;">✓</span>')
    _TICK_OFF = ('<span style="display:inline-block;width:15px;height:15px;border-radius:5px;'
                 'border:1.5px solid var(--line);vertical-align:-2px;"></span>')
    _r2a, _r2b, _r2c = st.columns([2, 1, 1], gap="small")

    with _r2a:
        with st.container(border=True, height=420, key='dash_tasks'):
            if db_available():
                _ho_rows = db_load_entries(today_vn())
                _handover_n = len(_ho_rows)
            else:
                _ho_rows = None
                _handover_n = len((st.session_state.get('handover') or {}).get('entries', []))
            _tasks = [
                ("Xử lý hàng ngày (KBTT · VNM · ĐK14)", _d is not None or _p_daily.get('done'),
                 "daily", None if _d else _p_daily.get('time')),
                ("Regcard + file ARR", _rc is not None or _p_regcard.get('done'),
                 "regcard", None if _rc else _p_regcard.get('time')),
                ("Đối chiếu người nước ngoài", _re_p is not None or _p_rp.get('done'),
                 "recon_person" if st.session_state.get("recon_ok") else "recon",
                 None if _re_p else _p_rp.get('time')),
                ("Đối chiếu hệ thống phòng", _re_r is not None or _p_rr.get('done'),
                 "recon_room" if st.session_state.get("recon_ok") else "recon",
                 None if _re_r else _p_rr.get('time')),
                (f"Sổ giao ca ({_handover_n} ghi chú)", _handover_n > 0, "handover", None),
            ]
            _done_n = sum(1 for _t in _tasks if _t[1])
            _pend = next((_t[2] for _t in _tasks if not _t[1]), None)
            _hc1, _hc2 = st.columns([2.5, 1], vertical_alignment="center")
            _hc1.markdown(f'<div class="pan-h" style="border:0;padding-bottom:0;margin-bottom:0">'
                          f'<span class="pan-t">Tiến độ công việc trong ca</span>'
                          f'<span class="tan-chip">{_done_n}/{len(_tasks)}</span></div>',
                          unsafe_allow_html=True)
            _hc2.button("Xem tất cả →", key="dash_seeall", use_container_width=True,
                        disabled=_pend is None,
                        help="Mở công cụ chưa chạy đầu tiên trong ca",
                        on_click=go_menu, args=(_pend or "daily",))
            st.markdown('<div style="border-bottom:1px solid var(--line2);margin:0 0 .35rem"></div>',
                        unsafe_allow_html=True)
            _row_html = []
            for _label, _done, _target, _stale_time in _tasks:
                _chip = ('<span class="tan-chip ok">Xong</span>' if _done
                         else '<span class="tan-chip">Chưa chạy</span>')
                _tm = _stale_time or ('✓' if _done else '—')
                _row_html.append(
                    f'<div class="pan-task">{_TICK_ON if _done else _TICK_OFF}'
                    f'<span class="pan-rt">{_label}</span>{_chip}'
                    f'<span class="pan-time">{_tm}</span></div>')
            st.markdown(''.join(_row_html), unsafe_allow_html=True)
            # 2 dòng tóm tắt cuối panel — số booking/phòng đến (khi Regcard + ARR đã chạy)
            _sum_as = (_rc or {}).get('arr_stats') or _p_regcard.get('summary')
            st.markdown(
                f'<div class="pan-foot"><div class="pan-line"><span>Phòng đến trong ngày</span>'
                f'<span class="pan-v">{(_sum_as or {}).get("rooms", "—")}</span></div>'
                f'<div class="pan-line"><span>Booking đã xử lý</span>'
                f'<span class="pan-v">{(_sum_as or {}).get("bookings", "—")}</span></div></div>',
                unsafe_allow_html=True)

    with _r2b:
        with st.container(border=True, height=420, key='dash_alerts'):
            _alerts = []
            if _hero:
                if _hero.get('red'):
                    _alerts.append(('r', '🔴', f"{_hero['red']} vấn đề cần sửa",
                                    'phải xử lý trước khi nộp hồ sơ công an'))
                if _hero.get('visa_watch'):
                    _alerts.append(('a', '🛂', f"{_hero['visa_watch']} khách sắp hết hạn tạm trú",
                                    'kiểm tra ở Xử lý hàng ngày'))
                if _hero.get('yellow'):
                    _alerts.append(('a', '🟡', f"{_hero['yellow']} vấn đề nên kiểm tra",
                                    'không chặn nộp hồ sơ'))
                if _hero.get('invalid_ids'):
                    _alerts.append(('a', '🪪', f"{_hero['invalid_ids']} khách chưa có hộ chiếu thật",
                                    'đã để trống số giấy tờ trong KBTT'))
                if _hero.get('unknown_nats'):
                    _alerts.append(('a', '🌐', f"{_hero['unknown_nats']} quốc tịch chưa có mã",
                                    'đã giữ nguyên chữ gốc, cần kiểm tra'))
            st.markdown(f'<div class="pan-h"><span class="pan-t">⚠️ Cảnh báo</span>'
                        f'<span class="tan-chip">{len(_alerts)}</span></div>', unsafe_allow_html=True)
            if _alerts:
                st.markdown(''.join(
                    f'<div class="pan-al {_k}"><div class="pan-al-ic">{_ic}</div>'
                    f'<div><div class="pan-al-t">{_t}</div><div class="pan-al-m">{_m}</div></div></div>'
                    for _k, _ic, _t, _m in _alerts), unsafe_allow_html=True)
            elif _hero:
                st.markdown('<div class="pan-empty">✅ Không có cảnh báo nào — dữ liệu hôm nay sạch.</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div class="pan-empty">Cảnh báo sẽ hiện sau khi chạy '
                            '<b>Xử lý hàng ngày</b>.</div>', unsafe_allow_html=True)

    with _r2c:
        with st.container(border=True, height=420, key='dash_ho'):
            st.markdown('<div class="pan-h"><span class="pan-t">Bàn giao gần nhất</span></div>',
                        unsafe_allow_html=True)
            _recent = []
            if _ho_rows is not None and len(_ho_rows):
                for _, _r in _ho_rows.head(5).iterrows():
                    _room = f" · P.{_r['room']}" if _r['room'] else ""
                    _recent.append((_r['entry_time'], f"{_r['category']}{_room}"))
            elif not db_available():
                for _e in reversed((st.session_state.get('handover') or {}).get('entries', [])[-5:]):
                    _room = f" · P.{_e['room']}" if _e.get('room') else ""
                    _recent.append((_e['time'], f"{_e['cat']}{_room}"))
            if _recent:
                st.markdown(''.join(
                    f'<div class="pan-line"><span>🏷 {_t}</span><span class="pan-mut">{_tm}</span></div>'
                    for _tm, _t in _recent), unsafe_allow_html=True)
            else:
                st.markdown('<div class="pan-empty">Chưa có ghi chú nào trong ca này.</div>',
                            unsafe_allow_html=True)
            st.button("Mở sổ giao ca →", key="dash_go_ho", use_container_width=True,
                      on_click=go_menu, args=("handover",))

    st.write("")

if st.session_state.menu == "dashboard":
    _render_dashboard()

# ── Daily processing screen ───────────────────────────────────────────────
@st.fragment
def _render_daily():
    st.write("")
    st.markdown('<div class="section-label">⚙️ Cài đặt</div>', unsafe_allow_html=True)
    with st.container(border=True):
        col1, col2 = st.columns(2)
        with col1:
            if 'rate_input' not in st.session_state:
                st.session_state.rate_input = 29535.15
            def _apply_vcb_rate():
                try:
                    _rates, _ts = fetch_vcb_rates()
                    st.session_state.rate_input = _rates['USD']
                    _eur = f" · EUR {_rates['EUR']:,.2f}" if _rates.get('EUR') else ''
                    st.session_state['vcb_note'] = (
                        f"✅ Tỷ giá chuyển khoản VCB lúc {_ts}: USD {_rates['USD']:,.2f}{_eur} "
                        f"— đã điền USD vào ô tỷ giá.")
                except Exception as _e:
                    st.session_state['vcb_note'] = (
                        f"⚠️ Không lấy được tỷ giá VCB ({_e}) — nhập tay như bình thường.")
            rate = st.number_input("💱 Tỷ giá USD/EUR → VNĐ", step=0.01, format="%.2f", key="rate_input")
            st.button("🔄 Lấy tỷ giá VCB (chuyển khoản)", on_click=_apply_vcb_rate, key="btn_vcb")
            if st.session_state.get('vcb_note'):
                st.caption(st.session_state['vcb_note'])
        with col2:
            today = today_vn()
            date_str = st.text_input("📅 Ngày (dùng cho tên file)", value=f"{today.day}_{today.month:02d}")

    st.write("")
    st.markdown('<div class="section-label">🚔 Cài đặt ĐK14</div>', unsafe_allow_html=True)
    with st.container(border=True):
        col_n, col_c = st.columns(2)
        with col_n:
            dk_notifier = st.text_input(
                "Người thông báo lưu trú", value="Đỗ Duy Tân", key="dk_notifier",
                help="Điền vào cột (11) khi file nguồn không có sẵn tên. Đổi được tuỳ ý.")
        with col_c:
            dk_checkin = st.date_input(
                "Ngày check-in mặc định", value=today, key="dk_checkin", format="DD/MM/YYYY",
                help="Dùng cho ô 'Từ ngày' ở đầu sổ, và cho khách thiếu ngày đến trong file nguồn.")

    st.write("")
    st.markdown('<div class="section-label">📂 Tải file lên</div>', unsafe_allow_html=True)
    with st.container(border=True):
        col_x, col_s = st.columns(2)
        with col_x:
            xlsx_file = st.file_uploader("File XLSX — Dữ liệu khách (bắt buộc)", type=['xlsx'], key="daily_xlsx")
        with col_s:
            xls_file = st.file_uploader("File IH — Nguồn ĐK14 (tùy chọn)", type=['xls', 'xlsx'], key="daily_xls")

        visa_file = st.file_uploader(
            "File Visa — dữ liệu thô date visa (tùy chọn, để tự điền cột 'Thời hạn tạm trú tại VN' trong KBTT)",
            type=['xlsx'], key="daily_visa")

    st.write("")

    if st.button("⚡ Bắt đầu xử lý", type="primary", disabled=(xlsx_file is None and xls_file is None), use_container_width=True):
        with st.spinner("Đang xử lý..."):
            try:
                progress = st.progress(0, text="Bắt đầu...")
                zip_buf = io.BytesIO()
                files_made = []
                has_xlsx = xlsx_file is not None
                has_dk14 = False
                conv = 0; gks_cnt = 0; gbl_cnt = 0
                df = None; df_intl = None; df_vn = None
                visa_map = {}; visa_unmatched = []; visa_source = None
                kbtt_invalid_ids = []; vnm_ward_unmatched = []; dk14_skipped = []; cnt_report = []
                # Đọc file visa (nếu có) → map tên → date visa
                if visa_file is not None:
                    try:
                        visa_map = parse_visa_file(visa_file.read())
                    except Exception as _ve:
                        st.warning(f"⚠️ Không đọc được file visa: {_ve}")

                out_files = {}   # tên file → bytes: dùng cho ZIP + nút tải từng file riêng
                _issues = None; _visa_watch = []
                with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    # ── Xử lý file XLSX (nếu có) ──
                    if has_xlsx:
                        progress.progress(10, text="Quy đổi tỷ giá...")
                        xlsx_bytes = xlsx_file.read()
                        wb, conv = process_xlsx(xlsx_bytes, rate)

                        # Chia giá phòng connecting SAU khi đã quy đổi tỷ giá — ĐƠN GIÁ
                        # nguồn có thể vẫn ở dạng ngoại tệ nhỏ (VD "278.35" EUR) lúc này
                        # còn là số NGUYÊN chưa quy đổi; chia trước rồi mới quy đổi từng
                        # nửa riêng lẻ sẽ làm tròn 2 lần → sai số hàng nghìn đồng so với
                        # quy đổi trước rồi mới chia (đã xảy ra thực tế: 278.35 EUR bị cắt
                        # về 278, chia 139/139, mỗi bên quy đổi ra ~4.150.550đ — lệch hẳn
                        # so với quy đổi cả 278.35 rồi mới chia đúng 8.313.043 → chia đôi).
                        progress.progress(15, text="Chia giá phòng connecting...")
                        wb_bytes_conv = wb_to_bytes(wb)
                        wb_bytes_conv, cnt_report = split_connecting_room_prices(wb_bytes_conv)
                        wb = load_workbook(io.BytesIO(wb_bytes_conv))
                        # Ép các cột dạng mã đọc bằng chuỗi — nếu không, pandas tự suy
                        # diễn cột "trông giống số" thành float, làm mất số 0 đứng đầu
                        # (vd số điện thoại "0912345678" → 912345678.0). An toàn kể cả
                        # khi cột không tồn tại trong file (bỏ qua, không lỗi).
                        _id_cols = {c: str for c in
                                   ('SỐ GIẤY TỜ', 'SỐ ĐIỆN THOẠI', 'SỐ PHÒNG', 'MÃ CHECKIN')}
                        df = pd.read_excel(io.BytesIO(xlsx_bytes), dtype=_id_cols)
                        df_intl = df[df['LOẠI KHÁCH']=='Quốc tế'].reset_index(drop=True)
                        df_vn   = df[df['LOẠI KHÁCH']=='Việt Nam'].reset_index(drop=True)

                        progress.progress(25, text="Kiểm tra chất lượng dữ liệu...")
                        _issues = validate_guests(df)
                        _visa_watch = check_visa_expiry(df_intl, visa_map=visa_map)

                        progress.progress(35, text="Tách file Quốc tế / Việt Nam...")
                        wb_intl = split_wb(wb, 'Quốc tế')
                        wb_vn   = split_wb(wb, 'Việt Nam')

                        progress.progress(50, text="Điền mẫu KBTT...")
                        wb_kbtt, visa_unmatched, visa_source, kbtt_invalid_ids = build_kbtt(df_intl, visa_map=visa_map)

                        progress.progress(65, text="Điền mẫu Thông báo lưu trú VNM...")
                        wb_vnm, gks_cnt, gbl_cnt, vnm_ward_unmatched = build_vnm(df_vn)

                        out_files[f'converted_{date_str}.xlsx']      = wb_to_bytes(wb)
                        out_files[f'KhachQuocTe_{date_str}.xlsx']    = wb_to_bytes(wb_intl)
                        out_files[f'KhachVietNam_{date_str}.xlsx']   = wb_to_bytes(wb_vn)
                        out_files[f'ho_so_KBTT_NNN_{date_str}.xlsx'] = wb_to_bytes(wb_kbtt)
                        out_files[f'thong_bao_luu_tru_VNM_{date_str}.xlsx'] = wb_to_bytes(wb_vnm)
                        files_made += ["📄 converted (file chung)", "🌍 KhachQuocTe", "🇻🇳 KhachVietNam",
                                       "📝 KBTT NNN", "📑 Thông báo lưu trú VNM"]

                    # ── Xử lý file ĐK14 (độc lập, chỉ cần file nguồn IH) ──
                    if xls_file:
                        progress.progress(85, text="Điền mẫu ĐK14...")
                        xls_bytes = xls_file.read()
                        (dk14_bytes, dk_count, dk14_skipped,
                         dk14_issues, dk14_map) = build_dk14(xls_bytes, dk_notifier, dk_checkin)
                        out_files[f'dk14_{date_str}.xlsx'] = dk14_bytes
                        has_dk14 = True
                        files_made.append("🚔 ĐK14")

                    for _ofn, _ofb in out_files.items():
                        zf.writestr(_ofn, _ofb)

                progress.progress(100, text="Hoàn tất!")
                progress.empty()

                # Lưu kết quả vào session — kết quả & nút tải không biến mất sau rerun
                _daily = {'files_made': files_made, 'zip': zip_buf.getvalue(),
                          'files': out_files, 'rate': rate,
                          'date_str': date_str, 'has_xlsx': has_xlsx, 'has_dk14': has_dk14,
                          'dk14_count': dk_count if has_dk14 else None, 'dk14_skipped': dk14_skipped,
                          'dk14_issues': dk14_issues if has_dk14 else [],
                          'dk14_map': dk14_map if has_dk14 else None,
                          'cnt_report': cnt_report}
                if has_xlsx:
                    unknown_nats = []
                    for q in df_intl.get('QUỐC TỊCH', pd.Series([], dtype=str)).dropna().unique():
                        mapped = lookup_nat_kbtt(q)
                        if not _re.match(r'^[A-Z]{2,3} - ', str(mapped)):
                            unknown_nats.append(str(q))
                    # Thống kê phục vụ báo cáo ngày cho quản lý
                    _nat_top = (df.get('QUỐC TỊCH', pd.Series(dtype=str)).dropna().astype(str)
                                .str.strip().value_counts().head(10))
                    _arr_s = pd.to_datetime(df.get('NGÀY ĐẾN'), dayfirst=True, errors='coerce')
                    _dep_col_d = 'NGÀY ÐI' if 'NGÀY ÐI' in df.columns else ('NGÀY ĐI' if 'NGÀY ĐI' in df.columns else None)
                    _avg_nights = None
                    if _dep_col_d:
                        _dep_s = pd.to_datetime(df[_dep_col_d], dayfirst=True, errors='coerce')
                        _nvals = (_dep_s - _arr_s).dt.days.dropna()
                        _nvals = _nvals[_nvals > 0]
                        if len(_nvals):
                            _avg_nights = round(float(_nvals.mean()), 1)
                    _rooms_cnt = int(df.get('SỐ PHÒNG', pd.Series(dtype=str)).dropna().astype(str)
                                     .str.strip().nunique())
                    # Khách check-in / check-out HÔM NAY — đếm theo ngày đến/đi
                    # trùng ngày hiện tại (giờ Việt Nam), phục vụ thẻ tổng quan.
                    _today_ts = pd.Timestamp(today_vn())
                    _checkin_n = int((_arr_s.dt.normalize() == _today_ts).sum()) if _arr_s is not None else 0
                    _checkout_n = 0
                    if _dep_col_d:
                        _checkout_n = int((pd.to_datetime(df[_dep_col_d], dayfirst=True, errors='coerce')
                                           .dt.normalize() == _today_ts).sum())
                    _daily.update({'issues': _issues, 'visa_watch': _visa_watch,
                                   'nat_top': [(str(k), int(v)) for k, v in _nat_top.items()],
                                   'avg_nights': _avg_nights, 'rooms_cnt': _rooms_cnt or None,
                                   'checkin_n': _checkin_n, 'checkout_n': _checkout_n})
                    _daily.update({'total': len(df), 'intl': len(df_intl), 'vn': len(df_vn),
                                   'gks': gks_cnt, 'gbl': gbl_cnt, 'conv': conv,
                                   'unknown_nats': unknown_nats,
                                   'visa_used': visa_source is not None,
                                   'visa_source': visa_source,
                                   'visa_matched': len(df_intl) - len(visa_unmatched) if visa_source else 0,
                                   'visa_unmatched': visa_unmatched,
                                   'visa_skipped_vn': visa_map.get('skipped_vn', 0) if isinstance(visa_map, dict) else 0,
                                   'kbtt_invalid_ids': kbtt_invalid_ids,
                                   'vnm_ward_unmatched': vnm_ward_unmatched})
                st.session_state['daily_results'] = _daily

                def _mark_daily_done(state, _daily=_daily, has_xlsx=has_xlsx, has_dk14=has_dk14):
                    task = {'done': True, 'time': now_vn().strftime('%H:%M:%S'),
                            'has_xlsx': has_xlsx, 'has_dk14': has_dk14}
                    if has_xlsx:
                        iss = _daily.get('issues')
                        task['summary'] = {
                            'total': _daily.get('total'), 'intl': _daily.get('intl'), 'vn': _daily.get('vn'),
                            'gks': _daily.get('gks'), 'gbl': _daily.get('gbl'), 'conv': _daily.get('conv'),
                            'red_issues': int((iss['Mức độ'] == '🔴').sum()) if iss is not None and len(iss) else 0,
                            'yellow_issues': int((iss['Mức độ'] == '🟡').sum()) if iss is not None and len(iss) else 0,
                            'visa_watch_count': len(_daily.get('visa_watch') or []),
                            'rooms_cnt': _daily.get('rooms_cnt'),
                            'checkin_n': _daily.get('checkin_n'), 'checkout_n': _daily.get('checkout_n'),
                        }
                    state.setdefault('tasks', {})['daily'] = task
                _progress_update(_mark_daily_done)
            except Exception as e:
                st.session_state.pop('daily_results', None)
                st.error(f"❌ Lỗi: {e}")
                st.exception(e)

    _dr = st.session_state.get('daily_results')
    if _dr:
        st.success("✅ Xử lý hoàn tất!")
        if _dr['has_xlsx']:
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Tổng khách", _dr['total'])
            c2.metric("Quốc tế", _dr['intl'])
            c3.metric("Việt Nam", _dr['vn'])
            c4.metric("GKS + GBL", f"{_dr['gks']} + {_dr['gbl']}")
            st.info(f"💱 Đã quy đổi tỷ giá cho **{_dr['conv']}** ô (tỷ giá {_dr.get('rate', 0):,.2f})")
            _cnt_rep = _dr.get('cnt_report') or []
            if _cnt_rep:
                with st.expander(f"🚪 Đã tự động chia giá {len(_cnt_rep)} cặp phòng connecting", expanded=False):
                    st.caption("Phòng connecting chung mã checkin nhưng chỉ 1 phòng có giá — đã chia đều "
                               "(lệch 1 đồng thì phòng số nhỏ hơn nhận phần làm tròn xuống).")
                    st.dataframe(pd.DataFrame([
                        {'Mã checkin': r['checkin'], 'Phòng A': r['room_a'], 'Giá A (mới)': r['split_a'],
                         'Phòng B': r['room_b'], 'Giá B (mới)': r['split_b'],
                         'Tổng gốc': r['total']} for r in _cnt_rep]),
                        use_container_width=True, hide_index=True)
            if _dr['unknown_nats']:
                st.warning("⚠️ Quốc tịch chưa có mã (giữ nguyên tên, cần kiểm tra): " + ", ".join(_dr['unknown_nats']))
            if _dr.get('visa_used'):
                st.info(f"🛂 Đã điền date visa cho **{_dr['visa_matched']}/{_dr['intl']}** khách quốc tế "
                        f"(khớp theo hộ chiếu/tên từ file Visa rời bạn đã upload).")
                if _dr.get('visa_skipped_vn'):
                    st.caption(f"ℹ️ Đã tự động bỏ qua {_dr['visa_skipped_vn']} khách Việt Nam trong file visa (không cần thời hạn tạm trú).")
                if _dr.get('visa_unmatched'):
                    st.warning("⚠️ Không tìm thấy date visa cho (cột tạm trú để trống): "
                               + ", ".join(_dr['visa_unmatched']))
            if _dr.get('kbtt_invalid_ids'):
                st.warning("⚠️ Số giấy tờ chỉ là mã tạm (chưa có hộ chiếu thật) — đã để trống trong hồ sơ KBTT, "
                           "cần bổ sung số hộ chiếu thật trước khi nộp cho: " + ", ".join(_dr['kbtt_invalid_ids']))
            if _dr.get('vnm_ward_unmatched'):
                st.warning("⚠️ Không tự tra được mã Phường/Xã (đã giữ nguyên chữ gốc trong file VNM, "
                           "cần chọn lại theo danh mục): "
                           + ", ".join(f"{n} ('{p}')" for n, p in _dr['vnm_ward_unmatched']))

            # ── Kiểm tra chất lượng dữ liệu trước khi nộp hồ sơ ──
            _iss = _dr.get('issues')
            if _iss is not None and len(_iss):
                _n_red = int((_iss['Mức độ'] == '🔴').sum())
                _n_yel = int((_iss['Mức độ'] == '🟡').sum())
                st.write("")
                st.markdown('<div class="section-label">✅ Kiểm tra dữ liệu trước khi nộp hồ sơ</div>',
                           unsafe_allow_html=True)
                if _n_red:
                    st.error(f"🔴 **{_n_red}** vấn đề cần sửa trước khi nộp hồ sơ công an "
                             f"+ 🟡 **{_n_yel}** vấn đề nên kiểm tra lại.")
                else:
                    st.warning(f"🟡 **{_n_yel}** vấn đề nên kiểm tra lại (không chặn nộp hồ sơ).")
                st.dataframe(_iss, use_container_width=True, hide_index=True)
            elif _iss is not None:
                st.success("✅ Kiểm tra dữ liệu: không phát hiện vấn đề nào.")

            # ── Cảnh báo visa/tạm trú sắp hết hạn ──
            _vw = _dr.get('visa_watch') or []
            if _vw:
                st.write("")
                st.markdown('<div class="section-label">🛂 Cảnh báo hạn tạm trú / visa</div>', unsafe_allow_html=True)
                _vw_days = st.slider("Cảnh báo khách còn lưu trú mà visa hết hạn trong vòng (ngày)",
                                     1, 30, 3, key="visa_warn_days")
                _cutoff = today_vn() + datetime.timedelta(days=_vw_days)
                _soon = []
                for _v in _vw:
                    _vd = datetime.date.fromisoformat(_v['visa'])
                    _dep = datetime.date.fromisoformat(_v['dep']) if _v.get('dep') else None
                    # chỉ cảnh báo khách còn ở (chưa checkout trước ngày visa hết hạn)
                    if _vd <= _cutoff and (_dep is None or _dep >= today_vn()):
                        _soon.append({'Phòng': _v['room'], 'Họ tên': _v['name'], 'Quốc tịch': _v['nat'],
                                     'Hết hạn tạm trú': _vd.strftime('%d/%m/%Y'),
                                     'Còn': (_vd - today_vn()).days,
                                     'Ngày đi dự kiến': (_dep.strftime('%d/%m/%Y') if _dep else '—')})
                if _soon:
                    _df_soon = pd.DataFrame(_soon).sort_values('Còn')
                    st.error(f"🔴 **{len(_soon)}** khách cần gia hạn/rời đi trước khi tạm trú hết hạn:")
                    st.dataframe(_df_soon, use_container_width=True, hide_index=True)
                else:
                    st.success(f"✅ Không có khách nào hết hạn tạm trú trong {_vw_days} ngày tới.")
        elif _dr['has_dk14']:
            st.info("ℹ️ Chỉ tạo file ĐK14 (không có file XLSX dữ liệu khách).")

        if _dr.get('has_dk14'):
            _dk_iss = _dr.get('dk14_issues') or []
            _n_skip = len(_dr.get('dk14_skipped') or [])
            _n_fix = sum(1 for t, *_ in _dk_iss if t == 'fix')
            _n_warn = sum(1 for t, *_ in _dk_iss if t == 'warn')
            _m1, _m2, _m3, _m4 = st.columns(4)
            _m1.metric("Tổng khách vào sổ", _dr['dk14_count'])
            _m2.metric("Đã tự sửa", _n_fix)
            _m3.metric("Cần xem lại", _n_warn)
            _m4.metric("Bỏ qua", _n_skip)

            if _dr.get('dk14_map'):
                _dcols, _dheads = _dr['dk14_map']
                _LBL = [('name', 'B — Họ và tên'), ('dob', 'C/D — Ngày sinh (tách Nam/Nữ)'),
                        ('gender', '(dùng để tách ngày sinh Nam/Nữ)'), ('nationality', 'E — Quốc tịch'),
                        ('country', 'E — Quốc tịch (dự phòng)'), ('id', 'F — Số giấy tờ'),
                        ('cuTru', 'G — Loại cư trú'), ('tinh', 'G — Tỉnh'), ('quan', 'G — Quận'),
                        ('phuong', 'G — Phường'), ('addressDetail', 'G — Địa chỉ chi tiết'),
                        ('arrival', 'H — Thời gian đến'), ('departure', 'I — Thời gian đi'),
                        ('room', 'J — Số phòng'), ('notifier', 'K — Người thông báo')]
                with st.expander("🔎 Cột đã tự nhận diện trong file nguồn", expanded=False):
                    for _k, _lb in _LBL:
                        _i = _dcols.get(_k)
                        if _i is None:
                            continue
                        _h = str(_dheads[_i]).strip() if _i < len(_dheads) and _dheads[_i] is not None else '(?)'
                        st.markdown(f"- cột **{chr(65 + _i)}** “{_h}” → ĐK14 **{_lb}**")
                    _miss = [k for k in ('name', 'dob', 'gender', 'id', 'arrival', 'departure', 'room')
                             if k not in _dcols]
                    if _miss:
                        st.warning("⚠️ Không tìm thấy cột: " + ", ".join(_miss))

            if _dk_iss:
                _ICON = {'fix': '🔧', 'warn': '⚠️', 'skip': '⏭️'}
                with st.expander(f"📋 Nhật ký {len(_dk_iss)} dòng đã sửa / cần xem lại", expanded=bool(_n_warn)):
                    st.dataframe(pd.DataFrame(
                        [{'': _ICON.get(t, ''), 'Dòng': r, 'Tên': n, 'Nội dung': msg}
                         for t, r, n, msg in _dk_iss]),
                        use_container_width=True, hide_index=True)
            else:
                st.success("✅ ĐK14: không phát hiện vấn đề nào trong dữ liệu nguồn.")

        st.markdown("**File đã tạo:** " + " · ".join(_dr['files_made']))

        st.download_button(
            label="⬇️ Tải về tất cả file (ZIP)",
            data=_dr['zip'],
            file_name=f"hotel_{_dr['date_str']}.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary"
        )

        # ── Tải riêng từng file (không cần giải nén ZIP) ──
        if _dr.get('files'):
            with st.expander("⬇️ Tải riêng từng file"):
                for _fn, _fb in _dr['files'].items():
                    st.download_button(f"⬇️ {_fn}", _fb, file_name=_fn,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key=f"dl_single_{_fn}")

if st.session_state.menu == "daily":
    _render_daily()

# ── Regcard screen ────────────────────────────────────────────────────────
@st.fragment
def _render_regcard():
    st.write("")
    st.markdown('<div class="section-label">🖨️ Tạo Registration Card + file ARR</div>', unsafe_allow_html=True)
    st.caption("Điền dữ liệu từ file Arrival Smile lên mẫu Regcard PDF gốc, đồng thời tạo file ARR "
               "(đúng định dạng ARR Converter — nhóm Conf# · đếm phòng · Cà Thẻ / Thu Tiền / Xem Lại BU / FOC Late C/O) — ra cùng lúc 2 file.")

    with st.container(border=True):
        rc_file = st.file_uploader("File Excel dữ liệu booking (.xlsx)", type=['xlsx'], key="rc_xlsx")

        only_main = st.checkbox("Chỉ tạo cho khách chính (có mã Conf#)", value=True,
                                help="Bỏ chọn để tạo regcard cho tất cả khách, kể cả khách đi cùng phòng")

    st.write("")

    if st.button("🖨️ Tạo Regcard PDF + file ARR", type="primary", disabled=rc_file is None, use_container_width=True):
        with st.spinner("Đang tạo Regcard PDF và file ARR..."):
            try:
                rc_bytes = rc_file.read()
                pdf_data, count = build_regcards(rc_bytes, only_main=only_main)

                # Tạo file ARR từ cùng file đầu vào (lỗi ARR không làm hỏng PDF)
                arr_bytes, arr_stats, arr_err = None, None, None
                try:
                    wb_arr, arr_stats = build_arr(rc_bytes)
                    arr_bytes = wb_to_bytes(wb_arr)
                except Exception as _e:
                    arr_err = str(_e)

                st.session_state['rc_results'] = {
                    'pdf': pdf_data, 'count': count,
                    'arr': arr_bytes, 'arr_stats': arr_stats, 'arr_err': arr_err,
                    'date': today_vn().strftime('%d_%m'),
                    'arr_date': today_vn().strftime('%d.%m.%Y'),
                }

                def _mark_regcard_done(state, count=count, arr_stats=arr_stats):
                    task = {'done': True, 'time': now_vn().strftime('%H:%M:%S'), 'regcards': count}
                    if arr_stats:
                        task['summary'] = {k: arr_stats.get(k) for k in
                                           ('bookings', 'rooms', 'ota', 'ca_the', 'thu_tien', 'xem_lai_bu', 'foc_lco')}
                    state.setdefault('tasks', {})['regcard'] = task
                _progress_update(_mark_regcard_done)
            except Exception as e:
                st.session_state.pop('rc_results', None)
                st.error(f"❌ Lỗi: {e}")
                st.exception(e)

    # Kết quả lưu trong session — 2 nút tải không biến mất sau khi bấm 1 nút
    _res = st.session_state.get('rc_results')
    if _res:
        if _res['count'] == 0:
            st.warning("⚠️ Không tìm thấy khách nào để tạo regcard. Kiểm tra lại file.")
        else:
            st.success(f"✅ Đã tạo {_res['count']} regcard"
                       + (" + file ARR!" if _res['arr'] else "!"))
            if _res['arr_stats']:
                a1, a2, a3, a4, a5 = st.columns(5)
                a1.metric("🖨️ Regcard", _res['count'])
                a2.metric("📦 Booking", _res['arr_stats']['bookings'])
                a3.metric("💳 Cà Thẻ", _res['arr_stats']['ca_the'])
                a4.metric("💵 Thu Tiền", _res['arr_stats']['thu_tien'])
                a5.metric("⚠️ Xem Lại BU", _res['arr_stats']['xem_lai_bu'])
                if _res['arr_stats'].get('foc_lco'):
                    st.caption(f"🛎️ FOC Late C/O: {_res['arr_stats']['foc_lco']} booking")
            else:
                st.metric("Số regcard", _res['count'])

            d1, d2 = st.columns(2)
            with d1:
                st.download_button(
                    label=f"⬇️ Tải {_res['count']} Regcard (PDF)",
                    data=_res['pdf'],
                    file_name=f"regcards_{_res['date']}.pdf",
                    mime="application/pdf",
                    use_container_width=True, type="primary", key="dl_rc_pdf")
            with d2:
                if _res['arr']:
                    st.download_button(
                        label="⬇️ Tải file ARR (Excel)",
                        data=_res['arr'],
                        file_name=f"Arr {_res['arr_date']}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, type="primary", key="dl_rc_arr")
            if _res['arr_err']:
                st.warning(f"⚠️ Không tạo được file ARR: {_res['arr_err']}")

if st.session_state.menu == "regcard":
    _render_regcard()

# ── Sổ giao ca ─────────────────────────────────────────────────────────────
@st.fragment
def _render_handover():
    st.write("")
    st.markdown('<div class="section-label">🤝 Sổ giao ca</div>', unsafe_allow_html=True)

    _db_on = db_available()

    if _db_on:
        st.caption("Ghi chú trong ca (khách nợ, yêu cầu đặc biệt, sự cố, đồ thất lạc...) để ca sau nắm được. "
                   "Lưu trên **đám mây (Supabase)** — không mất khi app deploy lại, xem lại được mọi ngày trong quá khứ.")

        with st.container(border=True):
            hc1, hc2, hc3 = st.columns([1, 1, 1])
            with hc1:
                h_shift = st.selectbox("Ca trực", ["Ca sáng", "Ca chiều", "Ca đêm"], key="h_shift")
            with hc2:
                h_staff = st.text_input("Tên lễ tân trực", key="h_staff", placeholder="VD: Tân")
            with hc3:
                _avail_dates = db_load_dates()
                _sel_date = st.date_input("📅 Xem ngày", value=today_vn(),
                                          max_value=today_vn(), key="h_view_date",
                                          help="Chọn lại ngày trong quá khứ để xem lịch sử sổ giao ca")
        _is_today = _sel_date == today_vn()

        _df_e = db_load_entries(_sel_date)
        _summary = compute_day_summary(_df_e)

        st.write("")
        if not _is_today:
            st.info(f"📜 Đang xem lại lịch sử ngày **{_sel_date.strftime('%d/%m/%Y')}** (chỉ xem, "
                    "không thêm/xóa được — quay lại hôm nay để ghi chú mới).")

        st.markdown('<div class="section-label">📊 Tổng hợp tự động</div>', unsafe_allow_html=True)
        if _summary['total'] == 0:
            st.caption(f"Chưa có ghi chú nào ngày {_sel_date.strftime('%d/%m/%Y')}.")
        else:
            sc1, sc2 = st.columns([1, 2])
            with sc1:
                st.metric("Tổng ghi chú", _summary['total'])
            with sc2:
                _cat_txt = " · ".join(f"{k}: {v}" for k, v in _summary['by_category'].items())
                st.markdown(f"**Theo phân loại:** {_cat_txt}")
            if _summary['rooms']:
                st.caption("🚪 Phòng được nhắc tới: " + ", ".join(_summary['rooms']))

        if _is_today:
            st.write("")
            with st.container(border=True):
                with st.form("handover_add_form", clear_on_submit=True):
                    fc1, fc2 = st.columns([1, 1])
                    with fc1:
                        h_cat = st.selectbox("Phân loại", ["Khách nợ", "Yêu cầu đặc biệt", "Sự cố",
                                                             "Đồ thất lạc", "Bảo trì", "Khác"], key="h_cat")
                    with fc2:
                        h_room = st.text_input("Số phòng (nếu có)", key="h_room")
                    h_note = st.text_area("Nội dung bàn giao", key="h_note", height=80)
                    h_submit = st.form_submit_button("➕ Thêm vào sổ giao ca", type="primary", use_container_width=True)
                    if h_submit:
                        if not h_note.strip():
                            st.warning("⚠️ Vui lòng nhập nội dung bàn giao.")
                        else:
                            db_add_entry(today_vn(), now_vn().strftime('%H:%M'),
                                        h_cat, h_room.strip(), h_note.strip())
                            st.rerun(scope="fragment")

        st.write("")
        if not _df_e.empty:
            st.markdown(f"**{len(_df_e)} ghi chú ngày {_sel_date.strftime('%d/%m/%Y')}**")
            for _, _row in _df_e.iterrows():
                with st.container(border=True):
                    ic1, ic2 = st.columns([10, 1])
                    with ic1:
                        _room_txt = f" · 🚪 Phòng {_row['room']}" if _row['room'] else ""
                        st.markdown(f"🕐 **{_row['entry_time']}** · 🏷️ {_row['category']}{_room_txt}")
                        st.write(_row['note'])
                    with ic2:
                        if _is_today and st.button("🗑️", key=f"h_del_{_row['id']}", help="Xóa ghi chú này"):
                            db_delete_entry(_row['id'])
                            st.rerun(scope="fragment")

            st.write("")
            _entries_for_xlsx = [{'time': r['entry_time'], 'cat': r['category'],
                                  'room': r['room'] or '', 'note': r['note']}
                                 for _, r in _df_e.iloc[::-1].iterrows()]
            _wb_ho = build_handover_xlsx(
                {'date': _sel_date.strftime('%d/%m/%Y'),
                 'shift': h_shift if _is_today else '', 'staff': h_staff if _is_today else ''},
                _entries_for_xlsx)
            st.download_button(f"⬇️ Tải sổ giao ca ngày {_sel_date.strftime('%d_%m')} (Excel)", wb_to_bytes(_wb_ho),
                               file_name=f"giao_ca_{_sel_date.strftime('%d_%m_%Y')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, type="primary", key="dl_handover")
        elif _is_today:
            st.info("Chưa có ghi chú nào trong ca này. Thêm ghi chú ở form phía trên.")

    else:
        # ── Chưa cấu hình Supabase — dùng lại lưu tạm trên đĩa server (chỉ sống
        # trong ngày, mất khi deploy lại). Xem secrets.toml.example để kết nối
        # lưu trữ đám mây bền vững + xem lại lịch sử nhiều ngày.
        st.caption("Ghi chú trong ca (khách nợ, yêu cầu đặc biệt, sự cố, đồ thất lạc...) để ca sau nắm được. "
                   "Tự động lưu trên server theo ngày — mở lại/tải lại trang trong ngày vẫn còn nguyên. "
                   "Bấm **Tải file Excel** cuối trang khi cần in hoặc lưu trữ lâu dài.")
        st.info("☁️ **Chưa kết nối lưu trữ đám mây** — ghi chú chỉ lưu tạm trên server, sẽ mất khi app deploy lại "
                "và không xem lại được các ngày trước. Xem hướng dẫn kết nối Supabase trong `secrets.toml.example` "
                "để lưu trữ bền vững + xem lại lịch sử nhiều ngày.")
        _db_err = st.session_state.get('_db_last_error')
        if _db_err:
            with st.expander("🔍 Xem lý do kết nối thất bại (để tự sửa secrets)"):
                st.code(_db_err, language=None)
        if st.button("🔄 Thử kết nối lại", key="db_retry",
                    help="Bấm sau khi đã sửa Secrets — không cần Reboot cả app"):
            _db_schema_ready.clear()
            st.rerun(scope="fragment")

        with st.container(border=True):
            hc1, hc2 = st.columns(2)
            with hc1:
                h_shift = st.selectbox("Ca trực", ["Ca sáng", "Ca chiều", "Ca đêm"], key="h_shift")
            with hc2:
                h_staff = st.text_input("Tên lễ tân trực", key="h_staff", placeholder="VD: Tân")

            st.write("")
            with st.form("handover_add_form", clear_on_submit=True):
                fc1, fc2 = st.columns([1, 1])
                with fc1:
                    h_cat = st.selectbox("Phân loại", ["Khách nợ", "Yêu cầu đặc biệt", "Sự cố",
                                                         "Đồ thất lạc", "Bảo trì", "Khác"], key="h_cat")
                with fc2:
                    h_room = st.text_input("Số phòng (nếu có)", key="h_room")
                h_note = st.text_area("Nội dung bàn giao", key="h_note", height=80)
                h_submit = st.form_submit_button("➕ Thêm vào sổ giao ca", type="primary", use_container_width=True)
                if h_submit:
                    if not h_note.strip():
                        st.warning("⚠️ Vui lòng nhập nội dung bàn giao.")
                    else:
                        _new_entry = {'time': now_vn().strftime('%H:%M'),
                                      'cat': h_cat, 'room': h_room.strip(), 'note': h_note.strip()}
                        _progress_update(lambda state: state.setdefault('handover_entries', []).append(_new_entry))
                        st.session_state.handover['entries'].append(_new_entry)
                        st.rerun(scope="fragment")

        _entries = st.session_state.handover['entries']
        st.write("")
        if _entries:
            st.markdown(f"**{len(_entries)} ghi chú trong ca này**")
            for _ei, _e in enumerate(reversed(_entries)):
                _real_i = len(_entries) - 1 - _ei
                with st.container(border=True):
                    ic1, ic2 = st.columns([10, 1])
                    with ic1:
                        _room_txt = f" · 🚪 Phòng {_e['room']}" if _e['room'] else ""
                        st.markdown(f"🕐 **{_e['time']}** · 🏷️ {_e['cat']}{_room_txt}")
                        st.write(_e['note'])
                    with ic2:
                        if st.button("🗑️", key=f"h_del_{_real_i}", help="Xóa ghi chú này"):
                            _target = _e
                            def _m(state, _target=_target):
                                entries = state.setdefault('handover_entries', [])
                                if _target in entries:
                                    entries.remove(_target)
                            _progress_update(_m)
                            st.session_state.handover['entries'] = list(
                                st.session_state.progress.get('handover_entries', []))
                            st.rerun(scope="fragment")

            st.write("")
            _wb_ho = build_handover_xlsx(
                {'date': today_vn().strftime('%d/%m/%Y'), 'shift': h_shift, 'staff': h_staff},
                _entries)
            st.download_button("⬇️ Tải sổ giao ca (Excel)", wb_to_bytes(_wb_ho),
                               file_name=f"giao_ca_{today_vn().strftime('%d_%m')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, type="primary", key="dl_handover")
        else:
            st.info("Chưa có ghi chú nào trong ca này. Thêm ghi chú ở form phía trên.")

if st.session_state.menu == "handover":
    _render_handover()

# ── Đối chiếu: sub-menu 2 lựa chọn (có cổng mật khẩu riêng) ────────────────
try:
    RECON_PASS = st.secrets.get("recon_pass", "368736")
except Exception:
    RECON_PASS = "368736"
if "recon_ok" not in st.session_state:
    st.session_state.recon_ok = False

def _check_recon():
    if st.session_state.get("recon_pass_input", "") == RECON_PASS:
        st.session_state.recon_ok = True
        st.session_state.recon_pass_err = False
    else:
        st.session_state.recon_pass_err = True

@st.fragment
def _render_recon():
    st.write("")

    # Cổng mật khẩu cho tính năng đối chiếu
    if not st.session_state.recon_ok:
        st.markdown('<div class="section-label">🔒 Nhập mật khẩu để truy cập</div>', unsafe_allow_html=True)
        st.caption("Tính năng Đối chiếu lưu trú được bảo vệ bằng mật khẩu riêng.")
        with st.form("recon_pass_form", clear_on_submit=False):
            st.text_input("Mật khẩu", key="recon_pass_input", type="password", placeholder="Nhập mật khẩu")
            ok = st.form_submit_button("Mở khóa →", type="primary", use_container_width=True)
            if ok:
                _check_recon()
                if st.session_state.recon_ok:
                    st.session_state.menu = "recon_person"
                    st.rerun()
        if st.session_state.get("recon_pass_err"):
            st.error("❌ Mật khẩu không đúng!")
        st.stop()

    st.info("Đã mở khóa — chọn công cụ Đối chiếu ở sidebar bên trái.")

if st.session_state.menu == "recon":
    _render_recon()

# ── Kiểm tra lưu trú người nước ngoài ─────────────────────────────────────
@st.fragment
def _render_recon_person():
    st.write("")
    st.markdown('<div class="section-label">🌏 Kiểm tra lưu trú người nước ngoài</div>', unsafe_allow_html=True)
    st.caption("So khớp khách Smile vs Trang quản lý người nước ngoài theo số hộ chiếu — tìm khách chưa đăng ký / đăng ký trùng.")

    rc1, rc2 = st.columns(2)
    with rc1:
        smile_file = st.file_uploader("File Smile — Inhouse (.xlsx)", type=['xlsx'], key="recon_smile")
    with rc2:
        luutru_file = st.file_uploader("File Trang lưu trú người nước ngoài (.xlsx)", type=['xlsx'], key="recon_luutru")

    today_str = st.text_input("📅 Ngày xuất file (hôm nay)", value=today_vn().strftime('%d/%m/%Y'),
                              help="Dùng để loại bỏ: khách arrival hôm nay (Smile) và khách ngày đi dự kiến hôm nay (Lưu trú)")

    st.write("")

    if st.button("🔍 Bắt đầu kiểm tra", type="primary",
                 disabled=(smile_file is None or luutru_file is None), use_container_width=True):
        with st.spinner("Đang đối chiếu..."):
            try:
                today = pd.to_datetime(today_str, format='%d/%m/%Y')
                _rp = reconcile(smile_file.read(), luutru_file.read(), today)
                st.session_state['recon_results'] = _rp

                def _mark_recon_person_done(state, _rp=_rp):
                    state.setdefault('tasks', {})['recon_person'] = {
                        'done': True, 'time': now_vn().strftime('%H:%M:%S'),
                        'summary': {'chua_dang_ky': len(_rp.get('chua_dk', [])),
                                   'thua': len(_rp.get('thua', [])), 'trung': len(_rp.get('dup', []))}}
                _progress_update(_mark_recon_person_done)
            except Exception as e:
                st.session_state.pop('recon_results', None)
                st.error(f"❌ Lỗi: {e}")
                st.exception(e)

    r = st.session_state.get('recon_results')
    if r:
        st.success("✅ Kiểm tra hoàn tất!")

        c1, c2 = st.columns(2)
        c1.metric("Smile (sau lọc)", r['smile_filtered'], f"từ {r['smile_total']} (bỏ VNM + arrival/departure hôm nay)")
        c2.metric("Lưu trú (sau lọc)", r['luutru_filtered'], f"từ {r['luutru_total']} (bỏ đi dự kiến hôm nay)")

        st.divider()
        st.markdown("### 👤 Kết quả đối chiếu người (theo số hộ chiếu)")

        n_chua = len(r['chua_dk']); n_thua = len(r['thua']); n_dup = len(r['dup'])

        # Bảng tổng hợp chênh lệch
        m1, m2, m3 = st.columns(3)
        m1.metric("🔴 Chưa đăng ký", n_chua)
        m2.metric("🟡 Có/lưu trú, thiếu/Smile", n_thua)
        m3.metric("🟠 Đăng ký trùng", n_dup)

        if n_chua == 0 and n_thua == 0 and n_dup == 0:
            st.success("✅ Khớp hoàn toàn! Không có ai thiếu/thừa/trùng.")

        if n_chua > 0:
            st.error(f"🔴 {n_chua} khách trên Smile nhưng CHƯA đăng ký lưu trú:")
            st.dataframe(r['chua_dk'], use_container_width=True, hide_index=True)
        else:
            st.success("✅ Không có khách nào chưa đăng ký lưu trú.")

        if n_thua > 0:
            st.warning(f"🟡 Chênh lệch **{n_thua} người**: có trên Trang quản lý người nước ngoài nhưng KHÔNG có trên Smile (có thể đã checkout nhưng chưa xóa khỏi lưu trú):")
            st.dataframe(r['thua'], use_container_width=True, hide_index=True)
            # Nút tải danh sách chênh lệch
            _csv = r['thua'].to_csv(index=False).encode('utf-8-sig')
            st.download_button("⬇️ Tải danh sách chênh lệch (CSV)", _csv,
                               file_name="chenh_lech_luu_tru.csv", mime="text/csv")

        if n_dup > 0:
            st.warning(f"🟠 {n_dup} dòng ĐĂNG KÝ TRÙNG trên lưu trú:")
            st.dataframe(r['dup'], use_container_width=True, hide_index=True)

if st.session_state.menu == "recon_person":
    _render_recon_person()

# ── Kiểm tra hệ thống quản lý lưu trú phòng ────────────────────────────────
def reconcile_rooms(smile_bytes, room_bytes, today):
    """Đối chiếu phòng inhouse từ file khách lưu trú Smile (trừ khách Arrival hôm nay)
    với file Excel chỉ chứa danh sách số phòng."""
    from collections import Counter

    # ── Đọc file khách lưu trú Smile ──
    df1 = pd.read_excel(io.BytesIO(smile_bytes), header=0)
    if 'Rm#' not in df1.columns:
        raise ValueError("File Smile không có cột 'Rm#'. Vui lòng dùng file khách lưu trú xuất từ Smile.")
    smile = df1.dropna(subset=['Rm#']).copy()
    smile['room'] = smile['Rm#'].apply(_norm_room)
    smile['Arrival'] = pd.to_datetime(smile['Arrival'], errors='coerce') if 'Arrival' in smile else pd.NaT
    _ln = smile['Last Name'].astype(str).str.strip() if 'Last Name' in smile else ''
    _fn = smile['First Name'].astype(str).str.strip() if 'First Name' in smile else ''
    smile['name'] = (_ln + ' ' + _fn).str.strip() if 'Last Name' in smile else ''
    smile_total = len(smile)
    # Cột Departure (dò tên linh hoạt)
    _dep_col = next((c for c in df1.columns if 'depart' in str(c).lower()), None)
    smile['Departure'] = pd.to_datetime(df1[_dep_col], errors='coerce') if _dep_col else pd.NaT
    # Trừ khách Arrival = hôm nay và khách Departure = hôm nay (trả phòng)
    if 'Arrival' in smile:
        smile_f = smile[(smile['Arrival'].dt.date != today.date()) &
                        (smile['Departure'].dt.date != today.date())].copy()
    else:
        smile_f = smile.copy()

    # ── Đọc file chỉ chứa số phòng: lấy tất cả ô có dữ liệu ──
    raw = pd.read_excel(io.BytesIO(room_bytes), header=None, dtype=str)
    rooms_sys = []
    for _, row_vals in raw.iterrows():
        for v in row_vals:
            if pd.isna(v): continue
            r = _norm_room(v)
            if not r: continue
            # bỏ ô tiêu đề nếu lỡ có (vd "Số phòng", "Room", "STT")
            if _norm_nat(r) in ('sophong', 'phong', 'room', 'rm', 'stt'): continue
            rooms_sys.append(r)
    if not rooms_sys:
        raise ValueError("File số phòng không có dữ liệu. Vui lòng kiểm tra lại file.")

    sys_rooms = set(rooms_sys)
    _cnt = Counter(rooms_sys)
    sys_dup = sorted((r for r, c in _cnt.items() if c > 1), key=lambda x: (len(x), x))

    import re as _re3
    def _is_virtual(r):
        return bool(_re3.fullmatch(r'9\d{3}', r))  # phòng ảo 9000-9999 (posting master)

    smile_rooms = set(r for r in smile_f['room'] if r and not _is_virtual(r))
    sys_rooms = set(r for r in sys_rooms if not _is_virtual(r))

    def _sortkey(x): return (len(x), x)
    room_chua = sorted(smile_rooms - sys_rooms, key=_sortkey)  # inhouse nhưng CHƯA có trong file phòng
    room_thua = sorted(sys_rooms - smile_rooms, key=_sortkey)  # có trong file phòng nhưng KHÔNG còn inhouse

    # Chi tiết khách trong các phòng chưa đăng ký (tiện đăng ký bổ sung)
    if room_chua:
        detail = smile_f[smile_f['room'].isin(room_chua)].copy()
        detail['_arr'] = detail['Arrival'].dt.strftime('%d/%m/%Y') if 'Arrival' in detail else ''
        _nat = detail['NAT'] if 'NAT' in detail else ''
        detail = pd.DataFrame({
            'Số phòng': detail['room'].values,
            'Họ tên': detail['name'].values,
            'Quốc tịch': _nat.values if hasattr(_nat, 'values') else _nat,
            'Ngày đến': detail['_arr'].values if hasattr(detail['_arr'], 'values') else '',
        }).sort_values('Số phòng', key=lambda s: s.map(lambda x: (len(x), x)))
    else:
        detail = pd.DataFrame(columns=['Số phòng', 'Họ tên', 'Quốc tịch', 'Ngày đến'])

    return {
        'smile_total': smile_total, 'smile_filtered': len(smile_f),
        'sys_total': len(rooms_sys), 'sys_unique': len(sys_rooms),
        'smile_rooms': len(smile_rooms),
        'room_chua': room_chua, 'room_thua': room_thua, 'sys_dup': sys_dup,
        'room_match': len(smile_rooms & sys_rooms),
        'detail_chua': detail,
    }


@st.fragment
def _render_recon_room():
    st.write("")
    st.markdown('<div class="section-label">🚪 Kiểm tra hệ thống quản lý lưu trú phòng</div>', unsafe_allow_html=True)
    st.caption("So khớp số phòng inhouse từ file khách lưu trú Smile với file danh sách số phòng — tìm phòng chưa đăng ký / thừa / trùng.")

    rr1, rr2 = st.columns(2)
    with rr1:
        smile_file_r = st.file_uploader("File khách lưu trú Smile (.xlsx)", type=['xlsx'], key="reconr_smile")
    with rr2:
        room_file = st.file_uploader("File số phòng (.xlsx — chỉ chứa danh sách số phòng)", type=['xlsx'], key="reconr_room")

    today_str_r = st.text_input("📅 Ngày xuất file (hôm nay)", value=today_vn().strftime('%d/%m/%Y'),
                                key="reconr_today",
                                help="Khách có Arrival = ngày này trên Smile sẽ được loại bỏ khỏi đối chiếu")

    st.write("")

    if st.button("🔍 Bắt đầu kiểm tra", type="primary", key="reconr_run",
                 disabled=(smile_file_r is None or room_file is None), use_container_width=True):
        with st.spinner("Đang đối chiếu phòng..."):
            try:
                today_r = pd.to_datetime(today_str_r, format='%d/%m/%Y')
                _rr = reconcile_rooms(smile_file_r.read(), room_file.read(), today_r)
                st.session_state['reconr_results'] = _rr

                def _mark_recon_room_done(state, _rr=_rr):
                    state.setdefault('tasks', {})['recon_room'] = {
                        'done': True, 'time': now_vn().strftime('%H:%M:%S'),
                        'summary': {'chua_dang_ky': len(_rr.get('room_chua', [])),
                                   'thua': len(_rr.get('room_thua', [])), 'trung': len(_rr.get('sys_dup', []))}}
                _progress_update(_mark_recon_room_done)
            except Exception as e:
                st.session_state.pop('reconr_results', None)
                st.error(f"❌ Lỗi: {e}")
                st.exception(e)

    rr = st.session_state.get('reconr_results')
    if rr:
        st.success("✅ Kiểm tra hoàn tất!")

        c1, c2, c3 = st.columns(3)
        c1.metric("Phòng inhouse (Smile)", rr['smile_rooms'],
                  f"{rr['smile_filtered']} khách (từ {rr['smile_total']}, đã trừ arrival hôm nay)")
        c2.metric("Phòng trong file", rr['sys_unique'],
                  (f"{rr['sys_total']} dòng" if rr['sys_total'] != rr['sys_unique'] else None))
        c3.metric("🟢 Phòng khớp", rr['room_match'])

        st.divider()
        st.markdown("### 🚪 Kết quả đối chiếu phòng")

        n_chua = len(rr['room_chua']); n_thua = len(rr['room_thua']); n_dup = len(rr['sys_dup'])

        m1, m2, m3 = st.columns(3)
        m1.metric("🔴 Chưa đăng ký", n_chua)
        m2.metric("🟡 Thừa trong file", n_thua)
        m3.metric("🟠 Trùng trong file", n_dup)

        if n_chua == 0 and n_thua == 0 and n_dup == 0:
            st.success("✅ Khớp hoàn toàn! Không có phòng thiếu/thừa/trùng.")
            st.balloons()

        if n_chua > 0:
            st.error(f"🔴 {n_chua} phòng có khách inhouse nhưng CHƯA có trong file số phòng: "
                     + ", ".join(rr['room_chua']))
            st.markdown("**Chi tiết khách trong các phòng chưa đăng ký:**")
            st.dataframe(rr['detail_chua'], use_container_width=True, hide_index=True)
            _csv_r = rr['detail_chua'].to_csv(index=False).encode('utf-8-sig')
            st.download_button("⬇️ Tải danh sách phòng chưa đăng ký (CSV)", _csv_r,
                               file_name="phong_chua_dang_ky.csv", mime="text/csv")
        else:
            st.success("✅ Tất cả phòng inhouse đều đã có trong file số phòng.")

        if n_thua > 0:
            st.warning(f"🟡 {n_thua} phòng có trong file nhưng KHÔNG còn khách inhouse "
                       f"(có thể đã checkout nhưng chưa gỡ): "
                       + ", ".join(rr['room_thua']))

        if n_dup > 0:
            st.warning(f"🟠 {n_dup} phòng bị TRÙNG (xuất hiện nhiều lần) trong file số phòng: "
                       + ", ".join(rr['sys_dup']))

if st.session_state.menu == "recon_room":
    _render_recon_room()

# ── Hiệu ứng mượt khi ĐỔI CÔNG CỤ (sidebar) ─────────────────────────────────
# Đặt Ở CUỐI file (sau mọi khối `if st.session_state.menu == ...`) — không
# phải ngẫu nhiên: đây là lệnh cuối cùng được gửi cho mỗi lượt rerun, nên khi
# JS bên dưới chạy, toàn bộ nội dung trang MỚI (đúng công cụ vừa chuyển sang)
# chắc chắn đã render xong. Đặt sớm hơn (như khối set data-theme ở trên) sẽ
# có nguy cơ JS thao tác nhầm lên nội dung TRANG CŨ vẫn còn trên DOM.
# So sánh menu hiện tại với menu LƯU TRÊN CHÍNH TRANG CHA (không phải session
# JS nội bộ iframe — iframe bị tạo mới mỗi lượt rerun, không nhớ được gì) để
# chỉ phát hiệu ứng khi THỰC SỰ đổi công cụ, không phát khi rerun vì lý do
# khác (gõ chữ, tick ô, mở expander...) — tránh nhấp nháy phiền mỗi thao tác.
st.iframe(f"""
<script>
(function(){{
    var doc = window.parent.document;
    var cur = {st.session_state.menu!r};
    var prev = doc.documentElement.getAttribute('data-tan-menu');
    if (prev !== null && prev !== cur) {{
        var main = doc.querySelector('[data-testid="stMainBlockContainer"]');
        if (main) {{
            main.classList.remove('tan-page-in');
            void main.offsetWidth;
            main.classList.add('tan-page-in');
        }}
    }}
    doc.documentElement.setAttribute('data-tan-menu', cur);
}})();
</script>
""", height=1)


