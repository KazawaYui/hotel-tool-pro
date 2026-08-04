import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from copy import copy
import xlrd, datetime, io, zipfile, base64, os
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.colors import white, black

# Load app icon (favicon) từ icon.b64
@st.cache_resource
def _load_app_icon():
    try:
        from PIL import Image
        p = os.path.join(os.path.dirname(__file__), 'icon.b64')
        with open(p, 'r') as f:
            return Image.open(io.BytesIO(base64.b64decode(f.read())))
    except Exception:
        return "🌸"

st.set_page_config(page_title="Tân Hotel", page_icon=_load_app_icon(), layout="wide")

# ── Load embedded templates ──────────────────────────────────────────────
@st.cache_resource
def load_template(name):
    path = os.path.join(os.path.dirname(__file__), f'tmpl_{name}.b64')
    with open(path, 'r') as f:
        return base64.b64decode(f.read())

# ── Lookup tables ─────────────────────────────────────────────────────────
# Full nationality mapping (normalized keys → "CODE - Name") — 350 entries
import unicodedata as _ud, re as _re
def _norm_nat(s):
    s = str(s).lower().strip()
    s = _ud.normalize('NFD', s)
    s = ''.join(c for c in s if _ud.category(c) != 'Mn')
    return _re.sub(r'[^a-z0-9]', '', s)

NAT_NORM = {
    "achentina": "ARG - Argentina",
    "acmenia": "ARM - Armenia",
    "adecbaigian": "AZE - Azerbaijan",
    "aicap": "EGY - Egypt",
    "airolen": "IRL - Ireland",
    "aixolen": "ISL - Iceland",
    "albania": "ALB - Albania",
    "anbani": "ALB - Albania",
    "andorra": "AND - Andorra",
    "anggola": "AGO - Angola",
    "angola": "AGO - Angola",
    "anh": "GBR - United Kingdom",
    "anmach": "DNK - Denmark",
    "ano": "IND - India",
    "ao": "AUT - Austria",
    "aosip": "CYP - Cyprus",
    "arapthongnhat": "ARE - United Arab Emirates",
    "arapxaui": "SAU - Saudi Arabia",
    "argentina": "ARG - Argentina",
    "armenia": "ARM - Armenia",
    "australia": "AUS - Australia",
    "austria": "AUT - Austria",
    "azerbaijan": "AZE - Azerbaijan",
    "bacbaot": "BRB - Barbados",
    "bahama": "BHS - Bahamas",
    "bahamas": "BHS - Bahamas",
    "bahrain": "BHR - Bahrain",
    "balan": "POL - Poland",
    "bangladesh": "BGD - Bangladesh",
    "banglaet": "BGD - Bangladesh",
    "barain": "BHR - Bahrain",
    "barbados": "BRB - Barbados",
    "becmua": "BMU - Bermuda",
    "belarus": "BLR - Belarus",
    "belarut": "BLR - Belarus",
    "belgium": "BEL - Belgium",
    "belixe": "BLZ - Belize",
    "belize": "BLZ - Belize",
    "benanh": "BEN - Benin",
    "benin": "BEN - Benin",
    "bermuda": "BMU - Bermuda",
    "bhutan": "BTN - Bhutan",
    "bi": "BEL - Belgium",
    "boaonha": "PRT - Portugal",
    "bolivia": "BOL - Bolivia",
    "bosniaandherzegovina": "BIH - Bosnia and Herzegovina",
    "botswana": "BWA - Botswana",
    "botxoana": "BWA - Botswana",
    "boxniahecdegovina": "BIH - Bosnia and Herzegovina",
    "bradin": "BRA - Brazil",
    "brazil": "BRA - Brazil",
    "britishindiaoceanterritory": "IOT - British India Ocean Territory",
    "bulgaria": "BGR - Bulgaria",
    "bungari": "BGR - Bulgaria",
    "buockinaphaxo": "BFA - Burkina Faso",
    "burkinafaso": "BFA - Burkina Faso",
    "burundi": "BDI - Burundi",
    "buruni": "BDI - Burundi",
    "butan": "BTN - Bhutan",
    "cameroon": "CMR - Cameroon",
    "camorun": "CMR - Cameroon",
    "canada": "CAN - Canada",
    "capeverde": "CPV - Cape Verde",
    "capve": "CPV - Cape Verde",
    "chad": "TCD - Chad",
    "charapxyri": "SYR - Syrian Arab Republic",
    "chdcndtrieutien": "PRK - Korea Democratic Peoples Republic of",
    "chhanquoc": "KOR - Korea (South)",
    "chhoigiaoiran": "IRN - Iran Ilasmic Republic of",
    "chile": "CHL - Chile",
    "china": "CHN - China",
    "chinataiwan": "CHN - China",
    "chlienbanguc": "D - Germany",
    "chmaxeonia": "MKD - Macedonia",
    "chominicana": "DMA - Dominica",
    "colombia": "COL - Colombia",
    "como": "COM - Comoros",
    "comoros": "COM - Comoros",
    "conggo": "COG - Congo",
    "conghoasec": "CZE - Czech Republic",
    "congo": "COG - Congo",
    "congquocanora": "AND - Andorra",
    "congquoclichtenxten": "LIE - Liechtenstein",
    "congquocmonaco": "MCO - Monaco",
    "cooet": "KWT - Kuwait",
    "costarica": "CRI - Costa Rica",
    "cotedivoire": "CIV - Cote d' Ivoire",
    "cotivoa": "CIV - Cote d' Ivoire",
    "coxtarica": "CRI - Costa Rica",
    "croatia": "HRV - Croatia",
    "cuba": "CUB - Cuba",
    "cyprus": "CYP - Cyprus",
    "czechrepublic": "CZE - Czech Republic",
    "dambia": "ZMB - Zambia",
    "denmark": "DNK - Denmark",
    "dimbabue": "ZWE - Zimbabwe",
    "djibouti": "DJI - Djibouti",
    "dominica": "DMA - Dominica",
    "dominicana": "DMA - Dominica",
    "ecuador": "ECU - Ecuador",
    "ecuao": "ECU - Ecuador",
    "egypt": "EGY - Egypt",
    "elsalvador": "SLV - El Salvado",
    "enxanvao": "SLV - El Salvado",
    "equatorialguinea": "GNQ - Equatorial Guinea",
    "eritoria": "ERI - Eritrea",
    "eritrea": "ERI - Eritrea",
    "estonia": "EST - Estonia",
    "ethiopia": "ETH - Ethiopia",
    "etiopia": "ETH - Ethiopia",
    "extonia": "EST - Estonia",
    "fiji": "FJI - Fiji",
    "finland": "FIN - Finland",
    "france": "FRA - France",
    "francemetropolitan": "FRA - France",
    "gabon": "GAB - Gabon",
    "gabong": "GAB - Gabon",
    "gambia": "GMB - Gambia",
    "gana": "GHA - Ghana",
    "georgia": "GEO - Georgia",
    "germany": "D - Germany",
    "ghana": "GHA - Ghana",
    "ghine": "GIN - Guinea",
    "ghinebitxao": "GNB - Guinea-Bissau",
    "ghinexichao": "GNQ - Equatorial Guinea",
    "giamahiriiaaraplibinhandan": "LBY - Libyan Arab Jamahiriya",
    "gibraltar": "GIB - Gibraltar",
    "gibranta": "GIB - Gibraltar",
    "goatemala": "GTM - Guatemala",
    "greece": "GRC - Greece",
    "greenland": "GRL - Greenland",
    "grenaa": "GRD - Grenada",
    "grenada": "GRD - Grenada",
    "grinlon": "GRL - Greenland",
    "grudia": "GEO - Georgia",
    "guatemala": "GTM - Guatemala",
    "guina": "GUY - Guyana",
    "guinea": "GIN - Guinea",
    "guineabissau": "GNB - Guinea-Bissau",
    "guyana": "GUY - Guyana",
    "haiti": "HTI - Haiti",
    "halan": "NLD - Netherland",
    "hanquoc": "KOR - Korea (South)",
    "honduras": "HND - Honduras",
    "hondurat": "HND - Honduras",
    "hylap": "GRC - Greece",
    "iaphanthuoclienhiepanh": "GBD - United Kingdom British Territories Citizen",
    "ibouti": "DJI - Djibouti",
    "iceland": "ISL - Iceland",
    "india": "IND - India",
    "indonesia": "IDN - Indonesia",
    "inonexia": "IDN - Indonesia",
    "irac": "IRQ - Iraq",
    "iran": "IRN - Iran Ilasmic Republic of",
    "iraq": "IRQ - Iraq",
    "ireland": "IRL - Ireland",
    "israel": "ISR - Israel",
    "italia": "ITA - Italy",
    "italy": "ITA - Italy",
    "ixraen": "ISR - Israel",
    "jamaica": "JAM - Jamaica",
    "japan": "JPN - Japan",
    "jocan": "JOR - Jordan",
    "jordan": "JOR - Jordan",
    "kadacxtan": "KAZ - Kazakhstan",
    "kazakhstan": "KAZ - Kazakhstan",
    "kenia": "KEN - Kenya",
    "kenya": "KEN - Kenya",
    "kiecghidia": "KGZ - Kyrgyzstan",
    "kiribati": "KIR - Kiribati",
    "koreademocraticpeoplesrepublic": "PRK - Korea Democratic Peoples Republic of",
    "koreasouth": "KOR - Korea (South)",
    "kosovo": "RKS - Kosovo",
    "kuwait": "KWT - Kuwait",
    "kyrgyzstan": "KGZ - Kyrgyzstan",
    "latvia": "LVA - Latvia",
    "lebanon": "LBN - Lebanon",
    "lesotho": "LSO - Lesotho",
    "lexotho": "LSO - Lesotho",
    "liban": "LBN - Lebanon",
    "liberia": "LBR - Liberia",
    "libya": "LBY - Libyan Arab Jamahiriya",
    "liechtenstein": "LIE - Liechtenstein",
    "lienbangnga": "RUS - Russia",
    "lithuania": "LTU - Lithuania",
    "luxembourg": "LUX - Luxembourg",
    "luychxembua": "LUX - Luxembourg",
    "maagaxca": "MDG - Madagascar",
    "macedonia": "MKD - Macedonia",
    "madagascar": "MDG - Madagascar",
    "malaixia": "MYS - Malaysia",
    "malauy": "MWI - Malawi",
    "malawi": "MWI - Malawi",
    "malaysia": "MYS - Malaysia",
    "maldives": "MDV - Maldives",
    "mali": "MLI - Mali",
    "malta": "MLT - Malta",
    "manivo": "MDV - Maldives",
    "manta": "MLT - Malta",
    "maroc": "MAR - Morocco",
    "marshallislands": "MHL - Marshall Islands",
    "mauritania": "MRT - Mauritania",
    "mauritius": "MUS - Mauritius",
    "mexico": "MEX - Mexico",
    "mianma": "MMR - Myanmar",
    "micronesia": "FSM - Micronesia",
    "modambich": "MOZ - Mozambique",
    "moldova": "MDA - Moldova",
    "monaco": "MCO - Monaco",
    "mongco": "MNG - Mongolia",
    "mongolia": "MNG - Mongolia",
    "monova": "MDA - Moldova",
    "montenegro": "MNE - Montenegro",
    "montserrat": "MSR - Montserrat",
    "monxerat": "MSR - Montserrat",
    "moratani": "MRT - Mauritania",
    "morixo": "MUS - Mauritius",
    "morocco": "MAR - Morocco",
    "mozambique": "MOZ - Mozambique",
    "my": "USA - United States of America",
    "myanmarburma": "MMR - Myanmar",
    "namibia": "NAM - Namibia",
    "nauru": "NRU - Nauru",
    "nauy": "NOR - Norway",
    "nepal": "NPL - Nepal",
    "nepan": "NPL - Nepal",
    "netherland": "NLD - Netherland",
    "netherlandantilles": "NLD - Netherland",
    "newzealand": "NZL - New Zealand",
    "nhatban": "JPN - Japan",
    "nicaragoa": "NIC - Nicaragua",
    "nicaragua": "NIC - Nicaragua",
    "niger": "NER - Niger",
    "nigeria": "NGA - Nigeria",
    "nigie": "NER - Niger",
    "nigieria": "NGA - Nigeria",
    "niudilan": "NZL - New Zealand",
    "norway": "NOR - Norway",
    "oman": "OMN - Oman",
    "ominica": "DMA - Dominica",
    "ongtimo": "TLS - Timor Leste",
    "oxtraylia": "AUS - Australia",
    "pakistan": "PAK - Pakistan",
    "pakixtan": "PAK - Pakistan",
    "palau": "PLW - Palau",
    "palestine": "PSE - Palestine",
    "palextin": "PSE - Palestine",
    "panama": "PAN - Panama",
    "papuanewguinea": "PNG - Papua New Guinea",
    "papuaniughine": "PNG - Papua New Guinea",
    "paragoay": "PRY - Paraguay",
    "paraguay": "PRY - Paraguay",
    "peru": "PER - Peru",
    "phanlan": "FIN - Finland",
    "phap": "FRA - France",
    "philippin": "PHL - Philippines",
    "philippine": "PHL - Philippines",
    "poland": "POL - Poland",
    "portugal": "PRT - Portugal",
    "qatar": "QAT - Qatar",
    "quanaoantithuochalan": "NLD - Netherland",
    "quanaomacsan": "MHL - Marshall Islands",
    "quanaonamgrudiavanamsanuych": "GEO - Georgia",
    "quanaoxaysen": "SYC - Seychelles",
    "quata": "QAT - Qatar",
    "romania": "ROU - Romania",
    "ruana": "RWA - Rwanda",
    "rumani": "ROU - Romania",
    "russia": "RUS - Russia",
    "rwanda": "RWA - Rwanda",
    "saintlucia": "LCA - Saint Lucia",
    "sanmarino": "SMR - San Marino",
    "sat": "TCD - Chad",
    "saudiarabia": "SAU - Saudi Arabia",
    "scotland": "SC- - Scotland",
    "senegal": "SEN - Senegal",
    "serbia": "SRB - Serbia",
    "seychelles": "SYC - Seychelles",
    "singapore": "SGP - Singapore",
    "slovakia": "SVK - Slovakia",
    "slovenia": "SVN - Slovenia",
    "somalia": "SOM - Somalia",
    "southgeorgiaandthesouths": "GEO - Georgia",
    "spain": "ESP - Spain",
    "srilanka": "LKA - Sri Lanka",
    "sudan": "SDN - Sudan",
    "suriname": "SUR - Suriname",
    "swaziland": "SWZ - Swaziland",
    "sweden": "SWE - Sweden",
    "switzerland": "CHE - Switzerland",
    "syria": "SYR - Syrian Arab Republic",
    "tagikixtan": "TJK - Tajikistan",
    "tajikistan": "TJK - Tajikistan",
    "taybannha": "ESP - Spain",
    "thailan": "THA - Thailand",
    "thailand": "THA - Thailand",
    "thonhiky": "TUR - Turkey",
    "thuyien": "SWE - Sweden",
    "thuysi": "CHE - Switzerland",
    "timorleste": "TLS - Timor Leste",
    "tochucdantocthongnhat": "UNO - United Nations Organization",
    "togo": "TGO - Togo",
    "tonga": "TON - Tonga",
    "trungquoc": "CHN - China",
    "trungquocailoan": "CHN - China",
    "tunidi": "TUN - Tunisia",
    "tunisia": "TUN - Tunisia",
    "tuocmenixtan": "TKM - Turkmenistan",
    "turkey": "TUR - Turkey",
    "turkmenistan": "TKM - Turkmenistan",
    "tuvalu": "TUV - Tuvalu",
    "uc": "D - Germany",
    "ucraina": "UKR - Ukraine",
    "udobekixtan": "UZB - Uzbekistan",
    "uganda": "UGA - Uganda",
    "ukraine": "UKR - Ukraine",
    "unitedarabemirates": "ARE - United Arab Emirates",
    "unitedkingdom": "GBD - United Kingdom British Territories Citizen",
    "unitednationsorganization": "UNO - United Nations Organization",
    "unitedstates": "USA - United States of America",
    "urugoay": "URY - Uruguay",
    "uruguay": "URY - Uruguay",
    "uzbekistan": "UZB - Uzbekistan",
    "vanuatu": "VUT - Vanuatu",
    "vaticancity": "VAT - Holy See (Vatican City State )",
    "vaticang": "VAT - Holy See (Vatican City State )",
    "veneduela": "VEN - Venezuela",
    "venezuela": "VEN - Venezuela",
    "vietnam": "VNM - Viet Nam",
    "vungatthuocanhoanoduong": "IOT - British India Ocean Territory",
    "vungthuophap": "FRA - France",
    "vuongquocnauy": "NOR - Norway",
    "westernsamoa": "WSM - Western Samoa",
    "xamoa": "WSM - Western Samoa",
    "xanhluxia": "LCA - Saint Lucia",
    "xanmarino": "SMR - San Marino",
    "xcolent": "SC- - Scotland",
    "xecbia": "SRB - Serbia",
    "xenegan": "SEN - Senegal",
    "xingapo": "SGP - Singapore",
    "xlovakia": "SVK - Slovakia",
    "xoadilen": "SWZ - Swaziland",
    "xomali": "SOM - Somalia",
    "xrilanca": "LKA - Sri Lanka",
    "xuang": "SDN - Sudan",
    "xurinam": "SUR - Suriname",
    "y": "ITA - Italy",
    "yemen": "YEM - Yemen",
    "zambia": "ZMB - Zambia",
    "zimbabwe": "ZWE - Zimbabwe"
}

def lookup_nat_kbtt(raw):
    """Khớp thông minh: chuẩn hóa dấu/khoảng trắng để tìm mã quốc tịch."""
    if not raw: return ''
    raw = str(raw).strip()
    key = _norm_nat(raw)
    if key in NAT_NORM:
        return NAT_NORM[key]
    # already in CODE - Name form?
    if _re.match(r'^[A-Z]{2,3} - ', raw):
        return raw
    return raw  # unknown -> keep original (sẽ hiện cảnh báo)

NAT_DK14 = {
    'RUS':'Russia  (Liên bang Nga)','UZB':'Uzbekistan  ( U-dơ-bê-ki-xtan )',
    'KAZ':'Kazakhstan  ( Ka-dắc-xtan )','KOR':'Korea (South)  ( CH Hàn Quốc )',
    'KGZ':'Kyrgyzstan  ( Kiếc-ghi-di-a )','TJK':'Tajikistan  ( Ta-gi-ki-xtan )',
    'UKR':'Ukraine  ( U-crai-na )','USA':'United States  ( Mỹ )',
    'VNM':'Vietnam  ( Việt Nam )','CAN':'Canada  ( Ca-na-da )',
    'GBR':'United Kingdom  ( Anh )','AUS':'Australia  ( Ô-xtrây-li-a )',
    'BLR':'Belarus  ( Bê-la-rút )','CHN':'China  ( Trung Quốc )',
    'DEU':'Germany  ( Đức )','MDA':'Moldova  ( Môn-đô-va )',
    'FIN':'Finland  ( Phần Lan )','FRA':'France  ( Pháp )',
    'DNK':'Denmark  ( Đan Mạch )','MUS':'Mauritius  ( Mô-ri-xơ )',
}
LOAI_GIAY = {
    'Căn cước công dân':'8 - Thẻ Căn Cước','Hộ chiếu':'4 - Hộ chiếu',
    'Chứng minh nhân dân':'2 - Thẻ CMND','Căn cước':'1 - Thẻ CCCD',
    'Giấy khai sinh':'5 - Giấy khai sinh',
}
TINH = {
    'DAK LAK':'605 - Đắk Lắk','DAK NONG':'607 - Đắk Nông','PHU YEN':'509 - Phú Yên',
    'HO CHI MINH':'701 - TP. Hồ Chí Minh','THP HO CHI MINH':'701 - TP. Hồ Chí Minh',
    'HCM':'701 - TP. Hồ Chí Minh','TP HCM':'701 - TP. Hồ Chí Minh',
    'GIA LAI':'603 - Gia Lai','QUANG NGAI':'505 - Quảng Ngãi','LAM DONG':'703 - Lâm Đồng',
    'LAM DONG.':'703 - Lâm Đồng','KHANH HOA':'511 - Khánh Hòa','BEN TRE':'817 - Bến Tre',
    'VINH LONG':'809 - Vĩnh Long','NINH THUAN':'513 - Ninh Thuận','DONG NAI':'713 - Đồng Nai',
    'TIEN GIANG':'807 - Tiền Giang','LONG AN':'801 - Long An','BINH DUONG':'707 - Bình Dương',
    'BINH THUAN':'705 - Bình Thuận','CAN THO':'815 - TP. Cần Thơ','DA NANG':'501 - TP. Đà Nẵng',
    'HA NOI':'101 - TP. Hà Nội','BA RIA VUNG TAU':'711 - Bà Rịa - Vũng Tàu',
    'DONG THAP':'803 - Đồng Tháp','AN GIANG':'805 - An Giang','KIEN GIANG':'819 - Kiên Giang',
    'HA TINH':'407 - Hà Tĩnh','BINH DINH':'507 - Bình Định','VIET NAM':'',
}

# ── Helpers ───────────────────────────────────────────────────────────────
def fmt(v):
    if v is None or str(v) in ('NaT','nan',''): return ''
    if hasattr(v,'strftime'): return v.strftime('%d/%m/%Y')
    return str(v).strip()[:10]

def make_code(prefix, ns):
    if not ns: return prefix
    p = ns.replace('-','/').split('/')
    return f"{prefix}{p[0].zfill(2)}{p[1].zfill(2)}{p[2][-2:]}" if len(p)==3 else prefix

def cp(src, dst):
    for a in ('font','fill','border','alignment'):
        v = getattr(src,a)
        if v: setattr(dst,a,copy(v))
    dst.number_format = src.number_format

def serial2date(s):
    if not s: return None
    try: return datetime.datetime(1899,12,30)+datetime.timedelta(days=int(s))
    except: return None

def wb_to_bytes(wb):
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

# ── Processing ────────────────────────────────────────────────────────────
def process_xlsx(xlsx_bytes, rate):
    """Điền dữ liệu file đầu vào lên FILE MẪU customer (QLLT) — giữ nguyên 100%
    template: sheet 'customer' + sheet 'Danh-muc', định dạng header, style ô dữ liệu.
    Map cột theo tên (không phân biệt hoa thường / dấu / khoảng trắng), quy đổi
    ĐƠN GIÁ (USD → VND, tô vàng ô đã đổi), đánh lại STT."""
    src_wb = load_workbook(io.BytesIO(xlsx_bytes))
    src_ws = src_wb.active

    # Map header nguồn (chuẩn hóa) → chỉ số cột nguồn
    src_map = {}
    for c in src_ws[1]:
        if c.value is not None and str(c.value).strip():
            src_map.setdefault(_norm_nat(c.value), c.column)

    # Nạp template QLLT (customer + Danh-muc), dòng 2 là mẫu định dạng ô dữ liệu
    wb = load_workbook(io.BytesIO(load_template('customer')))
    ws = wb['customer']
    n_cols = ws.max_column
    ref = [ws.cell(2, ci) for ci in range(1, n_cols + 1)]
    ref_styles = [(copy(c.font), copy(c.fill), copy(c.border), copy(c.alignment), c.number_format) for c in ref]

    # Với mỗi cột template, tìm cột nguồn tương ứng theo tên
    headers = [ws.cell(1, ci).value for ci in range(1, n_cols + 1)]
    col_src = [src_map.get(_norm_nat(h)) if h else None for h in headers]
    don_gia_idx = next((i + 1 for i, h in enumerate(headers) if h and _norm_nat(h) == _norm_nat('ĐƠN GIÁ')), None)

    ws.delete_rows(2)  # bỏ dòng mẫu

    conv = 0
    er = 1
    for row in src_ws.iter_rows(min_row=2, max_row=src_ws.max_row):
        if all(c.value is None or str(c.value).strip() == '' for c in row):
            continue
        er += 1
        for ci in range(1, n_cols + 1):
            cell = ws.cell(er, ci)
            f, fl, b, a, nf = ref_styles[ci - 1]
            cell.font = copy(f); cell.fill = copy(fl); cell.border = copy(b)
            cell.alignment = copy(a); cell.number_format = nf
            sc = col_src[ci - 1]
            if sc is not None and row[sc - 1].value is not None:
                cell.value = row[sc - 1].value
        ws.cell(er, 1).value = er - 1  # STT đánh lại
        if don_gia_idx:
            dg = ws.cell(er, don_gia_idx)
            if dg.value and isinstance(dg.value, (int, float)) and 0 < dg.value < 1000:
                dg.value = round(dg.value * rate)
                dg.fill = PatternFill("solid", start_color="FFFF00")
                conv += 1
    return wb, conv

def split_wb(wb, loai):
    wb2 = load_workbook(io.BytesIO(wb_to_bytes(wb)))
    ws2 = wb2.active
    lc = next(c.column for c in ws2[1] if c.value=='LOẠI KHÁCH')
    dels = [row[0].row for row in ws2.iter_rows(min_row=2,max_row=ws2.max_row)
            if row[lc-1].value != loai]
    for r in reversed(dels): ws2.delete_rows(r)
    for i, row in enumerate(ws2.iter_rows(min_row=2,max_row=ws2.max_row),1): row[0].value=i
    return wb2

def _norm_name(s):
    """Chuẩn hóa tên để khớp giữa 2 file: bỏ dấu, hoa thường, gộp khoảng trắng."""
    s = str(s).lower().strip()
    s = _ud.normalize('NFD', s)
    s = ''.join(c for c in s if _ud.category(c) != 'Mn')
    return _re.sub(r'\s+', ' ', _re.sub(r'[^a-z0-9 ]', '', s)).strip()

def parse_visa_file(visa_bytes):
    """Đọc file thô visa/quản lý người nước ngoài → dict {"by_pp": {...}, "by_name": {...}}.
    Hỗ trợ 2 định dạng cột:
    1) File "Trang quản lý người nước ngoài" thật: 'HỌ TÊN' (tên đầy đủ 1 cột),
       'SỐ HỘ CHIẾU', và 'THỜI HẠN ĐƯỢC PHÉP TẠM TRÚ TẠI VIỆT NAM'.
    2) File cũ dạng Last Name / First Name / Visa date.
    Khớp theo SỐ HỘ CHIẾU trước (chính xác nhất, không sợ trùng tên/sai thứ tự),
    tên đã chuẩn hóa dùng làm dự phòng khi không có/không khớp số hộ chiếu.
    Tự động BỎ QUA khách Việt Nam (nếu file có cột quốc tịch) — công dân VN
    không có "thời hạn tạm trú tại Việt Nam" nên không cần đưa vào visa_map."""
    df = pd.read_excel(io.BytesIO(visa_bytes))
    def _find(*names):
        for n in names:
            for c in df.columns:
                if _norm_nat(c) == _norm_nat(n):
                    return c
        return None
    c_full  = _find('HỌ TÊN', 'HO TEN', 'Họ và tên', 'Full Name', 'FullName')
    c_last  = _find('Last Name', 'LastName', 'Họ')
    c_first = _find('First Name', 'FirstName', 'Tên')
    c_pp    = _find('SỐ HỘ CHIẾU', 'So Ho Chieu', 'Passport', 'Passport Number', 'PassportNo')
    c_nat   = _find('QUỐC TỊCH', 'MÃ QUỐC TỊCH', 'Nationality', 'LOẠI KHÁCH')
    c_date  = _find('Visa date', 'Visadate', 'Thời hạn tạm trú',
                     'THỜI HẠN ĐƯỢC PHÉP TẠM TRÚ TẠI VIỆT NAM',
                     'Thoi han duoc phep tam tru tai Viet Nam', 'Tam tru den')
    if c_date is None:
        raise ValueError("File visa không có cột ngày visa/thời hạn tạm trú. Vui lòng kiểm tra lại file.")

    def _is_vn(val):
        """Nhận diện khách Việt Nam qua cột quốc tịch/loại khách (nếu có)."""
        s = _norm_nat(val)
        return s in ('vnm', 'viet nam', 'vietnam') or s.startswith('vnm') or s == 'viet nam'

    by_pp = {}
    by_name = {}
    n_skipped_vn = 0
    for _, r in df.iterrows():
        # Bỏ qua khách Việt Nam nếu nhận diện được qua cột quốc tịch/loại khách
        if c_nat and pd.notna(r[c_nat]) and _is_vn(r[c_nat]):
            n_skipped_vn += 1
            continue
        d = r[c_date]
        if pd.isna(d):
            continue
        # Cột ngày đọc THẲNG (month=tháng thật, day=ngày thật) — KHÔNG đảo như
        # cột Departure. Đã kiểm chứng: có ngày 25/26/31 nên day chính là ngày thật.
        if hasattr(d, 'year') and not isinstance(d, str):
            dstr = f"{d.day:02d}/{d.month:02d}/{d.year}"
        else:
            s = str(d).strip()
            p = s.split('/')
            if len(p) == 3:
                dd, mm, yy = p
                if len(yy) == 2: yy = '20' + yy
                try:
                    dstr = f"{int(dd):02d}/{int(mm):02d}/{yy}"
                except Exception:
                    continue
            else:
                try:
                    t = pd.to_datetime(s, dayfirst=True)
                    dstr = f"{t.day:02d}/{t.month:02d}/{t.year}"
                except Exception:
                    continue

        # Khớp theo số hộ chiếu — ưu tiên, chính xác nhất
        if c_pp and pd.notna(r[c_pp]):
            pp_key = _norm_pp(r[c_pp])
            if pp_key:
                by_pp.setdefault(pp_key, dstr)

        # Khớp theo tên — dự phòng khi không có/không khớp số hộ chiếu
        if c_full and pd.notna(r[c_full]):
            key = _norm_name(str(r[c_full]))
            if key:
                by_name.setdefault(key, dstr)
        else:
            ln = str(r[c_last]).strip() if c_last and pd.notna(r[c_last]) else ''
            fn = str(r[c_first]).strip() if c_first and pd.notna(r[c_first]) else ''
            # Lưu cả 2 thứ tự để khớp linh hoạt dù file NNN đảo Họ/Tên
            for combo in ((ln + ' ' + fn), (fn + ' ' + ln)):
                key = _norm_name(combo)
                if key:
                    by_name.setdefault(key, dstr)
    return {"by_pp": by_pp, "by_name": by_name, "skipped_vn": n_skipped_vn}

def build_kbtt(df_intl, visa_map=None):
    """Điền mẫu KBTT. Dòng 3 là dòng "[TEST] SAMPLE" BẮT BUỘC giữ nguyên (không
    bị ghi đè) — dữ liệu khách thật được điền bắt đầu từ dòng 4 trở xuống.
    Nếu có visa_map thì điền cột L 'THỜI HẠN ĐƯỢC PHÉP TẠM TRÚ TẠI VIỆT NAM' —
    khớp theo SỐ HỘ CHIẾU trước (chính xác nhất), tên là dự phòng; nếu không
    khớp thì để trống cột đó. Trả về (wb, danh_sách_tên_không_khớp)."""
    visa_map = visa_map or {}
    # Tương thích ngược: nếu visa_map là dict phẳng {tên: ngày} kiểu cũ, coi như by_name
    if isinstance(visa_map, dict) and ("by_pp" in visa_map or "by_name" in visa_map):
        by_pp = visa_map.get("by_pp", {})
        by_name = visa_map.get("by_name", {})
    else:
        by_pp = {}
        by_name = visa_map
    unmatched = []
    wb = load_workbook(io.BytesIO(load_template('kbtt')))
    ws = wb['KBTT']
    # Cấu trúc mẫu: dòng 1 = ô merge A1:L1 (tiêu đề + chú ý đỏ), dòng 2 = header,
    # dòng 3 = "[TEST] SAMPLE" BẮT BUỘC giữ nguyên, dữ liệu khách thật từ dòng 4.
    ref = [ws.cell(3,c) for c in range(1,13)]   # dùng style dòng 3 làm mẫu định dạng cho các dòng khách
    n = len(df_intl)
    # Xóa dòng dữ liệu thừa (nếu có), luôn giữ tối thiểu tới dòng 3 (dòng TEST)
    last_data_row = 3 + n
    if ws.max_row > last_data_row:
        ws.delete_rows(last_data_row + 1, ws.max_row - last_data_row)
    for i,(_,row) in enumerate(df_intl.iterrows(),1):
        er=i+3   # dữ liệu khách thật bắt đầu dòng 4 (dòng 3 là TEST, giữ nguyên)
        ht=str(row.get('HỌ TÊN ',row.get('HỌ TÊN',''))).strip()
        ns=fmt(row['NGÀY SINH']); nd=fmt(row['NGÀY ĐẾN']); ni=fmt(row.get('NGÀY ÐI',row.get('NGÀY ĐI','')))
        gt='M - Nam' if str(row.get('GIỚI TÍNH','')).strip()=='Nam' else 'F - Nữ'
        qt=lookup_nat_kbtt(row.get('QUỐC TỊCH',''))
        sh=str(row.get('SỐ GIẤY TỜ','')).strip(); sp=str(row.get('SỐ PHÒNG','')).strip()
        # Cột L (12) — THỜI HẠN ĐƯỢC PHÉP TẠM TRÚ TẠI VIỆT NAM:
        # khớp theo SỐ HỘ CHIẾU trước (chính xác nhất), tên là dự phòng
        if by_pp or by_name:
            vd = by_pp.get(_norm_pp(sh), '') or by_name.get(_norm_name(ht), '')
            if not vd:
                unmatched.append(ht)
        else:
            vd = ''   # không có file visa → cột L để trống
        vals=[i,ht,ns,'D - Ngày',gt,qt,sh,sp,nd,ni,ni,vd]
        for ci,val in enumerate(vals,1):
            cell=ws.cell(er,ci); cell.value=val if isinstance(val,int) else str(val)
            cp(ref[ci-1],cell)
    # Bảo toàn ô merge tiêu đề + chiều cao dòng 1 (phòng khi delete_rows làm xê dịch)
    if 'A1:L1' not in [str(m) for m in ws.merged_cells.ranges]:
        try: ws.merge_cells('A1:L1')
        except Exception: pass
    ws.row_dimensions[1].height = 41.1
    # Khôi phục rich text ô A1: tiêu đề (đen) + dòng chú ý (ĐỎ) — openpyxl làm mất khi save
    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
        from openpyxl.styles.colors import Color
        from openpyxl.styles import Alignment
        a1 = ws.cell(1,1)
        a1.value = CellRichText(
            TextBlock(InlineFont(rFont='Times New Roman', sz=16, b=True),
                      'DANH SÁCH HỒ SƠ KBTT\r\n'),
            TextBlock(InlineFont(rFont='Times New Roman', sz=16, b=True, color=Color(rgb='FFFF0000')),
                      '(*Lưu ý: Người khai báo chịu trách nhiệm trước pháp luật về các nội dung thông tin cung cấp)')
        )
        a1.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    except Exception:
        pass
    # Cập nhật vùng Table1 cho khớp số dòng thực (header + dòng TEST + dữ liệu khách)
    if 'Table1' in ws.tables:
        ws.tables['Table1'].ref = f"A2:L{3 + n}"
    return wb, unmatched

def build_vnm(df_vn):
    wb = load_workbook(io.BytesIO(load_template('vnm')))
    wsn = next((s for s in wb.sheetnames if 'KHACH' in s or 'DS' in s), wb.sheetnames[0])
    ws = wb[wsn]
    ref = [ws.cell(5,c) for c in range(1,ws.max_column+1)]
    for r in range(ws.max_row,4,-1): ws.delete_rows(r)
    gks_cnt=0; gbl_cnt=0
    for i,(_,row) in enumerate(df_vn.iterrows(),1):
        er=i+4
        ht=str(row.get('HỌ TÊN ',row.get('HỌ TÊN',''))).strip()
        ns=fmt(row['NGÀY SINH']); nd=fmt(row['NGÀY ĐẾN']); ni=fmt(row.get('NGÀY ÐI',row.get('NGÀY ĐI','')))
        gt='F - Nữ' if str(row.get('GIỚI TÍNH','')).strip()=='Nữ' else 'M - Nam'
        sg_raw=str(row.get('SỐ GIẤY TỜ','')).strip()
        lg_raw=str(row.get('LOẠI GIẤY TỜ','')).strip()
        is_gks='GKS' in sg_raw.upper(); is_gbl='GBL' in sg_raw.upper()
        ten_giay=''
        if is_gks:
            sg=make_code('GKS',ns); lg='5 - Giấy khai sinh'; gks_cnt+=1
        elif is_gbl:
            sg=make_code('GBL',ns); lg='9 - Giấy tờ khác'; ten_giay='Giấy bảo lãnh'; gbl_cnt+=1
        elif sg_raw and sg_raw[0].isalpha():
            # Số giấy tờ bắt đầu bằng chữ cái (vd: P02628567) → là hộ chiếu
            sg=sg_raw; lg='4 - Hộ chiếu'
        else:
            sg=sg_raw; lg=LOAI_GIAY.get(lg_raw,lg_raw)
        tinh=TINH.get(str(row.get('TP/TỈNH','')).strip().upper(),'')
        dc=str(row.get('ÐỊA CHỈ',row.get('ĐỊA CHỈ',''))).strip()
        sp=str(row.get('SỐ PHÒNG','')).strip()
        vals=[i,ht,ns,gt,'VNM - Viet Nam',lg,ten_giay,sg,'','1 - Thường trú',tinh,'',dc,nd,ni,sp,'1 - Du lịch','','']
        for ci,val in enumerate(vals,1):
            cell=ws.cell(er,ci); cell.value=val if isinstance(val,int) else str(val)
            if ci<=len(ref): cp(ref[ci-1],cell)
    return wb, gks_cnt, gbl_cnt

def build_dk14(xls_bytes):
    wb2=xlrd.open_workbook(file_contents=xls_bytes)
    ws2=wb2.sheet_by_index(0)
    data=[[ws2.cell_value(r,c) for c in range(ws2.ncols)]
          for r in range(1,ws2.nrows) if any(ws2.cell_value(r,c) for c in range(ws2.ncols))]
    wb_t=load_workbook(io.BytesIO(load_template('dk14')))
    ws_t=wb_t.active
    cw={col:ws_t.column_dimensions[col].width for col in ws_t.column_dimensions}
    rh={r:ws_t.row_dimensions[r].height for r in ws_t.row_dimensions if r<=17}
    wb_o=Workbook(); ws_o=wb_o.active
    for col,w in cw.items(): ws_o.column_dimensions[col].width=w
    for r,h in rh.items():
        if h: ws_o.row_dimensions[r].height=h
    def cc(s,d):
        d.value=s.value
        for a in ('font','fill','border','alignment'):
            v=getattr(s,a)
            if v: setattr(d,a,copy(v))
        d.number_format=s.number_format
    for r in range(1,18):
        for c in range(1,14): cc(ws_t.cell(r,c),ws_o.cell(r,c))
    for mc in ws_t.merged_cells.ranges:
        if mc.min_row<=17:
            ws_o.merge_cells(start_row=mc.min_row,start_column=mc.min_col,
                             end_row=mc.max_row,end_column=mc.max_col)
    wb_t.close()
    thin=Side(style='thin'); bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    fn=Font(name='Times New Roman',size=12)
    ac=Alignment(horizontal='center',vertical='center',wrap_text=True)
    al=Alignment(horizontal='left',vertical='center',wrap_text=True)
    for i,row in enumerate(data,1):
        er=i+17
        name=str(row[1]).strip() if row[1] else ''
        gender=str(row[3]).strip().upper() if row[3] else ''
        country=NAT_DK14.get(str(row[4]).strip().upper() if row[4] else '',str(row[4] or ''))
        passport=str(row[5]).strip() if row[5] else ''
        if passport.endswith('.0'): passport=passport[:-2]
        address=str(row[6]).strip() if row[6] else '   '
        dob=serial2date(row[2]); arr=serial2date(row[7]); dep=serial2date(row[8])
        room=str(row[9]).strip() if row[9] else ''
        notify=str(row[10]).strip() if row[10] else ''
        cols=[(1,i,ac),(2,name,al),(3,dob if gender=='M' else None,ac),
              (4,dob if gender=='F' else None,ac),(5,country,ac),(6,passport,ac),
              (7,address,al),(8,arr,ac),(9,dep,ac),(10,room,ac),(11,notify,al),(12,'',ac),(13,'',ac)]
        for ci,val,aln in cols:
            cell=ws_o.cell(er,ci); cell.value=val; cell.font=fn; cell.border=bdr; cell.alignment=aln
            if ci in (3,4) and val: cell.number_format='DD/MM/YYYY'
            elif ci in (8,9) and val: cell.number_format='DD/MM/YYYY'
    return wb_o, len(data)


# ── Regcard PDF builder ───────────────────────────────────────────────────
@st.cache_resource
def load_regcard_template():
    path = os.path.join(os.path.dirname(__file__), 'tmpl_regcard.b64')
    with open(path, 'r') as f:
        return base64.b64decode(f.read())

def load_group_template():
    path = os.path.join(os.path.dirname(__file__), 'tmpl_group.b64')
    with open(path, 'r') as f:
        return base64.b64decode(f.read())

def _grp_date(d):
    """Ngày cho regcard group: dd/mm/yyyy (như mẫu 24/07/2026)."""
    if pd.isna(d): return ''
    if hasattr(d, 'strftime'):
        return f"{d.day:02d}/{d.month:02d}/{d.year}"
    s = str(d).strip()
    if '/' in s:
        p = s.split('/')
        if len(p) == 3:
            dd, mm, yy = p
            if len(yy) == 2: yy = '20' + yy
            return f"{int(dd):02d}/{int(mm):02d}/{yy}"
        return s
    # Chuỗi ISO yyyy-mm-dd
    try:
        t = pd.to_datetime(s)
        return f"{t.day:02d}/{t.month:02d}/{t.year}"
    except Exception:
        return s

def _rc_clean_name(n):
    if pd.isna(n): return ''
    return str(n).strip().rstrip(',').strip()

def _rc_conf(c):
    if pd.isna(c): return ''
    return str(int(c)) if isinstance(c,(int,float)) else str(c)

def _rc_date(d):
    if pd.isna(d): return ''
    if hasattr(d,'strftime') and not isinstance(d, str):
        return f"{d.day:02d}/{d.month:02d}/{d.year}"   # dd/mm/yyyy
    s=str(d).strip()
    if '/' in s:
        p=s.split('/')
        if len(p)==3:
            dd,mm,yy=p
            if len(yy)==2: yy='20'+yy
            return f"{int(dd):02d}/{int(mm):02d}/{yy}"   # pad 7/8 → 07/08
        return s
    try:
        t=pd.to_datetime(s)                            # ISO yyyy-mm-dd
        return f"{t.day:02d}/{t.month:02d}/{t.year}"
    except Exception:
        return s

def _rc_nights(arr, dep):
    """Số đêm = Departure - Arrival. Xử lý cả datetime lẫn chuỗi dd/mm/yyyy."""
    def _to_ts(v):
        if pd.isna(v): return None
        if hasattr(v, 'year') and not isinstance(v, str):
            return pd.Timestamp(v)          # datetime/Timestamp — giữ nguyên
        s = str(v).strip()
        if '/' in s:                        # chuỗi dd/mm/yyyy hoặc dd/mm/yy
            p = s.split('/')
            if len(p) == 3:
                dd, mm, yy = p
                if len(yy) == 2: yy = '20' + yy
                try:
                    return pd.Timestamp(year=int(yy), month=int(mm), day=int(dd))
                except Exception:
                    return None
        try:
            return pd.to_datetime(s)        # ISO yyyy-mm-dd hoặc dạng khác
        except Exception:
            return None
    a, d = _to_ts(arr), _to_ts(dep)
    if a is None or d is None:
        return ''
    n = (d - a).days
    return str(n) if n >= 0 else ''

def build_group_regcard(grp_df, tmpl_bytes):
    """Vẽ 1 Registration Card for Group từ các dòng cùng 1 mã Group.
    Trả về trang PDF đã merge. Bảng Kind of rooms để trống (điền tay)."""
    H = 841.0
    FONT = "Times-Roman"; SIZE = 11
    first = grp_df.iloc[0]

    # Group Code = giá trị cột Group
    gcode = first.get('Group')
    if pd.notna(gcode):
        gcode = str(int(gcode)) if isinstance(gcode, (int, float)) else str(gcode)
    else:
        gcode = ''

    # Số phòng: đếm phòng unique, loại phòng ảo 9xxx
    rooms = set()
    for r in grp_df['Rm'].dropna():
        s = str(r).strip()
        if s.endswith('.0'): s = s[:-2]
        if s and not _re.fullmatch(r'9\d{3}', s):
            rooms.add(s.upper())
    n_rooms = len(rooms)

    # Số pax = tổng Adt + Chl + Enf
    def _sum(col):
        return int(grp_df[col].fillna(0).sum()) if col in grp_df.columns else 0
    n_pax = _sum('Adt') + _sum('Chl') + _sum('Enf')

    data = {
        'arrival':   (155.6, 181.2, _grp_date(first.get('Arrival'))),
        'departure': (436.9, 181.2, _grp_date(first.get('Departure'))),
        'gcode':     (48.2,  261.6, gcode),
        'gname':     (165.7, 263.9, _rc_clean_name(first.get('Name'))),
        'agent':     (364.4, 263.9, str(first.get('Company')) if pd.notna(first.get('Company')) else ''),
        'nrooms':    (50.9,  330.6, str(n_rooms)),
        'npax':      (187.2, 326.2, str(n_pax)),
    }

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(595, 841))
    c.setFillColor(black)
    c.setFont(FONT, SIZE)
    for key, (x, bottom, val) in data.items():
        if not val:
            continue
        # thu nhỏ nếu tên đoàn / hãng quá dài
        maxw = {'gname': 185, 'agent': 175}.get(key)
        fs = SIZE
        if maxw:
            while fs > 7 and c.stringWidth(val, FONT, fs) > maxw:
                fs -= 0.3
        c.setFont(FONT, fs)
        c.drawString(x, H - bottom, val)
        c.setFont(FONT, SIZE)
    c.save(); buf.seek(0)

    base = PdfReader(io.BytesIO(tmpl_bytes))
    overlay = PdfReader(buf)
    page = base.pages[0]
    page.merge_page(overlay.pages[0])
    return page


def _fix_date(v):
    """Chuẩn hóa ngày về pd.Timestamp đúng nghĩa dd/mm.

    File Smile export bị lỗi: ngày dạng dd/mm với ngày ≤ 12 (vd '7/8/2026' = 7 tháng 8)
    bị Excel hiểu nhầm kiểu Mỹ mm/dd → lưu thành datetime(month=7, day=8) kèm format
    mm-dd-yy. Khi đọc lại, cần HOÁN month↔day để khôi phục: datetime(y,7,8) → 7 tháng 8.
    Ngày ≥ 13 thì Excel không nhầm được nên giữ dạng text dd/mm bình thường.
    """
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    # datetime từ Excel → đã bị đảo month/day, khôi phục bằng cách hoán lại
    if hasattr(v, 'year') and not isinstance(v, str):
        try:
            return pd.Timestamp(year=v.year, month=v.day, day=v.month)
        except Exception:
            return pd.Timestamp(v)   # day>12: không đảo được, giữ nguyên
    s = str(v).strip()
    if '/' in s:
        p = s.split('/')
        if len(p) == 3:
            dd, mm, yy = p
            if len(yy) == 2: yy = '20' + yy
            try:
                return pd.Timestamp(year=int(yy), month=int(mm), day=int(dd))
            except Exception:
                return None
    try:
        return pd.to_datetime(s, dayfirst=True)
    except Exception:
        return None

def build_regcards(xlsx_bytes, only_main=True):
    """Tạo PDF regcard hàng loạt, gộp theo Conf# (đoàn nhiều phòng → 1 regcard,
    các số phòng gộp chung vào ô Room No)."""
    df = pd.read_excel(io.BytesIO(xlsx_bytes))
    # Chuẩn hóa ngày (sửa lỗi Excel đảo dd/mm↔mm/dd với ngày ≤12 từ Smile export)
    for _dc in ('Arrival', 'Departure'):
        if _dc in df.columns:
            df[_dc] = df[_dc].map(_fix_date)
    H = 841.0
    FONT = "Times-Roman"; SIZE = 9.8
    # Baseline chính xác (bottom) đo từ dữ liệu mẫu gốc — chữ trùng khít 100%
    POS = {
        'name':(125.50,109.60),'conf':(526.75,108.52),'arrival':(119.90,144.22),
        'departure':(329.90,144.22),'nights':(500.30,144.22),'type':(113.60,179.92),
        'rm':(360.60,179.92),'company':(221.40,216.32),
        'special':(137.40,288.40),
    }
    # Ô che dữ liệu cũ — vừa khít vùng chữ, không lấn đường kẻ bảng
    BLANK = [
        (125,99,250,110.5),(526,98,568,109.5),(119,133.5,168,145.2),
        (329,133.5,378,145.2),(500,133.5,512,145.2),(113,169,145,180.9),
        (360,169,410,180.9),(221,205.5,320,217.3),
        (135,277,575,289.4),   # che dòng "AI Lunch ( EUR )..." in sẵn trên template
    ]

    # ── Gộp theo Conf# ──
    # Điền mã Conf# xuống các dòng trống (khách đi cùng booking),
    # rồi gộp TẤT CẢ dòng cùng 1 mã Conf# thành 1 regcard, gộp mọi số phòng.
    def _clean_room(r):
        s = str(r).strip()
        if s.endswith('.0'): s = s[:-2]
        return s

    df = df.copy()
    df['_conf_ff'] = df['Conf#'].ffill()

    # ── Tách các booking ĐOÀN (có mã Group) để dùng mẫu Registration Card for Group ──
    # KHÔNG ffill cột Group — chỉ dòng có sẵn mã Group mới thuộc đoàn.
    # Các dòng cùng Conf# trong đoàn: điền Group xuống theo từng Conf# nếu dòng đầu có Group.
    if 'Group' in df.columns:
        # Điền mã Group xuống các dòng cùng Conf# (đoàn nhiều phòng, chỉ dòng đầu có Group)
        df['_group_ff'] = df.groupby('_conf_ff')['Group'].transform(
            lambda s: s.ffill().bfill() if s.notna().any() else s)
    else:
        df['_group_ff'] = pd.NA
    has_group = df['_group_ff'].notna()
    df_group = df[has_group].copy()

    writer = PdfWriter()
    count = 0

    tmpl_bytes = load_regcard_template()
    grp_tmpl = None  # nạp lười khi gặp đoàn đầu tiên
    _rendered_groups = set()  # tránh vẽ trùng 1 đoàn khi trải nhiều Conf#

    # Duyệt theo ĐÚNG THỨ TỰ xuất hiện trong file (gộp theo Conf#).
    # Nhóm nào có mã Group → vẽ 1 trang Registration Card for Group (gộp cả đoàn);
    # nhóm thường → regcard thường như cũ.
    for conf_val, grp in df.groupby('_conf_ff', sort=False):
        main_rows = grp[grp['Conf#'].notna()]
        if len(main_rows) == 0:
            continue
        main = main_rows.iloc[0]

        # ── Nếu booking này thuộc ĐOÀN → dùng mẫu group ──
        gid = grp['_group_ff'].dropna().iloc[0] if grp['_group_ff'].notna().any() else None
        if gid is not None:
            if gid in _rendered_groups:
                continue  # đoàn đã vẽ ở lần gặp trước
            _rendered_groups.add(gid)
            if grp_tmpl is None:
                grp_tmpl = load_group_template()
            gdf = df_group[df_group['_group_ff'] == gid]
            page = build_group_regcard(gdf, grp_tmpl)
            writer.add_page(page)
            count += 1
            continue

        # ── Booking thường → regcard thường ──
        rooms = []
        for r in grp['Rm']:
            if pd.notna(r):
                rs = _clean_room(r)
                # Bỏ phòng ảo đầu 9 dạng 9000-9999 (9002, 9005, 9010, 9040... — posting master)
                if rs and rs not in rooms and not _re.fullmatch(r'9\d{3}', rs):
                    rooms.append(rs)
        # Gom mã Specials của cả nhóm (để bắt CN/EB dù nằm ở dòng nào)
        spec_codes = set()
        if 'Specials' in grp.columns:
            for sv in grp['Specials'].dropna():
                for code in str(sv).split(','):
                    code = code.strip().upper()
                    if code:
                        spec_codes.add(code)
        g = {'main': main, 'rooms': rooms, 'specials': spec_codes}
        row = g['main']
        name = _rc_clean_name(row.get('Name'))
        if not name:
            continue
        company = str(row.get('Company')) if pd.notna(row.get('Company')) else ''
        # ── Ô SPECIAL REQUEST ──
        # Mặc định: AI Lunch, AI Dinner, Minibar set up.
        # CELERIS → thêm ( EUR ) sau Lunch & Dinner.
        # Specials có CN → thêm Connecting Room; có EB → thêm Extra Bed.
        if _norm_nat(company).startswith('celeris'):
            sr = 'AI Lunch ( EUR ), AI Dinner ( EUR ), Minibar set up'
        else:
            sr = 'AI Lunch, AI Dinner, Minibar set up'
        _spec = g.get('specials', set())
        if 'CN' in _spec:
            sr += ', Connecting Room'
        if 'EB' in _spec:
            sr += ', Extra Bed'

        data = {
            'name': name,
            'conf': _rc_conf(row.get('Conf#')),
            'arrival': _rc_date(row.get('Arrival')),
            'departure': _rc_date(row.get('Departure')),
            'nights': _rc_nights(row.get('Arrival'), row.get('Departure')),
            'type': str(row.get('Type')) if pd.notna(row.get('Type')) else '',
            'rm': ', '.join(g['rooms']),   # gộp các số phòng
            'company': company,
            'special': sr,
        }
        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(595,841))
        c.setFillColor(white)
        for x0,top,x1,bot in BLANK:
            c.rect(x0, H-bot, (x1-x0), (bot-top), fill=1, stroke=0)
        c.setFillColor(black)
        c.setFont(FONT, SIZE)
        MAXW = {'name': 300, 'company': 200}  # ô 1 dòng: thu nhỏ nếu tràn
        for key,(x,bottom) in POS.items():
            val = data[key]
            if not val:
                continue
            if key == 'rm':
                # Ô số phòng: 1 phòng → vị trí chuẩn như mẫu gốc.
                # Nhiều phòng → căn giữa CHÍNH XÁC giữa 2 nhãn (đo từ 2 ảnh crop độc lập,
                # sai lệch <1pt): nhãn "ROOM NO./(Số phòng)" kết thúc ~261pt,
                # nhãn "ROOM RATE/(Giá phòng)" bắt đầu ~441pt.
                # → Vùng vẽ 266–436pt, TÂM = 351pt, rộng 170pt
                #   (đủ ~8 phòng/dòng ở cỡ chữ nguyên vẹn 9.8pt → 24 phòng/3 dòng).
                rooms = [s.strip() for s in val.split(',') if s.strip()]
                RM_X0, RM_X1 = 266.0, 436.0
                RM_CX = (RM_X0 + RM_X1) / 2        # = 351.0
                RM_MAXW = RM_X1 - RM_X0            # = 170pt mỗi dòng
                def _wrap(items, fs):
                    lines=[]; cur=''
                    for it in items:
                        test = (cur + ', ' + it) if cur else it
                        if c.stringWidth(test, FONT, fs) <= RM_MAXW:
                            cur = test
                        else:
                            if cur: lines.append(cur)
                            cur = it
                    if cur: lines.append(cur)
                    return lines
                if len(rooms) <= 1:
                    c.drawString(x, H - bottom, val)
                else:
                    fs = SIZE
                    lines = _wrap(rooms, fs)
                    if len(lines) <= 3:
                        gap = fs + 1.6             # tới ~24 phòng: cỡ chữ 9.8pt nguyên vẹn
                    else:
                        gap = fs + 0.6             # 4 dòng sát nhau, vẫn nguyên cỡ chữ (~32 phòng)
                        while len(lines) > 4 and fs > 7:
                            fs -= 0.3              # chống tràn tuyệt đối cho case phi thực tế
                            lines = _wrap(rooms, fs)
                    c.setFont(FONT, fs)
                    for i, ln in enumerate(lines):
                        c.drawCentredString(RM_CX, H - bottom - i*gap, ln)
                    c.setFont(FONT, SIZE)
            else:
                maxw = MAXW.get(key)
                if key == 'special':
                    # Ô special: vẽ từ x=137.4, giới hạn mép phải bảng ~573 → rộng ~436pt.
                    # Thu nhỏ nhẹ nếu quá dài (hiếm), giữ nguyên baseline dòng in sẵn.
                    SP_MAXW = 573 - x
                    fs = SIZE
                    while fs > 7 and c.stringWidth(val, FONT, fs) > SP_MAXW:
                        fs -= 0.2
                    c.setFont(FONT, fs)
                    c.drawString(x, H - bottom, val)
                    c.setFont(FONT, SIZE)
                elif maxw:
                    fs = SIZE
                    while fs > 6 and c.stringWidth(val, FONT, fs) > maxw:
                        fs -= 0.3
                    c.setFont(FONT, fs)
                    c.drawString(x, H-bottom, val)
                    c.setFont(FONT, SIZE)
                else:
                    c.drawString(x, H-bottom, val)
        c.save(); buf.seek(0)
        base = PdfReader(io.BytesIO(tmpl_bytes))
        overlay = PdfReader(buf)
        page = base.pages[0]
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
        count += 1
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue(), count

# ── Đối chiếu Smile vs Trang lưu trú người nước ngoài ─────────────────────
def _norm_pp(p):
    """Chuẩn hóa số hộ chiếu."""
    if pd.isna(p): return ''
    s = str(p).strip().upper().replace(' ','')
    if s.endswith('.0'): s = s[:-2]
    return s

def _norm_room(r):
    if pd.isna(r): return ''
    s = str(r).strip()
    if s.endswith('.0'): s = s[:-2]
    return s.upper()  # 12a05 và 12A05 là một phòng

def reconcile(smile_bytes, luutru_bytes, today):
    """Đối chiếu file Smile (inhouse) với file trang quản lý lưu trú.
    today: pd.Timestamp ngày xuất file (hôm nay)."""
    # ── Đọc Smile ──
    df1 = pd.read_excel(io.BytesIO(smile_bytes), header=0)
    smile = df1[['Passport #','NAT','Rm#','Arrival','Last Name','First Name']].copy()
    smile = smile.dropna(subset=['Passport #'])
    smile['pp'] = smile['Passport #'].apply(_norm_pp)
    smile['Arrival'] = pd.to_datetime(smile['Arrival'], errors='coerce')
    smile['name'] = (smile['Last Name'].astype(str).str.strip() + ' ' +
                     smile['First Name'].astype(str).str.strip())
    smile['room'] = smile['Rm#'].apply(_norm_room)
    # Cột Departure (dò tên linh hoạt phòng khi khác tên)
    _dep_col = next((c for c in df1.columns if 'depart' in str(c).lower()), None)
    smile['Departure'] = pd.to_datetime(df1[_dep_col], errors='coerce') if _dep_col else pd.NaT
    # Lọc: bỏ VNM + bỏ arrival = hôm nay + bỏ departure = hôm nay (khách trả phòng)
    smile_f = smile[(smile['NAT'] != 'VNM') &
                    (smile['Arrival'].dt.date != today.date()) &
                    (smile['Departure'].dt.date != today.date())].copy()

    # ── Đọc Lưu trú ──
    df2 = pd.read_excel(io.BytesIO(luutru_bytes), header=9)
    df2 = df2.dropna(subset=['Họ tên'])
    df2['pp'] = df2['Số hộ chiếu'].apply(_norm_pp)
    # Cột ngày đi dự kiến: file mới đổi tên thành "Thời gian dự kiến tạm trú tại CSLT".
    # Dò linh hoạt để chạy được cả file mẫu mới lẫn cũ.
    _ddk_col = next((c for c in df2.columns
                     if 'dự kiến' in str(c).lower() and 'tạm trú' in str(c).lower()), None)
    if _ddk_col is None:
        _ddk_col = next((c for c in df2.columns if 'đi dự kiến' in str(c).lower()), None)
    df2['ddk'] = pd.to_datetime(df2[_ddk_col], format='%d/%m/%Y', errors='coerce') if _ddk_col else pd.NaT
    df2['room'] = df2['Số phòng'].apply(_norm_room)
    # Lọc: bỏ ngày đi dự kiến = hôm nay
    luutru_f = df2[df2['ddk'].dt.date != today.date()].copy()

    smile_pp = set(smile_f['pp'])
    luutru_pp = set(luutru_f['pp'])

    # Đối chiếu người
    chua_dk = smile_f[~smile_f['pp'].isin(luutru_pp)][['name','pp','NAT','room','Arrival']].copy()
    chua_dk.columns = ['Họ tên','Số hộ chiếu','Quốc tịch','Số phòng','Ngày đến']
    chua_dk['Ngày đến'] = chua_dk['Ngày đến'].dt.strftime('%d/%m/%Y')

    thua = luutru_f[~luutru_f['pp'].isin(smile_pp)].copy()
    # Giai đoạn lưu trú: Ngày đến → Ngày đi dự kiến
    _dd = pd.to_datetime(thua['Ngày đến '], format='%d/%m/%Y', errors='coerce') if 'Ngày đến ' in thua else pd.Series([pd.NaT]*len(thua))
    thua['_luutru'] = (_dd.dt.strftime('%d/%m/%Y').fillna('?') + ' → ' +
                       thua['ddk'].dt.strftime('%d/%m/%Y').fillna('?'))
    _qt = thua['QT'] if 'QT' in thua else ''
    thua = pd.DataFrame({
        'Họ tên': thua['Họ tên'].values,
        'Số hộ chiếu': thua['pp'].values,
        'Quốc tịch': _qt.values if hasattr(_qt,'values') else _qt,
        'Số phòng': thua['room'].values,
        'Giai đoạn lưu trú': thua['_luutru'].values,
    })

    dup = luutru_f[luutru_f['pp'].duplicated(keep=False) & (luutru_f['pp']!='')]
    dup = dup[['Họ tên','pp','room']].sort_values('pp').copy()
    dup.columns = ['Họ tên','Số hộ chiếu','Số phòng']

    # Đối chiếu phòng
    smile_rooms = set(r for r in smile_f['room'] if r)
    luutru_rooms = set(r for r in luutru_f['room'] if r)
    def _sortkey(x): return (len(x), x)
    room_chua = sorted(smile_rooms - luutru_rooms, key=_sortkey)
    room_thua = sorted(luutru_rooms - smile_rooms, key=_sortkey)

    return {
        'smile_total': len(smile), 'smile_filtered': len(smile_f),
        'luutru_total': len(df2), 'luutru_filtered': len(luutru_f),
        'chua_dk': chua_dk, 'thua': thua, 'dup': dup,
        'room_chua': room_chua, 'room_thua': room_thua,
        'room_match': len(smile_rooms & luutru_rooms),
        'smile_rooms': len(smile_rooms), 'luutru_rooms': len(luutru_rooms),
    }



# ── UI ────────────────────────────────────────────────────────────────────

# Custom CSS — Dark dev-tool style (sidebar + monospace số liệu)
if not st.session_state.get("_main_css_injected"):
    st.session_state["_main_css_injected"] = True
    components.html("""
<script>
(function(){
    var doc = window.parent.document;
    if (doc.getElementById('main-app-style')) return;
    var css = doc.createElement('style');
    css.id = 'main-app-style';
    css.textContent = `
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header[data-testid="stHeader"] {background: transparent;}

    .stApp {--ease: cubic-bezier(0.4, 0, 0.2, 1); --r-lg: 20px; --r-pill: 999px;}

    /* ── Sidebar (kiểu One UI: nav dạng pill nổi) ── */
    section[data-testid="stSidebar"] {
        background: #0e1116; border-right: 1px solid #1c2128;
    }
    section[data-testid="stSidebar"] .stButton button {
        background: transparent; border: 1px solid transparent;
        border-radius: var(--r-pill);
        color: #9ea7b3; font-weight: 500; text-align: left;
        justify-content: flex-start; padding: 0.55rem 0.9rem;
        box-shadow: none; will-change: transform, background;
        transform: translateX(0);
        transition: background 0.1s var(--ease), color 0.1s var(--ease),
                    border-color 0.1s var(--ease), transform 0.1s var(--ease);
    }
    section[data-testid="stSidebar"] .stButton button:hover {
        background: #161b22; color: #e6e8eb; border-color: #1c2128;
        transform: translateX(4px);
    }
    section[data-testid="stSidebar"] .stButton button[kind="primary"] {
        background: rgba(45,212,191,0.16) !important; color: #2dd4bf !important;
        border-color: transparent !important; font-weight: 600;
    }
    section[data-testid="stSidebar"] .stButton button[kind="primary"]:hover {
        background: rgba(45,212,191,0.24) !important;
    }
    .sb-brand {padding: 0.2rem 0.4rem 1rem;}
    .sb-brand-title {color:#e6e8eb; font-weight:600; font-size:0.95rem; letter-spacing:-0.01em;}
    .sb-mascot {display:inline-block; width:18px; height:18px; vertical-align:-4px; margin-right:2px;
        animation: flowerSway 3.4s ease-in-out infinite;}
    .sb-mascot svg {width:100%; height:100%;}
    .sb-brand-sub {color:#4d5561; font-size:0.72rem;}
    .sb-section {color:#4d5561; font-size:0.68rem; font-weight:600;
        text-transform:uppercase; letter-spacing:0.06em; padding:0.7rem 0.4rem 0.2rem;}
    .sb-status {display:flex; align-items:center; gap:6px; color:#4d5561;
        font-size:0.72rem; padding-top:0.7rem; margin-top:0.5rem; border-top:1px solid #1c2128;}
    .sb-dot {
        position: relative; width:6px; height:6px; border-radius:50%;
        background:#3fb950; flex-shrink:0;
    }
    .sb-dot::after {
        content: ""; position: absolute; inset: 0; border-radius: 50%;
        background: #3fb950;
        animation: pulse 2.2s var(--ease) infinite;
    }
    @keyframes pulse {
        0%   {transform: scale(1);   opacity: 0.55;}
        70%  {transform: scale(2.6); opacity: 0;}
        100% {transform: scale(2.6); opacity: 0;}
    }

    /* ── Banner chào mừng kiểu One UI: thẻ nổi, bo lớn, kính mờ, đuôi bong bóng thoại ── */
    .welcome-banner {
        position: relative;
        display: flex; align-items: center; gap: 14px;
        background: linear-gradient(135deg, #16232a, #131a22);
        border: 1px solid rgba(45,212,191,0.25);
        border-radius: var(--r-lg); padding: 1rem 1.3rem; margin-bottom: 1.4rem;
    }
    .welcome-banner::after {
        content: ""; position: absolute; bottom: -7px; left: 42px;
        width: 14px; height: 14px; background: #16232a;
        border-left: 1px solid rgba(45,212,191,0.25);
        border-bottom: 1px solid rgba(45,212,191,0.25);
        transform: rotate(-45deg); border-radius: 0 0 0 3px;
    }
    .welcome-emoji {width: 40px; height: 40px; flex-shrink: 0; animation: flowerSway 3.4s ease-in-out infinite;}
    .welcome-emoji svg {width: 100%; height: 100%;}
    @keyframes flowerSway {
        0%, 100% {transform: rotate(-8deg);}
        50%      {transform: rotate(8deg);}
    }
    .welcome-title {
        position: relative; overflow: hidden;
        font-weight: 600; font-size: 1.1rem; letter-spacing: -0.01em;
        background: linear-gradient(90deg, #5eead4, #7dd3fc 55%, #f9a8d4);
        -webkit-background-clip: text; background-clip: text; color: transparent;
    }
    .welcome-title span {display: inline-block; color: #f5f7fa; will-change: transform, opacity;}
    @keyframes wtLetterIn {
        from {opacity: 0; transform: scale(0.9);}
        to   {opacity: 1; transform: scale(1);}
    }
    .welcome-title::after {
        content: ""; position: absolute; top: 0; left: -30%; width: 24%; height: 100%;
        background: linear-gradient(100deg, transparent, rgba(255,255,255,0.55), transparent);
        transform: skewX(-20deg); opacity: 0; pointer-events: none;
    }
    .welcome-title.wt-shimmer::after {
        animation: wtShimmer 1.1s ease-out forwards;
    }
    @keyframes wtShimmer {
        0%   {left: -30%; opacity: 0;}
        12%  {opacity: 1;}
        100% {left: 130%; opacity: 0;}
    }
    .welcome-sub {color: #9ea7b3; font-size: 0.82rem; margin-top: 2px;}

    /* ── Nội dung chính ── */
    .section-label {
        font-size: 0.7rem; font-weight: 650; color: #7d8590;
        text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 0.75rem;
    }
    div[data-testid="stFileUploader"] {
        border: 1px dashed #2a3038; border-radius: var(--r-lg); padding: 0.4rem;
        transform: translateY(0); will-change: transform;
        transition: border-color 0.12s var(--ease), background 0.12s var(--ease), transform 0.15s var(--ease);
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #2dd4bf; background: rgba(45,212,191,0.05);
        transform: translateY(-2px);
    }
    /* Khung "panel" — dùng khi bọc st.container(border=True) */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background: #161b22; border: 1px solid #262b33 !important;
        border-radius: var(--r-lg); transform: translateY(0);
        will-change: transform;
        transition: border-color 0.12s var(--ease), transform 0.15s var(--ease), box-shadow 0.15s var(--ease);
    }
    div[data-testid="stVerticalBlockBorderWrapper"]:hover {
        border-color: #3a4048 !important;
        transform: translateY(-2px);
        box-shadow: 0 10px 22px rgba(0,0,0,0.28);
    }
    .stButton button, .stDownloadButton button {
        border-radius: var(--r-pill); font-weight: 550; will-change: transform;
        transform: perspective(500px) rotateX(0deg) translateY(0);
        transition: transform 0.12s var(--ease), background 0.12s var(--ease),
                    border-color 0.12s var(--ease), box-shadow 0.12s var(--ease);
    }
    .stButton button:hover, .stDownloadButton button:hover {
        transform: perspective(500px) rotateX(4deg) translateY(-1px);
        box-shadow: 0 8px 16px rgba(0,0,0,0.3);
    }
    .stButton button:active, .stDownloadButton button:active {
        transform: perspective(500px) rotateX(-3deg) scale(0.97) translateY(0);
        box-shadow: none;
    }
    div[data-testid="stMainBlockContainer"] .stButton button[kind="primary"], .stDownloadButton button {
        background: #2dd4bf; border-color: #2dd4bf; color: #04342c;
        position: relative; overflow: hidden;
    }
    div[data-testid="stMainBlockContainer"] .stButton button[kind="primary"]:hover, .stDownloadButton button:hover {
        background: #26b8a5; border-color: #26b8a5;
        box-shadow: 0 8px 18px rgba(45,212,191,0.35);
    }
    div[data-testid="stMainBlockContainer"] .stButton button[kind="primary"]::after, .stDownloadButton button::after {
        content: ""; position: absolute; top: 0; left: -60%; width: 35%; height: 100%;
        background: linear-gradient(120deg, transparent, rgba(255,255,255,0.55), transparent);
        transform: skewX(-20deg) translateX(0);
        transition: transform 0.65s ease;
        pointer-events: none; will-change: transform;
    }
    div[data-testid="stMainBlockContainer"] .stButton button[kind="primary"]:hover::after, .stDownloadButton button:hover::after {
        transform: skewX(-20deg) translateX(540%);
    }

    /* ── Input / checkbox: viền sáng dần khi hover / focus ── */
    .stTextInput input, .stNumberInput input, .stTextArea textarea {
        transition: border-color 0.1s var(--ease), box-shadow 0.1s var(--ease);
    }
    .stTextInput input:hover, .stNumberInput input:hover, .stTextArea textarea:hover {
        border-color: #3a4048 !important;
    }
    .stTextInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus {
        border-color: #2dd4bf !important; box-shadow: 0 0 0 1px rgba(45,212,191,0.35) !important;
    }
    .stCheckbox {transition: opacity 0.15s var(--ease);}
    .stCheckbox:hover {opacity: 0.85;}

    /* ── Metric: thẻ có viền + dải màu trên cùng, xoay vòng theo vị trí cột ── */
    div[data-testid="stMetric"] {
        background: #161b22; border: 1px solid #262b33; border-radius: var(--r-lg);
        padding: 0.9rem 0.75rem 0.75rem; position: relative; overflow: hidden;
        will-change: transform;
        transform: translateY(0) translateZ(0);
        transition: transform 0.15s var(--ease), border-color 0.12s var(--ease), box-shadow 0.15s var(--ease);
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-4px) translateZ(6px);
        border-color: #3a4048;
        box-shadow: 0 14px 26px rgba(0,0,0,0.38), 0 0 22px rgba(45,212,191,0.12);
    }
    div[data-testid="stMetric"]::before {
        content: ""; position: absolute; top: 0; left: 0; right: 0; height: 4px;
        transform: scaleY(0.75); transform-origin: top;
        transition: transform 0.2s var(--ease);
    }
    div[data-testid="stMetric"]:hover::before {transform: scaleY(1);}
    div[data-testid="stHorizontalBlock"] > div:nth-of-type(5n+1) div[data-testid="stMetric"]::before {background: #2dd4bf;}
    div[data-testid="stHorizontalBlock"] > div:nth-of-type(5n+2) div[data-testid="stMetric"]::before {background: #60a5fa;}
    div[data-testid="stHorizontalBlock"] > div:nth-of-type(5n+3) div[data-testid="stMetric"]::before {background: #f5a623;}
    div[data-testid="stHorizontalBlock"] > div:nth-of-type(5n+4) div[data-testid="stMetric"]::before {background: #f472b6;}
    div[data-testid="stHorizontalBlock"] > div:nth-of-type(5n+5) div[data-testid="stMetric"]::before {background: #a78bfa;}
    div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] * {
        font-family: ui-monospace, "SFMono-Regular", Menlo, monospace !important;
        font-weight: 600 !important; color: #f5f7fa !important;
        transition: color 0.2s var(--ease);
    }
    div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] * {
        font-size: 0.84rem !important; color: #b8c0cc !important; font-weight: 500 !important;
    }

    /* ── Alert (success/info/warning/error): xuất hiện mượt ── */
    div[data-testid="stAlert"] {
        transition: transform 0.1s var(--ease);
    }
    div[data-testid="stAlert"]:hover {transform: translateY(-1px);}

    /* Đã bỏ animation mờ dần toàn màn hình để chuyển tab/tương tác nhanh nhất có thể */
`;
    doc.head.appendChild(css);
})();
</script>
""", height=0)

# ── Hoa anh đào rơi nhẹ nhàng ở nền toàn bộ web (liên tục, không chỉ lúc khởi động) ──
# Chỉ tiêm 1 lần duy nhất trong session — các cánh hoa tự rơi vô hạn bằng CSS
# animation thuần transform/opacity (chạy trên compositor, không tốn hiệu năng
# dù chạy mãi mãi, đúng nguyên tắc đã tối ưu ở các phần khác của app).
if not st.session_state.get("_bg_sakura_injected"):
    st.session_state["_bg_sakura_injected"] = True
    components.html("""
<script>
(function(){
    var doc = window.parent.document;
    if (doc.getElementById('bg-sakura-layer')) return;

    var css = doc.createElement('style');
    css.id = 'bg-sakura-style';
    css.textContent = `
      #bg-sakura-layer { position: fixed; inset: 0; pointer-events: none; overflow: hidden; z-index: 0; }
      #bg-sakura-layer .petal {
          position: absolute; top: -20px; opacity: 0.5; will-change: transform;
          animation-name: bgSakuraFall; animation-timing-function: linear; animation-iteration-count: infinite;
      }
      @keyframes bgSakuraFall {
          0%   { transform: translate(0,0) rotate(0deg); }
          100% { transform: translate(var(--drift), 112vh) rotate(360deg); }
      }
    `;
    doc.head.appendChild(css);

    var layer = doc.createElement('div');
    layer.id = 'bg-sakura-layer';
    var colors = ['#ffd6e8', '#ffc2dd', '#ffe3ef'];
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
        // delay âm: cánh hoa bắt đầu ở giữa chừng vòng rơi, tránh cảm giác tất cả rơi đồng loạt lúc mới tải trang
        p.style.animationDelay = (-Math.random()*dur) + 's';
        layer.appendChild(p);
    }
    doc.body.insertBefore(layer, doc.body.firstChild);
})();
</script>
""", height=0)

# ── Hiệu ứng khởi động kiểu điện thoại: "Welcome, Tân" rồi mờ dần vào app ──
# QUAN TRỌNG: chỉ gọi components.html() đúng 1 lần trong session (guard bằng
# session_state phía Python) — nếu không, Streamlit sẽ tạo lại iframe này ở
# MỌI lần rerun (mọi cú click), dù JS bên trong có tự bỏ qua, việc tạo lại
# iframe + gửi lại toàn bộ HTML/JS qua lại vẫn tốn thời gian và gây khựng.
# ── Hiệu ứng chữ "Welcome, Tân" mờ dần mượt + vệt sáng lướt kiểu logo boot Samsung ──
# Chỉ chạy 1 lần khi mở web (không lặp lại mỗi lần chuyển tab — tránh lặp lại
# lỗi khựng đã tối ưu trước đó). Tự dò tìm phần tử vài lần vì banner có thể
# chưa kịp render ngay lúc script này chạy.
if not st.session_state.get("_welcome_reveal_done"):
    st.session_state["_welcome_reveal_done"] = True
    components.html("""
<script>
(function(){
    var doc = window.parent.document;
    // Streamlit thường render lại 2-3 lần liên tiếp rất nhanh lúc mới mở web
    // (khởi tạo session/widget) — mỗi lần render lại sẽ tạo banner MỚI (chưa
    // chạy hiệu ứng), thay thế banner cũ. Nên trong 3 giây đầu, hễ thấy banner
    // MỚI (chưa có cờ 'revealed') là áp hiệu ứng lại, đảm bảo bắt đúng bản
    // banner CUỐI CÙNG thực sự hiển thị cho người dùng.
    function reveal(){
        var el = doc.querySelector('.welcome-title');
        if (!el || el.dataset.revealed) return;
        el.dataset.revealed = '1';
        var text = el.textContent;
        el.textContent = '';
        var frag = doc.createDocumentFragment();
        var n = text.length;
        text.split('').forEach(function(ch, i){
            var span = doc.createElement('span');
            span.textContent = (ch === ' ') ? '\\u00A0' : ch;
            span.style.opacity = '0';
            span.style.animation = 'wtLetterIn 0.7s ease-out forwards';
            span.style.animationDelay = (i * 0.08) + 's';
            frag.appendChild(span);
        });
        el.appendChild(frag);
        // Sau khi chữ cuối cùng hiện xong mới cho vệt sáng lướt qua (giống logo boot thật)
        var totalMs = (n * 80) + 700 + 150;
        window.parent.setTimeout(function(){ el.classList.add('wt-shimmer'); }, totalMs);
    }
    var settleUntil = Date.now() + 3000;
    var iv = window.parent.setInterval(function(){
        reveal();
        if (Date.now() > settleUntil) window.parent.clearInterval(iv);
    }, 150);
})();
</script>
""", height=0)

if not st.session_state.get("_boot_splash_done"):
    st.session_state["_boot_splash_done"] = True
    components.html("""
<script>
(function(){
    var doc = window.parent.document;
    if (doc.getElementById('boot-splash')) return;

    var css = doc.createElement('style');
    css.id = 'boot-splash-style';
    css.textContent = `
      @import url('https://fonts.googleapis.com/css2?family=Patrick+Hand&display=swap');
      #boot-splash {
        position: fixed; inset: 0; z-index: 999999;
        background: #0b0d12;
        display: flex; flex-direction: column; align-items: center; justify-content: center;
        transition: opacity 0.55s ease;
      }
      #boot-splash.bs-hide { opacity: 0; pointer-events: none; }
      #boot-splash .bs-logo-wrap {
        position: relative; width: 100px; height: 100px;
        opacity: 0; transform: scale(0.5);
        animation: bsLogoIn 0.55s cubic-bezier(0.34,1.56,0.64,1) 0.15s forwards;
      }
      #boot-splash .bs-logo-wrap svg { width: 100%; height: 100%; }
      #boot-splash .bs-sparkle {
        position: absolute; color: #ffb3d1; opacity: 0;
        animation: bsTwinkle 1.6s ease-in-out infinite;
      }
      #boot-splash .bs-sparkle.s1 { top: -6px;  left: -18px; font-size: 1.1rem; animation-delay: 0.7s; }
      #boot-splash .bs-sparkle.s2 { top: 10px;  right: -22px; font-size: 0.85rem; animation-delay: 1.2s; }
      #boot-splash .bs-sparkle.s3 { bottom: -4px; left: -8px; font-size: 0.7rem; animation-delay: 1.6s; }
      @keyframes bsTwinkle {
        0%, 100% {opacity: 0; transform: scale(0.6) rotate(0deg);}
        50%      {opacity: 1; transform: scale(1.15) rotate(20deg);}
      }
      #boot-splash .bs-text {
        margin-top: 20px;
        font-family: 'ChocoCooky', 'Choco Cooky', 'Patrick Hand', cursive;
        font-size: 3.2rem; font-weight: 700;
        background: linear-gradient(90deg, #5eead4, #7dd3fc 55%, #f9a8d4);
        -webkit-background-clip: text; background-clip: text; color: transparent;
        transform: translateY(10px) rotate(-2deg); transform-origin: center;
        opacity: 0;
        animation: bsTextIn 0.5s ease 0.55s forwards;
      }
      #boot-splash .bs-sub {
        margin-top: 8px; font-family: -apple-system, "Segoe UI", Roboto, sans-serif;
        font-size: 1rem; color: #8b95a1;
        opacity: 0;
        animation: bsTextIn 0.5s ease 0.85s forwards;
      }
      @keyframes bsLogoIn { to {opacity: 1; transform: scale(1);} }
      @keyframes bsTextIn { to {opacity: 1; transform: translateY(0) rotate(-2deg);} }
      #boot-splash .sakura {
        position: absolute; top: -24px; will-change: transform, opacity;
        animation-name: sakuraFall; animation-timing-function: linear; animation-fill-mode: forwards;
      }
      @keyframes sakuraFall {
        0%   {transform: translate(0,0) rotate(0deg);   opacity: 0.95;}
        85%  {opacity: 0.9;}
        100% {transform: translate(var(--drift), 100vh) rotate(360deg); opacity: 0;}
      }
    `;
    doc.head.appendChild(css);

    var el = doc.createElement('div');
    el.id = 'boot-splash';
    el.innerHTML =
        '<div class="bs-logo-wrap">' +
          '<span class="bs-sparkle s1">&#10022;</span>' +
          '<span class="bs-sparkle s2">&#10022;</span>' +
          '<span class="bs-sparkle s3">&#10022;</span>' +
          '<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">' +
            '<g>' +
              '<path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#ffc2dd"/>' +
              '<path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#ffb3d1" transform="rotate(72 50 50)"/>' +
              '<path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#ffc2dd" transform="rotate(144 50 50)"/>' +
              '<path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#ffb3d1" transform="rotate(216 50 50)"/>' +
              '<path d="M50 50 C38 36 40 14 46 9 C48 7 52 7 54 9 C60 14 62 36 50 50 Z" fill="#ffc2dd" transform="rotate(288 50 50)"/>' +
              '<circle cx="50" cy="50" r="7" fill="#fff6ee"/>' +
              '<circle cx="47" cy="47" r="1.3" fill="#ffcf6b"/>' +
              '<circle cx="53" cy="47" r="1.3" fill="#ffcf6b"/>' +
              '<circle cx="50" cy="52.5" r="1.3" fill="#ffcf6b"/>' +
            '</g>' +
          '</svg>' +
        '</div>' +
        '<div class="bs-text">Welcome, Tân</div>' +
        '<div class="bs-sub">Tân Hotel &middot; Front Office toolkit</div>';
    doc.body.appendChild(el);

    // ── Hoa anh đào rơi nhẹ nhàng khắp màn hình khởi động ──
    var petalColors = ['#ffd6e8','#ffc2dd','#ffe3ef'];
    for (var i = 0; i < 18; i++) {
        var p = doc.createElement('div');
        p.className = 'sakura';
        var size = 9 + Math.random()*8;
        var left = Math.random()*100;
        var dur = 4 + Math.random()*3.5;
        var delay = Math.random()*2.5;
        var drift = (Math.random()*140 - 70) + 'px';
        var color = petalColors[i % petalColors.length];
        p.style.left = left + 'vw';
        p.style.width = size + 'px';
        p.style.height = size + 'px';
        p.style.background = 'radial-gradient(circle at 30% 30%, #fff, ' + color + ' 70%)';
        p.style.borderRadius = '0 60% 0 60%';
        p.style.setProperty('--drift', drift);
        p.style.animationDuration = dur + 's';
        p.style.animationDelay = delay + 's';
        el.appendChild(p);
    }

    setTimeout(function(){
        el.classList.add('bs-hide');
        setTimeout(function(){ if (el.parentNode) el.parentNode.removeChild(el); }, 600);
    }, 1900);
})();
</script>
""", height=0)



# Menu selection (session state)
if "menu" not in st.session_state:
    st.session_state.menu = "daily"

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
    st.button("Xử lý hàng ngày", key="nav_daily", use_container_width=True,
              icon=":material/checklist:",
              type="primary" if st.session_state.menu == "daily" else "secondary",
              on_click=go_menu, args=("daily",))
    st.button("Regcard + ARR", key="nav_regcard", use_container_width=True,
              icon=":material/print:",
              type="primary" if st.session_state.menu == "regcard" else "secondary",
              on_click=go_menu, args=("regcard",))

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

    st.markdown('<div class="sb-status"><span class="sb-dot"></span>Sẵn sàng</div>', unsafe_allow_html=True)

st.markdown('''
<div class="welcome-banner">
    <div class="welcome-emoji">
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
    </div>
    <div>
        <div class="welcome-title">Welcome, Tân</div>
        <div class="welcome-sub">Front Office toolkit · Tân Hotel</div>
    </div>
</div>
''', unsafe_allow_html=True)

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
        'dummy': dummy_count,
        'ca_the': sum(1 for x in result if x['type'] == 'sep' and x['conf'] == 'CÀ THẺ'),
        'thu_tien': sum(1 for x in result if x['type'] == 'sep' and x['conf'] == 'THU TIỀN'),
        'xem_lai_bu': sum(1 for x in result if x['type'] == 'sep' and x['conf'] == 'XEM LẠI BU'),
        'foc_lco': sum(1 for x in result if x['type'] == 'sep' and x['conf'].startswith('FOC LATE C/O')),
    }
    return wb, stats


# ── Daily processing screen ───────────────────────────────────────────────
if st.session_state.menu == "daily":
    st.write("")
    st.markdown('<div class="section-label">⚙️ Cài đặt</div>', unsafe_allow_html=True)
    with st.container(border=True):
        col1, col2 = st.columns(2)
        with col1:
            rate = st.number_input("💱 Tỷ giá USD/EUR → VNĐ", value=29535.15, step=0.01, format="%.2f")
        with col2:
            today = datetime.date.today()
            date_str = st.text_input("📅 Ngày (dùng cho tên file)", value=f"{today.day}_{today.month:02d}")

    st.write("")
    st.markdown('<div class="section-label">📂 Tải file lên</div>', unsafe_allow_html=True)
    with st.container(border=True):
        col_x, col_s = st.columns(2)
        with col_x:
            xlsx_file = st.file_uploader("File XLSX — Dữ liệu khách (bắt buộc)", type=['xlsx'], key="daily_xlsx")
        with col_s:
            xls_file = st.file_uploader("File XLS — Nguồn ĐK14 (tùy chọn)", type=['xls'], key="daily_xls")

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
                visa_map = {}; visa_unmatched = []
                # Đọc file visa (nếu có) → map tên → date visa
                if visa_file is not None:
                    try:
                        visa_map = parse_visa_file(visa_file.read())
                    except Exception as _ve:
                        st.warning(f"⚠️ Không đọc được file visa: {_ve}")

                with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    # ── Xử lý file XLSX (nếu có) ──
                    if has_xlsx:
                        progress.progress(10, text="Quy đổi tỷ giá...")
                        xlsx_bytes = xlsx_file.read()
                        wb, conv = process_xlsx(xlsx_bytes, rate)
                        df = pd.read_excel(io.BytesIO(xlsx_bytes))
                        df_intl = df[df['LOẠI KHÁCH']=='Quốc tế'].reset_index(drop=True)
                        df_vn   = df[df['LOẠI KHÁCH']=='Việt Nam'].reset_index(drop=True)

                        progress.progress(30, text="Tách file Quốc tế / Việt Nam...")
                        wb_intl = split_wb(wb, 'Quốc tế')
                        wb_vn   = split_wb(wb, 'Việt Nam')

                        progress.progress(45, text="Điền mẫu KBTT...")
                        wb_kbtt, visa_unmatched = build_kbtt(df_intl, visa_map=visa_map)

                        progress.progress(60, text="Điền mẫu Thông báo lưu trú VNM...")
                        wb_vnm, gks_cnt, gbl_cnt = build_vnm(df_vn)

                        zf.writestr(f'converted_{date_str}.xlsx',      wb_to_bytes(wb))
                        zf.writestr(f'KhachQuocTe_{date_str}.xlsx',    wb_to_bytes(wb_intl))
                        zf.writestr(f'KhachVietNam_{date_str}.xlsx',   wb_to_bytes(wb_vn))
                        zf.writestr(f'ho_so_KBTT_NNN_{date_str}.xlsx', wb_to_bytes(wb_kbtt))
                        zf.writestr(f'thong_bao_luu_tru_VNM_{date_str}.xlsx', wb_to_bytes(wb_vnm))
                        files_made += ["📄 converted (file chung)", "🌍 KhachQuocTe", "🇻🇳 KhachVietNam",
                                       "📝 KBTT NNN", "📑 Thông báo lưu trú VNM"]

                    # ── Xử lý file ĐK14 (độc lập, chỉ cần file XLS) ──
                    if xls_file:
                        progress.progress(85, text="Điền mẫu ĐK14...")
                        xls_bytes = xls_file.read()
                        wb_dk14, dk_count = build_dk14(xls_bytes)
                        zf.writestr(f'dk14_{date_str}.xlsx', wb_to_bytes(wb_dk14))
                        has_dk14 = True
                        files_made.append("🚔 ĐK14")

                progress.progress(100, text="Hoàn tất!")
                progress.empty()

                # Lưu kết quả vào session — kết quả & nút tải không biến mất sau rerun
                _daily = {'files_made': files_made, 'zip': zip_buf.getvalue(),
                          'date_str': date_str, 'has_xlsx': has_xlsx, 'has_dk14': has_dk14}
                if has_xlsx:
                    unknown_nats = []
                    for q in df_intl.get('QUỐC TỊCH', pd.Series([], dtype=str)).dropna().unique():
                        mapped = lookup_nat_kbtt(q)
                        if not _re.match(r'^[A-Z]{2,3} - ', str(mapped)):
                            unknown_nats.append(str(q))
                    _daily.update({'total': len(df), 'intl': len(df_intl), 'vn': len(df_vn),
                                   'gks': gks_cnt, 'gbl': gbl_cnt, 'conv': conv,
                                   'unknown_nats': unknown_nats,
                                   'visa_used': bool(visa_map),
                                   'visa_matched': len(df_intl) - len(visa_unmatched) if visa_map else 0,
                                   'visa_unmatched': visa_unmatched,
                                   'visa_skipped_vn': visa_map.get('skipped_vn', 0) if isinstance(visa_map, dict) else 0})
                st.session_state['daily_results'] = _daily
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
            st.info(f"💱 Đã quy đổi tỷ giá cho **{_dr['conv']}** ô (đã tô vàng)")
            if _dr['unknown_nats']:
                st.warning("⚠️ Quốc tịch chưa có mã (giữ nguyên tên, cần kiểm tra): " + ", ".join(_dr['unknown_nats']))
            if _dr.get('visa_used'):
                st.info(f"🛂 Đã điền date visa cho **{_dr['visa_matched']}/{_dr['intl']}** khách quốc tế (khớp theo tên).")
                if _dr.get('visa_skipped_vn'):
                    st.caption(f"ℹ️ Đã tự động bỏ qua {_dr['visa_skipped_vn']} khách Việt Nam trong file visa (không cần thời hạn tạm trú).")
                if _dr.get('visa_unmatched'):
                    st.warning("⚠️ Không tìm thấy date visa cho (cột tạm trú để trống): "
                               + ", ".join(_dr['visa_unmatched']))
        elif _dr['has_dk14']:
            st.info("ℹ️ Chỉ tạo file ĐK14 (không có file XLSX dữ liệu khách).")

        st.markdown("**File đã tạo:** " + " · ".join(_dr['files_made']))

        st.download_button(
            label="⬇️ Tải về tất cả file (ZIP)",
            data=_dr['zip'],
            file_name=f"hotel_{_dr['date_str']}.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary"
        )

# ── Regcard screen ────────────────────────────────────────────────────────
if st.session_state.menu == "regcard":
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
                    'date': datetime.date.today().strftime('%d_%m'),
                    'arr_date': datetime.date.today().strftime('%d.%m.%Y'),
                }
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

if st.session_state.menu == "recon":
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

# ── Kiểm tra lưu trú người nước ngoài ─────────────────────────────────────
if st.session_state.menu == "recon_person":
    st.write("")
    st.markdown('<div class="section-label">🌏 Kiểm tra lưu trú người nước ngoài</div>', unsafe_allow_html=True)
    st.caption("So khớp khách Smile vs Trang quản lý người nước ngoài theo số hộ chiếu — tìm khách chưa đăng ký / đăng ký trùng.")

    rc1, rc2 = st.columns(2)
    with rc1:
        smile_file = st.file_uploader("File Smile — Inhouse (.xlsx)", type=['xlsx'], key="recon_smile")
    with rc2:
        luutru_file = st.file_uploader("File Trang lưu trú người nước ngoài (.xlsx)", type=['xlsx'], key="recon_luutru")

    today_str = st.text_input("📅 Ngày xuất file (hôm nay)", value=datetime.date.today().strftime('%d/%m/%Y'),
                              help="Dùng để loại bỏ: khách arrival hôm nay (Smile) và khách ngày đi dự kiến hôm nay (Lưu trú)")

    st.write("")

    if st.button("🔍 Bắt đầu kiểm tra", type="primary",
                 disabled=(smile_file is None or luutru_file is None), use_container_width=True):
        with st.spinner("Đang đối chiếu..."):
            try:
                today = pd.to_datetime(today_str, format='%d/%m/%Y')
                st.session_state['recon_results'] = reconcile(smile_file.read(), luutru_file.read(), today)
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


if st.session_state.menu == "recon_room":
    st.write("")
    st.markdown('<div class="section-label">🚪 Kiểm tra hệ thống quản lý lưu trú phòng</div>', unsafe_allow_html=True)
    st.caption("So khớp số phòng inhouse từ file khách lưu trú Smile với file danh sách số phòng — tìm phòng chưa đăng ký / thừa / trùng.")

    rr1, rr2 = st.columns(2)
    with rr1:
        smile_file_r = st.file_uploader("File khách lưu trú Smile (.xlsx)", type=['xlsx'], key="reconr_smile")
    with rr2:
        room_file = st.file_uploader("File số phòng (.xlsx — chỉ chứa danh sách số phòng)", type=['xlsx'], key="reconr_room")

    today_str_r = st.text_input("📅 Ngày xuất file (hôm nay)", value=datetime.date.today().strftime('%d/%m/%Y'),
                                key="reconr_today",
                                help="Khách có Arrival = ngày này trên Smile sẽ được loại bỏ khỏi đối chiếu")

    st.write("")

    if st.button("🔍 Bắt đầu kiểm tra", type="primary", key="reconr_run",
                 disabled=(smile_file_r is None or room_file is None), use_container_width=True):
        with st.spinner("Đang đối chiếu phòng..."):
            try:
                today_r = pd.to_datetime(today_str_r, format='%d/%m/%Y')
                st.session_state['reconr_results'] = reconcile_rooms(smile_file_r.read(), room_file.read(), today_r)
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


