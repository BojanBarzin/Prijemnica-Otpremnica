import base64
from datetime import date
from io import BytesIO
import os
import re

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import streamlit.components.v1 as components

st.set_page_config(page_title="Prijemnica / Otpremnica", layout="wide")

BRAND_YELLOW = "#FFD700"
GRAPHITE = "#111111"
LIGHT_GRAY = "#F5F5F5"

DATA_FILE = "data.xlsx"
LOCATIONS_FILE = "LocationsSPTS.csv"
TEMPLATE_XLSX = "Prijamnica-otpremnica-template.xlsx"

COMMON_CLIENTS = [
    "Tendam",
    "Deichmann",
    "Takko",
    "Mercator-S",
    "H&M",
    "Metre Cash & Carry",
    "Ikea",
    "Decathlon",
    "Lidl",
]

PROJECT_LABELS = {
    "107": "Tendam",
    "108": "Deichmann",
    "109": "Takko",
    "112": "Mercator-S",
    "115": "H&M",
    "118": "Metre Cash & Carry",
    "119": "Ikea",
    "123": "Decathlon",
    "193": "Lidl",
}


def get_base64(path: str) -> str:
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""


bg_logo = get_base64("assets/fs_logo_white.png")


def apply_style():
    st.markdown(
        f"""
        <style>
        html, body, [class*="css"] {{
            font-family: 'Nunito Sans', 'Segoe UI', sans-serif;
        }}
        .stApp {{
            background-color: {LIGHT_GRAY};
            overflow-x: hidden;
        }}
        .stApp::before {{
            content: "";
            position: fixed;
            inset: -300px;
            background-image: url("data:image/png;base64,{bg_logo}");
            background-size: 340px;
            background-repeat: repeat;
            background-position: 0 0;
            opacity: 0.045;
            transform: rotate(-18deg);
            z-index: 0;
            pointer-events: none;
            animation: bgMove 55s linear infinite;
        }}
        @keyframes bgMove {{
            from {{ background-position: 0 0; }}
            to {{ background-position: 900px 900px; }}
        }}
        .block-container {{
            position: relative;
            z-index: 1;
            max-width: 1250px;
            padding-top: 1.4rem;
        }}
        .brand-header {{
            background: rgba(17,17,17,0.96);
            padding: 20px 26px;
            border-radius: 18px;
            margin-bottom: 18px;
            border-left: 10px solid {BRAND_YELLOW};
            box-shadow: 0 8px 24px rgba(0,0,0,0.18);
        }}
        .brand-title {{
            color: white;
            font-size: 31px;
            font-weight: 900;
            margin: 0;
        }}
        .brand-subtitle {{
            color: #d9d9d9;
            font-size: 15px;
            margin-top: 4px;
        }}
        .paper {{
            background: white;
            border: 1px solid #222;
            padding: 18px 20px;
            box-shadow: 0 8px 22px rgba(0,0,0,0.10);
            margin-bottom: 18px;
        }}
        .doc-top {{
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            align-items: start;
            margin-bottom: 8px;
        }}
        .company {{
            text-align: center;
            font-size: 13px;
            font-weight: 600;
        }}
        .doc-title {{
            font-size: 17px;
            font-weight: 800;
            text-align: left;
        }}
        .excel-label {{
            font-size: 13px;
            font-weight: 700;
            color: #111;
            margin-bottom: -0.35rem;
        }}
        .stTextInput label, .stDateInput label, div[data-baseweb="select"] label {{
            font-weight: 700 !important;
            color: #111 !important;
        }}
        .stTextInput input,
        .stDateInput input,
        .stNumberInput input {{
            background-color: white !important;
            color: black !important;
            border: 1px solid #222 !important;
            border-radius: 0 !important;
            min-height: 34px !important;
        }}
        div[data-baseweb="select"] > div {{
            background-color: white !important;
            color: black !important;
            border: 1px solid #222 !important;
            border-radius: 0 !important;
            min-height: 34px !important;
        }}
        div[data-baseweb="select"] svg {{ display: none !important; }}
        div[data-baseweb="select"] [aria-label="open"] {{ display: none !important; }}
        .section-line {{ border-top: 1px solid #222; margin: 12px 0 10px 0; }}
        .item-header {{
            display: grid;
            grid-template-columns: 55px 1.4fr 1.2fr 1.3fr 1.3fr 1.2fr;
            border-top: 1px solid #222;
            border-left: 1px solid #222;
            font-size: 13px;
            font-weight: 800;
            text-align: center;
        }}
        .item-header div {{
            border-right: 1px solid #222;
            border-bottom: 1px solid #222;
            padding: 7px 4px;
        }}
        .sign-row {{
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 18px;
            margin-top: 16px;
        }}
        .sign-box {{
            border-top: 1px solid #222;
            text-align: center;
            padding-top: 6px;
            font-size: 13px;
            font-weight: 700;
            min-height: 44px;
        }}
        div.stButton > button,
        button[kind="secondary"] {{
            background: {GRAPHITE} !important;
            color: white !important;
            border: 1px solid {GRAPHITE} !important;
            border-radius: 10px !important;
            font-weight: 700 !important;
        }}
        div.stButton > button:hover,
        button[kind="secondary"]:hover {{
            background: black !important;
            color: {BRAND_YELLOW} !important;
            border: 1px solid {BRAND_YELLOW} !important;
        }}
        .stDownloadButton > button {{
            background: {BRAND_YELLOW} !important;
            color: black !important;
            border-radius: 10px !important;
            font-weight: 800 !important;
            border: none !important;
        }}
        @media print {{
            .no-print, header, footer, [data-testid="stToolbar"] {{ display: none !important; }}
            .paper {{ box-shadow: none; margin: 0; page-break-inside: avoid; }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


apply_style()

# Header
col_logo, col_title = st.columns([1, 5])
with col_logo:
    try:
        st.image("assets/fs_logo.png", width=120)
    except Exception:
        pass
with col_title:
    st.markdown(
        """
        <div class="brand-header">
            <div class="brand-title">Prijemnica / Otpremnica sa terena</div>
            <div class="brand-subtitle">Prediktivno popunjavanje uređaja i lokacije iz CMDB i SPTS baze</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data
def load_cmdb():
    try:
        data = pd.read_excel(DATA_FILE, dtype=str).fillna("")
    except Exception:
        data = pd.DataFrame()
    return data


@st.cache_data
def load_locations():
    candidates = [
        LOCATIONS_FILE,
        "LocationsSPTS(1).csv",
        "locations.csv",
    ]

    for file_name in candidates:
        if os.path.exists(file_name):
            try:
                return pd.read_csv(file_name, dtype=str).fillna("")
            except Exception:
                continue

    return pd.DataFrame()


cmdb = load_cmdb()
locations = load_locations()

if cmdb.empty:
    st.error(f"Nije pronađen ili je prazan fajl: {DATA_FILE}")
    st.stop()

if locations.empty:
    st.warning(f"Nije pronađen ili je prazan fajl: {LOCATIONS_FILE}. Lokacije neće biti dostupne za predloge.")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out


cmdb = normalize_columns(cmdb)
locations = normalize_columns(locations)


def col_value(row, *names):
    for name in names:
        if name in row.index:
            val = row.get(name, "")
            if pd.notna(val):
                return str(val).strip()
    return ""


def get_options_from_df(df: pd.DataFrame, col: str, fallback=None):
    fallback = fallback or []
    if col not in df.columns:
        return fallback
    values = (
        df[col]
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .drop_duplicates()
        .tolist()
    )
    return sorted(list(dict.fromkeys(values + fallback)))


def map_project_to_name(project_value: str) -> str:
    value = str(project_value).strip()
    return PROJECT_LABELS.get(value, value)


def object_options():
    if locations.empty or "Name" not in locations.columns:
        return []
    return get_options_from_df(locations, "Name")


def object_row_by_name(name: str):
    if locations.empty or "Name" not in locations.columns or not name:
        return None
    exact = locations[locations["Name"].astype(str).str.strip().str.lower() == name.strip().lower()]
    if not exact.empty:
        return exact.iloc[0]
    starts = locations[locations["Name"].astype(str).str.strip().str.lower().str.startswith(name.strip().lower())]
    if len(starts) == 1:
        return starts.iloc[0]
    return None


def infer_city(address: str) -> str:
    """Best effort. CSV nema posebnu kolonu za grad, pa vraćamo prazan string osim ako adresa već sadrži poznat grad."""
    if not address:
        return ""
    known = [
        "Beograd", "Novi Sad", "Niš", "Kragujevac", "Subotica", "Zrenjanin", "Leskovac",
        "Čačak", "Cacak", "Kruševac", "Krusevac", "Pančevo", "Pancevo", "Kraljevo",
        "Užice", "Uzice", "Valjevo", "Sombor", "Kikinda", "Vršac", "Vrsac", "Šabac", "Sabac",
        "Sremska Mitrovica", "Loznica", "Jagodina", "Paraćin", "Paracin", "Pirot", "Zaječar", "Zajecar",
    ]
    a = address.lower()
    for city in known:
        if city.lower() in a:
            return city
    return ""


def find_device(inv: str, serial: str, sp: str):
    filters = []
    if inv and "inventory_number" in cmdb.columns:
        filters.append(("inventory_number", inv))
    if serial and "serial_number" in cmdb.columns:
        filters.append(("serial_number", serial))
    if sp and "sp_inventory_number" in cmdb.columns:
        filters.append(("sp_inventory_number", sp))

    for col, value in filters:
        exact = cmdb[cmdb[col].astype(str).str.strip().str.lower() == value.strip().lower()]
        if not exact.empty:
            return exact.iloc[0]

    for col, value in filters:
        contains = cmdb[cmdb[col].astype(str).str.contains(re.escape(value), case=False, na=False)]
        if len(contains) == 1:
            return contains.iloc[0]

    return None


def device_options(col: str):
    return get_options_from_df(cmdb, col) if col in cmdb.columns else []


def smart_select(label: str, options, key: str, value: str = ""):
    """Searchable input with free text when Streamlit supports accept_new_options."""
    if key not in st.session_state:
        st.session_state[key] = value or None
    try:
        current = st.session_state.get(key)
        index = None
        if current in options:
            index = options.index(current)
        return st.selectbox(
            label,
            options=options,
            index=index,
            placeholder="",
            key=key,
            accept_new_options=True,
        ) or ""
    except TypeError:
        # Fallback za stariji Streamlit
        return st.text_input(label, value=st.session_state.get(key, value) or "", key=f"{key}_text")


def set_default(key, value):
    if value and not st.session_state.get(key):
        st.session_state[key] = value


# =========================
# INPUT STATE INIT
# =========================
def init_defaults():
    defaults = {
        "pri_razduzio": "",
        "pri_broj": "",
        "pri_datum": date.today(),
        "pri_u_magacin": "",
        "pri_uredjaj_razduzio_ime": "",
        "pri_objekat": "",
        "pri_adresa": "",
        "pri_mesto": "",
        "pri_naziv": "",
        "pri_model": "",
        "pri_inv": "",
        "pri_sn": "",
        "pri_sp": "",
        "pri_predao": "",
        "otp_broj": "",
        "otp_datum": date.today(),
        "otp_iz_magacina": "",
        "otp_zaduzio": "",
        "otp_objekat": "",
        "otp_adresa": "",
        "otp_mesto": "",
        "otp_naziv": "",
        "otp_model": "",
        "otp_inv": "",
        "otp_sn": "",
        "otp_sp": "",
        "otp_otpremio": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_defaults()

# Auto-fill device based on current identifiers before rendering the document
matched_device = find_device(
    st.session_state.get("pri_inv", ""),
    st.session_state.get("pri_sn", ""),
    st.session_state.get("pri_sp", ""),
)
if matched_device is not None:
    set_default("pri_naziv", col_value(matched_device, "name", "Name"))
    set_default("pri_model", col_value(matched_device, "model", "Model"))
    set_default("pri_inv", col_value(matched_device, "inventory_number", "InventoryNumber"))
    set_default("pri_sn", col_value(matched_device, "serial_number", "SerialNumber"))
    set_default("pri_sp", col_value(matched_device, "sp_inventory_number", "SPInventoryNumber"))

# Auto-fill object address and project/client
obj_row = object_row_by_name(st.session_state.get("pri_objekat", ""))
if obj_row is not None:
    address = col_value(obj_row, "Address", "address")
    project_name = map_project_to_name(col_value(obj_row, "Project", "project"))
    set_default("pri_adresa", address)
    if not st.session_state.get("pri_mesto"):
        st.session_state["pri_mesto"] = infer_city(address)
    if project_name and not st.session_state.get("pri_razduzio"):
        st.session_state["pri_razduzio"] = project_name

# Mirror prijemnica -> otpremnica when fields are still blank
mirror_map = {
    "otp_broj": "pri_broj",
    "otp_datum": "pri_datum",
    "otp_iz_magacina": "pri_razduzio",
    "otp_zaduzio": "pri_u_magacin",
    "otp_objekat": "pri_objekat",
    "otp_adresa": "pri_adresa",
    "otp_mesto": "pri_mesto",
    "otp_naziv": "pri_naziv",
    "otp_model": "pri_model",
    "otp_inv": "pri_inv",
    "otp_sn": "pri_sn",
    "otp_sp": "pri_sp",
    "otp_otpremio": "pri_predao",
}
for dest, src in mirror_map.items():
    if not st.session_state.get(dest) and st.session_state.get(src):
        st.session_state[dest] = st.session_state[src]

name_opts = device_options("name")
model_opts = device_options("model")
inv_opts = device_options("inventory_number")
serial_opts = device_options("serial_number")
sp_opts = device_options("sp_inventory_number")
obj_opts = object_options()
client_opts = COMMON_CLIENTS


# =========================
# DOCUMENT UI
# =========================
st.markdown('<div class="paper">', unsafe_allow_html=True)

st.markdown(
    """
    <div class="doc-top">
        <div></div>
        <div class="company">Fiscal Solutions d.o.o. &nbsp; Temerinska 102, 21000 Novi Sad</div>
        <div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

c_title, c_broj, c_datum = st.columns([1.4, 1, 1])
with c_title:
    st.markdown('<div class="doc-title">PRIJEMNICA BR.</div>', unsafe_allow_html=True)
with c_broj:
    st.text_input("Broj prijemnice", key="pri_broj")
with c_datum:
    st.date_input("Datum", key="pri_datum")

c1, c2 = st.columns(2)
with c1:
    st.session_state["pri_razduzio"] = smart_select("UREĐAJ RAZDUŽIO (ime i prezime / naziv firme)", client_opts, "pri_razduzio")
with c2:
    st.text_input("U magacin / Ime i prezime", key="pri_u_magacin")

c1, c2, c3 = st.columns([1.4, 1.4, 1])
with c1:
    st.session_state["pri_uredjaj_razduzio_ime"] = st.text_input("Uređaj razdužio / Ime i prezime", key="pri_uredjaj_razduzio_ime")
with c2:
    st.session_state["pri_objekat"] = smart_select("Objekat", obj_opts, "pri_objekat")
with c3:
    st.text_input("Mesto", key="pri_mesto")
st.text_input("Adresa", key="pri_adresa")

st.markdown('<div class="item-header"><div>BR</div><div>NAZIV</div><div>MODEL</div><div>INV</div><div>SN</div><div>SP INV</div></div>', unsafe_allow_html=True)
ci, cn, cm, cinv, csn, csp = st.columns([0.35, 1.4, 1.2, 1.3, 1.3, 1.2])
with ci:
    st.text_input("", value="1", disabled=True, label_visibility="collapsed")
with cn:
    st.session_state["pri_naziv"] = smart_select("Naziv", name_opts, "pri_naziv")
with cm:
    st.session_state["pri_model"] = smart_select("Model", model_opts, "pri_model")
with cinv:
    st.session_state["pri_inv"] = smart_select("Inventarni broj", inv_opts, "pri_inv")
with csn:
    st.session_state["pri_sn"] = smart_select("Serijski broj", serial_opts, "pri_sn")
with csp:
    st.session_state["pri_sp"] = smart_select("SP/FS broj", sp_opts, "pri_sp")

s1, s2, s3 = st.columns(3)
with s1:
    st.text_input("Uređaj predao", key="pri_predao")
with s2:
    st.text_input("Uređaj zadužio", key="pri_uredjaj_razduzio_ime_bottom")
with s3:
    st.text_input("Uređaj zaprimio", key="pri_u_magacin_bottom")

st.markdown('<div class="section-line"></div>', unsafe_allow_html=True)

st.markdown(
    """
    <div class="doc-top">
        <div></div>
        <div class="company">Fiscal Solutions d.o.o. &nbsp; Temerinska 102, 21000 Novi Sad</div>
        <div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

c_title, c_broj, c_datum = st.columns([1.4, 1, 1])
with c_title:
    st.markdown('<div class="doc-title">OTPREMNICA BR.</div>', unsafe_allow_html=True)
with c_broj:
    st.text_input("Broj otpremnice", key="otp_broj")
with c_datum:
    st.date_input("Datum ", key="otp_datum")

c1, c2 = st.columns(2)
with c1:
    st.text_input("Iz magacina / Ime i prezime", key="otp_iz_magacina")
with c2:
    st.text_input("Uređaj zadužio / Ime i prezime", key="otp_zaduzio")

c1, c2, c3 = st.columns([1.4, 1.4, 1])
with c1:
    st.text_input("Objekat ", key="otp_objekat")
with c2:
    st.text_input("Adresa ", key="otp_adresa")
with c3:
    st.text_input("Mesto ", key="otp_mesto")

st.markdown('<div class="item-header"><div>BR</div><div>NAZIV</div><div>MODEL</div><div>INV</div><div>SN</div><div>SP INV</div></div>', unsafe_allow_html=True)
ci, cn, cm, cinv, csn, csp = st.columns([0.35, 1.4, 1.2, 1.3, 1.3, 1.2])
with ci:
    st.text_input("", value="1", disabled=True, label_visibility="collapsed", key="otp_br_disabled")
with cn:
    st.text_input("Naziv ", key="otp_naziv")
with cm:
    st.text_input("Model ", key="otp_model")
with cinv:
    st.text_input("Inventarni broj ", key="otp_inv")
with csn:
    st.text_input("Serijski broj ", key="otp_sn")
with csp:
    st.text_input("SP/FS broj ", key="otp_sp")

s1, s2, s3 = st.columns(3)
with s1:
    st.text_input("Uređaj otpremio", key="otp_otpremio")
with s2:
    st.text_input("Uređaj zadužio ", key="otp_zaduzio_bottom")
with s3:
    st.text_input("Uređaj primio", key="otp_primio_bottom")

st.markdown('</div>', unsafe_allow_html=True)


# =========================
# EXCEL EXPORT
# =========================
def fill_template() -> bytes:
    template_path = TEMPLATE_XLSX
    if not os.path.exists(template_path):
        template_path = "/mnt/data/Prijamnica-otpremnica-template.xlsx"
    wb = load_workbook(template_path)
    ws = wb["Interna prijemnica"] if "Interna prijemnica" in wb.sheetnames else wb.active

    def write(cell, value):
        for merged_range in ws.merged_cells.ranges:
            if cell in merged_range:
                coord = merged_range.start_cell.coordinate
                ws[coord] = value
                ws[coord].alignment = Alignment(horizontal="center", vertical="center")
                return
        ws[cell] = value
        ws[cell].alignment = Alignment(horizontal="center", vertical="center")

    # Prijemnica positions based on uploaded template
    write("D3", st.session_state.get("pri_broj", ""))
    write("E4", st.session_state.get("pri_datum", date.today()).strftime("%d.%m.%Y"))
    write("B7", st.session_state.get("pri_razduzio", ""))
    write("E7", st.session_state.get("pri_u_magacin", ""))
    write("B8", st.session_state.get("pri_objekat", ""))
    write("B9", st.session_state.get("pri_adresa", ""))
    write("B10", st.session_state.get("pri_mesto", ""))
    write("B13", st.session_state.get("pri_naziv", ""))
    write("C13", st.session_state.get("pri_model", ""))
    write("D13", st.session_state.get("pri_inv", ""))
    write("E13", st.session_state.get("pri_sn", ""))
    write("F13", st.session_state.get("pri_sp", ""))
    write("B20", st.session_state.get("pri_predao", ""))
    write("D20", st.session_state.get("pri_uredjaj_razduzio_ime", ""))
    write("E20", st.session_state.get("pri_u_magacin", ""))

    # Otpremnica positions based on uploaded template
    write("D26", st.session_state.get("otp_broj", ""))
    write("E27", st.session_state.get("otp_datum", date.today()).strftime("%d.%m.%Y"))
    write("B30", st.session_state.get("otp_iz_magacina", ""))
    write("D30", st.session_state.get("otp_zaduzio", ""))
    write("D31", st.session_state.get("otp_objekat", ""))
    write("D32", st.session_state.get("otp_adresa", ""))
    write("D33", st.session_state.get("otp_mesto", ""))
    write("B36", st.session_state.get("otp_naziv", ""))
    write("C36", st.session_state.get("otp_model", ""))
    write("D36", st.session_state.get("otp_inv", ""))
    write("E36", st.session_state.get("otp_sn", ""))
    write("F36", st.session_state.get("otp_sp", ""))
    write("B43", st.session_state.get("otp_otpremio", ""))
    write("D43", st.session_state.get("otp_zaduzio", ""))
    write("E43", st.session_state.get("otp_zaduzio", ""))

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def build_print_html() -> str:
    def esc(x):
        return str(x or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    return f"""
    <html>
    <head>
        <style>
            @page {{ size: A4 portrait; margin: 10mm; }}
            body {{ font-family: Arial, sans-serif; font-size: 12px; color: #000; }}
            button {{ margin-bottom: 10px; padding: 8px 14px; background:#111; color:white; border:0; border-radius:6px; font-weight:bold; }}
            .doc {{ border: 1px solid #000; padding: 12px; margin-bottom: 18px; }}
            .center {{ text-align:center; font-weight:bold; }}
            .title {{ font-size: 16px; font-weight:bold; margin-top: 12px; }}
            table {{ width:100%; border-collapse:collapse; margin-top:8px; }}
            td, th {{ border:1px solid #000; padding:6px; text-align:center; min-height:20px; }}
            th {{ font-weight:bold; background:#f4f4f4; }}
            .sign {{ display:grid; grid-template-columns:1fr 1fr 1fr; gap:24px; margin-top:42px; }}
            .sig {{ border-top:1px solid #000; text-align:center; padding-top:6px; font-weight:bold; }}
            @media print {{ button {{ display:none; }} }}
        </style>
    </head>
    <body>
        <button onclick="window.print()">Print</button>
        <div class="doc">
            <div class="center">Fiscal Solutions d.o.o. &nbsp; Temerinska 102, 21000 Novi Sad</div>
            <div class="title">PRIJEMNICA BR. {esc(st.session_state.get('pri_broj'))} &nbsp;&nbsp; Datum: {esc(st.session_state.get('pri_datum').strftime('%d.%m.%Y'))}</div>
            <table>
                <tr><th>UREĐAJ RAZDUŽIO</th><th>U magacin / Ime i prezime</th></tr>
                <tr><td>{esc(st.session_state.get('pri_razduzio'))}</td><td>{esc(st.session_state.get('pri_u_magacin'))}</td></tr>
                <tr><th>Objekat</th><th>Adresa</th><th>Mesto</th></tr>
                <tr><td>{esc(st.session_state.get('pri_objekat'))}</td><td>{esc(st.session_state.get('pri_adresa'))}</td><td>{esc(st.session_state.get('pri_mesto'))}</td></tr>
            </table>
            <table>
                <tr><th>BR</th><th>NAZIV</th><th>MODEL</th><th>INV</th><th>SN</th><th>SP INV</th></tr>
                <tr><td>1</td><td>{esc(st.session_state.get('pri_naziv'))}</td><td>{esc(st.session_state.get('pri_model'))}</td><td>{esc(st.session_state.get('pri_inv'))}</td><td>{esc(st.session_state.get('pri_sn'))}</td><td>{esc(st.session_state.get('pri_sp'))}</td></tr>
            </table>
            <div class="sign"><div class="sig">UREĐAJ PREDAO<br>{esc(st.session_state.get('pri_predao'))}</div><div class="sig">UREĐAJ ZADUŽIO<br>{esc(st.session_state.get('pri_uredjaj_razduzio_ime'))}</div><div class="sig">UREĐAJ ZAPRIMIO<br>{esc(st.session_state.get('pri_u_magacin'))}</div></div>
        </div>
        <div class="doc">
            <div class="center">Fiscal Solutions d.o.o. &nbsp; Temerinska 102, 21000 Novi Sad</div>
            <div class="title">OTPREMNICA BR. {esc(st.session_state.get('otp_broj'))} &nbsp;&nbsp; Datum: {esc(st.session_state.get('otp_datum').strftime('%d.%m.%Y'))}</div>
            <table>
                <tr><th>Iz magacina / Ime i prezime</th><th>UREĐAJ ZADUŽIO</th></tr>
                <tr><td>{esc(st.session_state.get('otp_iz_magacina'))}</td><td>{esc(st.session_state.get('otp_zaduzio'))}</td></tr>
                <tr><th>Objekat</th><th>Adresa</th><th>Mesto</th></tr>
                <tr><td>{esc(st.session_state.get('otp_objekat'))}</td><td>{esc(st.session_state.get('otp_adresa'))}</td><td>{esc(st.session_state.get('otp_mesto'))}</td></tr>
            </table>
            <table>
                <tr><th>BR</th><th>NAZIV</th><th>MODEL</th><th>INV</th><th>SN</th><th>SP INV</th></tr>
                <tr><td>1</td><td>{esc(st.session_state.get('otp_naziv'))}</td><td>{esc(st.session_state.get('otp_model'))}</td><td>{esc(st.session_state.get('otp_inv'))}</td><td>{esc(st.session_state.get('otp_sn'))}</td><td>{esc(st.session_state.get('otp_sp'))}</td></tr>
            </table>
            <div class="sign"><div class="sig">UREĐAJ OTPREMIO<br>{esc(st.session_state.get('otp_otpremio'))}</div><div class="sig">UREĐAJ ZADUŽIO<br>{esc(st.session_state.get('otp_zaduzio'))}</div><div class="sig">UREĐAJ PRIMIO<br>{esc(st.session_state.get('otp_zaduzio'))}</div></div>
        </div>
    </body>
    </html>
    """

st.markdown('<div class="no-print">', unsafe_allow_html=True)
col_a, col_b, col_c = st.columns([1, 1, 1])
with col_a:
    try:
        excel_bytes = fill_template()
        st.download_button(
            "Preuzmi Excel prijemnicu / otpremnicu",
            data=excel_bytes,
            file_name="prijemnica_otpremnica_teren.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Ne mogu da napravim Excel. Proveri da li postoji {TEMPLATE_XLSX}. Detalj: {e}")
with col_b:
    if st.button("Prikaži / Print dokument"):
        components.html(build_print_html(), height=1200, scrolling=True)
with col_c:
    if st.button("Osveži automatsko popunjavanje"):
        st.rerun()
st.markdown('</div>', unsafe_allow_html=True)
