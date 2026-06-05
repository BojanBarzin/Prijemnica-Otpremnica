import re
import base64
from io import BytesIO
from datetime import date

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Prijemnica / Otpremnica sa terena", layout="wide")

BRAND_YELLOW = "#FFD700"
GRAPHITE = "#111111"
LIGHT_GRAY = "#F5F5F5"

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

CITY_HINTS = [
    "Beograd", "Novi Sad", "Niš", "Nis", "Kragujevac", "Subotica", "Zrenjanin", "Pančevo", "Pancevo",
    "Čačak", "Cacak", "Kraljevo", "Kruševac", "Krusevac", "Leskovac", "Valjevo", "Šabac", "Sabac",
    "Sombor", "Kikinda", "Užice", "Uzice", "Vršac", "Vrsac", "Loznica", "Smederevo", "Požarevac", "Pozarevac",
    "Jagodina", "Paraćin", "Paracin", "Zaječar", "Zajecar", "Bor", "Pirot", "Vranje", "Prokuplje", "Ruma",
]

# =========================
# ASSETS / STYLE
# =========================
def get_base64(path: str) -> str:
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""

bg_logo = get_base64("assets/fs_logo_white.png")

st.markdown(f"""
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
}}
.brand-header {{
    background: rgba(17,17,17,0.96);
    padding: 22px 28px;
    border-radius: 18px;
    margin-bottom: 26px;
    border-left: 10px solid {BRAND_YELLOW};
    box-shadow: 0 8px 24px rgba(0,0,0,0.18);
}}
.brand-title {{
    color: white;
    font-size: 32px;
    font-weight: 900;
    margin: 0;
}}
.brand-subtitle {{
    color: #d9d9d9;
    font-size: 15px;
    margin-top: 4px;
}}
.stTextInput input,
.stNumberInput input,
.stDateInput input {{
    background-color: white !important;
    color: black !important;
    border: 1px solid #d0d0d0 !important;
    border-radius: 10px !important;
}}
div[data-baseweb="select"] > div {{
    background-color: white !important;
    color: black !important;
    border-radius: 10px !important;
    border: 1px solid #d0d0d0 !important;
}}
div.stButton > button,
button[kind="secondary"],
[data-testid="stNumberInput"] button,
[data-testid="stNumberInput"] div button {{
    background: {GRAPHITE} !important;
    color: white !important;
    border: 1px solid {GRAPHITE} !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
}}
div.stButton > button:hover,
button[kind="secondary"]:hover,
[data-testid="stNumberInput"] button:hover,
[data-testid="stNumberInput"] div button:hover {{
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
.doc-card {{
    background: white;
    border: 1px solid #dedede;
    border-radius: 14px;
    padding: 16px;
    box-shadow: 0 6px 18px rgba(0,0,0,0.06);
}}
.doc-title {{
    font-size: 22px;
    font-weight: 900;
    text-align: center;
    margin-bottom: 12px;
}}
.doc-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
}}
.doc-table td, .doc-table th {{
    border: 1px solid #111;
    padding: 6px;
    vertical-align: middle;
}}
.doc-label {{
    background: #f2f2f2;
    font-weight: 800;
    width: 28%;
}}
.doc-value {{
    min-height: 20px;
}}
.suggestions {{
    background: white;
    border: 1px solid #e0e0e0;
    border-radius: 10px;
    padding: 8px;
    margin-top: -8px;
    margin-bottom: 10px;
}}
.suggestion-chip {{
    display: inline-block;
    border: 1px solid #ddd;
    border-radius: 999px;
    padding: 4px 9px;
    margin: 3px;
    background: #fafafa;
    font-size: 12px;
}}
</style>
""", unsafe_allow_html=True)

# =========================
# LOAD DATA
# =========================
@st.cache_data
def load_cmdb() -> pd.DataFrame:
    try:
        return pd.read_excel("data.xlsx", dtype=str).fillna("")
    except Exception:
        return pd.DataFrame()

@st.cache_data
def load_locations() -> pd.DataFrame:
    paths = ["LocationsSPTS.csv", "LocationsSPTS(1).csv"]
    for path in paths:
        try:
            return pd.read_csv(path, dtype=str).fillna("")
        except Exception:
            pass
    return pd.DataFrame()

cmdb_df = load_cmdb()
locations_df = load_locations()

if cmdb_df.empty:
    st.warning("Nije pronađen data.xlsx ili je fajl prazan.")
if locations_df.empty:
    st.warning("Nije pronađen LocationsSPTS.csv ili je fajl prazan.")

# =========================
# NORMALIZATION HELPERS
# =========================
def norm(value) -> str:
    return str(value or "").strip()

def norm_lower(value) -> str:
    return norm(value).lower()

def get_col(df: pd.DataFrame, names: list[str]) -> str | None:
    lower_map = {c.lower(): c for c in df.columns}
    for name in names:
        if name.lower() in lower_map:
            return lower_map[name.lower()]
    return None

CMDB_COLS = {
    "number": get_col(cmdb_df, ["number", "Config Item", "ConfigItem", "CI"]),
    "name": get_col(cmdb_df, ["name", "Name"]),
    "vendor": get_col(cmdb_df, ["vendor", "Vendor"]),
    "model": get_col(cmdb_df, ["model", "Model"]),
    "serial": get_col(cmdb_df, ["serial_number", "SerialNumber", "serial", "SN"]),
    "inventory": get_col(cmdb_df, ["inventory_number", "InventoryNumber", "inventory", "INV"]),
    "spfs": get_col(cmdb_df, ["sp_inventory_number", "SPInventoryNumber", "SP/FS", "SPFS"]),
    "project": get_col(cmdb_df, ["project", "Project"]),
    "project_name": get_col(cmdb_df, ["project_name", "Project Name"]),
}

LOC_COLS = {
    "number": get_col(locations_df, ["Number", "number"]),
    "name": get_col(locations_df, ["Name", "name"]),
    "address": get_col(locations_df, ["Address", "address", "adress", "Adresa"]),
    "project": get_col(locations_df, ["Project", "project"]),
}

# =========================
# STATE HELPERS
# =========================
def init_key(key: str, value=""):
    if key not in st.session_state:
        st.session_state[key] = value

FIELDS = [
    # prijemnica
    "pri_uredjaj_razduzio", "pri_broj", "pri_datum", "pri_u_magacin", "pri_uredjaj_razduzio_ime",
    "pri_objekat", "pri_adresa", "pri_mesto", "pri_naziv", "pri_model", "pri_inv", "pri_serial", "pri_spfs", "pri_predao",
    # otpremnica
    "otp_broj", "otp_datum", "otp_iz_magacina", "otp_zaduzio", "otp_objekat", "otp_adresa", "otp_mesto",
    "otp_naziv", "otp_model", "otp_inv", "otp_serial", "otp_spfs", "otp_otpremio",
]

for key in FIELDS:
    init_key(key, date.today() if key in ["pri_datum", "otp_datum"] else "")

init_key("auto_sync_otp", True)

# =========================
# LOOKUP FUNCTIONS
# =========================
def lookup_device_by_any(inv="", serial="", spfs="") -> dict | None:
    if cmdb_df.empty:
        return None

    candidates = []
    if inv and CMDB_COLS["inventory"]:
        candidates.append((CMDB_COLS["inventory"], inv))
    if serial and CMDB_COLS["serial"]:
        candidates.append((CMDB_COLS["serial"], serial))
    if spfs and CMDB_COLS["spfs"]:
        candidates.append((CMDB_COLS["spfs"], spfs))

    for col, val in candidates:
        val_clean = norm_lower(val)
        if not val_clean:
            continue
        exact = cmdb_df[cmdb_df[col].astype(str).str.strip().str.lower() == val_clean]
        if not exact.empty:
            return exact.iloc[0].to_dict()

    for col, val in candidates:
        val_clean = norm_lower(val)
        if not val_clean:
            continue
        partial = cmdb_df[cmdb_df[col].astype(str).str.lower().str.contains(re.escape(val_clean), na=False)]
        if not partial.empty:
            return partial.iloc[0].to_dict()

    return None

def fill_device(prefix: str):
    inv = st.session_state.get(f"{prefix}_inv", "")
    serial = st.session_state.get(f"{prefix}_serial", "")
    spfs = st.session_state.get(f"{prefix}_spfs", "")
    row = lookup_device_by_any(inv=inv, serial=serial, spfs=spfs)
    if not row:
        return

    if CMDB_COLS["name"]:
        st.session_state[f"{prefix}_naziv"] = norm(row.get(CMDB_COLS["name"], ""))
    if CMDB_COLS["model"]:
        st.session_state[f"{prefix}_model"] = norm(row.get(CMDB_COLS["model"], ""))
    if CMDB_COLS["inventory"]:
        st.session_state[f"{prefix}_inv"] = norm(row.get(CMDB_COLS["inventory"], ""))
    if CMDB_COLS["serial"]:
        st.session_state[f"{prefix}_serial"] = norm(row.get(CMDB_COLS["serial"], ""))
    if CMDB_COLS["spfs"]:
        st.session_state[f"{prefix}_spfs"] = norm(row.get(CMDB_COLS["spfs"], ""))

def location_suggestions(text: str, limit=8) -> list[dict]:
    if locations_df.empty or not LOC_COLS["name"]:
        return []
    text = norm_lower(text)
    if not text:
        return []
    name_col = LOC_COLS["name"]
    address_col = LOC_COLS["address"]
    mask = locations_df[name_col].astype(str).str.lower().str.contains(re.escape(text), na=False)
    rows = locations_df[mask].head(limit)
    out = []
    for _, row in rows.iterrows():
        out.append({
            "name": norm(row.get(name_col, "")),
            "address": norm(row.get(address_col, "")) if address_col else "",
            "project": norm(row.get(LOC_COLS["project"], "")) if LOC_COLS["project"] else "",
        })
    return out

def lookup_location_by_name(text: str) -> dict | None:
    suggestions = location_suggestions(text, limit=1)
    return suggestions[0] if suggestions else None

def infer_city(address="", object_name="") -> str:
    combined = f"{address} {object_name}"
    for city in CITY_HINTS:
        if city.lower() in combined.lower():
            return "Niš" if city == "Nis" else city
    return ""

def fill_location(prefix: str):
    obj = st.session_state.get(f"{prefix}_objekat", "")
    loc = lookup_location_by_name(obj)
    if not loc:
        return
    st.session_state[f"{prefix}_objekat"] = loc["name"]
    st.session_state[f"{prefix}_adresa"] = loc["address"]
    city = infer_city(loc["address"], loc["name"])
    if city:
        st.session_state[f"{prefix}_mesto"] = city

def sync_otpremnica_from_prijemnica():
    if not st.session_state.get("auto_sync_otp", True):
        return
    mapping = {
        "otp_broj": "pri_broj",
        "otp_datum": "pri_datum",
        "otp_iz_magacina": "pri_uredjaj_razduzio",
        "otp_zaduzio": "pri_u_magacin",
        "otp_objekat": "pri_objekat",
        "otp_adresa": "pri_adresa",
        "otp_mesto": "pri_mesto",
        "otp_naziv": "pri_naziv",
        "otp_model": "pri_model",
        "otp_inv": "pri_inv",
        "otp_serial": "pri_serial",
        "otp_spfs": "pri_spfs",
        "otp_otpremio": "pri_predao",
    }
    for target, source in mapping.items():
        st.session_state[target] = st.session_state.get(source, "")

def client_suggestions(text: str) -> list[str]:
    text = norm_lower(text)
    if not text:
        return []
    return [c for c in COMMON_CLIENTS if text in c.lower()][:8]

# =========================
# UI HELPERS
# =========================
def suggestion_buttons(prefix: str, field: str, values: list[str]):
    if not values:
        return
    st.markdown("<div class='suggestions'>", unsafe_allow_html=True)
    cols = st.columns(min(len(values), 4))
    for idx, val in enumerate(values):
        with cols[idx % len(cols)]:
            if st.button(val, key=f"suggest_{prefix}_{field}_{idx}_{val}"):
                st.session_state[f"{prefix}_{field}"] = val
                st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

def location_suggestion_buttons(prefix: str):
    obj = st.session_state.get(f"{prefix}_objekat", "")
    suggestions = location_suggestions(obj, limit=6)
    if not suggestions:
        return
    st.caption("Predlozi lokacija iz CSV-a:")
    for idx, loc in enumerate(suggestions):
        label = loc["name"]
        if loc["address"]:
            label += f" — {loc['address']}"
        if st.button(label, key=f"loc_suggest_{prefix}_{idx}"):
            st.session_state[f"{prefix}_objekat"] = loc["name"]
            st.session_state[f"{prefix}_adresa"] = loc["address"]
            city = infer_city(loc["address"], loc["name"])
            if city:
                st.session_state[f"{prefix}_mesto"] = city
            st.rerun()

def draw_document_preview(title: str, data: dict, kind: str):
    if kind == "pri":
        rows = [
            ("Uređaj razdužio", data.get("uredjaj_razduzio", "")),
            ("Broj prijemnice", data.get("broj", "")),
            ("Datum", data.get("datum", "")),
            ("U magacin / Ime prezime", data.get("u_magacin", "")),
            ("Uređaj razdužio / Ime prezime", data.get("uredjaj_razduzio_ime", "")),
            ("Objekat", data.get("objekat", "")),
            ("Adresa", data.get("adresa", "")),
            ("Mesto", data.get("mesto", "")),
            ("Naziv", data.get("naziv", "")),
            ("Model", data.get("model", "")),
            ("Inventarni broj", data.get("inv", "")),
            ("Serijski broj", data.get("serial", "")),
            ("SP/FS broj", data.get("spfs", "")),
            ("Uređaj predao", data.get("predao", "")),
        ]
    else:
        rows = [
            ("Broj otpremnice", data.get("broj", "")),
            ("Datum", data.get("datum", "")),
            ("Iz magacina / Ime prezime", data.get("iz_magacina", "")),
            ("Uređaj zadužio / Ime prezime", data.get("zaduzio", "")),
            ("Objekat", data.get("objekat", "")),
            ("Adresa", data.get("adresa", "")),
            ("Mesto", data.get("mesto", "")),
            ("Naziv", data.get("naziv", "")),
            ("Model", data.get("model", "")),
            ("Inventarni broj", data.get("inv", "")),
            ("Serijski broj", data.get("serial", "")),
            ("SP/FS broj", data.get("spfs", "")),
            ("Uređaj otpremio", data.get("otpremio", "")),
        ]
    html_rows = "".join(f"<tr><td class='doc-label'>{a}</td><td class='doc-value'>{b}</td></tr>" for a, b in rows)
    st.markdown(f"""
    <div class="doc-card">
        <div class="doc-title">{title}</div>
        <table class="doc-table">{html_rows}</table>
    </div>
    """, unsafe_allow_html=True)

def fmt_date(value):
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    return norm(value)

def collect_prijemnica():
    return {
        "uredjaj_razduzio": st.session_state.get("pri_uredjaj_razduzio", ""),
        "broj": st.session_state.get("pri_broj", ""),
        "datum": fmt_date(st.session_state.get("pri_datum", "")),
        "u_magacin": st.session_state.get("pri_u_magacin", ""),
        "uredjaj_razduzio_ime": st.session_state.get("pri_uredjaj_razduzio_ime", ""),
        "objekat": st.session_state.get("pri_objekat", ""),
        "adresa": st.session_state.get("pri_adresa", ""),
        "mesto": st.session_state.get("pri_mesto", ""),
        "naziv": st.session_state.get("pri_naziv", ""),
        "model": st.session_state.get("pri_model", ""),
        "inv": st.session_state.get("pri_inv", ""),
        "serial": st.session_state.get("pri_serial", ""),
        "spfs": st.session_state.get("pri_spfs", ""),
        "predao": st.session_state.get("pri_predao", ""),
    }

def collect_otpremnica():
    return {
        "broj": st.session_state.get("otp_broj", ""),
        "datum": fmt_date(st.session_state.get("otp_datum", "")),
        "iz_magacina": st.session_state.get("otp_iz_magacina", ""),
        "zaduzio": st.session_state.get("otp_zaduzio", ""),
        "objekat": st.session_state.get("otp_objekat", ""),
        "adresa": st.session_state.get("otp_adresa", ""),
        "mesto": st.session_state.get("otp_mesto", ""),
        "naziv": st.session_state.get("otp_naziv", ""),
        "model": st.session_state.get("otp_model", ""),
        "inv": st.session_state.get("otp_inv", ""),
        "serial": st.session_state.get("otp_serial", ""),
        "spfs": st.session_state.get("otp_spfs", ""),
        "otpremio": st.session_state.get("otp_otpremio", ""),
    }

# =========================
# EXCEL EXPORT
# =========================
def make_excel(pri: dict, otp: dict) -> bytes:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Prijemnica"
    ws2 = wb.create_sheet("Otpremnica")

    def style_sheet(ws, title, rows):
        thin = Side(style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        header_fill = PatternFill("solid", fgColor="F2F2F2")
        title_fill = PatternFill("solid", fgColor="111111")

        ws.merge_cells("A1:B1")
        ws["A1"] = title
        ws["A1"].font = Font(bold=True, size=18, color="FFFFFF")
        ws["A1"].fill = title_fill
        ws["A1"].alignment = Alignment(horizontal="center")

        r = 3
        for label, value in rows:
            ws.cell(r, 1).value = label
            ws.cell(r, 2).value = value
            ws.cell(r, 1).font = Font(bold=True)
            ws.cell(r, 1).fill = header_fill
            ws.cell(r, 1).border = border
            ws.cell(r, 2).border = border
            ws.cell(r, 1).alignment = Alignment(vertical="center")
            ws.cell(r, 2).alignment = Alignment(vertical="center")
            r += 1

        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 55

    pri_rows = [
        ("Uređaj razdužio", pri["uredjaj_razduzio"]),
        ("Broj prijemnice", pri["broj"]),
        ("Datum", pri["datum"]),
        ("U magacin / Ime prezime", pri["u_magacin"]),
        ("Uređaj razdužio / Ime prezime", pri["uredjaj_razduzio_ime"]),
        ("Objekat", pri["objekat"]),
        ("Adresa", pri["adresa"]),
        ("Mesto", pri["mesto"]),
        ("Naziv", pri["naziv"]),
        ("Model", pri["model"]),
        ("Inventarni broj", pri["inv"]),
        ("Serijski broj", pri["serial"]),
        ("SP/FS broj", pri["spfs"]),
        ("Uređaj predao", pri["predao"]),
    ]
    otp_rows = [
        ("Broj otpremnice", otp["broj"]),
        ("Datum", otp["datum"]),
        ("Iz magacina / Ime prezime", otp["iz_magacina"]),
        ("Uređaj zadužio / Ime prezime", otp["zaduzio"]),
        ("Objekat", otp["objekat"]),
        ("Adresa", otp["adresa"]),
        ("Mesto", otp["mesto"]),
        ("Naziv", otp["naziv"]),
        ("Model", otp["model"]),
        ("Inventarni broj", otp["inv"]),
        ("Serijski broj", otp["serial"]),
        ("SP/FS broj", otp["spfs"]),
        ("Uređaj otpremio", otp["otpremio"]),
    ]

    style_sheet(ws1, "PRIJEMNICA", pri_rows)
    style_sheet(ws2, "OTPREMNICA", otp_rows)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================
# HEADER
# =========================
col_logo, col_title = st.columns([1, 6])
with col_logo:
    try:
        st.image("assets/fs_logo.png", width=120)
    except Exception:
        pass
with col_title:
    st.markdown("""
    <div class="brand-header">
        <div class="brand-title">Prijemnica / Otpremnica sa terena</div>
        <div class="brand-subtitle">Prediktivni unos uređaja, objekta, adrese i automatsko preslikavanje podataka</div>
    </div>
    """, unsafe_allow_html=True)

# =========================
# FORM UI
# =========================
st.session_state.auto_sync_otp = st.checkbox(
    "Automatski popuni otpremnicu iz prijemnice",
    value=st.session_state.get("auto_sync_otp", True)
)

left, right = st.columns(2)

with left:
    st.subheader("📥 Prijemnica")

    st.text_input("Uređaj razdužio", key="pri_uredjaj_razduzio")
    suggestion_buttons("pri", "uredjaj_razduzio", client_suggestions(st.session_state.get("pri_uredjaj_razduzio", "")))

    st.text_input("Broj prijemnice", key="pri_broj")
    st.date_input("Datum", key="pri_datum")
    st.text_input("U magacin / Ime prezime", key="pri_u_magacin")
    st.text_input("Uređaj razdužio / Ime prezime", key="pri_uredjaj_razduzio_ime")

    st.text_input("Objekat", key="pri_objekat", on_change=lambda: fill_location("pri"))
    location_suggestion_buttons("pri")
    st.text_input("Adresa", key="pri_adresa")
    st.text_input("Mesto", key="pri_mesto")

    st.markdown("---")
    st.caption("Unesi Inventarni broj, Serijski broj ili SP/FS broj — ostala polja uređaja se dopunjavaju iz data.xlsx ako postoji zapis.")
    st.text_input("Inventarni broj", key="pri_inv", on_change=lambda: fill_device("pri"))
    st.text_input("Serijski broj", key="pri_serial", on_change=lambda: fill_device("pri"))
    st.text_input("SP/FS broj", key="pri_spfs", on_change=lambda: fill_device("pri"))
    st.text_input("Naziv", key="pri_naziv")
    st.text_input("Model", key="pri_model")
    st.text_input("Uređaj predao", key="pri_predao")

# sync after prijemnica input render, before otpremnica render
sync_otpremnica_from_prijemnica()

with right:
    st.subheader("📤 Otpremnica")

    st.text_input("Broj otpremnice", key="otp_broj")
    st.date_input("Datum", key="otp_datum")
    st.text_input("Iz magacina / Ime prezime", key="otp_iz_magacina")
    st.text_input("Uređaj zadužio / Ime prezime", key="otp_zaduzio")

    st.text_input("Objekat", key="otp_objekat", on_change=lambda: fill_location("otp"))
    location_suggestion_buttons("otp")
    st.text_input("Adresa", key="otp_adresa")
    st.text_input("Mesto", key="otp_mesto")

    st.markdown("---")
    st.text_input("Inventarni broj", key="otp_inv", on_change=lambda: fill_device("otp"))
    st.text_input("Serijski broj", key="otp_serial", on_change=lambda: fill_device("otp"))
    st.text_input("SP/FS broj", key="otp_spfs", on_change=lambda: fill_device("otp"))
    st.text_input("Naziv", key="otp_naziv")
    st.text_input("Model", key="otp_model")
    st.text_input("Uređaj otpremio", key="otp_otpremio")

st.markdown("---")
pri_data = collect_prijemnica()
otp_data = collect_otpremnica()

prev1, prev2 = st.columns(2)
with prev1:
    draw_document_preview("PRIJEMNICA", pri_data, "pri")
with prev2:
    draw_document_preview("OTPREMNICA", otp_data, "otp")

st.markdown("---")
excel_bytes = make_excel(pri_data, otp_data)
st.download_button(
    "📥 Preuzmi prijemnicu i otpremnicu Excel",
    data=excel_bytes,
    file_name="prijemnica_otpremnica_teren.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
