import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import base64
import re

st.set_page_config(page_title="Prijemnica / Otpremnica sa terena", layout="wide")

BRAND_YELLOW = "#FFD700"
GRAPHITE = "#111111"
LIGHT_GRAY = "#F5F5F5"

DATA_FILE = "data.xlsx"
LOCATIONS_FILE = "LocationsSPTS.csv"
LOGO_HEADER = "assets/fs_logo.png"
LOGO_BG = "assets/fs_logo_white.png"

# =========================
# BRANDING
# =========================
def get_base64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""

bg_logo = get_base64(LOGO_BG)


def apply_branding():
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

    .document-preview {{
        background: white;
        border: 1px solid #d8d8d8;
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 8px 22px rgba(0,0,0,0.08);
    }}

    .doc-title {{
        text-align: right;
        font-weight: 900;
        font-size: 22px;
    }}

    .doc-meta {{
        border-collapse: collapse;
        width: 100%;
        margin-top: 12px;
    }}

    .doc-meta td {{
        border: 1px solid #222;
        padding: 7px;
        font-size: 13px;
    }}

    .items-table {{
        border-collapse: collapse;
        width: 100%;
        margin-top: 16px;
    }}

    .items-table th, .items-table td {{
        border: 1px solid #222;
        padding: 6px;
        font-size: 12px;
        text-align: center;
    }}

    .items-table th {{
        background: #f2f2f2;
        font-weight: 800;
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
    </style>
    """, unsafe_allow_html=True)


apply_branding()

col_logo, col_title = st.columns([1, 7])
with col_logo:
    try:
        st.image(LOGO_HEADER, width=130)
    except Exception:
        pass
with col_title:
    st.markdown("""
    <div class="brand-header">
        <div class="brand-title">Prijemnica / Otpremnica sa terena</div>
        <div class="brand-subtitle">Prediktivno popunjavanje uređaja iz CMDB-a i objekata iz SPTS lokacija</div>
    </div>
    """, unsafe_allow_html=True)

# =========================
# LOAD DATA
# =========================
@st.cache_data
def load_cmdb():
    try:
        data = pd.read_excel(DATA_FILE, sheet_name="CMDB", dtype=str).fillna("")
    except Exception:
        data = pd.read_excel(DATA_FILE, dtype=str).fillna("")

    data.columns = [str(c).strip() for c in data.columns]
    return data


@st.cache_data
def load_locations():
    try:
        data = pd.read_csv(LOCATIONS_FILE, dtype=str).fillna("")
    except Exception:
        data = pd.DataFrame()

    data.columns = [str(c).strip() for c in data.columns]
    return data


cmdb_df = load_cmdb()
locations_df = load_locations()

if cmdb_df.empty:
    st.error("Nije pronađen ili nije čitljiv data.xlsx.")
    st.stop()

if locations_df.empty:
    st.warning("CSV sa lokacijama nije pronađen ili je prazan. Lokacije neće biti dostupne za prediktivni unos.")


def pick_col(dataframe, candidates):
    lower_map = {str(c).lower().strip(): c for c in dataframe.columns}
    for candidate in candidates:
        key = candidate.lower().strip()
        if key in lower_map:
            return lower_map[key]
    return None


CMDB_COLS = {
    "number": pick_col(cmdb_df, ["number", "Config Item", "ConfigItem", "CI"]),
    "name": pick_col(cmdb_df, ["name", "Name"]),
    "vendor": pick_col(cmdb_df, ["vendor", "Vendor"]),
    "model": pick_col(cmdb_df, ["model", "Model"]),
    "type": pick_col(cmdb_df, ["type", "Type"]),
    "serial": pick_col(cmdb_df, ["serial_number", "SerialNumber", "Serial Number"]),
    "inventory": pick_col(cmdb_df, ["inventory_number", "InventoryNumber", "Inventory Number"]),
    "sp": pick_col(cmdb_df, ["sp_inventory_number", "SPInventoryNumber", "SP Inventory Number"]),
    "project": pick_col(cmdb_df, ["project", "Project"]),
    "project_name": pick_col(cmdb_df, ["project_name", "ProjectName", "Project Name"]),
}

LOC_COLS = {
    "number": pick_col(locations_df, ["Number", "number"]),
    "name": pick_col(locations_df, ["Name", "name"]),
    "address": pick_col(locations_df, ["Address", "address"]),
    "project": pick_col(locations_df, ["Project", "project"]),
}


def clean_value(value):
    if value is None:
        return ""
    return str(value).strip()


def get_options(dataframe, column, fallback=None, limit=5000):
    fallback = fallback or []
    if not column or column not in dataframe.columns:
        return fallback
    values = (
        dataframe[column]
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .drop_duplicates()
        .head(limit)
        .tolist()
    )
    return sorted(list(dict.fromkeys(values + fallback)))


def predictive_select(label, options, key):
    try:
        value = st.selectbox(
            label,
            options=options,
            index=None,
            placeholder="",
            accept_new_options=True,
            key=key
        )
    except TypeError:
        value = st.selectbox(
            label,
            options=[""] + options,
            index=0,
            key=key
        )
    return clean_value(value)


def row_to_device(row):
    if row is None:
        return {
            "Config Item": "",
            "Name": "",
            "Vendor": "",
            "Model": "",
            "Type": "",
            "SPInventoryNumber": "",
            "InventoryNumber": "",
            "SerialNumber": "",
            "Project": "",
            "ProjectName": "",
        }

    return {
        "Config Item": clean_value(row.get(CMDB_COLS["number"], "")) if CMDB_COLS["number"] else "",
        "Name": clean_value(row.get(CMDB_COLS["name"], "")) if CMDB_COLS["name"] else "",
        "Vendor": clean_value(row.get(CMDB_COLS["vendor"], "")) if CMDB_COLS["vendor"] else "",
        "Model": clean_value(row.get(CMDB_COLS["model"], "")) if CMDB_COLS["model"] else "",
        "Type": clean_value(row.get(CMDB_COLS["type"], "")) if CMDB_COLS["type"] else "",
        "SPInventoryNumber": clean_value(row.get(CMDB_COLS["sp"], "")) if CMDB_COLS["sp"] else "",
        "InventoryNumber": clean_value(row.get(CMDB_COLS["inventory"], "")) if CMDB_COLS["inventory"] else "",
        "SerialNumber": clean_value(row.get(CMDB_COLS["serial"], "")) if CMDB_COLS["serial"] else "",
        "Project": clean_value(row.get(CMDB_COLS["project"], "")) if CMDB_COLS["project"] else "",
        "ProjectName": clean_value(row.get(CMDB_COLS["project_name"], "")) if CMDB_COLS["project_name"] else "",
    }


def find_cmdb_match(params):
    exact_priority = [
        ("number", params.get("Config Item", "")),
        ("sp", params.get("SPInventoryNumber", "")),
        ("inventory", params.get("InventoryNumber", "")),
        ("serial", params.get("SerialNumber", "")),
    ]

    for key, value in exact_priority:
        col = CMDB_COLS.get(key)
        value = clean_value(value)
        if col and value:
            matches = cmdb_df[cmdb_df[col].astype(str).str.strip().str.lower() == value.lower()]
            if not matches.empty:
                return matches.iloc[0]

    contains_priority = [
        ("name", params.get("Name", "")),
        ("vendor", params.get("Vendor", "")),
        ("model", params.get("Model", "")),
        ("type", params.get("Type", "")),
    ]

    result = cmdb_df.copy()
    any_filter = False
    for key, value in contains_priority:
        col = CMDB_COLS.get(key)
        value = clean_value(value)
        if col and value:
            any_filter = True
            result = result[result[col].astype(str).str.contains(value, case=False, na=False)]

    if any_filter and not result.empty:
        return result.iloc[0]

    return None


def row_to_location(row):
    if row is None:
        return {
            "Number": "",
            "Name": "",
            "Address": "",
            "Project": "",
            "City": "",
        }

    name = clean_value(row.get(LOC_COLS["name"], "")) if LOC_COLS["name"] else ""
    return {
        "Number": clean_value(row.get(LOC_COLS["number"], "")) if LOC_COLS["number"] else "",
        "Name": name,
        "Address": clean_value(row.get(LOC_COLS["address"], "")) if LOC_COLS["address"] else "",
        "Project": clean_value(row.get(LOC_COLS["project"], "")) if LOC_COLS["project"] else "",
        "City": guess_city_from_name(name),
    }


def guess_city_from_name(name):
    name = clean_value(name)
    if not name:
        return ""

    known_cities = [
        "Beograd", "BG", "Novi Sad", "NS", "Niš", "Nis", "Kragujevac", "Subotica", "Kraljevo",
        "Čačak", "Cacak", "Zrenjanin", "Pančevo", "Pancevo", "Banja Luka", "Sarajevo", "Tuzla",
        "Zenica", "Kruševac", "Krusevac", "Leskovac", "Valjevo", "Šabac", "Sabac", "Užice", "Uzice"
    ]

    for city in known_cities:
        if re.search(rf"\b{re.escape(city)}\b", name, flags=re.IGNORECASE):
            if city == "BG":
                return "Beograd"
            if city == "NS":
                return "Novi Sad"
            if city == "Nis":
                return "Niš"
            return city

    return ""


def find_location_match(params):
    exact_priority = [
        ("number", params.get("Number", "")),
    ]

    for key, value in exact_priority:
        col = LOC_COLS.get(key)
        value = clean_value(value)
        if col and value:
            matches = locations_df[locations_df[col].astype(str).str.strip().str.lower() == value.lower()]
            if not matches.empty:
                return matches.iloc[0]

    result = locations_df.copy()
    any_filter = False
    for key, field in [("name", "Name"), ("address", "Address"), ("project", "Project")]:
        col = LOC_COLS.get(key)
        value = clean_value(params.get(field, ""))
        if col and value:
            any_filter = True
            result = result[result[col].astype(str).str.contains(value, case=False, na=False)]

    if any_filter and not result.empty:
        return result.iloc[0]

    return None


# =========================
# OPTIONS
# =========================
number_options = get_options(cmdb_df, CMDB_COLS["number"])
sp_options = get_options(cmdb_df, CMDB_COLS["sp"])
inventory_options = get_options(cmdb_df, CMDB_COLS["inventory"])
serial_options = get_options(cmdb_df, CMDB_COLS["serial"])
name_options = get_options(cmdb_df, CMDB_COLS["name"])
vendor_options = get_options(cmdb_df, CMDB_COLS["vendor"])
model_options = get_options(cmdb_df, CMDB_COLS["model"])
type_options = get_options(cmdb_df, CMDB_COLS["type"])

location_number_options = get_options(locations_df, LOC_COLS["number"])
location_name_options = get_options(locations_df, LOC_COLS["name"])
location_address_options = get_options(locations_df, LOC_COLS["address"])
location_project_options = get_options(locations_df, LOC_COLS["project"])

# =========================
# DOCUMENT INPUTS
# =========================
st.markdown("---")
left, right = st.columns([1, 1])

with left:
    doc_type = st.radio("Tip dokumenta", ["Prijemnica", "Otpremnica"], horizontal=True)

with right:
    doc_date = st.date_input("Datum", value=date.today())

c1, c2 = st.columns(2)
with c1:
    from_warehouse = st.text_input("Iz magacina / Ime i prezime", value="")
with c2:
    to_warehouse = st.text_input("Magacin / Uređaj zadužio", value="")

# =========================
# LOCATION
# =========================
st.markdown("---")
st.subheader("📍 Objekat / lokacija")

l1, l2, l3, l4 = st.columns(4)
with l1:
    loc_number = predictive_select("Broj lokacije", location_number_options, "loc_number")
with l2:
    loc_name = predictive_select("Naziv objekta", location_name_options, "loc_name")
with l3:
    loc_address = predictive_select("Adresa", location_address_options, "loc_address")
with l4:
    loc_project = predictive_select("Project", location_project_options, "loc_project")

location_row = find_location_match({
    "Number": loc_number,
    "Name": loc_name,
    "Address": loc_address,
    "Project": loc_project,
})
location = row_to_location(location_row)

if location_row is not None:
    st.success(f"Lokacija pronađena: {location['Name']} | {location['Address']}")
else:
    st.info("Unesi makar jedan parametar lokacije da se objekat automatski pronađe.")

lc1, lc2, lc3 = st.columns(3)
with lc1:
    final_object = st.text_input("Objekat", value=location["Name"], key="final_object")
with lc2:
    final_address = st.text_input("Adresa", value=location["Address"], key="final_address")
with lc3:
    final_city = st.text_input("Mesto / grad", value=location["City"], key="final_city")

# =========================
# DEVICES
# =========================
st.markdown("---")
st.subheader("📦 Uređaji")

item_count = st.number_input("Broj uređaja", min_value=1, max_value=30, value=1)

devices = []

for i in range(int(item_count)):
    st.markdown(f"#### Uređaj {i + 1}")

    d1, d2, d3, d4, d5 = st.columns(5)
    with d1:
        input_number = predictive_select("Config Item", number_options, f"dev_number_{i}")
    with d2:
        input_sp = predictive_select("SPInventoryNumber", sp_options, f"dev_sp_{i}")
    with d3:
        input_inventory = predictive_select("InventoryNumber", inventory_options, f"dev_inv_{i}")
    with d4:
        input_serial = predictive_select("SerialNumber", serial_options, f"dev_serial_{i}")
    with d5:
        input_name = predictive_select("Name", name_options, f"dev_name_{i}")

    d6, d7, d8 = st.columns(3)
    with d6:
        input_vendor = predictive_select("Vendor", vendor_options, f"dev_vendor_{i}")
    with d7:
        input_model = predictive_select("Model", model_options, f"dev_model_{i}")
    with d8:
        input_type = predictive_select("Type", type_options, f"dev_type_{i}")

    match_row = find_cmdb_match({
        "Config Item": input_number,
        "SPInventoryNumber": input_sp,
        "InventoryNumber": input_inventory,
        "SerialNumber": input_serial,
        "Name": input_name,
        "Vendor": input_vendor,
        "Model": input_model,
        "Type": input_type,
    })

    device = row_to_device(match_row)

    if match_row is not None:
        st.success(
            f"Pronađeno: {device['Name']} | {device['Vendor']} {device['Model']} | "
            f"SP: {device['SPInventoryNumber']} | INV: {device['InventoryNumber']} | SN: {device['SerialNumber']}"
        )
    else:
        device = {
            "Config Item": input_number,
            "Name": input_name,
            "Vendor": input_vendor,
            "Model": input_model,
            "Type": input_type,
            "SPInventoryNumber": input_sp,
            "InventoryNumber": input_inventory,
            "SerialNumber": input_serial,
            "Project": "",
            "ProjectName": "",
        }

    devices.append(device)

# =========================
# PREVIEW
# =========================
st.markdown("---")
st.subheader("👁️ Pregled dokumenta")

rows_html = ""
for idx, device in enumerate(devices, start=1):
    rows_html += f"""
    <tr>
        <td>{idx}</td>
        <td>{device.get('Name', '')}</td>
        <td>{device.get('Vendor', '')}</td>
        <td>{device.get('Model', '')}</td>
        <td>{device.get('Type', '')}</td>
        <td>{device.get('Config Item', '')}</td>
        <td>{device.get('SPInventoryNumber', '')}</td>
        <td>{device.get('InventoryNumber', '')}</td>
        <td>{device.get('SerialNumber', '')}</td>
    </tr>
    """

st.markdown(f"""
<div class="document-preview">
    <div class="doc-title">{doc_type.upper()}</div>
    <table class="doc-meta">
        <tr>
            <td><b>Datum</b></td><td>{doc_date.strftime('%d.%m.%Y')}</td>
            <td><b>Iz magacina / Ime i prezime</b></td><td>{from_warehouse}</td>
        </tr>
        <tr>
            <td><b>Magacin / Uređaj zadužio</b></td><td>{to_warehouse}</td>
            <td><b>Objekat</b></td><td>{final_object}</td>
        </tr>
        <tr>
            <td><b>Adresa</b></td><td>{final_address}</td>
            <td><b>Mesto</b></td><td>{final_city}</td>
        </tr>
    </table>

    <table class="items-table">
        <tr>
            <th>BR</th>
            <th>NAZIV</th>
            <th>VENDOR</th>
            <th>MODEL</th>
            <th>TYPE</th>
            <th>CONFIG ITEM</th>
            <th>SP</th>
            <th>INV</th>
            <th>SN</th>
        </tr>
        {rows_html}
    </table>
</div>
""", unsafe_allow_html=True)

# =========================
# EXCEL EXPORT
# =========================
def create_document_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = doc_type

    title_fill = PatternFill("solid", fgColor="111111")
    yellow_fill = PatternFill("solid", fgColor="FFD700")
    light_fill = PatternFill("solid", fgColor="F2F2F2")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:I1")
    ws["A1"] = doc_type.upper()
    ws["A1"].font = Font(bold=True, size=18, color="FFFFFF")
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    meta_rows = [
        ("Datum", doc_date.strftime("%d.%m.%Y"), "Iz magacina / Ime i prezime", from_warehouse),
        ("Magacin / Uređaj zadužio", to_warehouse, "Objekat", final_object),
        ("Adresa", final_address, "Mesto", final_city),
    ]

    start_meta = 3
    for r, values in enumerate(meta_rows, start=start_meta):
        for c, value in enumerate(values, start=1):
            ws.cell(r, c).value = value
            ws.cell(r, c).border = border
            ws.cell(r, c).alignment = Alignment(horizontal="center", vertical="center")
            if c in [1, 3]:
                ws.cell(r, c).font = Font(bold=True)
                ws.cell(r, c).fill = light_fill

    headers = ["BR", "NAZIV", "VENDOR", "MODEL", "TYPE", "CONFIG ITEM", "SP", "INV", "SN"]
    start_row = 8

    for c, header in enumerate(headers, start=1):
        cell = ws.cell(start_row, c)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = yellow_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, device in enumerate(devices, start=start_row + 1):
        row_values = [
            r_idx - start_row,
            device.get("Name", ""),
            device.get("Vendor", ""),
            device.get("Model", ""),
            device.get("Type", ""),
            device.get("Config Item", ""),
            device.get("SPInventoryNumber", ""),
            device.get("InventoryNumber", ""),
            device.get("SerialNumber", ""),
        ]

        for c, value in enumerate(row_values, start=1):
            cell = ws.cell(r_idx, c)
            cell.value = value
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    signature_row = start_row + len(devices) + 5
    ws.merge_cells(start_row=signature_row, start_column=1, end_row=signature_row, end_column=3)
    ws.merge_cells(start_row=signature_row, start_column=7, end_row=signature_row, end_column=9)
    ws.cell(signature_row, 1).value = "Robu izdao"
    ws.cell(signature_row, 7).value = "Robu primio"
    ws.cell(signature_row, 1).alignment = Alignment(horizontal="center")
    ws.cell(signature_row, 7).alignment = Alignment(horizontal="center")
    ws.cell(signature_row, 1).font = Font(bold=True)
    ws.cell(signature_row, 7).font = Font(bold=True)

    widths = [8, 26, 18, 22, 20, 18, 16, 18, 22]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

st.markdown("---")
file_name = f"{doc_type.lower()}_teren_{doc_date.strftime('%Y%m%d')}.xlsx"

st.download_button(
    "📥 Preuzmi Excel dokument",
    data=create_document_excel(),
    file_name=file_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
