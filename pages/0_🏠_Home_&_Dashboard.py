import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
from google.oauth2.service_account import Credentials
import datetime
import re
import io
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Monitoring & SLA Control Tower - HPI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# CUSTOM CSS: GRADASI SIDEBAR BIRU MUDA -> BIRU TUA, PUTIH & ORANYE AKSEN
# ==============================================================================
st.markdown("""
<style>
    .stApp {
        background-color: #F4F6F9;
    }
    .block-container {
        padding-top: 1.3rem;
        padding-bottom: 2rem;
        max-width: 96%;
    }

    /* Sidebar Gradasi Biru Muda ke Biru Tua */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0077B6 0%, #023E8A 45%, #0A2540 100%) !important;
    }
    [data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }
    [data-testid="stSidebarNav"] a {
        border-radius: 8px;
        margin: 4px 8px;
        padding: 8px 12px;
        transition: all 0.2s ease;
    }
    [data-testid="stSidebarNav"] a:hover {
        background-color: rgba(255, 255, 255, 0.15) !important;
    }
    [data-testid="stSidebarNav"] a[aria-current="page"] {
        background-color: #FF7A00 !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    /* Pastikan Selectbox & Input Terlihat Jelas */
    div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #0A2540 !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
    }
    div[data-baseweb="select"] span {
        color: #0A2540 !important;
        font-weight: 600 !important;
    }

    /* Top Header Bar */
    .top-header-box {
        background-color: #FFFFFF;
        padding: 14px 22px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(10, 37, 64, 0.06);
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 16px;
        border-left: 5px solid #FF7A00;
    }
    .date-pill {
        background-color: #0A2540;
        color: #FFFFFF;
        padding: 7px 16px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: 600;
        display: inline-flex;
        align-items: center;
        gap: 8px;
    }

    /* 3 Kartu KPI Utama */
    .kpi-card {
        background-color: #FFFFFF;
        border-radius: 14px;
        padding: 20px 22px;
        box-shadow: 0 2px 10px rgba(10, 37, 64, 0.05);
        border: 1px solid #E6ECF2;
        display: flex;
        align-items: center;
        gap: 18px;
        height: 100%;
    }
    .kpi-icon-navy {
        background-color: #0A2540;
        color: #FFFFFF;
        width: 54px;
        height: 54px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 23px;
        flex-shrink: 0;
    }
    .kpi-icon-orange {
        background-color: #FF7A00;
        color: #FFFFFF;
        width: 54px;
        height: 54px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 23px;
        flex-shrink: 0;
    }
    .kpi-icon-blue {
        background-color: #0077B6;
        color: #FFFFFF;
        width: 54px;
        height: 54px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 23px;
        flex-shrink: 0;
    }
    .kpi-title {
        font-size: 13px;
        color: #5A6B7C;
        font-weight: 600;
        margin-bottom: 3px;
    }
    .kpi-value {
        font-size: 22px;
        color: #0A2540;
        font-weight: 800;
        line-height: 1.2;
    }
    .kpi-sub {
        font-size: 12px;
        color: #16A34A;
        font-weight: 600;
        margin-top: 4px;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# HELPER FUNCTIONS: PARSING ANGKA INDONESIA & FORMAT RUPIAH
# ==============================================================================
def clean_numeric(val):
    if pd.isna(val) or val == "":
        return 0.0
    if isinstance(val, int):
        return float(val)
    s = str(val).strip()
    s = re.sub(r'[^\d.,-]', '', s)
    if not s:
        return 0.0

    if "." in s and "," in s:
        if s.rfind(".") < s.rfind(","):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
            s = s.replace(".", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return 0.0

def format_rp(val):
    return f"Rp. {val:,.0f}".replace(",", ".")

# ==============================================================================
# DATA FETCHING DARI GOOGLE SHEETS
# ==============================================================================
@st.cache_data(ttl=60, show_spinner="⏳ Memuat data terbaru dari Google Sheets...")
def fetch_monitoring_data():
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        client = gspread.authorize(creds)

        # 1. Rekap_Monitoring_TK
        df_tk = pd.DataFrame(columns=["Timestamp", "Cabang", "Jumlah TK", "Total tagihan"])
        try:
            sh_tk = client.open_by_url(st.secrets["SPREADSHEET_URL"])
            ws_tk = next((ws for ws in sh_tk.worksheets() if "rekap_monitoring_tk" in ws.title.strip().lower()), None)
            if ws_tk:
                rec_tk = ws_tk.get_all_records()
                if rec_tk:
                    df_tk = pd.DataFrame(rec_tk)
        except Exception as e:
            st.warning(f"Info Sheet TK: {repr(e)}")

        # 2. Rekap_Monitoring_KS (HPI + ICON)
        list_df_ks = []
        try:
            sh_ks = client.open_by_url(st.secrets["SPREADSHEET_URL_KES"])
            ws_ks = next((ws for ws in sh_ks.worksheets() if "rekap_monitoring_ks" in ws.title.strip().lower()), None)
            if ws_ks:
                rec_ks = ws_ks.get_all_records()
                if rec_ks:
                    list_df_ks.append(pd.DataFrame(rec_ks))
        except Exception as e:
            st.warning(f"Info Sheet KS HPI: {repr(e)}")

        try:
            if "SPREADSHEET_URL_KES_ICON" in st.secrets:
                sh_ks_icon = client.open_by_url(st.secrets["SPREADSHEET_URL_KES_ICON"])
                ws_ks_icon = next((ws for ws in sh_ks_icon.worksheets() if "rekap_monitoring_ks" in ws.title.strip().lower()), None)
                if ws_ks_icon:
                    rec_ks_icon = ws_ks_icon.get_all_records()
                    if rec_ks_icon:
                        list_df_ks.append(pd.DataFrame(rec_ks_icon))
        except Exception as e:
            st.warning(f"Info Sheet KS ICON: {repr(e)}")

        df_ks = pd.concat(list_df_ks, ignore_index=True) if list_df_ks else pd.DataFrame(columns=["Timestamp", "Cabang", "Jumlah TK", "Total tagihan"])
        return df_ks, df_tk
    except Exception as e:
        st.error(f"Gagal terhubung ke Google Sheets: {e}")
        return pd.DataFrame(), pd.DataFrame()

def parse_monitoring_df(df_raw, tahun, bulan_angka, limit_tgl):
    if df_raw is None or df_raw.empty or "Timestamp" not in df_raw.columns:
        return pd.DataFrame(columns=["Cabang", "Timestamp_DT", "Tanggal_Upload", "Jumlah TK", "Total tagihan", "Status_SLA"])

    df = df_raw.copy()
    df["Timestamp_DT"] = pd.to_datetime(df["Timestamp"], errors="coerce")
    df = df.dropna(subset=["Timestamp_DT"])
    df = df[(df["Timestamp_DT"].dt.year == int(tahun)) & (df["Timestamp_DT"].dt.month == int(bulan_angka))]

    if df.empty:
        return pd.DataFrame(columns=["Cabang", "Timestamp_DT", "Tanggal_Upload", "Jumlah TK", "Total tagihan", "Status_SLA"])

    df["Cabang"] = df["Cabang"].astype(str).str.strip().str.upper()
    df["Jumlah TK"] = df["Jumlah TK"].apply(clean_numeric).astype(int)
    df["Total tagihan"] = df["Total tagihan"].apply(clean_numeric)

    df = df.sort_values("Timestamp_DT").groupby("Cabang", as_index=False).last()
    df["Tanggal_Upload"] = df["Timestamp_DT"].dt.day
    df["Status_SLA"] = df["Tanggal_Upload"].apply(lambda d: f"✅ ON-TIME (Tgl {d})" if d <= limit_tgl else f"⚠️ LATE (Tgl {d})")
    return df

# ==============================================================================
# GENERATOR EXCEL SLA TAHUNAN (IDENTIK PERSENTASE_PENILAIAN_CABANG.XLSX)
# ==============================================================================
def generate_excel_sla_tahunan(raw_df_ks, raw_df_tk, tahun_target):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    alias_cabang = {
    # BABEL
    "BANGKA BELITUNG": "BABEL", "BABEL": "BABEL",
    # BANTEN
    "BANTEN": "BANTEN",
    # DKI (termasuk DKI2 & PLN Pusat)
    "DKI JAKARTA": "DKI", "DKI": "DKI", "DKI2": "DKI",
    "PLNPUSAT": "DKI", "PLN PUSAT": "DKI",
    # ICON
    "ICON": "ICON",
    # JABAR (termasuk BANDUNG & BANDUNG JABAR)
    "JAWA BARAT": "JABAR", "JABAR": "JABAR",
    "BANDUNG": "JABAR", "BANDUNG JABAR": "JABAR",
    # JAWA TIMUR (termasuk JATIM & JBTT)
    "JAWA TIMUR": "JAWA TIMUR", "JATIM": "JAWA TIMUR", "JBTT": "JAWA TIMUR",
    # DIY / JATENG
    "JAWA TENGAH": "DIY", "JATENG": "DIY", "DIY": "DIY",
    # MEDAN
    "MEDAN": "MEDAN",
    # PADANG (termasuk SUMBAR)
    "PADANG": "PADANG", "SUMBAR": "PADANG",
    # PALEMBANG
    "PALEMBANG": "PALEMBANG",
    # PEKANBARU
    "PEKANBARU": "PEKANBARU",
    # SIBOLGA
    "SIBOLGA NIAS": "SIBOLGA", "SIBOLGA": "SIBOLGA"
}
    daftar_cabang = [
        "BABEL", "BANTEN", "DKI", "ICON", "JABAR", "JAWA TIMUR",
        "DIY", "MEDAN", "PADANG", "PALEMBANG", "PEKANBARU", "SIBOLGA"
    ]
    bulan_list = [
        "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
        "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOPEMBER", "DESEMBER"
    ]

    def extract_tgl_map(df_raw):
        hasil = {}
        if df_raw is None or df_raw.empty or "Timestamp" not in df_raw.columns:
            return hasil
        df_c = df_raw.copy()
        df_c["DT"] = pd.to_datetime(df_c["Timestamp"], errors="coerce")
        df_c = df_c.dropna(subset=["DT"])
        df_c = df_c[df_c["DT"].dt.year == int(tahun_target)]
        if df_c.empty:
            return hasil
        df_c["Cab_Raw"] = df_c["Cabang"].astype(str).str.strip().str.upper()
        df_c["Cab_Std"] = df_c["Cab_Raw"].map(alias_cabang).fillna(df_c["Cab_Raw"])
        df_c["Bulan"] = df_c["DT"].dt.month
        df_c = df_c.sort_values("DT").groupby(["Cab_Std", "Bulan"], as_index=False).last()
        for _, r in df_c.iterrows():
            hasil[(r["Cab_Std"], int(r["Bulan"]))] = int(r["DT"].day)
        return hasil

    map_ks = extract_tgl_map(raw_df_ks)
    map_tk = extract_tgl_map(raw_df_tk)
    map_dplk = {}

    font_title = Font(name="Calibri", size=12, bold=True, color="0A2540")
    font_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_subhdr = Font(name="Calibri", size=10, bold=True, color="0A2540")
    font_data = Font(name="Calibri", size=10)
    font_bold = Font(name="Calibri", size=10, bold=True)

    fill_navy = PatternFill(start_color="0A2540", end_color="0A2540", fill_type="solid")
    fill_orange = PatternFill(start_color="FF7A00", end_color="FF7A00", fill_type="solid")
    fill_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_total = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin = Side(style="thin", color="808080")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")

    konfigurasi_sheet = [
        ("BPJS KES", 2, map_ks, 0.75, [
            (0, 1.00), (1, 0.98), (2, 0.96), (3, 0.93), (4, 0.90), (5, 0.85), (">5", 0.75),
            (-1, 1.02), (-2, 1.04), (-3, 1.06), (-4, 1.08), (-5, 1.10), (-6, 1.15)
        ]),
        ("BPJS TK", 16, map_tk, 0.88, [
            (0, 1.00), (1, 0.98), (2, 0.96), (3, 0.94), (4, 0.92), (5, 0.90), (">5", 0.88),
            (-1, 1.02), (-2, 1.04), (-3, 1.06), (-4, 1.08), (-5, 1.10), (-6, 1.15)
        ]),
        ("DPLK", 16, map_dplk, 0.88, [
            (0, 1.00), (1, 0.98), (2, 0.96), (3, 0.94), (4, 0.92), (5, 0.90), (">5", 0.88),
            (-1, 1.02), (-2, 1.04), (-3, 1.06), (-4, 1.08), (-5, 1.10), (-6, 1.15)
        ]),
    ]

    for sheet_name, target_tgl, data_map, min_pct, tabel_skor in konfigurasi_sheet:
        ws = wb.create_sheet(title=sheet_name)
        ws["A1"] = f"PERSENTASE PENILAIAN KEPATUHAN SLA {sheet_name} - TAHUN {tahun_target}"
        ws["A1"].font = font_title

        ws.merge_cells("A2:A4")
        ws["A2"] = "NO"
        ws.merge_cells("B2:B4")
        ws["B2"] = "CABANG/UNIT"
        ws["C2"] = "TARGET"
        ws["C3"] = "PENGIRIMAN"
        ws["C4"] = "TGL"

        ws.merge_cells("D2:AA2")
        ws["D2"] = f"REALISASI PENGIRIMAN TAHUN {tahun_target}"

        ws.merge_cells("AB2:AB4")
        ws["AB2"] = "TOTAL\n%"

        for col_idx in [1, 2, 3, 28]:
            for r_idx in [2, 3, 4]:
                c = ws.cell(r_idx, col_idx)
                c.fill = fill_navy
                c.font = font_hdr
                c.alignment = align_center
                c.border = border_all

        ws["D2"].fill = fill_orange
        ws["D2"].font = font_hdr
        ws["D2"].alignment = align_center

        pct_cols = []
        for m_idx, nama_bln in enumerate(bulan_list):
            c_tgl = 4 + (m_idx * 2)
            c_pct = c_tgl + 1
            let_tgl = get_column_letter(c_tgl)
            let_pct = get_column_letter(c_pct)
            pct_cols.append(let_pct)

            ws.merge_cells(f"{let_tgl}3:{let_pct}3")
            ws.cell(3, c_tgl, nama_bln).fill = fill_navy
            ws.cell(3, c_tgl).font = font_hdr
            ws.cell(3, c_tgl).alignment = align_center

            ws.cell(4, c_tgl, "TGL").fill = fill_sub
            ws.cell(4, c_tgl).font = font_subhdr
            ws.cell(4, c_tgl).alignment = align_center

            ws.cell(4, c_pct, "%").fill = fill_sub
            ws.cell(4, c_pct).font = font_subhdr
            ws.cell(4, c_pct).alignment = align_center

            for r_h in [2, 3, 4]:
                ws.cell(r_h, c_tgl).border = border_all
                ws.cell(r_h, c_pct).border = border_all

        min_pct_str = f"{int(min_pct*100)}%"
        for idx_cab, nama_cab in enumerate(daftar_cabang, start=1):
            r = 4 + idx_cab
            ws.cell(r, 1, idx_cab).alignment = align_center
            ws.cell(r, 1).font = font_data
            ws.cell(r, 1).border = border_all

            ws.cell(r, 2, nama_cab).alignment = align_left
            ws.cell(r, 2).font = font_bold
            ws.cell(r, 2).border = border_all

            ws.cell(r, 3, target_tgl).alignment = align_center
            ws.cell(r, 3).font = font_data
            ws.cell(r, 3).border = border_all

            for m_num in range(1, 13):
                c_tgl = 4 + ((m_num - 1) * 2)
                c_pct = c_tgl + 1
                let_tgl = get_column_letter(c_tgl)

                tgl_real = data_map.get((nama_cab, m_num), None)
                cell_tgl = ws.cell(r, c_tgl, tgl_real)
                cell_tgl.alignment = align_center
                cell_tgl.font = font_data
                cell_tgl.border = border_all

                formula_pct = (
                    f'=IFERROR(IF({let_tgl}{r}>0,'
                    f'IF(({let_tgl}{r}-$C{r})>5,{min_pct_str},'
                    f'IF(({let_tgl}{r}-$C{r})<-6,115%,'
                    f'VLOOKUP(({let_tgl}{r}-$C{r}),$C$20:$D$32,2,0))),""),"")'
                )
                cell_pct = ws.cell(r, c_pct, formula_pct)
                cell_pct.number_format = "0%"
                cell_pct.alignment = align_center
                cell_pct.font = font_data
                cell_pct.border = border_all

            avg_refs = ",".join([f"{col_let}{r}" for col_let in pct_cols])
            cell_tot = ws.cell(r, 28, f'=IFERROR(AVERAGE({avg_refs}),"")')
            cell_tot.number_format = "0.0%"
            cell_tot.font = font_bold
            cell_tot.alignment = align_center
            cell_tot.fill = fill_total
            cell_tot.border = border_all

        ws.cell(17, 2, "RATA-RATA NASIONAL").font = font_bold
        ws.cell(17, 2).alignment = align_left
        for c_i in range(1, 29):
            ws.cell(17, c_i).fill = fill_sub
            ws.cell(17, c_i).border = border_all

        for col_let in pct_cols + ["AB"]:
            col_idx_num = openpyxl.utils.column_index_from_string(col_let)
            c_avg = ws.cell(17, col_idx_num, f'=IFERROR(AVERAGE({col_let}5:{col_let}16),"")')
            c_avg.number_format = "0.0%"
            c_avg.font = font_bold
            c_avg.alignment = align_center

        ws.cell(19, 2, "Penilaian").font = font_bold
        ws.cell(19, 3, "Selisih Hari").font = font_hdr
        ws.cell(19, 3).fill = fill_navy
        ws.cell(19, 3).alignment = align_center
        ws.cell(19, 3).border = border_all

        ws.cell(19, 4, "%").font = font_hdr
        ws.cell(19, 4).fill = fill_navy
        ws.cell(19, 4).alignment = align_center
        ws.cell(19, 4).border = border_all

        for idx_s, (selisih_val, pct_val) in enumerate(tabel_skor):
            r_s = 20 + idx_s
            c1 = ws.cell(r_s, 3, selisih_val)
            c1.alignment = align_center
            c1.font = font_data
            c1.border = border_all

            c2 = ws.cell(r_s, 4, pct_val)
            c2.number_format = "0%"
            c2.alignment = align_center
            c2.font = font_data
            c2.border = border_all

        ws.column_dimensions["A"].width = 5
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 12
        for c_k in range(4, 28):
            ws.column_dimensions[get_column_letter(c_k)].width = 6.5
        ws.column_dimensions["AB"].width = 10

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ==============================================================================
# LOAD DATA & TOP HEADER BAR (HANYA TANGGAL SAJA)
# ==============================================================================
raw_df_ks, raw_df_tk = fetch_monitoring_data()
wib_now = datetime.datetime.utcnow() + datetime.timedelta(hours=7)

st.markdown(f"""
<div class="top-header-box">
    <div>
        <div style="font-size:19px; font-weight:800; color:#0A2540;">📊 Dashboard Monitoring & SLA Cabang</div>
        <div style="font-size:12.5px; color:#64748B;">PT Haleyora Powerindo • Sinkronisasi Real-Time BPJS Kesehatan, Ketenagakerjaan & DPLK</div>
    </div>
    <div class="date-pill">
        <span>📅 {wib_now.strftime("%d %b %Y, %H:%M WIB")}</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# FILTER BAR
# ==============================================================================
col_f1, col_f2, col_f3, col_f4 = st.columns([1.2, 1.5, 1.5, 1])
with col_f1:
    tahun_opsi = [2025, 2026, 2027, 2028]
    idx_thn = tahun_opsi.index(wib_now.year) if wib_now.year in tahun_opsi else 1
    sel_tahun = st.selectbox("📅 Tahun Monitoring", tahun_opsi, index=idx_thn)

bulan_dict = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}
with col_f2:
    sel_bulan_num = st.selectbox("🗓️ Bulan Target", list(bulan_dict.keys()), index=wib_now.month - 1, format_func=lambda x: bulan_dict[x])
    sel_bulan_label = bulan_dict[sel_bulan_num]

df_ks_parsed = parse_monitoring_df(raw_df_ks, sel_tahun, sel_bulan_num, limit_tgl=2)
df_tk_parsed = parse_monitoring_df(raw_df_tk, sel_tahun, sel_bulan_num, limit_tgl=16)

semua_cabang_aktif = sorted(list(set(df_ks_parsed["Cabang"].tolist() + df_tk_parsed["Cabang"].tolist())))
with col_f3:
    sel_cabang = st.selectbox("🏢 Filter Cabang", ["Semua Cabang"] + semua_cabang_aktif)

with col_f4:
    st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
    if st.button("🔄 Refresh Data", use_container_width=True, type="primary"):
        st.cache_data.clear()
        st.rerun()

if sel_cabang != "Semua Cabang":
    df_ks_filtered = df_ks_parsed[df_ks_parsed["Cabang"] == sel_cabang]
    df_tk_filtered = df_tk_parsed[df_tk_parsed["Cabang"] == sel_cabang]
else:
    df_ks_filtered = df_ks_parsed
    df_tk_filtered = df_tk_parsed

# Hitung 3 Metrik Utama
tot_tagihan_ks = df_ks_filtered["Total tagihan"].sum() if not df_ks_filtered.empty else 0.0
tot_tk_ks = int(df_ks_filtered["Jumlah TK"].sum()) if not df_ks_filtered.empty else 0
tot_tagihan_tk = df_tk_filtered["Total tagihan"].sum() if not df_tk_filtered.empty else 0.0
tot_tk_tk = int(df_tk_filtered["Jumlah TK"].sum()) if not df_tk_filtered.empty else 0

# ==============================================================================
# BARIS 1: 3 KARTU KPI (TAGIHAN KES, TAGIHAN TK, TOTAL TENAGA KERJA)
# ==============================================================================
k1, k2, k3 = st.columns(3)

with k1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-icon-navy">🏥</div>
        <div>
            <div class="kpi-title">Total Tagihan BPJS Kesehatan</div>
            <div class="kpi-value">{format_rp(tot_tagihan_ks)}</div>
            <div class="kpi-sub">▲ {len(df_ks_filtered)} Cabang Tersinkronisasi</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with k2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-icon-orange">👷</div>
        <div>
            <div class="kpi-title">Total Tagihan BPJS Ketenagakerjaan</div>
            <div class="kpi-value">{format_rp(tot_tagihan_tk)}</div>
            <div class="kpi-sub">▲ {len(df_tk_filtered)} Cabang Tersinkronisasi</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with k3:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-icon-blue">👥</div>
        <div>
            <div class="kpi-title">Total Tenaga Kerja Terdaftar</div>
            <div class="kpi-value">{tot_tk_tk:,} TK • {tot_tk_ks:,} Kes</div>
            <div class="kpi-sub">● Aktif Periode {sel_bulan_label} {sel_tahun}</div>
        </div>
    </div>
    """.replace(",", "."), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ==============================================================================
# BARIS 2: GRAFIK TREN REKAPITULASI (LEBAR PENUH TANPA PEMISAH ANEH)
# ==============================================================================
hdr_c1, hdr_c2 = st.columns([2, 1])
with hdr_c1:
    st.subheader(f"📈 Tren Rekapitulasi per Cabang ({sel_bulan_label} {sel_tahun})")
with hdr_c2:
    mode_grafik = st.selectbox(
        "Pilih Kategori Grafik",
        ["Tagihan BPJS Ketenagakerjaan", "Tenaga Kerja BPJS Ketenagakerjaan", "Tagihan BPJS Kesehatan", "Tenaga Kerja BPJS Kesehatan"],
        label_visibility="collapsed"
    )

if mode_grafik == "Tagihan BPJS Ketenagakerjaan":
    df_plot, y_col, is_rp, bar_color = df_tk_filtered.copy(), "Total tagihan", True, "#0A2540"
elif mode_grafik == "Tenaga Kerja BPJS Ketenagakerjaan":
    df_plot, y_col, is_rp, bar_color = df_tk_filtered.copy(), "Jumlah TK", False, "#FF7A00"
elif mode_grafik == "Tagihan BPJS Kesehatan":
    df_plot, y_col, is_rp, bar_color = df_ks_filtered.copy(), "Total tagihan", True, "#0077B6"
else:
    df_plot, y_col, is_rp, bar_color = df_ks_filtered.copy(), "Jumlah TK", False, "#FF7A00"

if not df_plot.empty:
    if is_rp:
        df_plot["Label_Val"] = df_plot[y_col].apply(format_rp)
    else:
        df_plot["Label_Val"] = df_plot[y_col].apply(lambda x: f"{x:,.0f}".replace(",", "."))

    fig_bar = px.bar(
        df_plot,
        x="Cabang",
        y=y_col,
        text="Label_Val",
        color_discrete_sequence=[bar_color]
    )
    fig_bar.update_traces(
        textposition="outside",
        marker_line_width=0,
        hovertemplate="<b>%{x}</b><br>" + mode_grafik + ": %{text}<extra></extra>"
    )
    fig_bar.update_layout(
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        margin=dict(t=30, b=20, l=15, r=15),
        height=350,
        xaxis_title="",
        yaxis_title="",
        showlegend=False,
        yaxis=dict(gridcolor="#E2E8F0")
    )
    st.plotly_chart(fig_bar, use_container_width=True)
else:
    st.info(f"Belum ada data {mode_grafik} pada periode {sel_bulan_label} {sel_tahun}.")

st.markdown("---")

# ==============================================================================
# BARIS 3: TABEL REKAPITULASI CABANG & DOWNLOAD EXCEL SLA TAHUNAN
# ==============================================================================
master_cabang = [
    "BABEL", "BANTEN", "DKI", "ICON", "JABAR", "JAWA TIMUR",
    "DIY", "MEDAN", "PADANG", "PALEMBANG", "PEKANBARU", "SIBOLGA"
]
alias_rev = {
    "BANGKA BELITUNG": "BABEL", "DKI JAKARTA": "DKI", "JAWA BARAT": "JABAR",
    "JAWA TENGAH": "DIY", "SIBOLGA NIAS": "SIBOLGA"
}

map_ks_row = {}
for _, r in df_ks_parsed.iterrows():
    c_std = alias_rev.get(r["Cabang"], r["Cabang"])
    map_ks_row[c_std] = r

map_tk_row = {}
for _, r in df_tk_parsed.iterrows():
    c_std = alias_rev.get(r["Cabang"], r["Cabang"])
    map_tk_row[c_std] = r

rekap_rows = []
for idx_c, cab in enumerate(master_cabang, start=1):
    if sel_cabang != "Semua Cabang" and alias_rev.get(sel_cabang, sel_cabang) != cab:
        continue
    r_ks = map_ks_row.get(cab)
    r_tk = map_tk_row.get(cab)

    rekap_rows.append({
        "No": idx_c,
        "Cabang / Unit": cab,
        "SLA BPJS Kes (Limit Tgl 2)": r_ks["Status_SLA"] if r_ks is not None else "⏳ Belum Upload",
        "TK Kes": int(r_ks["Jumlah TK"]) if r_ks is not None else 0,
        "Tagihan BPJS Kes": format_rp(r_ks["Total tagihan"]) if r_ks is not None else "Rp. 0",
        "SLA BPJS TK (Limit Tgl 16)": r_tk["Status_SLA"] if r_tk is not None else "⏳ Belum Upload",
        "TK BPJS TK": int(r_tk["Jumlah TK"]) if r_tk is not None else 0,
        "Tagihan BPJS TK": format_rp(r_tk["Total tagihan"]) if r_tk is not None else "Rp. 0",
    })

# 4 METRIK SLA RINGKAS DI ATAS TABEL MONITORING SLA
st.markdown(
    f"""
    <h3 style="text-align: center; color: #0A2540; font-weight: 800; margin-top: 6px; margin-bottom: 18px;">
        📋 Monitoring & Kepatuhan SLA Cabang ({sel_bulan_label} {sel_tahun})
    </h3>
    """,
    unsafe_allow_html=True
)

def hitung_skor_sla_rata(df_parsed, target_tgl, min_pct):
    if df_parsed.empty:
        return None
    skor_map = {
        0: 100.0, 1: 98.0, 2: 96.0, 3: 94.0, 4: 92.0, 5: 90.0,
        -1: 102.0, -2: 104.0, -3: 106.0, -4: 108.0, -5: 110.0, -6: 115.0
    }
    if target_tgl == 2:  # BPJS Kesehatan
        skor_map.update({3: 93.0, 4: 90.0, 5: 85.0})
    list_skor = []
    for tgl_up in df_parsed["Tanggal_Upload"]:
        selisih = int(tgl_up) - target_tgl
        if selisih > 5:
            list_skor.append(min_pct)
        elif selisih < -6:
            list_skor.append(115.0)
        else:
            list_skor.append(skor_map.get(selisih, 100.0))
    return sum(list_skor) / len(list_skor)

ontime_ks_cnt = len(df_ks_filtered[df_ks_filtered["Tanggal_Upload"] <= 2]) if not df_ks_filtered.empty else 0
late_ks_cnt = len(df_ks_filtered[df_ks_filtered["Tanggal_Upload"] > 2]) if not df_ks_filtered.empty else 0
avg_sla_ks = hitung_skor_sla_rata(df_ks_filtered, target_tgl=2, min_pct=75.0)

ontime_tk_cnt = len(df_tk_filtered[df_tk_filtered["Tanggal_Upload"] <= 16]) if not df_tk_filtered.empty else 0
late_tk_cnt = len(df_tk_filtered[df_tk_filtered["Tanggal_Upload"] > 16]) if not df_tk_filtered.empty else 0
avg_sla_tk = hitung_skor_sla_rata(df_tk_filtered, target_tgl=16, min_pct=88.0)

s1, s2, s3, s4 = st.columns(4)
with s1:
    st.metric(label="🟢 Cabang On-Time (Kes)", value=f"{ontime_ks_cnt} Cabang", delta=f"Late: {late_ks_cnt} Cabang", delta_color="off")
with s2:
    st.metric(label="⭐ Rata-rata SLA Kesehatan", value=f"{avg_sla_ks:.2f}%" if avg_sla_ks is not None else "-", delta="Limit Tgl 2", delta_color="normal")
with s3:
    st.metric(label="🟢 Cabang On-Time (TK)", value=f"{ontime_tk_cnt} Cabang", delta=f"Late: {late_tk_cnt} Cabang", delta_color="off")
with s4:
    st.metric(label="⭐ Rata-rata SLA BPJS TK", value=f"{avg_sla_tk:.2f}%" if avg_sla_tk is not None else "-", delta="Limit Tgl 16", delta_color="normal")

st.markdown("<br>", unsafe_allow_html=True)

df_tabel_utama = pd.DataFrame(rekap_rows)
st.dataframe(df_tabel_utama, use_container_width=True, hide_index=True)

excel_sla_bytes = generate_excel_sla_tahunan(raw_df_ks, raw_df_tk, sel_tahun)
st.download_button(
    label=f"📥 Download Rekap SLA Tahunan ({sel_tahun}) - Format Excel (.xlsx)",
    data=excel_sla_bytes,
    file_name=f"persentase_penilaian_cabang_{sel_tahun}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary"
)
