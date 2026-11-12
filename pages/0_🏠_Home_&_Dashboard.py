import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from google.oauth2.service_account import Credentials
import datetime
import json
import os
import re

# ==============================================================================
# FUNGSI LOAD DATA RIIL DARI GOOGLE SHEETS (REKAP MONITORING KS & TK)
# ==============================================================================
@st.cache_data(ttl=60, show_spinner="⏳ Sedang Memuat Data Monitoring Terbaru dari Google Sheets")
def fetch_monitoring_data():
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        client = gspread.authorize(creds)

        # 1. Baca Rekap_Monitoring_TK (HPI + ICON TK)
        df_tk = pd.DataFrame(columns=["Timestamp", "Cabang", "Jumlah TK", "Total tagihan"])
        try:
            sh_tk = client.open_by_url(st.secrets["SPREADSHEET_URL"])
            ws_tk = sh_tk.worksheet("Rekap_Monitoring_TK")
            records_tk = ws_tk.get_all_records()
            if records_tk:
                df_tk = pd.DataFrame(records_tk)
        except Exception as e:
            st.warning(f"Info Sheet TK: {e}")

        # 2. Baca Rekap_Monitoring_KS (HPI) + Rekap_Monitoring_KS_Icon (ICON)
        list_df_ks = []
        try:
            sh_ks = client.open_by_url(st.secrets["SPREADSHEET_URL_KES"])
            ws_ks = sh_ks.worksheet("Rekap_Monitoring_KS")
            rec_ks = ws_ks.get_all_records()
            if rec_ks:
                list_df_ks.append(pd.DataFrame(rec_ks))
        except Exception as e:
            st.warning(f"Info Sheet KS HPI: {e}")

        try:
            if "SPREADSHEET_URL_KES_ICON" in st.secrets:
                sh_ks_icon = client.open_by_url(st.secrets["SPREADSHEET_URL_KES_ICON"])
                # Cari tab yang namanya mengandung "rekap_monitoring_ks" (abaikan spasi/huruf besar)
                ws_ks_icon = next(
                    (ws for ws in sh_ks_icon.worksheets() if "rekap_monitoring_ks" in ws.title.strip().lower()),
                    None
                )
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

def clean_numeric(val):
    if pd.isna(val) or val == "":
        return 0.0
    s = re.sub(r'[^\d.-]', '', str(val).replace(',', ''))
    try:
        return float(s)
    except ValueError:
        return 0.0

def parse_monitoring_df(df, tahun_target, bulan_target_num):
    if df.empty or "Timestamp" not in df.columns:
        return pd.DataFrame(columns=["Cabang", "Timestamp_DT", "Tanggal_Str", "Hari", "Jumlah TK", "Total tagihan"])

    df = df.copy()
    df["Timestamp_DT"] = pd.to_datetime(df["Timestamp"], errors="coerce")
    df = df.dropna(subset=["Timestamp_DT"])

    # Filter Tahun & Bulan
    df = df[(df["Timestamp_DT"].dt.year == int(tahun_target)) & (df["Timestamp_DT"].dt.month == int(bulan_target_num))]
    if df.empty:
        return pd.DataFrame(columns=["Cabang", "Timestamp_DT", "Tanggal_Str", "Hari", "Jumlah TK", "Total tagihan"])

    df["Cabang"] = df["Cabang"].astype(str).str.strip().str.upper()
    df["Jumlah TK"] = df["Jumlah TK"].apply(clean_numeric).astype(int)
    df["Total tagihan"] = df["Total tagihan"].apply(clean_numeric)

    # Urutkan berdasarkan waktu dan ambil entri terbaru jika cabang melakukan revisi/sync ulang
    df = df.sort_values("Timestamp_DT").groupby("Cabang", as_index=False).last()
    df["Hari"] = df["Timestamp_DT"].dt.day
    df["Tanggal_Str"] = df["Timestamp_DT"].dt.strftime("%d %b %Y (%H:%M)")
    return df

def hitung_sla(hari, batas_tgl):
    if hari <= batas_tgl:
        return 100.0
    selisih = hari - batas_tgl
    skor = max(0.0, 100.0 - (selisih * 2.5))
    return round(skor, 1)

# ==============================================================================
# HERO SECTION & FILTER
# ==============================================================================
col_title, col_refresh = st.columns([5, 1])
with col_title:
    st.title("📊 Dashboard Monitoring SLA & Portal BPJS")
    st.caption("Control Tower Kepatuhan SLA Pengiriman Cabang & Visualisasi Data Riil Real-Time")
with col_refresh:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔄 Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.markdown("---")

BULAN_MAP = {
    "01. Januari": 1, "02. Februari": 2, "03. Maret": 3, "04. April": 4,
    "05. Mei": 5, "06. Juni": 6, "07. Juli": 7, "08. Agustus": 8,
    "09. September": 9, "10. Oktober": 10, "11. November": 11, "12. Desember": 12
}

now_wib = datetime.datetime.utcnow() + datetime.timedelta(hours=7)
default_bulan_idx = now_wib.month - 1

f1, f2, f3 = st.columns(3)
with f1:
    sel_tahun = st.selectbox("📅 Pilih Tahun Monitoring", ["2026", "2027", "2028"], index=0)
with f2:
    sel_bulan_label = st.selectbox("📆 Pilih Bulan Target", list(BULAN_MAP.keys()), index=default_bulan_idx)
    sel_bulan_num = BULAN_MAP[sel_bulan_label]

# Muat Data Riil dari Google Sheets
raw_df_ks, raw_df_tk = fetch_monitoring_data()
df_ks_filtered = parse_monitoring_df(raw_df_ks, sel_tahun, sel_bulan_num)
df_tk_filtered = parse_monitoring_df(raw_df_tk, sel_tahun, sel_bulan_num)

# Daftar master seluruh cabang (gabungan dari JSON + yang tercatat di Sheet)
master_cabang = [
    "BANGKA BELITUNG", "BANTEN", "DKI JAKARTA", "ICON", "JAWA BARAT",
    "JAWA TENGAH", "JAWA TIMUR", "MEDAN", "PADANG", "PALEMBANG",
    "PEKANBARU", "SIBOLGA NIAS"
]
semua_cabang_terdeteksi = sorted(list(set(master_cabang) | set(df_ks_filtered["Cabang"]) | set(df_tk_filtered["Cabang"])))

with f3:
    sel_cabang = st.selectbox("🏢 Filter Cabang", ["Semua Cabang"] + semua_cabang_terdeteksi)

# ==============================================================================
# PENYUSUNAN MATRIKS SLA & METRIK
# ==============================================================================
rows_sla = []
on_time_kes_count = 0
late_kes_count = 0
on_time_tk_count = 0
late_tk_count = 0
list_sla_kes = []
list_sla_tk = []

cabang_loop = semua_cabang_terdeteksi if sel_cabang == "Semua Cabang" else [sel_cabang]

for idx, cab in enumerate(cabang_loop, start=1):
    row_ks = df_ks_filtered[df_ks_filtered["Cabang"] == cab]
    row_tk = df_tk_filtered[df_tk_filtered["Cabang"] == cab]

    # Status Kesehatan (Limit Tgl 2)
    if not row_ks.empty:
        tgl_ks_str = row_ks.iloc[0]["Tanggal_Str"]
        hari_ks = int(row_ks.iloc[0]["Hari"])
        sla_ks_val = hitung_sla(hari_ks, 2)
        sla_ks_str = f"{sla_ks_val:g}%"
        list_sla_kes.append(sla_ks_val)
        if hari_ks <= 2:
            on_time_kes_count += 1
            stat_ks = "OK"
        else:
            late_kes_count += 1
            stat_ks = "LATE"
    else:
        tgl_ks_str = "Menunggu (Limit Tgl 2)"
        sla_ks_str = "-"
        stat_ks = "WAIT"

    # Status Ketenagakerjaan (Limit Tgl 16)
    if not row_tk.empty:
        tgl_tk_str = row_tk.iloc[0]["Tanggal_Str"]
        hari_tk = int(row_tk.iloc[0]["Hari"])
        sla_tk_val = hitung_sla(hari_tk, 16)
        sla_tk_str = f"{sla_tk_val:g}%"
        list_sla_tk.append(sla_tk_val)
        if hari_tk <= 16:
            on_time_tk_count += 1
            stat_tk = "OK"
        else:
            late_tk_count += 1
            stat_tk = "LATE"
    else:
        tgl_tk_str = "Menunggu (Limit Tgl 16)"
        sla_tk_str = "-"
        stat_tk = "WAIT"

    # Penentuan Status Gabungan
    if stat_ks == "OK" and stat_tk == "OK":
        status_total = "🟢 COMPLETE & ON TIME"
    elif "LATE" in (stat_ks, stat_tk):
        status_total = "🟡 LATE SUBMISSION"
    elif stat_ks == "OK" and stat_tk == "WAIT":
        status_total = "🔵 KES OK | MENUNGGU TK"
    elif stat_ks == "WAIT" and stat_tk == "OK":
        status_total = "🔵 TK OK | MENUNGGU KES"
    else:
        status_total = "⚪ BELUM UPLOAD"

    rows_sla.append({
        "No": idx,
        "Cabang": cab,
        "Tgl Upload Kes": tgl_ks_str,
        "SLA Kes": sla_ks_str,
        "Tgl Upload TK": tgl_tk_str,
        "SLA TK": sla_tk_str,
        "Status Total": status_total
    })

df_sla_real = pd.DataFrame(rows_sla)
avg_sla_kes = f"{sum(list_sla_kes)/len(list_sla_kes):.2f}%" if list_sla_kes else "-"
avg_sla_tk = f"{sum(list_sla_tk)/len(list_sla_tk):.2f}%" if list_sla_tk else "-"

# ==============================================================================
# 1. EXECUTIVE SUMMARY (METRIC CARDS REAL-TIME)
# ==============================================================================
m1, m2, m3, m4 = st.columns(4)
with m1:
    st.metric(label="🟢 Cabang On-Time (Kes)", value=f"{on_time_kes_count} Cabang", delta=f"Late: {late_kes_count} Cabang", delta_color="off")
with m2:
    st.metric(label="⭐ Rata-rata SLA Kesehatan", value=avg_sla_kes, delta="Limit Tgl 2")
with m3:
    st.metric(label="🟢 Cabang On-Time (TK)", value=f"{on_time_tk_count} Cabang", delta=f"Late: {late_tk_count} Cabang", delta_color="off")
with m4:
    st.metric(label="⭐ Rata-rata SLA BPJS TK", value=avg_sla_tk, delta="Limit Tgl 16")

st.markdown("---")

# ==============================================================================
# 2. VISUALISASI GRAFIK TREN (DATA RIIL DARI LOGGING SHEET)
# ==============================================================================
st.subheader(f"📈 Tren Rekapitulasi Tagihan & Tenaga Kerja ({sel_bulan_label} {sel_tahun})")

tab_chart1, tab_chart2, tab_chart3, tab_chart4 = st.tabs([
    "💰 Tagihan BPJS Kesehatan",
    "👥 Tenaga Kerja BPJS Kesehatan",
    "💰 Tagihan BPJS Ketenagakerjaan",
    "👥 Tenaga Kerja BPJS Ketenagakerjaan"
])

with tab_chart1:
    if not df_ks_filtered.empty:
        df_ks_plot = df_ks_filtered.copy()
        df_ks_plot["Label_Rp"] = df_ks_plot["Total tagihan"].apply(lambda x: f"Rp. {x:,.0f}".replace(",", "."))
        
        fig_ks_bayar = px.bar(
            df_ks_plot,
            x="Cabang",
            y="Total tagihan",
            text="Label_Rp",
            color="Cabang",
            title=f"Total Tagihan BPJS Kesehatan per Cabang ({sel_bulan_label} {sel_tahun})",
            color_discrete_sequence=px.colors.qualitative.Bold
        )
        fig_ks_bayar.update_traces(
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>Total Tagihan: %{text}<extra></extra>"
        )
        fig_ks_bayar.update_layout(
            yaxis_title="Rupiah (Rp)",
            showlegend=False,
            uniformtext_minsize=8,
            uniformtext_mode="show"
        )
        st.plotly_chart(fig_ks_bayar, use_container_width=True)
    else:
        st.info("Belum ada data log BPJS Kesehatan yang masuk pada periode ini.")

with tab_chart2:
    if not df_ks_filtered.empty:
        fig_ks_tk = px.bar(
            df_ks_filtered, x="Cabang", y="Jumlah TK", text="Jumlah TK", color="Cabang",
            title=f"Jumlah Tenaga Kerja BPJS Kesehatan ({sel_bulan_label} {sel_tahun})",
            color_discrete_sequence=px.colors.qualitative.Prism
        )
        fig_ks_tk.update_layout(showlegend=False)
        st.plotly_chart(fig_ks_tk, use_container_width=True)
    else:
        st.info("Belum ada data log Tenaga Kerja Kesehatan pada periode ini.")

with tab_chart3:
    if not df_tk_filtered.empty:
        df_tk_plot = df_tk_filtered.copy()
        df_tk_plot["Label_Rp"] = df_tk_plot["Total tagihan"].apply(lambda x: f"Rp. {x:,.0f}".replace(",", "."))
        
        fig_tk_bayar = px.bar(
            df_tk_plot,
            x="Cabang",
            y="Total tagihan",
            text="Label_Rp",
            color="Cabang",
            title=f"Total Tagihan BPJS Ketenagakerjaan per Cabang ({sel_bulan_label} {sel_tahun})",
            color_discrete_sequence=px.colors.qualitative.Bold
        )
        fig_tk_bayar.update_traces(
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>Total Tagihan: %{text}<extra></extra>"
        )
        fig_tk_bayar.update_layout(
            yaxis_title="Rupiah (Rp)",
            showlegend=False,
            uniformtext_minsize=8,
            uniformtext_mode="show"
        )
        st.plotly_chart(fig_tk_bayar, use_container_width=True)
    else:
        st.info("Belum ada data log BPJS Ketenagakerjaan yang masuk pada periode ini.")

with tab_chart4:
    if not df_tk_filtered.empty:
        fig_tk_jml = px.bar(
            df_tk_filtered, x="Cabang", y="Jumlah TK", text="Jumlah TK", color="Cabang",
            title=f"Jumlah Tenaga Kerja BPJS Ketenagakerjaan ({sel_bulan_label} {sel_tahun})",
            color_discrete_sequence=px.colors.qualitative.Prism
        )
        fig_tk_jml.update_layout(showlegend=False)
        st.plotly_chart(fig_tk_jml, use_container_width=True)
    else:
        st.info("Belum ada data log Tenaga Kerja Ketenagakerjaan pada periode ini.")

# ==============================================================================
# 3. TABEL CONTROL TOWER MONITORING SLA REAL
# ==============================================================================
st.subheader("📋 Matriks Status Pengiriman & SLA Cabang")
st.caption("Limit Target Pengiriman: BPJS Kesehatan (Maks. Tanggal 2) | BPJS Ketenagakerjaan (Maks. Tanggal 16)")

st.dataframe(
    df_sla_real,
    use_container_width=True,
    hide_index=True,
    column_config={
        "No": st.column_config.TextColumn("No", width="small"),
        "Cabang": st.column_config.TextColumn("Cabang", width="large"),
        "Tgl Upload Kes": st.column_config.TextColumn("Tgl Upload Kes", width="medium"),
        "SLA Kes": st.column_config.TextColumn("SLA Kes", width="small"),
        "Tgl Upload TK": st.column_config.TextColumn("Tgl Upload TK", width="medium"),
        "SLA TK": st.column_config.TextColumn("SLA TK", width="small"),
        "Status Total": st.column_config.TextColumn("Status Total", width="medium"),
    }
)
def generate_excel_sla_tahunan(raw_df_ks, raw_df_tk, tahun_target):
    wb = openpyxl.Workbook()
    # Hapus sheet default
    wb.remove(wb.active)

    # Mapping nama cabang di log ke nama baku di tabel SLA Excel
    alias_cabang = {
        "BANGKA BELITUNG": "BABEL", "BABEL": "BABEL",
        "BANTEN": "BANTEN",
        "DKI JAKARTA": "DKI", "DKI": "DKI",
        "ICON": "ICON",
        "JAWA BARAT": "JABAR", "JABAR": "JABAR",
        "JAWA TIMUR": "JAWA TIMUR",
        "JAWA TENGAH": "DIY", "DIY": "DIY",
        "MEDAN": "MEDAN",
        "PADANG": "PADANG",
        "PALEMBANG": "PALEMBANG",
        "PEKANBARU": "PEKANBARU",
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

    # Ekstrak tanggal upload per (cabang, bulan) untuk tahun terpilih
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
    map_dplk = {}  # Siap terisi otomatis begitu modul DPLK aktif

    # Styling Excel
    font_title = Font(name="Calibri", size=12, bold=True, color="0A2540")
    font_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_subhdr = Font(name="Calibri", size=10, bold=True, color="0A2540")
    font_data = Font(name="Calibri", size=10)
    font_bold = Font(name="Calibri", size=10, bold=True)

    fill_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
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

        # Judul Atas
        ws["A1"] = f"PERSENTASE PENILAIAN KEPATUHAN SLA {sheet_name} - TAHUN {tahun_target}"
        ws["A1"].font = font_title

        # Merge Header Utama (Baris 2-4)
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

        # Sub-Header Bulan (JANUARI - DESEMBER)
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

        # Isi Data 12 Cabang (Baris 5 - 16)
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

                # Rumus VLOOKUP identik dengan persentase_penilaian_cabang.xlsx
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

            # Kolom AB: Rata-rata Total % per Cabang selama setahun
            avg_refs = ",".join([f"{col_let}{r}" for col_let in pct_cols])
            cell_tot = ws.cell(r, 28, f'=IFERROR(AVERAGE({avg_refs}),"")')
            cell_tot.number_format = "0.0%"
            cell_tot.font = font_bold
            cell_tot.alignment = align_center
            cell_tot.fill = fill_total
            cell_tot.border = border_all

        # Baris 17: Rata-Rata Nasional per Bulan
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

        # Tabel Referensi Penilaian (Baris 19 - 32)
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

        # Atur Lebar Kolom
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

st.markdown("<br>", unsafe_allow_html=True)
excel_sla_bytes = generate_excel_sla_tahunan(raw_df_ks, raw_df_tk, sel_tahun)
st.download_button(
    label=f"📥 Download Rekap SLA Tahunan Lengkap ({sel_tahun}) - Format Excel (.xlsx)",
    data=excel_sla_bytes,
    file_name=f"persentase_penilaian_cabang_{sel_tahun}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary"
)