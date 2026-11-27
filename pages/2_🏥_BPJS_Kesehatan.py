import streamlit as st
import pandas as pd
import pdfplumber
import json
import os
import re
import gc
import io
import docx
import datetime
import openpyxl
import gspread
from google.oauth2.service_account import Credentials
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docxtpl import DocxTemplate
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Set Page Config (Harus di paling atas)
st.set_page_config(
    page_title="BPJS Kesehatan",
    page_icon="🏥",
    layout="wide"
)

# Custom Styling Tema Utama
from ui_theme import apply_glass_theme

apply_glass_theme()

# Format Tanggal Bahasa Indonesia
NAMA_BULAN = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]
now = datetime.datetime.now()
tanggal_hari_ini = f"{now.day} {NAMA_BULAN[now.month]} {now.year}"

# ==============================================================================
# FUNGSI HELPER & KEAMANAN
# ==============================================================================
def safe_int(val):
    try:
        if pd.isna(val) or val is None or str(val).strip() == "" or str(val).strip().lower() == "none":
            return 0
        return int(float(val))
    except (ValueError, TypeError):
        return 0

def custom_round(val):
    try:
        val_float = float(val)
        integer_part = int(val_float)
        decimal_part = val_float - integer_part
        if decimal_part > 0.50:
            return integer_part + 1
        else:
            return integer_part
    except (ValueError, TypeError):
        return 0

def extract_periode(tanggal_str):
    if not tanggal_str:
        return "-"
    parts = tanggal_str.strip().split()
    if len(parts) >= 3:
        return f"{parts[-2]} {parts[-1]}"
    elif len(parts) == 2:
        return tanggal_str
    return tanggal_str

def terbilang(n):
    satuan = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"]
    n = int(abs(n))
    if n < 12:
        return satuan[n]
    elif n < 20:
        return terbilang(n - 10) + " Belas"
    elif n < 100:
        return terbilang(n // 10) + " Puluh " + terbilang(n % 10)
    elif n < 200:
        return "Seratus " + terbilang(n - 100)
    elif n < 1000:
        return terbilang(n // 100) + " Ratus " + terbilang(n % 100)
    elif n < 2000:
        return "Seribu " + terbilang(n - 1000)
    elif n < 1000000:
        return terbilang(n // 1000) + " Ribu " + terbilang(n % 1000)
    elif n < 1000000000:
        return terbilang(n // 1000000) + " Juta " + terbilang(n % 1000000)
    elif n < 1000000000000:
        return terbilang(n // 1000000000) + " Miliar " + terbilang(n % 1000000000)
    return str(n)

def set_cell_border(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(r'''
        <w:tcBorders %s>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        </w:tcBorders>
    ''' % nsdecls('w'))
    tcPr.append(tcBorders)

def alokasi_prorata_bulat(total_tagihan_billing, total_tk_billing, list_tk_spk):
    """
    Menghitung pembagian rupiah per SPK secara proporsional (prorata) dalam bilangan bulat (integer)
    menggunakan metode Largest Remainder sehingga jumlah total pecahan 100% identik dengan total billing.
    """
    total_tagihan_int = custom_round(total_tagihan_billing)
    total_tk_int = safe_int(total_tk_billing)

    if total_tk_int <= 0 or total_tagihan_int <= 0 or not list_tk_spk:
        return [0 for _ in list_tk_spk]

    tarif_per_orang = total_tagihan_int / total_tk_int
    exact_shares = [safe_int(tk) * tarif_per_orang for tk in list_tk_spk]
    floored_shares = [int(val) for val in exact_shares]

    # Hitung sisa pembulatan hanya jika total TK yang diinput sudah sama dengan total TK billing
    # (Jika input masih sebagian/kurang, hitung proporsional murni tanpa memaksa sisa ke baris terakhir)
    sum_input_tk = sum(safe_int(tk) for tk in list_tk_spk)
    if sum_input_tk == total_tk_int:
        sisa_rupiah = total_tagihan_int - sum(floored_shares)
        remainders = [(exact_shares[i] - floored_shares[i], i) for i in range(len(list_tk_spk)) if safe_int(list_tk_spk[i]) > 0]
        remainders.sort(key=lambda x: (-x[0], x[1]))
        for k in range(min(sisa_rupiah, len(remainders))):
            idx_target = remainders[k][1]
            floored_shares[idx_target] += 1
        return floored_shares
    else:
        return [custom_round(val) for val in exact_shares]

# ==============================================================================
# AUTO-SYNC GOOGLE SHEETS (MENDUKUNG MODE SESUAI BILLING & SESUAI SPK)
# ==============================================================================
def sync_to_google_sheet_kesehatan(df_target, nama_tab_sheet, cabang_nama, is_icon=False, mode_sync="SESUAI_BILLING"):
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        client = gspread.authorize(creds)

        if is_icon:
            sheet_url = st.secrets["SPREADSHEET_URL_KES_ICON"]
            col_key_search = 8   # Kolom H: VIRTUAL ACCOUNT BANK MANDIRI
            col_spk = 4          # Kolom D: NO KONTRAK
            col_tk = 11          # Kolom K: Jumlah TK
            col_iuran = 13       # Kolom M: Total Tagihan
        else:
            sheet_url = st.secrets["SPREADSHEET_URL_KES"]
            col_key_search = 8   # Kolom H: Nomor VA di sheet HPI
            col_spk = 4          # Kolom D: Nomor Kontrak
            # Tab SESUAI SPK: Kolom I = JUMLAH TK DI INVOICE, Kolom J = JUMLAH TK DIKONTRAK (hasil breakdown SPK), Kolom K = IURAN
            col_tk = 10 if mode_sync == "SESUAI_SPK" else 9
            col_iuran = 11       # Kolom K: Total Iuran

        sh = client.open_by_url(sheet_url)
        worksheet = sh.worksheet(nama_tab_sheet)

        # 1x API Read untuk seluruh kolom VA dan SPK agar super cepat dan ringan
        raw_keys = worksheet.col_values(col_key_search)
        key_list = [re.sub(r'\D', '', str(x)) for x in raw_keys]
        spk_list = [str(x).strip() for x in worksheet.col_values(col_spk)]

        let_spk = get_column_letter(col_spk)
        let_tk = get_column_letter(col_tk)
        let_iuran = get_column_letter(col_iuran)

        batch_data = []
        updated_count = 0
        unmatched_rows = []
        baris_terpakai = set()
        records = df_target.to_dict(orient="records")

        for row in records:
            kode_pdf = re.sub(r'\D', '', str(row.get("Nomor VA", "")))
            spk_pdf = str(row.get("Nomor Kontrak", "")).strip()
            tk_val = safe_int(row.get("Jumlah Tenaga Kerja", 0))
            iuran_val = custom_round(row.get("Total Iuran", 0))

            if tk_val <= 0 and iuran_val <= 0:
                continue

            kandidat_baris = [
                idx + 1 for idx, val in enumerate(key_list)
                if kode_pdf and val == kode_pdf and (idx + 1) not in baris_terpakai
            ]

            row_idx = None
            if kandidat_baris:
                # Prioritas 1: Cocokkan VA sama DAN SPK di sheet sama persis (Wajib untuk Mode SESUAI SPK)
                for r in kandidat_baris:
                    spk_sheet = spk_list[r - 1] if (r - 1) < len(spk_list) else ""
                    if spk_sheet.lower() == spk_pdf.lower():
                        row_idx = r
                        break

                # Prioritas 2: Jika SPK di sheet target masih KOSONG (belum dipilih dropdown-nya) di bawah VA yang sama, isi ke baris kosong tersebut!
                if row_idx is None:
                    for r in kandidat_baris:
                        spk_sheet = spk_list[r - 1] if (r - 1) < len(spk_list) else ""
                        if spk_sheet in ["", "-"]:
                            row_idx = r
                            break

                # Prioritas 3: Fallback khusus mode SESUAI BILLING atau pencocokan SPK langsung jika VA belum terisi
                if row_idx is None:
                    if mode_sync == "SESUAI_BILLING":
                        row_idx = kandidat_baris[0]
                    else:
                        for idx_s, s_val in enumerate(spk_list):
                            if (idx_s + 1) not in baris_terpakai and s_val.lower() == spk_pdf.lower() and spk_pdf not in ["", "-"]:
                                row_idx = idx_s + 1
                                break

            if row_idx is not None:
                baris_terpakai.add(row_idx)
                # Pada mode SESUAI SPK, SPK di sheet sudah menjadi acuan master; tetap pastikan sinkron
                if spk_pdf and spk_pdf != "-":
                    batch_data.append({
                        "range": f"{let_spk}{row_idx}",
                        "values": [[spk_pdf]]
                    })
                batch_data.append({
                    "range": f"{let_tk}{row_idx}",
                    "values": [[tk_val]]
                })
                batch_data.append({
                    "range": f"{let_iuran}{row_idx}",
                    "values": [[iuran_val]]
                })
                updated_count += 1
            else:
                unmatched_rows.append(f"VA {kode_pdf} | SPK {spk_pdf}")

        if batch_data:
            # 1x Batch Update API Call untuk seluruh baris
            worksheet.batch_update(batch_data, value_input_option="USER_ENTERED")

        # Auto-Logging hanya dicatat 1x pada tahap SESUAI BILLING agar tidak double-counting di Rekap_Monitoring_KS
        if updated_count > 0 and mode_sync == "SESUAI_BILLING":
            try:
                wib_now = datetime.datetime.utcnow() + datetime.timedelta(hours=7)
                timestamp_str = wib_now.strftime("%Y-%m-%d %H:%M:%S")
                total_tk_log = int(df_target["Jumlah Tenaga Kerja"].sum())
                total_tagihan_log = custom_round(df_target["Total Iuran"].sum())

                if is_icon:
                    sh_kes_icon = client.open_by_url(st.secrets["SPREADSHEET_URL_KES_ICON"])
                    ws_mon_ks = next(
                        (ws for ws in sh_kes_icon.worksheets() if "rekap_monitoring_ks" in ws.title.strip().lower()),
                        None
                    )
                    if ws_mon_ks:
                        ws_mon_ks.append_row([timestamp_str, "ICON", total_tk_log, total_tagihan_log], value_input_option="USER_ENTERED")
                else:
                    sh_kes_utama = client.open_by_url(st.secrets["SPREADSHEET_URL_KES"])
                    ws_mon_ks = sh_kes_utama.worksheet("Rekap_Monitoring_KS")
                    ws_mon_ks.append_row([timestamp_str, cabang_nama.upper(), total_tk_log, total_tagihan_log], value_input_option="USER_ENTERED")
            except Exception as log_err:
                st.warning(f"Sync VA sukses, namun log monitoring KS gagal dicatat: {log_err}")

        return True, {"updated": updated_count, "unmatched": unmatched_rows}
    except Exception as e:
        return False, str(e)

# ==============================================================================
# DATABASE KONTRAK & DATABASE MAPPING VA MANDIRI (JSON RINGAN)
# ==============================================================================
@st.cache_data(ttl=300, max_entries=3)
def load_database():
    possible_paths = [
        "database_kontrak.json",
        "database.json",
        os.path.join("..", "database_kontrak.json")
    ]
    for path in possible_paths:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
    return {}

@st.cache_data(ttl=300, max_entries=3)
def load_va_npp_database():
    possible_paths = [
        "database_va_npp.json",
        os.path.join("..", "database_va_npp.json")
    ]
    for path in possible_paths:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
    return {"kesehatan": {}, "ketenagakerjaan": {}}

db_kontrak = load_database()
db_va_npp = load_va_npp_database()
db_va_kes = db_va_npp.get("kesehatan", {})

st.sidebar.header("📌 Pengaturan Dokumen Word")
wilayah_list = list(db_kontrak.keys()) if db_kontrak else ["Cab. Palembang", "Cab. Medan", "Cab. Sibolga Nias", "Cab. Bangka Belitung"]
selected_wilayah = st.sidebar.selectbox("Pilih Cabang / Wilayah", wilayah_list)

wilayah_display = re.sub(r'(?i)\bcab\.?\s*', '', selected_wilayah).strip()
data_cabang = db_kontrak.get(selected_wilayah, {})

nama_kota = data_cabang.get("kota", wilayah_display.title())
default_manager = data_cabang.get("manager", "NAMA MANAGER")
kontrak_options = data_cabang.get("spk", ["-"])
spk_io_map_cabang = data_cabang.get("spk_io", {})

if not kontrak_options:
    kontrak_options = ["-"]

st.sidebar.markdown("---")
st.sidebar.header("📝 Informasi Header Surat")
nomor_surat = st.sidebar.text_input("Nomor Surat", "001/SP/HPI/X/2026")
tanggal_surat = st.sidebar.text_input("Tanggal Surat", value=tanggal_hari_ini)
nama_manager = st.sidebar.text_input("Nama Manager Cabang", value=default_manager)

st.sidebar.markdown("---")
st.sidebar.caption(f"🗄️ **Database VA Mandiri:** `{len(db_va_kes)}` VA aktif terhubung")

# ==============================================================================
# KONTEN UTAMA: BPJS KESEHATAN
# ==============================================================================
st.title("🏥 Rekap PDF Billing Statement & Generator Surat BPJS Kesehatan")
st.caption("Otomatisasi pembacaan VA Mandiri -> SPK, kertas kerja pemecahan VA Gabung per SPK, export Word/Excel, dan Auto-Sync Google Sheet.")

uploaded_files_kes = st.file_uploader(
    "Upload File PDF Billing Statement BPJS Kesehatan (Bisa banyak file sekaligus)",
    type=["pdf"],
    accept_multiple_files=True,
    key="uploader_kes"
)

if uploaded_files_kes:
    # Bangun fingerprint file agar cache pembacaan PDF tidak mengulang parsing saat user edit tabel
    file_signature = tuple((f.name, f.size) for f in uploaded_files_kes)

    if st.session_state.get("kes_last_files_sig") != file_signature:
        all_data_kes = []
        for file in uploaded_files_kes:
            try:
                file.seek(0)
                with pdfplumber.open(file) as pdf:
                    raw_text = ""
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            raw_text += text + "\n"

                clean_text = raw_text.replace('\n', ' ').replace('|', '')
                default_spk_cabang = kontrak_options[0] if (kontrak_options and len(kontrak_options) > 0) else "-"

                data_kes = {
                    "File PDF": file.name,
                    "Nama Perusahaan": "-",
                    "Nomor VA": "-",
                    "Status VA": "Belum Terpetakan",
                    "Nomor Kontrak": str(default_spk_cabang),
                    "Kandidat SPK": [],
                    "Jumlah Tenaga Kerja": 0,
                    "Total Iuran": 0
                }

                # 1. NAMA PERUSAHAAN
                match_nama = re.search(r'Yth\.\s*(.*?)(?:\n\s*JL\.|\s*JL\.|\s*Nomor Virtual|\s*Pelaporan|\n\n|$)', raw_text, re.IGNORECASE | re.DOTALL)
                if match_nama:
                    nama_bu = match_nama.group(1).replace('\n', ' ').strip()
                    nama_bu = re.sub(r'(?:Surat Tagihan|Billing Statement|Pelaporan Penagihan|JKN-KIS).*$', '', nama_bu, flags=re.IGNORECASE).strip()
                    data_kes["Nama Perusahaan"] = re.sub(r'\s+', ' ', nama_bu)

                # 2. NOMOR VA (DIUTAMAKAN MANDIRI / MDR)
                va_found = ""
                match_va_mdr = re.search(r'(?:MDR|Mandiri)\s*:?\s*([0-9]{13,16})', clean_text, re.IGNORECASE)
                if match_va_mdr:
                    va_found = match_va_mdr.group(1).strip()
                else:
                    match_va_gen = re.search(r'(?:BNI|BRI|BTN|Virtual Account)\s*:?\s*([0-9]{13,16})', clean_text, re.IGNORECASE)
                    if match_va_gen:
                        va_found = match_va_gen.group(1).strip()

                data_kes["Nomor VA"] = va_found if va_found else "-"

                # 3. PENCOCOKAN OTOMATIS VA MANDIRI -> DATABASE MAPPING SPK (JSON RINGAN)
                va_digits = re.sub(r'\D', '', va_found)
                if va_digits and va_digits in db_va_kes:
                    spk_candidates = db_va_kes[va_digits].get("spk_list", [])
                    data_kes["Kandidat SPK"] = spk_candidates
                    if len(spk_candidates) == 1:
                        data_kes["Nomor Kontrak"] = spk_candidates[0]
                        data_kes["Status VA"] = "✅ Auto (1 SPK)"
                    elif len(spk_candidates) > 1:
                        # Sesuai instruksi: ambil salah satu SPK default sebagai penanda di tahap Billing
                        data_kes["Nomor Kontrak"] = spk_candidates[0]
                        data_kes["Status VA"] = f"⚠️ VA Gabung ({len(spk_candidates)} SPK)"
                    else:
                        data_kes["Status VA"] = "⏳ Kosong di DB (Pakai Default)"
                else:
                    data_kes["Status VA"] = "❓ VA Baru (Pakai Default)"

                # 4. JUMLAH TENAGA KERJA
                match_pegawai = re.search(r'(?:Rp\.?|Rp)?\s*[\d.]+\s*(?:Rp\.?|Rp)?\s*(\d+)\s+\d+\s+\d{2}/\d{2}/\d{4}', clean_text, re.IGNORECASE)
                if match_pegawai:
                    data_kes["Jumlah Tenaga Kerja"] = int(match_pegawai.group(1).strip())
                else:
                    match_peg_alt = re.search(r'(?:Jumlah Pegawai|Number of Employees).*?\s+(\d+)', clean_text, re.IGNORECASE)
                    if match_peg_alt:
                        data_kes["Jumlah Tenaga Kerja"] = int(match_peg_alt.group(1).strip())

                # 5. TOTAL IURAN (DIBULATKAN PRESISI)
                match_total = re.search(r'(?:TOTAL TAGIHAN|Total Balance).?(?:Rp\.?|Rp)\s([\d.]+)', clean_text, re.IGNORECASE)
                if not match_total:
                    match_total = re.search(r'(?:Premi/Iuran|Contribution).?(?:Rp\.?|Rp)\s([\d.]+)', clean_text, re.IGNORECASE)

                if match_total:
                    raw_angka = match_total.group(1).strip()
                    clean_angka = raw_angka.replace('.', '')
                    data_kes["Total Iuran"] = custom_round(float(clean_angka))

                all_data_kes.append(data_kes)

            except Exception as e:
                st.error(f"Terjadi kesalahan saat membaca file {file.name}: {e}")

        gc.collect()
        st.session_state["kes_last_files_sig"] = file_signature
        st.session_state["kes_raw_records"] = all_data_kes

    all_data_kes = st.session_state.get("kes_raw_records", [])

    if all_data_kes:
        df_kes_raw = pd.DataFrame(all_data_kes)

        # Ringkasan Deteksi VA Mandiri
        jml_auto = sum(1 for r in all_data_kes if "Auto" in r["Status VA"])
        jml_gabung = sum(1 for r in all_data_kes if "VA Gabung" in r["Status VA"])
        jml_kosong = len(all_data_kes) - jml_auto - jml_gabung

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total File PDF Billing", f"{len(all_data_kes)} File")
        m2.metric("Auto-SPK (1-to-1)", f"{jml_auto} Billing")
        m3.metric("VA Gabung (Multi-SPK)", f"{jml_gabung} Billing")
        m4.metric("SPK Kosong / VA Baru", f"{jml_kosong} Billing")

        # ======================================================================
        # PILIHAN 2 ALUR PENGERJAAN: SESUAI BILLING vs SESUAI SPK
        # ======================================================================
        st.markdown("---")
        tab_billing, tab_spk = st.tabs([
            "1️⃣ Tahap 1: Rekap SESUAI BILLING (Apa Adanya dari PDF)",
            "2️⃣ Tahap 2: Kertas Kerja & Rekap SESUAI SPK (Pecah VA Gabung Proporsional)"
        ])

        # ======================================================================
        # TAB 1: REKAP SESUAI BILLING
        # ======================================================================
        with tab_billing:
            st.subheader("✏️ Edit Data & Nomor Kontrak (SPK) — Mode Sesuai Billing")
            st.caption(
                "💡 VA Mandiri otomatis dicocokkan dengan `database_va_npp.json`. "
                "Untuk VA Gabung (>1 SPK), sistem otomatis mengisi salah satu SPK default sebagai penanda awal."
            )

            with st.expander("📌 Lihat Daftar Nomor Kontrak (SPK) Cabang Ini"):
                for opt in kontrak_options:
                    io_tag = f" [IO: {spk_io_map_cabang[opt]}]" if opt in spk_io_map_cabang else ""
                    st.code(f"{opt}{io_tag}", language="text")

            df_display_billing = df_kes_raw[["Nama Perusahaan", "Nomor VA", "Status VA", "Nomor Kontrak", "Jumlah Tenaga Kerja", "Total Iuran"]].copy()

            column_config_kes = {
                "Nama Perusahaan": st.column_config.TextColumn("Nama Perusahaan / BU"),
                "Nomor VA": st.column_config.TextColumn("Nomor VA Mandiri"),
                "Status VA": st.column_config.TextColumn("Deteksi VA", disabled=True),
                "Nomor Kontrak": st.column_config.TextColumn("Nomor Kontrak (SPK)", width="large"),
                "Jumlah Tenaga Kerja": st.column_config.NumberColumn("Jml TK", format="%d"),
                "Total Iuran": st.column_config.NumberColumn("Total Tagihan (Rp)", format="Rp %d")
            }

            edited_df_kes = st.data_editor(
                df_display_billing,
                column_config=column_config_kes,
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_combine_kes_{selected_wilayah}_{len(uploaded_files_kes)}"
            )

            edited_df_kes.reset_index(drop=True, inplace=True)
            edited_df_kes["NO"] = [idx + 1 for idx in range(len(edited_df_kes))]
            edited_df_kes["Total Iuran"] = edited_df_kes["Total Iuran"].apply(custom_round)

            total_iuran_kes = int(edited_df_kes["Total Iuran"].sum())
            total_tk_kes = int(edited_df_kes["Jumlah Tenaga Kerja"].sum())
            terbilang_kes = (terbilang(total_iuran_kes).strip() + " Rupiah") if total_iuran_kes > 0 else "-"

            st.markdown(f"### **Total Tagihan BPJS Kesehatan (Sesuai Billing):** Rp {total_iuran_kes:,.0f}".replace(",", "."))
            st.markdown(f"**Terbilang:** *{terbilang_kes}*")

            st.subheader("📥 Unduh Hasil Dokumen BPJS Kesehatan (Sesuai Billing)")
            col1, col2 = st.columns(2)

            # 1. Export Excel BPJS Kesehatan
            with col1:
                st.markdown("##### 📊 Export Master Excel")
                clean_filename_wilayah = wilayah_display.replace(".", "").replace(" ", "_")
                excel_filename_kes = f"Rekap_BPJS_Kesehatan_{clean_filename_wilayah}.xlsx"

                output_excel_kes = io.BytesIO()
                wb_kes = openpyxl.Workbook()
                ws_kes = wb_kes.active
                ws_kes.title = "Rekap BPJS Kesehatan"

                font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                font_data = Font(name="Calibri", size=11)
                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'),
                    right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'),
                    bottom=Side(style='thin', color='D9D9D9')
                )

                headers_kes = ["NO", "Nama Perusahaan", "Nomor VA", "Nomor Kontrak", "Jml TK", "Total Tagihan"]
                ws_kes.append(headers_kes)
                for col_num in range(1, len(headers_kes) + 1):
                    cell = ws_kes.cell(row=1, column=col_num)
                    cell.font = font_header
                    cell.fill = fill_header
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                records_kes = edited_df_kes.to_dict(orient="records")
                for r_idx, row in enumerate(records_kes, start=2):
                    ws_kes.cell(row=r_idx, column=1, value=int(row["NO"])).alignment = Alignment(horizontal="center")
                    ws_kes.cell(row=r_idx, column=2, value=str(row["Nama Perusahaan"])).alignment = Alignment(horizontal="left")
                    ws_kes.cell(row=r_idx, column=3, value=str(row["Nomor VA"])).alignment = Alignment(horizontal="center")
                    ws_kes.cell(row=r_idx, column=4, value=str(row["Nomor Kontrak"])).alignment = Alignment(horizontal="left")
                    ws_kes.cell(row=r_idx, column=5, value=safe_int(row["Jumlah Tenaga Kerja"])).alignment = Alignment(horizontal="center")

                    c_cell = ws_kes.cell(row=r_idx, column=6, value=custom_round(row["Total Iuran"]))
                    c_cell.number_format = '"Rp "#,##0'
                    c_cell.alignment = Alignment(horizontal="right")

                    for c_idx in range(1, 7):
                        c = ws_kes.cell(row=r_idx, column=c_idx)
                        c.font = font_data
                        c.border = thin_border

                for col in ws_kes.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws_kes.column_dimensions[col_letter].width = max(max_len + 3, 12)

                wb_kes.save(output_excel_kes)
                output_excel_kes.seek(0)

                st.download_button(
                    label="📥 Download Master Excel BPJS Kesehatan",
                    data=output_excel_kes.getvalue(),
                    file_name=excel_filename_kes,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # 2. Export Word Docx BPJS Kesehatan
            with col2:
                st.markdown("##### 📝 Export Surat Word (.docx)")
                template_kes_path = os.path.join("templates", "Template_DEFAULT_kesehatan.docx")
                if not os.path.exists(template_kes_path):
                    template_kes_path = os.path.join("templates", "Template_DEFAULT.docx")

                if os.path.exists(template_kes_path):
                    doc_tpl = DocxTemplate(template_kes_path)
                    periode_kes = extract_periode(tanggal_surat)

                    context = {
                        "Kota": nama_kota,
                        "Wilayah": wilayah_display,
                        "Tanggal": tanggal_surat,
                        "Periode": periode_kes,
                        "NomorSurat": nomor_surat,
                        "Terbilang": terbilang_kes,
                        "NamaManager": nama_manager,
                        "TOTAL_TK": str(total_tk_kes),
                        "TOTAL_IURAN": f"Rp {total_iuran_kes:,.0f}".replace(",", ".")
                    }
                    doc_tpl.render(context)

                    temp_stream = io.BytesIO()
                    doc_tpl.save(temp_stream)
                    temp_stream.seek(0)

                    doc_native = docx.Document(temp_stream)

                    target_p = None
                    for p in doc_native.paragraphs:
                        if "sebesar" in p.text.lower() or "anggaran" in p.text.lower():
                            target_p = p
                            break

                    for tbl in list(doc_native.tables):
                        tbl._element.getparent().remove(tbl._element)

                    if target_p:
                        spacer_top_kes = doc_native.add_paragraph()
                        spacer_top_kes.paragraph_format.space_before = Pt(4)
                        spacer_top_kes.paragraph_format.space_after = Pt(4)
                        target_p._p.addnext(spacer_top_kes._p)
                        insert_point_kes = spacer_top_kes
                    else:
                        insert_point_kes = target_p

                    new_table = doc_native.add_table(rows=0, cols=6)
                    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER

                    if insert_point_kes:
                        insert_point_kes._p.addnext(new_table._tbl)

                    hdr_cells = new_table.add_row().cells
                    headers_kes_docx = ["NO", "NAMA BADAN USAHA", "NOMOR SPK", "NOMOR VA", "JML TK", "TOTAL TAGIHAN"]
                    widths_kes = [Inches(0.4), Inches(2.2), Inches(2.2), Inches(1.3), Inches(0.6), Inches(1.3)]

                    for idx, text in enumerate(headers_kes_docx):
                        hdr_cells[idx].text = text
                        hdr_cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        hdr_cells[idx].paragraphs[0].runs[0].font.bold = True
                        set_cell_border(hdr_cells[idx])
                        hdr_cells[idx].width = widths_kes[idx]

                    for r_data in records_kes:
                        val_tagihan = custom_round(r_data["Total Iuran"])
                        formatted_tagihan = f"Rp\t{val_tagihan:,.0f}".replace(",", ".")

                        row_cells = new_table.add_row().cells
                        row_cells[0].text = str(r_data["NO"])
                        row_cells[1].text = str(r_data["Nama Perusahaan"])
                        row_cells[2].text = str(r_data["Nomor Kontrak"])
                        row_cells[3].text = str(r_data["Nomor VA"])
                        row_cells[4].text = str(safe_int(r_data["Jumlah Tenaga Kerja"]))

                        p_tagihan = row_cells[5].paragraphs[0]
                        p_tagihan.text = formatted_tagihan
                        p_tagihan.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        p_tagihan.paragraph_format.tab_stops.add_tab_stop(Inches(1.25), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)

                        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                        row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                        row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                        for c_idx in range(6):
                            set_cell_border(row_cells[c_idx])
                            row_cells[c_idx].width = widths_kes[c_idx]

                    total_cells = new_table.add_row().cells
                    total_cells[0].text = "Jumlah"
                    total_cells[4].text = str(total_tk_kes)

                    p_tot_kes = total_cells[5].paragraphs[0]
                    p_tot_kes.text = f"Rp\t{total_iuran_kes:,.0f}".replace(",", ".")
                    p_tot_kes.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p_tot_kes.paragraph_format.tab_stops.add_tab_stop(Inches(1.25), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)
                    if p_tot_kes.runs:
                        p_tot_kes.runs[0].font.bold = True

                    total_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    total_cells[0].paragraphs[0].runs[0].font.bold = True
                    total_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    total_cells[4].paragraphs[0].runs[0].font.bold = True

                    for c_idx in range(6):
                        set_cell_border(total_cells[c_idx])
                        total_cells[c_idx].width = widths_kes[c_idx]

                    spacer_bottom_kes = doc_native.add_paragraph()
                    spacer_bottom_kes.paragraph_format.space_before = Pt(6)
                    spacer_bottom_kes.paragraph_format.space_after = Pt(6)
                    new_table._tbl.addnext(spacer_bottom_kes._p)

                    doc_io_kes = io.BytesIO()
                    doc_native.save(doc_io_kes)
                    doc_io_kes.seek(0)

                    word_filename_kes = f"Surat_BPJS_Kesehatan_{clean_filename_wilayah}.docx"
                    st.download_button(
                        label="📥 Download Surat Word (.docx)",
                        data=doc_io_kes.getvalue(),
                        file_name=word_filename_kes,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.error("Template Word belum ditemukan di `templates/`!")

            # AUTO-SYNC GOOGLE SHEETS: TAB SESUAI BILLING
            st.markdown("---")
            st.subheader("🚀 Sync Tahap 1 ke Google Sheets (Tab: SESUAI BILLING)")

            target_sheet_choice_ks = st.radio(
                "Pilih Target Google Sheet:",
                ["Sheet Rekap Utama HPI", "Sheet Khusus Cabang ICON"],
                horizontal=True,
                key="target_sheet_choice_ks_billing"
            )

            @st.cache_data(ttl=300, max_entries=4)
            def get_sheet_tabs_ks(url):
                scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
                creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
                client = gspread.authorize(creds)
                sh = client.open_by_url(url)
                return [ws.title for ws in sh.worksheets()]

            try:
                sheet_url_ks = st.secrets["SPREADSHEET_URL_KES_ICON"] if target_sheet_choice_ks == "Sheet Khusus Cabang ICON" else st.secrets["SPREADSHEET_URL_KES"]
                available_tabs_ks = get_sheet_tabs_ks(sheet_url_ks)

                # Otomatis pilihkan tab yang mengandung kata "SESUAI BILLING" jika ada
                default_idx_bill = next(
                    (i for i, t in enumerate(available_tabs_ks) if "sesuai billing" in t.lower()),
                    0
                )
                tab_sheet_target_ks = st.selectbox(
                    "Pilih Tab Google Sheet Target (Sesuai Billing):",
                    available_tabs_ks,
                    index=default_idx_bill,
                    key="target_sheet_ks_billing"
                )
            except Exception:
                st.warning("⚠️ Server Google Sheet API sedang sibuk (Error 503). Silakan ketik nama tab secara manual di bawah ini:")
                tab_sheet_target_ks = st.text_input("Nama Tab Google Sheet Target (Manual)", "1. NOV 26 SPSH SESUAI BILLING", key="target_sheet_ks_billing_man")

            if st.button("⚡ Sync Tab SESUAI BILLING ke Google Sheet Sekarang", type="primary", key="btn_sync_ks_billing"):
                with st.spinner("Sedang mencocokkan Nomor VA dan mengirim data ke Tab Sesuai Billing..."):
                    is_icon_selected_ks = (target_sheet_choice_ks == "Sheet Khusus Cabang ICON")
                    success, res = sync_to_google_sheet_kesehatan(
                        edited_df_kes,
                        tab_sheet_target_ks,
                        wilayah_display,
                        is_icon=is_icon_selected_ks,
                        mode_sync="SESUAI_BILLING"
                    )
                    if success:
                        st.success(f"BERHASIL! {res['updated']} baris berhasil di-Update ke **{tab_sheet_target_ks}** & dicatat ke Rekap_Monitoring_KS!")
                        if res["unmatched"]:
                            st.warning(f"⚠️ Baris tidak ditemukan di sheet: {', '.join(res['unmatched'])}")
                    else:
                        st.error(f"Gagal melakukan sync: {res}")

        # ======================================================================
        # TAB 2: KERTAS KERJA & REKAP SESUAI SPK (PECAH VA GABUNG PROPORSIONAL)
        # ======================================================================
        with tab_spk:
            st.subheader("🧮 Kertas Kerja Pemecahan Billing Gabung -> Rekap SESUAI SPK")
            st.info(
                "💡 **Cara Kerja Kertas Kerja Otomatis:**\n"
                "1. Untuk **VA Tunggal (1 SPK)**, data otomatis masuk 100% sesuai PDF Billing tanpa perlu dipecah.\n"
                "2. Untuk **VA Gabung (>1 SPK)**, sistem otomatis menampilkan daftar SPK di bawah VA tersebut beserta **Internal Order (IO)**-nya. "
                "PIC cukup mengetik **Jumlah TK per SPK**, sistem otomatis menghitung **Nominal Tagihan Bulat (Rp)** secara proporsional "
                "dan menjaga agar total pecahan **100% identik** dengan total PDF Billing."
            )

            # Bangun struktur kertas kerja dari hasil edited_df_kes (Tahap 1)
            kertas_kerja_rows = []
            for idx_b, row_b in edited_df_kes.iterrows():
                va_str = str(row_b["Nomor VA"]).strip()
                va_clean = re.sub(r'\D', '', va_str)
                bu_str = str(row_b["Nama Perusahaan"]).strip()
                spk_billing_str = str(row_b["Nomor Kontrak"]).strip()
                tk_billing_tot = safe_int(row_b["Jumlah Tenaga Kerja"])
                rp_billing_tot = custom_round(row_b["Total Iuran"])

                # Ambil semua SPK yang terdaftar di VA Mandiri dari database_va_npp.json
                kandidat_spk = []
                if va_clean and va_clean in db_va_kes:
                    kandidat_spk = list(db_va_kes[va_clean].get("spk_list", []))

                # Pastikan SPK yang dipilih PIC di Tahap 1 ikut masuk jika belum ada di kandidat
                if spk_billing_str and spk_billing_str != "-" and spk_billing_str not in kandidat_spk:
                    kandidat_spk.insert(0, spk_billing_str)

                if len(kandidat_spk) <= 1:
                    spk_single = kandidat_spk[0] if kandidat_spk else spk_billing_str
                    io_single = spk_io_map_cabang.get(spk_single, "-")
                    kertas_kerja_rows.append({
                        "ID Billing": f"BILL-{idx_b+1}",
                        "Nama Perusahaan": bu_str,
                        "Nomor VA": va_str,
                        "Tipe VA": "Tunggal",
                        "Total TK Billing": tk_billing_tot,
                        "Total Rp Billing": rp_billing_tot,
                        "Nomor Kontrak": spk_single,
                        "No IO": io_single,
                        "Jml TK SPK": tk_billing_tot
                    })
                else:
                    # VA Gabung: pecah menjadi beberapa baris sesuai jumlah SPK terdaftar
                    for s_item in kandidat_spk:
                        io_item = spk_io_map_cabang.get(s_item, "-")
                        kertas_kerja_rows.append({
                            "ID Billing": f"BILL-{idx_b+1}",
                            "Nama Perusahaan": bu_str,
                            "Nomor VA": va_str,
                            "Tipe VA": f"VA Gabung ({len(kandidat_spk)} SPK)",
                            "Total TK Billing": tk_billing_tot,
                            "Total Rp Billing": rp_billing_tot,
                            "Nomor Kontrak": s_item,
                            "No IO": io_item,
                            "Jml TK SPK": 0  # Diisi PIC di tabel kertas kerja
                        })

            df_kk_init = pd.DataFrame(kertas_kerja_rows)

            col_kk_config = {
                "ID Billing": st.column_config.TextColumn("ID", disabled=True, width="small"),
                "Nama Perusahaan": st.column_config.TextColumn("Badan Usaha", disabled=True),
                "Nomor VA": st.column_config.TextColumn("Nomor VA Mandiri", disabled=True),
                "Tipe VA": st.column_config.TextColumn("Status VA", disabled=True),
                "Total TK Billing": st.column_config.NumberColumn("TK Billing", format="%d", disabled=True),
                "Total Rp Billing": st.column_config.NumberColumn("Tagihan Billing (Rp)", format="Rp %d", disabled=True),
                "Nomor Kontrak": st.column_config.TextColumn("Nomor Kontrak (SPK)", width="large"),
                "No IO": st.column_config.TextColumn("No IO", disabled=True, width="small"),
                "Jml TK SPK": st.column_config.NumberColumn("✏️ Isi Jml TK SPK", format="%d", min_value=0, step=1)
            }

            st.caption("✏️ **Silakan isi kolom `✏️ Isi Jml TK SPK` untuk baris berstatus `VA Gabung` (Anda juga bisa tambah/hapus baris SPK jika perlu):**")
            df_kk_edited = st.data_editor(
                df_kk_init,
                column_config=col_kk_config,
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_kertas_kerja_spk_{selected_wilayah}_{len(uploaded_files_kes)}"
            )

            # Hitung alokasi prorata bulat per kelompok ID Billing / Nomor VA
            df_kk_calc = df_kk_edited.copy()
            df_kk_calc["Jml TK SPK"] = df_kk_calc["Jml TK SPK"].apply(safe_int)
            df_kk_calc["Total TK Billing"] = df_kk_calc["Total TK Billing"].apply(safe_int)
            df_kk_calc["Total Rp Billing"] = df_kk_calc["Total Rp Billing"].apply(custom_round)
            df_kk_calc["Tagihan SPK (Rp)"] = 0
            df_kk_calc["No IO"] = df_kk_calc["Nomor Kontrak"].apply(lambda s: spk_io_map_cabang.get(str(s).strip(), "-"))

            status_validasi_billing = []
            semua_va_valid = True

            for bill_id, grp in df_kk_calc.groupby("ID Billing", sort=False):
                idxs = grp.index.tolist()
                tot_tk_bill = int(grp["Total TK Billing"].iloc[0])
                tot_rp_bill = int(grp["Total Rp Billing"].iloc[0])
                va_label = str(grp["Nomor VA"].iloc[0])
                bu_label = str(grp["Nama Perusahaan"].iloc[0])
                tk_per_spk_list = grp["Jml TK SPK"].tolist()
                sum_tk_input = sum(tk_per_spk_list)

                alokasi_rp_list = alokasi_prorata_bulat(tot_rp_bill, tot_tk_bill, tk_per_spk_list)
                for i_pos, row_index in enumerate(idxs):
                    df_kk_calc.at[row_index, "Tagihan SPK (Rp)"] = alokasi_rp_list[i_pos]

                sum_rp_input = sum(alokasi_rp_list)
                selisih_tk = tot_tk_bill - sum_tk_input
                selisih_rp = tot_rp_bill - sum_rp_input

                if selisih_tk == 0 and selisih_rp == 0:
                    status_validasi_billing.append({
                        "ID": bill_id,
                        "Nomor VA": va_label,
                        "Badan Usaha": bu_label,
                        "TK Billing": tot_tk_bill,
                        "TK Input SPK": sum_tk_input,
                        "Tagihan Billing (Rp)": tot_rp_bill,
                        "Total Pecahan SPK (Rp)": sum_rp_input,
                        "Status": "✅ 100% Cocok"
                    })
                else:
                    semua_va_valid = False
                    status_validasi_billing.append({
                        "ID": bill_id,
                        "Nomor VA": va_label,
                        "Badan Usaha": bu_label,
                        "TK Billing": tot_tk_bill,
                        "TK Input SPK": sum_tk_input,
                        "Tagihan Billing (Rp)": tot_rp_bill,
                        "Total Pecahan SPK (Rp)": sum_rp_input,
                        "Status": f"⚠️ Selisih TK: {selisih_tk:+d} | Rp: {selisih_rp:+d}"
                    })

            # Tampilkan Panel Kontrol Rekonsiliasi (Billing vs Pecahan SPK)
            st.markdown("##### 📊 Kontrol Validasi & Hasil Kalkulasi Prorata per SPK")
            df_validasi_show = pd.DataFrame(status_validasi_billing)
            st.dataframe(df_validasi_show, use_container_width=True, hide_index=True)

            # Tampilkan Tabel Final Siap Sync ke Tab SESUAI SPK (Hanya SPK dengan TK > 0 atau Tagihan > 0)
            df_final_spk = df_kk_calc[(df_kk_calc["Jml TK SPK"] > 0) | (df_kk_calc["Tagihan SPK (Rp)"] > 0)].copy()
            df_final_spk.reset_index(drop=True, inplace=True)
            df_final_spk["NO"] = [i + 1 for i in range(len(df_final_spk))]

            df_final_spk_view = df_final_spk[[
                "NO", "Nama Perusahaan", "Nomor VA", "Nomor Kontrak", "No IO", "Jml TK SPK", "Tagihan SPK (Rp)"
            ]].rename(columns={
                "Jml TK SPK": "Jumlah Tenaga Kerja",
                "Tagihan SPK (Rp)": "Total Iuran"
            })

            st.markdown("##### 📋 Preview Tabel Final Rekap SESUAI SPK (Siap Sync & Export)")
            st.dataframe(
                df_final_spk_view,
                column_config={
                    "Total Iuran": st.column_config.NumberColumn("Total Iuran (Rp)", format="Rp %d")
                },
                use_container_width=True,
                hide_index=True
            )

            tot_tk_spk_all = int(df_final_spk_view["Jumlah Tenaga Kerja"].sum()) if not df_final_spk_view.empty else 0
            tot_rp_spk_all = int(df_final_spk_view["Total Iuran"].sum()) if not df_final_spk_view.empty else 0

            c_sp1, c_sp2, c_sp3 = st.columns(3)
            c_sp1.metric("Total TK (Sesuai SPK)", f"{tot_tk_spk_all} TK", delta=f"{tot_tk_spk_all - total_tk_kes:+d} vs Billing")
            c_sp2.metric("Total Tagihan (Sesuai SPK)", f"Rp {tot_rp_spk_all:,.0f}".replace(",", "."), delta=f"{tot_rp_spk_all - total_iuran_kes:+d} vs Billing")
            c_sp3.metric("Status Keseimbangan", "100% IDENTIK" if semua_va_valid else "BELUM SEIMBANG")

            if not semua_va_valid:
                st.warning("⚠️ Masih ada VA Gabung yang jumlah TK per SPK-nya belum sama dengan Total TK Billing. Lengkapi pengisian TK di atas agar hasil 100% identik sebelum sync ke Google Sheet.")

            # Download Excel Khusus Kertas Kerja & Rekap Sesuai SPK
            if not df_final_spk_view.empty:
                out_excel_spk = io.BytesIO()
                wb_spk = openpyxl.Workbook()
                ws_spk = wb_spk.active
                ws_spk.title = "Rekap Sesuai SPK"

                h_spk = ["NO", "Nama Perusahaan", "Nomor VA", "Nomor Kontrak (SPK)", "No IO", "Jumlah TK", "Total Tagihan (Rp)"]
                ws_spk.append(h_spk)
                for col_i in range(1, len(h_spk) + 1):
                    c = ws_spk.cell(row=1, column=col_i)
                    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                    c.fill = PatternFill(start_color="0F5132", end_color="0F5132", fill_type="solid")
                    c.alignment = Alignment(horizontal="center", vertical="center")

                for r_i, r_row in enumerate(df_final_spk_view.to_dict(orient="records"), start=2):
                    ws_spk.cell(row=r_i, column=1, value=int(r_row["NO"])).alignment = Alignment(horizontal="center")
                    ws_spk.cell(row=r_i, column=2, value=str(r_row["Nama Perusahaan"]))
                    ws_spk.cell(row=r_i, column=3, value=str(r_row["Nomor VA"])).alignment = Alignment(horizontal="center")
                    ws_spk.cell(row=r_i, column=4, value=str(r_row["Nomor Kontrak"]))
                    ws_spk.cell(row=r_i, column=5, value=str(r_row["No IO"])).alignment = Alignment(horizontal="center")
                    ws_spk.cell(row=r_i, column=6, value=int(r_row["Jumlah Tenaga Kerja"])).alignment = Alignment(horizontal="center")
                    c_rp = ws_spk.cell(row=r_i, column=7, value=int(r_row["Total Iuran"]))
                    c_rp.number_format = '"Rp "#,##0'
                    c_rp.alignment = Alignment(horizontal="right")

                for col in ws_spk.columns:
                    max_l = max(len(str(cell.value or '')) for cell in col)
                    ws_spk.column_dimensions[get_column_letter(col[0].column)].width = max(max_l + 3, 12)

                wb_spk.save(out_excel_spk)
                out_excel_spk.seek(0)

                st.download_button(
                    label="📥 Download Excel Kertas Kerja & Rekap SESUAI SPK",
                    data=out_excel_spk.getvalue(),
                    file_name=f"Rekap_BPJS_Kesehatan_SESUAI_SPK_{clean_filename_wilayah}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_excel_sesuai_spk"
                )

            # AUTO-SYNC GOOGLE SHEETS: TAB SESUAI SPK
            st.markdown("---")
            st.subheader("🚀 Sync Tahap 2 ke Google Sheets (Tab: SESUAI SPK)")

            target_sheet_choice_spk = st.radio(
                "Pilih Target Google Sheet:",
                ["Sheet Rekap Utama HPI", "Sheet Khusus Cabang ICON"],
                horizontal=True,
                key="target_sheet_choice_ks_spk"
            )

            try:
                sheet_url_spk = st.secrets["SPREADSHEET_URL_KES_ICON"] if target_sheet_choice_spk == "Sheet Khusus Cabang ICON" else st.secrets["SPREADSHEET_URL_KES"]
                available_tabs_spk = get_sheet_tabs_ks(sheet_url_spk)

                # Otomatis pilihkan tab yang mengandung kata "SESUAI SPK" jika ada
                default_idx_spk = next(
                    (i for i, t in enumerate(available_tabs_spk) if "sesuai spk" in t.lower()),
                    0
                )
                tab_sheet_target_spk = st.selectbox(
                    "Pilih Tab Google Sheet Target (Sesuai SPK):",
                    available_tabs_spk,
                    index=default_idx_spk,
                    key="target_sheet_ks_spk"
                )
            except Exception:
                st.warning("⚠️ Server Google Sheet API sedang sibuk (Error 503). Silakan ketik nama tab secara manual di bawah ini:")
                tab_sheet_target_spk = st.text_input("Nama Tab Google Sheet Target (Manual)", "2. NOV 26 SPSH SESUAI SPK", key="target_sheet_ks_spk_man")

            if st.button("⚡ Sync Tab SESUAI SPK ke Google Sheet Sekarang", type="primary", key="btn_sync_ks_spk"):
                if df_final_spk_view.empty:
                    st.error("Belum ada data SPK dengan Jumlah TK > 0 untuk dikirim.")
                else:
                    with st.spinner("Sedang mencocokkan pasangan VA + SPK dan mengirim data ke Tab Sesuai SPK..."):
                        is_icon_selected_spk = (target_sheet_choice_spk == "Sheet Khusus Cabang ICON")
                        success_spk, res_spk = sync_to_google_sheet_kesehatan(
                            df_final_spk_view,
                            tab_sheet_target_spk,
                            wilayah_display,
                            is_icon=is_icon_selected_spk,
                            mode_sync="SESUAI_SPK"
                        )
                        if success_spk:
                            st.success(f"BERHASIL! {res_spk['updated']} baris SPK berhasil di-Update secara presisi ke **{tab_sheet_target_spk}**!")
                            if res_spk["unmatched"]:
                                st.warning(
                                    "⚠️ Beberapa kombinasi VA + SPK berikut belum ditemukan di tab target Google Sheet "
                                    "(pastikan nomor kontrak di sheet sama dengan database):\n- " + "\n- ".join(res_spk["unmatched"])
                                )
                        else:
                            st.error(f"Gagal melakukan sync: {res_spk}")