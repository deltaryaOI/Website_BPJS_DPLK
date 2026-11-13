import streamlit as st
import pandas as pd
import pdfplumber
import requests
import base64
import mimetypes
import json
import os
import re
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
    page_title="BPJS Ketenagakerjaan",
    page_icon="🏢",
    layout="wide"
)

# Custom Styling (Mengikuti tema utama)
st.markdown("""
    <style>
        [data-testid="stSidebar"] {
            background-color: #F0F4F9 !important;
            border-right: 3px solid #FF7A00 !important;
        }
        h1, h2, h3 { color: #0A2540 !important; font-weight: 700 !important; }
    </style>
""", unsafe_allow_html=True)

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

def parse_amount(val_str):
    if not val_str:
        return 0.0
    clean_str = str(val_str).strip()
    if '.' in clean_str and ',' in clean_str:
        if clean_str.rfind(',') > clean_str.rfind('.'):
            clean_str = clean_str.replace('.', '').replace(',', '.')
        else:
            clean_str = clean_str.replace(',', '')
    elif ',' in clean_str and '.' not in clean_str:
        clean_str = clean_str.replace('.', '').replace(',', '.')
    else:
        clean_str = clean_str.replace(',', '')
    try:
        return float(clean_str)
    except ValueError:
        return 0.0

def lengkapi_npp_dan_kode_iuran(data_tk):
    """Otomatis melengkapi Kode Iuran dari NPP (4+NPP+000) atau sebaliknya"""
    npp = str(data_tk.get("NPP", "-")).strip()
    kode = str(data_tk.get("Kode Iuran", "-")).strip()

    # Jika NPP ada tapi Kode Iuran kosong -> Kode Iuran = 4 + NPP + 000
    if npp != "-" and npp.isdigit() and (kode == "-" or not kode):
        data_tk["Kode Iuran"] = f"4{npp}000"

    # Jika Kode Iuran ada tapi NPP kosong -> ambil digit setelah '4' dan sebelum '000'
    elif kode != "-" and kode.isdigit() and (npp == "-" or not npp):
        if kode.startswith("4") and kode.endswith("000") and len(kode) > 4:
            data_tk["NPP"] = kode[1:-3]
        elif len(kode) > 4:
            data_tk["NPP"] = kode[1:-3]

    return data_tk

def parse_f2a_pdf(pdf, default_spk):
    """Membaca PDF Rincian Tenaga Kerja / Form F2A multi-halaman"""
    data_tk = {
        "Nama Perusahaan": "-",
        "NPP": "-",
        "Nomor Kontrak": str(default_spk),
        "Kode Iuran": "-",
        "Jumlah Tenaga Kerja": 0,
        "Total Iuran JKK": 0.0,
        "Total Iuran JHT": 0.0,
        "Total Iuran JKM": 0.0,
        "Total Iuran JPN": 0.0,
        "Total Pembayaran": 0.0
    }

    max_no_tk = 0
    # 1. Ekstrak NPP & Nama Unit Kerja dari Halaman 1
    first_page_text = pdf.pages[0].extract_text() or ""
    lines = first_page_text.split('\n')
    for idx, line in enumerate(lines):
        if "NPP" in line and ("Nama Perusahaan" in line or "Nama Unit" in line):
            if idx + 1 < len(lines):
                val_line = lines[idx + 1].strip()
                parts = val_line.split()
                if parts and parts[0].isdigit():
                    data_tk["NPP"] = str(parts[0])
                    sisa_teks = " ".join(parts[1:]).strip()
                else:
                    sisa_teks = val_line

                split_unit = re.split(r'(?=HALEYORA|ICON)', sisa_teks, flags=re.IGNORECASE)
                split_unit = [s.strip() for s in split_unit if s.strip() and not s.strip().isdigit()]
                if split_unit:
                    data_tk["Nama Perusahaan"] = split_unit[-1]
                elif sisa_teks:
                    data_tk["Nama Perusahaan"] = sisa_teks
            break

    # 2. Loop semua halaman untuk mencari nomor urut terakhir & baris "Jumlah Seluruhnya"
    for page in pdf.pages:
        tables = page.extract_tables()
        for table in tables:
            for row in table:
                if not row:
                    continue
                
                first_col = str(row[0]).strip() if row[0] else ""
                if first_col.isdigit():
                    no_urut = int(first_col)
                    if no_urut > max_no_tk:
                        max_no_tk = no_urut

                row_joined = " ".join([str(c) for c in row if c])
                if "Jumlah Seluruhnya" in row_joined:
                    numeric_cells = []
                    for c in row:
                        if c and re.search(r'\d', str(c)) and "Jumlah" not in str(c):
                            numeric_cells.append(parse_amount(c))

                    if len(numeric_cells) >= 8:
                        jkk = numeric_cells[2]
                        jkm = numeric_cells[3]
                        jht = numeric_cells[4] + numeric_cells[5]
                        jpn = numeric_cells[6] + numeric_cells[7]
                        total = numeric_cells[-1] if len(numeric_cells) >= 11 else (jkk + jkm + jht + jpn)

                        data_tk["Total Iuran JKK"] = jkk
                        data_tk["Total Iuran JKM"] = jkm
                        data_tk["Total Iuran JHT"] = jht
                        data_tk["Total Iuran JPN"] = jpn
                        data_tk["Total Pembayaran"] = total

    data_tk["Jumlah Tenaga Kerja"] = max_no_tk
    return lengkapi_npp_dan_kode_iuran(data_tk)

def parse_npp_eps_pdf(pdf, default_spk):
    """Membaca PDF Billing NPP / EPS"""
    raw_text = ""
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            raw_text += text + "\n"

    clean_text = raw_text.replace('\n', ' ').replace('|', '')

    data_tk = {
        "Nama Perusahaan": "-",
        "NPP": "-",
        "Nomor Kontrak": str(default_spk),
        "Kode Iuran": "-",
        "Jumlah Tenaga Kerja": 0, 
        "Total Iuran JKK": 0.0,
        "Total Iuran JHT": 0.0,
        "Total Iuran JKM": 0.0,
        "Total Iuran JPN": 0.0,
        "Total Pembayaran": 0.0
    }

    # 1. Nama Perusahaan
    match_nama_sipp = re.search(r'Nama.?\s:\s*(.*?)\s(?:Total|Jumlah)\s*Tenaga', clean_text, re.IGNORECASE)
    match_nama_eps = re.search(r'PEMBAYARAN\s*UNTUK\s+(.*?)\s+NPP', clean_text, re.IGNORECASE)
    match_nama_eps_ringkas = re.search(r'Nama\s+Perusahaan\s*:\s*(.*?)(?:\s+Total\s+Tenaga|\s+Bulan|$)', clean_text, re.IGNORECASE)

    if match_nama_sipp:
        nama_pt = match_nama_sipp.group(1).replace('Perusahaan', '').replace('Kerja', '').strip()
        data_tk["Nama Perusahaan"] = re.sub(r'\s+', ' ', nama_pt)
    elif match_nama_eps:
        data_tk["Nama Perusahaan"] = re.sub(r'\s+', ' ', match_nama_eps.group(1).strip())
    elif match_nama_eps_ringkas:
        data_tk["Nama Perusahaan"] = re.sub(r'\s+', ' ', match_nama_eps_ringkas.group(1).strip())

    # 2. NPP
    match_npp = re.search(r'NPP\s*:\s*(\d+)', clean_text, re.IGNORECASE)
    if match_npp:
        data_tk["NPP"] = str(match_npp.group(1).strip())

    # 3. Kode Iuran
    match_kode = re.search(r'(?:KODE\s*IURAN\s*:?|KODE\s*PEMBAYARAN)\s*([0-9\s]{10,20})', clean_text, re.IGNORECASE)
    if match_kode: 
        data_tk["Kode Iuran"] = match_kode.group(1).replace(' ', '').strip()

    # 4. Jumlah Tenaga Kerja
    match_tk = re.search(r'(?:JUMLAH|Total)\s*TENAGA\s*KERJA\s*:\s*(\d+)', clean_text, re.IGNORECASE)
    if match_tk:
        data_tk["Jumlah Tenaga Kerja"] = int(match_tk.group(1).strip())

    # 5. Rincian Iuran
    match_jkk = re.search(r'JKK.*?(?:Rp\.?|Rp)?\s*([\d,.]+)', clean_text, re.IGNORECASE)
    if match_jkk:
        data_tk["Total Iuran JKK"] = parse_amount(match_jkk.group(1))

    match_jht = re.search(r'JHT.*?(?:Rp\.?|Rp)?\s*([\d,.]+)', clean_text, re.IGNORECASE)
    if match_jht:
        data_tk["Total Iuran JHT"] = parse_amount(match_jht.group(1))

    match_jkm = re.search(r'JKM.*?(?:Rp\.?|Rp)?\s*([\d,.]+)', clean_text, re.IGNORECASE)
    if match_jkm:
        data_tk["Total Iuran JKM"] = parse_amount(match_jkm.group(1))

    match_jpn = re.search(r'(?:JPN|JP)\s.*?(?:Rp\.?|Rp)?\s*([\d,.]+)', clean_text, re.IGNORECASE)
    if match_jpn:
        data_tk["Total Iuran JPN"] = parse_amount(match_jpn.group(1))

    # 6. Total Pembayaran
    match_total = re.search(r'(?:TOTAL\s*YANG\s*DIBAYARKAN|TOTAL\s*IURAN)\s*(?:\(\s*PEMBULATAN\s*\))?\s*:?\s*(?:Rp\.?|Rp)?\s*([\d,.]+)', clean_text, re.IGNORECASE)
    if match_total:
        data_tk["Total Pembayaran"] = parse_amount(match_total.group(1))

    return lengkapi_npp_dan_kode_iuran(data_tk)

from openpyxl.utils import get_column_letter

def sync_to_google_sheet(df_edited, nama_tab_sheet, cabang_nama, is_icon=False):
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        client = gspread.authorize(creds)
        
        if is_icon:
            sheet_url = st.secrets.get("SPREADSHEET_URL_ICON", st.secrets["SPREADSHEET_URL"])
            col_npp_search = 8   # Kolom H
            col_spk = 5          # Kolom E
            col_tk = 9           # Kolom I
            col_jkk = 14         # Kolom N
            col_jht = 15         # Kolom O
            col_jkm = 16         # Kolom P
            col_jpn = 17         # Kolom Q
        else:
            sheet_url = st.secrets["SPREADSHEET_URL"]
            col_npp_search = 9   # Kolom I
            col_spk = 5          # Kolom E
            col_tk = 10          # Kolom J
            col_jkk = 15         # Kolom O
            col_jht = 16         # Kolom P
            col_jkm = 17         # Kolom Q
            col_jpn = 18         # Kolom R
            
        sh = client.open_by_url(sheet_url)
        worksheet = sh.worksheet(nama_tab_sheet)
        
        # Ambil seluruh kolom NPP dan kolom SPK sekaligus untuk mendukung 1 NPP dengan banyak SPK
        npp_list = [str(x).strip() for x in worksheet.col_values(col_npp_search)]
        spk_list = [str(x).strip() for x in worksheet.col_values(col_spk)]
        
        let_spk = get_column_letter(col_spk)
        let_tk = get_column_letter(col_tk)
        let_jkk = get_column_letter(col_jkk)
        let_jht = get_column_letter(col_jht)
        let_jkm = get_column_letter(col_jkm)
        let_jpn = get_column_letter(col_jpn)

        batch_data = []
        updated_count = 0
        baris_terpakai = set()  # Mencegah baris yang sama tertimpa dua kali dalam 1 kali sync
        records = df_edited.to_dict(orient="records")
        
        for row in records:
            npp_pdf = str(row["NPP"]).strip()
            spk_pdf = str(row["Nomor Kontrak"]).strip()
            
            # Cari semua nomor baris di sheet yang memiliki NPP yang sama dan belum terpakai
            kandidat_baris = [
                idx + 1 for idx, val in enumerate(npp_list)
                if val == npp_pdf and (idx + 1) not in baris_terpakai
            ]
            
            row_idx = None
            if kandidat_baris:
                # Prioritas 1: Cari baris dengan NPP sama DAN SPK di sheet sudah sama
                for r in kandidat_baris:
                    spk_sheet = spk_list[r - 1] if (r - 1) < len(spk_list) else ""
                    if spk_sheet.lower() == spk_pdf.lower():
                        row_idx = r
                        break
                # Prioritas 2: Jika SPK di sheet masih kosong/berbeda, ambil baris kandidat pertama yang belum terpakai
                if row_idx is None:
                    row_idx = kandidat_baris[0]

            if row_idx is not None:
                baris_terpakai.add(row_idx)
                batch_data.append({
                    "range": f"{let_spk}{row_idx}",
                    "values": [[spk_pdf]]
                })
                batch_data.append({
                    "range": f"{let_tk}{row_idx}",
                    "values": [[safe_int(row["Jumlah Tenaga Kerja"])]]
                })
                batch_data.append({
                    "range": f"{let_jkk}{row_idx}:{let_jpn}{row_idx}",
                    "values": [[
                        custom_round(row["Total Iuran JKK"]),
                        custom_round(row["Total Iuran JHT"]),
                        custom_round(row["Total Iuran JKM"]),
                        custom_round(row["Total Iuran JPN"])
                    ]]
                })
                updated_count += 1
                
        if batch_data:
            worksheet.batch_update(batch_data, value_input_option="USER_ENTERED")

        # ======================================================================
        # AUTO-LOGGING KE TAB "Rekap_Monitoring_TK"
        # ======================================================================
        if updated_count > 0:
            try:
                sh_utama = client.open_by_url(st.secrets["SPREADSHEET_URL"])
                ws_mon = sh_utama.worksheet("Rekap_Monitoring_TK")
                wib_now = datetime.datetime.utcnow() + datetime.timedelta(hours=7)
                timestamp_str = wib_now.strftime("%Y-%m-%d %H:%M:%S")
                
                nama_cabang_log = "ICON" if is_icon else cabang_nama.upper()
                total_tk_log = int(df_edited["Jumlah Tenaga Kerja"].sum())
                total_tagihan_log = custom_round(df_edited["Total Pembayaran"].sum())
                
                ws_mon.append_row([timestamp_str, nama_cabang_log, total_tk_log, total_tagihan_log], value_input_option="USER_ENTERED")
            except Exception as log_err:
                st.warning(f"Sync NPP sukses, namun log monitoring gagal dicatat: {log_err}")
                
        return True, updated_count
    except Exception as e:
        return False, str(e)

# ==============================================================================
# DATABASE KONTRAK & SIDEBAR
# ==============================================================================
@st.cache_data
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

db_kontrak = load_database()

st.sidebar.header("📌 Pengaturan Dokumen Word")
wilayah_list = list(db_kontrak.keys()) if db_kontrak else ["Cab. Palembang", "Cab. Medan", "Cab. Sibolga Nias", "Cab. Bangka Belitung"]
selected_wilayah = st.sidebar.selectbox("Pilih Cabang / Wilayah", wilayah_list)

wilayah_display = re.sub(r'(?i)\bcab\.?\s*', '', selected_wilayah).strip()
data_cabang = db_kontrak.get(selected_wilayah, {})

nama_kota = data_cabang.get("kota", wilayah_display.title())
default_manager = data_cabang.get("manager", "NAMA MANAGER")
kontrak_options = data_cabang.get("spk", ["-"])

if not kontrak_options:
    kontrak_options = ["-"]

st.sidebar.markdown("---")
st.sidebar.header("📝 Informasi Header Surat")
nomor_surat = st.sidebar.text_input("Nomor Surat", "001/SP/HPI/X/2026")
tanggal_surat = st.sidebar.text_input("Tanggal Surat", value=tanggal_hari_ini) 
nama_manager = st.sidebar.text_input("Nama Manager Cabang", value=default_manager)

# ==============================================================================
# KONTEN UTAMA: BPJS KETENAGAKERJAAN
# ==============================================================================
st.title("🏢 Rekap PDF & Generator Surat BPJS Ketenagakerjaan")
st.caption("Unggah PDF NPP/EPS/F2A Ketenagakerjaan, edit data SPK, buat dokumen Word/Excel, dan sync ke Google Sheet.")

uploaded_files_tk = st.file_uploader(
    "Silahkan upload NPP/EPS/F2A (Bisa banyak & campur format sekaligus)", 
    type=["pdf"], 
    accept_multiple_files=True,
    key="uploader_tk"
)

if uploaded_files_tk:
    st.info(f"Memproses {len(uploaded_files_tk)} file BPJS Ketenagakerjaan...")
    all_data_tk = []
    default_spk = kontrak_options[0] if (kontrak_options and len(kontrak_options) > 0) else "-"

    for file in uploaded_files_tk:
        try:
            with pdfplumber.open(file) as pdf:
                first_text = (pdf.pages[0].extract_text() or "").lower()
                # Deteksi otomatis apakah dokumen berjenis F2A (Rincian Tenaga Kerja) atau NPP/EPS
                if "nomor referensi" in first_text or "iuran jht tk" in first_text or "jumlah seluruhnya" in first_text:
                    data_tk = parse_f2a_pdf(pdf, default_spk)
                else:
                    data_tk = parse_npp_eps_pdf(pdf, default_spk)

                all_data_tk.append(data_tk)
            
        except Exception as e:
            st.error(f"Terjadi kesalahan saat membaca file {file.name}: {e}")

    if all_data_tk:
        df_raw = pd.DataFrame(all_data_tk)
        
        st.subheader("✏️ Edit Data & Nomor Kontrak (SPK)")
        st.caption("💡 Klik 2x sel di bawah ini untuk mengedit data secara langsung.")

        with st.expander("📌 Lihat Daftar Nomor Kontrak (SPK) Cabang Ini"):
            for opt in kontrak_options:
                st.code(opt, language="text")

        column_config = {
            "Nama Perusahaan": st.column_config.TextColumn("Nama Perusahaan / Unit"),
            "NPP": st.column_config.TextColumn("NPP"),
            "Nomor Kontrak": st.column_config.TextColumn("Nomor Kontrak (SPK)", width="large"),
            "Kode Iuran": st.column_config.TextColumn("Kode Iuran"),
            "Jumlah Tenaga Kerja": st.column_config.NumberColumn("Jml TK", format="%d"),
            "Total Iuran JKK": st.column_config.NumberColumn("Total JKK", format="Rp %.2f"),
            "Total Iuran JHT": st.column_config.NumberColumn("Total JHT", format="Rp %.2f"),
            "Total Iuran JKM": st.column_config.NumberColumn("Total JKM", format="Rp %.2f"),
            "Total Iuran JPN": st.column_config.NumberColumn("Total JPN", format="Rp %.2f"),
            "Total Pembayaran": st.column_config.NumberColumn("Total Pembayaran (Rp)", format="Rp %.2f")
        }

        edited_df = st.data_editor(
            df_raw,
            column_config=column_config,
            num_rows="dynamic",
            use_container_width=True,
            key=f"editor_combine_{selected_wilayah}_{len(uploaded_files_tk)}"
        )

        edited_df.reset_index(drop=True, inplace=True)
        edited_df["NO"] = [idx + 1 for idx in range(len(edited_df))]

        total_iuran_sum = float(edited_df["Total Pembayaran"].sum())
        total_tk_sum = int(edited_df["Jumlah Tenaga Kerja"].sum())
        terbilang_iuran = (terbilang(total_iuran_sum).strip() + " Rupiah") if total_iuran_sum > 0 else "-"

        st.markdown(f"### **Total Anggaran Iuran:** Rp {total_iuran_sum:,.0f}".replace(",", "."))
        st.markdown(f"**Terbilang:** *{terbilang_iuran}*")

        st.subheader("📥 Unduh Hasil Dokumen")
        col1, col2 = st.columns(2)

        # 1. Download Excel
        with col1:
            st.markdown("##### 📊 Export Master Excel")
            clean_filename_wilayah = wilayah_display.replace(".", "").replace(" ", "_")
            excel_filename = f"Rekap_BPJS_TK_{clean_filename_wilayah}.xlsx"
            
            output_excel = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Rekap BPJS TK"

            font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            font_data = Font(name="Calibri", size=11)
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )

            headers = ["NO", "Nama Perusahaan", "NPP", "Nomor Kontrak", "Kode Iuran", "Jml TK", "Total JKK", "Total JHT", "Total JKM", "Total JPN", "Total Pembayaran"]
            ws.append(headers)
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = Alignment(horizontal="center", vertical="center")

            records = edited_df.to_dict(orient="records")
            for r_idx, row in enumerate(records, start=2):
                ws.cell(row=r_idx, column=1, value=int(row["NO"])).alignment = Alignment(horizontal="center")
                ws.cell(row=r_idx, column=2, value=str(row["Nama Perusahaan"])).alignment = Alignment(horizontal="left")
                ws.cell(row=r_idx, column=3, value=str(row["NPP"])).alignment = Alignment(horizontal="center")
                ws.cell(row=r_idx, column=4, value=str(row["Nomor Kontrak"])).alignment = Alignment(horizontal="left")
                ws.cell(row=r_idx, column=5, value=str(row["Kode Iuran"])).alignment = Alignment(horizontal="center")
                ws.cell(row=r_idx, column=6, value=safe_int(row["Jumlah Tenaga Kerja"])).alignment = Alignment(horizontal="center")
                
                for c_off, col_name in enumerate(["Total Iuran JKK", "Total Iuran JHT", "Total Iuran JKM", "Total Iuran JPN", "Total Pembayaran"], start=7):
                    c_cell = ws.cell(row=r_idx, column=c_off, value=float(row[col_name]))
                    c_cell.number_format = '"Rp "#,##0.00'
                    c_cell.alignment = Alignment(horizontal="right")

                for c_idx in range(1, 12):
                    c = ws.cell(row=r_idx, column=c_idx)
                    c.font = font_data
                    c.border = thin_border

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

            wb.save(output_excel)
            output_excel.seek(0)

            st.download_button(
                label="📥 Download Master Excel (.xlsx)",
                data=output_excel.getvalue(),
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # 2. Download Word Docx
        with col2:
            st.markdown("##### 📝 Export Surat Word (.docx)")
            template_path = os.path.join("templates", "Template_DEFAULT_ketenagakerjaan.docx")
            if not os.path.exists(template_path):
                template_path = os.path.join("templates", "Template_DEFAULT.docx")

            if os.path.exists(template_path):
                doc_tpl = DocxTemplate(template_path)
                periode_surat = extract_periode(tanggal_surat)

                context = {
                    "Kota": nama_kota,
                    "Wilayah": wilayah_display,
                    "Tanggal": tanggal_surat,
                    "Periode": periode_surat,
                    "NomorSurat": nomor_surat,
                    "Terbilang": terbilang_iuran,
                    "NamaManager": nama_manager,
                    "TOTAL_TK": str(total_tk_sum),
                    "TOTAL_IURAN": f"Rp {total_iuran_sum:,.0f}".replace(",", ".")
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
                    spacer_top = doc_native.add_paragraph()
                    spacer_top.paragraph_format.space_before = Pt(4)
                    spacer_top.paragraph_format.space_after = Pt(4)
                    target_p._p.addnext(spacer_top._p)
                    insert_point = spacer_top
                else:
                    insert_point = target_p

                new_table = doc_native.add_table(rows=0, cols=5)
                new_table.alignment = WD_TABLE_ALIGNMENT.CENTER

                if insert_point:
                    insert_point._p.addnext(new_table._tbl)

                hdr_cells = new_table.add_row().cells
                headers_docx = ["NO", "NPP", "Nomor Kontrak", "Jml TK", "Total Iuran"]
                widths = [Inches(0.5), Inches(1.3), Inches(3.1), Inches(0.8), Inches(1.8)]
                
                for idx, text in enumerate(headers_docx):
                    hdr_cells[idx].text = text
                    hdr_cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    hdr_cells[idx].paragraphs[0].runs[0].font.bold = True
                    set_cell_border(hdr_cells[idx])
                    hdr_cells[idx].width = widths[idx]

                records = edited_df.to_dict(orient="records")
                for row_data in records:
                    val_iuran = float(row_data["Total Pembayaran"])
                    formatted_iuran = f"Rp\t{val_iuran:,.0f}".replace(",", ".")

                    row_cells = new_table.add_row().cells
                    row_cells[0].text = str(row_data["NO"])
                    row_cells[1].text = str(row_data["NPP"])
                    row_cells[2].text = str(row_data["Nomor Kontrak"])
                    row_cells[3].text = str(safe_int(row_data["Jumlah Tenaga Kerja"]))
                    
                    p_iuran = row_cells[4].paragraphs[0]
                    p_iuran.text = formatted_iuran
                    p_iuran.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p_iuran.paragraph_format.tab_stops.add_tab_stop(Inches(1.75), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)

                    row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                    for c_idx in range(5):
                        set_cell_border(row_cells[c_idx])
                        row_cells[c_idx].width = widths[c_idx]

                total_cells = new_table.add_row().cells
                total_cells[0].text = "Jumlah"
                total_cells[3].text = str(total_tk_sum)

                p_tot_iuran = total_cells[4].paragraphs[0]
                p_tot_iuran.text = f"Rp\t{total_iuran_sum:,.0f}".replace(",", ".")
                p_tot_iuran.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p_tot_iuran.paragraph_format.tab_stops.add_tab_stop(Inches(1.75), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)
                if p_tot_iuran.runs:
                    p_tot_iuran.runs[0].font.bold = True

                total_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                total_cells[0].paragraphs[0].runs[0].font.bold = True
                total_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                total_cells[3].paragraphs[0].runs[0].font.bold = True

                for c_idx in range(5):
                    set_cell_border(total_cells[c_idx])
                    total_cells[c_idx].width = widths[c_idx]

                spacer_bottom = doc_native.add_paragraph()
                spacer_bottom.paragraph_format.space_before = Pt(6)
                spacer_bottom.paragraph_format.space_after = Pt(6)
                new_table._tbl.addnext(spacer_bottom._p)

                doc_io = io.BytesIO()
                doc_native.save(doc_io)
                doc_io.seek(0)
                
                word_filename = f"Surat_BPJS_TK_{clean_filename_wilayah}.docx"
                st.download_button(
                    label="📥 Download Surat Word (.docx)",
                    data=doc_io.getvalue(),
                    file_name=word_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.error("Template Word Ketenagakerjaan belum ditemukan di `templates/`!")

        # AUTO-SYNC GOOGLE SHEETS
        st.markdown("---")
        st.subheader("🚀 Auto-Sync ke Google Sheets (BPJS TK)")

        target_sheet_choice = st.radio(
            "Pilih Target Google Sheet:",
            ["Sheet Rekap Cabang HPI", "Sheet Rekap Cabang ICON"],
            horizontal=True,
            key="target_sheet_choice_tk"
        )

        try:
            scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
            client = gspread.authorize(creds)
            
            sheet_url = st.secrets.get("SPREADSHEET_URL_ICON") if target_sheet_choice == "Sheet Rekap Cabang ICON" else st.secrets["SPREADSHEET_URL"]
            sh = client.open_by_url(sheet_url)
            
            available_tabs = [ws.title for ws in sh.worksheets()]
            
            tab_sheet_target_tk = st.selectbox(
                "Pilih Tab Google Sheet Target:", 
                available_tabs,
                key="target_sheet_tk"
            )
        except Exception as e:
            st.warning(f"Tidak dapat memuat daftar tab secara otomatis: {e}")
            tab_sheet_target_tk = st.text_input("Nama Tab Google Sheet Target (Manual)", "01. REKAP 11.26", key="target_sheet_tk")

        if st.button("⚡ Sync Data PDF ke Google Sheet Sekarang", type="primary", key="btn_sync_tk"):
            with st.spinner("Sedang mencocokkan NPP dan mengirim data ke Google Sheet"):
                is_icon_selected = (target_sheet_choice == "Sheet Rekap Cabang ICON")
                success, res = sync_to_google_sheet(edited_df, tab_sheet_target_tk, wilayah_display, is_icon=is_icon_selected)
                if success:
                    st.success(f"BERHASIL! {res} data baris NPP berhasil di-Update & dicatat ke Rekap_Monitoring_TK!")
                else:
                    st.error(f"Gagal melakukan sync: {res}")
