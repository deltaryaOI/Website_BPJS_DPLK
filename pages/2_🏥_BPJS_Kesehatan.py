import streamlit as st
import pandas as pd
import pdfplumber
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
    page_title="BPJS Kesehatan",
    page_icon="🏥",
    layout="wide"
)

# Custom Styling Tema Utama
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

# Fungsi Auto-Sync Khusus BPJS Kesehatan
from openpyxl.utils import get_column_letter

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

def sync_to_google_sheet_kesehatan(df_edited, nama_tab_sheet, cabang_nama, is_icon=False):
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        client = gspread.authorize(creds)
        
        if is_icon:
            sheet_url = st.secrets["SPREADSHEET_URL_KES_ICON"]
            col_key_search = 8   # Kolom H: VIRTUAL ACCAOUNT BANK MANDIRI
            col_spk = 4          # Kolom D: NO KONTRAK
            col_tk = 9          # Sesuaikan nomor kolom Jumlah TK di sheet ICON
            col_iuran = 11       # Sesuaikan nomor kolom Total Tagihan di sheet ICON
        else:
            sheet_url = st.secrets["SPREADSHEET_URL_KES"]
            col_key_search = 8   # Kolom H: Nomor VA di sheet HPI
            col_spk = 4          # Kolom D: Nomor Kontrak
            col_tk = 9          # Kolom I: Jumlah TK
            col_iuran = 11       # Kolom K: Total Iuran

        sh = client.open_by_url(sheet_url)
        worksheet = sh.worksheet(nama_tab_sheet)
        
        # Bersihkan spasi dan karakter non-digit agar pencocokan VA 100% akurat
        raw_keys = worksheet.col_values(col_key_search)
        key_list = [re.sub(r'\D', '', str(x)) for x in raw_keys]
        
        let_spk = get_column_letter(col_spk)
        let_tk = get_column_letter(col_tk)
        let_iuran = get_column_letter(col_iuran)

        batch_data = []
        updated_count = 0
        records = df_edited.to_dict(orient="records")
        
        for row in records:
            kode_pdf = re.sub(r'\D', '', str(row.get("Nomor VA", "")))
            if kode_pdf and kode_pdf in key_list:
                row_idx = key_list.index(kode_pdf) + 1
                
                batch_data.append({
                    "range": f"{let_spk}{row_idx}",
                    "values": [[str(row["Nomor Kontrak"])]]
                })
                batch_data.append({
                    "range": f"{let_tk}{row_idx}",
                    "values": [[safe_int(row["Jumlah Tenaga Kerja"])]]
                })
                batch_data.append({
                    "range": f"{let_iuran}{row_idx}",
                    "values": [[custom_round(row["Total Iuran"])]]
                })
                updated_count += 1
                
        if batch_data:
            worksheet.batch_update(batch_data, value_input_option="USER_ENTERED")

        # ======================================================================
        # AUTO-LOGGING KE TAB MONITORING MASING-MASING (HPI vs ICON)
        # ======================================================================
        if updated_count > 0:
            try:
                wib_now = datetime.datetime.utcnow() + datetime.timedelta(hours=7)
                timestamp_str = wib_now.strftime("%Y-%m-%d %H:%M:%S")
                total_tk_log = int(df_edited["Jumlah Tenaga Kerja"].sum())
                total_tagihan_log = custom_round(df_edited["Total Iuran"].sum())

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
# KONTEN UTAMA: BPJS KESEHATAN
# ==============================================================================
st.title("🏥 Rekap PDF Billing Statement & Generator Surat BPJS Kesehatan")
st.caption("Unggah Billing Statement BPJS Kesehatan, edit data SPK, export Word/Excel, dan auto-sync ke Google Sheet.")

uploaded_files_kes = st.file_uploader(
    "Upload File PDF Billing Statement BPJS Kesehatan (Bisa banyak file sekaligus)", 
    type=["pdf"], 
    accept_multiple_files=True, 
    key="uploader_kes"
)

if uploaded_files_kes:
    st.info(f"Memproses {len(uploaded_files_kes)} file BPJS Kesehatan...")
    all_data_kes = []

    for file in uploaded_files_kes:
        try:
            with pdfplumber.open(file) as pdf:
                raw_text = ""
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        raw_text += text + "\n"

            clean_text = raw_text.replace('\n', ' ').replace('|', '')
            default_spk = kontrak_options[0] if (kontrak_options and len(kontrak_options) > 0) else "-"

            data_kes = {
                "Nama Perusahaan": "-",
                "Nomor VA": "-",
                "Nomor Kontrak": str(default_spk),
                "Jumlah Tenaga Kerja": 0,
                "Total Iuran": 0.0
            }

            # 1. NAMA PERUSAHAAN
            match_nama = re.search(r'Yth\.\s*(.*?)(?:\n\s*JL\.|\s*JL\.|\s*Nomor Virtual|\s*Pelaporan|\n\n|$)', raw_text, re.IGNORECASE | re.DOTALL)
            if match_nama:
                nama_bu = match_nama.group(1).replace('\n', ' ').strip()
                nama_bu = re.sub(r'(?:Surat Tagihan|Billing Statement|Pelaporan Penagihan|JKN-KIS).*$', '', nama_bu, flags=re.IGNORECASE).strip()
                data_kes["Nama Perusahaan"] = re.sub(r'\s+', ' ', nama_bu)

            # 2. NOMOR VA (DIUTAMAKAN MANDIRI / MDR)
            match_va_mdr = re.search(r'(?:MDR|Mandiri)\s*:?\s*([0-9]{13,16})', clean_text, re.IGNORECASE)
            if match_va_mdr:
                data_kes["Nomor VA"] = match_va_mdr.group(1).strip()
            else:
                match_va_gen = re.search(r'(?:BNI|BRI|BTN|Virtual Account)\s*:?\s*([0-9]{13,16})', clean_text, re.IGNORECASE)
                if match_va_gen:
                    data_kes["Nomor VA"] = match_va_gen.group(1).strip()

            # 3. JUMLAH TENAGA KERJA
            match_pegawai = re.search(r'(?:Rp\.?|Rp)?\s*[\d.]+\s*(?:Rp\.?|Rp)?\s*(\d+)\s+\d+\s+\d{2}/\d{2}/\d{4}', clean_text, re.IGNORECASE)
            if match_pegawai:
                data_kes["Jumlah Tenaga Kerja"] = int(match_pegawai.group(1).strip())
            else:
                match_peg_alt = re.search(r'(?:Jumlah Pegawai|Number of Employees).*?\s+(\d+)', clean_text, re.IGNORECASE)
                if match_peg_alt:
                    data_kes["Jumlah Tenaga Kerja"] = int(match_peg_alt.group(1).strip())

            # 4. TOTAL IURAN
            match_total = re.search(r'(?:TOTAL TAGIHAN|Total Balance).?(?:Rp\.?|Rp)\s([\d.]+)', clean_text, re.IGNORECASE)
            if not match_total:
                match_total = re.search(r'(?:Premi/Iuran|Contribution).?(?:Rp\.?|Rp)\s([\d.]+)', clean_text, re.IGNORECASE)

            if match_total:
                raw_angka = match_total.group(1).strip()
                clean_angka = raw_angka.replace('.', '')
                data_kes["Total Iuran"] = float(clean_angka)

            all_data_kes.append(data_kes)

        except Exception as e:
            st.error(f"Terjadi kesalahan saat membaca file {file.name}: {e}")

    if all_data_kes:
        df_kes_raw = pd.DataFrame(all_data_kes)
        
        st.subheader("✏️ Edit Data & Nomor Kontrak (SPK) BPJS Kesehatan")
        st.caption("💡 Klik 2x sel di bawah ini untuk mengedit data secara langsung.")

        with st.expander("📌 Lihat Daftar Nomor Kontrak (SPK) Cabang Ini"):
            for opt in kontrak_options:
                st.code(opt, language="text")

        column_config_kes = {
            "Nama Perusahaan": st.column_config.TextColumn("Nama Perusahaan / BU"),
            "Nomor VA": st.column_config.TextColumn("Nomor VA"),
            "Nomor Kontrak": st.column_config.TextColumn("Nomor Kontrak (SPK)", width="large"),
            "Jumlah Tenaga Kerja": st.column_config.NumberColumn("Jml TK", format="%d"),
            "Total Iuran": st.column_config.NumberColumn("Total Tagihan (Rp)", format="Rp %.2f")
        }

        edited_df_kes = st.data_editor(
            df_kes_raw,
            column_config=column_config_kes,
            num_rows="dynamic",
            use_container_width=True,
            key=f"editor_combine_kes_{selected_wilayah}_{len(uploaded_files_kes)}"
        )

        edited_df_kes.reset_index(drop=True, inplace=True)
        edited_df_kes["NO"] = [idx + 1 for idx in range(len(edited_df_kes))]

        total_iuran_kes = float(edited_df_kes["Total Iuran"].sum())
        total_tk_kes = int(edited_df_kes["Jumlah Tenaga Kerja"].sum())
        terbilang_kes = (terbilang(total_iuran_kes).strip() + " Rupiah") if total_iuran_kes > 0 else "-"

        st.markdown(f"### **Total Tagihan BPJS Kesehatan:** Rp {total_iuran_kes:,.0f}".replace(",", "."))
        st.markdown(f"**Terbilang:** *{terbilang_kes}*")

        st.subheader("📥 Unduh Hasil Dokumen BPJS Kesehatan")
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
                
                c_cell = ws_kes.cell(row=r_idx, column=6, value=float(row["Total Iuran"]))
                c_cell.number_format = '"Rp "#,##0.00'
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
                    val_tagihan = float(r_data["Total Iuran"])
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

        # AUTO-SYNC GOOGLE SHEETS KHUSUS KESEHATAN
        st.markdown("---")
        st.subheader("🚀 Auto-Sync ke Google Sheets (BPJS Kesehatan)")

        target_sheet_choice_ks = st.radio(
            "Pilih Target Google Sheet:",
            ["Sheet Rekap Utama HPI", "Sheet Khusus Cabang ICON"],
            horizontal=True,
            key="target_sheet_choice_ks"
        )

        try:
            scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
            client = gspread.authorize(creds)
            
            sheet_url_ks = st.secrets["SPREADSHEET_URL_KES_ICON"] if target_sheet_choice_ks == "Sheet Khusus Cabang ICON" else st.secrets["SPREADSHEET_URL_KES"]
            sh_ks = client.open_by_url(sheet_url_ks)
            
            available_tabs_ks = [ws.title for ws in sh_ks.worksheets()]
            
            tab_sheet_target_ks = st.selectbox(
                "Pilih Tab Google Sheet Target:", 
                available_tabs_ks,
                key="target_sheet_ks"
            )
        except Exception as e:
            st.warning("⚠️ Server Google Sheet API sedang sibuk (Error 503). Silakan ketik nama tab secara manual di bawah ini:")
            tab_sheet_target_ks = st.text_input("Nama Tab Google Sheet Target (Manual)", "1. NOV 26 SPSH SESUAI BILLI", key="target_sheet_ks")

        if st.button("⚡ Sync Data PDF BPJS Kesehatan ke Google Sheet Sekarang", type="primary", key="btn_sync_ks"):
            with st.spinner("Sedang mencocokkan Nomor VA dan mengirim data ke Google Sheet"):
                is_icon_selected_ks = (target_sheet_choice_ks == "Sheet Khusus Cabang ICON")
                success, res = sync_to_google_sheet_kesehatan(edited_df_kes, tab_sheet_target_ks, wilayah_display, is_icon=is_icon_selected_ks)
                
                if success:
                    st.success(f"BERHASIL! {res} data berhasil di-Update & dicatat ke Rekap_Monitoring_KS!")
                else:
                    st.error(f"Gagal melakukan sync: {res}")