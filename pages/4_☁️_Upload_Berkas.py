import streamlit as st
import requests
import base64
import mimetypes
import json
import os

# Konfigurasi Halaman Standalone
st.set_page_config(
    page_title="Upload & Arsip Berkas",
    page_icon="☁️",
    layout="wide"
)

# Custom Styling Body
st.markdown("""
    <style>
        h1, h2, h3 { color: #0A2540 !important; font-weight: 700 !important; }
        hr { border-top: 2px solid #FF7A00 !important; margin: 1rem 0 !important; }
    </style>
""", unsafe_allow_html=True)

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
wilayah_list = list(db_kontrak.keys()) if db_kontrak else [
    "Cab. Palembang", "Cab. Medan", "Cab. Sibolga Nias", "Cab. Bangka Belitung",
    "Cab. DKI", "Cab. Jawa Timur", "Cab. Jabar", "Cab. DIY", "Cab. Banten", "Cab. Padang", "Cab. Pekanbaru"
]

def upload_to_dynamic_drive_folder(file_bytes, file_name, jenis_bpjs, tahun_str, bulan_str, cabang_str):
    try:
        gas_url = st.secrets.get("GAS_WEB_APP_URL", "")
        if not gas_url:
            return False, "GAS_WEB_APP_URL belum dikonfigurasi di secrets.toml!"

        mime_type, _ = mimetypes.guess_type(file_name)
        if not mime_type:
            mime_type = 'application/octet-stream'

        base64_data = base64.b64encode(file_bytes).decode('utf-8')
        cabang_clean = cabang_str if cabang_str.startswith("Cabang") else f"Cabang {cabang_str.replace('Cab.', '').strip()}"

        payload = {
            "fileName": file_name,
            "mimeType": mime_type,
            "fileData": base64_data,
            "jenis": jenis_bpjs,
            "tahun": str(tahun_str),
            "bulan": str(bulan_str),
            "cabang": cabang_clean
        }

        response = requests.post(gas_url, json=payload, timeout=60)
        res_json = response.json()

        if res_json.get("status") == "success":
            return True, res_json.get("fileId")
        else:
            return False, res_json.get("message")

    except Exception as e:
        return False, str(e)

# ==============================================================================
# Halaman Upload Berkas Google Drive
# ==============================================================================
st.title("☁️ Upload & Pengarsipan Berkas Google Drive")
st.caption("Pusat penyimpanan berkas pendukung (PDF, RAR, ZIP, Excel, Word) langsung ke hirarki folder Google Drive.")

st.markdown("---")

col_d1, col_d2 = st.columns(2)

with col_d1:
    drive_kategori = st.selectbox("Target Kategori Folder:", ["BPJS Ketenagakerjaan", "BPJS Kesehatan", "DPLK"], key="drive_kategori")
    drive_cabang = st.selectbox("Target Cabang / Wilayah:", wilayah_list, key="drive_cabang")
    
with col_d2:
    drive_tahun = st.selectbox("Target Tahun Folder:", ["2026", "2027", "2028"], index=0, key="drive_tahun")
    drive_bulan = st.selectbox("Target Bulan Folder:", [
        "01. Januari", "02. Februari", "03. Maret", "04. April", 
        "05. Mei", "06. Juni", "07. Juli", "08. Agustus", 
        "09. September", "10. Oktober", "11. November", "12. Desember"
    ], index=10, key="drive_bulan")

drive_uploaded_files = st.file_uploader(
    "Pilih file yang akan diunggah ke Google Drive (Bisa upload sekaligus banyak file):",
    type=["pdf", "rar", "zip", "xlsx", "xls", "docx", "csv"],
    accept_multiple_files=True,
    key="uploader_google_drive"
)

if st.button("📤 Upload Berkas ke Google Drive Sekarang", type="primary", key="btn_upload_drive"):
    if drive_uploaded_files:
        with st.spinner("Sedang mengunggah berkas ke Google Drive..."):
            success_count = 0
            for f in drive_uploaded_files:
                f_bytes = f.getvalue()
                ok, res = upload_to_dynamic_drive_folder(
                    file_bytes=f_bytes,
                    file_name=f.name,
                    jenis_bpjs=drive_kategori,
                    tahun_str=drive_tahun,
                    bulan_str=drive_bulan,
                    cabang_str=drive_cabang
                )
                if ok:
                    success_count += 1
                else:
                    st.error(f"Gagal mengunggah {f.name}: {res}")
            
            if success_count > 0:
                st.success(f"BERHASIL! {success_count} berkas berhasil diarsipkan ke Google Drive di folder: {drive_kategori} / {drive_tahun} / {drive_bulan} / {drive_cabang}")
    else:
        st.warning("Pilih minimal 1 file terlebih dahulu sebelum menekan tombol upload!")
