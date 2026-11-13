import streamlit as st
import os
import json

# 1. Konfigurasi Halaman
st.set_page_config(
    page_title="DPLK (Dana Pensiun)",
    page_icon="🏦",
    layout="wide"
)

# 2. Custom Styling Tema Utama (Sesuai palette Biru Navy & Oranye)
st.markdown("""
    <style>
        [data-testid="stSidebar"] {
            background-color: #F0F4F9 !important;
            border-right: 3px solid #FF7A00 !important;
        }
        h1, h2, h3 { color: #0A2540 !important; font-weight: 700 !important; }
        
        .dev-card {
            background-color: #FFF8F0;
            border-left: 5px solid #FF7A00;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 25px;
        }
        .dev-badge {
            background-color: #FF7A00;
            color: white;
            padding: 4px 12px;
            border-radius: 15px;
            font-size: 0.85rem;
            font-weight: bold;
            display: inline-block;
            margin-bottom: 10px;
        }
    </style>
""", unsafe_allow_html=True)

# 3. Sidebar Setup
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
wilayah_list = list(db_kontrak.keys()) if db_kontrak else ["Cab. Palembang", "Cab. Medan", "Cab. Sibolga Nias", "Cab. Bangka Belitung"]

st.sidebar.header("📌 Pengaturan Dokumen DPLK")
selected_wilayah = st.sidebar.selectbox("Pilih Cabang / Wilayah", wilayah_list, disabled=True)
st.sidebar.markdown("---")
st.sidebar.info("💡 Halaman ini dalam mode pemeliharaan / pengembangan (On Development).")

# ==============================================================================
# HEADER & BANNER ON DEVELOPMENT
# ==============================================================================
st.title("🏦 Pengolahan & Rekapitulasi DPLK (Dana Pensiun)")
st.caption("Modul Otomasi Administrasi, Rekapitulasi Iuran & Surat Pengantar DPLK")

# Card Peringatan Feature On Development
st.markdown("""
    <div class="dev-card">
        <div class="dev-badge">🚧 FEATURE ON DEVELOPMENT</div>
        <h3 style="margin-top:0; color: #0A2540;">Modul DPLK Sedang Dalam Tahap Perancangan</h3>
        <p style="color: #555; margin-bottom: 10px;">
            Fitur pengolahan otomatis Dana Pensiun Lembaga Keuangan (DPLK) sedang disiapkan. 
            Modul ini nantinya akan mengolah data kepesertaan, pembagian porsi iuran pemberi kerja vs pekerja, 
            serta otomatisasi pencetakan Surat Pengantar ke Bank/Lembaga Pengelola DPLK.
        </p>
    </div>
""", unsafe_allow_html=True)

# Callout Alert Tambahan dari Streamlit
st.warning("⚠️ **Status Modul:** Non-Aktif (Locked). Fitur ini tidak mengganggu operasional BPJS TK dan BPJS Kesehatan yang sudah berjalan.")

st.markdown("---")

# ==============================================================================
# SKELETON / RANCANGAN KERANGKA UI MASA DEPAN (MODE DISABLED)
# ==============================================================================
st.subheader("📋 Kerangka Form DPLK (Preview UI Masa Depan)")
st.caption("Dibawah ini adalah bayangan antarmuka yang siap diisi kodingan jika kerangka data sudah final.")

col_a, col_b = st.columns(2)

with col_a:
    st.markdown("##### 1. Parameter Iuran DPLK")
    st.text_input("Lembaga Pengelola DPLK", value="DPLK Bank Mandiri / BNI / Muamalat", disabled=True)
    st.number_input("Persentase Porsi Perusahaan (%)", value=5.0, disabled=True)
    st.number_input("Persentase Porsi Pekerja (%)", value=2.5, disabled=True)

with col_b:
    st.markdown("##### 2. File Input & Sumber Data")
    st.file_uploader("Upload Rekap Iuran DPLK Bulanan (Excel/CSV/Billing)", type=["xlsx", "csv", "pdf"], disabled=True, key="upl_dplk_dummy")

st.markdown("---")

st.subheader("📊 Preview Tabel Hasil Olah Data DPLK")
st.caption("Tampilan simulasi hasil ekstraksi porsi iuran DPLK per karyawan.")

# Data Dummy untuk Simulasi Visualisasi
df_dummy = [
    {"NO": 1, "NIK": "10920001", "Nama Karyawan": "Fulan A", "Cabang": "Palembang", "Gaji Pokok": 5000000, "Iuran Perusahaan (5%)": 250000, "Potongan Pekerja (2.5%)": 125000, "Total Iuran DPLK": 375000},
    {"NO": 2, "NIK": "10920002", "Nama Karyawan": "Fulan B", "Cabang": "Palembang", "Gaji Pokok": 6000000, "Iuran Perusahaan (5%)": 300000, "Potongan Pekerja (2.5%)": 150000, "Total Iuran DPLK": 450000},
]
st.dataframe(df_dummy, use_container_width=True)

st.markdown("---")

# Action Buttons (Disabled)
c1, c2, c3 = st.columns(3)
with c1:
    st.button("📥 Download Master Excel DPLK", disabled=True, use_container_width=True)
with c2:
    st.button("📝 Export Surat Word DPLK", disabled=True, use_container_width=True)
with c3:
    st.button("🚀 Sync to Google Sheet DPLK", disabled=True, use_container_width=True)