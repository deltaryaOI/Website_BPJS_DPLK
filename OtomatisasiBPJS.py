import streamlit as st
import os
import pandas as pd
import json
import base64


# ==============================================================================
# 1. KONFIGURASI HALAMAN (HARUS PALING PERTAMA)
# ==============================================================================
st.set_page_config(
    page_title="SDM HPI Control Tower",
    page_icon="logo.png" if os.path.exists("logo.png") else "📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# 2. DEFINISI DEKLARASI HALAMAN (MULTI-PAGE ROUTING)
# ==============================================================================
page_home = st.Page(
    page="pages/0_🏠_Home_&_Dashboard.py",
    title="Home & Dashboard",
    icon="🏠",
    default=True
)

page_tk = st.Page(
    page="pages/1_🏢_BPJS_Ketenagakerjaan.py",
    title="BPJS Ketenagakerjaan",
    icon="🏢"
)

page_kes = st.Page(
    page="pages/2_🏥_BPJS_Kesehatan.py",
    title="BPJS Kesehatan",
    icon="🏥"
)

page_dplk = st.Page(
    page="pages/3_🏦_DPLK_Pensiun.py",
    title="DPLK Pensiun",
    icon="🏦"
)

page_upload = st.Page(
    page="pages/4_☁️_Upload_Berkas.py",
    title="Upload Berkas Drive",
    icon="☁️"
)

page_update_spk = st.Page(
    page="pages/5_🔄_Update_Database_SPK.py",
    title="Update Database SPK",
    icon="🔃"
)

pg = st.navigation({
    "": [
        page_home
    ],
    "📌 Main Menu": [
        page_tk,
        page_kes,
        page_dplk,
        page_upload,
        page_update_spk  # <-- Halaman khusus Update SPK masuk ke grup Main Menu
    ]
})

# ==============================================================================
# 3. CUSTOM CSS PRESISI (GARIS ORANYE MUNCUL KEMBALI)
# ==============================================================================
# Logo HPI tunggal di puncak paling atas sidebar (berlaku untuk semua menu)
b64_logo = ""
if os.path.exists("logo.png"):
    with open("logo.png", "rb") as img_f:
        b64_logo = base64.b64encode(img_f.read()).decode()

st.markdown(f"""
<style>
    .stApp {{
        background-color: #F4F6F9;
    }}
    /* Sidebar Gradasi Biru Muda ke Biru Tua */
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, #0077B6 0%, #023E8A 45%, #0A2540 100%) !important;
    }}
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {{
        color: #FFFFFF !important;
    }}
    /* Input & Dropdown Sidebar Tetap Putih dengan Teks Gelap */
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] div[data-baseweb="select"] > div {{
        background-color: #FFFFFF !important;
        color: #0A2540 !important;
    }}
    [data-testid="stSidebar"] div[data-baseweb="select"] span {{
        color: #0A2540 !important;
    }}

    /* LOGO HPI TRANSPARAN LANGSUNG DI ATAS MENU NAVIGASI */
    [data-testid="stSidebarNav"]::before {{
        content: "";
        display: block;
        background-color: transparent;
        background-image: url("data:image/png;base64,{b64_logo}");
        background-repeat: no-repeat;
        background-position: center center;
        background-size: 85% auto;
        height: 80px;
        margin: 8px 16px 14px 16px;
        padding-bottom: 14px;
        border-bottom: 4px solid #FFFFFF;
    }}

    /* 2. GARIS PUTIH TEBAL DI BAWAH MENU PALING BAWAH (UPDATE DATABASE SPK) */
    [data-testid="stSidebarNav"] {{
        border-bottom: 2px solid #FFFFFF !important;
        margin: 0 20px 20px 20px !important;
        padding-bottom: 14px !important;
    }}
    /* Menu Navigasi Sidebar */
    [data-testid="stSidebarNav"] a {{
        border-radius: 8px;
        margin: 4px 12px;
        padding: 8px 12px;
    }}
    [data-testid="stSidebarNav"] a:hover {{
        background-color: rgba(255, 255, 255, 0.15) !important;
    }}
    [data-testid="stSidebarNav"] a[aria-current="page"] {{
        background-color: #FF7A00 !important;
        font-weight: 700 !important;
    }}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 4. JALANKAN ROUTER NAVIGASI
# ==============================================================================
pg.run()
