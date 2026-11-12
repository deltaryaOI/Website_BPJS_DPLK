import streamlit as st
import os
import pandas as pd
import json


# ==============================================================================
# 1. KONFIGURASI HALAMAN (HARUS PALING PERTAMA)
# ==============================================================================
st.set_page_config(
    page_title="Dashboard Portal & SLA Monitoring BPJS",
    page_icon="🏠",
    layout="wide"
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
st.markdown("""
    <style>
        /* Background & Border Kanan Sidebar */
        [data-testid="stSidebar"] {
            background-color: #F0F4F9 !important;
            border-right: 3px solid #FF7A00 !important;
        }

        /* 1. Sembunyikan HANYA Garis Pemisah Bawaan Router Streamlit */
        [data-testid="stSidebarNav"] div:has(> hr),
        [data-testid="stSidebarNavSeparator"] {
            display: none !important;
        }

        /* 2. PAKSA MUNCULKAN GARIS ORANYE DI BAWAH NAVIGASI / DI ATAS LOGO */
        [data-testid="stSidebar"] > div:first-child > hr,
        [data-testid="stSidebarUserContent"] hr {
            display: block !important;
            border: none !important;
            border-top: 3px solid #FF7A00 !important;
            margin: 1.5rem 0 1rem 0 !important;
            opacity: 1 !important;
        }

        /* 3. Teks "Home & Dashboard" (Menu Atas) */
        [data-testid="stSidebarNavItems"] > li:first-child * {
            font-size: 24px !important;
            font-weight: 800 !important;
            color: #0A2540 !important;
        }

        /* 4. Header "📌 Main Menu" */
        [data-testid="stSidebarNavHeader"],
        [data-testid="stSidebarNavHeader"] * {
            font-size: 20px !important;
            font-weight: 800 !important;
            color: #0A2540 !important;
            padding-top: 10px !important;
        }

        /* 5. Sub-menu (BPJS TK, Kes, DPLK, Upload) */
        [data-testid="stSidebarNavItems"]:nth-of-type(2) * {
            font-size: 18px !important;
            font-weight: 600 !important;
            color: #0A2540 !important;
        }

        /* Highlight Menu Aktif */
        [data-testid="stSidebarNav"] a[aria-current="page"] {
            background-color: #E2EAF4 !important;
            border-radius: 8px !important;
        }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 4. LOGO HALEYORA POWERINDO & GARIS ORANYE TERPUSAT
# ==============================================================================
if os.path.exists("logo.png"):
    st.sidebar.markdown("---")  # Garis pemisah Oranye tebal 3px
    col1, col2, col3 = st.sidebar.columns([1, 3, 1])
    with col2:
        st.image("logo.png", use_container_width=True)

# ==============================================================================
# 5. JALANKAN ROUTER NAVIGASI
# ==============================================================================
pg.run()