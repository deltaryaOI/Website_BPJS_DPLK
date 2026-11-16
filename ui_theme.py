import streamlit as st

def apply_glass_theme():
    st.markdown("""
    <style>
        /* Latar Belakang Ambient Glow Halus */
        .stApp {
            background: 
                radial-gradient(circle at 12% 18%, rgba(0, 119, 182, 0.16) 0%, transparent 40%),
                radial-gradient(circle at 88% 15%, rgba(255, 122, 0, 0.14) 0%, transparent 38%),
                radial-gradient(circle at 50% 85%, rgba(2, 62, 138, 0.10) 0%, transparent 45%),
                linear-gradient(135deg, #E2E8F0 0%, #EDF2F7 50%, #DCE6F2 100%) !important;
        }
        .block-container {
            padding-top: 1.3rem;
            padding-bottom: 2rem;
            max-width: 96%;
        }

        /* Header Bawaan Streamlit Transparan Menyatu */
        header[data-testid="stHeader"] {
            background: transparent !important;
            box-shadow: none !important;
        }

        /* Sidebar Kaca Transparan Ringan */
        [data-testid="stSidebar"] {
            background: linear-gradient(
                180deg,
                rgba(0, 119, 182, 0.85) 0%,
                rgba(2, 62, 138, 0.88) 45%,
                rgba(10, 37, 64, 0.92) 100%
            ) !important;
            backdrop-filter: blur(6px);
            -webkit-backdrop-filter: blur(6px);
            border-right: 1px solid rgba(255, 255, 255, 0.35) !important;
            box-shadow: 4px 0 24px rgba(10, 37, 64, 0.12) !important;
        }
        [data-testid="stSidebar"] * {
            color: #FFFFFF !important;
        }
        [data-testid="stSidebarNav"] a {
            border-radius: 10px;
            margin: 4px 8px;
            padding: 8px 12px;
            border: 1px solid transparent;
            transition: background-color 0.2s ease, border-color 0.2s ease;
        }
        [data-testid="stSidebarNav"] a:hover {
            background-color: rgba(255, 255, 255, 0.14) !important;
            border-color: rgba(255, 255, 255, 0.28) !important;
        }
        [data-testid="stSidebarNav"] a[aria-current="page"] {
            background: linear-gradient(135deg, rgba(255, 122, 0, 0.92) 0%, rgba(255, 158, 68, 0.88) 100%) !important;
            border: 1px solid rgba(255, 255, 255, 0.45) !important;
            color: #FFFFFF !important;
            font-weight: 700 !important;
            box-shadow: 0 4px 12px rgba(255, 122, 0, 0.28);
        }

        /* Input & Selectbox (Halaman Utama & Sidebar) */
        div[data-baseweb="select"] > div,
        [data-testid="stSidebar"] div[data-baseweb="select"] > div {
            background-color: rgba(255, 255, 255, 0.96) !important;
            border: 1.5px solid #FF7A00 !important;
            border-radius: 10px !important;
            padding-right: 0 !important;
            overflow: hidden !important;
        }
        div[data-baseweb="select"] span,
        div[data-baseweb="select"] input,
        [data-testid="stSidebar"] div[data-baseweb="select"] span,
        [data-testid="stSidebar"] div[data-baseweb="select"] input {
            color: #0A2540 !important;
            font-weight: 600 !important;
        }

        /* Kotak Klik Dropdown di Sisi Kanan: Oranye + Panah Bawah Biru */
        div[data-baseweb="select"] > div > div:last-child,
        [data-testid="stSidebar"] div[data-baseweb="select"] > div > div:last-child {
            background: linear-gradient(135deg, #FF7A00 0%, #FF9E44 100%) !important;
            padding: 0 12px !important;
            margin: 0 !important;
            align-self: stretch !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            border-left: 1.5px solid #FF7A00 !important;
        }
        div[data-baseweb="select"]:hover > div > div:last-child,
        [data-testid="stSidebar"] div[data-baseweb="select"]:hover > div > div:last-child {
            background: linear-gradient(135deg, #E66C00 0%, #FF7A00 100%) !important;
        }

        /* Paksa Warna Panah Bawah Jadi Biru Tua */
        div[data-baseweb="select"] svg,
        div[data-baseweb="select"] svg path,
        [data-testid="stSidebar"] div[data-baseweb="select"] svg,
        [data-testid="stSidebar"] div[data-baseweb="select"] svg path {
            fill: #023E8A !important;
            color: #023E8A !important;
        }

        /* Top Header Box & Kontainer Kaca Ringan */
        .top-header-box, .glass-box {
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.78) 0%, rgba(255, 255, 255, 0.52) 100%);
            padding: 16px 22px;
            border-radius: 14px;
            box-shadow: 0 8px 24px rgba(10, 37, 64, 0.06), inset 0 1px 1px rgba(255, 255, 255, 0.95);
            border: 1px solid rgba(255, 255, 255, 0.8);
            margin-bottom: 16px;
        }
        .top-header-box {
            border-left: 5px solid #FF7A00;
            display: flex;
            justify-content: space-between;
            align-items: center;
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

        /* Kartu KPI & SLA Kaca Ringan (Tanpa Blur Berat) */
        .kpi-card {
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.75) 0%, rgba(255, 255, 255, 0.45) 100%);
            border-radius: 16px;
            padding: 20px 22px;
            box-shadow: 0 8px 24px rgba(10, 37, 64, 0.07), inset 0 1px 1px rgba(255, 255, 255, 0.95);
            border: 1px solid rgba(255, 255, 255, 0.85);
            display: flex;
            align-items: center;
            gap: 18px;
            height: 100%;
            transition: transform 0.22s ease, box-shadow 0.22s ease, background 0.22s ease;
        }
        .kpi-card:hover {
            transform: translateY(-4px);
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.88) 0%, rgba(255, 255, 255, 0.62) 100%);
            box-shadow: 0 12px 28px rgba(10, 37, 64, 0.12), inset 0 1px 1px rgba(255, 255, 255, 1);
            border-color: rgba(255, 122, 0, 0.55);
        }
        .kpi-icon-navy {
            background: linear-gradient(135deg, #0A2540 0%, #1E3A8A 100%);
            color: #FFFFFF;
            width: 54px; height: 54px; border-radius: 14px;
            display: flex; align-items: center; justify-content: center;
            font-size: 24px; flex-shrink: 0;
            box-shadow: 0 6px 14px rgba(10, 37, 64, 0.25);
        }
        .kpi-icon-orange {
            background: linear-gradient(135deg, #FF7A00 0%, #FF9E44 100%);
            color: #FFFFFF;
            width: 54px; height: 54px; border-radius: 14px;
            display: flex; align-items: center; justify-content: center;
            font-size: 24px; flex-shrink: 0;
            box-shadow: 0 6px 14px rgba(255, 122, 0, 0.28);
        }
        .kpi-icon-blue {
            background: linear-gradient(135deg, #0077B6 0%, #00B4D8 100%);
            color: #FFFFFF;
            width: 54px; height: 54px; border-radius: 14px;
            display: flex; align-items: center; justify-content: center;
            font-size: 24px; flex-shrink: 0;
            box-shadow: 0 6px 14px rgba(0, 119, 182, 0.25);
        }
        .kpi-title { font-size: 13px; color: #5A6B7C; font-weight: 600; margin-bottom: 3px; }
        .kpi-value { font-size: 22px; color: #0A2540; font-weight: 800; line-height: 1.2; }
        .kpi-sub { font-size: 12px; color: #16A34A; font-weight: 600; margin-top: 4px; }

        .sla-glass-card {
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.72) 0%, rgba(255, 255, 255, 0.42) 100%);
            border-radius: 14px;
            padding: 16px 20px;
            border: 1px solid rgba(255, 255, 255, 0.85);
            border-left: 4px solid #0077B6;
            box-shadow: 0 6px 20px rgba(10, 37, 64, 0.06), inset 0 1px 1px rgba(255, 255, 255, 0.95);
            transition: transform 0.22s ease, box-shadow 0.22s ease;
        }
        .sla-glass-card.orange-accent { border-left: 4px solid #FF7A00; }
        .sla-glass-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 10px 24px rgba(10, 37, 64, 0.11);
        }
        .sla-label { font-size: 12.5px; font-weight: 600; color: #5A6B7C; margin-bottom: 4px; }
        .sla-val { font-size: 24px; font-weight: 800; color: #0A2540; line-height: 1.2; }
        .sla-badge {
            display: inline-block; margin-top: 6px; padding: 2px 10px;
            border-radius: 12px; font-size: 11.5px; font-weight: 600;
            background-color: rgba(10, 37, 64, 0.07); color: #0A2540;
        }

        /* Styling st.metric bawaan di halaman lain agar otomatis jadi kartu kaca */
        div[data-testid="stMetric"] {
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.74) 0%, rgba(255, 255, 255, 0.44) 100%);
            border: 1px solid rgba(255, 255, 255, 0.85);
            border-left: 4px solid #0077B6;
            border-radius: 14px;
            padding: 14px 18px;
            box-shadow: 0 6px 20px rgba(10, 37, 64, 0.06), inset 0 1px 1px rgba(255, 255, 255, 0.95);
            transition: transform 0.22s ease, box-shadow 0.22s ease;
        }
        div[data-testid="stMetric"]:hover {
            transform: translateY(-3px);
            box-shadow: 0 10px 24px rgba(10, 37, 64, 0.11);
        }

        /* File Uploader & Expander bernuansa kaca */
        section[data-testid="stFileUploaderDropzone"],
        details[data-testid="stExpander"] {
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.75) 0%, rgba(255, 255, 255, 0.50) 100%) !important;
            border: 1px solid rgba(255, 255, 255, 0.85) !important;
            border-radius: 14px !important;
            box-shadow: 0 6px 20px rgba(10, 37, 64, 0.05) !important;
        }
        
        /* Tombol Utama (Primary): Biru Muda Bergaris Oranye + Hover Biru Lebih Gelap */
        button[data-testid="stBaseButton-primary"],
        button[kind="primary"],
        div[data-testid="stDownloadButton"] > button[data-testid="stBaseButton-primary"],
        div[data-testid="stDownloadButton"] > button[kind="primary"] {
            background: linear-gradient(135deg, #38BDF8 0%, #0EA5E9 100%) !important;
            color: #FFFFFF !important;
            border: 2px solid #FF7A00 !important;
            border-radius: 10px !important;
            font-weight: 700 !important;
            box-shadow: 0 4px 12px rgba(14, 165, 233, 0.25) !important;
            transition: all 0.22s ease !important;
        }
        button[data-testid="stBaseButton-primary"]:hover,
        button[kind="primary"]:hover,
        div[data-testid="stDownloadButton"] > button[data-testid="stBaseButton-primary"]:hover,
        div[data-testid="stDownloadButton"] > button[kind="primary"]:hover {
            background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%) !important;
            border-color: #FF7A00 !important;
            color: #FFFFFF !important;
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(2, 132, 199, 0.35) !important;
        }
        button[data-testid="stBaseButton-primary"]:active,
        button[kind="primary"]:active,
        div[data-testid="stDownloadButton"] > button[data-testid="stBaseButton-primary"]:active,
        div[data-testid="stDownloadButton"] > button[kind="primary"]:active {
            transform: translateY(0px);
        }
    </style>
    """, unsafe_allow_html=True)


def show_center_loader():
    placeholder = st.empty()
    placeholder.markdown("""
    <div style="
        position: fixed;
        inset: 0;
        z-index: 999999;
        display: flex;
        align-items: center;
        justify-content: center;
        background: rgba(244, 246, 249, 0.25);
        pointer-events: none;
    ">
        <div class="loadingio-spinner-spinner-rurz9ansjwk"><div class="ldio-freyurszfti">
        <div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div>
        </div></div>
    </div>
    <style type="text/css">
    @keyframes ldio-freyurszfti { 0% { opacity: 1 } 100% { opacity: 0 } }
    .ldio-freyurszfti div {
      left: 94px; top: 48px; position: absolute;
      animation: ldio-freyurszfti linear 1s infinite;
      background: #f8b26a; width: 12px; height: 24px;
      border-radius: 6px / 12px; transform-origin: 6px 52px;
    }
    .ldio-freyurszfti div:nth-child(1) { transform: rotate(0deg); animation-delay: -0.916s; background: #f8b26a; }
    .ldio-freyurszfti div:nth-child(2) { transform: rotate(30deg); animation-delay: -0.833s; background: #5cb4c9; }
    .ldio-freyurszfti div:nth-child(3) { transform: rotate(60deg); animation-delay: -0.75s; background: #f8b26a; }
    .ldio-freyurszfti div:nth-child(4) { transform: rotate(90deg); animation-delay: -0.666s; background: #5cb4c9; }
    .ldio-freyurszfti div:nth-child(5) { transform: rotate(120deg); animation-delay: -0.583s; background: #f8b26a; }
    .ldio-freyurszfti div:nth-child(6) { transform: rotate(150deg); animation-delay: -0.5s; background: #5cb4c9; }
    .ldio-freyurszfti div:nth-child(7) { transform: rotate(180deg); animation-delay: -0.416s; background: #f8b26a; }
    .ldio-freyurszfti div:nth-child(8) { transform: rotate(210deg); animation-delay: -0.333s; background: #5cb4c9; }
    .ldio-freyurszfti div:nth-child(9) { transform: rotate(240deg); animation-delay: -0.25s; background: #f8b26a; }
    .ldio-freyurszfti div:nth-child(10) { transform: rotate(270deg); animation-delay: -0.166s; background: #5cb4c9; }
    .ldio-freyurszfti div:nth-child(11) { transform: rotate(300deg); animation-delay: -0.083s; background: #f8b26a; }
    .ldio-freyurszfti div:nth-child(12) { transform: rotate(330deg); animation-delay: 0s; background: #5cb4c9; }
    .loadingio-spinner-spinner-rurz9ansjwk { width: 200px; height: 200px; display: inline-block; overflow: hidden; background: transparent; }
    .ldio-freyurszfti { width: 100%; height: 100%; position: relative; transform: translateZ(0) scale(0.75); backface-visibility: hidden; transform-origin: 100px 100px; }
    .ldio-freyurszfti div { box-sizing: content-box; }
    </style>
    """, unsafe_allow_html=True)
    return placeholder