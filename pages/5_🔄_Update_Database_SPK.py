import streamlit as st
import pandas as pd
import json
import os
import base64
import requests

# Set Page Config
st.set_page_config(
    page_title="Update Database SPK",
    page_icon="🔃",
    layout="wide"
)

# Custom Styling (Menyatu dengan tema gradasi sidebar biru & aksen oranye)
st.markdown("""
    <style>
        .stApp {
            background-color: #F4F6F9;
        }
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #0077B6 0%, #023E8A 45%, #0A2540 100%) !important;
        }
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            color: #FFFFFF !important;
        }
        [data-testid="stSidebarNav"] a {
            border-radius: 8px;
            margin: 4px 8px;
            padding: 8px 12px;
        }
        [data-testid="stSidebarNav"] a:hover {
            background-color: rgba(255, 255, 255, 0.15) !important;
        }
        [data-testid="stSidebarNav"] a[aria-current="page"] {
            background-color: #FF7A00 !important;
            font-weight: 700 !important;
        }
        [data-testid="stSidebarHeader"] {
            padding: 16px 16px 10px 16px !important;
        }
        [data-testid="stLogo"] {
            height: auto !important;
            max-height: none !important;
            width: 210px !important;
        }
        [data-testid="stSidebarHeader"] img {
            background-color: #FFFFFF;
            padding: 10px 18px;
            border-radius: 12px;
            height: 72px !important;
            width: 200px !important;
            max-width: none !important;
            object-fit: contain;
            box-shadow: 0 3px 8px rgba(0, 0, 0, 0.18);
        }
        h1, h2, h3 { color: #0A2540 !important; font-weight: 700 !important; }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# FUNGSI AUTO-PUSH KE GITHUB (HARUS DI ATAS TOMBOL SYNC)
# ==============================================================================
def push_database_spk_to_github(json_content_str, commit_msg="Auto-update database_kontrak.json via Web App"):
    """Otomatis memperbarui file database_kontrak.json di repository GitHub via REST API."""
    if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
        return False, "Token GitHub belum diatur di secrets.toml (hanya tersimpan lokal)."

    token = st.secrets["GITHUB_TOKEN"]
    repo = st.secrets["GITHUB_REPO"]
    file_path = "database_kontrak.json"
    api_url = f"https://api.github.com/repos/{repo}/contents/{file_path}"

    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    try:
        res_get = requests.get(api_url, headers=headers, timeout=15)
        sha_val = res_get.json().get("sha") if res_get.status_code == 200 else None

        encoded_content = base64.b64encode(json_content_str.encode("utf-8")).decode("utf-8")
        payload = {
            "message": commit_msg,
            "content": encoded_content,
            "branch": "main"
        }
        if sha_val:
            payload["sha"] = sha_val

        res_put = requests.put(api_url, headers=headers, json=payload, timeout=15)
        if res_put.status_code in [200, 201]:
            return True, "Berhasil tersinkronisasi permanen ke GitHub!"
        else:
            return False, f"Gagal push ke GitHub ({res_put.status_code}): {res_put.json().get('message', '')}"
    except Exception as e:
        return False, f"Error koneksi GitHub: {e}"

# Deteksi lokasi file database_kontrak.json
possible_paths = ["database_kontrak.json", "database.json", os.path.join("..", "database_kontrak.json")]
json_target_path = next((p for p in possible_paths if os.path.exists(p)), "database_kontrak.json")

st.title("🔃 Update & Manajemen Database SPK")
st.caption("Sinkronisasi otomatis daftar Nomor Kontrak (SPK) seluruh cabang dari file Excel ke `database_kontrak.json` & GitHub.")

# ==============================================================================
# 1. UPLOAD & SINKRONISASI EXCEL KE JSON + AUTO-PUSH GITHUB
# ==============================================================================
st.subheader("📤 Sinkronisasi via Excel SPK Terbaru")
st.info("💡 **Logika Sinkronisasi:** SPK baru di Excel akan ditambahkan, SPK yang sama tetap dipertahankan (tanpa duplikat), dan SPK lama yang sudah tidak ada di Excel akan otomatis dihapus.")

uploaded_spk_excel = st.file_uploader(
    "Pilih File Excel Daftar SPK (.xlsx)", 
    type=["xlsx"], 
    key="uploader_spk_page"
)

if uploaded_spk_excel is not None:
    if st.button("⚡ Sinkronkan SPK Sekarang", type="primary", key="btn_sync_spk_page"):
        try:
            df_spk_new = pd.read_excel(uploaded_spk_excel)
            
            col_to_json_map = {
                "CAB. BABEL": "Cab. Bangka Belitung",
                "CAB. DKI": "Cab. DKI Jakarta",
                "CAB. BANTEN": "Cab. Banten",
                "CAB. JABAR": "Cab. Jawa Barat",
                "CAB. PADANG": "Cab. Padang",
                "CAB. PALEMBANG": "Cab. Palembang",
                "CAB. JATENG": "Cab. Jawa Tengah",
                "CAB. ICON": "Cab. Icon",
                "CAB. SIBOLGA NIAS": "Cab. Sibolga Nias",
                "CAB. MEDAN": "Cab. Medan",
                "CAB. JATIM": "Cab. Jawa Timur",
                "CAB. PEKANBARU": "Cab. Pekanbaru"
            }

            with open(json_target_path, "r", encoding="utf-8") as f:
                current_db = json.load(f)

            total_added = 0
            total_removed = 0
            detail_log = []

            for excel_col, json_key in col_to_json_map.items():
                matched_col = next((c for c in df_spk_new.columns if str(c).strip().upper() == excel_col), None)
                
                if matched_col and json_key in current_db:
                    excel_list = []
                    for val in df_spk_new[matched_col].dropna().tolist():
                        clean_val = str(val).strip()
                        if clean_val and clean_val not in excel_list:
                            excel_list.append(clean_val)

                    old_list = current_db[json_key].get("spk", [])

                    kept_list = []
                    added_cabang = 0
                    removed_cabang = 0

                    for old_spk in old_list:
                        if old_spk in excel_list and old_spk not in kept_list:
                            kept_list.append(old_spk)
                        elif old_spk not in excel_list:
                            total_removed += 1
                            removed_cabang += 1

                    for new_spk in excel_list:
                        if new_spk not in kept_list:
                            kept_list.append(new_spk)
                            total_added += 1
                            added_cabang += 1

                    current_db[json_key]["spk"] = kept_list
                    if added_cabang > 0 or removed_cabang > 0:
                        detail_log.append(f"- **{json_key}**: `+{added_cabang}` SPK baru, `-{removed_cabang}` SPK dihapus")

            # Simpan lokal ke database_kontrak.json
            json_str = json.dumps(current_db, indent=2, ensure_ascii=False)
            with open(json_target_path, "w", encoding="utf-8") as f:
                f.write(json_str)

            # Otomatis push ke GitHub agar permanen di Streamlit Cloud
            ok_gh, msg_gh = push_database_spk_to_github(json_str)

            st.session_state["spk_sync_msg_page"] = {
                "added": total_added,
                "removed": total_removed,
                "details": detail_log,
                "gh_ok": ok_gh,
                "gh_msg": msg_gh
            }

            st.cache_data.clear()
            st.rerun()

        except Exception as e:
            st.error(f"Gagal memproses Excel SPK: {e}")

# Notifikasi Persisten Setelah Rerun
if "spk_sync_msg_page" in st.session_state:
    msg = st.session_state["spk_sync_msg_page"]
    st.success(f"✅ **Sinkronisasi Berhasil!** (+{msg['added']} SPK baru ditambahkan | -{msg['removed']} SPK lama dihapus)")
    if msg.get("gh_ok"):
        st.success(f"☁️ **GitHub Cloud Sync:** {msg.get('gh_msg')}")
    else:
        st.warning(f"⚠️ **GitHub Cloud Sync:** {msg.get('gh_msg')}")

    if msg["details"]:
        st.markdown("\n".join(msg["details"]))
    else:
        st.info("Semua nomor SPK sudah sama persis dengan Excel (0 perubahan).")
    if st.button("Tutup Notifikasi", key="close_spk_notif_page"):
        del st.session_state["spk_sync_msg_page"]
        st.rerun()

st.markdown("---")

# ==============================================================================
# 2. PREVIEW & BACKUP DATABASE KONTRAK SAAT INI
# ==============================================================================
st.subheader("📋 Preview Database SPK Aktif per Cabang")

if os.path.exists(json_target_path):
    with open(json_target_path, "r", encoding="utf-8") as f:
        db_view = json.load(f)

    col_sel, col_dl = st.columns([2, 1])
    with col_sel:
        cabang_preview = st.selectbox("Pilih Cabang untuk Melihat Daftar SPK Aktif:", list(db_view.keys()))
    with col_dl:
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label="📥 Download Backup database_kontrak.json",
            data=json.dumps(db_view, indent=2, ensure_ascii=False),
            file_name="database_kontrak.json",
            mime="application/json",
            use_container_width=True
        )

    info_cab = db_view.get(cabang_preview, {})
    st.markdown(f"**Kota:** `{info_cab.get('kota', '-')}` &nbsp;|&nbsp; **Manager:** `{info_cab.get('manager', '-')}` &nbsp;|&nbsp; **Total SPK Aktif:** `{len(info_cab.get('spk', []))}`")

    spk_list_aktif = info_cab.get("spk", [])

    if spk_list_aktif:
        rows_html = ""
        for idx, spk in enumerate(spk_list_aktif, start=1):
            rows_html += (
                f'<tr style="border-bottom: 1px solid #E6EAF1;">'
                f'<td style="width: 60px; text-align: center; padding: 10px 8px; color: #31333F; font-weight: 500;">{idx}</td>'
                f'<td style="padding: 10px 14px; color: #0A2540; font-family: monospace; font-size: 14px;">{spk}</td>'
                f'</tr>'
            )

        table_html = (
            '<div style="width: 100%; overflow-x: auto; border: 1px solid #D6DCE5; border-radius: 8px; margin-top: 8px;">'
            '<table style="width: 100%; border-collapse: collapse; background-color: #FFFFFF; font-size: 14px;">'
            '<thead>'
            '<tr style="background-color: #F0F4F9; border-bottom: 2px solid #FF7A00; text-align: left;">'
            '<th style="width: 60px; text-align: center; padding: 12px 8px; color: #0A2540; font-weight: 700;">No</th>'
            '<th style="padding: 12px 14px; text-align: center; color: #0A2540; font-weight: 700;">Nomor Kontrak (SPK)</th>'
            '</tr>'
            '</thead>'
            f'<tbody>{rows_html}</tbody>'
            '</table>'
            '</div>'
        )
        st.markdown(table_html, unsafe_allow_html=True)
    else:
        st.info("Belum ada daftar SPK aktif untuk cabang ini.")
else:
    st.warning("File `database_kontrak.json` belum ditemukan.")