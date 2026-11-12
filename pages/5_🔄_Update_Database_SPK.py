import streamlit as st
import pandas as pd
import json
import os

# Set Page Config
st.set_page_config(
    page_title="Update Database SPK",
    page_icon="🔃",
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

# Deteksi lokasi file database_kontrak.json
possible_paths = ["database_kontrak.json", "database.json", os.path.join("..", "database_kontrak.json")]
json_target_path = next((p for p in possible_paths if os.path.exists(p)), "database_kontrak.json")

st.title("🔃 Update & Manajemen Database SPK")
st.caption("Sinkronisasi otomatis daftar Nomor Kontrak (SPK) seluruh cabang dari file Excel ke `database_kontrak.json`.")

# ==============================================================================
# 1. UPLOAD & SINKRONISASI EXCEL KE JSON
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

                    # 1. Pertahankan SPK lama yang masih ada di Excel
                    for old_spk in old_list:
                        if old_spk in excel_list and old_spk not in kept_list:
                            kept_list.append(old_spk)
                        elif old_spk not in excel_list:
                            total_removed += 1
                            removed_cabang += 1

                    # 2. Tambahkan SPK baru dari Excel
                    for new_spk in excel_list:
                        if new_spk not in kept_list:
                            kept_list.append(new_spk)
                            total_added += 1
                            added_cabang += 1

                    current_db[json_key]["spk"] = kept_list
                    if added_cabang > 0 or removed_cabang > 0:
                        detail_log.append(f"- **{json_key}**: `+{added_cabang}` SPK baru, `-{removed_cabang}` SPK dihapus")

            with open(json_target_path, "w", encoding="utf-8") as f:
                json.dump(current_db, f, indent=2, ensure_ascii=False)

            st.session_state["spk_sync_msg_page"] = {
                "added": total_added,
                "removed": total_removed,
                "details": detail_log
            }

            st.cache_data.clear()
            st.rerun()

        except Exception as e:
            st.error(f"Gagal memproses Excel SPK: {e}")

# Notifikasi Persisten Setelah Rerun
if "spk_sync_msg_page" in st.session_state:
    msg = st.session_state["spk_sync_msg_page"]
    st.success(f"✅ **Sinkronisasi Berhasil!** (+{msg['added']} SPK baru ditambahkan | -{msg['removed']} SPK lama dihapus)")
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