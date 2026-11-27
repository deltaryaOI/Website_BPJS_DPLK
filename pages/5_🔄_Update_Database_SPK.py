import streamlit as st
import pandas as pd
import json
import os
import re
import base64
import requests

# Set Page Config
st.set_page_config(
    page_title="Update Database SPK & IO",
    page_icon="🔃",
    layout="wide"
)

# Custom Styling (Menyatu dengan tema gradasi sidebar biru & aksen oranye)
from ui_theme import apply_glass_theme

apply_glass_theme()

# ==============================================================================
# KONFIGURASI MAPPING UNIT KERJA EXCEL -> KEY DATABASE JSON
# ==============================================================================
UNIT_TO_JSON_MAP = {
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
    "CAB. PEKANBARU": "Cab. Pekanbaru",
    "HPI PUSAT": "HPI Pusat"
}

DEFAULT_META_CABANG = {
    "Cab. Bangka Belitung": {"kota": "Bangka", "manager": "Ardi Pangestu"},
    "Cab. DKI Jakarta": {"kota": "Jakarta", "manager": "Cecep Iwan"},
    "Cab. Banten": {"kota": "Banten", "manager": "Gugus Waluyo Pambudi"},
    "Cab. Jawa Barat": {"kota": "Bandung", "manager": "Eko Satria"},
    "Cab. Padang": {"kota": "Padang", "manager": "Ridwan Syahril Ismail"},
    "Cab. Palembang": {"kota": "Palembang", "manager": "Debby Marison"},
    "Cab. Jawa Tengah": {"kota": "Semarang", "manager": "Akhmad Bajuri"},
    "Cab. Icon": {"kota": "Jakarta", "manager": "Muhammad Chaidir Safikri Kesuma"},
    "Cab. Sibolga Nias": {"kota": "Sibolga", "manager": "Rizal Taher Khadafi"},
    "Cab. Medan": {"kota": "Medan", "manager": "Rizal Taher Khadafi"},
    "Cab. Jawa Timur": {"kota": "Surabaya", "manager": "Didik Supriyanto"},
    "Cab. Pekanbaru": {"kota": "Pekanbaru", "manager": "Shelly Indah Permata Sari"},
    "HPI Pusat": {"kota": "Jakarta", "manager": "NAMA MANAGER"}
}

# ==============================================================================
# FUNGSI HELPER PARSING & AUTO-PUSH GITHUB
# ==============================================================================
def clean_io_str(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return ""
    try:
        if "e+" in s.lower() or s.endswith(".0"):
            return str(int(float(s)))
    except ValueError:
        pass
    if s.endswith(".0"):
        s = s[:-2]
    return s

def find_header_row_kontrak(excel_file):
    """Mendeteksi baris header secara otomatis pada file master kontrak kantor."""
    df_preview = pd.read_excel(excel_file, header=None, nrows=15)
    for idx, row in df_preview.iterrows():
        row_upper = [str(c).replace("\n", " ").strip().upper() for c in row.tolist()]
        has_unit = any("UNIT KERJA" in c for c in row_upper)
        has_spk = any("NOMOR KONTRAK" in c or "SPK" in c for c in row_upper)
        if has_unit and has_spk:
            return idx
    return 2

def parse_master_kontrak_excel(excel_file, hanya_aktif=True, gunakan_perubahan_spk=False):
    """Mengekstrak Cabang, Nomor SPK, dan Internal Order (IO) dari file Kontrak Update kantor."""
    excel_file.seek(0)
    header_idx = find_header_row_kontrak(excel_file)
    excel_file.seek(0)
    df = pd.read_excel(excel_file, header=header_idx)
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

    col_status = next((c for c in df.columns if c.upper() == "KET STATUS"), None)
    col_io = next((c for c in df.columns if c.upper() in ["INT_ORDER", "INTERNAL ORDER", "IO", "NOMOR IO"]), None)
    col_unit = next((c for c in df.columns if "UNIT KERJA" in c.upper()), None)
    col_spk = next((c for c in df.columns if c.upper() in ["NOMOR KONTRAK / SPK", "NOMOR KONTRAK", "NO SPK"]), None)
    col_perubahan = next((c for c in df.columns if "PERUBAHAN NO KONTRAK" in c.upper()), None)

    if not col_unit or not col_spk:
        raise ValueError(
            f"Kolom wajib ('Unit Kerja PT HPI' dan 'Nomor Kontrak / SPK') tidak ditemukan. "
            f"Kolom terbaca: {list(df.columns[:10])}"
        )

    df_valid = df[df[col_unit].notna() & df[col_spk].notna()].copy()

    if hanya_aktif and col_status:
        df_valid = df_valid[df_valid[col_status].astype(str).str.strip().str.upper() == "AKTIF"].copy()

    extracted_rows = []
    unmapped_units = set()

    for _, r in df_valid.iterrows():
        unit_raw = str(r[col_unit]).strip().upper()
        json_key = UNIT_TO_JSON_MAP.get(unit_raw)
        if not json_key:
            unmapped_units.add(unit_raw)
            json_key = unit_raw.title()

        spk_asal = str(r[col_spk]).strip()
        spk_baru = str(r[col_perubahan]).strip() if (col_perubahan and pd.notna(r[col_perubahan])) else ""
        if spk_baru.lower() == "nan":
            spk_baru = ""

        spk_final = spk_baru if (gunakan_perubahan_spk and spk_baru) else spk_asal
        io_val = clean_io_str(r[col_io]) if col_io else ""
        status_val = str(r[col_status]).strip() if (col_status and pd.notna(r[col_status])) else "-"

        if spk_final and spk_final.lower() != "nan":
            extracted_rows.append({
                "Unit Excel": unit_raw,
                "Cabang JSON": json_key,
                "Nomor Kontrak (SPK)": spk_final,
                "SPK Perubahan": spk_baru if spk_baru else "-",
                "Internal Order (IO)": io_val if io_val else "-",
                "Status": status_val
            })

    return pd.DataFrame(extracted_rows), sorted(list(unmapped_units))

def parse_mapping_va_npp_excel(excel_file):
    """Membaca file Mapping NPP dan VA ALL.xlsx (sheet BPJS Kesehatan & BPJS Ketenagakerjaan) ke struktur JSON."""
    xls = pd.ExcelFile(excel_file)
    result = {"kesehatan": {}, "ketenagakerjaan": {}}

    for sheet_name in xls.sheet_names:
        sname_low = sheet_name.strip().lower()
        target_key = None
        if "kesehatan" in sname_low:
            target_key = "kesehatan"
        elif "ketenagakerjaan" in sname_low:
            target_key = "ketenagakerjaan"

        if not target_key:
            continue

        df_s = pd.read_excel(excel_file, sheet_name=sheet_name)
        df_s.columns = [str(c).strip().upper() for c in df_s.columns]

        col_cab = next((c for c in df_s.columns if "CABANG" in c), None)
        col_kode = next((c for c in df_s.columns if "NPP" in c or "VA" in c), None)
        col_spk = next((c for c in df_s.columns if "SPK" in c or "KONTRAK" in c), None)

        if not col_kode or not col_spk:
            continue

        for _, r in df_s.iterrows():
            cab = str(r[col_cab]).strip().upper() if (col_cab and pd.notna(r[col_cab])) else "-"
            kode_raw = str(r[col_kode]).strip() if pd.notna(r[col_kode]) else ""
            kode_clean = re.sub(r'\D', '', kode_raw)
            spk_clean = str(r[col_spk]).strip() if pd.notna(r[col_spk]) else ""
            if spk_clean.lower() == "nan":
                spk_clean = ""

            if kode_clean:
                if kode_clean not in result[target_key]:
                    result[target_key][kode_clean] = {
                        "cabang": cab,
                        "spk_list": []
                    }
                if spk_clean and spk_clean not in result[target_key][kode_clean]["spk_list"]:
                    result[target_key][kode_clean]["spk_list"].append(spk_clean)

    return result

def push_file_to_github(file_path_repo, json_content_str, commit_msg):
    """Otomatis memperbarui file JSON di repository GitHub via REST API."""
    if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
        return False, "Token GitHub belum diatur di secrets.toml (tersimpan lokal di perangkat ini)."

    token = st.secrets["GITHUB_TOKEN"]
    repo = st.secrets["GITHUB_REPO"]
    api_url = f"https://api.github.com/repos/{repo}/contents/{file_path_repo}"

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
            return True, f"Berhasil tersinkronisasi permanen ke GitHub (`{file_path_repo}`)!"
        else:
            return False, f"Gagal push ke GitHub ({res_put.status_code}): {res_put.json().get('message', '')}"
    except Exception as e:
        return False, f"Error koneksi GitHub: {e}"

# Deteksi lokasi file JSON (otomatis mendeteksi localhost / parent folder)
possible_paths_kontrak = ["database_kontrak.json", "database.json", os.path.join("..", "database_kontrak.json")]
json_target_path = next((p for p in possible_paths_kontrak if os.path.exists(p)), "database_kontrak.json")
abs_json_path = os.path.abspath(json_target_path)

possible_paths_va = ["database_va_npp.json", os.path.join("..", "database_va_npp.json")]
va_json_target_path = next((p for p in possible_paths_va if os.path.exists(p)), "database_va_npp.json")
abs_va_json_path = os.path.abspath(va_json_target_path)

st.title("🔃 Update & Manajemen Database SPK, IO & Mapping VA/NPP")
st.caption(
    "Kelola database ringan berbasis JSON (`database_kontrak.json` & `database_va_npp.json`) "
    "agar pembacaan otomatis di seluruh halaman web berjalan instan tanpa membebani Google Sheets API."
)

tab_kontrak, tab_va_npp = st.tabs([
    "1️⃣ Master Kontrak (SPK & IO)",
    "2️⃣ Mapping VA & NPP -> SPK"
])

# ==============================================================================
# TAB 1: UPLOAD & EKSTRAKSI MASTER KONTRAK KANTOR -> JSON + AUTO-PUSH GITHUB
# ==============================================================================
with tab_kontrak:
    st.subheader("📤 Ekstraksi Master Kontrak Kantor")
    st.info(
        "💡 **Update SPK dan IO dari Monitoring Kontrak**
    )

    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
        hanya_aktif = st.checkbox("✅ Hanya ambil kontrak berstatus AKTIF", value=True)
    with col_opt2:
        gunakan_perubahan_spk = st.checkbox(
            "🔄 Gunakan 'Perubahan No Kontrak/SPK' jika tersedia",
            value=False,
            help="Jika dicentang, nomor kontrak awal (misal LoI / SP2K) otomatis diganti dengan nomor kontrak perubahan resmi jika kolomnya terisi."
        )

    uploaded_spk_excel = st.file_uploader(
        "Unggah File Master Kontrak Kantor (.xlsx / .xls)",
        type=["xlsx", "xls"],
        key="uploader_spk_page"
    )

    if uploaded_spk_excel is not None:
        try:
            df_extracted, unmapped_units = parse_master_kontrak_excel(
                uploaded_spk_excel,
                hanya_aktif=hanya_aktif,
                gunakan_perubahan_spk=gunakan_perubahan_spk
            )

            if unmapped_units:
                st.warning(f"⚠️ Ditemukan nama unit baru yang belum ada di peta default: `{', '.join(unmapped_units)}` (akan dibuatkan grup cabang otomatis).")

            c_m1, c_m2, c_m3 = st.columns(3)
            c_m1.metric("Total Kontrak (SPK) Terbaca", f"{len(df_extracted)} SPK")
            c_m2.metric("Total Internal Order (IO)", f"{df_extracted[df_extracted['Internal Order (IO)'] != '-']['Internal Order (IO)'].nunique()} IO")
            c_m3.metric("Jumlah Cabang / Unit", f"{df_extracted['Cabang JSON'].nunique()} Unit")

            with st.expander("🔍 Preview Hasil Ekstraksi SPK & IO Sebelum Disimpan", expanded=True):
                filter_cab_prev = st.selectbox(
                    "Filter Preview Cabang:",
                    ["Semua Cabang"] + sorted(df_extracted["Cabang JSON"].unique().tolist()),
                    key="prev_extract_cab"
                )
                df_show_prev = df_extracted if filter_cab_prev == "Semua Cabang" else df_extracted[df_extracted["Cabang JSON"] == filter_cab_prev]
                st.dataframe(df_show_prev, use_container_width=True, hide_index=True)

            if st.button("⚡ Simpan Database Kontrak & Push ke GitHub Sekarang", type="primary", key="btn_sync_spk_page"):
                if os.path.exists(json_target_path):
                    with open(json_target_path, "r", encoding="utf-8") as f:
                        old_db = json.load(f)
                else:
                    old_db = {}

                updated_db = {}
                all_cabang_keys = list(dict.fromkeys(list(old_db.keys()) + df_extracted["Cabang JSON"].tolist()))

                total_added = 0
                total_removed = 0
                total_io_mapped = 0
                detail_log = []

                for json_key in all_cabang_keys:
                    meta_old = old_db.get(json_key, {})
                    meta_def = DEFAULT_META_CABANG.get(json_key, {"kota": json_key.replace("Cab. ", ""), "manager": "NAMA MANAGER"})

                    updated_db[json_key] = {
                        "kota": meta_old.get("kota", meta_def["kota"]),
                        "manager": meta_old.get("manager", meta_def["manager"]),
                        "spk": [],
                        "spk_io": {}
                    }

                    sub_df = df_extracted[df_extracted["Cabang JSON"] == json_key]
                    excel_spk_list = []
                    excel_io_map = {}

                    for _, row in sub_df.iterrows():
                        s_val = str(row["Nomor Kontrak (SPK)"]).strip()
                        i_val = str(row["Internal Order (IO)"]).strip()
                        if s_val and s_val not in excel_spk_list:
                            excel_spk_list.append(s_val)
                        if s_val and i_val and i_val != "-":
                            excel_io_map[s_val] = i_val

                    old_spk_list = meta_old.get("spk", [])
                    kept_list = []
                    added_cab = 0
                    removed_cab = 0

                    for old_s in old_spk_list:
                        if old_s in excel_spk_list and old_s not in kept_list:
                            kept_list.append(old_s)
                        elif old_s not in excel_spk_list:
                            total_removed += 1
                            removed_cab += 1

                    for new_s in excel_spk_list:
                        if new_s not in kept_list:
                            kept_list.append(new_s)
                            total_added += 1
                            added_cab += 1

                    updated_db[json_key]["spk"] = kept_list
                    updated_db[json_key]["spk_io"] = {s: excel_io_map[s] for s in kept_list if s in excel_io_map}
                    total_io_mapped += len(updated_db[json_key]["spk_io"])

                    if added_cab > 0 or removed_cab > 0:
                        detail_log.append(
                            f"- **{json_key}**: `+{added_cab}` SPK baru, `-{removed_cab}` SPK dihapus "
                            f"(Total aktif: `{len(kept_list)}` SPK | `{len(updated_db[json_key]['spk_io'])}` IO)"
                        )

                json_str = json.dumps(updated_db, indent=2, ensure_ascii=False)
                with open(json_target_path, "w", encoding="utf-8") as f:
                    f.write(json_str)

                ok_gh, msg_gh = push_file_to_github("database_kontrak.json", json_str, "Auto-update database_kontrak.json (SPK + IO) via Web App")

                st.session_state["spk_sync_msg_page"] = {
                    "added": total_added,
                    "removed": total_removed,
                    "io_count": total_io_mapped,
                    "local_path": abs_json_path,
                    "details": detail_log,
                    "gh_ok": ok_gh,
                    "gh_msg": msg_gh
                }

                st.cache_data.clear()
                st.rerun()

        except Exception as e:
            st.error(f"Gagal memproses file Excel Kontrak: {e}")

    if "spk_sync_msg_page" in st.session_state:
        msg = st.session_state["spk_sync_msg_page"]
        st.success(
            f"✅ **Database Berhasil Diperbarui!** "
            f"(+{msg['added']} SPK baru | -{msg['removed']} SPK lama dihapus | {msg['io_count']} pasangan SPK ↔ IO tersimpan)"
        )
        st.info(f"💾 **Database Offline / Localhost:** Tersimpan otomatis di `{msg.get('local_path', json_target_path)}`")
        if msg.get("gh_ok"):
            st.success(f"☁️ **GitHub Cloud Sync:** {msg.get('gh_msg')}")
        else:
            st.warning(f"⚠️ **GitHub Cloud Sync:** {msg.get('gh_msg')}")

        if msg["details"]:
            st.markdown("".join(msg["details"]))
        else:
            st.info("Seluruh daftar SPK sudah sesuai dengan file Excel terbaru (0 penambahan/pengurangan SPK), mapping IO tetap diperbarui.")

        if st.button("Tutup Notifikasi", key="close_spk_notif_page"):
            del st.session_state["spk_sync_msg_page"]
            st.rerun()

    st.markdown("---")
    st.subheader("📋 Preview Database SPK & Internal Order (IO) Aktif per Cabang")

    if os.path.exists(json_target_path):
        with open(json_target_path, "r", encoding="utf-8") as f:
            db_view = json.load(f)

        col_sel, col_dl = st.columns([2, 1])
        with col_sel:
            cabang_preview = st.selectbox("Pilih Cabang untuk Melihat Daftar SPK & IO Aktif:", list(db_view.keys()))
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
        spk_list_aktif = info_cab.get("spk", [])
        spk_io_aktif = info_cab.get("spk_io", {})

        st.markdown(
            f"**Kota:** `{info_cab.get('kota', '-')}` &nbsp;|&nbsp; "
            f"**Manager:** `{info_cab.get('manager', '-')}` &nbsp;|&nbsp; "
            f"**Total SPK Aktif:** `{len(spk_list_aktif)}` &nbsp;|&nbsp; "
            f"**IO Terpetakan:** `{len(spk_io_aktif)}`"
        )

        if spk_list_aktif:
            rows_html = ""
            for idx, spk in enumerate(spk_list_aktif, start=1):
                io_val = spk_io_aktif.get(spk, "-")
                rows_html += (
                    f'<tr style="border-bottom: 1px solid #E6EAF1;">'
                    f'<td style="width: 60px; text-align: center; padding: 10px 8px; color: #31333F; font-weight: 500;">{idx}</td>'
                    f'<td style="width: 170px; text-align: center; padding: 10px 12px; color: #0A2540; font-family: monospace; font-weight: 600;">{io_val}</td>'
                    f'<td style="padding: 10px 14px; color: #0A2540; font-family: monospace; font-size: 14px;">{spk}</td>'
                    f'</tr>'
                )

            table_html = (
                '<div style="width: 100%; overflow-x: auto; border: 1px solid #D6DCE5; border-radius: 8px; margin-top: 8px;">'
                '<table style="width: 100%; border-collapse: collapse; background-color: #FFFFFF; font-size: 14px;">'
                '<thead>'
                '<tr style="background-color: #F0F4F9; border-bottom: 2px solid #FF7A00; text-align: left;">'
                '<th style="width: 60px; text-align: center; padding: 12px 8px; color: #0A2540; font-weight: 700;">No</th>'
                '<th style="width: 170px; text-align: center; padding: 12px 12px; color: #0A2540; font-weight: 700;">Internal Order (IO)</th>'
                '<th style="padding: 12px 14px; text-align: left; color: #0A2540; font-weight: 700;">Nomor Kontrak (SPK)</th>'
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

# ==============================================================================
# TAB 2: UPLOAD & EKSTRAKSI MAPPING VA MANDIRI & NPP -> database_va_npp.json
# ==============================================================================
with tab_va_npp:
    st.subheader("🔗 Database Mapping VA & NPP")
    st.info(
        "💡 **Fitur Otomatisasi SPK dari PDF Billing / NPP**"
    )

    uploaded_va_excel = st.file_uploader(
        "Unggah File Mapping NPP & VA",
        type=["xlsx", "xls"],
        key="uploader_va_npp_page"
    )

    if uploaded_va_excel is not None:
        try:
            parsed_va_data = parse_mapping_va_npp_excel(uploaded_va_excel)
            kes_dict = parsed_va_data.get("kesehatan", {})
            tk_dict = parsed_va_data.get("ketenagakerjaan", {})

            kes_tunggal = sum(1 for v in kes_dict.values() if len(v["spk_list"]) == 1)
            kes_gabung = sum(1 for v in kes_dict.values() if len(v["spk_list"]) > 1)
            kes_kosong = sum(1 for v in kes_dict.values() if len(v["spk_list"]) == 0)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total VA Kesehatan", f"{len(kes_dict)} VA")
            c2.metric("VA Tunggal (1 SPK)", f"{kes_tunggal} VA")
            c3.metric("VA Gabung (>1 SPK)", f"{kes_gabung} VA")
            c4.metric("Total NPP Ketenagakerjaan", f"{len(tk_dict)} NPP")

            # Buat tabel preview gabungan
            prev_rows = []
            for jenis_lbl, d_src in [("BPJS Kesehatan (VA)", kes_dict), ("BPJS Ketenagakerjaan (NPP)", tk_dict)]:
                for kode, info in d_src.items():
                    spks = info["spk_list"]
                    status_map = "Kosong" if len(spks) == 0 else ("Tunggal (Auto)" if len(spks) == 1 else f"Gabung ({len(spks)} SPK)")
                    prev_rows.append({
                        "Program": jenis_lbl,
                        "Cabang": info["cabang"],
                        "VA Mandiri / NPP": kode,
                        "Status": status_map,
                        "Daftar SPK": " | ".join(spks) if spks else "-"
                    })
            df_va_prev = pd.DataFrame(prev_rows)

            with st.expander("🔍 Preview Hasil Mapping VA & NPP Sebelum Disimpan", expanded=True):
                f_prog = st.selectbox("Filter Program:", ["Semua", "BPJS Kesehatan (VA)", "BPJS Ketenagakerjaan (NPP)"], key="f_prog_va")
                df_show_va = df_va_prev if f_prog == "Semua" else df_va_prev[df_va_prev["Program"] == f_prog]
                st.dataframe(df_show_va, use_container_width=True, hide_index=True)

            if st.button("⚡ Simpan database_va_npp.json & Push ke GitHub Sekarang", type="primary", key="btn_sync_va_page"):
                json_va_str = json.dumps(parsed_va_data, indent=2, ensure_ascii=False)
                with open(va_json_target_path, "w", encoding="utf-8") as f:
                    f.write(json_va_str)

                ok_gh_va, msg_gh_va = push_file_to_github(
                    "database_va_npp.json",
                    json_va_str,
                    "Auto-update database_va_npp.json (Mapping VA & NPP) via Web App"
                )

                st.session_state["va_sync_msg_page"] = {
                    "kes_total": len(kes_dict),
                    "kes_tunggal": kes_tunggal,
                    "kes_gabung": kes_gabung,
                    "kes_kosong": kes_kosong,
                    "tk_total": len(tk_dict),
                    "local_path": abs_va_json_path,
                    "gh_ok": ok_gh_va,
                    "gh_msg": msg_gh_va
                }
                st.cache_data.clear()
                st.rerun()

        except Exception as e:
            st.error(f"Gagal memproses file Mapping VA & NPP: {e}")

    if "va_sync_msg_page" in st.session_state:
        vmsg = st.session_state["va_sync_msg_page"]
        st.success(
            f"✅ **Database Mapping VA & NPP Berhasil Diperbarui!** "
            f"({vmsg['kes_total']} VA Kesehatan: {vmsg['kes_tunggal']} Tunggal, {vmsg['kes_gabung']} Gabung, {vmsg['kes_kosong']} Kosong | "
            f"{vmsg['tk_total']} NPP Ketenagakerjaan)"
        )
        st.info(f"💾 **Database Offline / Localhost:** Tersimpan otomatis di `{vmsg.get('local_path', va_json_target_path)}`")
        if vmsg.get("gh_ok"):
            st.success(f"☁️ **GitHub Cloud Sync:** {vmsg.get('gh_msg')}")
        else:
            st.warning(f"⚠️ **GitHub Cloud Sync:** {vmsg.get('gh_msg')}")

        if st.button("Tutup Notifikasi", key="close_va_notif_page"):
            del st.session_state["va_sync_msg_page"]
            st.rerun()

    st.markdown("---")
    st.subheader("📋 Preview Database Mapping VA & NPP Aktif")

    if os.path.exists(va_json_target_path):
        with open(va_json_target_path, "r", encoding="utf-8") as f:
            db_va_view = json.load(f)

        kes_v = db_va_view.get("kesehatan", {})
        tk_v = db_va_view.get("ketenagakerjaan", {})

        col_v1, col_v2 = st.columns([2, 1])
        with col_v1:
            st.markdown(f"**Status File Aktif:** `{len(kes_v)}` VA Kesehatan &nbsp;|&nbsp; `{len(tk_v)}` NPP Ketenagakerjaan")
        with col_v2:
            st.download_button(
                label="📥 Download Backup database_va_npp.json",
                data=json.dumps(db_va_view, indent=2, ensure_ascii=False),
                file_name="database_va_npp.json",
                mime="application/json",
                use_container_width=True
            )

        active_rows = []
        for kode, info in kes_v.items():
            spks = info.get("spk_list", [])
            active_rows.append({
                "Jenis": "BPJS Kesehatan",
                "Cabang": info.get("cabang", "-"),
                "VA Mandiri / NPP": kode,
                "Tipe": "Kosong" if not spks else ("Tunggal" if len(spks) == 1 else f"VA Gabung ({len(spks)} SPK)"),
                "Daftar SPK": " | ".join(spks) if spks else "-"
            })
        for kode, info in tk_v.items():
            spks = info.get("spk_list", [])
            active_rows.append({
                "Jenis": "BPJS Ketenagakerjaan",
                "Cabang": info.get("cabang", "-"),
                "VA Mandiri / NPP": kode,
                "Tipe": "Kosong" if not spks else ("Tunggal" if len(spks) == 1 else f"NPP Gabung ({len(spks)} SPK)"),
                "Daftar SPK": " | ".join(spks) if spks else "-"
            })

        if active_rows:
            df_active_va = pd.DataFrame(active_rows)
            f_cab_act = st.selectbox(
                "Filter Cabang Mapping Aktif:",
                ["Semua Cabang"] + sorted(df_active_va["Cabang"].unique().tolist()),
                key="filter_cab_va_active"
            )
            df_disp_act = df_active_va if f_cab_act == "Semua Cabang" else df_active_va[df_active_va["Cabang"] == f_cab_act]
            st.dataframe(df_disp_act, use_container_width=True, hide_index=True)
    else:
        st.warning("File `database_va_npp.json` belum ditemukan. Silakan unggah file `Mapping NPP dan VA ALL.xlsx` di atas.")
