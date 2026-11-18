import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter

from modul_bni import ekstrak_pdf_bni, proses_rekonsiliasi_bni
from modul_mandiri import ekstrak_pdf_mandiri, proses_rekonsiliasi_mandiri

# ============================================================
# KONFIGURASI HALAMAN WEB
# ============================================================
st.set_page_config(
    page_title="Rekonsiliasi Rekening Koran vs General Ledger",
    page_icon="🏦",
    layout="wide"
)

# ============================================================
# HALAMAN LOGIN PASSWORD (SELARAS TEMA SAGE GREEN REKON PRO)
# ============================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"], [data-testid="collapsedControl"], header[data-testid="stHeader"] {
            display: none !important;
        }
        .stApp {
            background-color: #F4F6F5;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background-color: #FFFFFF;
            border: 1px solid #DCE3DF;
            border-top: 5px solid #5B7B65;
            border-radius: 14px;
            box-shadow: 0 10px 30px rgba(45, 65, 53, 0.08);
            padding: 2rem 2.2rem 1.8rem 2.2rem;
            margin-top: 4vh;
        }
        div[data-testid="stTextInput"] input {
            background-color: #F8FAF9;
            border: 1px solid #CBD5D0;
            border-radius: 8px;
            padding: 0.6rem 0.85rem;
            color: #1F2933;
        }
        div[data-testid="stTextInput"] input:focus {
            border-color: #5B7B65;
            box-shadow: 0 0 0 2px rgba(91, 123, 101, 0.2);
        }
        div[data-testid="stButton"] > button[kind="primary"] {
            background-color: #5B7B65 !important;
            border: none !important;
            color: #FFFFFF !important;
            font-weight: 600;
            border-radius: 8px;
            padding: 0.6rem 1rem;
            transition: all 0.2s ease;
        }
        div[data-testid="stButton"] > button[kind="primary"]:hover {
            background-color: #496552 !important;
            box-shadow: 0 4px 12px rgba(91, 123, 101, 0.28);
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    _, col_center, _ = st.columns([1.2, 1.3, 1.2])
    with col_center:
        with st.container(border=True):
            st.markdown(
                """
                <div style="text-align: center; margin-bottom: 1.2rem;">
                    <div style="display: inline-block; background-color: #EBF1ED; color: #496552;
                                font-size: 0.72rem; font-weight: 700; letter-spacing: 0.8px;
                                padding: 4px 12px; border-radius: 20px; margin-bottom: 10px;">
                        REKON PRO • ENTERPRISE WORKSPACE
                    </div>
                    <div style="font-size: 2.2rem; line-height: 1.1; margin-bottom: 6px;">🏦</div>
                    <h2 style="margin: 0; font-size: 1.45rem; color: #1E2923; font-weight: 700;">
                        Rekonsiliasi Bank vs GL
                        <sup style="font-size: 0.55em; font-weight: 600; color: #6B7B72;">by FERDX</sup>
                    </h2>
                    <p style="margin: 6px 0 0 0; font-size: 0.85rem; color: #6B7B72;">
                        Automated Bank Statement & General Ledger Reconciliation
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )

            pwd = st.text_input("Password Akses", type="password", placeholder="Masukkan kata sandi tim...", autocomplete="current-password")
            st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)

            if st.button("🔐 Masuk ke Aplikasi", type="primary", use_container_width=True):
                if pwd == st.secrets["APP_PASSWORD"]:
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Password salah! Silakan coba lagi.")

            st.markdown(
                """
                <div style="text-align: center; margin-top: 1.2rem; padding-top: 0.8rem;
                            border-top: 1px solid #EEF2F0; font-size: 0.75rem; color: #8A9990;">
                    🔒 Akses Terbatas • Pemrosesan Aman di Memori
                </div>
                """,
                unsafe_allow_html=True
            )
    st.stop()
 
# ============================================================
# FUNGSI PEMBANTU & PENERJEMAH KOLOM UNIVERSAL
# ============================================================
def deteksi_kolom_otomatis(columns):
    kamus = {
        "Date": ["date", "tanggal", "tgl", "posting date", "trans date", "transaction date", "tgl transaksi"],
        "Label": ["label", "keterangan", "deskripsi", "description", "uraian", "memo", "remarks", "catatan"],
        "Debit": ["debit", "debet", "masuk", "penerimaan", "db", "mutasi debit", "mutasi debet"],
        "Credit": ["credit", "kredit", "keluar", "pengeluaran", "cr", "mutasi credit", "mutasi kredit"],
        "Number": ["number", "no", "nomor", "no bukti", "no. bukti", "ref", "reference", "no transaksi", "voucher"],
        "Partner": ["partner", "rekanan", "customer", "vendor", "pelanggan", "supplier", "nama", "pihak"]
    }
    hasil = {}
    cols_lower = {str(c).strip().lower(): c for c in columns}
    for target, sinonim in kamus.items():
        hasil[target] = None
        for s in sinonim:
            if s in cols_lower:
                hasil[target] = cols_lower[s]
                break
    return hasil


def normalisasi_excel_pembukuan(df_raw, mapping):
    df = df_raw.copy()
    rename_dict = {}
    for target in ["Date", "Label", "Debit", "Credit", "Number", "Partner"]:
        sumber = mapping.get(target)
        if sumber and sumber != "(Tidak Ada)" and sumber in df.columns:
            rename_dict[sumber] = target

    df = df.rename(columns=rename_dict)
    if "Number" not in df.columns:
        df["Number"] = ""
    if "Partner" not in df.columns:
        df["Partner"] = ""
    return df


def perhalus_label_universal(df_export, df_laporan_final, pesan_audit):
    ganti_teks = {
        "HANYA DI ODOO": "HANYA DI GENERAL LEDGER",
        "Transaksi Odoo": "Transaksi General Ledger",
        "Trx Odoo": "Trx GL",
        "Tgl Odoo": "Tgl GL",
        "di Odoo": "di General Ledger",
        "SUM Hasil Rekon Odoo": "SUM Hasil Rekon General Ledger",
        "Total Bawaan File Odoo": "Total Bawaan File General Ledger",
        "TOTAL ODOO": "TOTAL GENERAL LEDGER",
        "& Odoo": "& General Ledger",
        "Total Odoo": "Total General Ledger",
        "| Odoo:": "| GL:",
        "STATUS AUDIT": "CONTROL CHECK"
    }

    for lama, baru in ganti_teks.items():
        df_export["Tanggal_Clean"] = df_export["Tanggal_Clean"].astype(str).str.replace(lama, baru, regex=False)
        df_export["Status_Rekonsiliasi"] = df_export["Status_Rekonsiliasi"].astype(str).str.replace(lama, baru, regex=False)
        if "Label" in df_export.columns:
            df_export["Label"] = df_export["Label"].fillna("").astype(str).str.replace(lama, baru, regex=False)

        df_laporan_final["Status_Rekonsiliasi"] = df_laporan_final["Status_Rekonsiliasi"].astype(str).str.replace(lama, baru, regex=False)
        if "Label" in df_laporan_final.columns:
            df_laporan_final["Label"] = df_laporan_final["Label"].fillna("").astype(str).str.replace(lama, baru, regex=False)

        pesan_audit = pesan_audit.replace(lama, baru)

    return df_export, df_laporan_final, pesan_audit


def tambah_kolom_selisih(df_export, df_laporan_final):
    nama_kolom = "Selisih Bank dengan GL"
    for df in [df_laporan_final, df_export]:
        kredit_bank = pd.to_numeric(df.get("Kredit", 0), errors="coerce").fillna(0)
        debet_bank = pd.to_numeric(df.get("Debet", 0), errors="coerce").fillna(0)
        debit_gl = pd.to_numeric(df.get("Debit", 0), errors="coerce").fillna(0)
        credit_gl = pd.to_numeric(df.get("Credit", 0), errors="coerce").fillna(0)
        df[nama_kolom] = ((kredit_bank - debet_bank) - (debit_gl - credit_gl)).round(2)

    mask_transaksi = ~df_export["Tanggal_Clean"].isin(["TOTAL REKON", "TOTAL ASLI", "CONTROL CHECK"])
    total_selisih_sum = round(df_export.loc[mask_transaksi, nama_kolom].sum(), 2)
    df_export.loc[df_export["Tanggal_Clean"] == "TOTAL REKON", nama_kolom] = total_selisih_sum
    return df_export, df_laporan_final
def cek_kelengkapan_bulan(df_bank, df_gl):
    tgl_bank = pd.to_datetime(df_bank["Tanggal"], errors="coerce").dropna()
    tgl_gl = pd.to_datetime(df_gl["Date"], errors="coerce").dropna()

    bulan_bank = sorted(tgl_bank.dt.strftime("%Y-%m").unique().tolist())
    bulan_gl = sorted(tgl_gl.dt.strftime("%Y-%m").unique().tolist())

    hilang_di_gl = sorted(list(set(bulan_bank) - set(bulan_gl)))
    hilang_di_bank = sorted(list(set(bulan_gl) - set(bulan_bank)))
    return bulan_bank, bulan_gl, hilang_di_gl, hilang_di_bank


def buat_rekap_per_bulan(df_laporan_final):
    df_temp = df_laporan_final.copy()
    df_temp["Bulan"] = df_temp["Tanggal_Clean"].astype(str).str[:7]

    rekap_rows = []
    for bln, grp in df_temp.groupby("Bulan", sort=True):
        debet_b = round(grp["Debet"].fillna(0).sum(), 2)
        kredit_b = round(grp["Kredit"].fillna(0).sum(), 2)
        debit_g = round(grp["Debit"].fillna(0).sum(), 2)
        credit_g = round(grp["Credit"].fillna(0).sum(), 2)
        selisih_net = round((kredit_b - debet_b) - (debit_g - credit_g), 2)
        jml_unmatched = grp[~grp["Status_Rekonsiliasi"].str.startswith(("MATCH", "NETTING JAGIR", "KONTRA BANK"))].shape

        if selisih_net == 0 and jml_unmatched == 0:
            status = "✅ COCOK 100%"
        elif selisih_net == 0 and jml_unmatched > 0:
            status = f"⚠️ NETT NOL ({jml_unmatched} Baris Belum Cocok)"
        else:
            status = f"❌ ADA SELISIH ({jml_unmatched} Baris Belum Cocok)"

        rekap_rows.append({
            "Bulan": bln,
            "Debet Bank": debet_b,
            "Kredit Bank": kredit_b,
            "Debit GL": debit_g,
            "Credit GL": credit_g,
            "Selisih Bank dengan GL": selisih_net,
            "Baris Belum Match": jml_unmatched,
            "Status Bulan": status
        })

    return pd.DataFrame(rekap_rows)


def proses_satu_rekening(pilihan_bank, uploaded_pdfs, df_book_raw, mapping_kolom):
    df_book_normalized = normalisasi_excel_pembukuan(df_book_raw, mapping_kolom)

    hasil_ekstraksi_list = []
    for pdf_item in uploaded_pdfs:
        if pilihan_bank == "BNI Direct":
            df_p, s_awal, s_akhir, err_p = ekstrak_pdf_bni(pdf_item)
        else:
            df_p, s_awal, s_akhir, err_p = ekstrak_pdf_mandiri(pdf_item)

        df_p["Sumber_File"] = pdf_item.name
        hasil_ekstraksi_list.append({
            "tgl_awal": pd.to_datetime(df_p["Tanggal"]).min(),
            "df": df_p,
            "saldo_awal": s_awal,
            "saldo_akhir": s_akhir,
            "errors": err_p
        })

    hasil_ekstraksi_list.sort(key=lambda x: x["tgl_awal"])
    df_bank = pd.concat([x["df"] for x in hasil_ekstraksi_list], ignore_index=True)
    saldo_awal = hasil_ekstraksi_list["saldo_awal"]
    saldo_akhir = hasil_ekstraksi_list[-1]["saldo_akhir"]
    pdf_errors = [err for x in hasil_ekstraksi_list for err in x["errors"]]

    bulan_bank, bulan_gl, hilang_di_gl, hilang_di_bank = cek_kelengkapan_bulan(df_bank, df_book_normalized)

    if pilihan_bank == "BNI Direct":
        df_export, df_laporan_final, pesan_audit = proses_rekonsiliasi_bni(df_bank.copy(), df_book_normalized)
    else:
        df_export, df_laporan_final, pesan_audit = proses_rekonsiliasi_mandiri(df_bank.copy(), df_book_normalized)

    df_export, df_laporan_final, pesan_audit = perhalus_label_universal(df_export, df_laporan_final, pesan_audit)
    df_export, df_laporan_final = tambah_kolom_selisih(df_export, df_laporan_final)
    df_rekap_bulan = buat_rekap_per_bulan(df_laporan_final)

    return {
        "df_bank": df_bank,
        "df_export": df_export,
        "df_laporan_final": df_laporan_final,
        "df_rekap_bulan": df_rekap_bulan,
        "saldo_awal": saldo_awal,
        "saldo_akhir": saldo_akhir,
        "pdf_errors": pdf_errors,
        "pesan_audit": pesan_audit,
        "bulan_bank": bulan_bank,
        "hilang_di_gl": hilang_di_gl,
        "hilang_di_bank": hilang_di_bank
    }


def tulis_sheet_rekon_dengan_rumus(writer, sheet_name, df_export, df_laporan_final):
    df_export.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]
    headers = [cell.value for cell in ws]

    last_data_row = len(df_laporan_final) + 1
    row_total_rekon = last_data_row + 1
    row_total_asli = last_data_row + 2
    row_control_check = last_data_row + 3

    col_idx = {name: i + 1 for i, name in enumerate(headers)}
    col_letter = {name: get_column_letter(i + 1) for i, name in enumerate(headers)}

    if all(k in col_letter for k in ["Kredit", "Debet", "Debit", "Credit", "Selisih Bank dengan GL"]):
        c_kredit, c_debet = col_letter["Kredit"], col_letter["Debet"]
        c_debit, c_credit = col_letter["Debit"], col_letter["Credit"]
        idx_selisih = col_idx["Selisih Bank dengan GL"]
        for r in range(2, last_data_row + 1):
            ws.cell(row=r, column=idx_selisih).value = f"=({c_kredit}{r}-{c_debet}{r})-({c_debit}{r}-{c_credit}{r})"

    for nama_kol in ["Debet", "Kredit", "Debit", "Credit", "Selisih Bank dengan GL"]:
        if nama_kol in col_idx:
            c_idx, c_let = col_idx[nama_kol], col_letter[nama_kol]
            ws.cell(row=row_total_rekon, column=c_idx).value = f"=SUM({c_let}2:{c_let}{last_data_row})"
            ws.cell(row=row_control_check, column=c_idx).value = f"={c_let}{row_total_rekon}-{c_let}{row_total_asli}"
# ============================================================
# KOSMETIK HALAMAN UTAMA (TOP BAR HIJAU + STEPPER + CARD)
# ============================================================
st.markdown(
    """
    <style>
    header[data-testid="stHeader"] {
        background-color: #466650 !important;
        height: 2.6rem !important;
        border-bottom: 1px solid #395441;
    }
    header[data-testid="stHeader"] * {
        color: #FFFFFF !important;
    }
    .stApp {
        background-color: #F2F5F3;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #FFFFFF;
        border: 1px solid #DCE3DF !important;
        border-radius: 12px !important;
        box-shadow: 0 3px 12px rgba(45, 65, 53, 0.05);
        padding: 1rem 1.25rem;
    }
    div[data-testid="stColumn"]:nth-of-type(1) [data-testid="stFileUploaderDropzone"] {
        background-color: #EBF3EE !important;
        border: 1.5px dashed #7EA189 !important;
        border-radius: 10px !important;
        padding: 1.4rem 1rem !important;
    }
    div[data-testid="stColumn"]:nth-of-type(2) [data-testid="stFileUploaderDropzone"] {
        background-color: #F4F6F5 !important;
        border: 1.5px dashed #B8C4BD !important;
        border-radius: 10px !important;
        padding: 1.4rem 1rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <h1 style='margin-bottom: 0.2rem;'>
        🏦 Aplikasi Rekonsiliasi Rekening Koran Bank vs General Ledger
        <sup style='font-size: 0.42em; font-weight: 600; color: #6c757d; letter-spacing: 0.5px;'>by FERDX</sup>
    </h1>
    """,
    unsafe_allow_html=True
)

st.markdown("Unggah **PDF Rekening Koran** dan **Excel Catatan Transaksi Internal (ERP / Accounting Software / Excel)** untuk validasi otomatis.")

# ============================================================
# PILIHAN MODE: SINGLE VS MULTI-REKENING (BATCH)
# ============================================================
col_mode1, col_mode2 = st.columns([3, 1])
with col_mode1:
    mode_rekon = st.radio(
        "Mode Rekonsiliasi:",
        ["Single Rekening (1 Akun Bank)", "Multi-Rekening Batch (2–5 Akun Sekaligus)"],
        horizontal=True
    )
with col_mode2:
    if mode_rekon.startswith("Multi-Rekening"):
        jumlah_akun = st.number_input("Jumlah Rekening:", min_value=2, max_value=5, value=2, step=1)
    else:
        jumlah_akun = 1
daftar_input_akun = []

for idx in range(int(jumlah_akun)):
    judul_slot = f"📂 Rekening #{idx + 1}" if jumlah_akun > 1 else "📂 Unggah Dokumen Rekonsiliasi"
    with st.expander(judul_slot, expanded=True):
        c_bank, c_nama = st.columns([1.5, 1.5])
        with c_bank:
            bank_slot = st.radio(
                "Pilih Bank:",
                ["BNI Direct", "Bank Mandiri"],
                horizontal=True,
                key=f"bank_slot_{idx}"
            )
        with c_nama:
            default_nama = f"Rekening_{idx + 1}_{bank_slot.split()[0]}"
            label_akun = st.text_input(
                "Label / Nama Rekening (untuk nama sheet Excel):",
                value=default_nama,
                key=f"label_slot_{idx}"
            )

        st.markdown(
            """
            <div style="background-color: #FFFFFF; border: 1px solid #DCE3DF; border-radius: 12px;
                        padding: 10px 24px; margin: 8px 0 14px 0;
                        display: flex; align-items: center; justify-content: space-between;">
                <div style="display: flex; align-items: center; justify-content: center;
                            width: 32px; height: 32px; border-radius: 8px;
                            background-color: #E5EFE8; border: 1px solid #7EA189; font-size: 15px;">
                    📄
                </div>
                <div style="flex-grow: 1; height: 3px; background: linear-gradient(90deg, #5B7B65 50%, #D5DDD8 50%);
                            margin: 0 14px; border-radius: 2px;"></div>
                <div style="display: flex; align-items: center; justify-content: center;
                            width: 32px; height: 32px; border-radius: 8px;
                            background-color: #ECEFED; border: 1px solid #CBD3CE; font-size: 15px;">
                    📊
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        c_pdf, c_gl = st.columns(2)
        with c_pdf:
            pdfs_slot = st.file_uploader(
                f"1. Rekening Koran {bank_slot} (.pdf — Bisa 1–12 Bulan)",
                type=["pdf"],
                accept_multiple_files=True,
                key=f"pdf_slot_{idx}"
            )
        with c_gl:
            gl_slot = st.file_uploader(
                "2. Excel General Ledger (.xlsx)",
                type=["xlsx"],
                key=f"gl_slot_{idx}"
            )

        map_slot = {}
        df_gl_slot = None
        if gl_slot:
            df_gl_slot = pd.read_excel(gl_slot)
            kolom_tersedia = df_gl_slot.columns.tolist()
            auto_map = deteksi_kolom_otomatis(kolom_tersedia)
            wajib_ok = all(auto_map[k] is not None for k in ["Date", "Label", "Debit", "Credit"])

            with st.expander("⚙️ Pemetaan Kolom GL (Otomatis Terdeteksi)", expanded=not wajib_ok):
                m1, m2, m3, m4 = st.columns(4)
                map_slot["Date"] = m1.selectbox("Kolom Tanggal *", kolom_tersedia, index=kolom_tersedia.index(auto_map["Date"]) if auto_map["Date"] in kolom_tersedia else 0, key=f"d_{idx}")
                map_slot["Label"] = m2.selectbox("Kolom Deskripsi *", kolom_tersedia, index=kolom_tersedia.index(auto_map["Label"]) if auto_map["Label"] in kolom_tersedia else 0, key=f"l_{idx}")
                map_slot["Debit"] = m3.selectbox("Kolom Debit (Masuk) *", kolom_tersedia, index=kolom_tersedia.index(auto_map["Debit"]) if auto_map["Debit"] in kolom_tersedia else 0, key=f"db_{idx}")
                map_slot["Credit"] = m4.selectbox("Kolom Credit (Keluar) *", kolom_tersedia, index=kolom_tersedia.index(auto_map["Credit"]) if auto_map["Credit"] in kolom_tersedia else 0, key=f"cr_{idx}")

                opsi_opt = ["(Tidak Ada)"] + kolom_tersedia
                o1, o2 = st.columns(2)
                map_slot["Number"] = o1.selectbox("No. Bukti (Opsional)", opsi_opt, index=opsi_opt.index(auto_map["Number"]) if auto_map["Number"] in opsi_opt else 0, key=f"no_{idx}")
                map_slot["Partner"] = o2.selectbox("Partner (Opsional)", opsi_opt, index=opsi_opt.index(auto_map["Partner"]) if auto_map["Partner"] in opsi_opt else 0, key=f"pt_{idx}")

        if pdfs_slot and gl_slot and df_gl_slot is not None:
            daftar_input_akun.append({
                "label": label_akun.strip()[:22] or f"Rekening_{idx + 1}",
                "bank": bank_slot,
                "pdfs": pdfs_slot,
                "df_gl": df_gl_slot,
                "mapping": map_slot
            })
if len(daftar_input_akun) == int(jumlah_akun):
    teks_btn = f"🚀 Proses Rekonsiliasi ({jumlah_akun} Rekening Sekaligus)" if jumlah_akun > 1 else f"🚀 Proses Rekonsiliasi {daftar_input_akun[0]['bank']} Sekarang"
    if st.button(teks_btn, type="primary", use_container_width=True):
        with st.spinner("Sedang memproses ekstraksi dan rekonsiliasi seluruh rekening..."):
            try:
                hasil_semua_akun = []
                ringkasan_batch_rows = []

                for item in daftar_input_akun:
                    res = proses_satu_rekening(item["bank"], item["pdfs"], item["df_gl"], item["mapping"])
                    res["label"] = item["label"]
                    res["bank"] = item["bank"]
                    hasil_semua_akun.append(res)

                    df_lap = res["df_laporan_final"]
                    jml_unmatch = df_lap[~df_lap["Status_Rekonsiliasi"].str.startswith(("MATCH", "NETTING JAGIR", "KONTRA BANK"))].shape[0]
                    total_selisih = round(df_lap["Selisih Bank dengan GL"].sum(), 2)

                    ringkasan_batch_rows.append({
                        "Nama Rekening": item["label"],
                        "Bank": item["bank"],
                        "Saldo Awal PDF": res["saldo_awal"],
                        "Saldo Akhir PDF": res["saldo_akhir"],
                        "Total Baris": len(df_lap),
                        "Baris Belum Match": jml_unmatch,
                        "Selisih Bersih Bank vs GL": total_selisih,
                        "Status Audit": "OKE 100%" if "OKE 100%" in res["pesan_audit"] else "CEK SELISIH"
                    })

                df_ringkasan_batch = pd.DataFrame(ringkasan_batch_rows)
                st.divider()

                if jumlah_akun > 1:
                    st.subheader("🌐 Ringkasan Rekonsiliasi Multi-Rekening")
                    st.dataframe(df_ringkasan_batch, use_container_width=True)

                tab_labels = [f"🏦 {h['label']} ({h['bank']})" for h in hasil_semua_akun]
                tabs_akun = st.tabs(tab_labels)

                for tab_ui, h in zip(tabs_akun, hasil_semua_akun):
                    with tab_ui:
                        if h["hilang_di_gl"]:
                            st.warning(f"⚠️ **Kelengkapan GL:** PDF memuat bulan `{', '.join(h['bulan_bank'])}`, tetapi GL kosong di bulan: **{', '.join(h['hilang_di_gl'])}**.")
                        if h["hilang_di_bank"]:
                            st.warning(f"⚠️ **Kelengkapan PDF:** GL memiliki data di bulan **{', '.join(h['hilang_di_bank'])}** yang belum diunggah PDF-nya.")

                        if "OKE 100%" in h["pesan_audit"]:
                            st.success(f"**Status Audit ({h['label']}):** {h['pesan_audit']}")
                        else:
                            st.error(f"**Status Audit ({h['label']}):** {h['pesan_audit']}")

                        df_lap = h["df_laporan_final"]
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("Saldo Akhir PDF", f"Rp {h['saldo_akhir']:,.0f}")
                        m2.metric("Cocok (MATCH)", df_lap["Status_Rekonsiliasi"].str.startswith("MATCH").sum())
                        m3.metric("Hanya di Bank", df_lap["Status_Rekonsiliasi"].str.startswith("HANYA DI BANK").sum())
                        m4.metric("Hanya di GL", df_lap["Status_Rekonsiliasi"].str.startswith("HANYA DI GENERAL LEDGER").sum())

                        sub_rekap, sub_detail, sub_pdf = st.tabs(["📅 Rekap Per Bulan", "📊 Detail Rekonsiliasi", "📄 Ekstraksi PDF"])
                        with sub_rekap:
                            st.dataframe(h["df_rekap_bulan"], use_container_width=True)
                        with sub_detail:
                            f1, f2 = st.columns(2)
                            with f1:
                                opsi_bln = sorted(df_lap['Tanggal_Clean'].astype(str).str[:7].unique().tolist())
                                sel_bln = st.multiselect("Filter Bulan:", opsi_bln, default=opsi_bln, key=f"fbln_{h['label']}")
                            with f2:
                                opsi_sts = df_lap['Status_Rekonsiliasi'].unique().tolist()
                                sel_sts = st.multiselect("Filter Status:", opsi_sts, default=opsi_sts, key=f"fsts_{h['label']}")
                            df_filtered = df_lap[
                                df_lap['Tanggal_Clean'].astype(str).str[:7].isin(sel_bln)
                                & df_lap['Status_Rekonsiliasi'].isin(sel_sts)
                            ]
                            st.dataframe(df_filtered, use_container_width=True, height=420)
                        with sub_pdf:
                            st.dataframe(h["df_bank"], use_container_width=True, height=420)

                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                    if jumlah_akun > 1:
                        df_ringkasan_batch.to_excel(writer, sheet_name="Ringkasan_Multi_Rekening", index=False)

                    for h in hasil_semua_akun:
                        nama_sheet = h["label"][:25] if jumlah_akun > 1 else "Hasil Rekonsiliasi"
                        tulis_sheet_rekon_dengan_rumus(writer, nama_sheet, h["df_export"], h["df_laporan_final"])
                        if jumlah_akun == 1:
                            h["df_rekap_bulan"].to_excel(writer, sheet_name="Rekap Per Bulan", index=False)
                            h["df_bank"].to_excel(writer, sheet_name=f"Ekstraksi PDF {h['bank'][:10]}", index=False)

                nama_file_out = f"Hasil_Rekon_Multi_{jumlah_akun}_Rekening.xlsx" if jumlah_akun > 1 else f"Hasil_Rekonsiliasi_{hasil_semua_akun[0]['label']}.xlsx"
                st.download_button(
                    label="📥 Download Laporan Rekonsiliasi Lengkap (.xlsx)",
                    data=output_buffer.getvalue(),
                    file_name=nama_file_out,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses file: {e}")
