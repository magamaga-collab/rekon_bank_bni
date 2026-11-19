import streamlit as st
import pandas as pd
import io
import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font

from modul_bni import ekstrak_pdf_bni, proses_rekonsiliasi_bni
from modul_mandiri import ekstrak_pdf_mandiri, proses_rekonsiliasi_mandiri

st.set_page_config(
    page_title="Rekonsiliasi Rekening Koran vs General Ledger",
    page_icon="🏦",
    layout="wide"
)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"], [data-testid="collapsedControl"], header[data-testid="stHeader"] {
            display: none !important;
        }
        .stApp { background-color: #F4F6F5; }
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

st.markdown(
    """
    <style>
    header[data-testid="stHeader"] {
        background-color: #466650 !important;
        height: 2.6rem !important;
        border-bottom: 1px solid #395441;
    }
    header[data-testid="stHeader"] * { color: #FFFFFF !important; }
    .stApp { background-color: #F2F5F3; }
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
        padding: 1.2rem 1rem !important;
    }
    div[data-testid="stColumn"]:nth-of-type(2) [data-testid="stFileUploaderDropzone"] {
        background-color: #F4F6F5 !important;
        border: 1.5px dashed #B8C4BD !important;
        border-radius: 10px !important;
        padding: 1.2rem 1rem !important;
    }
    div[data-testid="stButton"] > button[kind="primary"],
    div[data-testid="stDownloadButton"] > button[kind="primary"] {
        background-color: #5B7B65 !important;
        border: none !important;
        color: #FFFFFF !important;
        font-weight: 600;
        border-radius: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True
)
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
    df_export.loc[df_export["Tanggal_Clean"] == "TOTAL REKON", nama_kolom] = round(df_export.loc[mask_transaksi, nama_kolom].sum(), 2)
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
        jml_unmatched = grp[~grp["Status_Rekonsiliasi"].str.startswith(("MATCH", "NETTING JAGIR", "KONTRA BANK"))].shape[0]
        if selisih_net == 0 and jml_unmatched == 0:
            status = "✅ COCOK 100%"
        elif selisih_net == 0 and jml_unmatched > 0:
            status = f"⚠️ NETT NOL ({jml_unmatched} Baris Belum Cocok)"
        else:
            status = f"❌ ADA SELISIH ({jml_unmatched} Baris Belum Cocok)"
        rekap_rows.append({
            "Bulan": bln, "Debet Bank": debet_b, "Kredit Bank": kredit_b,
            "Debit GL": debit_g, "Credit GL": credit_g,
            "Selisih Bank dengan GL": selisih_net, "Baris Belum Match": jml_unmatched, "Status Bulan": status
        })
    return pd.DataFrame(rekap_rows)
def tulis_sheet_hasil_dengan_rumus(writer, sheet_name, df_export, df_laporan_final):
    df_export.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]
    headers = [cell.value for cell in ws[1]]
    col_idx = {name: i + 1 for i, name in enumerate(headers)}
    col_letter = {name: get_column_letter(i + 1) for i, name in enumerate(headers)}

    fill_subtotal = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
    fill_grand = PatternFill(start_color="D8E8DD", end_color="D8E8DD", fill_type="solid")
    font_bold = Font(bold=True)

    kolom_hitung = [k for k in ["Debet", "Kredit", "Debit", "Credit", "Selisih Bank dengan GL"] if k in col_idx]
    has_selisih = all(k in col_letter for k in ["Kredit", "Debet", "Debit", "Credit", "Selisih Bank dengan GL"])

    start_trx_row = 2
    subtotal_rekon_rows = []
    subtotal_asli_rows = []

    for r in range(2, ws.max_row + 1):
        tag = str(ws.cell(row=r, column=1).value or "")

        if tag.startswith(("SUBTOTAL REKON", "TOTAL REKON")):
            for k in kolom_hitung:
                c_i, c_l = col_idx[k], col_letter[k]
                ws.cell(row=r, column=c_i).value = f"=SUM({c_l}{start_trx_row}:{c_l}{r-1})"
            subtotal_rekon_rows.append(r)
        elif tag.startswith(("SUBTOTAL ASLI", "TOTAL ASLI")):
            subtotal_asli_rows.append(r)
        elif tag.startswith("CONTROL CHECK"):
            for k in kolom_hitung:
                c_i, c_l = col_idx[k], col_letter[k]
                ws.cell(row=r, column=c_i).value = f"={c_l}{r-2}-{c_l}{r-1}"
            start_trx_row = r + 1
        elif tag == "GRAND TOTAL REKON":
            for k in kolom_hitung:
                c_i, c_l = col_idx[k], col_letter[k]
                ws.cell(row=r, column=c_i).value = f"=SUM({','.join(f'{c_l}{sr}' for sr in subtotal_rekon_rows)})"
        elif tag == "GRAND TOTAL ASLI":
            for k in kolom_hitung:
                c_i, c_l = col_idx[k], col_letter[k]
                ws.cell(row=r, column=c_i).value = f"=SUM({','.join(f'{c_l}{sa}' for sa in subtotal_asli_rows)})"
        elif tag == "GRAND CONTROL CHECK":
            for k in kolom_hitung:
                c_i, c_l = col_idx[k], col_letter[k]
                ws.cell(row=r, column=c_i).value = f"={c_l}{r-2}-{c_l}{r-1}"
        else:
            if has_selisih:
                c_kr, c_db = col_letter["Kredit"], col_letter["Debet"]
                c_dg, c_cg = col_letter["Debit"], col_letter["Credit"]
                ws.cell(row=r, column=col_idx["Selisih Bank dengan GL"]).value = f"=({c_kr}{r}-{c_db}{r})-({c_dg}{r}-{c_cg}{r})"

        if tag.startswith(("SUBTOTAL REKON", "SUBTOTAL ASLI", "CONTROL CHECK [")):
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = fill_subtotal
                ws.cell(row=r, column=c).font = font_bold
        elif tag.startswith(("GRAND ", "TOTAL REKON", "TOTAL ASLI", "CONTROL CHECK")):
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = fill_grand
                ws.cell(row=r, column=c).font = font_bold
def jalankan_rekon_satu_rekening(nama_rek, bank_pilihan, pdf_files, df_gl_raw, mapping_kolom):
    df_book_normalized = normalisasi_excel_pembukuan(df_gl_raw, mapping_kolom)
    hasil_ekstraksi_list = []
    for pdf_item in pdf_files:
        if bank_pilihan == "BNI Direct":
            df_p, s_awal, s_akhir, err_p = ekstrak_pdf_bni(pdf_item)
        else:
            df_p, s_awal, s_akhir, err_p = ekstrak_pdf_mandiri(pdf_item)
        df_p["Sumber_File"] = pdf_item.name
        hasil_ekstraksi_list.append({
            "tgl_awal": pd.to_datetime(df_p["Tanggal"]).min(),
            "df": df_p, "saldo_awal": s_awal, "saldo_akhir": s_akhir, "errors": err_p
        })

    hasil_ekstraksi_list.sort(key=lambda x: x["tgl_awal"])
    df_bank = pd.concat([x["df"] for x in hasil_ekstraksi_list], ignore_index=True)
    saldo_awal = hasil_ekstraksi_list[0]["saldo_awal"]
    saldo_akhir = hasil_ekstraksi_list[-1]["saldo_akhir"]
    pdf_errors = [err for x in hasil_ekstraksi_list for err in x["errors"]]

    bulan_bank, bulan_gl, hilang_di_gl, hilang_di_bank = cek_kelengkapan_bulan(df_bank, df_book_normalized)

    if bank_pilihan == "BNI Direct":
        df_export, df_laporan_final, pesan_audit = proses_rekonsiliasi_bni(df_bank.copy(), df_book_normalized)
    else:
        df_export, df_laporan_final, pesan_audit = proses_rekonsiliasi_mandiri(df_bank.copy(), df_book_normalized)

    df_export, df_laporan_final, pesan_audit = perhalus_label_universal(df_export, df_laporan_final, pesan_audit)
    df_export, df_laporan_final = tambah_kolom_selisih(df_export, df_laporan_final)
    df_rekap_bulan = buat_rekap_per_bulan(df_laporan_final)

    df_lf_bln = df_laporan_final.copy()
    df_lf_bln["_Bln"] = df_lf_bln["Tanggal_Clean"].astype(str).str[:7]
    daftar_bulan = sorted(df_lf_bln["_Bln"].unique().tolist())

    if len(daftar_bulan) > 1:
        df_b_tmp = df_bank.copy()
        df_b_tmp["_Bln"] = pd.to_datetime(df_b_tmp["Tanggal"], errors="coerce").dt.strftime("%Y-%m")
        df_g_tmp = df_book_normalized.copy()
        df_g_tmp["_Tgl"] = pd.to_datetime(df_g_tmp["Date"], errors="coerce")
        df_g_tmp = df_g_tmp[df_g_tmp["_Tgl"].notna()].copy()
        df_g_tmp["_Bln"] = df_g_tmp["_Tgl"].dt.strftime("%Y-%m")
        df_g_tmp["Debit"] = pd.to_numeric(df_g_tmp["Debit"], errors="coerce").fillna(0)
        df_g_tmp["Credit"] = pd.to_numeric(df_g_tmp["Credit"], errors="coerce").fillna(0)

        blocks = []
        for bln in daftar_bulan:
            blocks.append(df_lf_bln[df_lf_bln["_Bln"] == bln].drop(columns=["_Bln"]))
            b_db = round(pd.to_numeric(df_b_tmp[df_b_tmp["_Bln"] == bln]["Debet"], errors="coerce").fillna(0).sum(), 2)
            b_kr = round(pd.to_numeric(df_b_tmp[df_b_tmp["_Bln"] == bln]["Kredit"], errors="coerce").fillna(0).sum(), 2)
            g_db = round(df_g_tmp[df_g_tmp["_Bln"] == bln]["Debit"].sum(), 2)
            g_cr = round(df_g_tmp[df_g_tmp["_Bln"] == bln]["Credit"].sum(), 2)
            blocks.append(pd.DataFrame([
                {"Tanggal_Clean": f"SUBTOTAL REKON [{bln}]", "Status_Rekonsiliasi": f"SUM Hasil Rekon Bulan {bln}"},
                {"Tanggal_Clean": f"SUBTOTAL ASLI [{bln}]", "Status_Rekonsiliasi": f"Total Asli PDF & GL Bulan {bln}",
                 "Debet": b_db, "Kredit": b_kr, "Debit": g_db, "Credit": g_cr,
                 "Selisih Bank dengan GL": round((b_kr - b_db) - (g_db - g_cr), 2)},
                {"Tanggal_Clean": f"CONTROL CHECK [{bln}]", "Status_Rekonsiliasi": f"Validasi Kelengkapan Bulan {bln}"}
            ]))

        blocks.append(pd.DataFrame([
            {"Tanggal_Clean": "GRAND TOTAL REKON", "Status_Rekonsiliasi": "SUM Gabungan Seluruh Bulan"},
            {"Tanggal_Clean": "GRAND TOTAL ASLI", "Status_Rekonsiliasi": "Total Asli Gabungan Seluruh Bulan"},
            {"Tanggal_Clean": "GRAND CONTROL CHECK", "Status_Rekonsiliasi": pesan_audit}
        ]))
        df_export = pd.concat(blocks, ignore_index=True)

    return {
        "nama_rek": nama_rek, "bank": bank_pilihan, "jml_pdf": len(pdf_files),
        "df_bank": df_bank, "df_export": df_export, "df_laporan_final": df_laporan_final,
        "df_rekap_bulan": df_rekap_bulan, "saldo_awal": saldo_awal, "saldo_akhir": saldo_akhir,
        "pdf_errors": pdf_errors, "pesan_audit": pesan_audit,
        "bulan_bank": bulan_bank, "hilang_di_gl": hilang_di_gl, "hilang_di_bank": hilang_di_bank
    }


# ============================================================
# HEADER UTAMA & TOMBOL LOGOUT (RESET TOTAL)
# ============================================================
head_col1, head_col2 = st.columns([5, 1])
with head_col1:
    st.markdown(
        """
        <h1 style='margin-bottom: 0.2rem;'>
            🏦 Aplikasi Rekonsiliasi Rekening Koran Bank vs General Ledger
            <sup style='font-size: 0.42em; font-weight: 600; color: #6c757d; letter-spacing: 0.5px;'>by FERDX</sup>
        </h1>
        """,
        unsafe_allow_html=True
    )
with head_col2:
    st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
    if st.button("🚪 Logout", help="Keluar & bersihkan semua data pekerjaan", use_container_width=True):
        st.session_state.clear()
        st.rerun()

st.markdown("Unggah **PDF Rekening Koran (1 bulan atau multi-bulan)** dan **Excel General Ledger** untuk satu atau beberapa rekening sekaligus.")
with st.container(border=True):
    c_jml, c_info = st.columns([1, 2.5])
    with c_jml:
        jumlah_rekening = st.number_input("Jumlah Rekening Diproses:", min_value=1, max_value=10, value=1, step=1)
    with c_info:
        st.info("💡 **Tips Multi-Rekening:** Atur angka di samping jika Anda ingin merekonsiliasi beberapa nomor rekening berbeda (BNI maupun Mandiri) dalam satu kali proses.")

daftar_input_rekening = []

for idx in range(int(jumlah_rekening)):
    with st.container(border=True):
        h1, h2 = st.columns([1.5, 2])
        with h1:
            nama_rek = st.text_input(f"Label / No. Rekening #{idx + 1}", value=f"Rekening_{idx + 1}", key=f"nama_rek_{idx}")
        with h2:
            bank_pilihan = st.radio(f"Pilih Bank Rekening #{idx + 1}:", ["BNI Direct", "Bank Mandiri"], horizontal=True, key=f"bank_rek_{idx}")

        col1, col2 = st.columns(2)
        with col1:
            pdfs = st.file_uploader(f"1. PDF Rekening Koran {nama_rek} (Bisa Multi-Bulan)", type=["pdf"], accept_multiple_files=True, key=f"pdf_rek_{idx}")
        with col2:
            book = st.file_uploader(f"2. Excel General Ledger {nama_rek} (.xlsx)", type=["xlsx"], key=f"book_rek_{idx}")

        mapping_kolom = {}
        df_book_raw = None
        if book:
            df_book_raw = pd.read_excel(book)
            kolom_tersedia = df_book_raw.columns.tolist()
            auto_map = deteksi_kolom_otomatis(kolom_tersedia)
            wajib_lengkap = all(auto_map[k] is not None for k in ["Date", "Label", "Debit", "Credit"])

            with st.expander(f"⚙️ Pengaturan Pemetaan Kolom GL ({nama_rek})", expanded=not wajib_lengkap):
                mc1, mc2, mc3, mc4 = st.columns(4)
                with mc1:
                    mapping_kolom["Date"] = st.selectbox("Tanggal *", kolom_tersedia, index=kolom_tersedia.index(auto_map["Date"]) if auto_map["Date"] in kolom_tersedia else 0, key=f"m_date_{idx}")
                with mc2:
                    mapping_kolom["Label"] = st.selectbox("Keterangan *", kolom_tersedia, index=kolom_tersedia.index(auto_map["Label"]) if auto_map["Label"] in kolom_tersedia else 0, key=f"m_lbl_{idx}")
                with mc3:
                    mapping_kolom["Debit"] = st.selectbox("Debit (Masuk) *", kolom_tersedia, index=kolom_tersedia.index(auto_map["Debit"]) if auto_map["Debit"] in kolom_tersedia else 0, key=f"m_deb_{idx}")
                with mc4:
                    mapping_kolom["Credit"] = st.selectbox("Kredit (Keluar) *", kolom_tersedia, index=kolom_tersedia.index(auto_map["Credit"]) if auto_map["Credit"] in kolom_tersedia else 0, key=f"m_cre_{idx}")

                opsi_ops = ["(Tidak Ada)"] + kolom_tersedia
                oc1, oc2 = st.columns(2)
                with oc1:
                    mapping_kolom["Number"] = st.selectbox("No. Bukti (Opsional)", opsi_ops, index=opsi_ops.index(auto_map["Number"]) if auto_map["Number"] in opsi_ops else 0, key=f"m_num_{idx}")
                with oc2:
                    mapping_kolom["Partner"] = st.selectbox("Partner (Opsional)", opsi_ops, index=opsi_ops.index(auto_map["Partner"]) if auto_map["Partner"] in opsi_ops else 0, key=f"m_par_{idx}")

        daftar_input_rekening.append({
            "nama_rek": nama_rek.strip() or f"Rekening_{idx + 1}",
            "bank": bank_pilihan, "pdfs": pdfs, "df_gl_raw": df_book_raw, "mapping": mapping_kolom
        })

rekening_siap = [r for r in daftar_input_rekening if r["pdfs"] and r["df_gl_raw"] is not None]

if rekening_siap:
    label_btn = f"🚀 Proses Rekonsiliasi ({len(rekening_siap)} Rekening)" if len(rekening_siap) > 1 else f"🚀 Proses Rekonsiliasi {rekening_siap[0]['bank']} Sekarang"
    if st.button(label_btn, type="primary", use_container_width=True):
        with st.spinner("Sedang mengekstrak PDF dan menjalankan audit rekonsiliasi..."):
            try:
                hasil_semua_rekening = [jalankan_rekon_satu_rekening(it["nama_rek"], it["bank"], it["pdfs"], it["df_gl_raw"], it["mapping"]) for it in rekening_siap]
                st.divider()

                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                    if len(hasil_semua_rekening) == 1:
                        h = hasil_semua_rekening[0]
                        tulis_sheet_hasil_dengan_rumus(writer, "Hasil Rekonsiliasi", h["df_export"], h["df_laporan_final"])
                        h["df_rekap_bulan"].to_excel(writer, sheet_name="Rekap Per Bulan", index=False)
                        h["df_bank"].to_excel(writer, sheet_name=f"Ekstraksi PDF {h['bank'][:10]}", index=False)
                    else:
                        summary_rows = []
                        for h in hasil_semua_rekening:
                            df_lf = h["df_laporan_final"]
                            selisih_tot = round((df_lf["Kredit"].fillna(0).sum() - df_lf["Debet"].fillna(0).sum()) - (df_lf["Debit"].fillna(0).sum() - df_lf["Credit"].fillna(0).sum()), 2)
                            periode = f"{h['bulan_bank'][0]} s.d. {h['bulan_bank'][-1]}" if len(h["bulan_bank"]) > 1 else (h["bulan_bank"][0] if h["bulan_bank"] else "-")
                            summary_rows.append({
                                "Nama Rekening": h["nama_rek"], "Bank": h["bank"], "Periode": periode,
                                "Saldo Awal": h["saldo_awal"], "Saldo Akhir": h["saldo_akhir"],
                                "Selisih Bank vs GL": selisih_tot, "Status Audit": h["pesan_audit"]
                            })
                        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary_Semua_Rekening", index=False)

                        for idx_h, h in enumerate(hasil_semua_rekening):
                            pref = re.sub(r'[^A-Za-z0-9_]', '', h["nama_rek"])[:14] or f"Rek{idx_h+1}"
                            tulis_sheet_hasil_dengan_rumus(writer, f"{pref}_Hasil_Rekon", h["df_export"], h["df_laporan_final"])
                            h["df_rekap_bulan"].to_excel(writer, sheet_name=f"{pref}_Rekap_Bulan", index=False)
                            h["df_bank"].to_excel(writer, sheet_name=f"{pref}_Ekstrak_PDF", index=False)

                nama_dl = "Hasil_Rekonsiliasi_Multi_Rekening.xlsx" if len(hasil_semua_rekening) > 1 else f"Hasil_Rekonsiliasi_{hasil_semua_rekening['nama_rek']}.xlsx"
                st.download_button("📥 Download File Excel Hasil Rekonsiliasi (.xlsx)", data=output_buffer.getvalue(), file_name=nama_dl, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)

                tabs_rek = st.tabs([f"🏦 {h['nama_rek']} ({h['bank']})" for h in hasil_semua_rekening])
                for tab_item, h in zip(tabs_rek, hasil_semua_rekening):
                    with tab_item:
                        if h["hilang_di_gl"]:
                            st.warning(f"⚠️ **Peringatan GL:** PDF mencakup `{h['bulan_bank']}` s.d. `{h['bulan_bank'][-1]}`, tetapi GL kosong di bulan: **{', '.join(h['hilang_di_gl'])}**.")
                        if h["hilang_di_bank"]:
                            st.warning(f"⚠️ **Peringatan PDF:** GL memiliki data pada bulan **{', '.join(h['hilang_di_bank'])}** yang belum diunggah PDF-nya.")
                        if len(h["pdf_errors"]) == 0:
                            st.success(f"**Ekstraksi {h['jml_pdf']} PDF Valid 100%!** ({len(h['df_bank'])} trx | Saldo Awal: Rp {h['saldo_awal']:,.2f} → Saldo Akhir: Rp {h['saldo_akhir']:,.2f})")
                        else:
                            st.warning(f"Ditemukan {len(h['pdf_errors'])} baris selisih pada pembacaan saldo PDF.")

                        if "OKE 100%" in h["pesan_audit"]:
                            st.success(f"**Status Audit:** {h['pesan_audit']}")
                        else:
                            st.error(f"**Status Audit:** {h['pesan_audit']}")

                        df_lf = h["df_laporan_final"]
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("Total Baris", len(df_lf))
                        m2.metric("Cocok (MATCH)", df_lf['Status_Rekonsiliasi'].str.startswith('MATCH').sum())
                        m3.metric("Hanya di Bank", df_lf['Status_Rekonsiliasi'].str.startswith('HANYA DI BANK').sum())
                        m4.metric("Hanya di GL", df_lf['Status_Rekonsiliasi'].str.startswith('HANYA DI GENERAL LEDGER').sum())

                        sub_rekap, sub_detail, sub_pdf = st.tabs(["📅 Rekap Per Bulan", "📊 Hasil Rekonsiliasi Lengkap", "📄 Hasil Ekstraksi PDF"])
                        with sub_rekap:
                            st.dataframe(h["df_rekap_bulan"], use_container_width=True)
                        with sub_detail:
                            opsi_bln = sorted(df_lf['Tanggal_Clean'].astype(str).str[:7].unique().tolist())
                            fc1, fc2 = st.columns(2)
                            with fc1:
                                f_bln = st.multiselect("Filter Bulan:", opsi_bln, default=opsi_bln, key=f"f_bln_{h['nama_rek']}")
                            with fc2:
                                f_sts = st.multiselect("Filter Status:", df_lf['Status_Rekonsiliasi'].unique().tolist(), default=df_lf['Status_Rekonsiliasi'].unique().tolist(), key=f"f_sts_{h['nama_rek']}")
                            st.dataframe(df_lf[(df_lf['Tanggal_Clean'].astype(str).str[:7].isin(f_bln)) & (df_lf['Status_Rekonsiliasi'].isin(f_sts))], use_container_width=True, height=420)
                        with sub_pdf:
                            st.dataframe(h["df_bank"], use_container_width=True, height=420)

            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses file: {e}")
