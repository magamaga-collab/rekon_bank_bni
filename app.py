import streamlit as st
import pandas as pd
import io

from modul_bni import ekstrak_pdf_bni, proses_rekonsiliasi_bni
from modul_mandiri import ekstrak_pdf_mandiri, proses_rekonsiliasi_mandiri

# ============================================================
# KONFIGURASI HALAMAN WEB
# ============================================================
st.set_page_config(
    page_title="Rekonsiliasi Rekening Koran vs Pembukuan",
    page_icon="🏦",
    layout="wide"
)

# ============================================================
# HALAMAN LOGIN PASSWORD
# ============================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 Login Aplikasi Rekonsiliasi Bank")
    pwd = st.text_input("Masukkan Password Akses:", type="password")
    if st.button("Masuk", type="primary"):
        if pwd == st.secrets["APP_PASSWORD"]:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Password salah! Silakan coba lagi.")
    st.stop()
# ============================================================
# FUNGSI PENERJEMAH KOLOM UNIVERSAL (ODOO / ACCURATE / EXCEL)
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
                "STATUS AUDIT": "CONTROL CHECK",

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

    mask_transaksi = ~df_export["Tanggal_Clean"].isin(["TOTAL REKON", "TOTAL ASLI", "STATUS AUDIT"])
    total_selisih_sum = round(df_export.loc[mask_transaksi, nama_kolom].sum(), 2)
    df_export.loc[df_export["Tanggal_Clean"] == "TOTAL REKON", nama_kolom] = total_selisih_sum

    return df_export, df_laporan_final


# ============================================================
# HEADER & PILIHAN BANK
# ============================================================
st.markdown(
    """
    <h1 style='margin-bottom: 0.2rem;'>
        🏦 Aplikasi Rekonsiliasi Rekening Koran Bank vs General Ledger
        <sup style='font-size: 0.42em; font-weight: 600; color: #6c757d; letter-spacing: 0.5px;'>by FERDX</sup>
    </h1>
    """,
    unsafe_allow_html=True
)

pilihan_bank = st.radio(
    "Pilih Bank Rekening Koran:",
    ["BNI Direct", "Bank Mandiri"],
    horizontal=True
)

st.markdown(f"Unggah **PDF Rekening Koran {pilihan_bank}** dan **Excel Catatan Transaksi Internal/GL (ERP / Accounting Software / Excel)** untuk validasi otomatis.")
# ============================================================
# ANTARMUKA PENGGUNA (STREAMLIT UI)
# ============================================================
col1, col2 = st.columns(2)
with col1:
    uploaded_pdf = st.file_uploader(
        f"1. Unggah Rekening Koran {pilihan_bank} (.pdf)",
        type=["pdf"],
        key=f"pdf_{pilihan_bank}"
    )
with col2:
    uploaded_book = st.file_uploader(
        "2. Unggah Excel General Ledger (.xlsx)",
        type=["xlsx"],
        key=f"book_{pilihan_bank}"
    )

mapping_kolom = {}
df_book_raw = None

if uploaded_book:
    df_book_raw = pd.read_excel(uploaded_book)
    kolom_tersedia = df_book_raw.columns.tolist()
    auto_map = deteksi_kolom_otomatis(kolom_tersedia)

    wajib_lengkap = all(auto_map[k] is not None for k in ["Date", "Label", "Debit", "Credit"])

    with st.expander("⚙️ Pengaturan Pemetaan Kolom Pembukuan (Otomatis Terdeteksi)", expanded=not wajib_lengkap):
        st.caption("Sistem otomatis mencocokkan nama kolom dari Odoo, Accurate, maupun Excel manual. Ubah jika diperlukan:")
        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        with m_col1:
            mapping_kolom["Date"] = st.selectbox(
                "Kolom Tanggal *",
                options=kolom_tersedia,
                index=kolom_tersedia.index(auto_map["Date"]) if auto_map["Date"] in kolom_tersedia else 0
            )
        with m_col2:
            mapping_kolom["Label"] = st.selectbox(
                "Kolom Keterangan / Deskripsi *",
                options=kolom_tersedia,
                index=kolom_tersedia.index(auto_map["Label"]) if auto_map["Label"] in kolom_tersedia else 0
            )
        with m_col3:
            mapping_kolom["Debit"] = st.selectbox(
                "Kolom Debit (Uang Masuk) *",
                options=kolom_tersedia,
                index=kolom_tersedia.index(auto_map["Debit"]) if auto_map["Debit"] in kolom_tersedia else 0
            )
        with m_col4:
            mapping_kolom["Credit"] = st.selectbox(
                "Kolom Kredit (Uang Keluar) *",
                options=kolom_tersedia,
                index=kolom_tersedia.index(auto_map["Credit"]) if auto_map["Credit"] in kolom_tersedia else 0
            )

        opsi_opsional = ["(Tidak Ada)"] + kolom_tersedia
        o_col1, o_col2 = st.columns(2)
        with o_col1:
            mapping_kolom["Number"] = st.selectbox(
                "Kolom No. Bukti / Referensi (Opsional)",
                options=opsi_opsional,
                index=opsi_opsional.index(auto_map["Number"]) if auto_map["Number"] in opsi_opsional else 0
            )
        with o_col2:
            mapping_kolom["Partner"] = st.selectbox(
                "Kolom Nama Rekanan / Partner (Opsional)",
                options=opsi_opsional,
                index=opsi_opsional.index(auto_map["Partner"]) if auto_map["Partner"] in opsi_opsional else 0
            )
if uploaded_pdf and uploaded_book and df_book_raw is not None:
    if st.button(f"🚀 Proses Rekonsiliasi {pilihan_bank} Sekarang", type="primary", use_container_width=True):
        with st.spinner(f"Sedang mengekstrak PDF {pilihan_bank} dan mencocokkan dengan catatan pembukuan..."):
            try:
                df_book_normalized = normalisasi_excel_pembukuan(df_book_raw, mapping_kolom)

                if pilihan_bank == "BNI Direct":
                    df_bank, saldo_awal, saldo_akhir, pdf_errors = ekstrak_pdf_bni(uploaded_pdf)
                    df_export, df_laporan_final, pesan_audit = proses_rekonsiliasi_bni(df_bank.copy(), df_book_normalized)
                else:
                    df_bank, saldo_awal, saldo_akhir, pdf_errors = ekstrak_pdf_mandiri(uploaded_pdf)
                    df_export, df_laporan_final, pesan_audit = proses_rekonsiliasi_mandiri(df_bank.copy(), df_book_normalized)

                df_export, df_laporan_final, pesan_audit = perhalus_label_universal(df_export, df_laporan_final, pesan_audit)
                df_export, df_laporan_final = tambah_kolom_selisih(df_export, df_laporan_final)
                st.divider()

                if len(pdf_errors) == 0:
                    st.success(f"**Ekstraksi PDF {pilihan_bank} Valid 100%!** ({len(df_bank)} transaksi | Saldo Awal: Rp {saldo_awal:,.2f} → Saldo Akhir: Rp {saldo_akhir:,.2f})")
                else:
                    st.warning(f"Ditemukan {len(pdf_errors)} baris selisih pada pembacaan saldo PDF.")

                if "OKE 100%" in pesan_audit:
                    st.success(f"**Status Audit Rekonsiliasi:** {pesan_audit}")
                else:
                    st.error(f"**Status Audit Rekonsiliasi:** {pesan_audit}")

                m1, m2, m3, m4 = st.columns(4)
                jml_match = df_laporan_final['Status_Rekonsiliasi'].str.startswith('MATCH').sum()
                jml_bank_only = df_laporan_final['Status_Rekonsiliasi'].str.startswith('HANYA DI BANK').sum()
                jml_book_only = df_laporan_final['Status_Rekonsiliasi'].str.startswith('HANYA DI PEMBUKUAN').sum()

                m1.metric("Total Baris Laporan", len(df_laporan_final))
                m2.metric("Cocok (MATCH)", jml_match)
                m3.metric("Hanya di Bank", jml_bank_only)
                m4.metric("Hanya di General Ledger", jml_book_only)

                tab1, tab2 = st.tabs(["📊 Hasil Rekonsiliasi Lengkap", f"📄 Hasil Ekstraksi PDF {pilihan_bank}"])
                with tab1:
                    filter_status = st.multiselect(
                        "Filter Status Rekonsiliasi:",
                        options=df_laporan_final['Status_Rekonsiliasi'].unique().tolist(),
                        default=df_laporan_final['Status_Rekonsiliasi'].unique().tolist()
                    )
                    st.dataframe(
                        df_laporan_final[df_laporan_final['Status_Rekonsiliasi'].isin(filter_status)],
                        use_container_width=True,
                        height=450
                    )
                with tab2:
                    st.dataframe(df_bank, use_container_width=True, height=450)

                from openpyxl.utils import get_column_letter

                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                    df_export.to_excel(writer, sheet_name="Hasil Rekonsiliasi", index=False)
                    df_bank.to_excel(writer, sheet_name=f"Ekstraksi PDF {pilihan_bank[:10]}", index=False)

                    ws = writer.sheets["Hasil Rekonsiliasi"]
                    headers = [cell.value for cell in ws]
                    last_data_row = len(df_laporan_final) + 1
                    row_total_rekon = last_data_row + 1
                    row_total_asli = last_data_row + 2
                    row_control_check = last_data_row + 3

                    col_idx = {name: i + 1 for i, name in enumerate(headers)}
                    col_letter = {name: get_column_letter(i + 1) for i, name in enumerate(headers)}

                    # 1. Tulis rumus per baris di kolom "Selisih Bank dengan GL"
                    if all(k in col_letter for k in ["Kredit", "Debet", "Debit", "Credit", "Selisih Bank dengan GL"]):
                        c_kredit = col_letter["Kredit"]
                        c_debet = col_letter["Debet"]
                        c_debit = col_letter["Debit"]
                        c_credit = col_letter["Credit"]
                        idx_selisih = col_idx["Selisih Bank dengan GL"]

                        for r in range(2, last_data_row + 1):
                            ws.cell(row=r, column=idx_selisih).value = f"=({c_kredit}{r}-{c_debet}{r})-({c_debit}{r}-{c_credit}{r})"

                    # 2. Tulis rumus =SUM(...) di baris TOTAL REKON & pengurangan di CONTROL CHECK
                    for nama_kol in ["Debet", "Kredit", "Debit", "Credit", "Selisih Bank dengan GL"]:
                        if nama_kol in col_idx:
                            c_idx = col_idx[nama_kol]
                            c_let = col_letter[nama_kol]
                            ws.cell(row=row_total_rekon, column=c_idx).value = f"=SUM({c_let}2:{c_let}{last_data_row})"
                            ws.cell(row=row_control_check, column=c_idx).value = f"={c_let}{row_total_rekon}-{c_let}{row_total_asli}"


                nama_file_keluar = f"Hasil_Rekonsiliasi_{uploaded_pdf.name.replace('.pdf', '')}.xlsx"
                st.download_button(
                    label=f"📥 Download Hasil Rekonsiliasi {pilihan_bank} (.xlsx)",
                    data=output_buffer.getvalue(),
                    file_name=nama_file_keluar,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses file: {e}")
