import streamlit as st
import pandas as pd
import io

from modul_bni import ekstrak_pdf_bni, proses_rekonsiliasi_bni
from modul_mandiri import ekstrak_pdf_mandiri, proses_rekonsiliasi_mandiri

# ============================================================
# KONFIGURASI HALAMAN WEB
# ============================================================
st.set_page_config(
    page_title="Rekonsiliasi Rekening Koran vs Odoo",
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
# HEADER & PILIHAN BANK
# ============================================================
st.title("🏦 Aplikasi Rekonsiliasi Rekening Koran Bank vs Odoo")

pilihan_bank = st.radio(
    "Pilih Bank Rekening Koran:",
    ["BNI Direct", "Bank Mandiri"],
    horizontal=True
)

st.markdown(f"Unggah **PDF Rekening Koran {pilihan_bank}** dan **Excel Mutasi Odoo** untuk memproses ekstraksi dan rekonsiliasi otomatis.")

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
    uploaded_odoo = st.file_uploader(
        "2. Unggah Mutasi Bank Odoo (.xlsx)",
        type=["xlsx"],
        key=f"odoo_{pilihan_bank}"
    )

if uploaded_pdf and uploaded_odoo:
    if st.button(f"🚀 Proses Rekonsiliasi {pilihan_bank} Sekarang", type="primary", use_container_width=True):
        with st.spinner(f"Sedang mengekstrak PDF {pilihan_bank} dan mencocokkan dengan data Odoo..."):
            try:
                df_odoo_raw = pd.read_excel(uploaded_odoo)

                if pilihan_bank == "BNI Direct":
                    df_bank, saldo_awal, saldo_akhir, pdf_errors = ekstrak_pdf_bni(uploaded_pdf)
                    df_export, df_laporan_final, pesan_audit = proses_rekonsiliasi_bni(df_bank.copy(), df_odoo_raw)
                else:
                    df_bank, saldo_awal, saldo_akhir, pdf_errors = ekstrak_pdf_mandiri(uploaded_pdf)
                    df_export, df_laporan_final, pesan_audit = proses_rekonsiliasi_mandiri(df_bank.copy(), df_odoo_raw)

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
                jml_odoo_only = df_laporan_final['Status_Rekonsiliasi'].str.startswith('HANYA DI ODOO').sum()

                m1.metric("Total Baris Laporan", len(df_laporan_final))
                m2.metric("Cocok (MATCH)", jml_match)
                m3.metric("Hanya di Bank", jml_bank_only)
                m4.metric("Hanya di Odoo", jml_odoo_only)

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

                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                    df_export.to_excel(writer, sheet_name="Hasil Rekonsiliasi", index=False)
                    df_bank.to_excel(writer, sheet_name=f"Ekstraksi PDF {pilihan_bank[:10]}", index=False)

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
