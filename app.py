return output_buffer.getvalue()


def warnai_baris_status(row):
    status = str(row.get("Status_Rekonsiliasi", ""))
    if status.startswith("MATCH"):
        return ["background-color: #F0FDF4"] * len(row)
    elif status.startswith(("NETTING JAGIR", "KONTRA BANK")):
        return ["background-color: #EFF6FF"] * len(row)
    elif status.startswith(("HANYA DI", "SELISIH")):
        return ["background-color: #FEF2F2"] * len(row)
    return [""] * len(row)


# ============================================================
# CSS ENTERPRISE WORKSPACE (SAGE GREEN & CALM UI)

# ============================================================
# CSS ENTERPRISE WORKSPACE (SAGE GREEN & CALM UI)
# ============================================================
st.markdown(
    """
    <style>
    .stApp {
        background-color: #F4F7F6;
    }
    header[data-testid="stHeader"] {
        background: transparent !important;
    }
    [data-testid="stToolbarActions"], [data-testid="stStatusWidget"] {
        display: none !important;
    }
    button[data-testid="stExpandSidebarButton"], [data-testid="collapsedControl"] {
        display: flex !important;
        visibility: visible !important;
        background-color: #6B8E78 !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        margin: 0.4rem !important;
    }
    section[data-testid="stSidebar"] {
        background-color: #6B8E78;
        min-width: 260px !important;
    }
    section[data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label {
        background-color: rgba(255, 255, 255, 0.08);
        padding: 0.55rem 0.85rem;
        border-radius: 8px;
        margin-bottom: 0.35rem;
        border: 1px solid rgba(255, 255, 255, 0.12);
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        background-color: rgba(255, 255, 255, 0.18);
    }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #FFFFFF;
        border-radius: 12px;
        border: 1px solid #DCE5E0;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.03);
        padding: 0.6rem;
    }
    div.stButton > button[kind="primary"] {
        background-color: #4A7259;
        border-color: #4A7259;
        color: #FFFFFF;
        font-weight: 600;
        border-radius: 8px;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #3B5D48;
        border-color: #3B5D48;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ============================================================
# SIDEBAR NAVIGASI UTAMA & PROFIL PENGGUNA
# ============================================================
with st.sidebar:
    st.markdown(
        """
        <div style="padding: 0.5rem 0 1rem 0; border-bottom: 1px solid rgba(255,255,255,0.2); margin-bottom: 1rem;">
            <div style="font-size: 1.35rem; font-weight: 700; letter-spacing: 0.3px;">🏦 REKON PRO</div>
            <div style="font-size: 0.78rem; opacity: 0.85;">Enterprise Bank vs GL • by FERDX</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    menu_aktif = st.radio(
        "Navigasi Workspace",
        ["📊 Dashboard", "🔄 Reconciliation", "📑 Reports", "⚙️ Settings"],
        index=1,
        label_visibility="collapsed"
    )

    st.markdown("<div style='height: 18rem;'></div>", unsafe_allow_html=True)
    st.markdown(
        """
        <div style="padding-top: 0.8rem; border-top: 1px solid rgba(255,255,255,0.2); display: flex; align-items: center; gap: 10px;">
            <div style="background: rgba(255,255,255,0.22); width: 34px; height: 34px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 700;">F</div>
            <div>
                <div style="font-weight: 600; font-size: 0.92rem;">Ferdy</div>
                <div style="font-size: 0.75rem; opacity: 0.85;">Finance & Audit Lead</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ============================================================
# HEADER & TOMBOL KUNCI SESI
# ============================================================
h_col1, h_col2 = st.columns([5, 1])
with h_col1:
    st.markdown(
        """
        <h2 style='margin-bottom: 0.15rem; color: #0F172A;'>
            🏦 Aplikasi Rekonsiliasi Rekening Koran Bank vs General Ledger
            <sup style='font-size: 0.45em; font-weight: 600; color: #64748B; letter-spacing: 0.5px;'>by FERDX</sup>
        </h2>
        """,
        unsafe_allow_html=True
    )
with h_col2:
    st.write("")
    if st.button("🔒 Kunci Sesi", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()


if "last_result" not in st.session_state:
    st.session_state.last_result = None

# ============================================================
# MENU 1: DASHBOARD RINGKASAN EKSEKUTIF
# ============================================================
if menu_aktif == "📊 Dashboard":
    st.subheader("📊 Executive Reconciliation Dashboard")
    res = st.session_state.last_result

    if res is None:
        with st.container(border=True):
            st.info("Belum ada data rekonsiliasi yang diproses pada sesi ini. Silakan buka menu **🔄 Reconciliation** di sidebar kiri untuk mengunggah file PDF Rekening Koran dan Excel General Ledger.")
    else:
        with st.container(border=True):
            st.caption(f"Ringkasan Sesi Terakhir: **{res['pilihan_bank']}** ({res['jumlah_pdf']} File PDF) • Saldo Awal: **Rp {res['saldo_awal']:,.2f}** → Saldo Akhir: **Rp {res['saldo_akhir']:,.2f}**")

            df_lap = res["df_laporan_final"]
            jml_match = df_lap['Status_Rekonsiliasi'].str.startswith('MATCH').sum()
            jml_jagir = df_lap['Status_Rekonsiliasi'].str.startswith(('NETTING JAGIR', 'KONTRA BANK')).sum()
            jml_unmatched = len(df_lap) - jml_match - jml_jagir
            rasio_match = round((jml_match / len(df_lap) * 100), 1) if len(df_lap) > 0 else 0.0

            d1, d2, d3, d4 = st.columns(4)
            d1.metric("Total Mutasi Diproses", f"{len(df_lap)} Baris")
            d2.metric("Tingkat Kecocokan", f"{rasio_match}%", delta=f"{jml_match} Baris Match")
            d3.metric("Netting & Kontra", f"{jml_jagir} Baris")
            d4.metric("Perlu Tindak Lanjut", f"{jml_unmatched} Baris", delta_color="inverse")

        col_chart, col_month = st.columns([1.1, 1.4])
        with col_chart:
            with st.container(border=True):
                st.markdown("**Distribusi Status Rekonsiliasi**")
                status_counts = df_lap["Status_Rekonsiliasi"].apply(
                    lambda s: "MATCH" if str(s).startswith("MATCH") else ("NETTING / KONTRA" if str(s).startswith(("NETTING JAGIR", "KONTRA BANK")) else "BELUM COCOK")
                ).value_counts()
                st.bar_chart(status_counts, color="#4A7259")

        with col_month:
            with st.container(border=True):
                st.markdown("**Ringkasan Per Bulan**")
                st.dataframe(res["df_rekap_bulan"], use_container_width=True, height=280)
# ============================================================
# MENU 2: RECONCILIATION (WORKSPACE UTAMA)
# ============================================================
elif menu_aktif == "🔄 Reconciliation":
    with st.container(border=True):
        pilihan_bank = st.radio(
            "Pilih Bank Rekening Koran:",
            ["BNI Direct", "Bank Mandiri"],
            horizontal=True
        )
        st.caption(f"Unggah **PDF Rekening Koran {pilihan_bank}** (1 s.d. 12 bulan) dan **Excel Catatan Transaksi Internal (ERP / Accounting Software / Excel)** untuk validasi otomatis.")

        col1, col2 = st.columns(2)
        with col1:
            uploaded_pdfs = st.file_uploader(
                f"1. Unggah Rekening Koran {pilihan_bank} (.pdf)",
                type=["pdf"],
                accept_multiple_files=True,
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

        with st.expander("⚙️ Pengaturan Pemetaan Kolom General Ledger (Otomatis Terdeteksi)", expanded=not wajib_lengkap):
            st.caption("Sistem otomatis mencocokkan nama kolom dari sistem akuntansi/ERP maupun Excel manual. Ubah jika diperlukan:")
            m_col1, m_col2, m_col3, m_col4 = st.columns(4)
            with m_col1:
                mapping_kolom["Date"] = st.selectbox("Kolom Tanggal *", options=kolom_tersedia, index=kolom_tersedia.index(auto_map["Date"]) if auto_map["Date"] in kolom_tersedia else 0)
            with m_col2:
                mapping_kolom["Label"] = st.selectbox("Kolom Keterangan / Deskripsi *", options=kolom_tersedia, index=kolom_tersedia.index(auto_map["Label"]) if auto_map["Label"] in kolom_tersedia else 0)
            with m_col3:
                mapping_kolom["Debit"] = st.selectbox("Kolom Debit (Uang Masuk) *", options=kolom_tersedia, index=kolom_tersedia.index(auto_map["Debit"]) if auto_map["Debit"] in kolom_tersedia else 0)
            with m_col4:
                mapping_kolom["Credit"] = st.selectbox("Kolom Kredit (Uang Keluar) *", options=kolom_tersedia, index=kolom_tersedia.index(auto_map["Credit"]) if auto_map["Credit"] in kolom_tersedia else 0)

            opsi_opsional = ["(Tidak Ada)"] + kolom_tersedia
            o_col1, o_col2 = st.columns(2)
            with o_col1:
                mapping_kolom["Number"] = st.selectbox("Kolom No. Bukti / Referensi (Opsional)", options=opsi_opsional, index=opsi_opsional.index(auto_map["Number"]) if auto_map["Number"] in opsi_opsional else 0)
            with o_col2:
                mapping_kolom["Partner"] = st.selectbox("Kolom Nama Rekanan / Partner (Opsional)", options=opsi_opsional, index=opsi_opsional.index(auto_map["Partner"]) if auto_map["Partner"] in opsi_opsional else 0)

    if uploaded_pdfs and uploaded_book and df_book_raw is not None:
        label_tombol = f"🚀 Proses Rekonsiliasi {pilihan_bank} ({len(uploaded_pdfs)} File PDF)" if len(uploaded_pdfs) > 1 else f"🚀 Proses Rekonsiliasi {pilihan_bank} Sekarang"
        if st.button(label_tombol, type="primary", use_container_width=True):
            with st.spinner(f"Sedang mengekstrak {len(uploaded_pdfs)} PDF {pilihan_bank} dan mencocokkan dengan General Ledger..."):
                try:
                    df_book_normalized = normalisasi_excel_pembukuan(df_book_raw, mapping_kolom)

                    hasil_ekstraksi_list = []
                    for pdf_item in uploaded_pdfs:
                        if pilihan_bank == "BNI Direct":
                            df_p, s_awal, s_akhir, err_p = ekstrak_pdf_bni(pdf_item)
                        else:
                            df_p, s_awal, s_akhir, err_p = ekstrak_pdf_mandiri(pdf_item)

                        df_p["Sumber_File"] = pdf_item.name
                        tgl_awal_file = pd.to_datetime(df_p["Tanggal"]).min()
                        hasil_ekstraksi_list.append({
                            "nama": pdf_item.name, "tgl_awal": tgl_awal_file,
                            "df": df_p, "saldo_awal": s_awal, "saldo_akhir": s_akhir, "errors": err_p
                        })

                    hasil_ekstraksi_list.sort(key=lambda x: x["tgl_awal"])
                    df_bank = pd.concat([x["df"] for x in hasil_ekstraksi_list], ignore_index=True)
                    saldo_awal = hasil_ekstraksi_list[0]["saldo_awal"]
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

                    from openpyxl.utils import get_column_letter
                    output_buffer = io.BytesIO()
                    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                        df_export.to_excel(writer, sheet_name="Hasil Rekonsiliasi", index=False)
                        df_rekap_bulan.to_excel(writer, sheet_name="Rekap Per Bulan", index=False)
                        df_bank.to_excel(writer, sheet_name=f"Ekstraksi PDF {pilihan_bank[:10]}", index=False)

                        ws = writer.sheets["Hasil Rekonsiliasi"]
                        headers = [cell.value for cell in ws[1]]
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

                    nama_prefix = f"{len(uploaded_pdfs)}_Bulan_{pilihan_bank.replace(' ', '_')}" if len(uploaded_pdfs) > 1 else uploaded_pdfs[0].name.replace('.pdf', '')
                    nama_file_keluar = f"Hasil_Rekonsiliasi_{nama_prefix}.xlsx"

                    st.session_state.last_result = {
                        "pilihan_bank": pilihan_bank,
                        "jumlah_pdf": len(uploaded_pdfs),
                        "saldo_awal": saldo_awal,
                        "saldo_akhir": saldo_akhir,
                        "pdf_errors": pdf_errors,
                        "pesan_audit": pesan_audit,
                        "hilang_di_gl": hilang_di_gl,
                        "hilang_di_bank": hilang_di_bank,
                        "bulan_bank": bulan_bank,
                        "df_bank": df_bank,
                        "df_export": df_export,
                        "df_laporan_final": df_laporan_final,
                        "df_rekap_bulan": df_rekap_bulan,
                        "excel_bytes": output_buffer.getvalue(),
                        "nama_file_keluar": nama_file_keluar
                    }
                except Exception as e:
                    st.error(f"Terjadi kesalahan saat memproses file: {e}")
    if st.session_state.last_result is not None:
        res = st.session_state.last_result
        st.divider()

        if res["hilang_di_gl"]:
            st.warning(f"⚠️ **PERINGATAN KELENGKAPAN GENERAL LEDGER:** PDF Rekening Koran mencakup **{len(res['bulan_bank'])} bulan** (`{res['bulan_bank'][0]}` s.d. `{res['bulan_bank'][-1]}`), tetapi file General Ledger **tidak memiliki data transaksi** pada bulan: **{', '.join(res['hilang_di_gl'])}**.")
        if res["hilang_di_bank"]:
            st.warning(f"⚠️ **PERINGATAN KELENGKAPAN PDF BANK:** File General Ledger memiliki transaksi pada bulan **{', '.join(res['hilang_di_bank'])}** yang belum diunggah file PDF Rekening Korannya.")

        if len(res["pdf_errors"]) == 0:
            st.success(f"**Ekstraksi {res['jumlah_pdf']} PDF {res['pilihan_bank']} Valid 100%!** ({len(res['df_bank'])} transaksi | Saldo Awal: Rp {res['saldo_awal']:,.2f} → Saldo Akhir: Rp {res['saldo_akhir']:,.2f})")
        else:
            st.warning(f"Ditemukan {len(res['pdf_errors'])} baris selisih pada pembacaan saldo PDF.")

        if "OKE 100%" in res["pesan_audit"]:
            st.success(f"**Status Audit Rekonsiliasi:** {res['pesan_audit']}")
        else:
            st.error(f"**Status Audit Rekonsiliasi:** {res['pesan_audit']}")

        with st.container(border=True):
            df_lap = res["df_laporan_final"]
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Baris Laporan", len(df_lap))
            m2.metric("Cocok (MATCH)", df_lap['Status_Rekonsiliasi'].str.startswith('MATCH').sum())
            m3.metric("Hanya di Bank", df_lap['Status_Rekonsiliasi'].str.startswith('HANYA DI BANK').sum())
            m4.metric("Hanya di General Ledger", df_lap['Status_Rekonsiliasi'].str.startswith('HANYA DI GENERAL LEDGER').sum())

        with st.container(border=True):
            tab_rekap, tab1, tab2 = st.tabs([
                "📅 Rekap Selisih Per Bulan",
                "📊 Hasil Rekonsiliasi Lengkap",
                f"📄 Hasil Ekstraksi PDF {res['pilihan_bank']}"
            ])

            with tab_rekap:
                st.dataframe(res["df_rekap_bulan"], use_container_width=True)

            with tab1:
                f_col1, f_col2 = st.columns(2)
                with f_col1:
                    opsi_bulan = sorted(df_lap['Tanggal_Clean'].astype(str).str[:7].unique().tolist())
                    filter_bulan = st.multiselect("Filter Bulan:", options=opsi_bulan, default=opsi_bulan)
                with f_col2:
                    filter_status = st.multiselect(
                        "Filter Status Rekonsiliasi:",
                        options=df_lap['Status_Rekonsiliasi'].unique().tolist(),
                        default=df_lap['Status_Rekonsiliasi'].unique().tolist()
                    )
                df_tampil = df_lap[
                    (df_lap['Tanggal_Clean'].astype(str).str[:7].isin(filter_bulan))
                    & (df_lap['Status_Rekonsiliasi'].isin(filter_status))
                ]
                st.dataframe(df_tampil.style.apply(warnai_baris_status, axis=1), use_container_width=True, height=450)

            with tab2:
                st.dataframe(res["df_bank"], use_container_width=True, height=450)

        st.download_button(
            label=f"📥 Download Hasil Rekonsiliasi {res['pilihan_bank']} (.xlsx)",
            data=res["excel_bytes"],
            file_name=res["nama_file_keluar"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

# ============================================================
# MENU 3: REPORTS (ARSIP & RINGKASAN AUDIT)
# ============================================================
elif menu_aktif == "📑 Reports":
    st.subheader("📑 Reconciliation Audit Reports")
    res = st.session_state.last_result

    if res is None:
        with st.container(border=True):
            st.info("Belum ada laporan yang tersedia. Jalankan proses pada menu **🔄 Reconciliation** terlebih dahulu.")
    else:
        with st.container(border=True):
            st.markdown(f"**Laporan Audit:** `{res['nama_file_keluar']}`")
            st.write(f"**Status Validasi:** {res['pesan_audit']}")
            st.dataframe(res["df_export"].tail(3), use_container_width=True)
            st.download_button(
                label="📥 Unduh Ulang Laporan Excel (.xlsx)",
                data=res["excel_bytes"],
                file_name=res["nama_file_keluar"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

# ============================================================
# MENU 4: SETTINGS (PARAMETER & PANDUAN STANDAR)
# ============================================================
elif menu_aktif == "⚙️ Settings":
    st.subheader("⚙️ Workspace Configuration & Rules")
    with st.container(border=True):
        st.markdown("**Parameter Mesin Rekonsiliasi Aktif**")
        st.markdown(
            """
            * **Toleransi Pembulatan Nominal:** Maksimal **Rp 1,00** untuk transaksi reguler dan **Rp 5,00** untuk *Netting Jagir*.
            * **Toleransi Selisih Kurs Valas (BNI):** Maksimal **Rp 500,00** jika label mengandung simbol `$` atau `€`.
            * **Rentang Pencocokan Beda Tanggal (BNI):** Hingga **±10 hari** untuk transaksi satuan dan **±5 hari** untuk transaksi gabungan (*split match*).
            * **Pengelompokan Biaya Admin:** Otomatis diakumulasikan per bulan (`YYYY-MM`).
            * **Standar Kolom Excel General Ledger:** Mendukung otomatis kolom `Date/Tanggal`, `Label/Keterangan`, `Debit/Debet` (Uang Masuk), dan `Credit/Kredit` (Uang Keluar).
            """
        )





