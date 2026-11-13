import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Border, Side
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
# HALAMAN LOGIN PASSWORD (ENTERPRISE CALM UI)
# ============================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown(
        """
        <style>
        .stApp {
            background-color: #F4F7FA;
        }
        header[data-testid="stHeader"] {
            background: transparent;
        }
        [data-testid="stToolbar"] {
            display: none;
        }
        div[data-testid="stForm"] {
            background-color: #FFFFFF;
            padding: 2.2rem 2rem;
            border-radius: 14px;
            border: 1px solid #E2E8F0;
            box-shadow: 0 10px 25px rgba(15, 23, 42, 0.06);
        }
        div[data-testid="stForm"] button[kind="primaryFormSubmit"] {
            background-color: #1E3A8A;
            border-color: #1E3A8A;
            color: #FFFFFF;
            font-weight: 600;
            border-radius: 8px;
            padding: 0.55rem 1rem;
        }
        div[data-testid="stForm"] button[kind="primaryFormSubmit"]:hover {
            background-color: #1D4ED8;
            border-color: #1D4ED8;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.write("")
    st.write("")
    _, col_center, _ = st.columns([1.2, 1.3, 1.2])

    with col_center:
        with st.form("login_form", clear_on_submit=False):
            st.markdown(
                """
                <div style="text-align: center; margin-bottom: 1.2rem;">
                    <div style="font-size: 2.2rem; margin-bottom: 0.2rem;">🏦</div>
                    <h3 style="margin: 0; color: #0F172A; font-weight: 700;">
                        Rekonsiliasi Bank vs GL
                        <sup style="font-size: 0.5em; color: #64748B; font-weight: 600;">by FERDX</sup>
                    </h3>
                    <p style="margin: 0.35rem 0 0 0; font-size: 0.88rem; color: #64748B;">
                        Enterprise Automated Reconciliation Workspace
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )

            pwd = st.text_input("Password Akses", type="password", placeholder="Masukkan kata sandi tim...", autocomplete="current-password")
            st.write("")
            submitted = st.form_submit_button("🔐 Masuk ke Aplikasi", type="primary", use_container_width=True)

            if submitted:
                if pwd == st.secrets["APP_PASSWORD"]:
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Password salah! Silakan periksa kembali.")

            st.markdown(
                """
                <div style="text-align: center; margin-top: 1rem; font-size: 0.78rem; color: #94A3B8;">
                    🔒 Akses Terbatas • Pemrosesan Aman di Memori
                </div>
                """,
                unsafe_allow_html=True
            )
    st.stop()



# ============================================================
# FUNGSI PENERJEMAH KOLOM & PEMBANTU UNIVERSAL
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
def buat_excel_dengan_kontrol_bulanan(df_laporan_final, df_bank, df_gl_norm, df_export_asli, df_rekap_bulan, pilihan_bank):
    kolom_utama = df_laporan_final.columns.tolist()

    # Hitung total asli per bulan dari masing-masing sumber mentah
    df_b = df_bank.copy()
    df_b["Bulan"] = pd.to_datetime(df_b["Tanggal"], errors="coerce").dt.strftime("%Y-%m")
    df_b["Debet"] = pd.to_numeric(df_b["Debet"], errors="coerce").fillna(0)
    df_b["Kredit"] = pd.to_numeric(df_b["Kredit"], errors="coerce").fillna(0)

    df_g = df_gl_norm.copy()
    df_g["Bulan"] = pd.to_datetime(df_g["Date"], errors="coerce").dt.strftime("%Y-%m")
    df_g["Debit"] = pd.to_numeric(df_g["Debit"], errors="coerce").fillna(0)
    df_g["Credit"] = pd.to_numeric(df_g["Credit"], errors="coerce").fillna(0)
    df_g_valid = df_g[df_g["Bulan"].notna()].copy()

    # Ambil Grand Total Asli dari df_export_asli (sudah termasuk header bawaan GL)
    baris_asli_global = df_export_asli[df_export_asli["Tanggal_Clean"] == "TOTAL ASLI"].iloc[0]
    grand_debet_bank = float(baris_asli_global.get("Debet", 0) or 0)
    grand_kredit_bank = float(baris_asli_global.get("Kredit", 0) or 0)
    grand_debit_gl = float(baris_asli_global.get("Debit", 0) or 0)
    grand_credit_gl = float(baris_asli_global.get("Credit", 0) or 0)

    df_lap = df_laporan_final.copy()
    df_lap["Bulan"] = df_lap["Tanggal_Clean"].astype(str).str[:7]
    daftar_bulan = sorted(df_lap["Bulan"].unique().tolist())

    rows_terstruktur = []
    info_blok_bulan = []  # menyimpan koordinat baris Excel tiap bulan

    current_excel_row = 2  # baris 1 adalah header
    for bln in daftar_bulan:
        grp = df_lap[df_lap["Bulan"] == bln].drop(columns=["Bulan"])
        start_row = current_excel_row
        for _, r in grp.iterrows():
            rows_terstruktur.append(r.to_dict())
            current_excel_row += 1
        end_row = current_excel_row - 1

        # Hitung total asli khusus bulan ini
        asli_debet_b = round(df_b[df_b["Bulan"] == bln]["Debet"].sum(), 2)
        asli_kredit_b = round(df_b[df_b["Bulan"] == bln]["Kredit"].sum(), 2)
        asli_debit_g = round(df_g_valid[df_g_valid["Bulan"] == bln]["Debit"].sum(), 2)
        asli_credit_g = round(df_g_valid[df_g_valid["Bulan"] == bln]["Credit"].sum(), 2)
        asli_selisih = round((asli_kredit_b - asli_debet_b) - (asli_debit_g - asli_credit_g), 2)

        row_rekon = current_excel_row
        rows_terstruktur.append({
            "Tanggal_Clean": f"TOTAL REKON ({bln})",
            "Status_Rekonsiliasi": f"SUM Hasil Rekon Bulan {bln}"
        })
        current_excel_row += 1

        row_asli = current_excel_row
        rows_terstruktur.append({
            "Tanggal_Clean": f"TOTAL ASLI ({bln})",
            "Status_Rekonsiliasi": f"Total Asli PDF & GL Bulan {bln}",
            "Debet": asli_debet_b,
            "Kredit": asli_kredit_b,
            "Debit": asli_debit_g,
            "Credit": asli_credit_g,
            "Selisih Bank dengan GL": asli_selisih
        })
        current_excel_row += 1

        row_ctrl = current_excel_row
        rows_terstruktur.append({
            "Tanggal_Clean": f"CONTROL CHECK ({bln})",
            "Status_Rekonsiliasi": f"Selisih Kontrol Bulan {bln} (Harus 0)"
        })
        current_excel_row += 1

        # 1 baris kosong pemisah antarbulan
        rows_terstruktur.append({})
        current_excel_row += 1

        info_blok_bulan.append({
            "start": start_row,
            "end": end_row,
            "rekon": row_rekon,
            "asli": row_asli,
            "ctrl": row_ctrl
        })

    # Tambahkan blok GRAND TOTAL keseluruhan di paling bawah
    row_grand_rekon = current_excel_row
    rows_terstruktur.append({
        "Tanggal_Clean": "GRAND TOTAL REKON",
        "Status_Rekonsiliasi": "Total Gabungan Seluruh Bulan Hasil Rekon"
    })
    current_excel_row += 1

    row_grand_asli = current_excel_row
    grand_selisih = round((grand_kredit_bank - grand_debet_bank) - (grand_debit_gl - grand_credit_gl), 2)
    rows_terstruktur.append({
        "Tanggal_Clean": "GRAND TOTAL ASLI",
        "Status_Rekonsiliasi": "Total Asli Keseluruhan PDF & Header GL",
        "Debet": grand_debet_bank,
        "Kredit": grand_kredit_bank,
        "Debit": grand_debit_gl,
        "Credit": grand_credit_gl,
        "Selisih Bank dengan GL": grand_selisih
    })
    current_excel_row += 1

    row_grand_ctrl = current_excel_row
    rows_terstruktur.append({
        "Tanggal_Clean": "GRAND CONTROL CHECK",
        "Status_Rekonsiliasi": "Validasi Akhir Keseluruhan (Harus 0)"
    })

    df_final_excel = pd.DataFrame(rows_terstruktur, columns=kolom_utama)

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        df_final_excel.to_excel(writer, sheet_name="Hasil Rekonsiliasi", index=False)
        df_rekap_bulan.to_excel(writer, sheet_name="Rekap Per Bulan", index=False)
        df_bank.to_excel(writer, sheet_name=f"Ekstraksi PDF {pilihan_bank[:10]}", index=False)

        ws = writer.sheets["Hasil Rekonsiliasi"]
        headers = [cell.value for cell in ws[1]]
        col_idx = {name: i + 1 for i, name in enumerate(headers)}
        col_letter = {name: get_column_letter(i + 1) for i, name in enumerate(headers)}

        fill_rekon = PatternFill(start_color="EDF2F8", end_color="EDF2F8", fill_type="solid")
        fill_asli = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
        fill_ctrl = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        fill_grand = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        font_bold = Font(bold=True)
        border_top_thin = Border(top=Side(style="thin", color="808080"))
        border_bottom_double = Border(bottom=Side(style="double", color="000000"), top=Side(style="thin", color="808080"))

        kolom_angka = ["Debet", "Kredit", "Saldo", "Debit", "Credit", "Selisih Bank dengan GL"]
        kolom_hitung = ["Debet", "Kredit", "Debit", "Credit", "Selisih Bank dengan GL"]

        # Tulis rumus per baris transaksi dan blok 3 baris kontrol tiap bulan
        for blok in info_blok_bulan:
            if all(k in col_letter for k in ["Kredit", "Debet", "Debit", "Credit", "Selisih Bank dengan GL"]):
                c_kr, c_db, c_dgl, c_cgl = col_letter["Kredit"], col_letter["Debet"], col_letter["Debit"], col_letter["Credit"]
                idx_sel = col_idx["Selisih Bank dengan GL"]
                for r in range(blok["start"], blok["end"] + 1):
                    ws.cell(row=r, column=idx_sel).value = f"=({c_kr}{r}-{c_db}{r})-({c_dgl}{r}-{c_cgl}{r})"

            for k_col in kolom_hitung:
                if k_col in col_idx:
                    ci, cl = col_idx[k_col], col_letter[k_col]
                    ws.cell(row=blok["rekon"], column=ci).value = f"=SUM({cl}{blok['start']}:{cl}{blok['end']})"
                    ws.cell(row=blok["ctrl"], column=ci).value = f"={cl}{blok['rekon']}-{cl}{blok['asli']}"

            for c_i in range(1, len(headers) + 1):
                ws.cell(row=blok["rekon"], column=c_i).fill = fill_rekon
                ws.cell(row=blok["rekon"], column=c_i).font = font_bold
                ws.cell(row=blok["rekon"], column=c_i).border = border_top_thin

                ws.cell(row=blok["asli"], column=c_i).fill = fill_asli
                ws.cell(row=blok["asli"], column=c_i).font = font_bold

                ws.cell(row=blok["ctrl"], column=c_i).fill = fill_ctrl
                ws.cell(row=blok["ctrl"], column=c_i).font = font_bold
                ws.cell(row=blok["ctrl"], column=c_i).border = border_bottom_double

        # Tulis rumus Grand Total Tahunan
        for k_col in kolom_hitung:
            if k_col in col_idx:
                ci, cl = col_idx[k_col], col_letter[k_col]
                daftar_sel_rekon = "+".join([f"{cl}{b['rekon']}" for b in info_blok_bulan])
                ws.cell(row=row_grand_rekon, column=ci).value = f"={daftar_sel_rekon}"
                ws.cell(row=row_grand_ctrl, column=ci).value = f"={cl}{row_grand_rekon}-{cl}{row_grand_asli}"

        for r_g in [row_grand_rekon, row_grand_asli, row_grand_ctrl]:
            for c_i in range(1, len(headers) + 1):
                ws.cell(row=r_g, column=c_i).fill = fill_grand
                ws.cell(row=r_g, column=c_i).font = font_bold
                if r_g == row_grand_ctrl:
                    ws.cell(row=r_g, column=c_i).border = border_bottom_double

        # Format Accounting & Auto-Fit Lebar Kolom untuk SEMUA Sheet
        acc_format = '#,##0.00;(#,##0.00);"-"'
        for s_name in writer.sheets:
            s_ws = writer.sheets[s_name]
            s_headers = [c.value for c in s_ws[1]]
            for idx_h, h_name in enumerate(s_headers, start=1):
                if h_name in kolom_angka:
                    for r_i in range(2, s_ws.max_row + 1):
                        s_ws.cell(row=r_i, column=idx_h).number_format = acc_format

                max_len = max((len(str(s_ws.cell(row=r_i, column=idx_h).value or "")) for r_i in range(1, min(s_ws.max_row + 1, 200))), default=10)
                s_ws.column_dimensions[get_column_letter(idx_h)].width = max(min(max_len + 5, 55), 16)

    return output_buffer.getvalue()

# ============================================================
# CSS ENTERPRISE WORKSPACE (CALM & MODERN UI)
# ============================================================
st.markdown(
    """
    <style>
    .stApp {
        background-color: #F4F7FA;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #FFFFFF;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.03);
        padding: 0.5rem;
    }
    div.stButton > button[kind="primary"] {
        background-color: #1E3A8A;
        border-color: #1E3A8A;
        color: #FFFFFF;
        font-weight: 600;
        border-radius: 8px;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #1D4ED8;
        border-color: #1D4ED8;
    }
    </style>
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

# ============================================================
# KARTU PANEL UNGGAH & PILIHAN BANK
# ============================================================
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
if "konfirmasi_lanjut_beda_bulan" not in st.session_state:
    st.session_state.konfirmasi_lanjut_beda_bulan = False

if uploaded_pdfs and uploaded_book and df_book_raw is not None:
    label_tombol = f"🚀 Proses Rekonsiliasi {pilihan_bank} ({len(uploaded_pdfs)} File PDF)" if len(uploaded_pdfs) > 1 else f"🚀 Proses Rekonsiliasi {pilihan_bank} Sekarang"
    klik_proses = st.button(label_tombol, type="primary", use_container_width=True)

    if klik_proses:
        st.session_state.konfirmasi_lanjut_beda_bulan = False

    if klik_proses or st.session_state.konfirmasi_lanjut_beda_bulan:
        try:
            df_book_normalized = normalisasi_excel_pembukuan(df_book_raw, mapping_kolom)

            with st.spinner(f"Memeriksa kelengkapan dan mengekstrak {len(uploaded_pdfs)} PDF {pilihan_bank}..."):
                hasil_ekstraksi_list = []
                for pdf_item in uploaded_pdfs:
                    if pilihan_bank == "BNI Direct":
                        df_p, s_awal, s_akhir, err_p = ekstrak_pdf_bni(pdf_item)
                    else:
                        df_p, s_awal, s_akhir, err_p = ekstrak_pdf_mandiri(pdf_item)

                    df_p["Sumber_File"] = pdf_item.name
                    tgl_awal_file = pd.to_datetime(df_p["Tanggal"]).min()
                    hasil_ekstraksi_list.append({
                        "nama": pdf_item.name,
                        "tgl_awal": tgl_awal_file,
                        "df": df_p,
                        "saldo_awal": s_awal,
                        "saldo_akhir": s_akhir,
                        "errors": err_p
                    })

                hasil_ekstraksi_list.sort(key=lambda x: x["tgl_awal"])
                df_bank = pd.concat([x["df"] for x in hasil_ekstraksi_list], ignore_index=True)
                saldo_awal = hasil_ekstraksi_list[0]["saldo_awal"]
                saldo_akhir = hasil_ekstraksi_list[-1]["saldo_akhir"]
                pdf_errors = [err for x in hasil_ekstraksi_list for err in x["errors"]]

                bulan_bank, bulan_gl, hilang_di_gl, hilang_di_bank = cek_kelengkapan_bulan(df_bank, df_book_normalized)

            # Tahan proses & tawarkan pilihan jika ada bulan yang tidak cocok antara PDF vs GL
            ada_beda_bulan = bool(hilang_di_gl or hilang_di_bank)
            if ada_beda_bulan and not st.session_state.konfirmasi_lanjut_beda_bulan:
                st.warning("⚠️ **DITEMUKAN PERBEDAAN KELENGKAPAN BULAN ANTARA PDF BANK DAN GENERAL LEDGER!**")
                if hilang_di_gl:
                    st.write(f"- **PDF Rekening Koran** memuat bulan `{', '.join(bulan_bank)}`, tetapi **General Ledger tidak memiliki transaksi** pada bulan: **`{', '.join(hilang_di_gl)}`**.")
                if hilang_di_bank:
                    st.write(f"- **General Ledger** memuat bulan `{', '.join(bulan_gl)}`, tetapi **PDF Rekening Koran belum diunggah** untuk bulan: **`{', '.join(hilang_di_bank)}`**.")

                st.info("Apakah Anda ingin melengkapi file unggahan terlebih dahulu, atau tetap melanjutkan proses rekonsiliasi dengan data yang ada?")
                c_btn1, c_btn2 = st.columns(2)
                with c_btn1:
                    if st.button("🔄 Lengkapi File Dulu (Batal Proses)", use_container_width=True):
                        st.session_state.konfirmasi_lanjut_beda_bulan = False
                        st.rerun()
                with c_btn2:
                    if st.button("▶️ Tetap Lanjutkan Proses Rekonsiliasi", type="primary", use_container_width=True):
                        st.session_state.konfirmasi_lanjut_beda_bulan = True
                        st.rerun()
                st.stop()

            # Jalankan Rekonsiliasi Penuh
            with st.spinner("Sedang mencocokkan transaksi dan menyusun laporan audit per bulan..."):
                if pilihan_bank == "BNI Direct":
                    df_export, df_laporan_final, pesan_audit = proses_rekonsiliasi_bni(df_bank.copy(), df_book_normalized)
                else:
                    df_export, df_laporan_final, pesan_audit = proses_rekonsiliasi_mandiri(df_bank.copy(), df_book_normalized)

                df_export, df_laporan_final, pesan_audit = perhalus_label_universal(df_export, df_laporan_final, pesan_audit)
                df_export, df_laporan_final = tambah_kolom_selisih(df_export, df_laporan_final)
                df_rekap_bulan = buat_rekap_per_bulan(df_laporan_final)

                excel_bytes = buat_excel_dengan_kontrol_bulanan(
                    df_laporan_final, df_bank, df_book_normalized, df_export, df_rekap_bulan, pilihan_bank
                )

            st.divider()

            if len(pdf_errors) == 0:
                st.success(f"**Ekstraksi {len(uploaded_pdfs)} PDF {pilihan_bank} Valid 100%!** ({len(df_bank)} transaksi | Saldo Awal: Rp {saldo_awal:,.2f} → Saldo Akhir: Rp {saldo_akhir:,.2f})")
            else:
                st.warning(f"Ditemukan {len(pdf_errors)} baris selisih pada pembacaan saldo PDF.")

            if "OKE 100%" in pesan_audit:
                st.success(f"**Status Audit Rekonsiliasi:** {pesan_audit}")
            else:
                st.error(f"**Status Audit Rekonsiliasi:** {pesan_audit}")

            m1, m2, m3, m4 = st.columns(4)
            jml_match = df_laporan_final['Status_Rekonsiliasi'].str.startswith('MATCH').sum()
            jml_bank_only = df_laporan_final['Status_Rekonsiliasi'].str.startswith('HANYA DI BANK').sum()
            jml_book_only = df_laporan_final['Status_Rekonsiliasi'].str.startswith('HANYA DI GENERAL LEDGER').sum()

            m1.metric("Total Baris Laporan", len(df_laporan_final))
            m2.metric("Cocok (MATCH)", jml_match)
            m3.metric("Hanya di Bank", jml_bank_only)
            m4.metric("Hanya di General Ledger", jml_book_only)

            # Format Accounting untuk tampilan tabel di layar Streamlit
            fmt_uang = {
                "Debet Bank": "{:,.2f}", "Kredit Bank": "{:,.2f}",
                "Debit GL": "{:,.2f}", "Credit GL": "{:,.2f}",
                "Debet": "{:,.2f}", "Kredit": "{:,.2f}", "Saldo": "{:,.2f}",
                "Debit": "{:,.2f}", "Credit": "{:,.2f}",
                "Selisih Bank dengan GL": "{:,.2f}"
            }

            tab_rekap, tab1, tab2 = st.tabs([
                "📅 Rekap Selisih Per Bulan",
                "📊 Hasil Rekonsiliasi Lengkap",
                f"📄 Hasil Ekstraksi PDF {pilihan_bank}"
            ])

            with tab_rekap:
                st.caption("Ringkasan total mutasi dan selisih per bulan (Format Accounting):")
                fmt_rekap = {k: v for k, v in fmt_uang.items() if k in df_rekap_bulan.columns}
                st.dataframe(df_rekap_bulan.style.format(fmt_rekap), use_container_width=True)

            with tab1:
                f_col1, f_col2 = st.columns(2)
                with f_col1:
                    opsi_bulan = sorted(df_laporan_final['Tanggal_Clean'].astype(str).str[:7].unique().tolist())
                    filter_bulan = st.multiselect("Filter Bulan:", options=opsi_bulan, default=opsi_bulan)
                with f_col2:
                    filter_status = st.multiselect(
                        "Filter Status Rekonsiliasi:",
                        options=df_laporan_final['Status_Rekonsiliasi'].unique().tolist(),
                        default=df_laporan_final['Status_Rekonsiliasi'].unique().tolist()
                    )
                df_tampil = df_laporan_final[
                    (df_laporan_final['Tanggal_Clean'].astype(str).str[:7].isin(filter_bulan))
                    & (df_laporan_final['Status_Rekonsiliasi'].isin(filter_status))
                ]
                                    def warnai_baris_status(row):
                        status = str(row.get("Status_Rekonsiliasi", ""))
                        if status.startswith("MATCH"):
                            return ["background-color: #F0FDF4"] * len(row)
                        elif status.startswith(("NETTING JAGIR", "KONTRA BANK")):
                            return ["background-color: #EFF6FF"] * len(row)
                        elif status.startswith(("HANYA DI", "SELISIH")):
                            return ["background-color: #FEF2F2"] * len(row)
                        return [""] * len(row)

                    fmt_lap = {k: v for k, v in fmt_uang.items() if k in df_tampil.columns}
                    st.dataframe(
                        df_tampil.style.format(fmt_lap).apply(warnai_baris_status, axis=1),
                        use_container_width=True,
                        height=450
                    )


            with tab2:
                fmt_bank = {k: v for k, v in fmt_uang.items() if k in df_bank.columns}
                st.dataframe(df_bank.style.format(fmt_bank), use_container_width=True, height=450)

            nama_prefix = f"{len(uploaded_pdfs)}_Bulan_{pilihan_bank.replace(' ', '_')}" if len(uploaded_pdfs) > 1 else uploaded_pdfs[0].name.replace('.pdf', '')
            nama_file_keluar = f"Hasil_Rekonsiliasi_{nama_prefix}.xlsx"
            st.download_button(
                label=f"📥 Download Hasil Rekonsiliasi {pilihan_bank} (.xlsx)",
                data=excel_bytes,
                file_name=nama_file_keluar,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses file: {e}")




