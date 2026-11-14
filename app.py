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
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# HALAMAN LOGIN PASSWORD (SELARAS TEMA SAGE GREEN ReconIt!)
# ============================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown(
        """
        <style>
        /* Sembunyikan sidebar & header bawaan di halaman login */
        [data-testid="stSidebar"], [data-testid="collapsedControl"], header[data-testid="stHeader"] {
            display: none !important;
        }
        /* Warna latar belakang lembut selaras halaman utama */
        .stApp {
            background-color: #F4F6F5;
        }
        /* Kartu Login Tengah */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background-color: #FFFFFF;
            border: 1px solid #DCE3DF;
            border-top: 5px solid #5B7B65;
            border-radius: 14px;
            box-shadow: 0 10px 30px rgba(45, 65, 53, 0.08);
            padding: 2rem 2.2rem 1.8rem 2.2rem;
            margin-top: 4vh;
        }
        /* Input Password */
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
        /* Tombol Masuk Hijau Sage */
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
                        ReconIt! • Faster 'n Easier!
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
# KOSMETIK HALAMAN UTAMA (TOP BAR HIJAU + STEPPER + CARD)
# ============================================================
st.markdown(
    """
    <style>
    /* 1. Top Bar Hijau Tua Penuh di Paling Atas (Tinggi +1 cm) */
    header[data-testid="stHeader"], .stAppHeader, .stAppToolbar, [data-testid="stDecoration"] {
        background-color: #466650 !important;
        background-image: none !important;
        height: 4.2rem !important;
        min-height: 4.2rem !important;
        border-bottom: 1px solid #395441 !important;
        display: block !important;
        visibility: visible !important;
    }
    header[data-testid="stHeader"] *, .stAppToolbar * {
        color: #FFFFFF !important;
    }
    /* Dorong konten utama & sidebar sedikit ke bawah agar pas dengan tinggi header baru */
    .block-container {
        padding-top: 5.2rem !important;
    }
    section[data-testid="stSidebar"] > div:first-child {
        padding-top: 4.5rem !important;
    }

    /* 2. Latar Belakang Halaman Utama */
    .stApp {
        background-color: #F2F5F3;
    }

    /* 3. Kartu Kontainer Putih Rapi */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #FFFFFF;
        border: 1px solid #DCE3DF !important;
        border-radius: 12px !important;
        box-shadow: 0 3px 12px rgba(45, 65, 53, 0.05);
        padding: 1rem 1.25rem;
    }

    /* 4. Dropzone Upload Kiri (Aksen Hijau Lembut) */
    div[data-testid="stColumn"]:nth-of-type(1) [data-testid="stFileUploaderDropzone"] {
        background-color: #EBF3EE !important;
        border: 1.5px dashed #7EA189 !important;
        border-radius: 10px !important;
        padding: 1.4rem 1rem !important;
    }

    /* 5. Dropzone Upload Kanan (Aksen Abu-Abu Lembut) */
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
        jml_unmatched = grp[~grp["Status_Rekonsiliasi"].str.startswith(("MATCH", "NETTING JAGIR", "KONTRA BANK"))].shape[0]

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
# CSS Faster 'n Easier! (SAGE GREEN & CALM UI)
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
            <div style="font-size: 1.35rem; font-weight: 700; letter-spacing: 0.3px;">🏦 ReconIt!</div>
            <div style="font-size: 0.78rem; opacity: 0.85;">Reconcile Bank vs GL • by FERDX</div>
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
                <div style="font-weight: 600; font-size: 0.92rem;">github</div>
                <div style="font-size: 0.75rem; opacity: 0.85;">magamaga-collab</div>
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

    
    
stepper_slot = st.empty()

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

    # Hitung status progress unggahan secara dinamis (0%, 50%, atau 100%)
    ada_pdf = bool(uploaded_pdfs)
    ada_excel = uploaded_book is not None

    if ada_pdf and ada_excel:
        persen_bar = "100%"
        bg_kiri, brd_kiri = "#D4E8DA", "#5B7B65"
        bg_kanan, brd_kanan = "#D4E8DA", "#5B7B65"
    elif ada_pdf or ada_excel:
        persen_bar = "50%"
        bg_kiri = "#D4E8DA" if ada_pdf else "#ECEFED"
        brd_kiri = "#5B7B65" if ada_pdf else "#CBD3CE"
        bg_kanan = "#D4E8DA" if ada_excel else "#ECEFED"
        brd_kanan = "#5B7B65" if ada_excel else "#CBD3CE"
    else:
        persen_bar = "0%"
        bg_kiri, brd_kiri = "#ECEFED", "#CBD3CE"
        bg_kanan, brd_kanan = "#ECEFED", "#CBD3CE"

    stepper_slot.markdown(
        f"""
        <div style="background-color: #FFFFFF; border: 1px solid #DCE3DF; border-radius: 12px;
                    padding: 12px 28px; margin: 12px 0 16px 0;
                    box-shadow: 0 3px 10px rgba(45, 65, 53, 0.04);
                    display: flex; align-items: center; justify-content: space-between;">
            <div style="display: flex; align-items: center; justify-content: center;
                        width: 34px; height: 34px; border-radius: 8px;
                        background-color: {bg_kiri}; border: 1.5px solid {brd_kiri}; font-size: 16px;">
                📄
            </div>
            <div style="flex-grow: 1; height: 4px; background: linear-gradient(90deg, #5B7B65 {persen_bar}, #DCE3DF {persen_bar});
                        margin: 0 16px; border-radius: 3px; transition: all 0.3s ease;"></div>
            <div style="display: flex; align-items: center; justify-content: center;
                        width: 34px; height: 34px; border-radius: 8px;
                        background-color: {bg_kanan}; border: 1.5px solid {brd_kanan}; font-size: 16px;">
                📊
            </div>
        </div>
        """,
        unsafe_allow_html=True
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




