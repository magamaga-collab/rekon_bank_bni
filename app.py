import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==============================================================================
# BAGIAN 1: KONFIGURASI HALAMAN & GAYA TAMPILAN (CSS)
# ==============================================================================
st.set_page_config(
    page_title="Rekonsiliasi Bank BNI",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

def apply_custom_styles():
    st.markdown(
        """
        <style>
        .stApp {
            background-color: #F8FAFC;
        }
        [data-testid="InputInstructions"] {
            display: none !important;
        }
        .login-header {
            text-align: center;
            margin-bottom: 1.5rem;
        }
        .login-title {
            font-size: 1.5rem;
            font-weight: 700;
            color: #0F172A;
            margin-top: 0.75rem;
            margin-bottom: 0.25rem;
        }
        .login-subtitle {
            font-size: 0.875rem;
            color: #64748B;
        }
        div.stButton > button[kind="primary"],
        div[data-testid="stFormSubmitButton"] > button[kind="primary"] {
            background: linear-gradient(135deg, #005E6A 0%, #007A8A 100%);
            border: none;
            border-radius: 10px;
            padding: 0.6rem 1.2rem;
            font-weight: 600;
            color: white;
            transition: all 0.2s ease;
        }
        div.stButton > button[kind="primary"]:hover,
        div[data-testid="stFormSubmitButton"] > button[kind="primary"]:hover {
            box-shadow: 0 4px 12px rgba(0, 94, 106, 0.25);
            transform: translateY(-1px);
        }
        .metric-card {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 1.1rem 1.25rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.03);
        }
        .metric-label {
            font-size: 0.8rem;
            color: #64748B;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }
        .metric-value {
            font-size: 1.35rem;
            font-weight: 700;
            color: #0F172A;
            margin-top: 0.25rem;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

apply_custom_styles()
# ==============================================================================
# BAGIAN 2: SISTEM AUTENTIKASI (LOGIN)
# ==============================================================================
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.markdown("<div style='height: 3.5rem;'></div>", unsafe_allow_html=True)
    _, col_center, _ = st.columns([1, 1.35, 1])

    with col_center:
        with st.container(border=True):
            st.markdown(
                """
                <div class="login-header">
                    <div style="font-size: 2.5rem;">🔐</div>
                    <div class="login-title">Akses Terbatas - Rekon Bank BNI</div>
                    <div class="login-subtitle">Silakan masukkan kata sandi untuk mengakses sistem rekonsiliasi</div>
                </div>
                """,
                unsafe_allow_html=True
            )

            with st.form("login_form", clear_on_submit=False):
                pwd = st.text_input(
                    "Password Akses",
                    type="password",
                    placeholder="Masukkan kata sandi tim...",
                    autocomplete="one-time-code"
                )
                st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
                submitted = st.form_submit_button(
                    "🔓 Masuk ke Aplikasi",
                    type="primary",
                    use_container_width=True
                )

            if submitted:
                app_pwd = st.secrets.get("APP_PASSWORD", "admin123")
                if pwd == app_pwd:
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Kata sandi salah. Silakan coba lagi.")

            st.markdown(
                """
                <div style="text-align: center; margin-top: 1.2rem; padding-top: 0.8rem; border-top: 1px solid #EEF2F0; font-size: 0.75rem; color: #8A9990;">
                    🔒 Akses Terbatas • Pemrosesan Aman di Memori
                </div>
                """,
                unsafe_allow_html=True
            )

    return False

if not check_password():
    st.stop()
# ==============================================================================
# BAGIAN 3: FUNGSI PARSING & PEMBERSIHAN DATA
# ==============================================================================
def clean_currency(val):
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s in ["", "-", "nan", "None"]:
        return 0.0
    s = re.sub(r"[^\d,.-]", "", s)
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts[-1]) == 2:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0

def normalize_date(val):
    if pd.isna(val) or val is None:
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip().split(" ")
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return s

def read_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, sep=";")
    else:
        return pd.read_excel(uploaded_file)
# ==============================================================================
# BAGIAN 4: DETEKSI KOLOM OTOMATIS (REKENING KORAN BNI & BUKU BESAR)
# ==============================================================================
def detect_columns(df, source_type="bank"):
    cols = list(df.columns)
    lower_cols = {c: str(c).lower().strip() for c in cols}

    date_col = None
    desc_col = None
    debit_col = None
    credit_col = None
    amount_col = None
    type_col = None

    for c, lc in lower_cols.items():
        if not date_col and any(k in lc for k in ["tgl", "tanggal", "date", "post date", "trx date"]):
            date_col = c
        if not desc_col and any(k in lc for k in ["uraian", "keterangan", "deskripsi", "description", "remark", "narrative", "memo"]):
            desc_col = c
        if not debit_col and any(k in lc for k in ["debit", "debet", "keluar", "dr"]):
            debit_col = c
        if not credit_col and any(k in lc for k in ["kredit", "credit", "masuk", "cr"]):
            credit_col = c
        if not amount_col and any(k in lc for k in ["nominal", "jumlah", "amount", "mutasi", "nilai"]):
            amount_col = c
        if not type_col and any(k in lc for k in ["db/cr", "d/c", "tipe", "type", "jenis", "dk"]):
            type_col = c

    return {
        "date": date_col or cols,
        "desc": desc_col or (cols if len(cols) > 1 else cols),
        "debit": debit_col,
        "credit": credit_col,
        "amount": amount_col,
        "type": type_col
    }
# ==============================================================================
# BAGIAN 5: STANDARISASI TRANSAKSI
# ==============================================================================
def standardize_df(df, mapping, source_label):
    records = []
    for idx, row in df.iterrows():
        tgl = normalize_date(row.get(mapping["date"]))
        if not tgl or str(tgl).lower() in ["nan", "none", ""]:
            continue

        desc = str(row.get(mapping["desc"], "")).strip()
        if desc.lower() in ["nan", "none"]:
            desc = ""

        debit = 0.0
        credit = 0.0

        if mapping["debit"] and mapping["credit"]:
            debit = abs(clean_currency(row.get(mapping["debit"])))
            credit = abs(clean_currency(row.get(mapping["credit"])))
        elif mapping["amount"]:
            amt = clean_currency(row.get(mapping["amount"]))
            if mapping["type"]:
                t_val = str(row.get(mapping["type"], "")).upper().strip()
                if t_val in ["D", "DB", "DR", "DEBIT", "DEBET", "KELUAR"]:
                    debit = abs(amt)
                elif t_val in ["C", "CR", "K", "KREDIT", "CREDIT", "MASUK"]:
                    credit = abs(amt)
                else:
                    if amt < 0:
                        debit = abs(amt)
                    else:
                        credit = abs(amt)
            else:
                if amt < 0:
                    debit = abs(amt)
                else:
                    credit = abs(amt)

        if debit == 0 and credit == 0:
            continue

        records.append({
            "ID": f"{source_label}-{idx+1}",
            "Tanggal": tgl,
            "Keterangan": desc,
            "Debit": round(debit, 2),
            "Kredit": round(credit, 2),
            "Netto": round(credit - debit, 2)
        })

    return pd.DataFrame(records)
# ==============================================================================
# BAGIAN 6: MESIN REKONSILIASI OTOMATIS (EXACT & TOLERANSI TANGGAL)
# ==============================================================================
def run_reconciliation(df_bank, df_book, date_tolerance=2):
    bank_used = set()
    book_used = set()
    matched_rows = []

    bank_list = df_bank.to_dict("records")
    book_list = df_book.to_dict("records")

    # Tahap 1: Cocok Tanggal Sama Persis & Nominal Sama Persis
    for i, b in enumerate(bank_list):
        if i in bank_used:
            continue
        for j, k in enumerate(book_list):
            if j in book_used:
                continue
            if b["Tanggal"] == k["Tanggal"] and b["Debit"] == k["Debit"] and b["Kredit"] == k["Kredit"]:
                bank_used.add(i)
                book_used.add(j)
                matched_rows.append({
                    "Status": "Cocok (Exact)",
                    "Tgl Bank": b["Tanggal"],
                    "Ket Bank": b["Keterangan"],
                    "Debit Bank": b["Debit"],
                    "Kredit Bank": b["Kredit"],
                    "Tgl Buku": k["Tanggal"],
                    "Ket Buku": k["Keterangan"],
                    "Debit Buku": k["Debit"],
                    "Kredit Buku": k["Kredit"],
                    "Selisih": 0.0
                })
                break

    # Tahap 2: Cocok Toleransi Tanggal (± N hari) & Nominal Sama Persis
    for i, b in enumerate(bank_list):
        if i in bank_used:
            continue
        try:
            d_bank = datetime.strptime(b["Tanggal"], "%Y-%m-%d")
        except Exception:
            continue

        for j, k in enumerate(book_list):
            if j in book_used:
                continue
            try:
                d_book = datetime.strptime(k["Tanggal"], "%Y-%m-%d")
            except Exception:
                continue

            diff_days = abs((d_bank - d_book).days)
            if diff_days <= date_tolerance and b["Debit"] == k["Debit"] and b["Kredit"] == k["Kredit"]:
                bank_used.add(i)
                book_used.add(j)
                matched_rows.append({
                    "Status": f"Cocok (Beda {diff_days} Hari)",
                    "Tgl Bank": b["Tanggal"],
                    "Ket Bank": b["Keterangan"],
                    "Debit Bank": b["Debit"],
                    "Kredit Bank": b["Kredit"],
                    "Tgl Buku": k["Tanggal"],
                    "Ket Buku": k["Keterangan"],
                    "Debit Buku": k["Debit"],
                    "Kredit Buku": k["Kredit"],
                    "Selisih": 0.0
                })
                break

    unmatched_bank = [b for i, b in enumerate(bank_list) if i not in bank_used]
    unmatched_book = [k for j, k in enumerate(book_list) if j not in book_used]

    return (
        pd.DataFrame(matched_rows),
        pd.DataFrame(unmatched_bank),
        pd.DataFrame(unmatched_book)
    )
# ==============================================================================
# BAGIAN 7: EKSPOR HASIL KE EXCEL RAPI
# ==============================================================================
def export_to_excel(df_matched, df_unmatched_bank, df_unmatched_book):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_matched.to_excel(writer, sheet_name="1_Transaksi_Cocok", index=False)
        df_unmatched_bank.to_excel(writer, sheet_name="2_Hanya_di_Bank_BNI", index=False)
        df_unmatched_book.to_excel(writer, sheet_name="3_Hanya_di_Buku_Besar", index=False)

        wb = writer.book
        header_fill = PatternFill(start_color="005E6A", end_color="005E6A", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        thin_border = Border(
            left=Side(style="thin", color="D1D5DB"),
            right=Side(style="thin", color="D1D5DB"),
            top=Side(style="thin", color="D1D5DB"),
            bottom=Side(style="thin", color="D1D5DB")
        )

        for ws in wb.worksheets:
            for cell in ws:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.border = thin_border
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0.00"

            for col in ws.columns:
                max_len = 0
                col_letter = col.column_letter
                for cell in col:
                    val_len = len(str(cell.value or ""))
                    if val_len > max_len:
                        max_len = val_len
                ws.column_dimensions[col_letter].width = min(max_len + 4, 45)

    return output.getvalue()
# ==============================================================================
# BAGIAN 8: ANTARMUKA UTAMA (DASHBOARD REKONSILIASI BNI)
# ==============================================================================
with st.sidebar:
    st.markdown("### 🏦 Rekon Bank BNI")
    st.caption("Sistem Rekonsiliasi Rekening Koran & Buku Besar")
    st.divider()

    tolerance_days = st.slider(
        "Toleransi Selisih Tanggal (Hari)",
        min_value=0,
        max_value=7,
        value=2,
        help="Mengizinkan pencocokan otomatis jika tanggal bank dan pembukuan berbeda beberapa hari."
    )

    st.divider()
    if st.button("🚪 Keluar / Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

st.title("🏦 Rekonsiliasi Rekening Koran BNI")
st.write("Unggah file **Rekening Koran BNI** dan **Buku Besar / Catatan Internal** dalam format Excel (`.xlsx`, `.xls`) atau `.csv`.")

col_up1, col_up2 = st.columns(2)
with col_up1:
    file_bank = st.file_uploader("1️⃣ Unggah Rekening Koran Bank BNI", type=["xlsx", "xls", "csv"], key="bank")
with col_up2:
    file_book = st.file_uploader("2️⃣ Unggah Buku Besar / Catatan Internal", type=["xlsx", "xls", "csv"], key="book")

if file_bank and file_book:
    try:
        raw_bank = read_uploaded_file(file_bank)
        raw_book = read_uploaded_file(file_book)

        map_bank = detect_columns(raw_bank, "bank")
        map_book = detect_columns(raw_book, "book")

        with st.expander("⚙️ Pengaturan Pemetaan Kolom (Opsional - Sudah Terdeteksi Otomatis)"):
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Kolom Rekening Koran BNI**")
                bank_cols = [None] + list(raw_bank.columns)
                map_bank["date"] = st.selectbox("Tanggal (Bank)", list(raw_bank.columns), index=list(raw_bank.columns).index(map_bank["date"]))
                map_bank["desc"] = st.selectbox("Keterangan (Bank)", list(raw_bank.columns), index=list(raw_bank.columns).index(map_bank["desc"]))
                map_bank["debit"] = st.selectbox("Debit / Keluar (Bank)", bank_cols, index=bank_cols.index(map_bank["debit"]) if map_bank["debit"] in bank_cols else 0)
                map_bank["credit"] = st.selectbox("Kredit / Masuk (Bank)", bank_cols, index=bank_cols.index(map_bank["credit"]) if map_bank["credit"] in bank_cols else 0)
                map_bank["amount"] = st.selectbox("Nominal Tunggal (Jika tidak terpisah)", bank_cols, index=bank_cols.index(map_bank["amount"]) if map_bank["amount"] in bank_cols else 0)
                map_bank["type"] = st.selectbox("Tipe DB/CR (Jika nominal tunggal)", bank_cols, index=bank_cols.index(map_bank["type"]) if map_bank["type"] in bank_cols else 0)

            with c2:
                st.markdown("**Kolom Buku Besar Internal**")
                book_cols = [None] + list(raw_book.columns)
                map_book["date"] = st.selectbox("Tanggal (Buku)", list(raw_book.columns), index=list(raw_book.columns).index(map_book["date"]))
                map_book["desc"] = st.selectbox("Keterangan (Buku)", list(raw_book.columns), index=list(raw_book.columns).index(map_book["desc"]))
                map_book["debit"] = st.selectbox("Debit / Masuk (Buku)", book_cols, index=book_cols.index(map_book["debit"]) if map_book["debit"] in book_cols else 0)
                map_book["credit"] = st.selectbox("Kredit / Keluar (Buku)", book_cols, index=book_cols.index(map_book["credit"]) if map_book["credit"] in book_cols else 0)
                map_book["amount"] = st.selectbox("Nominal Tunggal (Buku)", book_cols, index=book_cols.index(map_book["amount"]) if map_book["amount"] in book_cols else 0)
                map_book["type"] = st.selectbox("Tipe DB/CR (Buku)", book_cols, index=book_cols.index(map_book["type"]) if map_book["type"] in book_cols else 0)

        df_bank_clean = standardize_df(raw_bank, map_bank, "BNI")
        df_book_clean = standardize_df(raw_book, map_book, "BUKU")

        if st.button("🚀 Jalankan Proses Rekonsiliasi", type="primary", use_container_width=True):
            df_matched, df_un_bank, df_un_book = run_reconciliation(
                df_bank_clean, df_book_clean, date_tolerance=tolerance_days
            )

            st.divider()
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.markdown(f"<div class='metric-card'><div class='metric-label'>Total Transaksi Bank</div><div class='metric-value'>{len(df_bank_clean)}</div></div>", unsafe_allow_html=True)
            with m2:
                st.markdown(f"<div class='metric-card'><div class='metric-label'>Berhasil Cocok</div><div class='metric-value' style='color:#059669;'>{len(df_matched)}</div></div>", unsafe_allow_html=True)
            with m3:
                st.markdown(f"<div class='metric-card'><div class='metric-label'>Selisih di Bank BNI</div><div class='metric-value' style='color:#DC2626;'>{len(df_un_bank)}</div></div>", unsafe_allow_html=True)
            with m4:
                st.markdown(f"<div class='metric-card'><div class='metric-label'>Selisih di Buku Besar</div><div class='metric-value' style='color:#D97706;'>{len(df_un_book)}</div></div>", unsafe_allow_html=True)

            st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)

            tab1, tab2, tab3 = st.tabs([
                f"✅ Cocok ({len(df_matched)})",
                f"🏦 Hanya di Bank BNI ({len(df_un_bank)})",
                f"📒 Hanya di Buku Besar ({len(df_un_book)})"
            ])

            with tab1:
                st.dataframe(df_matched, use_container_width=True, hide_index=True)
            with tab2:
                st.dataframe(df_un_bank, use_container_width=True, hide_index=True)
            with tab3:
                st.dataframe(df_un_book, use_container_width=True, hide_index=True)

            excel_bytes = export_to_excel(df_matched, df_un_bank, df_un_book)
            st.download_button(
                label="📥 Unduh Laporan Rekonsiliasi (.xlsx)",
                data=excel_bytes,
                file_name=f"Hasil_Rekon_BNI_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
