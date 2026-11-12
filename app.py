import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime
from collections import Counter
from itertools import combinations

# ============================================================
# KONFIGURASI HALAMAN WEB
# ============================================================
st.set_page_config(
    page_title="Rekonsiliasi BNI Direct vs Odoo",
    page_icon="🏦",
    layout="wide"
)
# ============================================================
# HALAMAN LOGIN PASSWORD
# ============================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 Login Aplikasi Rekonsiliasi BNI")
    pwd = st.text_input("Masukkan Password Akses:", type="password")
    if st.button("Masuk", type="primary"):
        if pwd == st.secrets["APP_PASSWORD"]:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Password salah! Silakan coba lagi.")
    st.stop()
st.title("🏦 Aplikasi Rekonsiliasi Rekening Koran BNI vs Odoo")
st.markdown("Unggah **PDF Rekening Koran BNI Direct** dan **Excel Mutasi Odoo** untuk memproses ekstraksi dan rekonsiliasi otomatis.")

# ============================================================
# FUNGSI EKSTRAKSI PDF BNI (SESUAI VERSI TESTED)
# ============================================================
def clean_double_printed(text):
    s = str(text).strip().replace(" ", "")
    if len(s) % 2 == 0 and len(s) > 0:
        pairs = [s[i:i+2] for i in range(0, len(s), 2)]
        if all(len(p) == 2 and p[0] == p[1] for p in pairs):
            return "".join(p[0] for p in pairs)

    s = re.sub(r'\.\.+', '.', s)
    if '.' in s:
        bagian_kiri, bagian_kanan = s.split('.', 1)
        if len(bagian_kanan) == 4 and bagian_kanan[0] == bagian_kanan[1] and bagian_kanan[2] == bagian_kanan[3]:
            bagian_kanan = bagian_kanan[0] + bagian_kanan[2]
        if ',' in bagian_kiri:
            grup_koma = bagian_kiri.split(',')
            grup_bersih = []
            for idx_g, g in enumerate(grup_koma):
                if idx_g > 0 and len(g) == 6 and all(g[j] == g[j+1] for j in range(0, 6, 2)):
                    grup_bersih.append(g[::2])
                elif idx_g == 0 and len(g) >= 2 and g[0] == g[1] and len(grup_koma) > 1 and len(g) > 3:
                    grup_bersih.append(g[1:])
                else:
                    grup_bersih.append(g)
            bagian_kiri = ",".join(grup_bersih)
        s = f"{bagian_kiri}.{bagian_kanan}"
    return s

def parse_money(text):
    if text is None:
        return None
    s = clean_double_printed(text)
    s = re.sub(r"[^0-9.\-]", "", s)
    if not re.fullmatch(r"-?[0-9]+\.[0-9]{2}", s):
        return None
    try:
        return float(s)
    except ValueError:
        return None
def get_opening_balance(pdf):
    text = pdf.pages[0].extract_text() or ""
    match = re.search(r"Ledger\s+Balance\s*:\s*([\d,]+\.\d{2})", text, re.IGNORECASE)
    if not match:
        raise ValueError("Ledger Balance awal tidak ditemukan di halaman 1.")
    balance = parse_money(match.group(1))
    if balance is None:
        raise ValueError("Nominal Ledger Balance tidak valid.")
    return balance

def extract_page_words(page):
    return page.extract_words(x_tolerance=1, y_tolerance=2, keep_blank_chars=False, use_text_flow=False)

def find_transaction_dates(words):
    dates = []
    for w in words:
        if w["x0"] < 100 and re.fullmatch(r"\d{2}/\d{2}/\d{4}", w["text"]):
            dates.append(w)
    return sorted(dates, key=lambda x: x["top"])

def get_amount(words, date_y, dbcr):
    candidates = []
    for w in words:
        if not (430 <= w["x0"] < 620):
            continue
        value = parse_money(w["text"])
        if value is None or abs(w["top"] - date_y) > 90:
            continue
        candidates.append({"value": value, "y": w["top"], "distance": abs(w["top"] - date_y)})
    if not candidates:
        return None
    if dbcr == "K":
        nonzero = [c for c in candidates if c["value"] != 0]
        if nonzero:
            return min(nonzero, key=lambda c: c["distance"])["value"]
        return 0.0
    return min(candidates, key=lambda c: c["distance"])["value"]

def get_balance(words, date_y):
    candidates = []
    for w in words:
        if w["x0"] < 615 or abs(w["top"] - date_y) > 8:
            continue
        value = parse_money(w["text"])
        if value is not None:
            candidates.append({"value": value, "distance": abs(w["top"] - date_y)})
    if not candidates:
        return None
    return min(candidates, key=lambda c: c["distance"])["value"]

def get_dbcr(words, date_y):
    for w in words:
        if 580 <= w["x0"] < 620 and abs(w["top"] - date_y) <= 8:
            if w["text"].strip() in ("D", "K"):
                return w["text"].strip()
    return None

def get_journal(words, top, bottom):
    candidates = []
    for w in words:
        if (350 <= w["x0"] < 430) and (top <= w["top"] <= bottom) and re.fullmatch(r"\d{5,7}", w["text"]):
            candidates.append(w)
    return candidates if candidates else ""

def ekstrak_pdf_bni(pdf_file):
    transactions = []
    page_transaction_counts = Counter()

    with pdfplumber.open(pdf_file) as pdf:
        opening_balance = get_opening_balance(pdf)

        for page_number, page in enumerate(pdf.pages, start=1):
            words = extract_page_words(page)
            date_words = find_transaction_dates(words)
            if not date_words:
                continue

            centers = [w["top"] for w in date_words]
            for i, date_word in enumerate(date_words):
                date_y = date_word["top"]
                top_limit = 300 if i == 0 else (centers[i-1] + date_y) / 2
                bottom_limit = 850 if i == len(date_words) - 1 else (date_y + centers[i+1]) / 2

                row_words = [w for w in words if top_limit <= w["top"] < bottom_limit]
                tanggal = datetime.strptime(date_word["text"], "%d/%m/%Y").strftime("%Y-%m-%d")
                dbcr = get_dbcr(row_words, date_y)
                if dbcr not in ("D", "K"):
                    raise ValueError(f"DB/CR tidak ditemukan: Halaman {page_number}, tanggal {tanggal}")

                nominal = get_amount(row_words, date_y, dbcr)
                saldo = get_balance(row_words, date_y)
                prev_bal = transactions[-1]["Saldo"] if transactions else opening_balance

                if nominal is None:
                    if saldo is not None:
                        nominal = round(abs(saldo - prev_bal), 2)
                    else:
                        raise ValueError(f"Nominal dan Saldo tidak ditemukan: Halaman {page_number}, tanggal {tanggal}")

                if saldo is None:
                    saldo = round(prev_bal - (nominal if dbcr == "D" else 0.0) + (nominal if dbcr == "K" else 0.0), 2)

                prev_bal_check = transactions[-1]["Saldo"] if transactions else opening_balance
                expected_bal_check = round(prev_bal_check - (nominal if dbcr == "D" else 0.0) + (nominal if dbcr == "K" else 0.0), 2)
                if abs(expected_bal_check - saldo) > 0.05:
                    nominal = round(abs(saldo - prev_bal_check), 2)

                description_words = [w for w in row_words if 330 <= w["x0"] < 480]
                description_words.sort(key=lambda w: (round(w["top"], 1), w["x0"]))
                keterangan = re.sub(r"\s+", " ", " ".join(w["text"] for w in description_words)).strip()

                if "760360420803001" in keterangan and nominal <= 10000:
                    keterangan = "BIAYA TRANSFER BNI | " + keterangan

                debet = nominal if dbcr == "D" else 0.0
                kredit = nominal if dbcr == "K" else 0.0

                transactions.append({
                    "Tanggal": tanggal,
                    "Keterangan": keterangan,
                    "Debet": debet,
                    "Kredit": kredit,
                    "Saldo": saldo,
                    "Halaman": page_number
                })
                page_transaction_counts[page_number] += 1

    columns = ["Tanggal", "Keterangan", "Debet", "Kredit", "Saldo", "Halaman"]
    df = pd.DataFrame(transactions, columns=columns)
    for col in ["Debet", "Kredit", "Saldo"]:
        df[col] = pd.to_numeric(df[col], errors="raise")
    df["Halaman"] = df["Halaman"].astype(int)

    running_balance = opening_balance
    errors = []
    for index, row in df.iterrows():
        expected_balance = running_balance - row["Debet"] + row["Kredit"]
        pdf_balance = row["Saldo"]
        difference = round(expected_balance - pdf_balance, 2)
        if abs(difference) > 0.01:
            errors.append({
                "Baris": index + 1, "Tanggal": row["Tanggal"], "Halaman": row["Halaman"],
                "Keterangan": row["Keterangan"], "Saldo Hitungan": expected_balance,
                "Saldo PDF": pdf_balance, "Selisih": difference
            })
        running_balance = pdf_balance

    return df, opening_balance, running_balance, errors
# ============================================================
# FUNGSI REKONSILIASI BNI VS ODOO (SESUAI VERSI TESTED)
# ============================================================
def proses_rekonsiliasi(df_bank_raw, df_odoo_raw):
    df_odoo_raw['Tanggal_Clean'] = pd.to_datetime(df_odoo_raw['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    df_odoo_raw['Debit'] = pd.to_numeric(df_odoo_raw['Debit'], errors='coerce').fillna(0)
    df_odoo_raw['Credit'] = pd.to_numeric(df_odoo_raw['Credit'], errors='coerce').fillna(0)

    baris_header_odoo = df_odoo_raw[df_odoo_raw['Tanggal_Clean'].isna()]
    total_debit_bawaan = baris_header_odoo['Debit'].sum() if not baris_header_odoo.empty else df_odoo_raw['Debit'].sum()
    total_credit_bawaan = baris_header_odoo['Credit'].sum() if not baris_header_odoo.empty else df_odoo_raw['Credit'].sum()

    df_bank_raw['Tanggal_Clean'] = pd.to_datetime(df_bank_raw['Tanggal'], errors='coerce').dt.strftime('%Y-%m-%d')
    df_bank = df_bank_raw[df_bank_raw['Tanggal_Clean'].notna()].copy()
    df_odoo = df_odoo_raw[df_odoo_raw['Tanggal_Clean'].notna()].copy()

    df_bank['Debet'] = pd.to_numeric(df_bank['Debet'], errors='coerce').fillna(0)
    df_bank['Kredit'] = pd.to_numeric(df_bank['Kredit'], errors='coerce').fillna(0)

    total_debet_bank_asli = df_bank['Debet'].sum()
    total_kredit_bank_asli = df_bank['Kredit'].sum()

    df_bank['Keterangan_Upper'] = df_bank['Keterangan'].astype(str).str.upper()
    df_odoo['Label_Upper'] = df_odoo['Label'].astype(str).str.upper()

    # 1. Akumulasi Biaya Admin & Transfer BNI
    filter_bank_fee = (
        (
            df_bank['Keterangan_Upper'].str.contains('BIAYA TRANSFER BNI|BY TRX BIFAST|BY TRX ATM PRIMA|TRANSACTION FEE|BIAYA ADM|TRANSFER FEE|INHOUSETRF KE GNC')
            & (df_bank['Debet'] > 0) & (df_bank['Debet'] <= 249000)
        )

        | (df_bank['Debet'] == 600) | (df_bank['Debet'] == 2500) | (df_bank['Debet'] == 6500)
        | (df_bank['Keterangan_Upper'].str.contains('RTGS KE') & (df_bank['Debet'] == 30000))
    )
    total_fee_bank = df_bank[filter_bank_fee]['Debet'].sum()

    filter_odoo_fee = (
        df_odoo['Label_Upper'].str.contains('BY TRANSFER|BIAYA ADM|ADM MAN|ADM BNI')
        & ~df_odoo['Label_Upper'].str.contains('NET JAGIR|VENDOR PAYMENT')
        & (df_odoo['Credit'] > 0)
    )
    total_fee_odoo = df_odoo[filter_odoo_fee]['Credit'].sum()

    rekon_fee_rows = []
    if total_fee_bank > 0 or total_fee_odoo > 0:
        status_fee = 'MATCH (Total Biaya Admin Cocok)' if round(total_fee_bank, 2) == round(total_fee_odoo, 2) else 'SELISIH BIAYA ADMIN'
        rekon_fee_rows.append({
            'Tanggal_Clean': df_bank[filter_bank_fee]['Tanggal_Clean'].max() if not df_bank[filter_bank_fee].empty else df_odoo[filter_odoo_fee]['Tanggal_Clean'].max(),
            'Status_Rekonsiliasi': status_fee, 'Tipe': 'Keluar (Total Fee)',
            'Keterangan': f"Akumulasi Biaya Transfer & Admin Bank Koran (Total: Rp{total_fee_bank:,.2f})",
            'Debet': total_fee_bank, 'Kredit': 0.0,
            'Label': f"Total Catatan Admin di Odoo (Total: Rp{total_fee_odoo:,.2f})",
            'Credit': total_fee_odoo, 'Debit': 0.0
        })
    df_rekon_fee = pd.DataFrame(rekon_fee_rows)

    # 2. Net Jagir / Jasa Giro
    is_bunga_murni = (
        df_bank['Keterangan_Upper'].str.contains(r'JASA GIRO|\bBUNGA\b', regex=True)
        & ~df_bank['Keterangan_Upper'].str.contains('TABUNGAN|TRANSFER DARI|PAPAN|ANGGREK')
        & (df_bank['Kredit'] > 0)
    )
    is_pajak_murni = ((df_bank['Keterangan_Upper'].str.strip() == 'PAJAK') | (df_bank['Keterangan_Upper'].str.strip() == 'PPH')) & (df_bank['Debet'] > 0)
    filter_bunga_pajak_bank = is_bunga_murni | is_pajak_murni

    bunga_bank = df_bank[is_bunga_murni].copy()
    pajak_bank = df_bank[is_pajak_murni].copy()

    net_jagir_bank_rows = []
    val_pajak_total_jagir = 0.0
    if not bunga_bank.empty:
        for tgl in bunga_bank['Tanggal_Clean'].unique():
            val_bunga = bunga_bank[bunga_bank['Tanggal_Clean'] == tgl]['Kredit'].sum()
            val_pajak = pajak_bank[pajak_bank['Tanggal_Clean'] == tgl]['Debet'].sum() if not pajak_bank.empty else 0.0
            val_pajak_total_jagir += val_pajak
            net_bank = round(val_bunga - val_pajak, 2)
            if net_bank > 0:
                net_jagir_bank_rows.append({'Tanggal_Clean': tgl, 'Net_Jagir_Bank': net_bank, 'Bunga_Asli': val_bunga, 'Pajak_Asli': val_pajak})
    df_net_bank = pd.DataFrame(net_jagir_bank_rows)

    filter_odoo_jagir = df_odoo['Label_Upper'].str.contains('NET JAGIR')
    df_odoo_jagir = df_odoo[filter_odoo_jagir].copy()

    if not df_net_bank.empty and not df_odoo_jagir.empty:
        df_rekon_jagir = pd.merge(df_net_bank, df_odoo_jagir, on='Tanggal_Clean', how='outer')
        def cek_status_jagir(r):
            if pd.isna(r['Net_Jagir_Bank']) or pd.isna(r['Debit']): return 'SELISIH NET JAGIR'
            selisih = abs(r['Net_Jagir_Bank'] - r['Debit'])
            if selisih == 0: return 'MATCH (Net Jagir Cocok)'
            elif selisih <= 5: return f'MATCH (Net Jagir Cocok - Selisih Pembulatan Rp{selisih:,.2f})'
            else: return f'SELISIH NET JAGIR (Selisih Rp{selisih:,.2f})'
        df_rekon_jagir['Status_Rekonsiliasi'] = df_rekon_jagir.apply(cek_status_jagir, axis=1)
        df_rekon_jagir['Tipe'] = 'Masuk (Netto)'
        df_rekon_jagir['Keterangan'] = df_rekon_jagir.apply(lambda r: f"Jasa Giro/Bunga (Rp{r['Bunga_Asli']:,.2f}) & Pajak (Rp{r['Pajak_Asli']:,.2f}) di-Nettkan" if pd.notna(r['Bunga_Asli']) else '', axis=1)
        df_rekon_jagir['Kredit'] = df_rekon_jagir['Net_Jagir_Bank']
        df_rekon_jagir['Debet'] = 0.0
    elif not df_odoo_jagir.empty:
        df_rekon_jagir = df_odoo_jagir.copy()
        df_rekon_jagir['Status_Rekonsiliasi'] = 'HANYA DI ODOO (Net Jagir)'
        df_rekon_jagir['Tipe'] = 'Masuk (Netto)'
    else:
        df_rekon_jagir = pd.DataFrame()
    # 3. Transaksi Normal
    filter_bank_normal = ~filter_bank_fee & ~filter_bunga_pajak_bank
    df_bank_normal = df_bank[filter_bank_normal].copy().reset_index(drop=True)
    filter_odoo_normal = ~filter_odoo_fee & ~filter_odoo_jagir
    df_odoo_normal = df_odoo[filter_odoo_normal].copy().reset_index(drop=True)

    def rekon_lengkap_bni(df_b, df_o, col_b, col_o, tipe_str):
        b_sub = df_b[df_b[col_b] > 0].copy().reset_index(drop=True)
        o_sub = df_o[df_o[col_o] > 0].copy().reset_index(drop=True)
        b_sub['id_b'], o_sub['id_o'] = b_sub.index, o_sub.index
        b_sub['dt_b'], o_sub['dt_o'] = pd.to_datetime(b_sub['Tanggal_Clean']), pd.to_datetime(o_sub['Tanggal_Clean'])

        b_sub['seq'] = b_sub.groupby(['Tanggal_Clean', col_b]).cumcount()
        o_sub['seq'] = o_sub.groupby(['Tanggal_Clean', col_o]).cumcount()
        exact = pd.merge(b_sub, o_sub, left_on=['Tanggal_Clean', col_b, 'seq'], right_on=['Tanggal_Clean', col_o, 'seq'], how='inner')
        exact['Status_Rekonsiliasi'] = 'MATCH (Cocok Sempurna)'
        used_b, used_o = set(exact['id_b']), set(exact['id_o'])

        fuzzy_amt_rows = []
        for _, rb in b_sub[~b_sub['id_b'].isin(used_b)].iterrows():
            kandidat = o_sub[(o_sub['Tanggal_Clean'] == rb['Tanggal_Clean']) & (~o_sub['id_o'].isin(used_o))]
            for _, ro in kandidat.iterrows():
                selisih = abs(rb[col_b] - ro[col_o])
                is_valas = ('$' in str(ro.get('Label', ''))) or ('€' in str(ro.get('Label', '')))
                if selisih <= (500 if is_valas else 1):
                    used_b.add(rb['id_b']); used_o.add(ro['id_o'])
                    ket = f"Selisih Kurs Valas Rp{selisih:,.0f}" if is_valas else f"Selisih Pembulatan Rp{selisih:,.2f}"
                    gabung = {**rb.to_dict(), **ro.to_dict(), 'Status_Rekonsiliasi': f'MATCH (Cocok - {ket})'}
                    fuzzy_amt_rows.append(gabung); break

        fuzzy_date_rows = []
        for _, rb in b_sub[~b_sub['id_b'].isin(used_b)].iterrows():
            kandidat = o_sub[(~o_sub['id_o'].isin(used_o)) & ((o_sub['dt_o'] - rb['dt_b']).abs().dt.days <= 10)]
            cocok = kandidat[(kandidat[col_o] - rb[col_b]).abs() <= 1]
            if not cocok.empty:
                ro = cocok.assign(jarak=(cocok['dt_o'] - rb['dt_b']).abs()).sort_values('jarak').head(1).squeeze()
                used_b.add(rb['id_b']); used_o.add(ro['id_o'])
                beda_hari = int((rb['dt_b'] - ro['dt_o']).days)
                gabung = {**rb.to_dict(), **ro.to_dict(), 'Tanggal_Clean': rb['Tanggal_Clean']}
                gabung['Status_Rekonsiliasi'] = f'MATCH (Cocok - Beda Tanggal {beda_hari:+d} Hari | Tgl Odoo: {ro["Tanggal_Clean"]})'
                fuzzy_date_rows.append(gabung)

        split_match_rows = []
        for _, ro in o_sub[~o_sub['id_o'].isin(used_o)].iterrows():
            kandidat_b = b_sub[(~b_sub['id_b'].isin(used_b)) & ((b_sub['dt_b'] - ro['dt_o']).abs().dt.days <= 5)]
            ketemu = False
            for tgl_b, grup_b in kandidat_b.groupby('Tanggal_Clean'):
                if len(grup_b) < 2: continue
                rows_list = [r for _, r in grup_b.iterrows()]
                for k in range(2, min(len(rows_list), 6) + 1):
                    for comb in combinations(rows_list, k):
                        if abs(sum(item[col_b] for item in comb) - ro[col_o]) <= 1:
                            for idx_g, rb_item in enumerate(comb):
                                used_b.add(rb_item['id_b'])
                                rd = rb_item.to_dict()
                                if idx_g == 0:
                                    rd.update({key: val for key, val in ro.to_dict().items() if key not in ['Tanggal_Clean', col_b, 'Saldo']})
                                    rd['Status_Rekonsiliasi'] = f'MATCH (Gabungan {k} Trx Bank -> 1 Trx Odoo | Tgl Odoo: {ro["Tanggal_Clean"]})'
                                else:
                                    rd['Status_Rekonsiliasi'] = f'MATCH (Bagian Gabungan ke-{idx_g+1} dari {k} Trx Bank)'
                                split_match_rows.append(rd)
                            used_o.add(ro['id_o']); ketemu = True; break
                    if ketemu: break
                if ketemu: break

        unmatched_b = b_sub[~b_sub['id_b'].isin(used_b)].copy()
        unmatched_b['Status_Rekonsiliasi'] = 'HANYA DI BANK (Unrecorded Expense)' if tipe_str == 'Keluar' else 'HANYA DI BANK (Unrecorded Deposit)'
        unmatched_o = o_sub[~o_sub['id_o'].isin(used_o)].copy()
        unmatched_o['Status_Rekonsiliasi'] = 'HANYA DI ODOO (Outstanding Payment)' if tipe_str == 'Keluar' else 'HANYA DI ODOO (Deposit in Transit)'
        res = pd.concat([exact, pd.DataFrame(fuzzy_amt_rows), pd.DataFrame(fuzzy_date_rows), pd.DataFrame(split_match_rows), unmatched_b, unmatched_o], ignore_index=True)
        res['Tipe'] = tipe_str
        return res

    rekon_keluar = rekon_lengkap_bni(df_bank_normal, df_odoo_normal, 'Debet', 'Credit', 'Keluar')
    rekon_masuk = rekon_lengkap_bni(df_bank_normal, df_odoo_normal, 'Kredit', 'Debit', 'Masuk')
    df_total_normal = pd.concat([rekon_keluar, rekon_masuk], ignore_index=True).reset_index(drop=True)

    # Kontra Bank
    idx_bank_out = df_total_normal[df_total_normal['Status_Rekonsiliasi'] == 'HANYA DI BANK (Unrecorded Expense)'].index
    idx_bank_in = df_total_normal[df_total_normal['Status_Rekonsiliasi'] == 'HANYA DI BANK (Unrecorded Deposit)'].index
    used_in = set()
    for i_out in idx_bank_out:
        r_out = df_total_normal.loc[i_out]
        for i_in in idx_bank_in:
            if i_in in used_in: continue
            r_in = df_total_normal.loc[i_in]
            if abs((pd.to_datetime(r_in['Tanggal_Clean']) - pd.to_datetime(r_out['Tanggal_Clean'])).days) <= 3 and abs(r_in['Kredit'] - r_out['Debet']) <= 1:
                used_in.add(i_in)
                df_total_normal.at[i_out, 'Status_Rekonsiliasi'] = f"KONTRA BANK (Reversal / Lawan Tgl {r_in['Tanggal_Clean']})"
                df_total_normal.at[i_in, 'Status_Rekonsiliasi'] = f"KONTRA BANK (Reversal / Lawan Tgl {r_out['Tanggal_Clean']})"
                break

    df_laporan_akhir = pd.concat([df_total_normal, df_rekon_jagir, df_rekon_fee], ignore_index=True)
    kolom_laporan = [c for c in ['Tanggal_Clean', 'Status_Rekonsiliasi', 'Tipe', 'Keterangan', 'Debet', 'Kredit', 'Saldo', 'Number', 'Partner', 'Label', 'Debit', 'Credit'] if c in df_laporan_akhir.columns]
    df_laporan_final = df_laporan_akhir[kolom_laporan].drop_duplicates(subset=['Tanggal_Clean', 'Keterangan', 'Debet', 'Number', 'Credit'], keep='first').sort_values(by='Tanggal_Clean').reset_index(drop=True)

    total_debit_rekon, total_credit_rekon = df_laporan_final['Debit'].fillna(0).sum(), df_laporan_final['Credit'].fillna(0).sum()
    selisih_debit_odoo, selisih_credit_odoo = round(total_debit_rekon - total_debit_bawaan, 2), round(total_credit_rekon - total_credit_bawaan, 2)
    total_debet_bank_rekon = round(df_laporan_final['Debet'].fillna(0).sum() + val_pajak_total_jagir, 2)
    total_kredit_bank_rekon = round(df_laporan_final['Kredit'].fillna(0).sum() + val_pajak_total_jagir, 2)
    selisih_debet_bank, selisih_kredit_bank = round(total_debet_bank_rekon - total_debet_bank_asli, 2), round(total_kredit_bank_rekon - total_kredit_bank_asli, 2)

    pesan_audit = "OKE 100% - Total Bank & Total Odoo Hasil Rekon SAMA PERSIS dengan Data Asli!" if (selisih_debit_odoo == 0 and selisih_credit_odoo == 0 and selisih_debet_bank == 0 and selisih_kredit_bank == 0) else f"CEK SELISIH! (Bank: {selisih_debet_bank:,.2f}/{selisih_kredit_bank:,.2f} | Odoo: {selisih_debit_odoo:,.2f}/{selisih_credit_odoo:,.2f})"

    baris_audit = pd.DataFrame([
        {'Tanggal_Clean': 'TOTAL REKON', 'Status_Rekonsiliasi': 'SUM Hasil Rekon (Termasuk Netting Pajak)', 'Debet': total_debet_bank_rekon, 'Kredit': total_kredit_bank_rekon, 'Debit': total_debit_rekon, 'Credit': total_credit_rekon},
        {'Tanggal_Clean': 'TOTAL ASLI', 'Status_Rekonsiliasi': 'Total Asli Rekening Koran & Odoo', 'Debet': total_debet_bank_asli, 'Kredit': total_kredit_bank_asli, 'Debit': total_debit_bawaan, 'Credit': total_credit_bawaan},
        {'Tanggal_Clean': 'STATUS AUDIT', 'Status_Rekonsiliasi': pesan_audit, 'Debet': selisih_debet_bank, 'Kredit': selisih_kredit_bank, 'Debit': selisih_debit_odoo, 'Credit': selisih_credit_odoo}
    ])

    df_export = pd.concat([df_laporan_final, baris_audit], ignore_index=True)
    return df_export, df_laporan_final, pesan_audit
# ============================================================
# ANTARMUKA PENGGUNA (STREAMLIT UI)
# ============================================================
col1, col2 = st.columns(2)
with col1:
    uploaded_pdf = st.file_uploader("1. Unggah Rekening Koran BNI (.pdf)", type=["pdf"])
with col2:
    uploaded_odoo = st.file_uploader("2. Unggah Mutasi Bank Odoo (.xlsx)", type=["xlsx"])

if uploaded_pdf and uploaded_odoo:
    if st.button("🚀 Proses Rekonsiliasi Sekarang", type="primary", use_container_width=True):
        with st.spinner("Sedang mengekstrak PDF BNI dan mencocokkan dengan data Odoo..."):
            try:
                # 1. Ekstraksi PDF
                df_bank, saldo_awal, saldo_akhir, pdf_errors = ekstrak_pdf_bni(uploaded_pdf)
                
                # 2. Proses Rekonsiliasi
                df_odoo_raw = pd.read_excel(uploaded_odoo)
                df_export, df_laporan_final, pesan_audit = proses_rekonsiliasi(df_bank.copy(), df_odoo_raw)

                st.divider()

                # Ringkasan Status Ekstraksi & Audit
                if len(pdf_errors) == 0:
                    st.success(f"**Ekstraksi PDF Valid 100%!** ({len(df_bank)} transaksi | Saldo Awal: Rp {saldo_awal:,.2f} → Saldo Akhir: Rp {saldo_akhir:,.2f})")
                else:
                    st.warning(f"Ditemukan {len(pdf_errors)} baris selisih pada pembacaan saldo PDF.")

                if "OKE 100%" in pesan_audit:
                    st.success(f"**Status Audit Rekonsiliasi:** {pesan_audit}")
                else:
                    st.error(f"**Status Audit Rekonsiliasi:** {pesan_audit}")

                # Kartu Metrik Ringkas
                m1, m2, m3, m4 = st.columns(4)
                jml_match = df_laporan_final['Status_Rekonsiliasi'].str.startswith('MATCH').sum()
                jml_bank_only = df_laporan_final['Status_Rekonsiliasi'].str.startswith('HANYA DI BANK').sum()
                jml_odoo_only = df_laporan_final['Status_Rekonsiliasi'].str.startswith('HANYA DI ODOO').sum()
                jml_kontra = df_laporan_final['Status_Rekonsiliasi'].str.startswith('KONTRA BANK').sum()

                m1.metric("Total Baris Laporan", len(df_laporan_final))
                m2.metric("Cocok (MATCH)", jml_match)
                m3.metric("Hanya di Bank", jml_bank_only)
                m4.metric("Hanya di Odoo", jml_odoo_only)

                # Filter Tampilan Tabel di Web
                tab1, tab2 = st.tabs(["📊 Hasil Rekonsiliasi Lengkap", "📄 Hasil Ekstraksi PDF BNI"])
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

                # Siapkan File Excel di Memori untuk Diunduh
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                    df_export.to_excel(writer, sheet_name="Hasil Rekonsiliasi", index=False)
                    df_bank.to_excel(writer, sheet_name="Ekstraksi PDF BNI", index=False)

                nama_file_keluar = f"Hasil_Rekonsiliasi_{uploaded_pdf.name.replace('.pdf', '')}.xlsx"
                st.download_button(
                    label="📥 Download Hasil Rekonsiliasi (.xlsx)",
                    data=output_buffer.getvalue(),
                    file_name=nama_file_keluar,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses file: {e}")
