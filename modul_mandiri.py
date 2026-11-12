import pdfplumber
import pandas as pd
import re
from datetime import datetime


def get_column_lines(words, xmin, xmax, ymin, ymax):
    selected = [
        w for w in words
        if ymin <= w["top"] < ymax
        and xmin <= w["x0"] < xmax
    ]
    selected.sort(key=lambda w: (w["top"], w["x0"]))

    lines = []
    for word in selected:
        if not lines or abs(word["top"] - lines[-1][0]) > 2.5:
            lines.append([word["top"], [word]])
        else:
            lines[-1][1].append(word)

    result = []
    for line in lines:
        text = " ".join(
            w["text"] for w in sorted(line[1], key=lambda x: x["x0"])
        )
        result.append(text)
    return result


def get_amount_mandiri(lines):
    values = []
    for line in lines:
        found = re.findall(r"-?[\d,]+\.\d{2}", line)
        values.extend(found)
    if values:
        return float(values[-1].replace(",", ""))
    return 0.0
def ekstrak_pdf_mandiri(pdf_file):
    transactions = []
    full_text = ""

    with pdfplumber.open(pdf_file) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            full_text += "\n" + (page.extract_text() or "")
            words = page.extract_words()

            posting_positions = [
                w["top"] for w in words
                if w["text"].lower() == "posting"
            ]
            header_y = min(posting_positions) if posting_positions else 0

            date_anchors = []
            for word in words:
                if (
                    word["top"] <= header_y + 5
                    or word["x0"] >= 100
                    or not re.fullmatch(r"\d{1,2}", word["text"])
                ):
                    continue

                same_line = [
                    w for w in words
                    if abs(w["top"] - word["top"]) < 2
                    and word["x1"] <= w["x0"] < 115
                ]

                date_text = word["text"] + " " + " ".join(
                    w["text"] for w in sorted(same_line, key=lambda x: x["x0"])
                )

                match = re.match(
                    r"(\d{1,2})\s+"
                    r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
                    r"\s+(\d{4}),?",
                    date_text,
                    re.IGNORECASE
                )

                if match:
                    date_anchors.append(
                        (word["top"], match.group(0).replace(",", ""))
                    )

            date_anchors.sort(key=lambda x: x[0])

            for i, (anchor, date_text) in enumerate(date_anchors):
                y_start = max(header_y + 5, anchor - 20)
                if i + 1 < len(date_anchors):
                    y_end = date_anchors[i + 1][0] - 20
                else:
                    y_end = page.height - 65

                transaction_date = datetime.strptime(date_text, "%d %b %Y")

                time_lines = get_column_lines(words, 0, 110, anchor, anchor + 20)
                transaction_time = None
                for line in time_lines:
                    time_match = re.search(r"\d{2}:\d{2}:\d{2}", line)
                    if time_match:
                        transaction_time = time_match.group(0)
                        break

                if transaction_time:
                    transaction_date = datetime.strptime(
                        transaction_date.strftime("%d %b %Y") + " " + transaction_time,
                        "%d %b %Y %H:%M:%S"
                    )

                remark = get_column_lines(words, 110, 235, y_start, y_end)
                reference = get_column_lines(words, 235, 320, y_start, y_end)
                debit = get_column_lines(words, 320, 410, y_start, y_end)
                credit = get_column_lines(words, 410, 500, y_start, y_end)
                balance = get_column_lines(words, 500, 590, y_start, y_end)

                transactions.append({
                    "Tanggal": transaction_date,
                    "Keterangan": " ".join(remark).strip(),
                    "No. Referensi": " ".join(reference).strip(),
                    "Debet": get_amount_mandiri(debit),
                    "Kredit": get_amount_mandiri(credit),
                    "Saldo": get_amount_mandiri(balance),
                    "Halaman": page_number
                })
    df = pd.DataFrame(transactions)
    if df.empty:
        raise ValueError("Tidak ditemukan transaksi. Periksa format PDF Mandiri.")

    df = df.sort_values("Tanggal").reset_index(drop=True)
    df = df[
        (df["Debet"] != 0)
        | (df["Kredit"] != 0)
        | (df["Saldo"] != 0)
    ].copy()

    # Baca ringkasan PDF asli Mandiri
    opening_match = re.search(
        r"Opening Balance\s+No\.\s*of Debit\s+"
        r"Total Amount Debited\s+"
        r"([\d,]+\.\d{2})\s+(\d+)\s+([\d,]+\.\d{2})",
        full_text,
        re.IGNORECASE
    )

    closing_match = re.search(
        r"Closing Balance\s+No\.\s*of Credit\s+"
        r"Total Amount Credited\s+"
        r"([\d,]+\.\d{2})\s+(\d+)\s+([\d,]+\.\d{2})",
        full_text,
        re.IGNORECASE
    )

    if not opening_match or not closing_match:
        raise ValueError("Ringkasan saldo tidak ditemukan. Periksa format PDF Mandiri.")

    opening_balance = float(opening_match.group(1).replace(",", ""))
    summary_debit_count = int(opening_match.group(2))
    summary_debit = float(opening_match.group(3).replace(",", ""))
    closing_balance = float(closing_match.group(1).replace(",", ""))
    summary_credit_count = int(closing_match.group(2))
    summary_credit = float(closing_match.group(3).replace(",", ""))

    total_debit = df["Debet"].sum()
    total_credit = df["Kredit"].sum()

    calculated_balance = opening_balance
    errors = []
    for index, row in df.iterrows():
        calculated_balance = round(calculated_balance + row["Kredit"] - row["Debet"], 2)
        difference = round(calculated_balance - row["Saldo"], 2)
        if abs(difference) > 0.01:
            errors.append({
                "Baris": index + 1, "Tanggal": str(row["Tanggal"]), "Halaman": row["Halaman"],
                "Keterangan": row["Keterangan"], "Saldo Hitungan": calculated_balance,
                "Saldo PDF": row["Saldo"], "Selisih": difference
            })

    balance_difference = round(calculated_balance - closing_balance, 2)
    if round(total_debit - summary_debit, 2) != 0 or round(total_credit - summary_credit, 2) != 0 or balance_difference != 0:
        errors.append({
            "Baris": "SUMMARY", "Tanggal": "-", "Halaman": "-",
            "Keterangan": "Selisih Ringkasan PDF Mandiri",
            "Saldo Hitungan": calculated_balance, "Saldo PDF": closing_balance,
            "Selisih": balance_difference
        })

    return df, opening_balance, closing_balance, errors
def proses_rekonsiliasi_mandiri(df_bank_raw, df_odoo_raw):
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
    df_bank['Keterangan_Upper'] = df_bank['Keterangan'].astype(str).str.upper()
    df_odoo['Label_Upper'] = df_odoo['Label'].astype(str).str.upper()

    # 1. Akumulasi Biaya Admin (Maksimal Rp249.000 & Khusus Keluar)
    filter_bank_fee = (
        df_bank['Keterangan_Upper'].str.contains('TRANSACTION FEE|BIAYA ADM|TRANSFER FEE|INHOUSETRF KE GNC')
        & (df_bank['Debet'] > 0) & (df_bank['Debet'] <= 249000)
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
            'Keterangan': f"Akumulasi Transaction Fee & Admin Bank Koran (Total: Rp{total_fee_bank:,.2f})",
            'Debet': total_fee_bank, 'Kredit': 0.0,
            'Label': f"Total Catatan Admin di Odoo (Total: Rp{total_fee_odoo:,.2f})",
            'Credit': total_fee_odoo, 'Debit': 0.0
        })
    df_rekon_fee = pd.DataFrame(rekon_fee_rows)

    # 2. Net Jagir Akhir Bulan (Toleransi Pembulatan Rp5 + Simpan Total Pajak untuk Audit Bank)
    bunga_bank = df_bank[df_bank['Keterangan_Upper'].str.contains('BUNGA')].copy()
    pajak_bank = df_bank[df_bank['Keterangan_Upper'].str.contains('PAJAK')].copy()

    net_jagir_bank_rows = []
    val_pajak_total_jagir = 0.0
    if not bunga_bank.empty and not pajak_bank.empty:
        for tgl in bunga_bank['Tanggal_Clean'].unique():
            val_bunga = bunga_bank[bunga_bank['Tanggal_Clean'] == tgl]['Kredit'].sum()
            val_pajak = pajak_bank[pajak_bank['Tanggal_Clean'] == tgl]['Debet'].sum()
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
        df_rekon_jagir['Keterangan'] = df_rekon_jagir.apply(lambda r: f"Bunga Koran (Rp{r['Bunga_Asli']:,.2f}) & Pajak Koran (Rp{r['Pajak_Asli']:,.2f}) di-Nettkan" if pd.notna(r['Bunga_Asli']) else '', axis=1)
        df_rekon_jagir['Kredit'] = df_rekon_jagir['Net_Jagir_Bank']
        df_rekon_jagir['Debet'] = 0.0
    else:
        df_rekon_jagir = pd.DataFrame()
    # 3. Transaksi Normal Satu-ke-Satu
    filter_bank_normal = ~filter_bank_fee & ~df_bank['Keterangan_Upper'].str.contains('BUNGA|PAJAK')
    df_bank_normal = df_bank[filter_bank_normal].copy()

    filter_odoo_normal = ~filter_odoo_fee & ~filter_odoo_jagir
    df_odoo_normal = df_odoo[filter_odoo_normal].copy()

    df_bank_normal['Nominal_Keluar'] = df_bank_normal['Debet']
    df_bank_normal['Nominal_Masuk'] = df_bank_normal['Kredit']
    df_bank_normal['Seq_Keluar'] = df_bank_normal.groupby(['Tanggal_Clean', 'Nominal_Keluar']).cumcount()
    df_bank_normal['Seq_Masuk'] = df_bank_normal.groupby(['Tanggal_Clean', 'Nominal_Masuk']).cumcount()

    df_odoo_normal['Nominal_Keluar'] = df_odoo_normal['Credit']
    df_odoo_normal['Nominal_Masuk'] = df_odoo_normal['Debit']
    df_odoo_normal['Seq_Keluar'] = df_odoo_normal.groupby(['Tanggal_Clean', 'Nominal_Keluar']).cumcount()
    df_odoo_normal['Seq_Masuk'] = df_odoo_normal.groupby(['Tanggal_Clean', 'Nominal_Masuk']).cumcount()

    rekon_keluar = pd.merge(df_bank_normal[df_bank_normal['Nominal_Keluar'] > 0], df_odoo_normal[df_odoo_normal['Nominal_Keluar'] > 0], on=['Tanggal_Clean', 'Nominal_Keluar', 'Seq_Keluar'], how='outer')
    rekon_keluar['Tipe'] = 'Keluar'

    rekon_masuk = pd.merge(df_bank_normal[df_bank_normal['Nominal_Masuk'] > 0], df_odoo_normal[df_odoo_normal['Nominal_Masuk'] > 0], on=['Tanggal_Clean', 'Nominal_Masuk', 'Seq_Masuk'], how='outer')
    rekon_masuk['Tipe'] = 'Masuk'

    df_total_normal = pd.concat([rekon_keluar, rekon_masuk], ignore_index=True)

    def tentukan_status_normal(row):
        if row['Tipe'] == 'Keluar':
            if pd.isna(row['Keterangan']): return 'HANYA DI ODOO (Outstanding Payment)'
            elif pd.isna(row['Label']): return 'HANYA DI BANK (Unrecorded Expense)'
            else: return 'MATCH (Cocok Sempurna)'
        else:
            if pd.isna(row['Keterangan']): return 'HANYA DI ODOO (Deposit in Transit)'
            elif pd.isna(row['Label']): return 'HANYA DI BANK (Unrecorded Deposit)'
            else: return 'MATCH (Cocok Sempurna)'

    df_total_normal['Status_Rekonsiliasi'] = df_total_normal.apply(tentukan_status_normal, axis=1)

    # 4. Gabungkan Laporan & Audit Cross-Check Total Bank + Total Odoo
    df_laporan_akhir = pd.concat([df_total_normal, df_rekon_jagir, df_rekon_fee], ignore_index=True)
    kolom_laporan = ['Tanggal_Clean', 'Status_Rekonsiliasi', 'Tipe', 'Keterangan', 'Debet', 'Kredit', 'Saldo', 'Number', 'Partner', 'Label', 'Debit', 'Credit']
    kolom_laporan = [c for c in kolom_laporan if c in df_laporan_akhir.columns]

    df_laporan_final = df_laporan_akhir[kolom_laporan].drop_duplicates(subset=['Tanggal_Clean', 'Keterangan', 'Debet', 'Number', 'Credit'], keep='first').sort_values(by='Tanggal_Clean').reset_index(drop=True)

    # Audit Odoo
    total_debit_rekon = df_laporan_final['Debit'].fillna(0).sum()
    total_credit_rekon = df_laporan_final['Credit'].fillna(0).sum()
    selisih_debit_odoo = round(total_debit_rekon - total_debit_bawaan, 2)
    selisih_credit_odoo = round(total_credit_rekon - total_credit_bawaan, 2)

    # Audit Bank Mandiri (Termasuk Netting Pajak Bunga)
    total_debet_bank_asli = df_bank['Debet'].sum()
    total_kredit_bank_asli = df_bank['Kredit'].sum()
    total_debet_bank_rekon = round(df_laporan_final['Debet'].fillna(0).sum() + val_pajak_total_jagir, 2)
    total_kredit_bank_rekon = round(df_laporan_final['Kredit'].fillna(0).sum() + val_pajak_total_jagir, 2)
    selisih_debet_bank = round(total_debet_bank_rekon - total_debet_bank_asli, 2)
    selisih_kredit_bank = round(total_kredit_bank_rekon - total_kredit_bank_asli, 2)

    if selisih_debit_odoo == 0 and selisih_credit_odoo == 0 and selisih_debet_bank == 0 and selisih_kredit_bank == 0:
        pesan_audit = "OKE 100% - Total Bank & Total Odoo Hasil Rekon SAMA PERSIS dengan Data Asli!"
    else:
        pesan_audit = f"CEK SELISIH! (Bank: {selisih_debet_bank:,.2f}/{selisih_kredit_bank:,.2f} | Odoo: {selisih_debit_odoo:,.2f}/{selisih_credit_odoo:,.2f})"

    baris_audit = pd.DataFrame([
        {'Tanggal_Clean': 'TOTAL REKON', 'Status_Rekonsiliasi': 'SUM Hasil Rekon (Termasuk Netting Pajak)', 'Debet': total_debet_bank_rekon, 'Kredit': total_kredit_bank_rekon, 'Debit': total_debit_rekon, 'Credit': total_credit_rekon},
        {'Tanggal_Clean': 'TOTAL ASLI', 'Status_Rekonsiliasi': 'Total Asli Rekening Koran & Odoo', 'Debet': total_debet_bank_asli, 'Kredit': total_kredit_bank_asli, 'Debit': total_debit_bawaan, 'Credit': total_credit_bawaan},
        {'Tanggal_Clean': 'STATUS AUDIT', 'Status_Rekonsiliasi': pesan_audit, 'Debet': selisih_debet_bank, 'Kredit': selisih_kredit_bank, 'Debit': selisih_debit_odoo, 'Credit': selisih_credit_odoo}
    ])

    df_export = pd.concat([df_laporan_final, baris_audit], ignore_index=True)
    return df_export, df_laporan_final, pesan_audit
