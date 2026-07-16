import os
import time
import pandas as pd # type: ignore
from openpyxl import load_workbook # type: ignore

from config import EXCEL_FILE, LOG_FILE, API_URL, DELAY, token, admission_date
from utils import (
    safe_str, backup_excel, tulis_log, update_status_excel,
    safe_save_workbook, is_excel_locked
)
from postal import autofill_postal_codes_pre_upload, format_phone_numbers_in_excel, format_birthdates_in_excel
from api import build_files, upload_data, parse_api_response

def tampilkan_ringkasan(data):
    print("\n📋 Ringkasan Data Siswa:")
    print(f"Nama Lengkap        : {data['full_name']}")
    print(f"NIK                 : {data['nik']}")
    print(f"Tempat, Tgl Lahir   : {data['birth_place']}, {data['birth_date']}")
    print(f"Alamat              : {data['address']}")
    print(f"Jenis Kelamin       : {data['gender']}")
    print(f"Nama Ayah           : {data['father_full_name']}")
    print(f"Nama Ibu            : {data['mother_full_name']}")
    print(f"File Kartu Keluarga : {str(data['full_name']).strip()}.jpg")

def main():
    while is_excel_locked(EXCEL_FILE):
        print(f"\n⚠️ File '{EXCEL_FILE}' sedang terbuka di program lain (misal: Microsoft Excel).")
        print("Silakan tutup file tersebut terlebih dahulu agar sistem dapat menyimpan status upload.")
        input("Tekan Enter untuk mencoba kembali setelah file ditutup...")

    try:
        backup_excel(EXCEL_FILE)
    except Exception as e:
        print(e)
        return

    if not token:
        print("❌ Token tidak ditemukan di file config.txt.")
        return
    
    autofill_postal_codes_pre_upload()
    format_phone_numbers_in_excel()
    format_birthdates_in_excel()
    
    headers = {
        'Accept': 'application/json',
        'Accept-Language': 'id,en-US;q=0.9,en;q=0.8,ms;q=0.7',
        'Authorization': f'Bearer {token}',
        'Connection': 'keep-alive',
        'Origin': 'https://emis.kemenag.go.id',
        'Referer': 'https://emis.kemenag.go.id/',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-site',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36 Edg/150.0.0.0',
        'sec-ch-ua': '"Not;A=Brand";v="8", "Chromium";v="150", "Microsoft Edge";v="150"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }

    # ✅ Validasi format admission_date
    try:
        admission_date_parsed = pd.to_datetime(admission_date, errors='raise')
    except Exception as e:
        print(f"❌ Format admission_date tidak valid: {admission_date}")
        tulis_log(f"❌ Format admission_date tidak valid: {admission_date}")
        return

    df = pd.read_excel(EXCEL_FILE).fillna("")
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active

    # ====== VALIDASI DUPLIKAT NIK ======
    duplikat_nik_mask = df['nik'].duplicated(keep=False)
    nik_duplikat_set = set(df.loc[duplikat_nik_mask, 'nik'])

    if nik_duplikat_set:
        print("⚠️ Terdapat NIK duplikat di file Excel. Baris berikut akan dilewati:")
        for nik in nik_duplikat_set:
            siswa_duplikat = df[df['nik'] == nik]
            for _, row in siswa_duplikat.iterrows():
                print(f"🔁 {row['full_name']} | NIK: {row['nik']}")
                tulis_log(f"🔁 Duplikat NIK dilewati: {row['full_name']} | NIK: {row['nik']}")

    col_nik_index = df.columns.get_loc("nik") + 1
    col_status_index = df.columns.get_loc("status") + 1
    
    if "upload_timestamp" not in [col.lower() for col in df.columns]:
        ws.cell(row=1, column=ws.max_column + 1).value = "upload_timestamp"
        col_timestamp_index = ws.max_column
    else:
        # Cari index kolom upload_timestamp (case-insensitive)
        col_timestamp_index = next(
            idx + 1 for idx, col in enumerate(df.columns) 
            if col.lower() == "upload_timestamp"
        )
    
    jumlah_sukses = 0
    jumlah_gagal = 0
    jumlah_dilewati = 0
    jumlah_dilewati_tanggal = 0
    jumlah_dilewati_nik = 0
    jumlah_dilewati_duplikat = 0

    for idx, row in df.iterrows():
        nik = safe_str(row.get("nik"))
        status = safe_str(row.get("status")).lower()

        if len(nik) != 16 or not nik.isdigit():
            tulis_log(f"❌ NIK tidak valid: {row['full_name']} | NIK: {nik}")
            jumlah_dilewati_nik += 1
            continue

        if nik in nik_duplikat_set:
            jumlah_dilewati_duplikat += 1
            continue

        if status not in ["", "belum"]:
            continue

        print(f"\n🚀 Upload siswa ke-{idx+1}: {row['full_name']} (NIK: {nik})")
        files_and_file = build_files(row)
        if files_and_file is None:
            tulis_log(f"⏭️ Dilewati karena data tidak valid: {row['full_name']}")
            jumlah_dilewati_tanggal += 1
            continue
        files, foto_file, foto_filename, is_temp = files_and_file
        tampilkan_ringkasan(row)
        lanjut = input("➡️ Lanjutkan upload? (Y/N): ").strip().lower()
        if lanjut != 'y':
            print("⏭️ Upload dibatalkan oleh pengguna.")
            if foto_file:
                foto_file.close()
            if is_temp and os.path.exists(foto_filename):
                os.remove(foto_filename)
            jumlah_dilewati += 1
            continue

        response = upload_data(files, headers, API_URL)
        full_name = row.get("full_name", "❓ Nama tidak ditemukan")
        
        log_message = parse_api_response(response, full_name)
        tulis_log(log_message)
        
        if response is not None and response.status_code == 200:
            jumlah_sukses += 1
            if update_status_excel(ws, nik, col_nik_index, col_status_index, col_timestamp_index):
                safe_save_workbook(wb, EXCEL_FILE)
        else:
            jumlah_gagal += 1

        if foto_file:
            foto_file.close()
        if is_temp and os.path.exists(foto_filename):
            os.remove(foto_filename)
        time.sleep(DELAY)

    print("\n📁 Proses upload selesai. Status diperbarui langsung di file 'data_siswa.xlsx'")
    print("\n📊 Ringkasan Harian Upload:")
    print(f"✅ Berhasil upload     : {jumlah_sukses}")
    print(f"❌ Gagal upload        : {jumlah_gagal}")
    print(f"⏭️ Dilewati (manual)   : {jumlah_dilewati}")
    print(f"⏭️ Dilewati (tanggal)  : {jumlah_dilewati_tanggal}")
    print(f"⏭️ Dilewati (NIK salah): {jumlah_dilewati_nik}")
    print(f"⏭️ Dilewati (duplikat) : {jumlah_dilewati_duplikat}")
    
    tulis_log("📊 Ringkasan Harian Upload:")
    tulis_log(f"✅ Berhasil upload     : {jumlah_sukses}")
    tulis_log(f"❌ Gagal upload        : {jumlah_gagal}")
    tulis_log(f"⏭️ Dilewati (manual)   : {jumlah_dilewati}")
    tulis_log(f"⏭️ Dilewati (tanggal)  : {jumlah_dilewati_tanggal}")
    tulis_log(f"⏭️ Dilewati (NIK salah): {jumlah_dilewati_nik}")
    tulis_log(f"⏭️ Dilewati (duplikat) : {jumlah_dilewati_duplikat}")
    
    wb.close()

if __name__ == "__main__":
    main()