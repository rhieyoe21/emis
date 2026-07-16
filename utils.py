import os
import re
import time
import pandas as pd # type: ignore
from datetime import datetime
from shutil import copyfile
from config import LOG_FILE

def safe_str(val):
    val_str = str(val).strip() if pd.notna(val) else ""
    if val_str.startswith("'"):
        val_str = val_str[1:]
    return val_str

def backup_excel(source_path):
    backup_folder = "backup"
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)
    
    backup_name = os.path.join(backup_folder, f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    if os.path.exists(source_path):
        copyfile(source_path, backup_name)
        print(f"📁 Backup dibuat: {backup_name}")
        return backup_name
    else:
        raise FileNotFoundError("❌ File Excel tidak ditemukan.")

def tulis_log(message, log_path=LOG_FILE):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{timestamp} {message}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def update_status_excel(ws, nik, col_nik_index, col_status_index, col_timestamp_index):
    for row_excel in ws.iter_rows(min_row=2, values_only=False):
        nik_excel = safe_str(row_excel[col_nik_index - 1].value)
        status_excel = safe_str(row_excel[col_status_index - 1].value).lower()

        if nik_excel == nik and status_excel in ["", "belum"]:
            row_excel[col_status_index - 1].value = "sudah"
            row_excel[col_timestamp_index - 1].value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            return True
    return False

def safe_save_workbook(wb, path, max_retry=3):
    for attempt in range(max_retry):
        try:
            wb.save(path)
            return True
        except Exception as e:
            print(f"⚠️ Gagal menyimpan Excel (percobaan {attempt+1}): {e}")
            time.sleep(2)
    print("❌ Gagal menyimpan Excel setelah beberapa percobaan.")
    return False

def format_nomor_hp(nomor):
    nomor = safe_str(nomor)
    if not nomor:
        return "", "false"

    # Hilangkan karakter non-digit
    nomor = re.sub(r"[^\d]", "", nomor)

    # Jika sudah diawali dengan 62 dan panjang cukup
    if re.match(r"^62\d{8,}$", nomor):
        return nomor, "true"

    # Jika diawali dengan 08 dan panjang cukup
    elif re.match(r"^08\d{8,}$", nomor):
        nomor_formatted = "62" + nomor[1:]
        return nomor_formatted, "true"

    # Jika diawali dengan 8 dan panjang cukup
    elif re.match(r"^8\d{8,}$", nomor):
        nomor_formatted = "62" + nomor
        return nomor_formatted, "true"

    # Format tidak valid
    return "", "false"

def format_tanggal(tanggal_str, min_year=1950, max_year=None):
    try:
        tanggal = pd.to_datetime(tanggal_str, errors='coerce')
        if pd.isna(tanggal):
            return "", False
        tahun = tanggal.year
        batas_atas = max_year if max_year else datetime.now().year
        if tahun < min_year or tahun > batas_atas:
            return "", False
        return tanggal.strftime("%Y-%m-%d"), True
    except:
        return "", False

def validasi_usia_masuk(birth_date_str, admission_date_str):
    try:
        birth_date = pd.to_datetime(birth_date_str, errors='coerce')
        admission_date = pd.to_datetime(admission_date_str, errors='coerce')
        if pd.isna(birth_date) or pd.isna(admission_date):
            return False
        return birth_date <= admission_date
    except:
        return False

def format_kode_wilayah_for_postal(kode):
    """Format kode wilayah untuk pencarian kode pos"""
    kode = str(kode).strip().replace(".", "")
    if len(kode) >= 12:
        return f"{kode[0:2]}.{kode[2:4]}.{kode[4:6]}.{kode[6:10]}"
    elif len(kode) == 10:
        return f"{kode[0:2]}.{kode[2:4]}.{kode[4:6]}.{kode[6:10]}"
    elif len(kode) == 6:
        return f"{kode[0:2]}.{kode[2:4]}.{kode[4:6]}"
    elif len(kode) == 4:
        return f"{kode[0:2]}.{kode[2:4]}"
    else:
        return kode
