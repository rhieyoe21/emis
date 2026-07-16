import os
import pandas as pd
import requests
import time
from bs4 import BeautifulSoup
from datetime import datetime
from shutil import copyfile

EXCEL_FILE = "data_siswa.xlsx"
LOG_FILE = "log_autofill_postal.txt"

def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{timestamp} {message}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def backup_excel(source_path):
    backup_folder = "backup"
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)
    
    backup_name = os.path.join(backup_folder, f"backup_postal_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    copyfile(source_path, backup_name)
    log_message(f"Backup dibuat: {backup_name}")
    return backup_name

def format_kode_wilayah(kode):
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

def cari_kodepos_web(kode_wilayah):
    try:
        kode_bertitik = format_kode_wilayah(kode_wilayah)
        url = f"https://kodepos.nomor.net/_kodepos.php?_i=cari-kodepos&jobs={kode_bertitik}"
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            tables = soup.find_all('table')
            
            for table in tables:
                rows = table.find_all('tr')
                for row in rows:
                    cols = row.find_all(['td', 'th'])
                    
                    for i in range(len(cols) - 6):
                        col_text = cols[i].get_text(strip=True)
                        
                        if col_text.isdigit() and len(col_text) == 5:
                            try:
                                kode_wil = cols[i+2].get_text(strip=True) if i+2 < len(cols) else ''
                                
                                if kode_wil and '.' in kode_wil:
                                    return col_text
                            except:
                                pass
            return None
        else:
            return None
    except Exception as e:
        log_message(f"Error cari kode pos: {e}")
        return None

def autofill_postal_codes():
    log_message("=== MULAI AUTO-FILL KODE POS ===")
    
    try:
        backup_excel(EXCEL_FILE)
    except Exception as e:
        log_message(f"Error backup: {e}")
        return
    
    df = pd.read_excel(EXCEL_FILE)
    
    df['postal_code_num'] = df['postal_code_num'].astype(str).str.strip()
    
    rows_to_update = df[(df['postal_code_num'] == '') | (df['postal_code_num'] == 'nan') | (df['postal_code_num'].isna())]
    
    log_message(f"Total baris perlu update: {len(rows_to_update)}")
    
    updated_count = 0
    failed_count = 0
    skipped_count = 0
    
    for idx, row in rows_to_update.iterrows():
        full_name = row.get('full_name', 'N/A')
        m_subdistrict_id = str(row.get('m_subdistrict_id', '')).strip()
        
        if not m_subdistrict_id or m_subdistrict_id == 'nan':
            log_message(f"[SKIP] {full_name}: m_subdistrict_id kosong")
            skipped_count += 1
            continue
        
        log_message(f"[PROSES] {full_name} - Kode wilayah: {m_subdistrict_id}")
        
        kode_pos = cari_kodepos_web(m_subdistrict_id)
        
        if kode_pos:
            df.at[idx, 'postal_code_num'] = kode_pos
            log_message(f"[OK] {full_name}: Kode pos = {kode_pos}")
            updated_count += 1
        else:
            log_message(f"[GAGAL] {full_name}: Kode pos tidak ditemukan")
            failed_count += 1
        
        time.sleep(2)
    
    df.to_excel(EXCEL_FILE, index=False)
    log_message(f"File Excel berhasil diupdate: {EXCEL_FILE}")
    
    log_message("\n=== RINGKASAN AUTO-FILL KODE POS ===")
    log_message(f"Berhasil diupdate: {updated_count}")
    log_message(f"Gagal: {failed_count}")
    log_message(f"Dilewati: {skipped_count}")
    log_message("=== SELESAI ===")

if __name__ == "__main__":
    autofill_postal_codes()
