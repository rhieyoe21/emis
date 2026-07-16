import requests
import time
import pandas as pd
import re
from bs4 import BeautifulSoup

def load_wilayah_data():
    """Load data wilayah dari CSV"""
    try:
        df = pd.read_csv("wilayah.csv", header=None, names=["kode", "nama"])
        df["kode"] = df["kode"].astype(str)
        df["kode_clean"] = df["kode"].str.replace(".", "", regex=False)
        return df
    except Exception as e:
        print(f"[ERROR] Gagal load wilayah.csv: {e}")
        return None

def get_nama_wilayah(kode_wilayah):
    """Ambil nama wilayah dari kode"""
    df = load_wilayah_data()
    if df is None:
        return None
    
    kode_clean = str(kode_wilayah).replace(".", "")
    
    result = df[df["kode_clean"] == kode_clean]
    if not result.empty:
        return result.iloc[0]["nama"]
    
    return None

def format_kode_wilayah(kode):
    """Convert kode wilayah tanpa titik menjadi format bertitik"""
    kode = str(kode).strip().replace(".", "")
    
    if len(kode) == 2:
        return kode
    elif len(kode) == 4:
        return f"{kode[0:2]}.{kode[2:4]}"
    elif len(kode) == 6:
        return f"{kode[0:2]}.{kode[2:4]}.{kode[4:6]}"
    elif len(kode) == 10:
        return f"{kode[0:2]}.{kode[2:4]}.{kode[4:6]}.{kode[6:10]}"
    elif len(kode) >= 12:
        return f"{kode[0:2]}.{kode[2:4]}.{kode[4:6]}.{kode[6:10]}"
    else:
        return kode

def cari_kodepos_web(kode_wilayah):
    """Cari kode pos dengan scraping halaman web kodepos.nomor.net"""
    try:
        kode_bertitik = format_kode_wilayah(kode_wilayah)
        url = f"https://kodepos.nomor.net/_kodepos.php?_i=cari-kodepos&jobs={kode_bertitik}"
        
        print(f"[*] Mencari kode pos untuk: {kode_bertitik}")
        print(f"[*] URL: {url}")
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            
            kode_pos_list = []
            tables = soup.find_all('table')
            
            for table in tables:
                rows = table.find_all('tr')
                
                for row in rows:
                    cols = row.find_all(['td', 'th'])
                    
                    for i in range(len(cols) - 6):
                        col_text = cols[i].get_text(strip=True)
                        
                        if col_text.isdigit() and len(col_text) == 5:
                            kode_pos = col_text
                            
                            try:
                                kelurahan = cols[i+1].get_text(strip=True) if i+1 < len(cols) else ''
                                kode_wil = cols[i+2].get_text(strip=True) if i+2 < len(cols) else ''
                                kecamatan = cols[i+3].get_text(strip=True) if i+3 < len(cols) else ''
                                kota1 = cols[i+4].get_text(strip=True) if i+4 < len(cols) else ''
                                kota2 = cols[i+5].get_text(strip=True) if i+5 < len(cols) else ''
                                provinsi = cols[i+6].get_text(strip=True) if i+6 < len(cols) else ''
                                
                                kota = f"{kota1} {kota2}".strip()
                                
                                if kode_wil and '.' in kode_wil:
                                    print(f"\n[OK] Kode Pos: {kode_pos}")
                                    print(f"  Kelurahan: {kelurahan}")
                                    print(f"  Kode Wilayah: {kode_wil}")
                                    print(f"  Kecamatan: {kecamatan}")
                                    print(f"  Kota/Kab: {kota}")
                                    print(f"  Provinsi: {provinsi}")
                                    
                                    kode_pos_list.append(kode_pos)
                            except:
                                pass
            
            if kode_pos_list:
                print(f"\n[OK] Total ditemukan {len(kode_pos_list)} kode pos")
                return kode_pos_list[0], kode_pos_list
            
            print("[ERROR] Tidak ada data kode pos ditemukan di halaman")
            return None, None
        else:
            print(f"[ERROR] HTTP Error: Status {response.status_code}")
            return None, None
            
    except Exception as e:
        print(f"[ERROR] Error: {e}")
        return None, None

def cari_kodepos_dengan_fallback(kode_wilayah):
    """Cari kode pos dengan web scraping"""
    
    print("\n" + "="*60)
    print("PENCARIAN KODE POS")
    print("="*60)
    
    kode_original = str(kode_wilayah).strip()
    print(f"\n[INPUT] Kode wilayah: {kode_original}")
    
    kodepos, results = cari_kodepos_web(kode_original)
    
    if kodepos:
        print("\n" + "="*60)
        print(f"[OK] HASIL AKHIR: Kode Pos = {kodepos}")
        if len(results) > 1:
            print(f"[INFO] Total {len(results)} kode pos ditemukan")
        print("="*60)
        return kodepos
    
    print("\n" + "="*60)
    print("[ERROR] HASIL AKHIR: Kode pos tidak ditemukan")
    print("="*60)
    
    return None

if __name__ == "__main__":
    kode_test = "3603222011"
    print(f"\nTEST PENCARIAN KODE POS")
    print(f"Kode Wilayah: {kode_test}")
    
    hasil = cari_kodepos_dengan_fallback(kode_test)
