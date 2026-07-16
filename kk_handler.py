import os
import requests # type: ignore
import tempfile
from config import FOTO_FOLDER
from utils import safe_str, tulis_log

def download_file_from_url(url, full_name):
    try:
        print(f"📥 Downloading kartu keluarga dari URL...")
        tulis_log(f"Download KK dari URL: {url[:100]}")
        
        response = requests.get(url, timeout=30, stream=True)
        if response.status_code != 200:
            msg = f"Gagal download file dari URL (Status: {response.status_code})"
            print(f"❌ {msg}")
            tulis_log(f"❌ {full_name} - {msg}")
            return None, None
        
        # Check file size
        file_size = int(response.headers.get('Content-Length', 0))
        if file_size > 5 * 1024 * 1024:  # 5MB
            print(f"⚠️ File besar: {file_size / 1024 / 1024:.2f}MB - upload mungkin lambat")
            tulis_log(f"⚠️ {full_name} - File size: {file_size / 1024 / 1024:.2f}MB")
        
        content_type = response.headers.get('Content-Type', '').lower()
        
        if 'pdf' in content_type or url.lower().endswith('.pdf'):
            extension = 'pdf'
            mime_type = 'application/pdf'
        elif 'image' in content_type or any(url.lower().endswith(ext) for ext in ['.jpg', '.jpeg', '.png']):
            if 'png' in content_type or url.lower().endswith('.png'):
                extension = 'png'
                mime_type = 'image/png'
            else:
                extension = 'jpg'
                mime_type = 'image/jpeg'
        else:
            extension = 'jpg'
            mime_type = 'image/jpeg'
        
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=f'.{extension}')
        temp_file.write(response.content)
        temp_file.close()
        
        print(f"✅ File downloaded: {os.path.getsize(temp_file.name) / 1024:.2f}KB")
        tulis_log(f"✅ {full_name} - Downloaded: {temp_file.name}")
        
        return temp_file.name, mime_type
    except requests.exceptions.Timeout:
        msg = "Timeout saat download file dari URL (>30s)"
        print(f"❌ {msg}")
        tulis_log(f"❌ {full_name} - {msg}")
        return None, None
    except Exception as e:
        msg = f"Error saat download file dari URL: {str(e)[:100]}"
        print(f"❌ {msg}")
        tulis_log(f"❌ {full_name} - {msg}")
        return None, None

def get_kk_file(full_name, kk_url_or_empty):
    kk_url = safe_str(kk_url_or_empty)
    
    if kk_url and (kk_url.startswith('http://') or kk_url.startswith('https://')):
        print(f"📥 Download kartu keluarga dari URL: {kk_url}")
        temp_path, mime_type = download_file_from_url(kk_url, full_name)
        if temp_path:
            try:
                foto_file = open(temp_path, "rb")
                return foto_file, temp_path, mime_type, True
            except Exception as e:
                print(f"❌ Gagal membuka file yang didownload: {e}")
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                return None, None, None, False
        else:
            return None, None, None, False
    else:
        foto_filename = os.path.join(FOTO_FOLDER, f"{full_name}.jpg")
        if not os.path.exists(foto_filename):
            print(f"❌ File foto {foto_filename} tidak ditemukan di folder lokal.")
            return None, None, None, False
        try:
            foto_file = open(foto_filename, "rb")
            return foto_file, foto_filename, 'image/jpeg', False
        except Exception as e:
            print(f"❌ Gagal membuka file foto: {e}")
            return None, None, None, False
