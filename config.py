import os

# ====== KONFIGURASI ======
EXCEL_FILE = "data_siswa.xlsx"
CONFIG_FILE = "config.txt"
LOG_FILE = "log_upload.txt"
FOTO_FOLDER = "kartu_keluarga"
API_URL = "https://api-emis.kemenag.go.id/v1/students/ppdb"
DELAY = 3

def load_config(config_path=CONFIG_FILE):
    config = {}
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            for line in f:
                if "=" in line:
                    key, value = line.strip().split("=", 1)
                    config[key.strip().lower()] = value.strip()
    return config

# Load runtime settings
config_data = load_config(CONFIG_FILE)
token = config_data.get("token", "")
academic_year_id = config_data.get("academic_year_id", "21")
admission_date = config_data.get("admission_date", "2026-07-13")
