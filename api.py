import os
import time
import requests # type: ignore
from config import FOTO_FOLDER, admission_date, academic_year_id
from utils import safe_str, format_nomor_hp, format_tanggal, validasi_usia_masuk, tulis_log
from kk_handler import get_kk_file

def parse_api_response(response, full_name):
    """Parse API response dan extract error details"""
    if response is None:
        return f"❌ {full_name} | Tidak ada response dari server (timeout/network error)"
    
    status = response.status_code
    
    try:
        data = response.json()
        
        if status == 200:
            message = data.get('message', 'Success')
            return f"✅ {full_name} | Status {status} | {message}"
        
        else:
            error_msg = data.get('message', '')
            errors = data.get('errors', {})
            
            if errors:
                error_details = []
                for field, messages in errors.items():
                    if isinstance(messages, list):
                        error_details.append(f"{field}: {', '.join(messages)}")
                    else:
                        error_details.append(f"{field}: {messages}")
                error_text = " | ".join(error_details)
                return f"❌ {full_name} | Status {status} | {error_msg} | {error_text}"
            else:
                return f"❌ {full_name} | Status {status} | {error_msg}"
    
    except:
        return f"❌ {full_name} | Status {status} | Raw: {response.text[:200]}"

def upload_data(files, headers, url, max_retry=3, timeout=120):
    last_error = None
    
    # Buat copy headers untuk menghindari modifikasi global
    upload_headers = headers.copy()
    # Hapus Content-Type dari headers, biar requests.post set otomatis untuk multipart
    if 'Content-Type' in upload_headers:
        del upload_headers['Content-Type']
    
    for attempt in range(max_retry):
        try:
            response = requests.post(url, headers=upload_headers, files=files, timeout=timeout)
            return response
        except requests.exceptions.Timeout as e:
            last_error = f"Timeout setelah {timeout} detik: {str(e)}"
            print(f"⚠️ Upload timeout (percobaan {attempt+1}/{max_retry}): {timeout}s")
            tulis_log(f"⚠️ Timeout percobaan {attempt+1}: {str(e)}")
            time.sleep(2)
        except requests.exceptions.ConnectionError as e:
            last_error = f"Connection error: {str(e)}"
            print(f"⚠️ Connection error (percobaan {attempt+1}/{max_retry}): {str(e)[:100]}")
            tulis_log(f"⚠️ Connection error percobaan {attempt+1}: {str(e)[:200]}")
            time.sleep(2)
        except Exception as e:
            last_error = f"Error: {str(e)}"
            print(f"⚠️ Gagal upload (percobaan {attempt+1}/{max_retry}): {str(e)[:100]}")
            tulis_log(f"⚠️ Exception percobaan {attempt+1}: {type(e).__name__}: {str(e)[:200]}")
            time.sleep(2)
    
    tulis_log(f"❌ Gagal upload setelah {max_retry} percobaan. Error terakhir: {last_error}")
    return None

def build_files(data):
    """Membangun payload files sesuai format"""
    full_name = str(data['full_name']).strip()

    gender_raw = str(data['gender']).strip().lower()
    gender_id = "2" if gender_raw in ["perempuan", "p", "wanita", "female", "2"] else "1"

    father_phone_number_raw = data['father_phone_number']
    father_phone_number, father_have_phone = format_nomor_hp(father_phone_number_raw)

    mother_phone_number_raw = data['mother_phone_number']
    mother_phone_number, mother_have_phone = format_nomor_hp(mother_phone_number_raw)

    birth_date, valid_birth = format_tanggal(data.get("birth_date"), min_year=2005)
    father_birth_date, valid_father = format_tanggal(data.get("father_birth_date"), min_year=1950)
    mother_birth_date, valid_mother = format_tanggal(data.get("mother_birth_date"), min_year=1950)

    if not validasi_usia_masuk(birth_date, admission_date):
        print(f"❌ Tanggal lahir siswa melebihi tanggal masuk: {full_name} | birth_date: {birth_date} > admission_date: {admission_date}")
        tulis_log(f"❌ Tanggal lahir siswa melebihi tanggal masuk: {full_name} | birth_date: {birth_date} > admission_date: {admission_date}")
        return None

    kk_url = data.get('kartu_keluarga', '')
    foto_file, foto_filename, mime_type, is_temp = get_kk_file(full_name, kk_url)
    
    if not foto_file:
        print(f"❌ Kartu keluarga tidak ditemukan (cek kolom kartu_keluarga atau folder {FOTO_FOLDER})")
        return None

    try:
        foto_file.seek(0)
        kk_data = foto_file.read()
    except Exception as e:
        print(f"❌ Gagal membaca file Kartu Keluarga: {e}")
        foto_file.close()
        if is_temp and os.path.exists(foto_filename):
            os.remove(foto_filename)
        return None

    if not full_name or full_name.lower() == 'nan':
        print(f"❌ Nama siswa tidak valid, batalkan input.")
        return None
    if not valid_birth:
        print(f"❌ Tanggal lahir siswa tidak valid: {full_name} | birth_date: {data.get('birth_date')}")
        return None

    father_life_status = safe_str(data.get('father_m_life_status_id'))
    if father_life_status == "1" and not valid_father:
        print(f"❌ Tanggal lahir ayah tidak valid: {full_name} | father_birth_date: {data.get('father_birth_date')}")
        return None

    mother_life_status = safe_str(data.get('mother_m_life_status_id'))
    if mother_life_status == "1" and not valid_mother:
        print(f"❌ Tanggal lahir ibu tidak valid: {full_name} | mother_birth_date: {data.get('mother_birth_date')}")
        return None

    siblings_num = safe_str(data.get('siblings_num'))
    if not siblings_num or siblings_num.lower() == 'nan':
        siblings_num = "0"

    files = {
        'full_name': (None, data['full_name']),
        'nationality': (None, 'wni'),
        'have_nik': (None, 'true'),
        'nik': (None, safe_str(data['nik'])),
        'have_nisn': (None, 'false'),
        'm_gender_id': (None, gender_id),
        'birth_place': (None, data['birth_place']),
        'birth_date': (None, birth_date),
        'siblings_num': (None, siblings_num),
        'child_of_num': (None, safe_str(data['child_of_num'])),
        'm_residence_status_id': (None, safe_str(data['m_residence_status_id'])),
        'm_province_id': (None, safe_str(data['m_province_id'])),
        'm_city_id': (None, safe_str(data['m_city_id'])),
        'm_district_id': (None, safe_str(data['m_district_id'])),
        'm_subdistrict_id': (None, safe_str(data['m_subdistrict_id'])),
        'address': (None, data['address']),
        'postal_code_num': (None, safe_str(data['postal_code_num'])),
        'have_handphone': (None, 'false'),
        'm_level_id': (None, safe_str(data['m_level_id'])),
        'kk_num': (None, safe_str(data['kk_num'])),
        'family_head_name': (None, data['family_head_name']),
        'upload_kk': (f"{full_name}-kk.jpg", kk_data, mime_type),

        # father
        'father_full_name': (None, data['father_full_name']),
        'father_m_life_status_id': (None, safe_str(data['father_m_life_status_id'])),
        'father_nationality': (None, 'wni'),
        'father_nik': (None, safe_str(data['father_nik'])),
        'father_birth_place': (None, data['father_birth_place']),
        'father_birth_date': (None, father_birth_date),
        'father_m_last_education_id': (None, safe_str(data['father_m_last_education_id'])),
        'father_m_job_id': (None, safe_str(data['father_m_job_id'])),
        'father_m_average_income_per_month_id': (None, safe_str(data['father_m_average_income_per_month_id'])),
        'father_domicile': (None, 'Dalam Negeri'),
        'father_m_residence_status_id': (None, safe_str(data['m_residence_status_id'])),
        'father_m_province_id': (None, safe_str(data['m_province_id'])),
        'father_m_city_id': (None, safe_str(data['m_city_id'])),
        'father_m_district_id': (None, safe_str(data['m_district_id'])),
        'father_m_sub_district_id': (None, safe_str(data['m_subdistrict_id'])),
        'father_address': (None, data['address']),
        'father_postal_code': (None, safe_str(data['postal_code_num'])),
        'father_phone_number': (None, father_phone_number),
        'father_have_phone_number': (None, father_have_phone),
        'father_kk_file': (f"{full_name}-kk.jpg", kk_data, mime_type),

        # mother
        'mother_residence': (None, '1'),
        'mother_full_name': (None, data['mother_full_name']),
        'mother_m_life_status_id': (None, safe_str(data['mother_m_life_status_id'])),
        'mother_nationality': (None, 'wni'),
        'mother_nik': (None, safe_str(data['mother_nik'])),
        'mother_birth_place': (None, data['mother_birth_place']),
        'mother_birth_date': (None, mother_birth_date),
        'mother_m_last_education_id': (None, safe_str(data['mother_m_last_education_id'])),
        'mother_m_job_id': (None, safe_str(data['mother_m_job_id'])),
        'mother_m_average_income_per_month_id': (None, safe_str(data['mother_m_average_income_per_month_id'])),
        'mother_domicile': (None, 'Dalam Negeri'),
        'mother_m_residence_status_id': (None, safe_str(data['m_residence_status_id'])),
        'mother_m_province_id': (None, safe_str(data['m_province_id'])),
        'mother_m_city_id': (None, safe_str(data['m_city_id'])),
        'mother_m_district_id': (None, safe_str(data['m_district_id'])),
        'mother_m_sub_district_id': (None, safe_str(data['m_subdistrict_id'])),
        'mother_address': (None, data['address']),
        'mother_postal_code': (None, safe_str(data['postal_code_num'])),
        'mother_phone_number': (None, mother_phone_number),
        'mother_have_phone_number': (None, mother_have_phone),
        'mother_kk_file': (f"{full_name}-kk.jpg", kk_data, mime_type),

        # wali
        'wali': (None, 'Sama dengan ayah kandung'),
        'wali_full_name': (None, data['father_full_name']),
        'wali_m_life_status_id': (None, safe_str(data['father_m_life_status_id'])),
        'wali_nationality': (None, 'wni'),
        'wali_nik': (None, safe_str(data['father_nik'])),
        'wali_birth_place': (None, data['father_birth_place']),
        'wali_birth_date': (None, father_birth_date),
        'wali_m_last_education_id': (None, safe_str(data['father_m_last_education_id'])),
        'wali_m_job_id': (None, safe_str(data['father_m_job_id'])),
        'wali_m_average_income_per_month_id': (None, safe_str(data['father_m_average_income_per_month_id'])),
        'wali_domicile': (None, 'Dalam Negeri'),
        'wali_m_residence_status_id': (None, safe_str(data['m_residence_status_id'])),
        'wali_m_province_id': (None, safe_str(data['m_province_id'])),
        'wali_m_city_id': (None, safe_str(data['m_city_id'])),
        'wali_m_district_id': (None, safe_str(data['m_district_id'])),
        'wali_m_sub_district_id': (None, safe_str(data['m_subdistrict_id'])),
        'wali_address': (None, data['address']),
        'wali_postal_code': (None, safe_str(data['postal_code_num'])),
        'wali_phone_number': (None, ''),   # kosong
        'wali_have_phone_number': (None, 'false'),
        'wali_kk_file': (f"{full_name}-kk.jpg", kk_data, mime_type),

        # umum
        'm_religion_id': (None, '1'), # Agama Islam
        'm_fund_source_id': (None, '1'), # Biaya Sekolah Dibayai Orang Tua
        'm_special_need_id': (None, '1'), # Kebutuhan Khusus Biasa
        'm_disability_id': (None, '1'), # Disabilitas Tidak
        'm_transportation_id': (None, '3'), # Sepeda Motor
        'm_residence_distance_id': (None, '1'), # Jarak 1 - 5 Km
        'm_interval_time_id': (None, '1'), # Waktu Tempuh 1 - 10 Menit
        'admission_date': (None, admission_date), # Tanggal Masuk Sekolah
        'academic_year_id': (None, academic_year_id), # ID Tahun Ajaran
        'is_pontren': (None, '0'),
    }
    return files, foto_file, foto_filename, is_temp
