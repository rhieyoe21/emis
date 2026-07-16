# AGENTS.md

## Setup

Install dependencies:
```
pip install pandas openpyxl requests requests-toolbelt
```

## Running the application

- `python upload.py` - Upload student data to EMIS API
- `python kode_wilayah.py` - Interactive region code lookup tool

## Critical files

- `config.txt` - Contains API token, academic_year_id, and admission_date (required)
- `data_siswa.xlsx` - Student data spreadsheet (use data_siswa_example.xlsx as template)
- `kartu_keluarga/` - Folder containing family card photos (`.jpg` files) - optional if using URLs
- `kartu_keluarga` column in Excel - Can contain URLs to family card images/PDFs (alternative to local files)
- `wilayah.csv` - Region code database for kode_wilayah.py

## upload.py behavior

- **Kartu keluarga (family card) handling** - supports two modes:
  - **Local files**: Place `{full_name}.jpg` in `kartu_keluarga/` folder
  - **URLs**: Add `kartu_keluarga` column in Excel with image/PDF URLs (http:// or https://)
  - Script checks URL column first, falls back to local folder if empty
  - Supports JPG, PNG, and PDF formats from URLs
  - Automatically downloads and cleans up temporary files
- Creates automatic backup before processing: `backup_YYYYMMDD_HHMMSS.xlsx`
- Updates `status` column to "sudah" after successful upload
- Skips rows where status is not empty or "belum"
- Validates NIK (must be exactly 16 digits), skips duplicates
- Validates dates (birth_date >= 2005, parent birth_date >= 1950)
- Validates birth_date <= admission_date
- Formats phone numbers to 62xxx format
- Interactive confirmation before each upload (Y/N prompt)
- 3-second delay between uploads
- Logs all operations to `log_upload.txt`

## config.txt format

```
academic_year_id = 21
admission_date = 2026-07-13
token = <JWT_token_here>
```

## Data validation

- NIK: exactly 16 digits
- birth_date: year >= 2005, must be <= admission_date
- father_birth_date, mother_birth_date: year >= 1950
- Phone numbers: automatically converted from 08xxx to 62xxx format
- Missing photos abort individual uploads

## Security

- `config.txt` contains sensitive JWT token - never commit
- `data_siswa.xlsx` contains PII - gitignored
- Photos in `kartu_keluarga/` contain PII - gitignored
