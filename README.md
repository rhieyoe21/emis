# EMIS Student Data Upload Tool

A Python tool for uploading student data to the EMIS (Education Management Information System) API.

## Features

- Upload student data to EMIS API from Excel spreadsheet
- Interactive confirmation before each upload
- Automatic data validation (NIK, dates, phone numbers)
- Support for local files or URLs for family card documents
- Automatic backup before processing
- Region code lookup tool

## Requirements

- Python 3.7+
- pandas
- openpyxl
- requests
- requests-toolbelt

## Installation

```bash
pip install pandas openpyxl requests requests-toolbelt
```

## Configuration

1. Copy `config.exp.txt` to `config.txt`
2. Fill in your JWT token:
   ```
   academic_year_id = 21
   admission_date = "2026-07-13"
   token = <your_jwt_token_here>
   ```

3. Prepare student data in `data_siswa.xlsx` (use `data_siswa_example.xlsx` as template)

## Usage

### Upload Student Data
```bash
python upload.py
```

### Lookup Region Codes
```bash
python kode_wilayah.py
```

## Family Card (Kartu Keluarga) Options

The tool supports two methods for providing family card documents:

### Option 1: Local Files
- Place `{full_name}.jpg` in `kartu_keluarga/` folder
- Filename must match student's full name exactly

### Option 2: URLs
- Add `kartu_keluarga` column in Excel
- Provide URLs to images or PDFs (http:// or https://)
- Supports JPG, PNG, and PDF formats

The script checks the URL column first, then falls back to local folder.

## File Structure

```
├── config.txt              # API token and settings (create from config.exp.txt)
├── data_siswa.xlsx         # Student data spreadsheet
├── data_siswa_example.xlsx # Template for student data
├── kartu_keluarga/         # Family card photos folder
├── wilayah.csv             # Region code database
├── upload.py               # Main upload script
├── kode_wilayah.py         # Region code lookup tool
└── log_upload.txt          # Upload log file
```

## Data Validation

- **NIK**: Must be exactly 16 digits, duplicates are skipped
- **birth_date**: Year >= 2005, must be <= admission_date
- **parent birth_date**: Year >= 1950
- **Phone numbers**: Automatically converted from 08xxx to 62xxx format

## Security Notes

- `config.txt` contains sensitive JWT token - **never commit**
- `data_siswa.xlsx` contains PII - **never commit**
- `kartu_keluarga/` contains PII photos - **never commit**
- These files are already in `.gitignore`

## License

MIT
