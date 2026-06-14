# PDF vs Excel Comparison Tool

**AMFS Life Insurance — QA Automation Team**

Tool otomatis untuk memvalidasi data **PDF Proposal/Laporan Asuransi** terhadap **data sistem (SQL Server)** secara batch per policy number. Setiap eksekusi menghasilkan folder hasil tersendiri dengan prefix datetime, berisi laporan per policy dan file test case yang sudah diupdate statusnya.

---

## Daftar Isi

- [Prerequisite](#prerequisite)
- [Instalasi](#instalasi)
- [Konfigurasi](#konfigurasi)
- [Cara Menggunakan](#cara-menggunakan)
- [Struktur Folder](#struktur-folder)
- [Struktur File Test Case](#struktur-file-test-case)
- [Alur Eksekusi](#alur-eksekusi)
- [Output per Eksekusi](#output-per-eksekusi)
- [Strategi Extract PDF](#strategi-extract-pdf)
- [Menambah Produk Baru](#menambah-produk-baru)
- [Menjalankan Unit Test](#menjalankan-unit-test)
- [Troubleshooting](#troubleshooting)
- [Status Code](#status-code)
- [Tim QA AMFS](#tim-qa-amfs)

---

## Prerequisite

- Python **3.10** ke atas
- **ODBC Driver 17 for SQL Server** terinstall di mesin
  - Download: https://aka.ms/downloadmsodbcsql
- Akses jaringan ke database SQL Server AMFS
- File PDF tersimpan di satu folder lokal

---

## Instalasi

```bash
# 1. Clone repository
git clone https://github.com/amfs-qa/pdf-excel-compare.git
cd pdf_excel_compare

# 2. Buat virtual environment
python -m venv venv

# 3. Activate virtual environment
# Windows:
venv\Scripts\activate
# Mac/Linux:
source venv/bin/activate

# 4. Install semua dependency
pip install -r requirements.txt
```

---

## Konfigurasi

### 1. Setup file .env

```bash
cp .env.example .env
```

Edit file `.env` sesuai environment:

```env
# Database SQL Server
DB_TYPE=sqlserver
DB_HOST=localhost
DB_PORT=1433
DB_NAME=nama_database
DB_USER=username
DB_PASSWORD=password
DB_DRIVER=ODBC Driver 17 for SQL Server

# Paths
PDF_FOLDER=data/pdf
TESTCASE_FILE=data/testcase/testcase.xlsx
RESULTS_FOLDER=results
```

> **Penting:** Jangan commit file `.env` ke repository. File ini sudah ada di `.gitignore`.

### 2. Letakkan file PDF

Taruh semua file PDF di folder `data/pdf/`. Nama file harus mengandung nomor polis.

```
data/pdf/
├── laporan_510-7071457_mei2026.pdf
├── laporan_510-7071458_mei2026.pdf
└── ...
```

### 3. Isi file test case

Buka `data/testcase/testcase.xlsx` dan isi kolom berikut:

| Kolom | Diisi | Keterangan |
|---|---|---|
| `test_case_id` | Manual | Format: TC-001, TC-002, ... |
| `policy_number` | Manual | Nomor polis yang akan ditest |
| `product_code` | Manual | Kode produk: `axa_mandiri`, dll |

Kolom lainnya akan diisi otomatis oleh script saat dijalankan.

---

## Cara Menggunakan

### Jalankan dengan default path (dari .env)

```bash
python main.py
```

### Jalankan dengan path test case custom

```bash
python main.py --testcase data/testcase/testcase.xlsx
```

### Lihat semua opsi

```bash
python main.py --help
```

---

## Struktur Folder

```
pdf_excel_compare/
│
├── config/
│   └── products/
│       ├── axa_mandiri.yaml      # Config produk AXA Mandiri
│       └── README.md             # Panduan menambah produk baru
│
├── data/
│   ├── pdf/                      # Letakkan semua file PDF di sini
│   └── testcase/
│       └── testcase.xlsx         # File test case master — jangan diubah langsung
│
├── results/                      # Dibuat otomatis setiap eksekusi
│   └── 20240601_143022/
│       ├── run.log               # Log lengkap eksekusi
│       ├── testcase_result.xlsx  # Test case + status terupdate
│       └── reports/
│           ├── report_510-7071457.xlsx
│           └── report_510-7071458.xlsx
│
├── src/
│   ├── config_loader.py          # Load dan validasi YAML config per produk
│   ├── db_extractor.py           # Koneksi SQL Server dan query data polis
│   ├── pdf_locator.py            # Cari file PDF berdasarkan nomor polis
│   ├── pdf_extractor.py          # Semua strategi extract data dari PDF
│   ├── excel_handler.py          # Baca, duplikat, dan update file Excel
│   ├── comparator.py             # Logic compare PDF vs DB field by field
│   └── report_generator.py       # Generate report Excel per policy
│
├── tests/
│   ├── conftest.py               # Shared fixtures untuk semua test
│   ├── test_pdf_locator.py
│   ├── test_pdf_extractor.py
│   ├── test_db_extractor.py
│   ├── test_comparator.py
│   └── test_excel_handler.py
│
├── .env                          # Kredensial dan path (tidak di-commit)
├── .env.example                  # Template .env
├── .gitignore
├── main.py                       # Entry point
├── README.md
└── requirements.txt
```

---

## Struktur File Test Case

File `testcase.xlsx` memiliki dua sheet:

**Sheet "Test Cases":**

| Kolom | Diisi oleh | Keterangan |
|---|---|---|
| `test_case_id` | Manual | ID unik: TC-001, TC-002, ... |
| `policy_number` | Manual | Nomor polis — key utama semua lookup |
| `product_code` | Manual | Kode produk untuk routing strategi PDF |
| `no_polis` | Otomatis (DB) | Nomor polis dari DB |
| `insured_name` | Otomatis (DB) | Nama tertanggung |
| `policy_status` | Otomatis (DB) | Status polis |
| `start_date` | Otomatis (DB) | Tanggal mulai asuransi |
| `product_name` | Otomatis (DB) | Nama produk |
| `sum_insured` | Otomatis (DB) | Uang pertanggungan |
| `status` | Otomatis | PASS / FAIL / ERROR |
| `mismatch_fields` | Otomatis | Field yang tidak cocok, dipisah koma |
| `notes` | Otomatis | Pesan error atau catatan |

---

## Alur Eksekusi

```
START
  │
  ▼
Baca testcase.xlsx
  │
  ▼
Buat folder results/YYYYMMDD_HHMMSS/
Duplikat testcase → testcase_result.xlsx
  │
  ▼  (per policy number)
  ├── Extract data dari SQL Server → tulis ke testcase_result.xlsx
  │
  ├── Cari PDF di data/pdf/ berdasarkan policy_number
  │     Tidak ditemukan → status FAIL, notes "File PDF tidak ditemukan" → SKIP
  │
  ├── Validasi PDF bisa dibuka
  │     Corrupt → status ERROR → SKIP
  │
  ├── Extract data dari PDF (strategi sesuai product_code)
  │     Mapping gagal → status ERROR, notes field yang hilang → SKIP
  │
  ├── Compare PDF vs DB field by field
  │
  ├── Generate report_{policy_number}.xlsx di reports/
  │
  └── Update status di testcase_result.xlsx (PASS/FAIL/ERROR)
  │
  ▼
Log summary: total PASS, FAIL, ERROR
END
```

> **Catatan:** Error pada satu policy tidak menghentikan eksekusi policy lainnya.

---

## Output per Eksekusi

Setiap kali script dijalankan, folder baru dibuat otomatis:

```
results/
└── 20240601_143022/
    ├── run.log                      # Log lengkap dengan timestamp
    ├── testcase_result.xlsx         # Semua policy + status terupdate
    └── reports/
        ├── report_510-7071457.xlsx  # Detail compare per field
        └── report_510-7071458.xlsx
```

### Format testcase_result.xlsx

Kolom `status` diisi warna otomatis:

| Status | Warna | Kondisi |
|---|---|---|
| `PASS` | 🟢 Hijau | Semua field cocok antara PDF dan DB |
| `FAIL` | 🔴 Merah | Ada field mismatch, atau PDF tidak ditemukan |
| `ERROR` | 🟡 Kuning | Script error: DB error, PDF corrupt, field tidak ditemukan |

### Format report per policy

Setiap file `report_{policy_number}.xlsx` berisi dua sheet:

- **Summary** — Satu baris per field dengan status MATCH/MISMATCH dan nilai perbandingan
- **Transaction Detail** — Perbandingan per baris transaksi (jika ada data transaksi)

---

## Strategi Extract PDF

Tool ini mendukung 7 strategi extract tergantung format PDF. Strategi dipilih per section dokumen di file YAML config produk.

| Strategi | Kapan Dipakai | Contoh Kondisi |
|---|---|---|
| `regex` | Ada label eksplisit | `No. Polis : 510-7071457` |
| `anchor_offset` | Data tanpa label, ada teks tetangga unik | Nama/alamat di bawah `Kepada Yth.` |
| `positional` | Urutan baris dalam section selalu konsisten | Section data tertanggung |
| `pattern_recognition` | Konten punya ciri khas | KTP = 16 digit, Jl. = alamat |
| `extract_table` | Tabel grid dengan border garis | Tabel rangkuman transaksi |
| `crop_columns` | Layout dua kolom yang bisa tercampur | Halaman pertama AXA Mandiri |
| `post_processing` | Tabel kompleks: multi-baris, nilai negatif | Rincian transaksi ratusan baris |

---

## Menambah Produk Baru

### 1. Salin config yang sudah ada

```bash
cp config/products/axa_mandiri.yaml config/products/nama_produk.yaml
```

### 2. Edit file YAML baru

Minimal yang perlu disesuaikan:

```yaml
metadata:
  product_name: "Nama Produk Baru"
  product_code: "nama_produk"       # harus sama dengan nama file

# Strategi per section dokumen
extract_strategy:
  header: "regex"                   # atau strategi lain sesuai layout PDF

# Pattern regex sesuai label di PDF produk ini
pdf_patterns:
  no_polis: 'No\.\s*Polis\s*:\s*(\S+)'
  # tambahkan field lain...

# Mapping field PDF ke kolom DB/Excel
excel_column_map:
  no_polis: "no_polis"
  # tambahkan mapping lain...

numeric_fields:
  - sum_insured
```

### 3. Test pattern regex

Sebelum menambahkan pattern, cek teks mentah PDF terlebih dahulu:

```python
import pdfplumber

with pdfplumber.open("data/pdf/nama_file.pdf") as pdf:
    for i, page in enumerate(pdf.pages):
        print(f"--- PAGE {i+1} ---")
        print(page.extract_text())
```

Copy teks yang relevan ke [regex101.com](https://regex101.com) dan test dengan flag `IGNORECASE`.

### 4. Gunakan di test case

Isi kolom `product_code` di `testcase.xlsx` dengan nama file YAML tanpa ekstensi:

```
product_code: nama_produk
```

Lihat panduan lengkap di `config/products/README.md`.

---

## Menjalankan Unit Test

```bash
# Semua test
pytest tests/ -v

# Dengan coverage report
pytest tests/ -v --cov=src --cov-report=term-missing

# Test file spesifik
pytest tests/test_comparator.py -v
pytest tests/test_pdf_extractor.py -v

# Test dengan nama tertentu
pytest tests/ -v -k "test_compare_policy"
```

Semua test menggunakan mock — tidak membutuhkan koneksi DB atau file PDF asli.

---

## Troubleshooting

### PDF tidak ditemukan (status FAIL)

**Gejala:** `notes = "File PDF tidak ditemukan"`

**Penyebab:** Nama file PDF tidak mengandung nomor polis, atau file ada di folder lain.

**Solusi:**
- Pastikan nama file mengandung nomor polis. Contoh yang valid:
  - `laporan_510-7071457.pdf`
  - `510-7071457_mei2026.pdf`
  - `AXA_510-7071457.pdf`
- Pastikan file ada di folder yang diset di `PDF_FOLDER` dalam `.env`

---

### Koneksi DB gagal (status ERROR)

**Gejala:** `notes = "DB Error: Gagal koneksi ke SQL Server..."`

**Penyebab:** Kredensial salah, host tidak terjangkau, atau ODBC Driver belum terinstall.

**Solusi:**
1. Cek isi file `.env` — pastikan `DB_HOST`, `DB_PORT`, `DB_USER`, `DB_PASSWORD` benar
2. Pastikan ODBC Driver 17 for SQL Server sudah terinstall:
   ```bash
   # Cek driver yang terinstall
   python -c "import pyodbc; print(pyodbc.drivers())"
   ```
3. Pastikan bisa ping ke `DB_HOST` dari jaringan yang digunakan

---

### Field tidak ditemukan di PDF (status ERROR)

**Gejala:** `notes = "Field tidak ditemukan di PDF: [nama, status_polis]"`

**Penyebab:** Regex pattern tidak cocok dengan teks aktual di PDF.

**Solusi:**
1. Jalankan script debug untuk lihat teks mentah:
   ```python
   import pdfplumber
   with pdfplumber.open("data/pdf/nama_file.pdf") as pdf:
       for page in pdf.pages:
           print(page.extract_text())
   ```
2. Perhatikan perbedaan spasi, karakter khusus, atau urutan kata
3. Update pattern di file YAML config produk yang sesuai
4. Test pattern baru di [regex101.com](https://regex101.com) sebelum disimpan

---

### Tabel tidak ter-extract

**Gejala:** Field `summary_table` atau `detail_table` kosong di hasil extract.

**Penyebab:** PDF tidak punya border garis pada tabel, sehingga `extract_table()` gagal mendeteksinya.

**Solusi:** Tambahkan `table_settings` di file YAML config produk:

```yaml
table_settings:
  vertical_strategy: "text"
  horizontal_strategy: "text"
```

---

### Nilai numerik tidak match padahal angkanya sama

**Gejala:** Field seperti `sum_insured` muncul MISMATCH padahal nilainya secara visual sama.

**Penyebab:** Format berbeda — PDF: `18.750.000,00` vs DB: `18750000`.

**Solusi:** Pastikan field tersebut ada di list `numeric_fields` di file YAML config produk:

```yaml
numeric_fields:
  - sum_insured
  - annual_premium
```

---

### Lihat daftar produk yang tersedia

```python
from src.config_loader import list_available_products
print(list_available_products())
```

---

## Status Code

| Status | Warna Excel | Kondisi |
|---|---|---|
| `PASS` | Hijau | Semua field match antara PDF dan DB |
| `FAIL` | Merah | Ada field mismatch, atau file PDF tidak ditemukan |
| `ERROR` | Kuning | Script error: DB tidak bisa diakses, PDF corrupt, field mapping gagal |

Perbedaan `FAIL` dan `ERROR`:
- **FAIL** = tool berhasil jalan, tapi data tidak cocok atau PDF tidak ada
- **ERROR** = tool tidak bisa menyelesaikan proses karena ada kesalahan teknis

---

## Tim QA AMFS

| Nama | Role |
|---|---|
| *(nama)* | QA Lead / Maintainer |
| *(nama)* | QA Engineer |

Untuk pertanyaan, bug report, atau request fitur baru — hubungi tim QA internal AMFS.

---

## Changelog

| Versi | Tanggal | Perubahan |
|---|---|---|
| 2.0.0 | 2024-06-01 | Full flow: SQL Server, multi-strategi PDF, folder result per datetime |
| 1.0.0 | 2024-01-01 | Initial release |
