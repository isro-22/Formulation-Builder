# Formulation Sheet Automation

Proyek ini dipakai untuk dua pekerjaan utama dalam satu workspace Streamlit:

1. generate `Formulation Sheet` dari input terstruktur,
2. scrape template Excel lama menjadi data JSON/XLSX normalisasi.

Engine perhitungan dan rendering Excel tetap berada di `formulation_generator.py`, sementara `app.py` sekarang menjadi pintu utama untuk kedua workflow tersebut.

## Ringkasan

Sistem ini dibuat untuk menggantikan proses Excel/VBA lama yang sulit dirawat, terutama pada:

- layout dinamis,
- formula premix dan BOM bertingkat,
- approval block,
- normalisasi output ke model data yang bisa dipakai ulang.

Hasil akhirnya adalah workbook Excel yang lebih stabil, plus model data JSON/XLSX untuk validasi, backup, dan integrasi lanjutan.

## Fitur Utama

- UI Streamlit untuk input metadata produk, phase metadata, material, dan approval
- Generate workbook formulasi dari template `templates/Template_Generate.xlsm`
- Download model formulasi dalam JSON dan XLSX
- Scrape template Excel lama dari `template_1.xlsx` atau file upload
- Export hasil scraping ke `outputs/scraped` dalam format JSON dan XLSX
- Validasi data material, phase, dan header sebelum generate
- Support mode input `mg/stick` dan `%`
- Unit test untuk engine generator

## Struktur Project

```text
03 Formulation Sheet Creation/
├── app.py
├── formulation_generator.py
├── template_scraper.py
├── requirements.txt
├── README.md
├── templates/
│   └── Template_Generate.xlsm
├── references/
│   ├── 02._SAAT_777.xlsx
│   └── BOL-SAAT List.xlsx
├── outputs/
│   ├── generated/
│   ├── scraped/
│   └── legacy/
├── assets/
└── tests/
```

## Peran File Utama

### `app.py`

UI utama Streamlit. Dari sini user bisa memilih workspace:

- `Formulation Generator`
- `Template Scraper`

### `formulation_generator.py`

Engine inti untuk:

- model data formulasi,
- validasi dan transformasi input,
- kalkulasi dosage, ratio, premix, dan harga,
- render workbook output,
- export model formulasi ke JSON/XLSX,
- pembuatan blank template `INPUT`.

### `template_scraper.py`

Modul untuk membaca template Excel legacy, lalu mengekstrak:

- product header,
- phase metadata,
- material rows,
- summary,
- approval block.

File ini sekarang bisa dipakai langsung sendiri, atau dirender dari `app.py`.

## Setup

### Prasyarat

- Python 3.8+
- terminal / shell

### Install

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Menjalankan Aplikasi

```bash
streamlit run app.py
```

Secara default aplikasi akan terbuka di localhost Streamlit, biasanya `http://localhost:8501`.

## Workflow 1: Formulation Generator

1. Jalankan `app.py`.
2. Pilih workspace `Formulation Generator`.
3. Upload file `INPUT` jika ingin memuat draft Excel.
4. Lengkapi metadata produk, approval, phase metadata, dan material.
5. Klik `Save Data / Validate Draft` untuk cek issue.
6. Klik `Generate Formulation Sheet` untuk membuat workbook final.

Output akan disimpan ke:

- `outputs/generated`

## Workflow 2: Template Scraper

1. Jalankan `app.py`.
2. Pilih workspace `Template Scraper`.
3. Gunakan `template_1.xlsx` default atau upload file template lama.
4. Pilih sheet yang ingin dibaca.
5. Review hasil scraping.
6. Download atau ambil hasil yang tersimpan otomatis.

Output akan disimpan ke:

- `outputs/scraped`

## Referensi Data

Material lookup dan harga chemical dibaca dari file:

`references/BOL-SAAT List.xlsx`

File ini dipakai generator sebagai database reference untuk material master.

## Relasi Database

Export model JSON/XLSX memakai tiga tabel utama:

- `product`
- `phase_metadata`
- `materials`

Relasi antar sheet dibuat lewat key berikut:

- `product.product_id` menjadi primary key product.
- `phase_metadata.product_id` mengarah ke `product.product_id`.
- `phase_metadata.phase_id` menjadi primary key phase.
- `materials.product_id` mengarah ke `product.product_id`.
- `materials.phase_id` mengarah ke `phase_metadata.phase_id`.
- `materials.material_id` menjadi primary key material row.

Dengan struktur ini, database berisi 50 product cukup dibuat dengan menggabungkan semua row `product`, `phase_metadata`, dan `materials` dari masing-masing model. Saat ingin memanggil satu product tertentu, filter dulu `product.product_id`, lalu ambil semua phase dan material dengan `product_id` yang sama.

## Testing

Jalankan test engine dengan:

```bash
python -m unittest tests/test_formulation_generator.py
```

## Catatan Teknis

- `app.py` sekarang menggabungkan workflow generator dan scraper.
- `template_scraper.py` sudah di-refactor agar aman di-import dari `app.py`.
- `formulation_generator.py` tetap menjadi source of truth untuk logika generate workbook.
- Output workbook sangat bergantung pada layout template Excel, jadi perubahan template perlu diuji ulang.

## Next Development

Area yang masuk akal untuk dilanjutkan:

- versioning formulasi,
- integrasi database,
- export PDF,
- approval workflow,
- deployment internal.
