# FileShift: Local-First File Converter

Aplikasi berbasis web modern dan alat command-line untuk konversi berkas berkecepatan tinggi dengan pemrosesan lokal mandiri (offline-ready) dan antarmuka berdesain **Bento Grid**.

## Fitur Utama

- **Konversi Dua Arah Lengkap**:
  - **XLSX ➔ CSV**: Konversi spreadsheet Microsoft Excel ke format teks CSV dengan penanganan multi-sheet dan pemisah (delimiter) kustom.
  - **CSV ➔ XLSX**: Transformasi file data CSV menjadi workbook Microsoft Excel berstandar OpenXML (ECMA-376).
  - **PDF ➔ Word (DOCX)**: Ekstraksi tata letak, judul, paragraf, dan pemisah halaman dari dokumen PDF menjadi dokumen Microsoft Word yang dapat diedit.
  - **Word (DOCX) ➔ PDF**: Kompilasi dokumen Microsoft Word dengan format heading, bobot font, dan tabel menjadi dokumen standar PDF.
- **Desain Antarmuka Bento Grid**:
  - Tata letak modular kartu dengan proporsi fungsional yang terstruktur rapi.
  - Dukungan drag-and-drop file dengan deteksi otomatis format input dan target.
  - Pratinjau langsung struktur file (tabel data untuk CSV/XLSX, ringkasan halaman dan teks untuk PDF/Word).
  - Riwayat konversi tersimpan pada sesi lokal dengan fitur unduh ulang seketika.
- **Aksesibilitas & Tampilan**:
  - Dukungan mode Terang (Light) dan Gelap (Dark) dengan deteksi preferensi sistem.
  - Rasio kontras teks memenuhi standar WCAG AA (>= 4.5:1).
  - Navigasi keyboard penuh (`Tab`, `Enter`, `Space`, `Esc`).
- **Privasi & Keamanan Penuh**:
  - Seluruh file diproses secara lokal di server/lingkungan Anda sendiri tanpa transmisi data ke API cloud pihak ketiga.

---

## Menjalankan Aplikasi Web

Pastikan Python 3.10+ dan utilitas dokumen (`poppler-utils`, `ghostscript`) terinstal di sistem Anda:

```bash
# Jalankan server web FileShift
python3 server.py 8080
```

Buka peramban web dan akses:
```
http://localhost:8080
```

---

## Penggunaan Command Line (CLI)

Anda juga dapat melakukan konversi langsung dari terminal:

```bash
# 1. Konversi XLSX ke CSV
python3 -m converter.cli data.xlsx data.csv

# 2. Konversi CSV ke XLSX dengan pemisah titik koma
python3 -m converter.cli --delimiter ";" transaksi.csv transaksi.xlsx

# 3. Konversi PDF ke Word (DOCX)
python3 -m converter.cli dokumen.pdf hasil.docx

# 4. Konversi Word (DOCX) ke PDF
python3 -m converter.cli laporan.docx laporan.pdf
```

---

## Dokumentasi REST API

### 1. Cek Status Mesin
- **Endpoint**: `GET /api/health`
- **Respons**:
  ```json
  {
    "status": "online",
    "app": "FileShift File Converter",
    "version": "1.0.0",
    "supported_conversions": [
      {"from": "xlsx", "to": "csv"},
      {"from": "csv", "to": "xlsx"},
      {"from": "pdf", "to": "docx"},
      {"from": "docx", "to": "pdf"}
    ]
  }
  ```

### 2. Konversi Berkas
- **Endpoint**: `POST /api/convert`
- **Headers (atau Multipart Form)**:
  - `X-From-Format`: `xlsx` | `csv` | `pdf` | `docx`
  - `X-To-Format`: `csv` | `xlsx` | `docx` | `pdf`
  - `X-Filename`: nama file asli
  - `X-Delimiter`: (opsional) `,` | `;` | `\t`
- **Respons**: File biner hasil konversi dengan header `Content-Disposition: attachment; filename="..."`.

### 3. Inspeksi Struktur Berkas
- **Endpoint**: `POST /api/inspect`
- **Body**: File biner atau multipart form
- **Respons**: JSON berisi metadata baris, halaman, dan data pratinjau tabel/dokumen.

---

## Menjalankan Pengujian Otomatis

Proyek ini dilengkapi dengan unit test dan integration test mandiri:

```bash
python3 -m unittest discover tests
```
