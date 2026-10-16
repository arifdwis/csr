# Prototipe Aplikasi Pendataan TSP/CSR Pemerintah Kota Samarinda

Prototipe ini disusun Diskominfo Kota Samarinda sebagai pengembang aplikasi. Isinya pengolahan data CSR 2022–2026, dashboard internal, paparan, dan rancangan awal sistem.

Aplikasi hanya menyajikan dan merapikan data. Aplikasi tidak melakukan audit, tidak memberi pendapat hukum, dan tidak menilai kewajaran nilai bantuan.

## Angka acuan

Hasil olahan cocok dengan subtotal setiap lembar dan sheet Total di Excel sumber.

| Ukuran | Nilai |
|---|---|
| Total nominal dalam rekap | Rp 29.333.365.268 |
| Catatan kegiatan dalam rekap | 776 |
| Perangkat daerah | 52 |
| Baris salinan di lembar 2023, di luar rekap | 148 |

Rincian per lembar ada di `output/rekap_sumber.json` dan lembar `REKONSILIASI` pada workbook hasil olahan.

## Struktur folder

```
CSR/
├── bahan/                  Data mentah dan materi rapat
│   ├── DATA CSR PEMKOT SAMARINDA.xlsx
│   └── IMG_4891.HEIC       Slide Inspektorat: enam langkah perbaikan tata kelola TSP
├── code/
│   ├── data_loader.py      Normalisasi Excel, rekonsiliasi per lembar, penanda kualitas data
│   ├── export_clean_excel.py  Workbook Excel hasil olahan
│   └── build_paparan.py    Paparan eksekutif HTML dan PDF dari data
├── viewer/
│   ├── index.html          Dashboard internal
│   └── dashboard_server.py Server lokal: /api/data dan /api/rekap
├── output/                 Hasil olahan (dibuat ulang oleh skrip)
├── dokumen/                SRS v0.2, diagram use case, ERD, dan skrip pembangun PDF
└── run_dashboard.sh
```

## Cara menjalankan

Pasang kebutuhan sekali saja:

```bash
python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt
```

Jalankan dashboard, lalu buka `http://localhost:8501`. Server hanya bisa diakses dari komputer ini.

```bash
./run_dashboard.sh
```

Buat ulang seluruh hasil olahan setelah Excel sumber berubah:

```bash
cd code
python3 data_loader.py          # data_cleaned.json dan rekap_sumber.json
python3 export_clean_excel.py   # workbook Excel
python3 build_paparan.py        # paparan HTML dan PDF (butuh Google Chrome)
cd ..
python3 dokumen/build_srs_pdf.py  # SRS HTML dan PDF
```

Diagram di folder `dokumen/` dibuat dari berkas `.puml` dengan PlantUML.

## Aturan pengolahan data

1. Kolom teks yang memakai sel gabungan diisi dari sel induknya. Kolom nilai tidak diisi ulang agar tidak terhitung ganda.
2. Baris lanjutan tanpa uraian, nilai, dan pemberi sendiri digabung ke catatan induknya.
3. Baris subtotal di bawah setiap lembar dibaca sebagai pembanding, bukan sebagai kegiatan.
4. Bila jumlah semua baris berbeda dari subtotal lembar, skrip mencoba aturan "hanya baris bertahun sama dengan lembar". Aturan ini membuat lembar 2023 cocok.
5. Duplikasi hanya ditandai bila OPD, uraian, pemberi, nilai, dan tahun kegiatan sama.
6. Catatan bernilai Rp 10 miliar atau lebih diberi tanda untuk dikonfirmasi OPD sumber. Skenario tanpa catatan itu tersedia di dashboard, tetapi bukan koreksi data.
7. Sektor, bentuk bantuan, jenis pemberi, dan barang fisik adalah label otomatis dari teks. Label ini perlu ditetapkan resmi oleh pengelola data.

## Isi workbook hasil olahan

| Lembar | Isi |
|---|---|
| KETERANGAN | Sumber, waktu olah, total, dan catatan penggunaan |
| DATA_LENGKAP | Semua catatan beserta lembar dan baris asal |
| REKONSILIASI | Subtotal Excel dibanding hasil olahan per lembar |
| REKAP_PER_OPD | Rekap per perangkat daerah |
| REKAP_TINGKAT | Rekap per tingkat: Dinas, Badan, Kecamatan, Kelurahan |
| REKAP_SEKTOR | Rekap per sektor |
| REKAP_PEMBERI | Rekap per nama pemberi |
| CATATAN_KUALITAS_DATA | Hal yang perlu dilengkapi atau dikonfirmasi |
