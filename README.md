# Direktori & Analisis Data CSR Pemerintah Kota Samarinda

Struktur repositori telah dikelompokkan secara rapi ke dalam 4 folder utama: **`bahan/`**, **`code/`**, **`viewer/`**, dan **`output/`**.

## Dokumen kebutuhan sistem

Draf [PDF gabungan SRS, use case, dan rancangan basis data](dokumen/SRS_Sistem_Pendataan_TSP_CSR_Samarinda.pdf) tersedia di folder `dokumen/`. Versi sumber SRS ada di [Markdown](dokumen/SRS_Sistem_Pendataan_TSP_CSR_Samarinda.md); diagramnya tersedia sebagai [use case](dokumen/use_case_tsp_csr.svg) dan [ERD basis data](dokumen/database_erd.svg). Dokumen masih memerlukan tinjauan dan konfirmasi ruang lingkup dari unit pemilik sistem.

---

## 🗂️ Struktur Folder Proyek

```
CSR/
├── 📁 bahan/               <-- SEMUA BAHAN DATA MENTAH & DOKUMEN INPUT
│   ├── DATA CSR PEMKOT SAMARINDA.xlsx    (File Excel sumber mentah dari Pemkot)
│   └── IMG_4891.HEIC                     (Materi slide arahan audit Inspektorat Kota)
│
├── 📁 code/                <-- SEMUA SKRIP PEMROSESAN / DATA ENGINEERING (CODER)
│   ├── data_loader.py                    (ETL normalisasi, klasifikasi 4 OPD, tagging sektor & risiko)
│   └── export_clean_excel.py             (Generator otomatis workbook Excel pasca-normalisasi)
│
├── 📁 viewer/              <-- SEMUA KOMPONEN TAMPILAN DASHBOARD (VIEWER)
│   ├── index.html                        (Antarmuka web dashboard eksekutif)
│   └── dashboard_server.py               (Server lokal Python untuk viewer & REST API)
│
├── 📁 output/              <-- SEMUA HASIL BERSIH PASCA-NORMALISASI
│   ├── DATA_CSR_SAMARINDA_BERSIH_TERNORMALISASI.xlsx  (Workbook Excel bersih 5 Sheet)
│   └── data_cleaned.json                 (806 baris hasil olahan; 1 nilai anomali ditandai untuk verifikasi)
│
├── run_dashboard.sh        <-- Skrip satu-klik peluncur dashboard di browser
├── requirements.txt        <-- Kebutuhan pustaka Python (pandas, openpyxl)
└── README.md               <-- Dokumentasi proyek
```

---

## 🚀 Cara Menjalankan Viewer Dashboard

Cukup jalankan skrip peluncur dari folder utama:
```bash
./run_dashboard.sh
```
Lalu buka peramban di: **`http://localhost:8501`**

---

## 📊 Isi File Excel Bersih (`output/DATA_CSR_SAMARINDA_BERSIH_TERNORMALISASI.xlsx`)

1. **`DATA_LENGKAP_TERNORMALISASI`**: 806 baris hasil olahan dengan 18 kolom atribut standar. Jika satu nilai anomali dikecualikan sementara, tersisa 805 baris; nilai koreksi tetap perlu diverifikasi.
2. **`REKAP_PER_OPD`**: Rekapitulasi per 52 Perangkat Daerah (10 Dinas, 1 Badan, 3 Kecamatan, 38 Kelurahan) membandingkan dana riil vs tercatat.
3. **`REKAP_TINGKAT_ENTITAS`**: Ringkasan alokasi per 4 tingkatan pemerintahan (Dinas, Badan, Kecamatan, Kelurahan).
4. **`REKAP_SEKTOR_CSR`**: Rekapitulasi penyaluran CSR per bidang pembangunan.
5. **`AUDIT_MUTU_INSPEKTORAT`**: Matriks 4 temuan tata kelola data Inspektorat beserta rekomendasi perbaikannya.
