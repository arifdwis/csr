import os
import re
import pandas as pd
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_FILE = os.path.join(BASE_DIR, "bahan", "DATA CSR PEMKOT SAMARINDA.xlsx")
OUTPUT_JSON = os.path.join(BASE_DIR, "output", "data_cleaned.json")

# Mapping Kelurahan ke Kecamatan Induk di Kota Samarinda
KELURAHAN_TO_KECAMATAN = {
    # Palaran
    'Bantuas': 'Kec. Palaran', 'Bukuan': 'Kec. Palaran', 'Handil Bakti': 'Kec. Palaran', 
    'Rawa Makmur': 'Kec. Palaran', 'Simpang Pasir': 'Kec. Palaran',
    # Samarinda Seberang
    'Baqa': 'Kec. Samarinda Seberang', 'Gunung Panjang': 'Kec. Samarinda Seberang', 
    'Mangkupalas': 'Kec. Samarinda Seberang', 'Mesjid': 'Kec. Samarinda Seberang', 'Tenun': 'Kec. Samarinda Seberang',
    # Loa Janan Ilir
    'Harapan Baru': 'Kec. Loa Janan Ilir', 'Rapak Dalam': 'Kec. Loa Janan Ilir', 'Tani Aman': 'Kec. Loa Janan Ilir',
    # Samarinda Ulu
    'Teluk Lerong Ulu': 'Kec. Samarinda Ulu',
    # Sungai Kunjang
    'Karang Asam Ilir': 'Kec. Sungai Kunjang', 'Loa Bakung': 'Kec. Sungai Kunjang', 
    'Loa Buah': 'Kec. Sungai Kunjang', 'Lok Bahu': 'Kec. Sungai Kunjang',
    # Samarinda Kota
    'Bugis': 'Kec. Samarinda Kota', 'Karang Mumus': 'Kec. Samarinda Kota', 
    'Pelabuhan': 'Kec. Samarinda Kota', 'Sungai Pinang Luar': 'Kec. Samarinda Kota',
    # Sungai Pinang
    'Gunung Lingai': 'Kec. Sungai Pinang', 'Mugirejo': 'Kec. Sungai Pinang', 'Sungai Pinang Dalam': 'Kec. Sungai Pinang',
    # Samarinda Utara
    'Budaya Pampang': 'Kec. Samarinda Utara', 'Lempake': 'Kec. Samarinda Utara', 
    'Sempaja Barat': 'Kec. Samarinda Utara', 'Sempaja Selatan': 'Kec. Samarinda Utara', 
    'Sempaja Timur': 'Kec. Samarinda Utara', 'Sempaja Utara': 'Kec. Samarinda Utara', 
    'Sungai Siring': 'Kec. Samarinda Utara', 'Tanah Merah': 'Kec. Samarinda Utara',
    # Sambutan
    'Makroman': 'Kec. Sambutan', 'Pulau Atas': 'Kec. Sambutan', 
    'Sambutan': 'Kec. Sambutan', 'Sungai Kapih': 'Kec. Sambutan',
    # Samarinda Ilir
    'Selili': 'Kec. Samarinda Ilir', 'Sidomulyo': 'Kec. Samarinda Ilir'
}

def classify_perangkat_daerah(raw_name):
    name = raw_name.strip()
    if name == 'Karang Asam Ilir':
        name = 'Kelurahan Karang Asam Ilir'
        
    lower = name.lower()
    
    if lower.startswith('kelurahan') or 'kel.' in lower:
        tipe = 'Kelurahan'
        base = name.replace('Kelurahan', '').replace('Kel.', '').strip()
        kec = KELURAHAN_TO_KECAMATAN.get(base, 'Kecamatan Lainnya')
        return name, tipe, kec
    elif lower.startswith('kecamatan') or 'kec.' in lower:
        tipe = 'Kecamatan'
        kec = name.replace('Kecamatan', 'Kec.').strip()
        return name, tipe, kec
    elif any(lower.startswith(k) or lower == k for k in ['bpbd', 'bapperida', 'bpkad', 'bapenda', 'bkpsdm', 'badan']):
        tipe = 'Badan'
        return name, tipe, 'Tingkat Kota (Badan)'
    else:
        tipe = 'Dinas'
        return name, tipe, 'Tingkat Kota (Dinas)'

def parse_currency(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s in ['-', '', 'NIHIL', 'nihil', 'Nihil', 'None', 'nan']:
        return 0.0
    s = s.replace('Rp', '').replace(' ', '').replace(',', '')
    try:
        return float(s)
    except:
        m = re.findall(r'\d+(?:\.\d+)?', s)
        if m:
            return float(m[0])
        return 0.0

def categorize_sector(uraian, sasaran=''):
    txt = f"{uraian} {sasaran}".lower()
    
    # Check Stunting & Kesehatan
    if any(k in txt for k in ['stunting', 'balita', 'posyandu', 'kesehatan', 'pmt', 'donor darah', 'puskesmas', 'ibu hamil']):
        return 'Kesehatan & Penanganan Stunting'
    
    # Check Infrastruktur & Fisik
    if any(k in txt for k in ['excavator', 'drainase', 'sumur bor', 'tandon', 'air bersih', 'penerangan', 'lampu', 'jalan', 'jembatan', 'kanopi', 'semen', 'pagar', 'taman', 'infrastruktur', 'rehab', 'bangunan', 'pos kamling', 'aspal', 'paving']):
        return 'Infrastruktur & Lingkungan'
        
    # Check Keagamaan
    if any(k in txt for k in ['masjid', 'musholla', 'mushola', 'ibadah', 'kurban', 'hewan kurban', 'mtq', 'gereja']) or re.search(r'\bsapi\b', txt):
        return 'Keagamaan & Sarana Ibadah'

    # Check Kebencanaan & Damkar
    if any(k in txt for k in ['banjir', 'bencana', 'damkar', 'rescue', 'logistik', 'perahu karet', 'perahu']):
        return 'Kebencanaan & Kedaruratan'

    # Check Pendidikan & SDM (with safe word matching)
    if any(k in txt for k in ['sekolah', 'guru', 'siswa', 'pendidikan', 'pelatihan', 'coding', 'bimtek', 'buku', 'beasiswa', 'perpustakaan', 'workshop']) or re.search(r'\bai\b', txt):
        return 'Pendidikan & Pengembangan SDM'

    # Check Sosial & Bantuan Pangan
    if any(k in txt for k in ['sembako', 'beras', 'minyak', 'gula', 'santunan', 'kurang mampu', 'miskin', 'dhuafa', 'uang tunai setiap bulan', 'bantuan tunai', 'yatim']):
        return 'Sosial & Bantuan Pangan'

    # Check Pemberdayaan Ekonomi & UMKM
    if any(k in txt for k in ['tenun', 'umkm', 'usaha', 'ekonomi kreatif', 'desa wisata', 'modal', 'kelompok usaha', 'pengrajin']):
        return 'Pemberdayaan UMKM & Ekonomi Kreatif'

    # Check Event & Acara
    if any(k in txt for k in ['hut', 'jalan santai', 'agustusan', 'lomba', 'hadiah', 'festival', 'olahraga', 'tuan rumah']):
        return 'Event & Kegiatan Masyarakat'

    return 'Lainnya / Fasilitas Umum'

def classify_bentuk(uraian, amt=0):
    u = str(uraian).lower().strip()
    if re.search(r'\b(uang|dana|tunai|kas|beasiswa|santunan)\b', u):
        return 'Uang'
    return 'Barang'

def categorize_funder(funder):
    if not funder:
        return 'Tidak Disebutkan'
    f = str(funder).lower()
    if any(k in f for k in ['bank', 'pln', 'telkom', 'pertamina', 'pelindo', 'kaltimtara', 'mandiri', 'bri', 'bni', 'btn', 'permata', 'bca', 'pt. pos']):
        return 'BUMN / BUMD / Perbankan'
    if any(k in f for k in ['baznas', 'zakat', 'bmm', 'yayasan', 'baitul maal', 'lazismu']):
        return 'Lembaga Filantropi & Zakat'
    if any(k in f for k in ['bapak', 'ibu', 'h. ', 'h.', 'cafe', 'toko', 'warung', 'sabina', 'aladin', 'teluk bajo', 'agung priyono', 'suharso', 'aas amelia', 'warga']):
        return 'Pelaku Usaha Lokal & Perorangan'
    if any(k in f for k in ['pgri', 'komura', 'karang taruna', 'lpm', 'forum', 'gabungan']):
        return 'Organisasi & Komunitas'
    return 'Perusahaan Swasta & Industri'

def load_and_clean_data(filepath=EXCEL_FILE):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    all_records = []
    rec_id = 1

    sheets_to_process = ['2022', '2023', '2024', '2025', '2026', '2024-2029']
    
    for sname in sheets_to_process:
        if sname not in wb.sheetnames:
            continue
        ws = wb[sname]
        curr_pd = 'Pemerintah Kota Samarinda'
        
        for r in range(7, ws.max_row + 1):
            pd_c = ws.cell(r, 3).value
            uraian_c = ws.cell(r, 4).value
            tahun_c = ws.cell(r, 5).value
            dana_c = ws.cell(r, 6).value
            pemberi_c = ws.cell(r, 7).value
            koordinasi_c = ws.cell(r, 8).value
            sasaran_c = ws.cell(r, 9).value
            
            if pd_c is not None and str(pd_c).strip() not in ['', '-']:
                curr_pd = str(pd_c).strip()
                
            has_content = any(x is not None and str(x).strip() not in ['', '-'] for x in [uraian_c, pemberi_c, dana_c])
            if uraian_c and any(w in str(uraian_c).upper() for w in ['TOTAL', 'JUMLAH']):
                continue
            if dana_c and any(w in str(dana_c).upper() for w in ['TOTAL', 'JUMLAH']):
                continue
                
            if has_content:
                uraian = str(uraian_c).strip() if uraian_c else '-'
                if uraian.upper() in ['NIHIL', 'TIDAK ADA', '-']:
                    continue
                    
                amt = parse_currency(dana_c)
                pemberi = str(pemberi_c).strip() if pemberi_c else 'Tidak Disebutkan'
                koordinasi = str(koordinasi_c).strip() if koordinasi_c else '-'
                sasaran = str(sasaran_c).strip() if sasaran_c else '-'
                
                # Classify into Dinas, Badan, Kecamatan, or Kelurahan + Parent Kecamatan
                pd_clean, tipe_pd, wilayah_kec = classify_perangkat_daerah(curr_pd)
                    
                sektor = categorize_sector(uraian, sasaran)
                tipe_pemberi = categorize_funder(pemberi)
                bentuk_bantuan = classify_bentuk(uraian, amt)
                
                is_outlier = (amt >= 10000000000 and 'perahu' in uraian.lower())
                is_zero = (amt == 0.0)
                is_individual = (tipe_pemberi == 'Pelaku Usaha Lokal & Perorangan')
                is_event = (sektor == 'Event & Kegiatan Masyarakat')
                is_asset = any(k in uraian.lower() for k in [
                    'excavator', 'perahu', 'kanopi', 'sumur bor', 'tandon', 'bangunan', 
                    'rehab', 'pagar', 'lampu', 'mesin', 'showcase'
                ])
                
                risk_notes = []
                if is_outlier:
                    risk_notes.append('Nilai Ekstrem (Perahu Karet Rp 15 Miliar - Disparitas Valuasi Wajar)')
                if is_zero:
                    risk_notes.append('Tanpa Valuasi Rupiah (Perlu Standar Penetapan Nilai BAST)')
                if is_individual:
                    risk_notes.append('Sumbangan Perorangan/Warga (Di luar definisi entitas TSP)')
                if is_event:
                    risk_notes.append('Sponsor Acara/HUT (Terkait SE Larangan Gratifikasi)')
                if is_asset:
                    risk_notes.append('Barang Fisik Modal (Wajib Pencatatan KIB/Neraca LKPD)')
                    
                all_records.append({
                    'ID': rec_id,
                    'Lembar_Sheet': sname,
                    'Tahun_Pelaporan': sname,
                    'Tahun_Detail': str(tahun_c).strip() if tahun_c else sname,
                    'Perangkat_Daerah': pd_clean,
                    'Tipe_Perangkat_Daerah': tipe_pd, # 'Dinas', 'Badan', 'Kecamatan', 'Kelurahan'
                    'Wilayah_Kecamatan': wilayah_kec, # 'Kec. Samarinda Utara', etc.
                    'Uraian_Bantuan': uraian,
                    'Bentuk_Bantuan': bentuk_bantuan, # 'Barang / Fasilitas Fisik', 'Uang Tunai', 'Jasa / Pelatihan'
                    'Kategori_Sektor': sektor,
                    'Jumlah_Dana': amt,
                    'Status_Nilai': 'Dengan Nilai Nominal' if amt > 0 else 'Non-Nominal / Hibah Barang',
                    'Nama_Pemberi': pemberi,
                    'Tipe_Pemberi': tipe_pemberi,
                    'Bentuk_Koordinasi': koordinasi,
                    'Sasaran_Hasil': sasaran,
                    'Is_Outlier': is_outlier,
                    'Is_Zero': is_zero,
                    'Potensi_Aset_KIB': is_asset,
                    'Catatan_Risiko': '; '.join(risk_notes) if risk_notes else 'Sesuai Prosedur Standar'
                })
                rec_id += 1

    df = pd.DataFrame(all_records)
    os.makedirs(os.path.dirname(OUTPUT_JSON), exist_ok=True)
    df.to_json(OUTPUT_JSON, orient='records', indent=2, force_ascii=False)
    return df

if __name__ == '__main__':
    df = load_and_clean_data()
    print("Dataset cleaned successfully!")
    print(f"Total rows: {len(df)}")
    print(f"Tipe Perangkat Daerah breakdown:\n{df['Tipe_Perangkat_Daerah'].value_counts()}")
    print(f"\nWilayah Kecamatan breakdown:\n{df['Wilayah_Kecamatan'].value_counts()}")
