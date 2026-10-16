import datetime
import json
import os
import re
from collections import defaultdict

import openpyxl
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_FILE = os.path.join(BASE_DIR, "bahan", "DATA CSR PEMKOT SAMARINDA.xlsx")
OUTPUT_JSON = os.path.join(BASE_DIR, "output", "data_cleaned.json")
OUTPUT_REKAP_JSON = os.path.join(BASE_DIR, "output", "rekap_sumber.json")

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

SHEETS_TO_PROCESS = ['2022', '2023', '2024', '2025', '2026', '2024-2029']
# Kolom Excel (1-based): B=No, C=Perangkat Daerah, D=Uraian, E=Tahun, F=Dana, G=Pemberi, H=Koordinasi, I=Sasaran
TEXT_COLS = [2, 3, 4, 5, 7, 8, 9]   # kolom teks yang boleh diisi dari sel gabungan
DANA_COL = 6                         # kolom dana TIDAK diisi dari sel gabungan agar tidak terhitung ganda
OUTLIER_THRESHOLD = 10_000_000_000   # penanda nilai sangat besar (>= Rp 10 miliar), hanya untuk skenario hitung


def parse_currency(val):
    """Kembalikan (nilai_float, status). Status: NOMINAL, KOSONG, TEKS_NON_ANGKA."""
    if val is None:
        return 0.0, 'KOSONG'
    if isinstance(val, (int, float)):
        return float(val), 'NOMINAL'
    s = str(val).strip()
    if s.lower() in ['', '-', 'nihil', 'none', 'nan']:
        return 0.0, 'KOSONG'
    t = re.sub(r'[Rr]p\.?', '', s).replace(' ', '')
    if re.fullmatch(r'\d{1,3}(,\d{3})+(\.\d+)?|\d+(\.\d+)?', t):      # format 1,234,567.00
        return float(t.replace(',', '')), 'NOMINAL'
    if re.fullmatch(r'\d{1,3}(\.\d{3})+(,\d+)?', t):                  # format 1.234.567,00
        return float(t.replace('.', '').replace(',', '.')), 'NOMINAL'
    return 0.0, 'TEKS_NON_ANGKA'


def has_kw(txt, keywords):
    """Pencocokan kata utuh agar 'usaha' tidak cocok dengan 'perusahaan'."""
    return any(re.search(r'\b' + re.escape(k) + r'\b', txt) for k in keywords)


def categorize_sector(uraian, sasaran='', lokasi=''):
    txt = f"{uraian} {sasaran} {lokasi}".lower()
    if has_kw(txt, ['jalan santai', 'agustusan', 'lomba', 'hadiah', 'festival', 'hut', 'olahraga', 'tuan rumah', 'peringatan']):
        return 'Event & Kegiatan Masyarakat'
    if has_kw(txt, ['internet', 'wifi', 'wi fi', 'access point', 'acess point', 'fiber optic', 'fiber optik', 'optic',
                    'switch', 'mbps', 'jaringan internet', 'free internet']):
        return 'Infrastruktur Digital & Internet'
    if has_kw(txt, ['stunting', 'balita', 'posyandu', 'kesehatan', 'pmt', 'donor darah', 'puskesmas', 'ibu hamil', 'gizi', 'lansia']):
        return 'Kesehatan & Penanganan Stunting'
    if has_kw(txt, ['banjir', 'bencana', 'damkar', 'rescue', 'kebakaran', 'perahu karet', 'perahu', 'pengungsi', 'logistik']):
        return 'Kebencanaan & Kedaruratan'
    if has_kw(txt, ['masjid', 'musholla', 'mushola', 'musala', 'langgar', 'ibadah', 'kurban', 'qurban', 'kambing', 'mtq', 'lptq', 'tpa', 'gereja', 'sapi', 'madrasah', 'pesantren']):
        return 'Keagamaan & Sarana Ibadah'
    if has_kw(txt, ['sekolah', 'guru', 'siswa', 'pendidikan', 'pelatihan', 'coding', 'bimtek', 'buku', 'beasiswa',
                    'perpustakaan', 'workshop', 'sdn', 'smp', 'smpn', 'sma', 'paud', 'tk', 'ai', 'komputer']):
        return 'Pendidikan & Pengembangan SDM'
    if has_kw(txt, ['excavator', 'drainase', 'sumur bor', 'tandon', 'air bersih', 'penerangan', 'lampu', 'jalan',
                    'jembatan', 'kanopi', 'semen', 'pagar', 'taman', 'infrastruktur', 'rehab', 'rehabilitasi',
                    'bangunan', 'pos kamling', 'aspal', 'paving', 'gazebo', 'toilet', 'pembangunan', 'renovasi',
                    'parit', 'gorong', 'semenisasi', 'high mast', 'rumah tidak layak huni', 'rtlh', 'bedah rumah',
                    'bersih desa', 'maggot', 'sampah', 'kampung iklim', 'penghijauan', 'sarana dan prasarana']):
        return 'Infrastruktur & Lingkungan'
    if has_kw(txt, ['sembako', 'beras', 'minyak', 'gula', 'santunan', 'kurang mampu', 'miskin', 'dhuafa', 'duafa',
                    'uang tunai', 'bantuan tunai', 'yatim', 'pangan', 'gpm']):
        return 'Sosial & Bantuan Pangan'
    if has_kw(txt, ['tenun', 'umkm', 'usaha', 'ekonomi kreatif', 'desa wisata', 'modal', 'kelompok usaha', 'pengrajin',
                    'hidroponik', 'pertanian', 'perikanan', 'bumrt']):
        return 'Pemberdayaan UMKM & Ekonomi'
    return 'Lainnya / Belum Terklasifikasi'


def classify_bentuk(uraian):
    u = str(uraian).lower()
    if has_kw(u, ['uang', 'dana', 'tunai', 'kas', 'beasiswa', 'santunan', 'transfer']):
        return 'Uang'
    return 'Barang/Jasa'


def categorize_funder(funder):
    """Label bantu otomatis. Bukan penetapan status hukum pemberi."""
    if not funder or funder == 'Tidak Disebutkan':
        return 'Tidak Disebutkan'
    f = str(funder).lower().strip()
    entity = re.match(r'^(pt|cv|ud|pd|fa|koperasi|perumda|perseroda)\b\.?', f) or has_kw(f, ['tbk', 'pt', 'cv'])
    if has_kw(f, ['bank', 'bri', 'bni', 'btn', 'bsi', 'pln', 'pertamina', 'telkom', 'telkomsel', 'pelindo', 'pegadaian',
                  'taspen', 'bpjs', 'pdam', 'perumda', 'bumd', 'bumn', 'kaltimtara', 'icon plus', 'patra niaga', 'pos indonesia']):
        return 'BUMN / BUMD / Perbankan'
    if has_kw(f, ['baznas', 'zakat', 'lazis', 'lazismu', 'baitul maal', 'bmm', 'yayasan', 'dompet dhuafa']):
        return 'Lembaga Filantropi & Zakat'
    if entity:
        return 'Perusahaan Swasta'
    if has_kw(f, ['toko', 'cafe', 'kafe', 'warung', 'travel', 'rumah makan', 'kedai', 'bengkel', 'apotek', 'salon',
                  'laundry', 'hotel', 'dealer', 'motor', 'indomaret', 'alfamart', 'astra', 'trakindo', 'united tractor',
                  'united tractors', 'mall', 'swalayan']):
        return 'Pelaku Usaha'
    if has_kw(f, ['relawan', 'ksr', 'komunitas', 'forum', 'karang taruna', 'lpm', 'pgri', 'ikatan', 'himpunan',
                  'paguyuban', 'organisasi', 'universitas', 'unmul', 'untag', 'its', 'pensiunan', 'perum', 'warga',
                  'kelompok', 'gabungan', 'majelis', 'mahasiswa', 'alumni', 'pramuka', 'lions', 'rotary', 'club',
                  'peduli', 'lpk', 'kampus', 'politeknik', 'sekolah', 'rt']):
        return 'Organisasi / Komunitas / Lembaga'
    if re.match(r'^(h|hj|bapak|bpk|ibu|dr|drs|ir)\b\.?', f):
        return 'Perorangan'
    return 'Belum Terklasifikasi'


def normalize_year(tahun_c, sname):
    """Ambil tahun kegiatan dari kolom TAHUN CSR. Rentang multi-tahun dikembalikan sebagai teks."""
    if isinstance(tahun_c, (datetime.date, datetime.datetime)):
        return str(tahun_c.year)
    s = str(tahun_c).strip() if tahun_c not in (None, '') else ''
    years = re.findall(r'(20\d{2})', s)
    if len(set(years)) >= 2:
        return f"{years[0]}-{years[-1]}"
    if years:
        return years[0]
    return sname


def build_merged_lookup(ws):
    lookup = {}
    for m in ws.merged_cells.ranges:
        top = ws.cell(m.min_row, m.min_col).value
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                if (r, c) != (m.min_row, m.min_col):
                    lookup[(r, c)] = top
    return lookup


def find_header_and_subtotal(ws):
    header_row = None
    for r in range(1, min(ws.max_row, 30) + 1):
        if str(ws.cell(r, 2).value).strip() == 'No':
            header_row = r
            break
    subtotal_row, subtotal_val = None, None
    if header_row:
        for r in range(header_row + 1, ws.max_row + 1):
            vals = [ws.cell(r, c).value for c in range(2, 10)]
            dana = ws.cell(r, DANA_COL).value
            others = [v for i, v in enumerate(vals) if i != DANA_COL - 2 and v not in (None, '')]
            if dana not in (None, '') and not others:
                amt, st = parse_currency(dana)
                if st == 'NOMINAL':
                    subtotal_row, subtotal_val = r, amt   # ambil baris subtotal terakhir
    return header_row, subtotal_row, subtotal_val


def read_total_sheet(wb):
    out = {}
    if 'Total' not in wb.sheetnames:
        return out
    for row in wb['Total'].iter_rows(values_only=True):
        if row and row[0] is not None and row[1] is not None:
            amt, st = parse_currency(row[1])
            if st == 'NOMINAL':
                out[str(row[0]).strip()] = amt
    return out


def norm_text(s):
    return re.sub(r'\W+', ' ', str(s or '')).strip().lower()


def load_and_clean_data(filepath=EXCEL_FILE, write_output=True):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    all_records = []
    rekap = {'sumber': os.path.basename(filepath), 'lembar': [], 'total_sheet': read_total_sheet(wb)}
    rec_id = 1

    for sname in SHEETS_TO_PROCESS:
        if sname not in wb.sheetnames:
            continue
        ws = wb[sname]
        merged = build_merged_lookup(ws)
        merged_dana_top = {r: m.min_row for m in ws.merged_cells.ranges
                           if m.min_col <= DANA_COL <= m.max_col
                           for r in range(m.min_row + 1, m.max_row + 1)}
        header_row, subtotal_row, subtotal_val = find_header_and_subtotal(ws)
        if not header_row:
            continue
        last_row = (subtotal_row - 1) if subtotal_row else ws.max_row
        curr_pd = None
        sheet_records = []

        def cell(r, c):
            v = ws.cell(r, c).value
            if v in (None, '') and c in TEXT_COLS:
                v = merged.get((r, c))
            return v

        for r in range(header_row + 1, last_row + 1):
            pd_c, uraian_c, tahun_c = cell(r, 3), cell(r, 4), cell(r, 5)
            dana_c = ws.cell(r, DANA_COL).value
            pemberi_c, koordinasi_c, sasaran_c = cell(r, 7), cell(r, 8), cell(r, 9)
            # Kolom tambahan tanpa judul di kanan kolom I (mis. rincian barang PLN Icon Plus di kolom J).
            extra_c = '; '.join(str(ws.cell(r, c).value).strip() for c in range(10, ws.max_column + 1)
                                if ws.cell(r, c).value not in (None, '') and str(ws.cell(r, c).value).strip())

            if pd_c is not None and str(pd_c).strip() not in ['', '-']:
                curr_pd = str(pd_c).strip()

            own_values = [ws.cell(r, c).value for c in range(2, 10)]
            if all(v is None or str(v).strip() in ['', '-'] for v in own_values):
                continue   # baris kosong atau sisa sel gabungan tanpa isi sendiri

            # Baris lanjutan: tidak punya uraian, dana, maupun pemberi sendiri.
            # Isinya (mis. rincian sasaran per unit) digabung ke catatan induk, bukan dihitung sebagai kegiatan baru.
            def blank(v):
                return v is None or str(v).strip() in ['', '-']
            if blank(ws.cell(r, 4).value) and blank(ws.cell(r, DANA_COL).value) and blank(ws.cell(r, 7).value) \
                    and sheet_records:
                parent = sheet_records[-1]
                for col, key in [(8, 'Bentuk_Koordinasi'), (9, 'Sasaran_Hasil')]:
                    v = ws.cell(r, col).value
                    if not blank(v) and str(v).strip() not in parent[key]:
                        parent[key] = (parent[key] + '; ' if parent[key] != '-' else '') + str(v).strip()
                parent['Baris_Lanjutan'] = parent.get('Baris_Lanjutan', 0) + 1
                continue

            uraian = str(uraian_c).strip() if uraian_c else ''
            if uraian.upper() in ['NIHIL', 'TIDAK ADA']:
                continue
            lokasi = ''
            if extra_c:
                # Pada baris seperti ini kolom URAIAN berisi nama tempat, sedangkan rincian bantuan ada di kolom tambahan.
                lokasi, uraian = uraian, extra_c
            amt, nilai_status = parse_currency(dana_c)
            if not uraian and nilai_status != 'NOMINAL' and not pemberi_c:
                continue
            if not uraian:
                uraian = '(uraian tidak tercantum)'

            pemberi = str(pemberi_c).strip() if pemberi_c else 'Tidak Disebutkan'
            koordinasi = str(koordinasi_c).strip() if koordinasi_c else '-'
            sasaran = str(sasaran_c).strip() if sasaran_c else '-'
            pd_clean, tipe_pd, wilayah_kec = classify_perangkat_daerah(curr_pd or 'Tidak Disebutkan')
            sektor = categorize_sector(uraian, sasaran, lokasi)
            tipe_pemberi = categorize_funder(pemberi)
            is_outlier = amt >= OUTLIER_THRESHOLD
            is_asset = has_kw(uraian.lower(), ['excavator', 'perahu', 'kanopi', 'sumur bor', 'tandon', 'bangunan', 'rehab',
                                               'pagar', 'lampu', 'mesin', 'showcase', 'genset', 'drone', 'gazebo',
                                               'kendaraan', 'motor', 'mobil', 'komputer', 'laptop',
                                               'access point', 'acess point', 'switch', 'optic', 'fiber optic'])

            notes = []
            if is_outlier:
                notes.append('Nilai sangat besar dibanding baris lain; perlu konfirmasi OPD sumber')
            dana_top = merged_dana_top.get(r)
            if nilai_status == 'KOSONG' and dana_top:
                notes.append(f'Nilai tergabung pada baris {dana_top} (sel gabungan)')
            elif nilai_status == 'KOSONG':
                notes.append('Nilai tidak dicantumkan')
            if nilai_status == 'TEKS_NON_ANGKA':
                notes.append(f'Kolom nilai berisi teks: "{str(dana_c).strip()}"')
            if uraian == '(uraian tidak tercantum)':
                notes.append('Uraian kosong di sumber')
            if extra_c:
                notes.append('Rincian bantuan diambil dari kolom tanpa judul di kanan kolom I; kolom URAIAN berisi lokasi penerima')
            if tipe_pemberi in ('Perorangan', 'Belum Terklasifikasi'):
                notes.append('Jenis pemberi perlu dikonfirmasi')
            if is_asset:
                notes.append('Barang fisik; referensi BAST/KIB diisi unit berwenang')

            rec = {
                'ID': rec_id,
                'Lembar_Sheet': sname,
                'Baris_Sumber': r,
                'Tahun_Pelaporan': sname,
                'Tahun_Detail': str(tahun_c).strip() if tahun_c not in (None, '') else '',
                'Tahun_Kegiatan': normalize_year(tahun_c, sname),
                'Perangkat_Daerah': pd_clean,
                'Tipe_Perangkat_Daerah': tipe_pd,
                'Wilayah_Kecamatan': wilayah_kec,
                'Uraian_Bantuan': uraian,
                'Lokasi_Penerima': lokasi,
                'Bentuk_Bantuan': classify_bentuk(uraian),
                'Kategori_Sektor': sektor,
                'Jumlah_Dana': amt,
                'Status_Nilai': ('Tergabung dengan baris lain' if (nilai_status == 'KOSONG' and dana_top) else
                                 {'NOMINAL': 'Nominal tercantum', 'KOSONG': 'Tidak dicantumkan',
                                  'TEKS_NON_ANGKA': 'Teks non-angka'}[nilai_status]),
                'Nama_Pemberi': pemberi,
                'Tipe_Pemberi': tipe_pemberi,
                'Bentuk_Koordinasi': koordinasi,
                'Sasaran_Hasil': sasaran,
                'Is_Outlier': is_outlier,
                'Is_Zero': nilai_status != 'NOMINAL' or amt == 0,
                'Potensi_Aset_KIB': is_asset,
                'Masuk_Rekap_Lembar': True,
                'Baris_Lanjutan': 0,
                'Kemungkinan_Duplikat_ID': '',
                'Catatan_Risiko': '; '.join(notes) if notes else '-',
            }
            sheet_records.append(rec)
            rec_id += 1

        # --- Rekonsiliasi lembar terhadap subtotal yang tertulis di Excel ---
        total_all = sum(x['Jumlah_Dana'] for x in sheet_records)
        rule = 'Semua baris'
        if subtotal_val is not None and abs(total_all - subtotal_val) > 1:
            same_year = sum(x['Jumlah_Dana'] for x in sheet_records if x['Tahun_Kegiatan'] == sname)
            if abs(same_year - subtotal_val) <= 1:
                rule = f'Hanya baris bertahun {sname}'
                for x in sheet_records:
                    if x['Tahun_Kegiatan'] != sname:
                        x['Masuk_Rekap_Lembar'] = False
                        extra = (f"Tercatat di lembar {sname} dengan tahun {x['Tahun_Kegiatan']}; "
                                 f"tidak termasuk subtotal lembar ini")
                        x['Catatan_Risiko'] = extra if x['Catatan_Risiko'] == '-' else extra + '; ' + x['Catatan_Risiko']
            else:
                rule = 'Tidak cocok; perlu penelusuran manual'
        in_rekap = [x for x in sheet_records if x['Masuk_Rekap_Lembar']]
        rekap['lembar'].append({
            'lembar': sname,
            'baris_subtotal_excel': subtotal_row,
            'subtotal_excel': subtotal_val,
            'total_sheet': rekap['total_sheet'].get(sname),
            'jumlah_baris': len(sheet_records),
            'jumlah_baris_dalam_rekap': len(in_rekap),
            'nilai_dalam_rekap': sum(x['Jumlah_Dana'] for x in in_rekap),
            'jumlah_baris_luar_rekap': len(sheet_records) - len(in_rekap),
            'nilai_luar_rekap': total_all - sum(x['Jumlah_Dana'] for x in in_rekap),
            'aturan_rekap': rule,
            'cocok': subtotal_val is not None and abs(sum(x['Jumlah_Dana'] for x in in_rekap) - subtotal_val) <= 1,
        })
        all_records.extend(sheet_records)

    # --- Penanda kemungkinan duplikat lintas lembar (tidak menghapus apa pun) ---
    def dup_key(x):
        # Tahun kegiatan ikut dibandingkan: sumbangan yang sama di tahun berbeda dianggap kegiatan berulang.
        return (x['Perangkat_Daerah'], norm_text(x['Uraian_Bantuan'])[:80], x['Jumlah_Dana'], norm_text(x['Nama_Pemberi']),
                x['Tahun_Kegiatan'])
    groups = defaultdict(list)
    for x in all_records:
        if x['Jumlah_Dana'] > 0:
            groups[dup_key(x)].append(x)
    for items in groups.values():
        if len({i['Lembar_Sheet'] for i in items}) < 2:
            continue
        for i in items:
            others = [o for o in items if o is not i and o['Lembar_Sheet'] != i['Lembar_Sheet']]
            if not others:
                continue
            ids = ', '.join(f"{o['ID']} (lembar {o['Lembar_Sheet']})" for o in others)
            # Hanya pasangan yang sama-sama masuk rekap yang berpotensi menggandakan total.
            if i['Masuk_Rekap_Lembar'] and any(o['Masuk_Rekap_Lembar'] for o in others):
                i['Kemungkinan_Duplikat_ID'] = ','.join(str(o['ID']) for o in others if o['Masuk_Rekap_Lembar'])
            note = f'Uraian, pemberi, dan nilai sama dengan ID {ids}'
            i['Catatan_Risiko'] = note if i['Catatan_Risiko'] == '-' else i['Catatan_Risiko'] + '; ' + note

    rekap['total_sheet_excel'] = rekap['total_sheet'].get('TOTAL')
    rekap['total_dalam_rekap'] = sum(x['Jumlah_Dana'] for x in all_records if x['Masuk_Rekap_Lembar'])
    rekap['total_semua_baris'] = sum(x['Jumlah_Dana'] for x in all_records)
    rekap['jumlah_baris'] = len(all_records)
    rekap['jumlah_baris_dalam_rekap'] = sum(1 for x in all_records if x['Masuk_Rekap_Lembar'])
    rekap['nilai_bertanda_besar_dalam_rekap'] = sum(
        x['Jumlah_Dana'] for x in all_records if x['Masuk_Rekap_Lembar'] and x['Is_Outlier'])
    rekap['jumlah_kemungkinan_duplikat'] = sum(1 for x in all_records if x['Kemungkinan_Duplikat_ID'])
    rekap['nilai_kemungkinan_duplikat'] = sum(
        x['Jumlah_Dana'] for x in all_records
        if x['Kemungkinan_Duplikat_ID'] and all(int(i) < x['ID'] for i in x['Kemungkinan_Duplikat_ID'].split(',')))
    rekap['dibuat'] = datetime.datetime.now().isoformat(timespec='seconds')

    df = pd.DataFrame(all_records)
    if write_output:
        os.makedirs(os.path.dirname(OUTPUT_JSON), exist_ok=True)
        df.to_json(OUTPUT_JSON, orient='records', indent=2, force_ascii=False)
        with open(OUTPUT_REKAP_JSON, 'w', encoding='utf-8') as f:
            json.dump(rekap, f, ensure_ascii=False, indent=2)
    return df


def load_rekap():
    if not os.path.exists(OUTPUT_REKAP_JSON):
        load_and_clean_data()
    with open(OUTPUT_REKAP_JSON, encoding='utf-8') as f:
        return json.load(f)


if __name__ == '__main__':
    df = load_and_clean_data()
    rk = load_rekap()
    print(f"Total baris: {len(df)} (dalam rekap lembar: {rk['jumlah_baris_dalam_rekap']})")
    for s in rk['lembar']:
        status = 'COCOK' if s['cocok'] else 'TIDAK COCOK'
        print(f"  {s['lembar']:>9}: rekap Rp {s['nilai_dalam_rekap']:,.0f} vs subtotal Excel Rp {s['subtotal_excel'] or 0:,.0f} "
              f"[{status}; {s['aturan_rekap']}]; di luar rekap {s['jumlah_baris_luar_rekap']} baris Rp {s['nilai_luar_rekap']:,.0f}")
    print(f"Total dalam rekap Rp {rk['total_dalam_rekap']:,.0f} vs sheet Total Rp {rk['total_sheet_excel'] or 0:,.0f}")
    print(f"Kemungkinan duplikat lintas lembar (dalam rekap): {rk['jumlah_kemungkinan_duplikat']} baris, "
          f"nilai salinan kedua Rp {rk['nilai_kemungkinan_duplikat']:,.0f}")
