"""Membuat workbook Excel bersih dari data CSR yang sudah dinormalisasi.

Semua angka dihitung dari data, tidak ada nilai yang ditulis tetap di skrip.
Rekap memakai baris yang masuk subtotal lembar Excel sumber (Masuk_Rekap_Lembar = True).
"""
import datetime
import os

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from data_loader import load_and_clean_data, load_rekap

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_FILE = os.path.join(BASE_DIR, "output", "DATA_CSR_SAMARINDA_BERSIH_TERNORMALISASI.xlsx")

FONT = Font(name='Calibri', size=10)
FONT_BOLD = Font(name='Calibri', size=10, bold=True)
FONT_HEAD = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
FILL_HEAD = PatternFill('solid', start_color='1E3A8A')
FILL_TOTAL = PatternFill('solid', start_color='E2E8F0')
FILL_FLAG = PatternFill('solid', start_color='FEF3C7')
FILL_OUT = PatternFill('solid', start_color='F1F5F9')
SIDE = Side(border_style='thin', color='CBD5E1')
BORDER = Border(left=SIDE, right=SIDE, top=SIDE, bottom=SIDE)
RUPIAH = '#,##0'


def rp(v):
    return 'Rp ' + f"{v:,.0f}".replace(',', '.')


def write_table(ws, headers, rows, money_cols=(), pct_cols=(), widths=None, total_row=None, row_fill=None):
    """Tulis tabel dengan header, format angka per nama kolom, dan baris total opsional."""
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(1, c)
        cell.font, cell.fill, cell.border = FONT_HEAD, FILL_HEAD, BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 30
    money_idx = {headers.index(h) + 1 for h in money_cols}
    pct_idx = {headers.index(h) + 1 for h in pct_cols}
    for i, row in enumerate(rows, start=2):
        ws.append(list(row))
        fill = row_fill(i - 2) if row_fill else None
        for c in range(1, len(headers) + 1):
            cell = ws.cell(i, c)
            cell.font, cell.border = FONT, BORDER
            cell.alignment = Alignment(vertical='top', wrap_text=isinstance(cell.value, str) and len(cell.value) > 40)
            if c in money_idx:
                cell.number_format = RUPIAH
            if c in pct_idx:
                cell.number_format = '0.00%'
            if fill:
                cell.fill = fill
    last = len(rows) + 1
    if total_row:
        r = last + 1
        ws.append([v.format(last=last) if isinstance(v, str) else v for v in total_row])
        for c in range(1, len(headers) + 1):
            cell = ws.cell(r, c)
            cell.font, cell.border, cell.fill = FONT_BOLD, BORDER, FILL_TOTAL
            if c in money_idx:
                cell.number_format = RUPIAH
            if c in pct_idx:
                cell.number_format = '0.00%'
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last}"
    for i, w in enumerate(widths or [], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def create_clean_excel():
    print("Memuat dan menormalisasi data sumber...")
    df = load_and_clean_data()
    rk = load_rekap()
    rek = df[df['Masuk_Rekap_Lembar']]
    total_rekap = rek['Jumlah_Dana'].sum()
    flagged = rek[rek['Is_Outlier']]['Jumlah_Dana'].sum()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # 1. Keterangan
    ws = wb.create_sheet('KETERANGAN')
    info = [
        ('Judul', 'Data CSR/TSP Pemkot Samarinda hasil normalisasi'),
        ('Sumber', rk['sumber']),
        ('Dibuat', datetime.datetime.now().strftime('%d-%m-%Y %H:%M')),
        ('Jumlah baris hasil olahan', len(df)),
        ('Baris yang masuk rekap lembar Excel', len(rek)),
        ('Total nominal dalam rekap', total_rekap),
        ('Total pada sheet Total Excel', rk.get('total_sheet_excel')),
        ('Status rekonsiliasi', 'Cocok' if abs(total_rekap - (rk.get('total_sheet_excel') or 0)) <= 1 else 'Tidak cocok'),
        ('Nilai baris bertanda besar (>= Rp 10 miliar)', flagged),
        ('Total jika baris bertanda besar dikecualikan (skenario)', total_rekap - flagged),
        ('Catatan', 'Label sektor, jenis pemberi, bentuk bantuan, dan potensi aset dibuat otomatis dari teks. '
                    'Label ini alat bantu baca, bukan penetapan status resmi. Skenario hitung tidak mengubah data sumber. '
                    'Verifikasi isi dan nilai menjadi kewenangan OPD sumber dan unit pengelola data.'),
    ]
    for k, v in info:
        ws.append([k, v])
    for r in range(1, len(info) + 1):
        ws.cell(r, 1).font = FONT_BOLD
        ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top')
        if isinstance(ws.cell(r, 2).value, float):
            ws.cell(r, 2).number_format = RUPIAH
    ws.column_dimensions['A'].width = 48
    ws.column_dimensions['B'].width = 90

    # 2. Data lengkap
    cols = [('ID', 'ID'), ('Lembar_Sheet', 'Lembar Excel'), ('Baris_Sumber', 'Baris Excel'),
            ('Tahun_Detail', 'Tahun (teks sumber)'), ('Tahun_Kegiatan', 'Tahun Kegiatan'),
            ('Masuk_Rekap_Lembar', 'Masuk Rekap Lembar'), ('Tipe_Perangkat_Daerah', 'Tingkat OPD'),
            ('Perangkat_Daerah', 'Perangkat Daerah'), ('Wilayah_Kecamatan', 'Wilayah'),
            ('Uraian_Bantuan', 'Uraian'), ('Lokasi_Penerima', 'Lokasi Penerima'), ('Bentuk_Bantuan', 'Bentuk (label otomatis)'),
            ('Kategori_Sektor', 'Sektor (label otomatis)'), ('Jumlah_Dana', 'Nilai (Rp)'),
            ('Status_Nilai', 'Status Nilai'), ('Nama_Pemberi', 'Nama Pemberi'),
            ('Tipe_Pemberi', 'Jenis Pemberi (label otomatis)'), ('Bentuk_Koordinasi', 'Bentuk Koordinasi'),
            ('Sasaran_Hasil', 'Hasil / Sasaran'), ('Potensi_Aset_KIB', 'Barang Fisik (label otomatis)'),
            ('Is_Outlier', 'Nilai Bertanda Besar'), ('Catatan_Risiko', 'Catatan Data')]
    rows = []
    for _, x in df.iterrows():
        row = [x[k] for k, _ in cols]
        row = ['Ya' if v is True else 'Tidak' if v is False else v for v in row]
        rows.append(row)
    flags = list(df['Is_Outlier'])
    inrek = list(df['Masuk_Rekap_Lembar'])
    write_table(wb.create_sheet('DATA_LENGKAP'), [h for _, h in cols], rows, money_cols=['Nilai (Rp)'],
                widths=[6, 9, 8, 14, 10, 10, 11, 28, 22, 50, 28, 12, 24, 16, 18, 30, 22, 30, 30, 10, 10, 60],
                row_fill=lambda i: FILL_FLAG if flags[i] else (FILL_OUT if not inrek[i] else None))

    # 3. Rekonsiliasi
    rows = []
    for s in rk['lembar']:
        rows.append([s['lembar'], s['subtotal_excel'], s['total_sheet'], s['nilai_dalam_rekap'],
                     s['nilai_dalam_rekap'] - (s['subtotal_excel'] or 0), s['jumlah_baris_dalam_rekap'],
                     s['jumlah_baris_luar_rekap'], s['nilai_luar_rekap'], s['aturan_rekap'],
                     'Cocok' if s['cocok'] else 'Tidak cocok'])
    write_table(wb.create_sheet('REKONSILIASI'),
                ['Lembar', 'Subtotal di Lembar Excel', 'Nilai di Sheet Total', 'Hasil Olahan (Dalam Rekap)',
                 'Selisih', 'Baris Dalam Rekap', 'Baris di Luar Rekap', 'Nilai di Luar Rekap',
                 'Aturan Rekap', 'Status'], rows,
                money_cols=['Subtotal di Lembar Excel', 'Nilai di Sheet Total', 'Hasil Olahan (Dalam Rekap)',
                            'Selisih', 'Nilai di Luar Rekap'],
                widths=[11, 20, 20, 22, 12, 12, 12, 18, 30, 12],
                total_row=['TOTAL', '=SUM(B2:B{last})', '=SUM(C2:C{last})', '=SUM(D2:D{last})', '=SUM(E2:E{last})',
                           '=SUM(F2:F{last})', '=SUM(G2:G{last})', '=SUM(H2:H{last})', '', ''])

    # 4. Rekap per OPD
    tier_order = {'Dinas': 1, 'Badan': 2, 'Kecamatan': 3, 'Kelurahan': 4}
    g = rek.groupby(['Tipe_Perangkat_Daerah', 'Perangkat_Daerah', 'Wilayah_Kecamatan'])
    opd = g.agg(keg=('ID', 'count'), nominal=('Jumlah_Dana', 'sum'),
                tanpa_nilai=('Is_Zero', 'sum'), barang=('Potensi_Aset_KIB', 'sum')).reset_index()
    opd['flag'] = [rek[(rek.Perangkat_Daerah == p) & rek.Is_Outlier]['Jumlah_Dana'].sum() for p in opd.Perangkat_Daerah]
    opd['tier'] = opd['Tipe_Perangkat_Daerah'].map(tier_order)
    opd = opd.sort_values(['tier', 'nominal'], ascending=[True, False])
    rows = [[i, x.Tipe_Perangkat_Daerah, x.Perangkat_Daerah, x.Wilayah_Kecamatan, x.keg, int(x.tanpa_nilai),
             x.nominal, x.nominal - x.flag, int(x.barang)] for i, x in enumerate(opd.itertuples(), start=1)]
    write_table(wb.create_sheet('REKAP_PER_OPD'),
                ['No', 'Tingkat', 'Perangkat Daerah', 'Wilayah', 'Jumlah Catatan', 'Catatan Tanpa Nilai',
                 'Total Nominal', 'Skenario Tanpa Nilai Bertanda Besar', 'Barang Fisik (label)'], rows,
                money_cols=['Total Nominal', 'Skenario Tanpa Nilai Bertanda Besar'],
                widths=[5, 11, 34, 26, 10, 10, 18, 22, 12],
                total_row=['', 'TOTAL', f'{len(opd)} Perangkat Daerah', '', '=SUM(E2:E{last})', '=SUM(F2:F{last})',
                           '=SUM(G2:G{last})', '=SUM(H2:H{last})', '=SUM(I2:I{last})'])

    # 5. Rekap tingkat
    rows = []
    for t in tier_order:
        d = rek[rek.Tipe_Perangkat_Daerah == t]
        rows.append([t, d.Perangkat_Daerah.nunique(), len(d), d.Jumlah_Dana.sum(),
                     d.Jumlah_Dana.sum() / total_rekap if total_rekap else 0,
                     d[~d.Is_Outlier].Jumlah_Dana.sum()])
    write_table(wb.create_sheet('REKAP_TINGKAT'),
                ['Tingkat', 'Jumlah OPD', 'Jumlah Catatan', 'Total Nominal', 'Pangsa Nominal',
                 'Skenario Tanpa Nilai Bertanda Besar'], rows,
                money_cols=['Total Nominal', 'Skenario Tanpa Nilai Bertanda Besar'], pct_cols=['Pangsa Nominal'],
                widths=[12, 11, 13, 20, 12, 24],
                total_row=['TOTAL', '=SUM(B2:B{last})', '=SUM(C2:C{last})', '=SUM(D2:D{last})', '=SUM(E2:E{last})',
                           '=SUM(F2:F{last})'])

    # 6. Rekap sektor
    rows = []
    sec = rek.groupby('Kategori_Sektor').agg(n=('ID', 'count'), v=('Jumlah_Dana', 'sum')).sort_values('v', ascending=False)
    for name, x in sec.iterrows():
        d = rek[rek.Kategori_Sektor == name]
        contoh = '; '.join(d.sort_values('Jumlah_Dana', ascending=False).Uraian_Bantuan.str[:60].drop_duplicates().head(3))
        rows.append([name, x.n, x.v, x.v / total_rekap if total_rekap else 0, d[~d.Is_Outlier].Jumlah_Dana.sum(), contoh])
    write_table(wb.create_sheet('REKAP_SEKTOR'),
                ['Sektor (label otomatis)', 'Jumlah Catatan', 'Total Nominal', 'Pangsa Nominal',
                 'Skenario Tanpa Nilai Bertanda Besar', 'Contoh Uraian Bernilai Terbesar'], rows,
                money_cols=['Total Nominal', 'Skenario Tanpa Nilai Bertanda Besar'], pct_cols=['Pangsa Nominal'],
                widths=[32, 12, 20, 12, 24, 90])

    # 7. Rekap pemberi
    pem = rek.groupby('Nama_Pemberi').agg(jenis=('Tipe_Pemberi', 'first'), n=('ID', 'count'),
                                         v=('Jumlah_Dana', 'sum'),
                                         opd=('Perangkat_Daerah', lambda s: ', '.join(sorted(set(s)))))
    pem = pem.sort_values(['v', 'n'], ascending=False).reset_index()
    rows = [[i, x.Nama_Pemberi, x.jenis, x.n, x.v, x.opd] for i, x in enumerate(pem.itertuples(), start=1)]
    write_table(wb.create_sheet('REKAP_PEMBERI'),
                ['No', 'Nama Pemberi (sesuai sumber)', 'Jenis (label otomatis)', 'Jumlah Catatan', 'Total Nominal',
                 'Perangkat Daerah'], rows, money_cols=['Total Nominal'], widths=[5, 40, 26, 10, 18, 60])

    # 8. Catatan kualitas data (dihitung dari data, tanpa kesimpulan audit)
    out = df[~df.Masuk_Rekap_Lembar]
    items = [
        ('Baris bertahun lain di lembar 2023', f"{len(out)} baris, {rp(out.Jumlah_Dana.sum())}",
         'Baris ini tidak masuk subtotal lembar 2023. Baris bernilai memiliki pasangan identik di lembar 2024-2026, '
         'sehingga tidak dihitung lagi dalam rekap.', 'Pengelola data: konfirmasi apakah baris tersebut cukup disimpan di lembar tahunnya.'),
        ('Nilai bertanda besar (>= Rp 10 miliar)',
         '; '.join(f"{x.Perangkat_Daerah}: {x.Uraian_Bantuan} {rp(x.Jumlah_Dana)}" for x in rek[rek.Is_Outlier].itertuples()),
         f"Satu baris setara {flagged / total_rekap:.1%} dari total rekap.".replace('.', ',', 1) if total_rekap else '-',
         'OPD sumber: konfirmasi nilai sesuai dokumen serah terima.'),
        ('Nilai tidak dicantumkan', f"{(rek.Status_Nilai == 'Tidak dicantumkan').sum()} catatan",
         'Nilai kosong dibedakan dari nol dan tidak menambah total.', 'OPD sumber: lengkapi bila tersedia.'),
        ('Nilai tergabung dengan baris lain', f"{(rek.Status_Nilai == 'Tergabung dengan baris lain').sum()} catatan",
         'Nilai ditulis sekali pada sel gabungan untuk beberapa pemberi.', 'Templat baru: satu baris satu nilai.'),
        ('Kolom nilai berisi teks', f"{(rek.Status_Nilai == 'Teks non-angka').sum()} catatan",
         'Contoh: "Bantuan Layak Pakai".', 'Templat baru: kolom nilai hanya angka; bentuk bantuan di kolom terpisah.'),
        ('Uraian kosong atau berisi nama OPD',
         f"{(rek.Uraian_Bantuan == '(uraian tidak tercantum)').sum() + rek.Uraian_Bantuan.str.upper().str.startswith('DINAS ').sum()} catatan",
         'Uraian tidak menjelaskan bentuk bantuan.', 'OPD sumber: lengkapi uraian.'),
        ('Format tahun tidak seragam', f"{df.Tahun_Detail.nunique()} variasi teks tahun",
         'Contoh: "Tahun 2024", "2024 (Mei)", tanggal lengkap.', 'Templat baru: kolom tanggal/tahun terstruktur.'),
        ('Nama pemberi tidak baku', f"{rek.Nama_Pemberi.nunique()} variasi nama",
         'Pemberi yang sama ditulis berbeda, misalnya singkatan dan nama lengkap.', 'Aplikasi: master pemberi dengan alias.'),
        ('Jenis pemberi perlu konfirmasi',
         f"{rek.Tipe_Pemberi.isin(['Perorangan', 'Belum Terklasifikasi']).sum()} catatan",
         'Label otomatis tidak dapat memastikan apakah pemberi badan usaha, lembaga, atau perorangan.',
         'Unit pengelola data: tetapkan jenis pemberi pada master pemberi.'),
    ]
    write_table(wb.create_sheet('CATATAN_KUALITAS_DATA'),
                ['Hal', 'Jumlah / Contoh', 'Penjelasan', 'Usulan Tindak Lanjut'], items, widths=[34, 50, 60, 55])

    wb.save(OUTPUT_FILE)
    print(f"Workbook berhasil disimpan: {OUTPUT_FILE}")


if __name__ == '__main__':
    create_clean_excel()
