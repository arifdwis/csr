import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from data_loader import load_and_clean_data

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_FILE = os.path.join(BASE_DIR, "output", "DATA_CSR_SAMARINDA_BERSIH_TERNORMALISASI.xlsx")

def create_clean_excel():
    print("Memuat dan menormalisasi data sumber...")
    df = load_and_clean_data()
    
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    # Styles
    font_header = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    fill_header_dark = PatternFill(start_color='0F172A', end_color='0F172A', fill_type='solid')
    fill_header_blue = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
    fill_header_emerald = PatternFill(start_color='065F46', end_color='065F46', fill_type='solid')
    
    font_data = Font(name='Segoe UI', size=10)
    font_bold = Font(name='Segoe UI', size=10, bold=True)
    font_mono = Font(name='Consolas', size=10)
    
    thin_border_side = Side(border_style='thin', color='CBD5E1')
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    fill_zebra = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
    fill_highlight_asset = PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid') # yellow alert for asset
    fill_highlight_outlier = PatternFill(start_color='FEE2E2', end_color='FEE2E2', fill_type='solid') # red alert for 15M
    
    # Currency format for IDR
    currency_fmt = '#,##0'

    # =========================================================================
    # SHEET 1: DATA LENGKAP TERNORMALISASI (806 BARIS SUMBER)
    # =========================================================================
    print("Menyusun Sheet 1: DATA_LENGKAP_TERNORMALISASI...")
    ws1 = wb.create_sheet(title='DATA_LENGKAP_TERNORMALISASI')
    ws1.views.sheetView[0].showGridLines = True
    
    headers1 = [
        'No', 'ID Kegiatan', 'Tahun Pelaporan', 'Tingkat Entitas OPD', 'Perangkat Daerah (OPD)', 
        'Wilayah Kecamatan', 'Uraian Kegiatan / Bantuan CSR', 'Bentuk Bantuan (Uang / Barang)', 'Kategori Sektor Program', 
        'Nilai Dana (Rp)', 'Status Valuasi', 'Nama Pemberi CSR', 'Kategori Mitra Pemberi', 
        'Bentuk Koordinasi Pelaksanaan', 'Sasaran Penerima Manfaat', 'Status Barang Modal (KIB)', 
        'Koreksi Anomali 15M', 'Catatan Kepatuhan Tata Kelola (Inspektorat)'
    ]
    
    ws1.append(headers1)
    
    # Style Header
    for col_idx in range(1, len(headers1) + 1):
        cell = ws1.cell(1, col_idx)
        cell.font = font_header
        cell.fill = fill_header_dark
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = cell_border
    ws1.row_dimensions[1].height = 28
    
    row_num = 2
    for idx, row in df.iterrows():
        is_outlier_str = 'Ya (Perahu Karet 15 Miliar - Disparitas Valuasi)' if row['Is_Outlier'] else 'Normal'
        is_asset_str = 'Ya (Wajib Catat KIB / Neraca LKPD)' if row['Potensi_Aset_KIB'] else 'Bukan Barang Modal'
        
        row_vals = [
            idx + 1,
            row['ID'],
            row['Tahun_Pelaporan'],
            row['Tipe_Perangkat_Daerah'],
            row['Perangkat_Daerah'],
            row['Wilayah_Kecamatan'],
            row['Uraian_Bantuan'],
            row['Bentuk_Bantuan'],
            row['Kategori_Sektor'],
            row['Jumlah_Dana'],
            row['Status_Nilai'],
            row['Nama_Pemberi'],
            row['Tipe_Pemberi'],
            row['Bentuk_Koordinasi'],
            row['Sasaran_Hasil'],
            is_asset_str,
            is_outlier_str,
            row['Catatan_Risiko']
        ]
        ws1.append(row_vals)
        
        # Style Data Row
        for col_idx in range(1, len(headers1) + 1):
            c = ws1.cell(row_num, col_idx)
            c.font = font_data
            c.border = cell_border
            
            # Alignments
            if col_idx in [1, 2, 3, 4, 10]:
                c.alignment = Alignment(horizontal='center', vertical='center')
            elif col_idx == 9:
                c.alignment = Alignment(horizontal='right', vertical='center')
                c.number_format = currency_fmt
                c.font = font_bold
            else:
                c.alignment = Alignment(horizontal='left', vertical='center')
                
            # Highlights
            if row['Is_Outlier'] and col_idx in [7, 9, 16]:
                c.fill = fill_highlight_outlier
            elif row['Potensi_Aset_KIB'] and col_idx in [7, 15]:
                c.fill = fill_highlight_asset
            elif row_num % 2 == 1:
                if not (row['Is_Outlier'] or row['Potensi_Aset_KIB']):
                    c.fill = fill_zebra
                    
        ws1.row_dimensions[row_num].height = 20
        row_num += 1

    ws1.freeze_panes = 'C2'
    ws1.auto_filter.ref = f"A1:{get_column_letter(len(headers1))}{len(df) + 1}"

    # Auto-fit columns
    col_widths = {
        'A': 6, 'B': 12, 'C': 15, 'D': 18, 'E': 28, 'F': 24, 'G': 45, 
        'H': 28, 'I': 18, 'J': 24, 'K': 32, 'L': 28, 'M': 35, 'N': 35, 
        'O': 26, 'P': 24, 'Q': 40
    }
    for col_letter, width in col_widths.items():
        ws1.column_dimensions[col_letter].width = width

    # =========================================================================
    # SHEET 2: REKAPITULASI PER OPD (52 ENTITAS)
    # =========================================================================
    print("Menyusun Sheet 2: REKAP_PER_OPD...")
    ws2 = wb.create_sheet(title='REKAP_PER_OPD')
    ws2.views.sheetView[0].showGridLines = True
    
    headers2 = [
        'No', 'Tingkat Entitas', 'Perangkat Daerah (OPD)', 'Wilayah Kecamatan Induk', 
        'Jumlah Kegiatan', 'Total Dana Riil (Koreksi Anomali)', 'Total Dana Tercatat (Gross)', 
        'Kegiatan Aset Fisik KIB', 'Keterangan Analisis'
    ]
    ws2.append(headers2)
    
    for col_idx in range(1, len(headers2) + 1):
        cell = ws2.cell(1, col_idx)
        cell.font = font_header
        cell.fill = fill_header_blue
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = cell_border
    ws2.row_dimensions[1].height = 28
    
    # Aggregation per OPD
    opd_groups = {}
    for _, row in df.iterrows():
        pd_name = row['Perangkat_Daerah']
        if pd_name not in opd_groups:
            opd_groups[pd_name] = {
                'tipe': row['Tipe_Perangkat_Daerah'],
                'wilayah': row['Wilayah_Kecamatan'],
                'count': 0,
                'dana_gross': 0,
                'dana_net': 0,
                'aset_count': 0,
                'outlier_count': 0
            }
        opd_groups[pd_name]['count'] += 1
        opd_groups[pd_name]['dana_gross'] += row['Jumlah_Dana']
        if not row['Is_Outlier']:
            opd_groups[pd_name]['dana_net'] += row['Jumlah_Dana']
        else:
            opd_groups[pd_name]['outlier_count'] += 1
        if row['Potensi_Aset_KIB']:
            opd_groups[pd_name]['aset_count'] += 1

    # Sort: Dinas first, then Badan, Kecamatan, Kelurahan, then by dana_net desc
    tier_order = {'Dinas': 1, 'Badan': 2, 'Kecamatan': 3, 'Kelurahan': 4}
    sorted_opds = sorted(opd_groups.items(), key=lambda x: (tier_order.get(x[1]['tipe'], 99), -x[1]['dana_net']))

    r2 = 2
    for idx, (opd_name, stats) in enumerate(sorted_opds, start=1):
        ket = []
        if stats['outlier_count'] > 0:
            ket.append(f"{stats['outlier_count']} baris anomali Rp 15 Miliar")
        if stats['aset_count'] > 0:
            ket.append(f"{stats['aset_count']} kegiatan aset modal fisik")
        ket_str = '; '.join(ket) if ket else 'Penyaluran reguler'
        
        row_vals = [
            idx,
            stats['tipe'],
            opd_name,
            stats['wilayah'],
            stats['count'],
            stats['dana_net'],
            stats['dana_gross'],
            stats['aset_count'],
            ket_str
        ]
        ws2.append(row_vals)
        
        for col_idx in range(1, len(headers2) + 1):
            c = ws2.cell(r2, col_idx)
            c.font = font_data
            c.border = cell_border
            if col_idx in [1, 2]:
                c.alignment = Alignment(horizontal='center', vertical='center')
            elif col_idx in [5, 8]:
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.number_format = '#,##0'
            elif col_idx in [6, 7]:
                c.alignment = Alignment(horizontal='right', vertical='center')
                c.number_format = currency_fmt
                c.font = font_bold
            else:
                c.alignment = Alignment(horizontal='left', vertical='center')
        ws2.row_dimensions[r2].height = 20
        r2 += 1

    # Total Row
    total_row = [
        '', 'TOTAL', '52 Perangkat Daerah', 'Kota Samarinda', 
        f'=SUM(E2:E{r2-1})', f'=SUM(F2:F{r2-1})', f'=SUM(G2:G{r2-1})', f'=SUM(H2:H{r2-1})', 
        'Konsolidasi Keseluruhan'
    ]
    ws2.append(total_row)
    for col_idx in range(1, len(headers2) + 1):
        c = ws2.cell(r2, col_idx)
        c.font = Font(name='Segoe UI', size=11, bold=True)
        c.border = cell_border
        c.fill = PatternFill(start_color='E2E8F0', end_color='E2E8F0', fill_type='solid')
        if col_idx in [5, 8]:
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.number_format = '#,##0'
        elif col_idx in [6, 7]:
            c.alignment = Alignment(horizontal='right', vertical='center')
            c.number_format = currency_fmt
        else:
            c.alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[r2].height = 22

    ws2.freeze_panes = 'D2'
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(headers2))}{r2-1}"
    
    col_widths2 = {'A': 6, 'B': 18, 'C': 30, 'D': 24, 'E': 16, 'F': 24, 'G': 24, 'H': 20, 'I': 35}
    for col_letter, width in col_widths2.items():
        ws2.column_dimensions[col_letter].width = width

    # =========================================================================
    # SHEET 3: REKAP TINGKAT ENTITAS (DINAS, BADAN, KECAMATAN, KELURAHAN)
    # =========================================================================
    print("Menyusun Sheet 3: REKAP_TINGKAT_ENTITAS...")
    ws3 = wb.create_sheet(title='REKAP_TINGKAT_ENTITAS')
    ws3.views.sheetView[0].showGridLines = True
    
    headers3 = [
        'No', 'Tingkat Entitas Pemerintahan', 'Jumlah OPD Terdata', 'Jumlah Kegiatan CSR', 
        'Total Dana Riil (Koreksi)', 'Pangsa Dana Riil (%)', 'Total Dana Gross (Termasuk 15M)'
    ]
    ws3.append(headers3)
    for col_idx in range(1, len(headers3) + 1):
        cell = ws3.cell(1, col_idx)
        cell.font = font_header
        cell.fill = fill_header_emerald
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = cell_border
    ws3.row_dimensions[1].height = 28
    
    tier_stats = {}
    for t in ['Dinas', 'Badan', 'Kecamatan', 'Kelurahan']:
        tier_df = df[df['Tipe_Perangkat_Daerah'] == t]
        tier_stats[t] = {
            'opd_count': tier_df['Perangkat_Daerah'].nunique(),
            'keg_count': len(tier_df),
            'dana_net': tier_df[~tier_df['Is_Outlier']]['Jumlah_Dana'].sum(),
            'dana_gross': tier_df['Jumlah_Dana'].sum()
        }
    
    total_dana_all_net = sum(s['dana_net'] for s in tier_stats.values())
    
    r3 = 2
    for idx, (t, s) in enumerate(tier_stats.items(), start=1):
        share = (s['dana_net'] / total_dana_all_net) if total_dana_all_net > 0 else 0
        row_vals = [
            idx, t, s['opd_count'], s['keg_count'], s['dana_net'], share, s['dana_gross']
        ]
        ws3.append(row_vals)
        for col_idx in range(1, len(headers3) + 1):
            c = ws3.cell(r3, col_idx)
            c.font = font_data
            c.border = cell_border
            if col_idx in [1, 2]:
                c.alignment = Alignment(horizontal='center', vertical='center')
            elif col_idx in [3, 4]:
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.number_format = '#,##0'
            elif col_idx == 6:
                c.alignment = Alignment(horizontal='right', vertical='center')
                c.number_format = '0.00%'
                c.font = font_bold
            elif col_idx in [5, 7]:
                c.alignment = Alignment(horizontal='right', vertical='center')
                c.number_format = currency_fmt
                c.font = font_bold
        ws3.row_dimensions[r3].height = 20
        r3 += 1

    # Total Row Sheet 3
    ws3.append([
        '', 'TOTAL', f'=SUM(C2:C{r3-1})', f'=SUM(D2:D{r3-1})', 
        f'=SUM(E2:E{r3-1})', 1.0, f'=SUM(G2:G{r3-1})'
    ])
    for col_idx in range(1, len(headers3) + 1):
        c = ws3.cell(r3, col_idx)
        c.font = Font(name='Segoe UI', size=11, bold=True)
        c.border = cell_border
        c.fill = PatternFill(start_color='E2E8F0', end_color='E2E8F0', fill_type='solid')
        if col_idx in [3, 4]:
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.number_format = '#,##0'
        elif col_idx == 6:
            c.alignment = Alignment(horizontal='right', vertical='center')
            c.number_format = '0.00%'
        elif col_idx in [5, 7]:
            c.alignment = Alignment(horizontal='right', vertical='center')
            c.number_format = currency_fmt
        else:
            c.alignment = Alignment(horizontal='center', vertical='center')
    ws3.row_dimensions[r3].height = 22

    col_widths3 = {'A': 6, 'B': 25, 'C': 20, 'D': 20, 'E': 25, 'F': 20, 'G': 25}
    for col_letter, width in col_widths3.items():
        ws3.column_dimensions[col_letter].width = width

    # =========================================================================
    # SHEET 4: REKAP KATEGORI SEKTOR CSR
    # =========================================================================
    print("Menyusun Sheet 4: REKAP_SEKTOR_CSR...")
    ws4 = wb.create_sheet(title='REKAP_SEKTOR_CSR')
    ws4.views.sheetView[0].showGridLines = True
    
    headers4 = ['No', 'Kategori Sektor Program', 'Jumlah Kegiatan', 'Total Nilai Riil (Rp)', 'Pangsa (%)', 'Fokus Utama Kegiatan']
    ws4.append(headers4)
    for col_idx in range(1, len(headers4) + 1):
        cell = ws4.cell(1, col_idx)
        cell.font = font_header
        cell.fill = fill_header_dark
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = cell_border
    ws4.row_dimensions[1].height = 28
    
    sector_summary = {}
    for _, row in df.iterrows():
        s = row['Kategori_Sektor']
        if s not in sector_summary:
            sector_summary[s] = {'count': 0, 'dana_net': 0}
        sector_summary[s]['count'] += 1
        if not row['Is_Outlier']:
            sector_summary[s]['dana_net'] += row['Jumlah_Dana']
            
    sorted_sectors = sorted(sector_summary.items(), key=lambda x: -x[1]['dana_net'])
    
    fokus_map = {
        'Infrastruktur & Lingkungan': 'Excavator DPUPR, sumur bor, tandon air, kanopi, rehab posyandu/kantor',
        'Kesehatan & Penanganan Stunting': 'Program OTA GENTING DPPKB, PMT Balita stunting, sarana posyandu',
        'Pendidikan & Pengembangan SDM': 'Bantuan sarana lab komputer SMP/SD (PLN ICON+), bimtek, beasiswa',
        'Keagamaan & Sarana Ibadah': 'Bantuan hewan kurban Idul Adha (Sapi), renovasi musholla/masjid, MTQ',
        'Kebencanaan & Kedaruratan': 'Bantuan logistik sembako bencana BPBD, rescue kit (dikurangi outlier)',
        'Sosial, Ekonomi & Pemberdayaan': 'Pelatihan UMKM, sembako duafa, peralatan usaha',
        'Event & Kegiatan Masyarakat': 'Sponsorship HUT Samarinda, jalan santai, peringatan hari besar'
    }
    
    r4 = 2
    for idx, (sec_name, s) in enumerate(sorted_sectors, start=1):
        share = (s['dana_net'] / total_dana_all_net) if total_dana_all_net > 0 else 0
        ws4.append([
            idx, sec_name, s['count'], s['dana_net'], share, fokus_map.get(sec_name, '-')
        ])
        for col_idx in range(1, len(headers4) + 1):
            c = ws4.cell(r4, col_idx)
            c.font = font_data
            c.border = cell_border
            if col_idx in [1]:
                c.alignment = Alignment(horizontal='center', vertical='center')
            elif col_idx == 3:
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.number_format = '#,##0'
            elif col_idx == 4:
                c.alignment = Alignment(horizontal='right', vertical='center')
                c.number_format = currency_fmt
                c.font = font_bold
            elif col_idx == 5:
                c.alignment = Alignment(horizontal='right', vertical='center')
                c.number_format = '0.00%'
            else:
                c.alignment = Alignment(horizontal='left', vertical='center')
        ws4.row_dimensions[r4].height = 20
        r4 += 1
        
    col_widths4 = {'A': 6, 'B': 32, 'C': 18, 'D': 24, 'E': 16, 'F': 55}
    for col_letter, width in col_widths4.items():
        ws4.column_dimensions[col_letter].width = width

    # =========================================================================
    # SHEET 5: AUDIT MUTU DATA & REKOMENDASI (INSPEKTORAT)
    # =========================================================================
    print("Menyusun Sheet 5: AUDIT_MUTU_INSPEKTORAT...")
    ws5 = wb.create_sheet(title='AUDIT_MUTU_INSPEKTORAT')
    ws5.views.sheetView[0].showGridLines = True
    
    ws5.column_dimensions['A'].width = 6
    ws5.column_dimensions['B'].width = 28
    ws5.column_dimensions['C'].width = 40
    ws5.column_dimensions['D'].width = 45
    ws5.column_dimensions['E'].width = 25
    
    ws5.append(['No', 'Kategori Temuan Audit', 'Uraian Kondisi Data', 'Dampak & Rekomendasi Tata Kelola', 'Regulasi / Acuan'])
    for col_idx in range(1, 6):
        cell = ws5.cell(1, col_idx)
        cell.font = font_header
        cell.fill = fill_header_dark
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = cell_border
    ws5.row_dimensions[1].height = 28
    
    audit_points = [
        (
            1,
            'Anomali Nilai 15 Miliar',
            'Pengadaan Perahu Karet (Rubber) di Disdamkar dari Bank Kaltimtara tercatat Rp 15.000.000.000,00 (15 Miliar).',
            'Mendistorsi 51% seluruh total dana CSR Kota Samarinda. Rekomendasi: Verifikasi fisik BAST dan Berita Acara Penerimaan di Disdamkar untuk koreksi nilai wajar (kemungkinan Rp 15 Juta atau Rp 150 Juta).',
            'PP 12/2019 & Standar Akuntansi Pemerintahan'
        ),
        (
            2,
            'Barang Modal Fisik Wajib KIB (44 Kegiatan)',
            'Bantuan fisik bernilai ekonomis tinggi (Excavator DPUPR, Sumur Bor Lempake, Kanopi Kantor Tanah Merah, Pagar SDN, Genset/Drone Distapangtani) belum tertelusur masuk KIB.',
            'Aset rawan hilang, tidak terpelihara melalui APBD, dan neraca LKPD Kota Samarinda berpotensi mendapat catatan BPK RI. Rekomendasi: Sahkan penetapan hibah barang dan serahkan ke BPKAD untuk dicatat ke KIB Aset Tetap.',
            'PP 27/2014 jo. PP 28/2020 & Perwali Samarinda 2/2018'
        ),
        (
            3,
            'Sumbangan Perorangan / Zakat (196 Entri)',
            'Di BPBD, 196 entri logistik kedaruratan (beras, mie instan, sarden) berasal dari perorangan warga atau lembaga amil zakat (Baznas, Rumah Zakat).',
            'Sumbangan kebencanaan warga dan zakat masyarakat secara hukum bukan merupakan Tanggung Jawab Sosial dan Lingkungan Perseroan (CSR). Rekomendasi: Pisahkan pembukuan antara Registri TSP Korporasi dengan Bantuan Hibah Kebencanaan Masyarakat.',
            'UU 40/2007 (UUPT) & Perda Kota Samarinda 3/2013'
        ),
        (
            4,
            'Sponsorship HUT & Gratifikasi Lokal',
            'Penerimaan sumbangan tunai/barang bernilai kecil (Rp 300.000 – Rp 1.000.000) dari warung/kafe/pelaku usaha lokal untuk hadiah jalan santai / HUT Samarinda.',
            'Risiko benturan kepentingan dan bersinggungan dengan larangan gratifikasi perizinan sesuai arahan KPK RI. Rekomendasi: Terbitkan Surat Edaran Wali Kota tentang Larangan Permohonan Sumbangan HUT ke Pelaku Usaha oleh Kelurahan/Kecamatan.',
            'Langkah 6 SE Integritas & UU 20/2001 (Tipikor)'
        )
    ]
    
    for r_idx, pt in enumerate(audit_points, start=2):
        ws5.append(list(pt))
        for col_idx in range(1, 6):
            c = ws5.cell(r_idx, col_idx)
            c.font = font_data
            c.border = cell_border
            if col_idx == 1:
                c.alignment = Alignment(horizontal='center', vertical='top')
            elif col_idx == 2:
                c.alignment = Alignment(horizontal='left', vertical='top')
                c.font = font_bold
            else:
                c.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        ws5.row_dimensions[r_idx].height = 45

    wb.save(OUTPUT_FILE)
    print(f"Workbook berhasil disimpan: {OUTPUT_FILE}")

if __name__ == '__main__':
    create_clean_excel()
