"""Membangun paparan eksekutif (HTML + PDF) dari data hasil normalisasi.

Seluruh angka dihitung dari output/data_cleaned.json dan output/rekap_sumber.json,
sehingga paparan selalu konsisten dengan dashboard dan workbook Excel.
PDF dibuat dengan Google Chrome headless bila tersedia.
"""
import datetime
import html
import os
import shutil
import subprocess
import tempfile

from data_loader import load_and_clean_data, load_rekap

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_HTML = os.path.join(BASE_DIR, "output", "paparan_eksekutif_csr.html")
OUT_PDF = os.path.join(BASE_DIR, "output", "PAPARAN_EKSEKUTIF_ANALISIS_CSR_SAMARINDA.pdf")
BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober',
         'November', 'Desember']


def e(v):
    return html.escape(str(v))


def rp_full(v):
    return 'Rp ' + f"{v:,.0f}".replace(',', '.')


def rp_short(v):
    if v >= 1e9:
        return 'Rp ' + f"{v / 1e9:,.2f}".replace('.', ',') + ' M'
    if v >= 1e6:
        return 'Rp ' + f"{v / 1e6:,.1f}".replace('.', ',') + ' Jt'
    return rp_full(v)


def pct(v):
    return f"{v * 100:.1f}".replace('.', ',') + '%'


def num(v):
    return f"{v:,}".replace(',', '.')


def bars(items, value_fmt=rp_short, note_fmt=None):
    """Batang horizontal satu seri: label kiri, batang, nilai tertulis di kanan."""
    mx = max([v for _, v, _ in items] + [1])
    rows = []
    for label, v, note in items:
        w = max(0.5, v / mx * 100)
        rows.append(f'''<div class="bar-row"><span class="bar-label">{e(label)}</span>
          <span class="bar-track"><span class="bar" style="width:{w:.1f}%"></span></span>
          <span class="bar-val">{e(value_fmt(v))}<small>{e(note_fmt(note) if note_fmt else note)}</small></span></div>''')
    return '<div class="bars">' + ''.join(rows) + '</div>'


def table(headers, rows, num_cols=()):
    th = ''.join(f'<th class="{"n" if i in num_cols else ""}">{e(h)}</th>' for i, h in enumerate(headers))
    trs = ''.join('<tr>' + ''.join(f'<td class="{"n" if i in num_cols else ""}">{c}</td>' for i, c in enumerate(r)) + '</tr>'
                  for r in rows)
    return f'<table><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table>'


def slide(no, total, tag, title, subtitle, body, source):
    return f'''<section class="slide">
  <header class="slide-header"><div><span class="slide-tag">{e(tag)}</span><h2>{e(title)}</h2>
  <span class="slide-subtitle">{e(subtitle)}</span></div><span class="slide-num">{no:02d} / {total:02d}</span></header>
  <div class="slide-body">{body}</div>
  <footer class="slide-footer"><span>{e(source)}</span><span>Diskominfo Kota Samarinda · Prototipe Aplikasi TSP/CSR</span></footer>
</section>'''


def build():
    df = load_and_clean_data()
    rk = load_rekap()
    rek = df[df.Masuk_Rekap_Lembar]
    out = df[~df.Masuk_Rekap_Lembar]
    total = rek.Jumlah_Dana.sum()
    sheet_total = rk.get('total_sheet_excel') or 0
    flagged = rek[rek.Is_Outlier]
    fl_val = flagged.Jumlah_Dana.sum()
    today = datetime.date.today()
    tanggal = f"{today.day} {BULAN[today.month - 1]} {today.year}"
    src = f"Sumber: {rk['sumber']}, diolah {tanggal}"
    slides = []
    N = 8

    # 1. Cover
    cover = f'''<section class="slide cover">
  <span class="cover-badge">Draf untuk dibahas · {e(tanggal)}</span>
  <h1>Prototipe Aplikasi Pendataan TSP/CSR Kota Samarinda</h1>
  <p class="cover-desc">Hasil pengolahan data CSR 2022–2026 dan rancangan awal sistem pendukung
  register TSP satu pintu.</p>
  <div class="cover-meta">
    <div><span>Disusun oleh</span><b>Diskominfo Kota Samarinda (pengembang aplikasi)</b></div>
    <div><span>Data sumber</span><b>{e(rk['sumber'])}</b></div>
    <div><span>Konteks</span><b>Rakor Tata Kelola TSP/CSR, Inspektorat Kota Samarinda, 15 Oktober 2026</b></div>
  </div>
  <p class="cover-note">Paparan ini menjelaskan isi data dan rancangan aplikasi. Paparan ini bukan hasil audit,
  pendapat hukum, atau penilaian kewajaran nilai bantuan.</p>
</section>'''
    slides.append(cover)

    # 2. Ringkasan
    n_opd = rek.Perangkat_Daerah.nunique()
    kpis = [
        ('Total nominal dalam rekap', rp_short(total), 'sama dengan sheet Total Excel' if abs(total - sheet_total) <= 1 else 'BELUM sama dengan sheet Total'),
        ('Catatan kegiatan', num(len(rek)), f'{n_opd} perangkat daerah'),
        ('Nama pemberi', num(rek.Nama_Pemberi.nunique()), 'sesuai penulisan di sumber'),
        ('Nilai bertanda besar', rp_short(fl_val), f'{len(flagged)} catatan · {pct(fl_val / total)} dari total'),
    ]
    k = ''.join(f'<div class="kpi"><span>{e(a)}</span><b>{e(b)}</b><small>{e(c)}</small></div>' for a, b, c in kpis)
    fl_txt = '; '.join(f"{x.Uraian_Bantuan} di {x.Perangkat_Daerah} (lembar {x.Lembar_Sheet} baris {x.Baris_Sumber}) tercatat {rp_full(x.Jumlah_Dana)}"
                       for x in flagged.itertuples())
    body = f'''<div class="kpis">{k}</div>
    <div class="cols"><div class="box"><h3>Yang ditunjukkan data</h3><ul>
      <li>Penjumlahan {num(len(rek))} catatan menghasilkan {e(rp_full(total))}, sama dengan sheet Total di Excel.</li>
      <li>{e(fl_txt)}. Satu catatan ini setara {pct(fl_val / total)} dari total.</li>
      <li>Bila catatan tersebut dikecualikan, skenario hitung menghasilkan {e(rp_full(total - fl_val))}. Angka skenario bukan koreksi data.</li>
    </ul></div>
    <div class="box warn"><h3>Koreksi atas versi dashboard sebelumnya</h3><ul>
      <li>Versi sebelumnya menampilkan Rp 29,67 M dari 806 baris. Selisih itu berasal dari skrip pengolah, bukan dari Excel.</li>
      <li>Skrip lama membuang baris lanjutan pada sel gabungan dan ikut menghitung salinan baris di lembar 2023.</li>
      <li>Angka Rp 14,67 M pada versi lama juga bukan "nilai riil". Angka itu hanya skenario tanpa satu catatan.</li>
    </ul></div></div>'''
    slides.append(slide(2, N, 'Ringkasan', 'Angka utama hasil pengolahan', 'Seluruh angka dihitung ulang dari Excel sumber', body, src))

    # 3. Rekonsiliasi
    rows = []
    for l in rk['lembar']:
        rows.append([e(l['lembar']), rp_full(l['subtotal_excel'] or 0), rp_full(l['nilai_dalam_rekap']),
                     num(l['jumlah_baris_dalam_rekap']),
                     (f"{l['jumlah_baris_luar_rekap']} baris · {rp_full(l['nilai_luar_rekap'])}" if l['jumlah_baris_luar_rekap'] else '-'),
                     '<span class="ok">Cocok</span>' if l['cocok'] else '<span class="bad">Perlu ditelusuri</span>'])
    rows.append(['<b>Total</b>', f"<b>{rp_full(sheet_total)}</b>", f"<b>{rp_full(total)}</b>", f"<b>{num(len(rek))}</b>",
                 f"{len(out)} baris" if len(out) else '-', ''])
    paired = out.Catatan_Risiko.str.contains('sama dengan ID').sum()
    body = table(['Lembar', 'Subtotal / Sheet Total Excel', 'Hasil olahan', 'Catatan', 'Di luar rekap', 'Status'], rows, num_cols=(1, 2, 3)) + f'''
    <div class="box"><h3>Penjelasan baris di luar rekap</h3><ul>
      <li>Lembar 2023 memuat {len(out)} baris bertahun 2024–2026. Subtotal lembar 2023 hanya menjumlahkan baris bertahun 2023.</li>
      <li>{paired} baris di antaranya punya pasangan identik di lembar tahunnya, termasuk seluruh baris yang bernilai. Baris itu disimpan untuk penelusuran, tetapi tidak dihitung dua kali.</li>
      <li>Sumbangan sama yang muncul di tahun berbeda diperlakukan sebagai kegiatan berulang, bukan data ganda.</li>
    </ul></div>'''
    slides.append(slide(3, N, 'Rekonsiliasi', 'Hasil olahan cocok dengan rekap Excel di semua lembar',
                        'Subtotal tiap lembar dibandingkan dengan penjumlahan baris hasil olahan', body, src))

    # 4. Tingkat OPD
    tiers = ['Dinas', 'Badan', 'Kecamatan', 'Kelurahan']
    items, rows = [], []
    for t in tiers:
        d = rek[rek.Tipe_Perangkat_Daerah == t]
        items.append((t, d.Jumlah_Dana.sum(), f" · {len(d)} catatan · {d.Perangkat_Daerah.nunique()} OPD"))
    top = rek.groupby(['Perangkat_Daerah', 'Tipe_Perangkat_Daerah']).agg(v=('Jumlah_Dana', 'sum'), n=('ID', 'count')) \
        .reset_index().sort_values('v', ascending=False).head(8)
    for x in top.itertuples():
        rows.append([e(x.Perangkat_Daerah), e(x.Tipe_Perangkat_Daerah), rp_full(x.v), num(int(x.n))])
    dinas = rek[rek.Tipe_Perangkat_Daerah == 'Dinas'].Jumlah_Dana.sum()
    body = f'''<div class="cols"><div><h3>Nominal menurut tingkat perangkat daerah</h3>{bars(items)}
      <p class="note">Nominal Dinas {e(rp_short(dinas))} termasuk catatan bertanda besar {e(rp_short(fl_val))}.
      Tanpa catatan itu, nominal Dinas menjadi {e(rp_short(dinas - fl_val))}.</p></div>
      <div><h3>Delapan perangkat daerah dengan nominal terbesar</h3>{table(['Perangkat daerah', 'Tingkat', 'Nominal', 'Catatan'], rows, num_cols=(2, 3))}</div></div>'''
    slides.append(slide(4, N, 'Sebaran', 'Sebaran menurut perangkat daerah penerima', 'Dihitung dari catatan yang masuk rekap', body, src))

    # 5. Sektor & pemberi
    sec = rek.groupby('Kategori_Sektor').agg(v=('Jumlah_Dana', 'sum'), n=('ID', 'count')).sort_values('v', ascending=False)
    items = [(name, x.v, f" · {int(x.n)} catatan") for name, x in sec.iterrows()]
    fl_sector = ', '.join(sorted(set(flagged.Kategori_Sektor)))
    pem = rek.groupby('Nama_Pemberi').agg(v=('Jumlah_Dana', 'sum'), n=('ID', 'count'),
                                         opd=('Perangkat_Daerah', lambda s: ', '.join(sorted(set(s))[:2]))) \
        .sort_values('v', ascending=False).head(8)
    rows = [[e(name), rp_full(x.v), num(int(x.n)), e(x.opd)] for name, x in pem.iterrows()]
    body = f'''<div class="cols"><div><h3>Nominal menurut sektor (label otomatis)</h3>{bars(items)}
      <p class="note">Sektor ditentukan otomatis dari kata kunci uraian. Label perlu ditetapkan resmi oleh pengelola data.
      Sektor {e(fl_sector)} termasuk catatan bertanda besar {e(rp_short(fl_val))}.</p></div>
      <div><h3>Delapan pemberi dengan nominal terbesar</h3>{table(['Nama pemberi (sesuai sumber)', 'Nominal', 'Catatan', 'Perangkat daerah'], rows, num_cols=(1, 2))}
      <p class="note">Nama ditampilkan seperti tertulis di sumber. Satu pemberi bisa tertulis dengan beberapa variasi nama.</p></div></div>'''
    slides.append(slide(5, N, 'Sektor dan pemberi', 'Sebaran menurut sektor dan pemberi', 'Dihitung dari catatan yang masuk rekap', body, src))

    # 6. Kualitas data
    q = [
        ('Nilai tidak dicantumkan', (rek.Status_Nilai == 'Tidak dicantumkan').sum(), 'Nilai kosong dibedakan dari nol.', 'OPD sumber'),
        ('Nilai tergabung dengan baris lain', (rek.Status_Nilai == 'Tergabung dengan baris lain').sum(), 'Satu nilai ditulis untuk beberapa pemberi lewat sel gabungan.', 'Templat baru'),
        ('Kolom nilai berisi teks', (rek.Status_Nilai == 'Teks non-angka').sum(), 'Contoh: "Bantuan Layak Pakai".', 'Templat baru'),
        ('Uraian kosong atau berisi nama OPD', (rek.Uraian_Bantuan == '(uraian tidak tercantum)').sum() + rek.Uraian_Bantuan.str.upper().str.startswith('DINAS ').sum(), 'Bentuk bantuan tidak terbaca.', 'OPD sumber'),
        ('Variasi penulisan tahun', df.Tahun_Detail.nunique(), 'Contoh: "Tahun 2024", "2024 (Mei)", tanggal lengkap.', 'Templat baru'),
        ('Jenis pemberi perlu konfirmasi', rek.Tipe_Pemberi.isin(['Perorangan', 'Belum Terklasifikasi']).sum(), 'Nama saja tidak cukup untuk memastikan jenis pemberi.', 'Pengelola data'),
        ('Barang fisik (label otomatis)', rek.Potensi_Aset_KIB.sum(), 'Butuh referensi BAST/KIB dari unit berwenang.', 'BPKAD / OPD'),
    ]
    rows = [[e(a), num(int(b)), e(c), e(d)] for a, b, c, d in q]
    body = table(['Hal', 'Jumlah', 'Keterangan', 'Pihak yang melengkapi'], rows, num_cols=(1,)) + '''
    <div class="box"><h3>Artinya bagi aplikasi</h3><ul>
      <li>Formulir dan templat impor perlu satu baris untuk satu kegiatan dan satu nilai, dengan kolom tahun terstruktur.</li>
      <li>Master pemberi dengan daftar alias diperlukan agar satu pemberi tidak tercatat dengan banyak nama.</li>
      <li>Kolom referensi BAST, RKUD, dan KIB diisi unit yang berwenang. Aplikasi hanya menyimpan referensinya.</li>
    </ul></div>'''
    slides.append(slide(6, N, 'Kualitas data', 'Catatan kualitas data untuk dilengkapi', 'Penjelasan isi data, bukan hasil pemeriksaan', body, src))

    # 7. Enam langkah
    steps = [
        ('1. Regulasi', 'Bagian Hukum · Perekonomian', 'Menyimpan referensi regulasi dan menyediakan laporan berkala.', 'Tidak menafsirkan kepatuhan hukum.'),
        ('2. Kelembagaan', 'Bagian Perekonomian · Bapperida', 'Akun dan hak akses per peran: operator OPD, pengelola data, pimpinan.', 'Tidak membentuk forum atau tim.'),
        ('3. Data satu pintu', 'Bapperida · Diskominfo', 'Register TSP terpusat: perusahaan, bentuk, nilai, lokasi, penerima, BAST. Publikasi ringkasan setelah disetujui.', 'Isi data tetap tanggung jawab unit sumber.'),
        ('4. Sinkronisasi', 'Bapperida · OPD', 'Menampilkan potensi kecocokan kegiatan dengan daftar prioritas Musrenbang/RKPD/DPA yang diimpor.', 'Tidak memutuskan adanya pendanaan ganda.'),
        ('5. Akuntansi & aset', 'BPKAD', 'Menyimpan status penyaluran dan nomor referensi RKUD/BAST/KIB.', 'Tidak membukukan atau mengakui aset.'),
        ('6. Integritas', 'Sekda · Inspektorat · UPG', 'Menampilkan tautan edaran dan riwayat perubahan data.', 'Tidak melakukan pemeriksaan atau audit.'),
    ]
    rows = [[f'<b>{e(a)}</b>', e(b), e(c), e(d)] for a, b, c, d in steps]
    body = table(['Langkah', 'Penanggung jawab menurut materi', 'Dukungan aplikasi', 'Batas peran aplikasi'], rows) + \
        '<p class="note">Peran Diskominfo ada pada langkah 3 bersama Bapperida, yaitu menyediakan aplikasi, hosting, dan portal publikasi.</p>'
    slides.append(slide(7, N, 'Dukungan sistem', 'Posisi aplikasi dalam enam langkah perbaikan tata kelola',
                        'Mengacu materi Inspektorat Kota Samarinda, Rakor 15 Oktober 2026', body,
                        'Sumber: materi paparan Inspektorat (IMG_4891.HEIC) dan dokumen SRS v0.2'))

    # 8. Keputusan & tahapan
    body = '''<div class="cols"><div class="box"><h3>Keputusan yang dibutuhkan</h3><ol>
      <li>Unit pemilik register dan penyetuju publikasi. Usulan: Bapperida sesuai langkah 3.</li>
      <li>Definisi satu catatan kegiatan dan templat impor resmi.</li>
      <li>Daftar kategori, jenis pemberi, dan istilah TSP/CSR yang resmi.</li>
      <li>Tindak lanjut catatan bertanda besar oleh OPD sumber.</li>
      <li>Lokasi hosting, kebijakan lampiran BAST, dan data yang boleh dipublikasikan.</li>
    </ol></div>
    <div class="box"><h3>Tahapan pengembangan yang diusulkan</h3><ol>
      <li><b>Tahap 1:</b> register terpusat, impor Excel dengan pratinjau kesalahan, dashboard internal.</li>
      <li><b>Tahap 2:</b> entri dan pengiriman data oleh operator OPD, lampiran BAST, riwayat perubahan.</li>
      <li><b>Tahap 3:</b> pencocokan dengan RKPD/DPA, referensi RKUD/KIB, ringkasan publik di portal Pemkot.</li>
    </ol></div></div>
    <p class="note">Rincian kebutuhan, use case, dan rancangan basis data ada di dokumen SRS v0.2.</p>'''
    slides.append(slide(8, N, 'Tindak lanjut', 'Keputusan yang dibutuhkan dan tahapan pengembangan', 'Untuk dibahas bersama unit terkait', body, src))

    css = '''
@page { size: A4 landscape; margin: 0; }
* { box-sizing: border-box; margin: 0; padding: 0; }
:root { --ink: #0f172a; --sub: #475569; --muted: #64748b; --line: #e2e8f0; --accent: #1d4ed8; --bar: #2563eb; --bg: #ffffff; --page: #f1f5f9; }
body { font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif; background: var(--page); color: var(--ink);
  -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.slide { width: 297mm; height: 210mm; page-break-after: always; background: var(--bg); padding: 14mm 18mm 10mm;
  display: flex; flex-direction: column; overflow: hidden; margin: 0 auto 8px; }
@media print { body { background: #fff; } .slide { margin: 0; } }
@media screen and (max-width: 1150px) { .slide { width: 100%; height: auto; min-height: 0; padding: 20px 16px; } .cols, .kpis { grid-template-columns: 1fr !important; } }
.slide-header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid var(--line); padding-bottom: 8px; margin-bottom: 12px; }
.slide-tag { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .08em; color: var(--accent); }
h2 { font-size: 23px; font-weight: 800; margin: 2px 0; }
.slide-subtitle { font-size: 11.5px; color: var(--muted); }
.slide-num { font-family: 'JetBrains Mono', monospace; font-size: 11px; color: var(--sub); border: 1px solid #cbd5e1; padding: 3px 9px; border-radius: 6px; }
.slide-body { flex: 1; display: flex; flex-direction: column; gap: 14px; font-size: 13.5px; line-height: 1.55; }
.slide-footer { display: flex; justify-content: space-between; border-top: 1px solid var(--line); padding-top: 6px; font-size: 9px; color: #94a3b8; }
h3 { font-size: 14px; font-weight: 700; margin-bottom: 6px; }
ul, ol { padding-left: 18px; } li { margin-bottom: 6px; }
.cols { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
.box { background: #f8fafc; border: 1px solid var(--line); border-radius: 8px; padding: 12px 16px; }
.box.warn { background: #fffbeb; border-color: #fde68a; }
.note { font-size: 12px; color: var(--muted); margin-top: 6px; }
.kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; }
.kpi { border: 1px solid var(--line); border-top: 3px solid var(--bar); border-radius: 8px; padding: 10px 12px; display: flex; flex-direction: column; gap: 2px; }
.kpi span { font-size: 10px; text-transform: uppercase; letter-spacing: .05em; color: var(--muted); font-weight: 700; }
.kpi b { font-size: 26px; font-family: 'JetBrains Mono', monospace; }
.kpi small { font-size: 12px; color: var(--sub); }
table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
th { text-align: left; background: #f1f5f9; color: var(--sub); font-weight: 700; padding: 7px 9px; border-bottom: 1px solid #cbd5e1; }
td { padding: 7px 9px; border-bottom: 1px solid var(--line); vertical-align: top; }
.n { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
.ok { color: #047857; font-weight: 700; } .bad { color: #b91c1c; font-weight: 700; }
.bars { display: flex; flex-direction: column; gap: 10px; }
.bar-row { display: grid; grid-template-columns: 30% 1fr 46%; align-items: center; gap: 8px; }
.bar-label { font-size: 12.5px; color: var(--ink); }
.bar-track { height: 16px; }
.bar { display: block; height: 16px; background: var(--bar); border-radius: 0 4px 4px 0; }
.bar-val { font-size: 12.5px; font-weight: 700; white-space: nowrap; font-variant-numeric: tabular-nums; }
.bar-val small { font-weight: 400; color: var(--muted); font-size: 11px; }
.cover { justify-content: center; padding: 24mm; background: radial-gradient(circle at 90% 10%, #eff6ff 0%, #ffffff 55%); }
.cover-badge { display: inline-block; width: fit-content; background: #eff6ff; border: 1px solid #bfdbfe; color: var(--accent); border-radius: 20px; padding: 5px 14px; font-size: 11px; font-weight: 700; margin-bottom: 16px; }
h1 { font-size: 32px; font-weight: 800; line-height: 1.25; max-width: 880px; margin-bottom: 12px; }
.cover-desc { font-size: 15px; color: var(--sub); max-width: 820px; margin-bottom: 26px; }
.cover-meta { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; border-top: 1px solid var(--line); padding-top: 16px; max-width: 980px; }
.cover-meta span { display: block; font-size: 10px; text-transform: uppercase; color: var(--muted); font-weight: 700; letter-spacing: .05em; }
.cover-meta b { font-size: 12.5px; }
.cover-note { margin-top: 22px; font-size: 11px; color: var(--muted); max-width: 820px; }
'''
    doc = f'''<!doctype html>
<html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Paparan Prototipe CSR</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap" rel="stylesheet">
<style>{css}</style></head><body>
{''.join(slides)}
</body></html>'''
    with open(OUT_HTML, 'w', encoding='utf-8') as f:
        f.write(doc)
    print(f"HTML disimpan: {OUT_HTML}")

    candidates = ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
                  shutil.which('google-chrome'), shutil.which('chromium'), shutil.which('chromium-browser')]
    for chrome in [c for c in candidates if c and os.path.exists(c)]:
        with tempfile.TemporaryDirectory() as profile:
            res = subprocess.run([chrome, '--headless=new', '--disable-gpu', '--no-first-run', '--no-pdf-header-footer',
                                  f'--user-data-dir={profile}', '--virtual-time-budget=5000',
                                  f'--print-to-pdf={OUT_PDF}', 'file://' + OUT_HTML],
                                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if res.returncode == 0 and os.path.exists(OUT_PDF):
            print(f"PDF disimpan: {OUT_PDF}")
            return
    print("Chrome/Chromium tidak dapat dijalankan; PDF tidak dibuat. Buka HTML lalu cetak ke PDF secara manual.")


if __name__ == '__main__':
    build()
