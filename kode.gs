/**
 * ============================================================
 *  SIMAS 1 — Web App Pencatatan Pemeriksaan (Google Apps Script)
 *  File: Code.gs
 *  Fitur: Form input per unit • Dashboard monitoring • Unduh Excel / Word / PDF
 * ============================================================
 *  CARA PAKAI SINGKAT
 *  1. Buka script.google.com → Proyek baru.
 *  2. Tempel file ini ke Code.gs, buat file HTML bernama "Index" lalu tempel isi Index.html.
 *  3. Ubah CONFIG.SPREADSHEET_ID bila perlu (default: sheet SIMAS 1 milik Anda).
 *  4. Deploy → Aplikasi web → Jalankan sebagai: Saya, Akses: Siapa saja (atau siapa saja dengan akun Google).
 * ============================================================
 */

const CONFIG = {
  SPREADSHEET_ID: '1FF_83ViaRu35LFunm3JK-S4P5LrWTMn3w3QxMpjhkoM',
  SHEET_NAME: 'DATA_INPUT',            // sheet penyimpanan (dibuat otomatis)
  APP_TITLE: 'SIMAS 1',
  APP_SUBTITLE: 'Sistem Pencatatan & Monitoring Pemeriksaan',
  TIMEZONE: 'Asia/Jakarta',
  TEAM_TARGET: 12,                      // jumlah tim/petugas yang mengisi
  MAX_REPORT_RECORDS: 60                // batas data pada laporan Word/PDF
};

// ---- Daftar unit & isian (urutan = urutan kolom di spreadsheet) ----
const UNITS = [
  { key: 'u04', name: 'UNIT 04' },
  { key: 'u07', name: 'UNIT 07' },
  { key: 'u08', name: 'UNIT 08' },
  { key: 'u09', name: 'UNIT 09' },
  { key: 'ubp', name: 'UNIT BANTUAN PROV.' }
];

const ITEMS = [
  { key: 'bbmUnit',      label: 'BBM UNIT (%)',      short: 'BBM Unit',      type: 'pct',    icon: '⛽' },
  { key: 'bbmPortable',  label: 'BBM PORTABLE (%)',  short: 'BBM Portable',  type: 'pct',    icon: '🛢️' },
  { key: 'oliHisap',     label: 'OLI HISAP (%)',     short: 'Oli Hisap',     type: 'pct',    icon: '🧴' },
  { key: 'mesinPortable',label: 'MESIN PORTABLE',    short: 'Mesin Portable',type: 'status', icon: '⚙️' },
  { key: 'lampuSirine',  label: 'LAMPU SIRINE',      short: 'Lampu Sirine',  type: 'status', icon: '🚨' },
  { key: 'lampuDepan',   label: 'LAMPU DEPAN',       short: 'Lampu Depan',   type: 'status', icon: '💡' },
  { key: 'lampuRem',     label: 'LAMPU REM',         short: 'Lampu Rem',     type: 'status', icon: '🛑' },
  { key: 'lampuSen',     label: 'LAMPU SEN',         short: 'Lampu Sen',     type: 'status', icon: '↔️' },
  { key: 'lampuRotari',  label: 'LAMPU ROTARI',      short: 'Lampu Rotari',  type: 'status', icon: '🔆' },
  { key: 'airRadiator',  label: 'AIR RADIATOR',      short: 'Air Radiator',  type: 'status', icon: '💧' },
  { key: 'oliMesin',     label: 'OLI MESIN',         short: 'Oli Mesin',     type: 'status', icon: '🔧' }
];

const STATUS_OPTIONS = ['Baik', 'Kurang', 'Rusak'];

const BASE_COLS = [
  'No.',
  'Tanggal Pemeriksaan',
  'Nama Petugas',
  'Peleton (1-6)',
  'Regu (1-6; pisahkan koma jika lebih dari satu)'
];

// ============================================================
//  WEB APP ENTRY
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.APP_TITLE + ' — ' + CONFIG.APP_SUBTITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Dipanggil sekali saat halaman dibuka: konfigurasi + seluruh data. */
function bootstrap() {
  return {
    config: {
      title: CONFIG.APP_TITLE,
      subtitle: CONFIG.APP_SUBTITLE,
      teamTarget: CONFIG.TEAM_TARGET,
      units: UNITS,
      items: ITEMS,
      statusOptions: STATUS_OPTIONS
    },
    records: readRecords_()
  };
}

function getRecords() {
  return readRecords_();
}

// ============================================================
//  SHEET HELPERS
// ============================================================
function getHeaders_() {
  const h = BASE_COLS.slice();
  UNITS.forEach(function (u) {
    ITEMS.forEach(function (i) { h.push(u.name + ' - ' + i.label); });
    h.push('KETERANGAN ' + u.name);
  });
  h.push('Waktu Input');
  return h;
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_NAME);

  const headers = getHeaders_();
  const n = headers.length;
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 2, sh.getMaxRows(), 1).setNumberFormat('@');   // tanggal sebagai teks
    sh.getRange(1, 5, sh.getMaxRows(), 1).setNumberFormat('@');   // regu sebagai teks ("1, 3")
    sh.getRange(1, 1, 1, n).setValues([headers])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#0f766e')
      .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
    sh.setRowHeight(1, 60);
    sh.setFrozenRows(1);
    sh.setFrozenColumns(3);
  }
  return sh;
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  return String(v || '');
}
function fmtDateTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  return String(v || '');
}

function rowToRecord_(r) {
  const rec = {
    no: r[0],
    tanggal: fmtDate_(r[1]),
    petugas: String(r[2] || ''),
    peleton: r[3],
    regu: String(r[4] || ''),
    waktu: fmtDateTime_(r[r.length - 1]),
    units: {}
  };
  UNITS.forEach(function (u, ui) {
    const base = 5 + ui * (ITEMS.length + 1);
    const d = {};
    let filled = false;
    ITEMS.forEach(function (it, ii) {
      const v = r[base + ii];
      d[it.key] = v;
      if (v !== '' && v !== null) filled = true;
    });
    d.ket = String(r[base + ITEMS.length] || '');
    d.active = filled;
    rec.units[u.key] = d;
  });
  return rec;
}

function readRecords_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = getHeaders_().length;
  const vals = sh.getRange(2, 1, last - 1, n).getValues();
  return vals
    .filter(function (r) { return r[1] !== '' || r[2] !== ''; })
    .map(rowToRecord_);
}

// ============================================================
//  SIMPAN DATA
// ============================================================
function clean_(s) {
  s = String(s == null ? '' : s).trim().substring(0, 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // cegah formula injection
  return s;
}

function validate_(p) {
  if (!p) throw new Error('Data kosong.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(p.tanggal || ''))) throw new Error('Tanggal pemeriksaan tidak valid.');
  if (!String(p.petugas || '').trim()) throw new Error('Nama petugas wajib diisi.');
  const pel = Number(p.peleton);
  if (!(pel >= 1 && pel <= 6)) throw new Error('Peleton harus 1 sampai 6.');
  if (!Array.isArray(p.regu) || !p.regu.length) throw new Error('Pilih minimal satu regu.');
  p.regu.forEach(function (x) { if (!(Number(x) >= 1 && Number(x) <= 6)) throw new Error('Regu harus 1 sampai 6.'); });

  let any = false;
  UNITS.forEach(function (u) {
    const d = p.units && p.units[u.key];
    if (!d || !d.active) return;
    any = true;
    ITEMS.forEach(function (it) {
      const v = d[it.key];
      if (it.type === 'pct') {
        const num = Number(v);
        if (v === '' || v === null || v === undefined || isNaN(num) || num < 0 || num > 100)
          throw new Error(u.name + ' — ' + it.short + ' harus 0–100%.');
      } else if (STATUS_OPTIONS.indexOf(v) < 0) {
        throw new Error(u.name + ' — ' + it.short + ' belum dipilih.');
      }
    });
  });
  if (!any) throw new Error('Minimal satu unit harus diperiksa.');
}

function saveEntry(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    validate_(p);
    const sh = getSheet_();
    const last = sh.getLastRow();
    const no = last;   // baris header = 1, jadi nomor berikut = last
    const reguSorted = p.regu.map(Number).sort(function (a, b) { return a - b; }).join(', ');

    const row = [no, p.tanggal, clean_(p.petugas), Number(p.peleton), reguSorted];
    UNITS.forEach(function (u) {
      const d = p.units[u.key];
      const on = d && d.active;
      ITEMS.forEach(function (it) {
        row.push(on ? (it.type === 'pct' ? Number(d[it.key]) : d[it.key]) : '');
      });
      row.push(on ? clean_(d.ket) : '');
    });
    row.push(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'));

    sh.getRange(last + 1, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();
    return { ok: true, no: no };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  EKSPOR (Excel / Word / PDF)
// ============================================================
function filterRecords_(recs, f) {
  f = f || {};
  return recs.filter(function (r) {
    return (!f.from || r.tanggal >= f.from) &&
           (!f.to || r.tanggal <= f.to) &&
           (!f.petugas || r.petugas === f.petugas);
  });
}

function exportFile(format, filters) {
  const recs = filterRecords_(readRecords_(), filters)
    .sort(function (a, b) { return a.tanggal < b.tanggal ? 1 : (a.tanggal > b.tanggal ? -1 : b.no - a.no); });
  if (!recs.length) throw new Error('Tidak ada data pada filter yang dipilih.');

  const stamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmm');
  const base = 'Laporan_Pemeriksaan_' + stamp;

  if (format === 'xlsx') return exportXlsx_(recs, base);

  const html = buildReportHtml_(recs, filters || {});
  if (format === 'doc') {
    const blob = Utilities.newBlob(html, 'application/msword', base + '.doc');
    return { filename: base + '.doc', mime: 'application/msword', base64: Utilities.base64Encode(blob.getBytes()) };
  }
  if (format === 'pdf') {
    const pdf = Utilities.newBlob(html, 'text/html', base + '.html').getAs('application/pdf');
    return { filename: base + '.pdf', mime: 'application/pdf', base64: Utilities.base64Encode(pdf.getBytes()) };
  }
  throw new Error('Format tidak dikenal.');
}

function recordToRow_(r, idx) {
  const row = [idx, r.tanggal, r.petugas, r.peleton, r.regu];
  UNITS.forEach(function (u) {
    const d = r.units[u.key];
    ITEMS.forEach(function (it) { row.push(d && d.active ? d[it.key] : ''); });
    row.push(d && d.active ? d.ket : '');
  });
  row.push(r.waktu);
  return row;
}

function exportXlsx_(recs, base) {
  const headers = getHeaders_();
  const rows = recs.map(function (r, i) { return recordToRow_(r, i + 1); });

  const ss = SpreadsheetApp.create(base);
  const id = ss.getId();
  try {
    const sh = ss.getSheets()[0];
    sh.setName('Data Pemeriksaan');
    if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
    sh.getRange(1, 2, rows.length + 1, 1).setNumberFormat('@');
    sh.getRange(1, 5, rows.length + 1, 1).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#0f766e')
      .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
    const body = sh.getRange(2, 1, rows.length, headers.length);
    body.setValues(rows).setVerticalAlignment('middle');
    sh.setRowHeight(1, 60);
    sh.setFrozenRows(1);
    sh.setFrozenColumns(3);
    sh.setColumnWidths(1, headers.length, 105);
    sh.setColumnWidth(3, 170);

    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Baik').setBackground('#dcfce7').setFontColor('#166534').setRanges([body]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Kurang').setBackground('#fef3c7').setFontColor('#92400e').setRanges([body]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Rusak').setBackground('#fee2e2').setFontColor('#991b1b').setRanges([body]).build()
    ];
    sh.setConditionalFormatRules(rules);
    SpreadsheetApp.flush();

    const res = UrlFetchApp.fetch(
      'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) throw new Error('Gagal membuat file Excel (kode ' + res.getResponseCode() + ').');
    return {
      filename: base + '.xlsx',
      mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      base64: Utilities.base64Encode(res.getContent())
    };
  } finally {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
  }
}

// ---------- Laporan HTML (dipakai untuk Word & PDF) ----------
function escHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
function statusStyle_(v) {
  if (v === 'Baik')   return 'background:#dcfce7;color:#166534;font-weight:bold;';
  if (v === 'Kurang') return 'background:#fef3c7;color:#92400e;font-weight:bold;';
  if (v === 'Rusak')  return 'background:#fee2e2;color:#991b1b;font-weight:bold;';
  return '';
}
function pctStyle_(v) {
  if (v === '' || v === null) return '';
  v = Number(v);
  if (v >= 60) return 'background:#dcfce7;color:#166534;font-weight:bold;';
  if (v >= 30) return 'background:#fef3c7;color:#92400e;font-weight:bold;';
  return 'background:#fee2e2;color:#991b1b;font-weight:bold;';
}

function statsPerUnit_(recs) {
  return UNITS.map(function (u) {
    const s = { name: u.name, n: 0, sum: {}, cnt: {}, baik: 0, kurang: 0, rusak: 0 };
    recs.forEach(function (r) {
      const d = r.units[u.key];
      if (!d || !d.active) return;
      s.n++;
      ITEMS.forEach(function (it) {
        const v = d[it.key];
        if (it.type === 'pct') {
          if (v !== '') { s.sum[it.key] = (s.sum[it.key] || 0) + Number(v); s.cnt[it.key] = (s.cnt[it.key] || 0) + 1; }
        } else if (v === 'Baik') s.baik++;
        else if (v === 'Kurang') s.kurang++;
        else if (v === 'Rusak') s.rusak++;
      });
    });
    return s;
  });
}

function buildReportHtml_(recs, f) {
  const e = escHtml_;
  const shown = recs.slice(0, CONFIG.MAX_REPORT_RECORDS);
  const periode = (f.from || f.to) ? ((f.from || 'awal') + ' s.d. ' + (f.to || 'sekarang')) : 'Semua periode';
  const cetak = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd-MM-yyyy HH:mm');
  const pctItems = ITEMS.filter(function (i) { return i.type === 'pct'; });
  const statItems = ITEMS.filter(function (i) { return i.type === 'status'; });

  let h = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8"><title>Laporan Pemeriksaan</title>' +
    '<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom></w:WordDocument></xml><![endif]-->' +
    '<style>@page{size:21cm 29.7cm;margin:1.5cm}' +
    'body{font-family:Arial,Helvetica,sans-serif;font-size:10px;color:#0f172a}' +
    'h1{font-size:17px;margin:0;color:#0f766e}h2{font-size:12px;margin:14px 0 6px;color:#0b1f3a}' +
    'p{margin:2px 0;color:#475569}' +
    'table{border-collapse:collapse;width:100%}' +
    'th{background:#0f766e;color:#ffffff;border:1px solid #0b5e57;padding:4px;font-size:9px}' +
    'td{border:1px solid #cbd5e1;padding:3px 4px;text-align:center;font-size:9px}' +
    'td.l{text-align:left}.rec{margin-top:12px}.bar{background:#0b1f3a;color:#ffffff;padding:5px 8px;font-size:10px;font-weight:bold}' +
    '</style></head><body>';

  h += '<h1>LAPORAN PEMERIKSAAN KONDISI UNIT &amp; PERLENGKAPAN</h1>';
  h += '<p>' + e(CONFIG.APP_TITLE) + ' — ' + e(CONFIG.APP_SUBTITLE) + '</p>';
  h += '<p>Periode: <b>' + e(periode) + '</b>' + (f.petugas ? ' &nbsp;|&nbsp; Petugas: <b>' + e(f.petugas) + '</b>' : '') +
       ' &nbsp;|&nbsp; Jumlah pemeriksaan: <b>' + recs.length + '</b> &nbsp;|&nbsp; Dicetak: ' + e(cetak) + '</p>';

  // Ringkasan per unit
  h += '<h2>Ringkasan per Unit</h2><table><tr><th>Unit</th><th>Diperiksa</th>';
  pctItems.forEach(function (it) { h += '<th>Rata-rata ' + e(it.short) + '</th>'; });
  h += '<th>Baik</th><th>Kurang</th><th>Rusak</th></tr>';
  statsPerUnit_(recs).forEach(function (s) {
    h += '<tr><td class="l"><b>' + e(s.name) + '</b></td><td>' + s.n + 'x</td>';
    pctItems.forEach(function (it) {
      const avg = s.cnt[it.key] ? Math.round(s.sum[it.key] / s.cnt[it.key]) : '';
      h += '<td style="' + pctStyle_(avg) + '">' + (avg === '' ? '-' : avg + '%') + '</td>';
    });
    h += '<td style="' + statusStyle_('Baik') + '">' + s.baik + '</td><td style="' + statusStyle_('Kurang') + '">' + s.kurang +
         '</td><td style="' + statusStyle_('Rusak') + '">' + s.rusak + '</td></tr>';
  });
  h += '</table>';

  // Detail per pemeriksaan
  h += '<h2>Rincian Pemeriksaan</h2>';
  shown.forEach(function (r, idx) {
    h += '<div class="rec"><div class="bar">#' + (idx + 1) + ' &nbsp; ' + e(r.tanggal) + ' &nbsp;|&nbsp; Petugas: ' + e(r.petugas) +
         ' &nbsp;|&nbsp; Peleton ' + e(r.peleton) + ' &nbsp;|&nbsp; Regu ' + e(r.regu) + '</div>';
    h += '<table><tr><th style="width:22%">Item</th>';
    UNITS.forEach(function (u) { h += '<th>' + e(u.name) + '</th>'; });
    h += '</tr>';
    ITEMS.forEach(function (it) {
      h += '<tr><td class="l"><b>' + e(it.short) + '</b></td>';
      UNITS.forEach(function (u) {
        const d = r.units[u.key];
        if (!d || !d.active) { h += '<td style="color:#94a3b8">-</td>'; return; }
        const v = d[it.key];
        if (it.type === 'pct') h += '<td style="' + pctStyle_(v) + '">' + e(v) + '%</td>';
        else h += '<td style="' + statusStyle_(v) + '">' + e(v) + '</td>';
      });
      h += '</tr>';
    });
    h += '<tr><td class="l"><b>Keterangan</b></td>';
    UNITS.forEach(function (u) {
      const d = r.units[u.key];
      h += '<td class="l">' + ((d && d.active && d.ket) ? e(d.ket) : '-') + '</td>';
    });
    h += '</tr></table></div>';
  });

  if (recs.length > shown.length) {
    h += '<p style="margin-top:10px"><i>Catatan: laporan menampilkan ' + shown.length + ' pemeriksaan terbaru dari ' + recs.length +
         ' data. Gunakan unduhan Excel untuk seluruh data.</i></p>';
  }
  h += '</body></html>';
  return h;
}
