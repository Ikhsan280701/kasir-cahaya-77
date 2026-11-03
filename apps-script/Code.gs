/**
 * Kios Cahaya 77 — Dashboard harga + Kasir & nota, berjalan di Google Apps Script.
 *
 * Pasang (sekali saja):
 *  1. Buka Google Sheet "Daftar Harga Kios Cahaya 77" > menu Ekstensi > Apps Script.
 *  2. Code.gs: hapus isinya, tempel seluruh file ini, lalu ganti PIN di bawah.
 *  3. Tombol + di samping "File" > HTML > beri nama  Index  (huruf I besar), hapus isinya,
 *     tempel seluruh isi file Index.html.
 *  4. Simpan, lalu Terapkan > Deployment baru > jenis "Aplikasi web"
 *     (Jalankan sebagai: Saya, Yang memiliki akses: Siapa saja) > Terapkan > izinkan akses.
 *  5. Buka alamat yang berakhiran /exec. Itu dashboard + kasirnya. Simpan di layar utama HP.
 *  Setelah mengubah kode: Terapkan > Kelola deployment > ikon pensil > Versi: Versi baru > Terapkan
 *  (alamat /exec tetap sama).
 *
 * Isi Sheet:
 *  - Tab "Harga": Nama barang | Harga (Rp) | Satuan | Kategori | Diubah (kolom E diisi otomatis)
 *  - Tab "Kios" : Nama kios, Alamat, Telepon / WA, Ucapan di bawah nota (kolom A label, kolom B isi)
 *
 * Tetap cocok dengan kasir GitHub Pages: alamat /exec?api=1 (atau ?t=...) mengembalikan JSON,
 * dan POST {action:'add', pin, n, h, s, k} menambah/memperbarui barang seperti sebelumnya.
 */

const PIN = 'GANTI-PIN-INI'; // ganti dengan PIN Anda sendiri, mis. 4-6 angka. Jangan bagikan.
const TAB_HARGA = 'Harga';
const TAB_KIOS = 'Kios';

/* ---------------- halaman web ---------------- */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.api || p.t) return json_(Object.assign({ ok: true }, readAll_()));
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Kios Cahaya 77')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* Untuk kasir GitHub Pages (versi lama). */
function doPost(e) {
  let d;
  try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'data' }); }
  if (d.action !== 'add') return json_({ ok: false, error: 'aksi' });
  try {
    cekPin_(d.pin);
    const r = simpan_({ n: d.n, h: d.h, s: d.s, k: d.k });
    return json_(Object.assign({ ok: true, status: r.status, nama: r.nama }, readAll_()));
  } catch (err) {
    return json_({ ok: false, error: /PIN/.test(err.message) ? 'pin' : 'isi' });
  }
}

/* ---------------- dipanggil dari halaman (google.script.run) ---------------- */

function ambilData() {
  return readAll_();
}

/** item = { n: nama, h: harga, s: satuan, k: kategori, lama: nama lama kalau sedang mengubah } */
function simpanBarang(item, pin) {
  cekPin_(pin);
  const r = simpan_(item || {});
  return Object.assign({ status: r.status, nama: r.nama }, readAll_());
}

function hapusBarang(nama, pin) {
  cekPin_(pin);
  const key = norm_(nama);
  if (!key) throw new Error('Nama barang kosong.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(TAB_HARGA);
    const row = cariBaris_(sh, key);
    if (!row) throw new Error('Barang "' + nama + '" tidak ditemukan. Muat ulang data.');
    sh.deleteRow(row);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return Object.assign({ status: 'dihapus', nama: String(nama).trim() }, readAll_());
}

/* ---------------- isi ---------------- */

function simpan_(item) {
  const nama = String(item.n || '').trim().replace(/\s+/g, ' ');
  const harga = Math.round(Number(item.h));
  const satuan = String(item.s || '').trim() || 'pcs';
  const kategori = String(item.k || '').trim();
  const lama = String(item.lama || '').trim();
  if (!nama) throw new Error('Isi nama barang.');
  if (!(harga > 0)) throw new Error('Isi harga lebih dari 0.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(TAB_HARGA);
    if (!String(sh.getRange(1, 5).getValue()).trim()) sh.getRange(1, 5).setValue('Diubah');

    const rowSama = cariBaris_(sh, norm_(nama));
    let row = lama ? cariBaris_(sh, norm_(lama)) : rowSama;
    if (lama && !row) throw new Error('Barang "' + lama + '" tidak ditemukan. Muat ulang data.');
    if (lama && rowSama && rowSama !== row) throw new Error('Nama "' + nama + '" sudah ada di daftar.');

    let status;
    if (row) {
      // Dari formulir ubah, kategori kosong berarti dikosongkan; dari kasir lama, kategori kosong = biarkan.
      const kat = lama || kategori ? kategori : String(sh.getRange(row, 4).getValue() || '');
      sh.getRange(row, 1, 1, 5).setValues([[nama, harga, satuan, kat, new Date()]]);
      status = 'diperbarui';
    } else {
      row = Math.max(sh.getLastRow(), 1) + 1;
      sh.getRange(row, 1, 1, 5).setValues([[nama, harga, satuan, kategori, new Date()]]);
      status = 'ditambah';
    }
    sh.getRange(row, 2).setNumberFormat('#,##0');
    sh.getRange(row, 5).setNumberFormat('dd/mm/yyyy hh:mm');
    SpreadsheetApp.flush();
    return { status: status, nama: nama };
  } finally {
    lock.releaseLock();
  }
}

function readAll_() {
  const ss = SpreadsheetApp.getActive();
  const rows = sheet_(TAB_HARGA).getDataRange().getValues().slice(1);
  const prices = rows
    .map(function (r) {
      const h = typeof r[1] === 'number' ? r[1] : Number(String(r[1] || '').replace(/[^\d]/g, ''));
      return {
        n: String(r[0] || '').trim(),
        h: Math.round(h) || 0,
        s: String(r[2] || '').trim() || 'pcs',
        k: String(r[3] || '').trim(),
        u: r[4] instanceof Date ? r[4].toISOString() : ''
      };
    })
    .filter(function (p) { return p.n && p.h > 0; });

  const kios = {};
  const ks = ss.getSheetByName(TAB_KIOS);
  if (ks) {
    ks.getRange('A1:B20').getDisplayValues().slice(1).forEach(function (r) {
      if (r[0]) kios[String(r[0]).trim().toLowerCase()] = String(r[1] || '').trim();
    });
  }
  return { prices: prices, kios: kios, sheetUrl: ss.getUrl(), at: new Date().toISOString() };
}

/* ---------------- bantuan ---------------- */

function cekPin_(pin) {
  if (PIN === 'GANTI-PIN-INI') throw new Error('PIN belum diatur. Ubah konstanta PIN di Code.gs lalu terapkan versi baru.');
  if (String(pin || '').trim() !== PIN) throw new Error('PIN salah.');
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Tab "' + name + '" tidak ada di Google Sheet.');
  return sh;
}

function cariBaris_(sh, key) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const names = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < names.length; i++) {
    if (norm_(names[i][0]) === key) return i + 2;
  }
  return 0;
}

function norm_(s) {
  return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
