/**
 * Skrip penghubung Kasir Cahaya 77 <-> Google Sheet "Daftar Harga Kios Cahaya 77".
 *
 * Pasang: di Google Sheet, menu Ekstensi > Apps Script, hapus isi Code.gs, tempel seluruh file ini,
 * ganti PIN di bawah, simpan, lalu Terapkan > Deployment baru > Aplikasi web
 * (Jalankan sebagai: Saya, Yang memiliki akses: Siapa saja).
 *
 * - GET  : mengembalikan daftar harga (tab Harga) dan data kios (tab Kios) sebagai JSON.
 *          Dengan ini Sheet tidak perlu dibuka untuk umum.
 * - POST : menambah barang ke tab Harga, atau memperbarui harganya kalau nama barang sudah ada.
 *          Wajib menyertakan PIN yang sama dengan di bawah.
 */

const PIN = 'GANTI-PIN-INI'; // ganti dengan PIN Anda sendiri, mis. 4-6 angka. Jangan bagikan.
const TAB_HARGA = 'Harga';
const TAB_KIOS = 'Kios';

function doGet() {
  return json_(Object.assign({ ok: true }, readAll_()));
}

function doPost(e) {
  let d;
  try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'data' }); }

  if (PIN === 'GANTI-PIN-INI' || String(d.pin || '') !== PIN) return json_({ ok: false, error: 'pin' });

  if (d.action === 'add') {
    const nama = String(d.n || '').trim();
    const harga = Math.round(Number(d.h));
    const satuan = String(d.s || '').trim() || 'pcs';
    const kategori = String(d.k || '').trim();
    if (!nama || !(harga > 0)) return json_({ ok: false, error: 'isi' });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = SpreadsheetApp.getActive().getSheetByName(TAB_HARGA);
      const last = sh.getLastRow();
      const names = last > 1 ? sh.getRange(2, 1, last - 1, 1).getDisplayValues() : [];
      const key = nama.toLowerCase();
      let row = 0;
      for (let i = 0; i < names.length; i++) {
        if (String(names[i][0]).trim().toLowerCase() === key) { row = i + 2; break; }
      }
      let status;
      if (row) {
        sh.getRange(row, 2, 1, 2).setValues([[harga, satuan]]);
        if (kategori) sh.getRange(row, 4).setValue(kategori);
        status = 'diperbarui';
      } else {
        row = last + 1;
        sh.getRange(row, 1, 1, 4).setValues([[nama, harga, satuan, kategori]]);
        status = 'ditambah';
      }
      sh.getRange(row, 2).setNumberFormat('#,##0');
      SpreadsheetApp.flush();
      return json_(Object.assign({ ok: true, status: status, nama: nama }, readAll_()));
    } finally {
      lock.releaseLock();
    }
  }

  return json_({ ok: false, error: 'aksi' });
}

function readAll_() {
  const ss = SpreadsheetApp.getActive();
  const h = ss.getSheetByName(TAB_HARGA).getDataRange().getValues().slice(1);
  const prices = h
    .map(function (r) {
      return { n: String(r[0] || '').trim(), h: Number(String(r[1]).replace(/[^\d.]/g, '')) || 0, s: String(r[2] || '').trim() || 'pcs', k: String(r[3] || '').trim() };
    })
    .filter(function (p) { return p.n && p.h > 0; });
  const kios = {};
  const ks = ss.getSheetByName(TAB_KIOS);
  if (ks) {
    ks.getRange('A1:B20').getDisplayValues().slice(1).forEach(function (r) {
      if (r[0]) kios[String(r[0]).trim().toLowerCase()] = String(r[1] || '').trim();
    });
  }
  return { prices: prices, kios: kios };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
