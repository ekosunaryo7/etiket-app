/**
 * E-Tiket Event — REST API Backend
 * Google Apps Script + Google Sheets.
 */

var SHEET_PENGATURAN = 'Pengaturan';
var SHEET_TIKET = 'Tiket';
var SHEET_PESERTA = 'Peserta';

var KOLOM_PESERTA = ['ID', 'Waktu Daftar', 'Nama', 'Email', 'WhatsApp', 'Kode Tiket', 'Instansi', 'Status', 'Waktu Check-in', 'Petugas'];
var TZ = 'Asia/Jakarta';
var HURUF_ID = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; 

// ========================= API Jembatan (Vercel ke GAS) =========================

function doPost(e) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    var body = JSON.parse(e.postData.contents);
    var funcName = body.func;
    var args = body.args || [];
    
    // Mengeksekusi fungsi berdasarkan request dari frontend
    var result = this[funcName].apply(this, args);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: result
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Backend API E-Tiket Aktif.").setMimeType(ContentService.MimeType.TEXT);
}

// ========================= Setup =========================

function setup() {
  var ss = SpreadsheetApp.getActive();

  var peng = ss.getSheetByName(SHEET_PENGATURAN) || ss.insertSheet(SHEET_PENGATURAN);
  if (peng.getLastRow() === 0) {
    var hariH = new Date(Date.now() + 21 * 86400000);
    var isi = [
      ['Kunci', 'Nilai', 'Keterangan'],
      ['nama_acara', 'Workshop Digital UMKM 2026', 'Judul besar di halaman pendaftaran'],
      ['tagline', 'Belajar jualan online, desain konten, dan pembukuan sederhana dalam satu hari.', 'Kalimat pendek di bawah judul'],
      ['tanggal', Utilities.formatDate(hariH, TZ, 'yyyy-MM-dd'), 'Format yyyy-mm-dd'],
      ['jam_mulai', '08:30', 'Format HH:mm'],
      ['jam_selesai', '15:00', 'Format HH:mm'],
      ['tempat', 'Aula Serbaguna Kecamatan', 'Nama tempat'],
      ['alamat', 'Jl. Merdeka No. 10, Yogyakarta', 'Alamat lengkap'],
      ['link_maps', 'https://maps.google.com/?q=Yogyakarta', 'Link Google Maps (boleh kosong)'],
      ['penyelenggara', 'Komunitas UMKM Naik Kelas', 'Nama penyelenggara'],
      ['deskripsi', 'Acara gratis untuk pelaku UMKM. Peserta mendapat materi, makan siang, dan sertifikat. Tunjukkan e-tiket QR saat registrasi ulang di lokasi.', 'Penjelasan acara'],
      ['pendaftaran_dibuka', 'YA', 'YA atau TIDAK'],
      ['batas_pendaftaran', Utilities.formatDate(new Date(hariH.getTime() - 86400000), TZ, 'yyyy-MM-dd') + ' 23:59', 'yyyy-mm-dd HH:mm (boleh kosong)'],
      ['kirim_email', 'YA', 'Kirim e-tiket ke email peserta (YA/TIDAK)'],
      ['warna_utama', '#ffc53d', 'Warna aksen tombol & tiket (kode hex)'],
      ['password_admin', 'admin123', 'WAJIB DIGANTI. Untuk /admin.html'],
      ['pin_panitia', '246810', 'WAJIB DIGANTI. Untuk /scan.html'],
    ];
    peng.getRange(1, 1, isi.length, 3).setNumberFormat('@').setValues(isi);
    peng.getRange('A1:C1').setFontWeight('bold').setBackground('#10162f').setFontColor('#ffffff');
    peng.setFrozenRows(1);
    peng.setColumnWidth(1, 170); peng.setColumnWidth(2, 360); peng.setColumnWidth(3, 300);
  }

  var tik = ss.getSheetByName(SHEET_TIKET) || ss.insertSheet(SHEET_TIKET);
  if (tik.getLastRow() === 0) {
    var tiket = [
      ['Kode', 'Nama', 'Keterangan', 'Kuota', 'Aktif'],
      ['REG', 'Reguler', 'Materi, makan siang, sertifikat', 150, 'YA'],
      ['VIP', 'VIP', 'Kursi depan + sesi tanya jawab dengan pemateri', 30, 'YA'],
      ['MHS', 'Mahasiswa', 'Wajib menunjukkan kartu mahasiswa saat registrasi ulang', 60, 'YA'],
    ];
    tik.getRange(1, 1, tiket.length, 5).setValues(tiket);
    tik.getRange('A1:E1').setFontWeight('bold').setBackground('#10162f').setFontColor('#ffffff');
    tik.setFrozenRows(1);
  }

  var pes = ss.getSheetByName(SHEET_PESERTA) || ss.insertSheet(SHEET_PESERTA);
  if (pes.getLastRow() === 0) {
    pes.getRange(1, 1, 1, KOLOM_PESERTA.length).setValues([KOLOM_PESERTA]);
    pes.getRange(1, 1, 1, KOLOM_PESERTA.length).setFontWeight('bold').setBackground('#10162f').setFontColor('#ffffff');
    pes.setFrozenRows(1);
  }
  pes.getRange('A:J').setNumberFormat('@');

  rahasia_();
  return 'Setup selesai. Isi sheet Pengaturan & Tiket, lalu Deploy sebagai Web app.';
}

// ========================= Data dasar =========================

function bacaPengaturan_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PENGATURAN);
  var cfg = {};
  if (!sh) return cfg;
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues().forEach(function (r) {
    if (r[0]) cfg[String(r[0]).trim()] = r[1] instanceof Date ? formatTanggalSel_(r[1]) : String(r[1]).trim();
  });
  return cfg;
}

function formatTanggalSel_(d) {
  if (d.getFullYear() < 1901) return Utilities.formatDate(d, TZ, 'HH:mm');
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function bacaJenisTiket_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TIKET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(function (r) { return r[0] && String(r[4]).toUpperCase() !== 'TIDAK'; })
    .map(function (r) {
      return { kode: String(r[0]).trim().toUpperCase(), nama: String(r[1]), keterangan: String(r[2] || ''), kuota: Number(r[3]) || 0 };
    });
}

function sheetPeserta_() {
  return SpreadsheetApp.getActive().getSheetByName(SHEET_PESERTA);
}

function bacaPeserta_() {
  var sh = sheetPeserta_();
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, KOLOM_PESERTA.length).getValues().map(function (r, i) {
    return {
      baris: i + 2, id: String(r[0]), waktu: String(r[1]), nama: String(r[2]), email: String(r[3]),
      wa: String(r[4]), kode: String(r[5]), instansi: String(r[6] || ''), status: String(r[7] || 'TERDAFTAR'),
      checkin: String(r[8] || ''), petugas: String(r[9] || ''),
    };
  });
}

function sekarang_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

// ========================= Publik =========================

function getAcara() {
  var cfg = bacaPengaturan_();
  var peserta = bacaPeserta_();
  var terpakai = {};
  peserta.forEach(function (p) { terpakai[p.kode] = (terpakai[p.kode] || 0) + 1; });
  var tiket = bacaJenisTiket_().map(function (t) {
    return { kode: t.kode, nama: t.nama, keterangan: t.keterangan, kuota: t.kuota, sisa: Math.max(0, t.kuota - (terpakai[t.kode] || 0)) };
  });
  return {
    nama: cfg.nama_acara || 'Acara',
    tagline: cfg.tagline || '',
    tanggal: cfg.tanggal || '',
    jamMulai: cfg.jam_mulai || '',
    jamSelesai: cfg.jam_selesai || '',
    tempat: cfg.tempat || '',
    alamat: cfg.alamat || '',
    maps: /^https:\/\//.test(cfg.link_maps || '') ? cfg.link_maps : '',
    penyelenggara: cfg.penyelenggara || '',
    deskripsi: cfg.deskripsi || '',
    warna: /^#[0-9a-fA-F]{6}$/.test(cfg.warna_utama || '') ? cfg.warna_utama : '#ffc53d',
    dibuka: statusPendaftaran_(cfg).dibuka,
    alasanTutup: statusPendaftaran_(cfg).alasan,
    totalPeserta: peserta.length,
    tiket: tiket,
  };
}

function statusPendaftaran_(cfg) {
  if (String(cfg.pendaftaran_dibuka || 'YA').toUpperCase() !== 'YA') return { dibuka: false, alasan: 'Pendaftaran sedang ditutup oleh panitia.' };
  var batas = String(cfg.batas_pendaftaran || '').trim();
  if (batas && Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') > batas) {
    return { dibuka: false, alasan: 'Pendaftaran sudah ditutup sejak ' + batas + '.' };
  }
  return { dibuka: true, alasan: '' };
}

function idBaru_(sudahAda) {
  var acak = function (n) {
    var s = '';
    for (var i = 0; i < n; i++) s += HURUF_ID.charAt(Math.floor(Math.random() * HURUF_ID.length));
    return s;
  };
  var id;
  do { id = 'TK-' + acak(4) + '-' + acak(4); } while (sudahAda[id]);
  return id;
}

function daftar(form) {
  form = form || {};
  if (form.website) throw new Error('Pendaftaran ditolak.'); 

  var nama = String(form.nama || '').trim().replace(/\s+/g, ' ');
  var email = String(form.email || '').trim().toLowerCase();
  var wa = String(form.wa || '').replace(/[^\d+]/g, '');
  var kode = String(form.kode || '').trim().toUpperCase();
  var instansi = String(form.instansi || '').trim().slice(0, 80);

  if (nama.length < 3 || nama.length > 80) throw new Error('Nama lengkap minimal 3 huruf.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email) || email.length > 100) throw new Error('Format email tidak valid.');
  if (wa.replace(/\D/g, '').length < 9 || wa.length > 16) throw new Error('Nomor WhatsApp tidak valid.');
  if (/^0/.test(wa)) wa = '62' + wa.slice(1);
  wa = wa.replace(/^\+/, '');

  var cfg = bacaPengaturan_();
  var st = statusPendaftaran_(cfg);
  if (!st.dibuka) throw new Error(st.alasan);

  var jenis = bacaJenisTiket_().filter(function (t) { return t.kode === kode; })[0];
  if (!jenis) throw new Error('Pilih jenis tiket.');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var tiket;
  try {
    var peserta = bacaPeserta_();
    var dobel = peserta.filter(function (p) { return p.email === email; })[0];
    if (dobel) throw new Error('Email ini sudah terdaftar. Gunakan menu "Lihat tiketku" untuk menampilkan tiketmu.');
    var terpakai = peserta.filter(function (p) { return p.kode === kode; }).length;
    if (terpakai >= jenis.kuota) throw new Error('Maaf, kuota tiket ' + jenis.nama + ' sudah habis.');

    var sudahAda = {};
    peserta.forEach(function (p) { sudahAda[p.id] = true; });
    var id = idBaru_(sudahAda);
    var waktu = sekarang_();
    sheetPeserta_().appendRow([id, waktu, nama, email, wa, kode, instansi, 'TERDAFTAR', '', '']);
    SpreadsheetApp.flush();
    tiket = { id: id, nama: nama, email: email, kode: kode, jenis: jenis.nama, instansi: instansi, waktu: waktu, status: 'TERDAFTAR', checkin: '' };
  } finally {
    lock.releaseLock();
  }

  if (String(cfg.kirim_email || 'YA').toUpperCase() === 'YA') kirimEmailTiket_(tiket, cfg);
  return tiket;
}

function kirimEmailTiket_(t, cfg) {
  try {
    var url = ScriptApp.getService().getUrl() + '?t=' + encodeURIComponent(t.id);
    var warna = /^#[0-9a-fA-F]{6}$/.test(cfg.warna_utama || '') ? cfg.warna_utama : '#ffc53d';
    var html =
      '<div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;color:#10162f">' +
      '<div style="background:#10162f;color:#fff;padding:24px;border-radius:16px 16px 0 0">' +
      '<div style="font-size:12px;letter-spacing:2px;color:' + warna + ';font-weight:bold">E-TIKET</div>' +
      '<div style="font-size:22px;font-weight:bold;margin-top:6px">' + esc_(cfg.nama_acara) + '</div>' +
      '<div style="margin-top:6px;opacity:.75">' + esc_(cfg.tanggal) + ' · ' + esc_(cfg.jam_mulai) + ' WIB · ' + esc_(cfg.tempat) + '</div></div>' +
      '<div style="border:1px solid #e3e6ee;border-top:0;padding:24px;border-radius:0 0 16px 16px">' +
      '<p>Halo <b>' + esc_(t.nama) + '</b>, pendaftaranmu berhasil.</p>' +
      '<p>Jenis tiket: <b>' + esc_(t.jenis) + '</b><br>Kode tiket: <b style="font-size:18px;letter-spacing:1px">' + esc_(t.id) + '</b></p>' +
      '<p><a href="' + url + '" style="display:inline-block;background:' + warna + ';color:#10162f;padding:12px 20px;border-radius:10px;font-weight:bold;text-decoration:none">Buka e-tiket (QR)</a></p>' +
      '<p style="color:#6b7280;font-size:13px">Tunjukkan QR pada e-tiket saat registrasi ulang. Simpan email ini.</p></div></div>';
    MailApp.sendEmail({ to: t.email, subject: 'E-tiket ' + cfg.nama_acara + ' · ' + t.id, htmlBody: html, name: cfg.penyelenggara || cfg.nama_acara });
  } catch (err) {
    console.warn('Email tiket gagal: ' + err); 
  }
}

function getTiket(id) {
  id = String(id || '').trim().toUpperCase();
  if (!/^TK-[A-Z0-9]{4}-[A-Z0-9]{4}$/.test(id)) throw new Error('Kode tiket tidak valid.');
  var p = bacaPeserta_().filter(function (x) { return x.id === id; })[0];
  if (!p) throw new Error('Tiket tidak ditemukan.');
  return publikTiket_(p);
}

function publikTiket_(p) {
  var jenis = bacaJenisTiket_().filter(function (t) { return t.kode === p.kode; })[0];
  var at = p.email.indexOf('@');
  return {
    id: p.id, nama: p.nama, instansi: p.instansi, kode: p.kode, jenis: jenis ? jenis.nama : p.kode,
    email: at > 2 ? p.email.slice(0, 2) + '***' + p.email.slice(at) : '***' + p.email.slice(at),
    waktu: p.waktu, status: p.status, checkin: p.checkin,
  };
}

function cariTiket(email, wa4) {
  email = String(email || '').trim().toLowerCase();
  wa4 = String(wa4 || '').replace(/\D/g, '');
  if (!email || wa4.length !== 4) throw new Error('Isi email dan 4 digit terakhir nomor WhatsApp.');
  var cache = CacheService.getScriptCache();
  var kunci = 'cari_' + email;
  var coba = Number(cache.get(kunci) || 0);
  if (coba >= 5) throw new Error('Terlalu banyak percobaan. Coba lagi 15 menit lagi.');
  var p = bacaPeserta_().filter(function (x) { return x.email === email && x.wa.slice(-4) === wa4; })[0];
  if (!p) {
    cache.put(kunci, String(coba + 1), 900);
    throw new Error('Tiket tidak ditemukan. Periksa email dan nomor WhatsApp.');
  }
  return publikTiket_(p);
}

// ========================= Token panitia & admin =========================

function rahasia_() {
  var props = PropertiesService.getScriptProperties();
  var r = props.getProperty('RAHASIA');
  if (!r) { r = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('RAHASIA', r); }
  return r;
}

function tandaTangan_(isi, kunciTambahan) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(isi + '|' + kunciTambahan, rahasia_()));
}

function buatToken_(peran, jam) {
  var cfg = bacaPengaturan_();
  var isi = peran + '.' + (Date.now() + jam * 3600000);
  return isi + '.' + tandaTangan_(isi, peran === 'admin' ? cfg.password_admin : cfg.pin_panitia);
}

function cekToken_(token, peran) {
  var bagian = String(token || '').split('.');
  if (bagian.length !== 3 || bagian[0] !== peran && !(peran === 'panitia' && bagian[0] === 'admin')) throw new Error('Sesi tidak valid. Silakan masuk lagi.');
  if (Number(bagian[1]) < Date.now()) throw new Error('Sesi berakhir. Silakan masuk lagi.');
  var cfg = bacaPengaturan_();
  var isi = bagian[0] + '.' + bagian[1];
  var harus = tandaTangan_(isi, bagian[0] === 'admin' ? cfg.password_admin : cfg.pin_panitia);
  if (harus !== bagian[2]) throw new Error('Sesi tidak valid. Silakan masuk lagi.');
  return bagian[0];
}

function masuk_(peran, rahasiaInput) {
  var cache = CacheService.getScriptCache();
  var kunci = 'gagal_' + peran;
  var gagal = Number(cache.get(kunci) || 0);
  if (gagal >= 8) throw new Error('Terlalu banyak percobaan salah. Tunggu 15 menit.');
  var cfg = bacaPengaturan_();
  var benar = peran === 'admin' ? cfg.password_admin : cfg.pin_panitia;
  if (!benar || String(rahasiaInput || '') !== String(benar)) {
    cache.put(kunci, String(gagal + 1), 900);
    throw new Error(peran === 'admin' ? 'Password salah.' : 'PIN salah.');
  }
  cache.remove(kunci);
  return buatToken_(peran, peran === 'admin' ? 12 : 24);
}

function masukPanitia(pin) { return masuk_('panitia', pin); }
function masukAdmin(password) { return masuk_('admin', password); }

// ========================= Check-in =========================

function checkin(token, kodeQr, petugas) {
  cekToken_(token, 'panitia');
  var id = String(kodeQr || '').trim().toUpperCase();
  var cocok = id.match(/TK-[A-Z0-9]{4}-[A-Z0-9]{4}/);
  if (!cocok) return { hasil: 'INVALID', pesan: 'QR bukan e-tiket acara ini.' };
  id = cocok[0];
  petugas = String(petugas || 'Panitia').trim().slice(0, 40) || 'Panitia';

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var p = bacaPeserta_().filter(function (x) { return x.id === id; })[0];
    if (!p) return { hasil: 'INVALID', pesan: 'Tiket ' + id + ' tidak terdaftar.' };
    var info = publikTiket_(p);
    if (p.status === 'HADIR') {
      return { hasil: 'SUDAH', pesan: 'Tiket sudah dipakai check-in pukul ' + p.checkin.slice(11, 16) + ' oleh ' + p.petugas + '.', tiket: info };
    }
    var waktu = sekarang_();
    sheetPeserta_().getRange(p.baris, 8, 1, 3).setValues([['HADIR', waktu, petugas]]);
    SpreadsheetApp.flush();
    info.status = 'HADIR';
    info.checkin = waktu;
    return { hasil: 'OK', pesan: 'Selamat datang, ' + p.nama + '!', tiket: info };
  } finally {
    lock.releaseLock();
  }
}

function statistikPanitia(token) {
  cekToken_(token, 'panitia');
  var peserta = bacaPeserta_();
  var hadir = peserta.filter(function (p) { return p.status === 'HADIR'; });
  hadir.sort(function (a, b) { return a.checkin < b.checkin ? 1 : -1; });
  return {
    total: peserta.length,
    hadir: hadir.length,
    terakhir: hadir.slice(0, 5).map(function (p) { return { nama: p.nama, jenis: p.kode, jam: p.checkin.slice(11, 16) }; }),
  };
}

function cariPeserta(token, kata) {
  cekToken_(token, 'panitia');
  kata = String(kata || '').trim().toLowerCase();
  if (kata.length < 3) return [];
  return bacaPeserta_()
    .filter(function (p) { return p.nama.toLowerCase().indexOf(kata) >= 0 || p.id.toLowerCase().indexOf(kata) >= 0 || p.wa.slice(-4) === kata; })
    .slice(0, 8)
    .map(publikTiket_);
}

// ========================= Admin =========================

function getDashboard(token) {
  cekToken_(token, 'admin');
  var acara = getAcara();
  var peserta = bacaPeserta_();
  var hadir = peserta.filter(function (p) { return p.status === 'HADIR'; });

  var perJenis = acara.tiket.map(function (t) {
    var semua = peserta.filter(function (p) { return p.kode === t.kode; });
    return { kode: t.kode, nama: t.nama, kuota: t.kuota, terdaftar: semua.length, hadir: semua.filter(function (p) { return p.status === 'HADIR'; }).length };
  });

  var harian = {};
  for (var i = 13; i >= 0; i--) harian[Utilities.formatDate(new Date(Date.now() - i * 86400000), TZ, 'yyyy-MM-dd')] = 0;
  peserta.forEach(function (p) { var d = p.waktu.slice(0, 10); if (d in harian) harian[d]++; });

  var perJam = {};
  hadir.forEach(function (p) { var j = p.checkin.slice(11, 13); perJam[j] = (perJam[j] || 0) + 1; });

  return {
    acara: { nama: acara.nama, tanggal: acara.tanggal, jamMulai: acara.jamMulai, tempat: acara.tempat, dibuka: acara.dibuka },
    total: peserta.length,
    hadir: hadir.length,
    kuotaTotal: acara.tiket.reduce(function (s, t) { return s + t.kuota; }, 0),
    perJenis: perJenis,
    harian: Object.keys(harian).map(function (d) { return { tanggal: d, jumlah: harian[d] }; }),
    perJam: Object.keys(perJam).sort().map(function (j) { return { jam: j, jumlah: perJam[j] }; }),
    instansiTeratas: hitungTeratas_(peserta.map(function (p) { return p.instansi; }).filter(String), 5),
  };
}

function hitungTeratas_(daftar, n) {
  var hit = {};
  daftar.forEach(function (x) { var k = x.trim(); if (k) hit[k] = (hit[k] || 0) + 1; });
  return Object.keys(hit).map(function (k) { return { nama: k, jumlah: hit[k] }; })
    .sort(function (a, b) { return b.jumlah - a.jumlah; }).slice(0, n);
}

function daftarPeserta(token, opsi) {
  cekToken_(token, 'admin');
  opsi = opsi || {};
  var kata = String(opsi.cari || '').trim().toLowerCase();
  var status = String(opsi.status || '');
  var kode = String(opsi.kode || '');
  var hasil = bacaPeserta_().filter(function (p) {
    if (status && p.status !== status) return false;
    if (kode && p.kode !== kode) return false;
    if (!kata) return true;
    return (p.nama + ' ' + p.email + ' ' + p.id + ' ' + p.instansi + ' ' + p.wa).toLowerCase().indexOf(kata) >= 0;
  });
  hasil.sort(function (a, b) { return a.waktu < b.waktu ? 1 : -1; });
  var per = 20, hal = Math.max(1, Number(opsi.hal) || 1);
  return {
    total: hasil.length,
    hal: hal,
    jumlahHal: Math.max(1, Math.ceil(hasil.length / per)),
    baris: hasil.slice((hal - 1) * per, hal * per).map(function (p) {
      return { id: p.id, nama: p.nama, email: p.email, wa: p.wa, kode: p.kode, instansi: p.instansi, waktu: p.waktu, status: p.status, checkin: p.checkin, petugas: p.petugas };
    }),
  };
}

function ubahStatus(token, id, hadir) {
  cekToken_(token, 'admin');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var p = bacaPeserta_().filter(function (x) { return x.id === String(id); })[0];
    if (!p) throw new Error('Peserta tidak ditemukan.');
    sheetPeserta_().getRange(p.baris, 8, 1, 3).setValues([hadir ? ['HADIR', sekarang_(), 'Admin'] : ['TERDAFTAR', '', '']]);
    SpreadsheetApp.flush();
    return true;
  } finally {
    lock.releaseLock();
  }
}

function eksporCsv(token) {
  cekToken_(token, 'admin');
  var sel = function (v) {
    var s = String(v == null ? '' : v);
    if (/^[=+\-@]/.test(s)) s = "'" + s; 
    return '"' + s.replace(/"/g, '""') + '"';
  };
  var baris = [KOLOM_PESERTA.map(sel).join(';')];
  bacaPeserta_().forEach(function (p) {
    baris.push([p.id, p.waktu, p.nama, p.email, p.wa, p.kode, p.instansi, p.status, p.checkin, p.petugas].map(sel).join(';'));
  });
  return baris.join('\r\n');
}

function buatDataContoh() {
  setup();
  var depan = ['Andi', 'Budi', 'Citra', 'Dewi', 'Eka', 'Fajar', 'Gita', 'Hana', 'Indra', 'Joko', 'Kartika', 'Lestari', 'Maya', 'Nanda', 'Oki', 'Putri', 'Rizky', 'Sari', 'Tono', 'Umi', 'Vina', 'Wahyu', 'Yuni', 'Zaki'];
  var belakang = ['Pratama', 'Saputra', 'Lestari', 'Wijaya', 'Kusuma', 'Hidayat', 'Rahmawati', 'Nugroho', 'Santoso', 'Permata', 'Siregar', 'Hasibuan'];
  var instansi = ['Toko Berkah Jaya', 'Kopi Senja', 'Batik Ayu', 'Universitas Gadjah Mada', 'UIN Sunan Kalijaga', 'Keripik Bu Tini', 'Laundry Bersih', 'Rumah Jahit Mawar', 'Bakso Mas Gino', 'Percetakan Cahaya', ''];
  var jenis = [['REG', 110], ['VIP', 22], ['MHS', 28]];
  var sh = sheetPeserta_();
  var sudahAda = {};
  var rows = [];
  var hariAcara = bacaPengaturan_().tanggal;
  var n = 0;
  jenis.forEach(function (j) {
    for (var i = 0; i < j[1]; i++) {
      var nama = depan[(n * 7 + i) % depan.length] + ' ' + belakang[(n * 5 + i * 3) % belakang.length];
      var id = idBaru_(sudahAda); sudahAda[id] = true;
      var daftarMs = Date.now() - Math.floor(Math.random() * 13 * 86400000) - Math.floor(Math.random() * 36000000);
      var hadir = Math.random() < 0.62;
      var jam = 7 + Math.floor(Math.random() * 3);
      var menit = ('0' + Math.floor(Math.random() * 60)).slice(-2);
      rows.push([
        id, Utilities.formatDate(new Date(daftarMs), TZ, 'yyyy-MM-dd HH:mm:ss'), nama,
        nama.toLowerCase().replace(/ /g, '.') + n + '@contoh.id', '62812' + (10000000 + n * 7919) % 100000000,
        j[0], instansi[(n * 3) % instansi.length], hadir ? 'HADIR' : 'TERDAFTAR',
        hadir ? hariAcara + ' 0' + jam + ':' + menit + ':00' : '', hadir ? ['Gerbang A', 'Gerbang B'][n % 2] : '',
      ]);
      n++;
    }
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, KOLOM_PESERTA.length).setValues(rows);
  return rows.length + ' peserta contoh ditambahkan.';
}