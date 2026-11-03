// Ganti dengan Web App URL dari Google Apps Script yang baru Anda deploy (Versi New Deployment)
const SCRIPT_URL = 'URL_WEB_APP_ANDA_DISINI'; 

async function rpc(fn, ...args) {
  try {
    // Memasukkan nama fungsi dan data ke dalam parameter URL (GET Method)
    const url = new URL(SCRIPT_URL);
    url.searchParams.append('func', fn);
    url.searchParams.append('args', JSON.stringify(args));
    url.searchParams.append('ts', Date.now()); // Mencegah browser menyimpan cache (cache-buster)

    const response = await fetch(url.toString(), {
      method: 'GET'
    });
    
    // Ambil respons sebagai teks terlebih dahulu untuk menghindari crash JSON parse
    const text = await response.text();
    let result;
    
    try {
      result = JSON.parse(text);
    } catch (e) {
      console.error("Respons backend bukan JSON:", text);
      throw new Error("Gagal terhubung ke server. Pastikan Web App GAS di-deploy dengan akses 'Anyone'.");
    }
    
    if (!result.success) {
      throw new Error(result.error);
    }
    
    return result.data;
  } catch (error) {
    throw error;
  }
}

function $(s, el) { return (el || document).querySelector(s); }
function $$(s, el) { return Array.prototype.slice.call((el || document).querySelectorAll(s)); }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function toast(msg, err) {
  var t = $('#toast'); if (!t) return;
  t.textContent = msg; t.className = 'toast' + (err ? ' err' : ''); t.style.display = 'block';
  clearTimeout(t._h); t._h = setTimeout(function () { t.style.display = 'none'; }, 3500);
}

var BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
var HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
function tglPanjang(s) {
  var m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/); if (!m) return s || '';
  var d = new Date(+m[1], +m[2] - 1, +m[3]);
  return HARI[d.getDay()] + ', ' + (+m[3]) + ' ' + BULAN[+m[2] - 1] + ' ' + m[1];
}
