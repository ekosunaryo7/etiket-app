const SCRIPT_URL = 'URL_WEB_APP_ANDA_DISINI'; 

async function rpc(fn, ...args) {
  try {
    const response = await fetch(SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify({ func: fn, args: args }),
      headers: {
        'Content-Type': 'text/plain;charset=utf-8',
      }
    });
    
    const result = await response.json();
    
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