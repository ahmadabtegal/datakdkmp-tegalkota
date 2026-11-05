/**
 * PORTAL MONITORING KDKMP KOTA TEGAL
 * Backend langsung untuk GitHub Pages (TANPA Cloudflare Worker).
 *
 * Transport:
 * - GET JSONP: data publik dan hasil transaksi.
 * - POST application/x-www-form-urlencoded: login dan perubahan data.
 * - POST result disimpan sementara di CacheService lalu diambil melalui JSONP.
 *
 * PERSIAPAN:
 * 1. Konversi Excel KDKMP Kota Tegal.xlsx menjadi Google Spreadsheet.
 * 2. Buka Extensions > Apps Script dari spreadsheet tersebut.
 * 3. Tempel seluruh file ini ke Code.gs.
 * 4. Jalankan setupDatabase() satu kali.
 * 5. Project Settings > Script Properties: ADMIN_PASSWORD = password admin yang kuat.
 * 6. Deploy > New deployment > Web app; Execute as Me; Who has access: Anyone.
 */

const CFG = {
  MASTER: 'MASTER_KDKMP',
  DOCS: 'DOKUMENTASI',
  ADMIN_PASSWORD_KEY: 'ADMIN_PASSWORD',
  TOKEN_PREFIX: 'KDKMP_ADMIN_',
  TOKEN_TTL: 21600,
  RESULT_PREFIX: 'KDKMP_RESULT_',
  RESULT_TTL: 180,
  LOGIN_FAIL_KEY: 'KDKMP_LOGIN_FAILS',
  LOGIN_LOCK_KEY: 'KDKMP_LOGIN_LOCK',
  LOGIN_FAIL_LIMIT: 8,
  LOGIN_LOCK_SECONDS: 900
};

const MASTER_HEADERS = [
  'ID_KDKMP','Kelurahan','Kecamatan','Nama_Koperasi','NIB','Nomor_Badan_Hukum','Tanggal_Pengesahan',
  'Nama_Ketua','No_HP_Ketua','Alamat_Sekretariat','Lokasi_Bangunan','Koordinat','Zona_Peruntukan',
  'Status_Lahan','Luas_Tanah_m2','Status_Kepemilikan_Lahan','Status_Bangunan','Progres_Pembangunan',
  'Tanggal_Mulai','Target_Selesai','Kondisi_Terakhir','Kendala_Pembangunan',
  'Status_Kelembagaan','Jumlah_Pengurus','Jumlah_Pengawas','Jumlah_Anggota','Status_AD_ART','Status_RAT',
  'Status_Operasional','Jenis_Usaha','Jumlah_Unit_Usaha','Jumlah_Tenaga_Kerja','Omzet_Bulanan',
  'Nama_Aset_Utama','Jumlah_Aset','Nilai_Aset','Kondisi_Aset','Catatan_Publik','Diperbarui_Pada','Diperbarui_Oleh'
];
const DOC_HEADERS = ['ID_DOKUMENTASI','ID_KDKMP','Tanggal','Kategori','Judul','Keterangan','URL_FOTO','Dibuat_Pada','Dibuat_Oleh'];
const DISTRICTS = ['Margadana','Tegal Barat','Tegal Selatan','Tegal Timur'];

/** Endpoint GET: JSONP untuk data publik dan polling hasil POST. */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || 'ping');
  let result;
  try {
    if (action === 'public') {
      result = getPublicDashboardData();
    } else if (action === 'result') {
      const requestId = String(p.requestId || '');
      if (!/^[a-f0-9]{32,64}$/i.test(requestId)) throw new Error('ID permintaan tidak valid.');
      const cached = CacheService.getScriptCache().get(CFG.RESULT_PREFIX + requestId);
      result = cached ? JSON.parse(cached) : {ok:false,pending:true};
    } else if (action === 'ping') {
      result = {ok:true,service:'KDKMP API',message:'API Apps Script aktif.'};
    } else {
      throw new Error('Aksi GET tidak dikenal.');
    }
  } catch (err) {
    result = {ok:false,error:safeError_(err)};
  }
  return outputJsonp_(result, p.callback);
}

/** Endpoint POST: menerima form-urlencoded agar tidak memicu CORS preflight. */
function doPost(e) {
  const p = (e && e.parameter) || {};
  const requestId = String(p.requestId || '');
  let result;
  try {
    if (!/^[a-f0-9]{32,64}$/i.test(requestId)) throw new Error('ID permintaan tidak valid.');
    const raw = String(p.payload || '');
    if (!raw || raw.length > 90000) throw new Error('Payload kosong atau terlalu besar.');
    const req = JSON.parse(raw);
    result = dispatchApi_(req || {});
    if (!result || typeof result !== 'object') result = {ok:true,result:result};
    if (result.ok === undefined) result.ok = true;
  } catch (err) {
    result = {ok:false,error:safeError_(err)};
  }
  try {
    CacheService.getScriptCache().put(CFG.RESULT_PREFIX + requestId, JSON.stringify(result), CFG.RESULT_TTL);
  } catch (cacheErr) {
    // Jangan pernah mengembalikan informasi sensitif dalam halaman respons POST.
    console.error('Gagal menyimpan hasil transaksi: ' + safeError_(cacheErr));
  }
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

function outputJsonp_(obj, callback) {
  const json = JSON.stringify(obj);
  const cb = String(callback || '');
  // Hanya nama callback sederhana yang diperbolehkan; mencegah injeksi JavaScript.
  if (cb && /^[A-Za-z_$][0-9A-Za-z_$]{0,80}$/.test(cb)) {
    return ContentService.createTextOutput('/**/' + cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function dispatchApi_(req) {
  const args = Array.isArray(req.args) ? req.args : [];
  switch (String(req.action || '')) {
    case 'adminLogin': return adminLogin(args[0]);
    case 'adminLogout': return adminLogout(args[0]);
    case 'getAdminData': return getAdminData(args[0]);
    case 'saveKdkmp': return saveKdkmp(args[0], args[1]);
    case 'addDocumentation': return addDocumentation(args[0], args[1]);
    case 'deleteDocumentation': return deleteDocumentation(args[0], args[1]);
    default: throw new Error('Perintah API tidak dikenal.');
  }
}

/** Jalankan satu kali dari spreadsheet hasil konversi Excel. */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script dari Google Spreadsheet database KDKMP.');
  let master = ss.getSheetByName(CFG.MASTER);
  if (!master) master = ss.insertSheet(CFG.MASTER);
  ensureHeaders_(master, MASTER_HEADERS);

  // Impor 27 data wilayah dari sheet Excel asal jika master masih kosong.
  if (master.getLastRow() < 2) {
    const source = ss.getSheetByName('Sheet1');
    if (source && source.getLastRow() > 1) {
      const width = Math.min(8, source.getLastColumn());
      const values = source.getRange(1, 1, source.getLastRow(), width).getDisplayValues();
      const rows = [];
      for (let i = 1; i < values.length; i++) {
        const r = values[i];
        const kel = String(r[1] || '').trim();
        const kec = String(r[2] || '').trim();
        if (!kel || !kec) continue;
        const obj = emptyObject_(MASTER_HEADERS);
        obj.ID_KDKMP = 'KDKMP-' + String(rows.length + 1).padStart(3, '0');
        obj.Kelurahan = kel;
        obj.Kecamatan = kec;
        obj.Status_Bangunan = r[3] || 'Belum Diisi';
        obj.Progres_Pembangunan = /sudah terbangun/i.test(String(r[3] || '')) ? 100 : 0;
        obj.NIB = String(r[4] || '').trim();
        obj.Lokasi_Bangunan = r[5] || '';
        obj.Koordinat = r[6] || '';
        obj.Zona_Peruntukan = r[7] || '';
        rows.push(MASTER_HEADERS.map(h => obj[h] == null ? '' : obj[h]));
      }
      if (rows.length) {
        master.getRange(2, 1, rows.length, MASTER_HEADERS.length).setNumberFormat('@');
        master.getRange(2, 1, rows.length, MASTER_HEADERS.length).setValues(rows);
      }
    }
  }

  let docs = ss.getSheetByName(CFG.DOCS);
  if (!docs) docs = ss.insertSheet(CFG.DOCS);
  ensureHeaders_(docs, DOC_HEADERS);
  [CFG.MASTER, CFG.DOCS].forEach(name => {
    const sh = ss.getSheetByName(name);
    sh.setFrozenRows(1);
    if (sh.getLastColumn() > 0) {
      sh.getRange(1,1,1,sh.getLastColumn()).setBackground('#173B65').setFontColor('#ffffff').setFontWeight('bold');
      sh.autoResizeColumns(1, Math.min(sh.getLastColumn(), 12));
    }
  });
  return 'Database siap. ' + Math.max(0, master.getLastRow() - 1) + ' data KDKMP tersedia.';
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    return;
  }
  const oldHeaders = sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn())).getDisplayValues()[0].map(String);
  if (oldHeaders.join('\u001f') === headers.join('\u001f')) return;
  const oldValues = sheet.getLastRow() > 1 ? sheet.getRange(2,1,sheet.getLastRow()-1,oldHeaders.length).getValues() : [];
  const oldIndex = {};
  oldHeaders.forEach((h,i) => { if (h) oldIndex[h] = i; });
  const migrated = oldValues.filter(r => r.some(v => v !== '' && v !== null))
    .map(r => headers.map(h => oldIndex[h] === undefined ? '' : r[oldIndex[h]]));
  sheet.clearContents();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  if (migrated.length) sheet.getRange(2,1,migrated.length,headers.length).setValues(migrated);
}

function emptyObject_(headers) { const o = {}; headers.forEach(h => o[h] = ''); return o; }
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Sheet ' + name + ' belum tersedia. Jalankan setupDatabase().');
  return s;
}
function readObjects_(name) {
  const s = sheet_(name), vals = s.getDataRange().getDisplayValues();
  if (vals.length < 2) return [];
  const heads = vals[0];
  return vals.slice(1).filter(r => r.some(v => v !== '')).map(r => {
    const o = {};
    heads.forEach((h,i) => o[h] = r[i] == null ? '' : r[i]);
    return o;
  });
}
function num_(v) { const n = Number(String(v || '').replace(',', '.')); return isFinite(n) ? n : 0; }
function safeError_(err) {
  const msg = err && err.message ? err.message : String(err || 'Terjadi kesalahan.');
  return msg.length > 300 ? msg.slice(0,300) : msg;
}

/** Daftar putih kolom publik: data pribadi dan nominal keuangan rinci tidak diekspos. */
function publicRecord_(r) {
  return {
    ID_KDKMP:r.ID_KDKMP, Kelurahan:r.Kelurahan, Kecamatan:r.Kecamatan,
    Nama_Koperasi:r.Nama_Koperasi, NIB:r.NIB, Nomor_Badan_Hukum:r.Nomor_Badan_Hukum,
    Tanggal_Pengesahan:r.Tanggal_Pengesahan, Alamat_Sekretariat:r.Alamat_Sekretariat,
    Lokasi_Bangunan:r.Lokasi_Bangunan, Koordinat:r.Koordinat, Zona_Peruntukan:r.Zona_Peruntukan,
    Status_Lahan:r.Status_Lahan, Luas_Tanah_m2:r.Luas_Tanah_m2,
    Status_Kepemilikan_Lahan:r.Status_Kepemilikan_Lahan, Status_Bangunan:r.Status_Bangunan,
    Progres_Pembangunan:num_(r.Progres_Pembangunan), Tanggal_Mulai:r.Tanggal_Mulai,
    Target_Selesai:r.Target_Selesai, Kondisi_Terakhir:r.Kondisi_Terakhir,
    Kendala_Pembangunan:r.Kendala_Pembangunan, Status_Kelembagaan:r.Status_Kelembagaan,
    Jumlah_Pengurus:r.Jumlah_Pengurus, Jumlah_Pengawas:r.Jumlah_Pengawas,
    Jumlah_Anggota:r.Jumlah_Anggota, Status_AD_ART:r.Status_AD_ART, Status_RAT:r.Status_RAT,
    Status_Operasional:r.Status_Operasional, Jenis_Usaha:r.Jenis_Usaha,
    Jumlah_Unit_Usaha:r.Jumlah_Unit_Usaha, Jumlah_Tenaga_Kerja:r.Jumlah_Tenaga_Kerja,
    Nama_Aset_Utama:r.Nama_Aset_Utama, Jumlah_Aset:r.Jumlah_Aset,
    Kondisi_Aset:r.Kondisi_Aset, Catatan_Publik:r.Catatan_Publik,
    Diperbarui_Pada:r.Diperbarui_Pada
  };
}

function getPublicDashboardData() {
  const rows = readObjects_(CFG.MASTER);
  const docs = readObjects_(CFG.DOCS).map(d => ({
    ID_DOKUMENTASI:d.ID_DOKUMENTASI, ID_KDKMP:d.ID_KDKMP, Tanggal:d.Tanggal,
    Kategori:d.Kategori, Judul:d.Judul, Keterangan:d.Keterangan, URL_FOTO:d.URL_FOTO
  }));
  const data = rows.map(publicRecord_);
  const total = data.length;
  const built = data.filter(r => /sudah terbangun|selesai/i.test(r.Status_Bangunan || '') || r.Progres_Pembangunan >= 100).length;
  const operational = data.filter(r => /operasional|beroperasi/i.test(r.Status_Operasional || '')).length;
  const avg = total ? Math.round(data.reduce((s,r) => s + Math.max(0,Math.min(100,r.Progres_Pembangunan)),0) / total) : 0;
  const byDistrict = DISTRICTS.map(k => {
    const list = data.filter(r => r.Kecamatan === k);
    return {
      kecamatan:k, total:list.length,
      terbangun:list.filter(r => /sudah terbangun|selesai/i.test(r.Status_Bangunan || '') || r.Progres_Pembangunan >= 100).length,
      progres:list.length ? Math.round(list.reduce((s,r) => s + r.Progres_Pembangunan,0) / list.length) : 0
    };
  });
  return {ok:true,updatedAt:new Date().toISOString(),stats:{total,terbangun:built,belum:total-built,operasional:operational,progres:avg},kecamatan:byDistrict,rows:data,dokumentasi:docs};
}

function adminLogin(password) {
  const cache = CacheService.getScriptCache();
  if (cache.get(CFG.LOGIN_LOCK_KEY)) throw new Error('Login sementara dikunci karena terlalu banyak percobaan. Coba lagi 15 menit kemudian.');
  const configured = PropertiesService.getScriptProperties().getProperty(CFG.ADMIN_PASSWORD_KEY);
  if (!configured) throw new Error('Password admin belum disetel. Tambahkan Script Property ADMIN_PASSWORD.');
  if (!constantTimeEquals_(String(password || ''), String(configured))) {
    const fails = Number(cache.get(CFG.LOGIN_FAIL_KEY) || 0) + 1;
    cache.put(CFG.LOGIN_FAIL_KEY, String(fails), CFG.LOGIN_LOCK_SECONDS);
    if (fails >= CFG.LOGIN_FAIL_LIMIT) cache.put(CFG.LOGIN_LOCK_KEY, '1', CFG.LOGIN_LOCK_SECONDS);
    throw new Error('Password admin tidak sesuai.');
  }
  cache.remove(CFG.LOGIN_FAIL_KEY);
  cache.remove(CFG.LOGIN_LOCK_KEY);
  const token = Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
  cache.put(CFG.TOKEN_PREFIX + token, '1', CFG.TOKEN_TTL);
  return {ok:true,token:token,expiresIn:CFG.TOKEN_TTL};
}
function constantTimeEquals_(a,b) {
  let diff = a.length ^ b.length;
  const n = Math.max(a.length,b.length);
  for (let i=0;i<n;i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}
function adminLogout(token) {
  if (token) CacheService.getScriptCache().remove(CFG.TOKEN_PREFIX + token);
  return {ok:true,message:'Anda telah keluar.'};
}
function assertAdmin_(token) {
  if (!token || CacheService.getScriptCache().get(CFG.TOKEN_PREFIX + token) !== '1') {
    throw new Error('Sesi admin berakhir. Silakan login kembali.');
  }
}
function getAdminData(token) {
  assertAdmin_(token);
  return {ok:true,rows:readObjects_(CFG.MASTER),dokumentasi:readObjects_(CFG.DOCS)};
}
function saveKdkmp(token, payload) {
  assertAdmin_(token);
  if (!payload || !payload.ID_KDKMP) throw new Error('ID KDKMP wajib diisi.');
  const sh = sheet_(CFG.MASTER), values = sh.getDataRange().getDisplayValues(), headers = values[0];
  const idCol = headers.indexOf('ID_KDKMP');
  let rowIndex = -1;
  for (let i=1;i<values.length;i++) if (values[i][idCol] === String(payload.ID_KDKMP)) { rowIndex = i+1; break; }
  if (rowIndex < 2) throw new Error('Data KDKMP tidak ditemukan.');
  const allowed = MASTER_HEADERS.filter(h => !['ID_KDKMP','Kelurahan','Kecamatan','Diperbarui_Pada','Diperbarui_Oleh'].includes(h));
  allowed.forEach(h => {
    if (Object.prototype.hasOwnProperty.call(payload,h)) {
      let value = payload[h];
      if (h === 'Progres_Pembangunan') value = Math.max(0,Math.min(100,num_(value)));
      const c = headers.indexOf(h);
      if (c >= 0) sh.getRange(rowIndex,c+1).setValue(value);
    }
  });
  const now = new Date();
  const updatedCol = headers.indexOf('Diperbarui_Pada');
  const byCol = headers.indexOf('Diperbarui_Oleh');
  if (updatedCol >= 0) sh.getRange(rowIndex,updatedCol+1).setValue(now);
  if (byCol >= 0) sh.getRange(rowIndex,byCol+1).setValue('Admin');
  return {ok:true,message:'Data KDKMP berhasil diperbarui.'};
}
function addDocumentation(token, payload) {
  assertAdmin_(token);
  if (!payload || !payload.ID_KDKMP || !payload.Judul) throw new Error('KDKMP dan judul dokumentasi wajib diisi.');
  const known = readObjects_(CFG.MASTER).some(r => r.ID_KDKMP === String(payload.ID_KDKMP));
  if (!known) throw new Error('ID KDKMP tidak ditemukan.');
  const sh = sheet_(CFG.DOCS), id = 'DOC-' + Utilities.getUuid().slice(0,8).toUpperCase();
  sh.appendRow([id,payload.ID_KDKMP,payload.Tanggal || new Date(),payload.Kategori || 'Umum',String(payload.Judul).slice(0,150),String(payload.Keterangan || '').slice(0,1000),String(payload.URL_FOTO || '').slice(0,1500),new Date(),'Admin']);
  return {ok:true,message:'Dokumentasi berhasil ditambahkan.'};
}
function deleteDocumentation(token, id) {
  assertAdmin_(token);
  const sh = sheet_(CFG.DOCS), vals = sh.getDataRange().getDisplayValues();
  for (let i=1;i<vals.length;i++) {
    if (vals[i][0] === String(id)) { sh.deleteRow(i+1); return {ok:true,message:'Dokumentasi dihapus.'}; }
  }
  throw new Error('Dokumentasi tidak ditemukan.');
}
