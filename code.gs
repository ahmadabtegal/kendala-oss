/** MONITORING KENDALA OSS - DPMPTSP KOTA TEGAL
 * Jalankan setupDatabase() sekali dari editor Apps Script.
 * Set Script Property ADMIN_PASSWORD sebelum aplikasi dipublikasikan.
 */
const APP = {
  DB_ID_PROPERTY: 'SPREADSHEET_ID',
  ADMIN_PASSWORD_PROPERTY: 'ADMIN_PASSWORD',
  SHEET_REPORTS: 'Kendala',
  SHEET_HISTORY: 'Riwayat',
  STATUSES: ['Baru', 'Sedang Ditangani', 'Menunggu Pihak OSS', 'Selesai']
};
const REPORT_HEADERS = [
  'ID Tiket','Waktu Laporan','Nama Pelaku Usaha','NIK','No WA',
  'Nama Perusahaan/Nama Usaha','NIB (jika ada)','Jenis Layanan OSS','Kendala',
  'Link Data Dukung','Solusi','Status','Nama Petugas FO','Petugas Penanganan',
  'Waktu Update','Waktu Selesai'
];
const HISTORY_HEADERS = ['Waktu','ID Tiket','Aktor','Status Sebelumnya','Status Baru','Catatan Tindak Lanjut'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Monitoring Kendala OSS | DPMPTSP Kota Tegal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Jalankan sekali. Membuat spreadsheet database dan sheet yang diperlukan. */
function setupDatabase() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(APP.DB_ID_PROPERTY);
  let ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('Database Monitoring Kendala OSS - DPMPTSP Kota Tegal');
    props.setProperty(APP.DB_ID_PROPERTY, ss.getId());
  }
  let report = ss.getSheetByName(APP.SHEET_REPORTS);
  if (!report) report = ss.getSheets()[0].setName(APP.SHEET_REPORTS);
  report.clear();
  report.getRange(1, 1, 1, REPORT_HEADERS.length).setValues([REPORT_HEADERS]);
  report.setFrozenRows(1);
  report.getRange(1,1,1,REPORT_HEADERS.length).setFontWeight('bold').setBackground('#123b63').setFontColor('#ffffff');
  report.autoResizeColumns(1, REPORT_HEADERS.length);
  report.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  report.getRange('O:P').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  let history = ss.getSheetByName(APP.SHEET_HISTORY);
  if (!history) history = ss.insertSheet(APP.SHEET_HISTORY);
  history.clear();
  history.getRange(1,1,1,HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
  history.setFrozenRows(1);
  history.getRange(1,1,1,HISTORY_HEADERS.length).setFontWeight('bold').setBackground('#123b63').setFontColor('#ffffff');
  history.autoResizeColumns(1, HISTORY_HEADERS.length);
  return {ok:true, message:'Database siap digunakan.', spreadsheetUrl:ss.getUrl(), spreadsheetId:ss.getId()};
}

function getDb_() {
  const id = PropertiesService.getScriptProperties().getProperty(APP.DB_ID_PROPERTY);
  if (!id) throw new Error('Database belum disiapkan. Jalankan setupDatabase() dari editor Apps Script.');
  return SpreadsheetApp.openById(id);
}
function getSheet_() { return getDb_().getSheetByName(APP.SHEET_REPORTS); }
function clean_(v, max) { return String(v == null ? '' : v).trim().slice(0, max || 2000); }
function dateText_(d) { return d instanceof Date && !isNaN(d) ? Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss') : ''; }
function readRows_() {
  const sh = getSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,REPORT_HEADERS.length).getValues().map((r,i)=>({
    row:i+2, id:String(r[0]||''), created:dateText_(r[1]), name:String(r[2]||''), nik:String(r[3]||''), phone:String(r[4]||''),
    company:String(r[5]||''), nib:String(r[6]||''), service:String(r[7]||''), issue:String(r[8]||''), evidence:String(r[9]||''),
    solution:String(r[10]||''), status:String(r[11]||'Baru'), fo:String(r[12]||''), handler:String(r[13]||''),
    updated:dateText_(r[14]), completed:dateText_(r[15])
  })).filter(x=>x.id);
}

/** Data untuk daftar publik: semua informasi laporan ditampilkan kecuali NIK. */
function getPublicData() {
  const rows = readRows_().sort((a,b)=>b.row-a.row);
  const stats = {total:rows.length, baru:0, proses:0, menunggu:0, selesai:0};
  rows.forEach(x=>{
    if (x.status === 'Baru') stats.baru++;
    else if (x.status === 'Sedang Ditangani' || x.status === 'Proses Tindak Lanjut') stats.proses++;
    else if (x.status === 'Menunggu Pihak OSS' || x.status === 'Menunggu Admin OSS') stats.menunggu++;
    else if (x.status === 'Selesai') stats.selesai++;
  });
  return {
    ok:true,
    statuses:APP.STATUSES,
    stats:stats,
    reports:rows.map(x=>({
      id:x.id, created:x.created, name:x.name, phone:x.phone, company:x.company,
      nib:x.nib, service:x.service, issue:x.issue, evidence:x.evidence,
      solution:x.solution, status:x.status, fo:x.fo, handler:x.handler,
      updated:x.updated, completed:x.completed
      // NIK sengaja tidak dikirim ke daftar publik.
    }))
  };
}

function submitReport(data) {
  data = data || {};
  const required = ['name','phone','company','service','issue','fo'];
  required.forEach(k=>{ if(!clean_(data[k], k==='issue'?2000:250)) throw new Error('Kolom wajib belum diisi: '+({name:'Nama Pelaku Usaha',phone:'Nomor WhatsApp',company:'Nama Perusahaan/Usaha',service:'Jenis Layanan OSS',issue:'Uraian Kendala',fo:'Nama Petugas FO'}[k])); });
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = getSheet_();
    const now = new Date();
    const ticket = 'OSS-' + Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyyMMdd-HHmmss') + '-' + Math.floor(100 + Math.random()*900);
    const row = [ticket,now,clean_(data.name,150),clean_(data.nik,30),clean_(data.phone,30),clean_(data.company,200),clean_(data.nib,50),clean_(data.service,150),clean_(data.issue,3000),clean_(data.evidence,1000),'','Baru',clean_(data.fo,150),'',now,''];
    sh.appendRow(row);
    addHistory_(ticket, clean_(data.fo,150), '', 'Baru', 'Laporan kendala dibuat');
    return {ok:true,ticket,message:'Kendala berhasil dicatat. Nomor tiket: '+ticket};
  } finally { lock.releaseLock(); }
}

function adminLogin(password) {
  const expected = PropertiesService.getScriptProperties().getProperty(APP.ADMIN_PASSWORD_PROPERTY);
  if (!expected) throw new Error('Password admin belum disetel. Buka Project Settings > Script Properties, lalu tambahkan ADMIN_PASSWORD.');
  if (String(password || '') !== expected) throw new Error('Password admin tidak sesuai.');
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('admin:'+token, '1', 21600); // 6 jam
  return {ok:true, token:token, expiresIn:21600};
}
function requireAdmin_(token) {
  if (!token || CacheService.getScriptCache().get('admin:'+String(token)) !== '1') throw new Error('Sesi admin berakhir. Silakan login kembali.');
}
function getAdminData(token) {
  requireAdmin_(token);
  return {ok:true, statuses:APP.STATUSES, reports:readRows_().sort((a,b)=>b.row-a.row)};
}
function updateReport(token, data) {
  requireAdmin_(token); data=data||{};
  const id=clean_(data.id,100), status=clean_(data.status,80);
  if (!id) throw new Error('ID tiket tidak valid.');
  if (APP.STATUSES.indexOf(status)<0) throw new Error('Status tidak valid.');
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh=getSheet_(), rows=readRows_(), item=rows.find(x=>x.id===id);
    if (!item) throw new Error('Tiket tidak ditemukan.');
    const old=item.status, now=new Date();
    sh.getRange(item.row,12).setValue(status);
    sh.getRange(item.row,11).setValue(clean_(data.solution,3000));
    sh.getRange(item.row,14).setValue(clean_(data.handler,150));
    sh.getRange(item.row,15).setValue(now);
    sh.getRange(item.row,16).setValue(status==='Selesai' ? now : '');
    addHistory_(id,'Admin',old,status,clean_(data.note,1500));
    return {ok:true,message:'Perubahan kendala berhasil disimpan.'};
  } finally {lock.releaseLock();}
}
function addHistory_(id, actor, oldStatus, newStatus, note) {
  const sh=getDb_().getSheetByName(APP.SHEET_HISTORY);
  sh.appendRow([new Date(),id,actor,oldStatus,newStatus,note]);
}
function getPublicHistory(ticket) {
  const sh=getDb_().getSheetByName(APP.SHEET_HISTORY);
  if (!sh || sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,HISTORY_HEADERS.length).getValues()
    .filter(r=>String(r[1])===String(ticket))
    .map(r=>({time:dateText_(r[0]),ticket:String(r[1]),oldStatus:String(r[3]||''),newStatus:String(r[4]||'')}))
    .reverse();
}
function getHistory(token, ticket) {
  requireAdmin_(token);
  const sh=getDb_().getSheetByName(APP.SHEET_HISTORY);
  if (!sh || sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,HISTORY_HEADERS.length).getValues().filter(r=>String(r[1])===String(ticket)).map(r=>({time:dateText_(r[0]),ticket:String(r[1]),actor:String(r[2]),oldStatus:String(r[3]),newStatus:String(r[4]),note:String(r[5])})).reverse();
}

/**
 * API JSON untuk frontend statis yang di-host di GitHub Pages/Vercel.
 * Frontend mengirim Content-Type text/plain agar request sederhana tidak memicu preflight CORS.
 */
function doPost(e) {
  try {
    const body = JSON.parse(e && e.postData && e.postData.contents ? e.postData.contents : '{}');
    const action = String(body.action || '');
    let result;
    switch (action) {
      case 'getPublicData':
        result = getPublicData();
        break;
      case 'submitReport':
        result = submitReport(body.data || {});
        break;
      case 'adminLogin':
        result = adminLogin(body.password || '');
        break;
      case 'getAdminData':
        result = getAdminData(body.token || '');
        break;
      case 'updateReport':
        result = updateReport(body.token || '', body.data || {});
        break;
      case 'getPublicHistory':
        result = getPublicHistory(body.ticket || '');
        break;
      case 'getHistory':
        result = getHistory(body.token || '', body.ticket || '');
        break;
      default:
        throw new Error('Aksi API tidak dikenal.');
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      message: err && err.message ? err.message : 'Terjadi kesalahan pada server.'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
