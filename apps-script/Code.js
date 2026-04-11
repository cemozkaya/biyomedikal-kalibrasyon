/**
 * Biyomedikal Kalibrasyon — Apps Script Backend
 * Container-bound script (clasp create --type sheets ile oluşturulmuş)
 * SpreadsheetApp.getActiveSpreadsheet() bağlı olduğu Sheet'i kullanır.
 */

// =====================================================
// KURULUM — sheetleri ve başlık satırlarını oluşturur
// İlk dağıtımdan sonra editörde bu fonksiyonu bir kez çalıştır
// =====================================================
// Sheet kolonları — sırası önemli, hep aynı sırada okunuyor.
const USER_COLS = ['tc', 'pin', 'ad', 'unvan', 'roller', 'aktif', 'son_giris'];
const REFCIHAZ_COLS = ['id', 'ad', 'tip', 'marka', 'model', 'seri_no',
  'sertifika_no', 'izlenebilirlik', 'gecerlilik_tarihi', 'sicaklik', 'nem'];
const RAPOR_COLS = [
  'raporNo', 'gonderim_tarihi', 'tc', 'ad', 'onay_durumu',
  'dev_type', 'saglik_tesisi', 'marka', 'model', 'seri_no', 'json_veri',
  // Onay/Red/Revizyon/Soft delete audit kolonları
  'onay_notu', 'red_sebebi', 'red_gecmisi', 'revizyon_notu',
  'onaylayan_tc', 'onaylayan_ad', 'onay_tarihi',
  'reddeden_tc',  'reddeden_ad',  'red_tarihi',
  'revizyon_isteyen_tc', 'revizyon_isteyen_ad', 'revizyon_tarihi',
  'tekrar_gonderilme_sayisi',
  'silindi', 'silen_tc', 'silindigi_tarih'
];

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Kullanicilar sayfası
  let uSheet = ss.getSheetByName('Kullanicilar');
  if (!uSheet) {
    uSheet = ss.insertSheet('Kullanicilar');
  }
  if (uSheet.getLastRow() === 0) {
    uSheet.appendRow(USER_COLS);
    uSheet.getRange('A:B').setNumberFormat('@'); // Düz metin (sıfırlar kaybolmasın)
    uSheet.setFrozenRows(1);
    uSheet.getRange(1, 1, 1, USER_COLS.length).setFontWeight('bold').setBackground('#1A3A5C').setFontColor('#ffffff');
  } else {
    // Eski kurulumdan kalan eksik kolonları otomatik tamamla
    upgradeSheetColumns_(uSheet, USER_COLS);
  }

  // RefCihazlar sayfası
  let rcSheet = ss.getSheetByName('RefCihazlar');
  if (!rcSheet) {
    rcSheet = ss.insertSheet('RefCihazlar');
  }
  if (rcSheet.getLastRow() === 0) {
    rcSheet.appendRow(REFCIHAZ_COLS);
    rcSheet.getRange('A:A').setNumberFormat('@'); // id düz metin
    rcSheet.setFrozenRows(1);
    rcSheet.getRange(1, 1, 1, REFCIHAZ_COLS.length).setFontWeight('bold').setBackground('#1A3A5C').setFontColor('#ffffff');
  } else {
    upgradeSheetColumns_(rcSheet, REFCIHAZ_COLS);
  }

  // Raporlar sayfası
  let rSheet = ss.getSheetByName('Raporlar');
  if (!rSheet) {
    rSheet = ss.insertSheet('Raporlar');
  }
  if (rSheet.getLastRow() === 0) {
    rSheet.appendRow(RAPOR_COLS);
    rSheet.getRange('A:A').setNumberFormat('@'); // raporNo düz metin
    rSheet.getRange('C:C').setNumberFormat('@'); // tc düz metin
    rSheet.setFrozenRows(1);
    rSheet.getRange(1, 1, 1, RAPOR_COLS.length).setFontWeight('bold').setBackground('#C8102E').setFontColor('#ffffff');
  } else {
    upgradeSheetColumns_(rSheet, RAPOR_COLS);
  }

  // Varsayılan "Sayfa1" / "Sheet1" sayfasını sil (içi boşsa)
  const def = ss.getSheetByName('Sayfa1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(def);
  }

  return 'Setup tamam. Şimdi Kullanicilar sayfasını açıp bir admin satırı ekleyin.';
}

// Mevcut bir sayfanın ilk satırındaki kolonları kontrol eder, eksik
// olanları sona ekler. İlk satırı bozmaz, sadece eklemeyi yapar.
function upgradeSheetColumns_(sheet, cols) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.appendRow(cols);
    return;
  }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = cols.filter(c => headers.indexOf(c) === -1);
  if (missing.length) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, 1, 1, lastCol + missing.length)
         .setFontWeight('bold');
  }
}

// Sheet içindeki bir kolonun index'ini döndürür (1-based). Yoksa -1.
function colIdx_(sheet, colName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const i = headers.indexOf(colName);
  return i === -1 ? -1 : i + 1;
}

// Bir sheet satırını cols sırasına göre objeye çevir
function rowToObj_(headers, row) {
  const o = {};
  for (let i = 0; i < headers.length; i++) o[headers[i]] = row[i];
  return o;
}

// =====================================================
// HTTP GİRİŞ
// =====================================================
function doPost(e) {
  try {
    // Güvenlik ağı: sayfalar yoksa otomatik kur
    ensureSheets_();

    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    // Auth & rapor temel
    if (action === 'login')         return jsonOut(handleLogin(body));
    if (action === 'submitRapor')   return jsonOut(handleSubmitRapor(body));
    if (action === 'saveDraft')     return jsonOut(handleSaveDraft(body));
    if (action === 'listRaporlar')  return jsonOut(handleListRaporlar(body));
    if (action === 'peekRaporNo')   return jsonOut(handlePeekRaporNo(body));
    if (action === 'listManagers')  return jsonOut(handleListManagers(body));

    // Rapor onay/red/revizyon/resubmit/sil
    if (action === 'approveRapor')  return jsonOut(handleApproveRapor(body));
    if (action === 'rejectRapor')   return jsonOut(handleRejectRapor(body));
    if (action === 'revisionRapor') return jsonOut(handleRevisionRapor(body));
    if (action === 'resubmitRapor') return jsonOut(handleResubmitRapor(body));
    if (action === 'deleteRapor')   return jsonOut(handleDeleteRapor(body));

    // Kullanıcı yönetimi
    if (action === 'listUsers')     return jsonOut(handleListUsers(body));
    if (action === 'addUser')       return jsonOut(handleAddUser(body));
    if (action === 'updateUser')    return jsonOut(handleUpdateUser(body));
    if (action === 'deleteUser')    return jsonOut(handleDeleteUser(body));
    if (action === 'resetUserPin')  return jsonOut(handleResetUserPin(body));
    if (action === 'changeOwnPin')  return jsonOut(handleChangeOwnPin(body));

    // Referans cihazlar
    if (action === 'listRefCihazlar')  return jsonOut(handleListRefCihazlar(body));
    if (action === 'saveRefCihaz')     return jsonOut(handleSaveRefCihaz(body));
    if (action === 'deleteRefCihaz')   return jsonOut(handleDeleteRefCihaz(body));

    return jsonOut({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    ensureSheets_();
    return ContentService
      .createTextOutput('Biyomedikal Kalibrasyon API çalışıyor. POST gerekiyor.')
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput('Setup hatası: ' + err)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Kullanicilar') || !ss.getSheetByName('Raporlar')) {
    setupSheets();
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

// =====================================================
// AUTH
// =====================================================
function normalizeRole_(s) {
  return String(s)
    .toLocaleLowerCase('tr-TR')
    .replace(/ö/g, 'o').replace(/ü/g, 'u').replace(/ç/g, 'c')
    .replace(/ı/g, 'i').replace(/ş/g, 's').replace(/ğ/g, 'g')
    .trim();
}

// Kullanıcıyı bulur. opts.checkAktif true ise pasif kullanıcılar
// reddedilir. opts.updateLogin true ise son_giris güncellenir.
function findUser_(tc, pin, opts) {
  opts = opts || {};
  const sheet = getSheet_('Kullanicilar');
  if (!sheet) return null;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const data = sheet.getDataRange().getValues();
  const tcCol     = headers.indexOf('tc');
  const pinCol    = headers.indexOf('pin');
  const adCol     = headers.indexOf('ad');
  const unvanCol  = headers.indexOf('unvan');
  const rollerCol = headers.indexOf('roller');
  const aktifCol  = headers.indexOf('aktif');
  const sgCol     = headers.indexOf('son_giris');

  for (let i = 1; i < data.length; i++) {
    const rowTc  = String(data[i][tcCol]).trim();
    const rowPin = String(data[i][pinCol]).trim();
    if (rowTc === String(tc).trim() && rowPin === String(pin).trim()) {
      const aktifVal = aktifCol !== -1 ? data[i][aktifCol] : true;
      const aktif = (aktifVal === '' || aktifVal === undefined || aktifVal === null)
        ? true
        : (aktifVal === true || String(aktifVal).toLowerCase() === 'true' || aktifVal === 1);
      if (opts.checkAktif && !aktif) {
        return { __pasif: true };
      }
      // Önce eski son_giris'i oku (welcome ekranı için), sonra üzerine yaz
      const eskiSonGiris = (sgCol !== -1 && data[i][sgCol]) ? String(data[i][sgCol]) : '';
      if (opts.updateLogin && sgCol !== -1) {
        const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        sheet.getRange(i + 1, sgCol + 1).setValue(ts);
      }
      return {
        tc:     rowTc,
        ad:     String(data[i][adCol]    || '').trim(),
        unvan:  String(data[i][unvanCol] || '').trim(),
        roller: String(data[i][rollerCol] || '')
                  .split(',')
                  .map(normalizeRole_)
                  .filter(Boolean),
        aktif:  aktif,
        oncekiGiris: eskiSonGiris
      };
    }
  }
  return null;
}

function handleLogin(body) {
  const user = findUser_(body.tc, body.pin, { checkAktif: true, updateLogin: true });
  if (!user) return { ok: false, error: 'invalid_credentials' };
  if (user.__pasif) return { ok: false, error: 'user_inactive' };
  return { ok: true, user: user };
}

// =====================================================
// RAPOR
// =====================================================
function handleSubmitRapor(body) {
  const user = findUser_(body.tc, body.pin);
  if (!user) return { ok: false, error: 'unauthorized' };

  const r = body.rapor || {};

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_('Raporlar');
    if (!sheet) throw new Error('Raporlar sayfası bulunamadı');

    // Mevcut taslak mı? (saveDraft ile sheet'e zaten yazılmışsa, aynı
    // satırı update et — yeni satır eklemiyoruz, no koruyoruz)
    if (r.raporNo) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
      const data    = sheet.getDataRange().getValues();
      const rnoCol  = headers.indexOf('raporNo');
      const tcCol   = headers.indexOf('tc');
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][rnoCol]) === String(r.raporNo)) {
          // Sahiplik kontrolü — sadece raporu açan kişi gönderebilir
          if (String(data[i][tcCol]) !== String(user.tc)) {
            return { ok: false, error: 'forbidden' };
          }
          // Update — taslaktan pending'e
          r.onayDurumu = 'pending';
          const updates = {
            'onay_durumu':   'pending',
            'dev_type':      r.devType || '',
            'saglik_tesisi': r.saglikTesisi || '',
            'marka':         r.marka || '',
            'model':         r.model || '',
            'seri_no':       r.seriNo || '',
            'json_veri':     JSON.stringify(r)
          };
          Object.keys(updates).forEach(key => {
            const c = headers.indexOf(key);
            if (c !== -1) sheet.getRange(i + 1, c + 1).setValue(updates[key]);
          });
          return { ok: true, raporNo: r.raporNo, action: 'updated' };
        }
      }
    }

    // Yeni rapor — yeni no üret, satır ekle
    const finalRaporNo = generateRaporNo_(sheet, r.devType);
    r.raporNo    = finalRaporNo;
    r.onayDurumu = 'pending';

    sheet.appendRow([
      finalRaporNo,
      new Date(),
      user.tc,
      user.ad,
      'pending',
      r.devType      || '',
      r.saglikTesisi || '',
      r.marka        || '',
      r.model        || '',
      r.seriNo       || '',
      JSON.stringify(r)
    ]);

    return { ok: true, raporNo: finalRaporNo, action: 'added' };
  } finally {
    lock.releaseLock();
  }
}

// Taslak kaydet — yeni rapor için yeni satır + yeni rapor no üretir,
// mevcut taslak için satırı günceller. onay_durumu boş bırakılır
// (frontend'de null = taslak). Sahibi olmayan başkasının taslağını
// kaydedemez.
function handleSaveDraft(body) {
  const user = findUser_(body.tc, body.pin);
  if (!user) return { ok: false, error: 'unauthorized' };

  const r = body.rapor || {};

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_('Raporlar');
    if (!sheet) throw new Error('Raporlar sayfası bulunamadı');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const data    = sheet.getDataRange().getValues();
    const rnoCol  = headers.indexOf('raporNo');

    // Mevcut bir rapor mu? raporNo varsa ve sheet'te bulunuyorsa update.
    let existingRow = -1;
    if (r.raporNo) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][rnoCol]) === String(r.raporNo)) {
          existingRow = i + 1;
          // Sahiplik kontrolü — sadece raporu yapan teknisyen veya
          // müdür/yönetici taslağı güncelleyebilir
          const tcCol = headers.indexOf('tc');
          const ownerTc = String(data[i][tcCol] || '');
          const role = (user.roller && user.roller[0]) || '';
          if (ownerTc !== String(user.tc) && role !== 'yonetici' && role !== 'mudur') {
            return { ok: false, error: 'forbidden' };
          }
          break;
        }
      }
    }

    if (existingRow !== -1) {
      // Update — sadece değişebilen kolonları yaz
      const updates = {
        'onay_durumu':   '',  // boş = taslak
        'dev_type':      r.devType || '',
        'saglik_tesisi': r.saglikTesisi || '',
        'marka':         r.marka || '',
        'model':         r.model || '',
        'seri_no':       r.seriNo || '',
        'json_veri':     JSON.stringify(r)
      };
      Object.keys(updates).forEach(key => {
        const c = headers.indexOf(key);
        if (c !== -1) sheet.getRange(existingRow, c + 1).setValue(updates[key]);
      });
      return { ok: true, raporNo: r.raporNo, action: 'updated' };
    }

    // Yeni taslak — yeni rapor no üret, satır ekle
    const finalRaporNo = generateRaporNo_(sheet, r.devType);
    r.raporNo    = finalRaporNo;
    r.onayDurumu = null;

    sheet.appendRow([
      finalRaporNo,
      new Date(),
      user.tc,
      user.ad,
      '',  // onay_durumu boş = taslak
      r.devType      || '',
      r.saglikTesisi || '',
      r.marka        || '',
      r.model        || '',
      r.seriNo       || '',
      JSON.stringify(r)
    ]);

    return { ok: true, raporNo: finalRaporNo, action: 'added' };
  } finally {
    lock.releaseLock();
  }
}

function generateRaporNo_(sheet, devType) {
  const prefixMap = {
    defibrillator: 'DEF',
    aspirator:     'ASP',
    infusion:      'INF',
    monitor:       'MON'
  };
  const prefix = prefixMap[devType] || 'XXX';

  const now = new Date();
  const dd  = pad2_(now.getDate());
  const mm  = pad2_(now.getMonth() + 1);
  const yy  = String(now.getFullYear()).slice(-2);
  const datePart = dd + mm + yy;
  const base     = prefix + datePart;  // ör. DEF110426

  // Max-based sıra: o gündeki mevcut raporlar arasındaki en büyük
  // numara + 1. Count-based yaklaşım, satır silinince (row delete)
  // numara çakışmasına yol açıyordu — max-based silmeden sonra bile
  // güvenli çalışır, kaybedilen numaralar boşluk olarak kalır.
  const data = sheet.getDataRange().getValues();
  let maxSeq = 0;
  for (let i = 1; i < data.length; i++) {
    const rno = String(data[i][0] || '');
    if (rno.indexOf(base) === 0) {
      const seq = parseInt(rno.slice(base.length), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }

  return base + pad3_(maxSeq + 1);
}

function pad2_(n) { return String(n).length < 2 ? '0' + n : String(n); }
function pad3_(n) {
  const s = String(n);
  if (s.length >= 3) return s;
  return ('000' + s).slice(-3);
}

// Bir sonraki raporNo'yu DÖNDÜRÜR ama kayıt yapmaz.
// Form açıldığında frontend'de göstermek için kullanılır.
// NOT: Sadece önizleme; gerçek raporNo submitRapor anında Sheets kilitli
// haldeyken yeniden hesaplanır (yarış durumu yaşanmaz).
function handlePeekRaporNo(body) {
  const user = findUser_(body.tc, body.pin);
  if (!user) return { ok: false, error: 'unauthorized' };

  const sheet = getSheet_('Raporlar');
  if (!sheet) return { ok: false, error: 'no_sheet' };

  const raporNo = generateRaporNo_(sheet, body.devType);
  return { ok: true, raporNo: raporNo };
}

function handleListRaporlar(body) {
  const user = findUser_(body.tc, body.pin);
  if (!user) return { ok: false, error: 'unauthorized' };

  const sheet = getSheet_('Raporlar');
  if (!sheet) return { ok: true, raporlar: [] };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const data    = sheet.getDataRange().getValues();
  const rows    = [];

  // Teknisyen rolü → sadece kendi raporları
  // Müdür/Yönetici → tüm raporları görür (pgVeriler arşivi için).
  // Sorumlu müdür filtresi pgOnay frontend'inde uygulanır — backend
  // bu filtreyi yapmaz, çünkü pgVeriler'de denetim amacıyla tüm
  // raporlar görünmeli.
  const role = (user.roller && user.roller[0]) || '';
  const onlyMine = role === 'teknisyen';

  for (let i = 1; i < data.length; i++) {
    const o = rowToObj_(headers, data[i]);
    // Soft delete filtresi
    const silindi = o.silindi === true || String(o.silindi).toLowerCase() === 'true';
    if (silindi) continue;
    // Teknisyen kendi raporları
    if (onlyMine && String(o.tc) !== String(user.tc)) continue;

    // JSON veriyi parse et — frontend tüm rapor objesini ister
    let parsed = {};
    try { if (o.json_veri) parsed = JSON.parse(o.json_veri); } catch (_) {}

    // Audit alanlarını parsed üstüne yaz (sheet üzerindekiler son söz)
    const merged = Object.assign({}, parsed, {
      raporNo:      String(o.raporNo),
      raporTarihi:  parsed.raporTarihi || (o.gonderim_tarihi ? Utilities.formatDate(new Date(o.gonderim_tarihi), Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''),
      onayDurumu:   o.onay_durumu === 'null' ? null : (o.onay_durumu || null),
      devType:      String(o.dev_type || parsed.devType || ''),
      saglikTesisi: String(o.saglik_tesisi || parsed.saglikTesisi || ''),
      marka:        String(o.marka || parsed.marka || ''),
      model:        String(o.model || parsed.model || ''),
      seriNo:       String(o.seri_no || parsed.seriNo || ''),
      testiUygulayan: String(o.ad || parsed.testiUygulayan || ''),
      // Audit alanları
      onayNotu:           o.onay_notu       || '',
      redSebebi:          o.red_sebebi      || '',
      redGecmisi:         (function () { try { return o.red_gecmisi ? JSON.parse(o.red_gecmisi) : []; } catch (_) { return []; } })(),
      revizyonNotu:       o.revizyon_notu   || '',
      onaylayanTc:        o.onaylayan_tc    || '',
      onaylayanAd:        o.onaylayan_ad    || '',
      onayTarihi:         o.onay_tarihi     || '',
      reddedenTc:         o.reddeden_tc     || '',
      reddedenAd:         o.reddeden_ad     || '',
      redTarihi:          o.red_tarihi      || '',
      revizyonIsteyenTc:  o.revizyon_isteyen_tc || '',
      revizyonIsteyenAd:  o.revizyon_isteyen_ad || '',
      revizyonTarihi:     o.revizyon_tarihi || '',
      tekrarGonderilmeSayisi: parseInt(o.tekrar_gonderilme_sayisi || 0, 10) || 0,
      // Sorumlu müdür (atama) — json_veri'den geliyor, parsed üstüne yazma
      sorumluMudurTc:     parsed.sorumluMudurTc    || '',
      sorumluMudurAd:     parsed.sorumluMudurAd    || '',
      sorumluMudurUnvan:  parsed.sorumluMudurUnvan || ''
    });
    rows.push(merged);
  }
  return { ok: true, raporlar: rows };
}

// =====================================================
// RAPOR ONAY / RED / REVİZYON / RESUBMIT / SİLME
// =====================================================

// Yardımcı: belirli raporNo'lu satırı bul, kolon güncellemelerini uygula.
// updates: { col_name: value, ... }
function updateRaporRow_(raporNo, updates) {
  const sheet = getSheet_('Raporlar');
  if (!sheet) throw new Error('Raporlar sheet yok');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const data    = sheet.getDataRange().getValues();
  const tcCol   = headers.indexOf('raporNo');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][tcCol]) === String(raporNo)) {
      const rowNum = i + 1;
      Object.keys(updates).forEach(key => {
        const c = headers.indexOf(key);
        if (c !== -1) sheet.getRange(rowNum, c + 1).setValue(updates[key]);
      });
      return true;
    }
  }
  return false;
}

function getRaporObj_(raporNo) {
  const sheet = getSheet_('Raporlar');
  if (!sheet) return null;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const data    = sheet.getDataRange().getValues();
  const rnoCol  = headers.indexOf('raporNo');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][rnoCol]) === String(raporNo)) {
      return rowToObj_(headers, data[i]);
    }
  }
  return null;
}

// Sadece müdür/yönetici onay aksiyonu yapabilir
function requireAdmin_(body) {
  const user = findUser_(body.tc, body.pin);
  if (!user) return { error: 'unauthorized' };
  const role = (user.roller && user.roller[0]) || '';
  if (role !== 'yonetici' && role !== 'mudur') return { error: 'forbidden' };
  return { user: user };
}

function handleApproveRapor(body) {
  const auth = requireAdmin_(body);
  if (auth.error) return { ok: false, error: auth.error };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const ok = updateRaporRow_(body.raporNo, {
      onay_durumu:  'approved',
      onay_notu:    body.note || '',
      onaylayan_tc: auth.user.tc,
      onaylayan_ad: auth.user.ad,
      onay_tarihi:  today
    });
    return ok ? { ok: true } : { ok: false, error: 'rapor_not_found' };
  } finally {
    lock.releaseLock();
  }
}

function handleRejectRapor(body) {
  const auth = requireAdmin_(body);
  if (auth.error) return { ok: false, error: auth.error };
  if (!body.note) return { ok: false, error: 'red_sebebi_zorunlu' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Eğer rapor zaten reddedilmişse, eski sebebi red_gecmisi'ne push
    const existing = getRaporObj_(body.raporNo);
    let gecmis = [];
    if (existing) {
      try { gecmis = existing.red_gecmisi ? JSON.parse(existing.red_gecmisi) : []; } catch (_) {}
      if (existing.red_sebebi) {
        gecmis.push({
          sebep:    existing.red_sebebi,
          reddeden: existing.reddeden_ad || '',
          tarih:    existing.red_tarihi || ''
        });
      }
    }
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const ok = updateRaporRow_(body.raporNo, {
      onay_durumu: 'rejected',
      red_sebebi:  body.note,
      red_gecmisi: JSON.stringify(gecmis),
      reddeden_tc: auth.user.tc,
      reddeden_ad: auth.user.ad,
      red_tarihi:  today
    });
    return ok ? { ok: true } : { ok: false, error: 'rapor_not_found' };
  } finally {
    lock.releaseLock();
  }
}

function handleRevisionRapor(body) {
  const auth = requireAdmin_(body);
  if (auth.error) return { ok: false, error: auth.error };
  if (!body.note) return { ok: false, error: 'revizyon_notu_zorunlu' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const ok = updateRaporRow_(body.raporNo, {
      onay_durumu:           'revision',
      revizyon_notu:         body.note,
      revizyon_isteyen_tc:   auth.user.tc,
      revizyon_isteyen_ad:   auth.user.ad,
      revizyon_tarihi:       today
    });
    return ok ? { ok: true } : { ok: false, error: 'rapor_not_found' };
  } finally {
    lock.releaseLock();
  }
}

// Teknisyen reddedilen veya revizyon istenen raporu düzeltip tekrar gönderir.
// Aynı rapor no korunur, redSebebi/revizyonNotu redGecmisi'ne push edilir,
// yeni rapor verisi (json_veri) ile satır güncellenir.
function handleResubmitRapor(body) {
  const user = findUser_(body.tc, body.pin);
  if (!user) return { ok: false, error: 'unauthorized' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existing = getRaporObj_(body.raporNo);
    if (!existing) return { ok: false, error: 'rapor_not_found' };

    // Sahibi olmayan resubmit yapamaz
    if (String(existing.tc) !== String(user.tc)) return { ok: false, error: 'forbidden' };

    // Sadece rejected veya revision raporlar resubmit edilebilir
    const dur = String(existing.onay_durumu || '');
    if (dur !== 'rejected' && dur !== 'revision') {
      return { ok: false, error: 'invalid_state' };
    }

    // Geçmişe push
    let gecmis = [];
    try { gecmis = existing.red_gecmisi ? JSON.parse(existing.red_gecmisi) : []; } catch (_) {}
    if (existing.red_sebebi) {
      gecmis.push({
        sebep:    existing.red_sebebi,
        reddeden: existing.reddeden_ad || '',
        tarih:    existing.red_tarihi || ''
      });
    }
    if (existing.revizyon_notu) {
      gecmis.push({
        sebep:    '[Revizyon] ' + existing.revizyon_notu,
        reddeden: existing.revizyon_isteyen_ad || '',
        tarih:    existing.revizyon_tarihi || ''
      });
    }
    const sayi = parseInt(existing.tekrar_gonderilme_sayisi || 0, 10) + 1;

    const updates = {
      onay_durumu:               'pending',
      red_sebebi:                '',
      revizyon_notu:             '',
      red_gecmisi:               JSON.stringify(gecmis),
      tekrar_gonderilme_sayisi:  sayi
    };
    // Yeni rapor verisini kaydet
    if (body.rapor) {
      updates.json_veri      = JSON.stringify(body.rapor);
      if (body.rapor.saglikTesisi) updates.saglik_tesisi = body.rapor.saglikTesisi;
      if (body.rapor.marka)        updates.marka         = body.rapor.marka;
      if (body.rapor.model)        updates.model         = body.rapor.model;
      if (body.rapor.seriNo)       updates.seri_no       = body.rapor.seriNo;
    }
    const ok = updateRaporRow_(body.raporNo, updates);
    return ok ? { ok: true } : { ok: false, error: 'rapor_not_found' };
  } finally {
    lock.releaseLock();
  }
}

// Soft delete — silinen rapor satırı kalır, sadece silindi=true işaretlenir.
// Numara çakışmasını ve audit kaybını önler.
function handleDeleteRapor(body) {
  const auth = requireAdmin_(body);
  if (auth.error) return { ok: false, error: auth.error };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const ok = updateRaporRow_(body.raporNo, {
      silindi:           true,
      silen_tc:          auth.user.tc,
      silindigi_tarih:   today
    });
    return ok ? { ok: true } : { ok: false, error: 'rapor_not_found' };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================
// KULLANICI YÖNETİMİ (sadece yönetici)
// =====================================================

function requireYonetici_(body) {
  const user = findUser_(body.tc, body.pin);
  if (!user) return { error: 'unauthorized' };
  const roller = user.roller || [];
  if (roller.indexOf('yonetici') === -1) return { error: 'forbidden' };
  return { user: user };
}

// Sorumlu müdür dropdown'u için herkesin çekebildiği endpoint.
// Sadece müdür+yönetici rolündeki AKTİF kullanıcıların tc/ad/unvan'ını döner.
// PIN/roller/sonGiris gibi hassas alanlar dahil değildir.
function handleListManagers(body) {
  const user = findUser_(body.tc, body.pin);
  if (!user) return { ok: false, error: 'unauthorized' };

  const sheet = getSheet_('Kullanicilar');
  if (!sheet) return { ok: true, managers: [] };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const data = sheet.getDataRange().getValues();
  const managers = [];
  for (let i = 1; i < data.length; i++) {
    const o = rowToObj_(headers, data[i]);
    if (!o.tc) continue;
    const aktifRaw = o.aktif;
    const aktif = (aktifRaw === '' || aktifRaw === undefined || aktifRaw === null)
      ? true
      : (aktifRaw === true || String(aktifRaw).toLowerCase() === 'true' || aktifRaw === 1);
    if (!aktif) continue;
    const roller = String(o.roller || '').split(',').map(normalizeRole_).filter(Boolean);
    if (roller.indexOf('mudur') === -1 && roller.indexOf('yonetici') === -1) continue;
    managers.push({
      tc:    String(o.tc),
      ad:    String(o.ad || ''),
      unvan: String(o.unvan || '')
    });
  }
  return { ok: true, managers: managers };
}

function handleListUsers(body) {
  const auth = requireYonetici_(body);
  if (auth.error) return { ok: false, error: auth.error };

  const sheet = getSheet_('Kullanicilar');
  if (!sheet) return { ok: true, users: [] };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const data = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const o = rowToObj_(headers, data[i]);
    if (!o.tc) continue;
    const aktifRaw = o.aktif;
    const aktif = (aktifRaw === '' || aktifRaw === undefined || aktifRaw === null)
      ? true
      : (aktifRaw === true || String(aktifRaw).toLowerCase() === 'true' || aktifRaw === 1);
    users.push({
      tc:       String(o.tc),
      ad:       String(o.ad || ''),
      unvan:    String(o.unvan || ''),
      roller:   String(o.roller || '').split(',').map(normalizeRole_).filter(Boolean),
      aktif:    aktif,
      sonGiris: o.son_giris ? String(o.son_giris) : ''
    });
  }
  return { ok: true, users: users };
}

function handleAddUser(body) {
  const auth = requireYonetici_(body);
  if (auth.error) return { ok: false, error: auth.error };
  const u = body.user || {};
  if (!u.tc || !u.pin || !u.ad) return { ok: false, error: 'eksik_alan' };
  if (!/^\d{11}$/.test(String(u.tc))) return { ok: false, error: 'tc_format' };
  if (!/^\d{4}$/.test(String(u.pin))) return { ok: false, error: 'pin_format' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_('Kullanicilar');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const data = sheet.getDataRange().getValues();
    const tcCol = headers.indexOf('tc');
    const adCol = headers.indexOf('ad');
    const pinCol = headers.indexOf('pin');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][tcCol]) === String(u.tc)) {
        // İdempotency: Apps Script Web App POST yanıtı bazen 302 redirect
        // ile follow ediliyor ve fetch POST'u tekrar gönderiyor. Aynı veri
        // ile ikinci çağrı için "zaten var" hatası yerine başarılı say —
        // ama farklı veri ise gerçek bir çakışma var, hata döndür.
        const sameAd  = String(data[i][adCol] || '') === String(u.ad);
        const samePin = String(data[i][pinCol] || '') === String(u.pin);
        if (sameAd && samePin) {
          return { ok: true, alreadyExists: true };
        }
        return { ok: false, error: 'tc_zaten_var' };
      }
    }
    // Satırı USER_COLS sırasına göre oluştur
    const row = USER_COLS.map(col => {
      if (col === 'roller') return (u.roller || []).join(',');
      if (col === 'aktif')  return u.aktif !== false;
      if (col === 'son_giris') return '';
      return u[col] != null ? u[col] : '';
    });
    sheet.appendRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function handleUpdateUser(body) {
  const auth = requireYonetici_(body);
  if (auth.error) return { ok: false, error: auth.error };
  const u = body.user || {};
  if (!u.tc) return { ok: false, error: 'tc_eksik' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_('Kullanicilar');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const data = sheet.getDataRange().getValues();
    const tcCol = headers.indexOf('tc');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][tcCol]) === String(u.tc)) {
        const rowNum = i + 1;
        if (u.ad    !== undefined) sheet.getRange(rowNum, headers.indexOf('ad') + 1).setValue(u.ad);
        if (u.unvan !== undefined) sheet.getRange(rowNum, headers.indexOf('unvan') + 1).setValue(u.unvan);
        if (u.roller) sheet.getRange(rowNum, headers.indexOf('roller') + 1).setValue((u.roller || []).join(','));
        if (u.aktif !== undefined) sheet.getRange(rowNum, headers.indexOf('aktif') + 1).setValue(u.aktif === true);
        if (u.pin) {
          if (!/^\d{4}$/.test(String(u.pin))) return { ok: false, error: 'pin_format' };
          sheet.getRange(rowNum, headers.indexOf('pin') + 1).setValue(u.pin);
        }
        return { ok: true };
      }
    }
    return { ok: false, error: 'user_not_found' };
  } finally {
    lock.releaseLock();
  }
}

function handleDeleteUser(body) {
  const auth = requireYonetici_(body);
  if (auth.error) return { ok: false, error: auth.error };
  if (!body.targetTc) return { ok: false, error: 'tc_eksik' };
  // Kendini silemez
  if (String(body.targetTc) === String(auth.user.tc)) return { ok: false, error: 'cannot_delete_self' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_('Kullanicilar');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const data = sheet.getDataRange().getValues();
    const tcCol = headers.indexOf('tc');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][tcCol]) === String(body.targetTc)) {
        sheet.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: 'user_not_found' };
  } finally {
    lock.releaseLock();
  }
}

function handleResetUserPin(body) {
  const auth = requireYonetici_(body);
  if (auth.error) return { ok: false, error: auth.error };
  if (!body.targetTc || !body.newPin) return { ok: false, error: 'eksik_alan' };
  if (!/^\d{4}$/.test(String(body.newPin))) return { ok: false, error: 'pin_format' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_('Kullanicilar');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const data = sheet.getDataRange().getValues();
    const tcCol = headers.indexOf('tc');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][tcCol]) === String(body.targetTc)) {
        sheet.getRange(i + 1, headers.indexOf('pin') + 1).setValue(body.newPin);
        return { ok: true };
      }
    }
    return { ok: false, error: 'user_not_found' };
  } finally {
    lock.releaseLock();
  }
}

// Kullanıcı kendi PIN'ini değiştirir. body.tc + body.pin (eski) doğrulanır,
// body.newPin yeni 4 haneli PIN olarak yazılır.
function handleChangeOwnPin(body) {
  const user = findUser_(body.tc, body.pin);
  if (!user) return { ok: false, error: 'unauthorized' };
  if (!body.newPin) return { ok: false, error: 'eksik_alan' };
  if (!/^\d{4}$/.test(String(body.newPin))) return { ok: false, error: 'pin_format' };
  if (String(body.newPin) === String(body.pin)) return { ok: false, error: 'ayni_pin' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_('Kullanicilar');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const data = sheet.getDataRange().getValues();
    const tcCol = headers.indexOf('tc');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][tcCol]) === String(body.tc)) {
        sheet.getRange(i + 1, headers.indexOf('pin') + 1).setValue(body.newPin);
        return { ok: true };
      }
    }
    return { ok: false, error: 'user_not_found' };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================
// REFERANS CİHAZLAR
// =====================================================

// Sadece müdür/yönetici düzenleyebilir, herkes okuyabilir.
function requireAdminOrSame_(body) {
  const user = findUser_(body.tc, body.pin);
  if (!user) return { error: 'unauthorized' };
  return { user: user };
}

function requireRefCihazAdmin_(body) {
  const user = findUser_(body.tc, body.pin);
  if (!user) return { error: 'unauthorized' };
  const role = (user.roller && user.roller[0]) || '';
  if (role !== 'mudur' && role !== 'yonetici') return { error: 'forbidden' };
  return { user: user };
}

function handleListRefCihazlar(body) {
  // Herkes çağırabilir — kalibrasyon formunda teknisyen de görmeli
  const user = findUser_(body.tc, body.pin);
  if (!user) return { ok: false, error: 'unauthorized' };

  const sheet = getSheet_('RefCihazlar');
  if (!sheet) return { ok: true, cihazlar: [] };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const data = sheet.getDataRange().getValues();
  const cihazlar = [];
  for (let i = 1; i < data.length; i++) {
    const o = rowToObj_(headers, data[i]);
    if (!o.id) continue;
    cihazlar.push({
      id:               String(o.id),
      ad:               String(o.ad || ''),
      tip:              String(o.tip || 'diger'),
      marka:            String(o.marka || ''),
      model:            String(o.model || ''),
      seriNo:           String(o.seri_no || ''),
      sertifikaNo:      String(o.sertifika_no || ''),
      izlenebilirlik:   String(o.izlenebilirlik || ''),
      gecerlilikTarihi: o.gecerlilik_tarihi
                          ? Utilities.formatDate(new Date(o.gecerlilik_tarihi), Session.getScriptTimeZone(), 'yyyy-MM-dd')
                          : '',
      sicaklik:         String(o.sicaklik || ''),
      nem:              String(o.nem || '')
    });
  }
  return { ok: true, cihazlar: cihazlar };
}

function handleSaveRefCihaz(body) {
  const auth = requireRefCihazAdmin_(body);
  if (auth.error) return { ok: false, error: auth.error };
  const c = body.cihaz || {};
  if (!c.id || !c.ad) return { ok: false, error: 'eksik_alan' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_('RefCihazlar');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const data = sheet.getDataRange().getValues();
    const idCol = headers.indexOf('id');
    // Mevcut satırı bul
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(c.id)) {
        // Update
        const rowNum = i + 1;
        const updates = {
          'ad':                c.ad || '',
          'tip':               c.tip || 'diger',
          'marka':             c.marka || '',
          'model':             c.model || '',
          'seri_no':           c.seriNo || '',
          'sertifika_no':      c.sertifikaNo || '',
          'izlenebilirlik':    c.izlenebilirlik || '',
          'gecerlilik_tarihi': c.gecerlilikTarihi || '',
          'sicaklik':          c.sicaklik || '',
          'nem':               c.nem || ''
        };
        Object.keys(updates).forEach(key => {
          const col = headers.indexOf(key);
          if (col !== -1) sheet.getRange(rowNum, col + 1).setValue(updates[key]);
        });
        return { ok: true, action: 'updated' };
      }
    }
    // Yeni kayıt — REFCIHAZ_COLS sırasına göre satır oluştur
    const row = REFCIHAZ_COLS.map(col => {
      switch (col) {
        case 'id':                return c.id;
        case 'ad':                return c.ad || '';
        case 'tip':               return c.tip || 'diger';
        case 'marka':             return c.marka || '';
        case 'model':             return c.model || '';
        case 'seri_no':           return c.seriNo || '';
        case 'sertifika_no':      return c.sertifikaNo || '';
        case 'izlenebilirlik':    return c.izlenebilirlik || '';
        case 'gecerlilik_tarihi': return c.gecerlilikTarihi || '';
        case 'sicaklik':          return c.sicaklik || '';
        case 'nem':               return c.nem || '';
        default: return '';
      }
    });
    sheet.appendRow(row);
    return { ok: true, action: 'added' };
  } finally {
    lock.releaseLock();
  }
}

function handleDeleteRefCihaz(body) {
  const auth = requireRefCihazAdmin_(body);
  if (auth.error) return { ok: false, error: auth.error };
  if (!body.id) return { ok: false, error: 'id_eksik' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_('RefCihazlar');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const data = sheet.getDataRange().getValues();
    const idCol = headers.indexOf('id');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(body.id)) {
        sheet.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: 'cihaz_not_found' };
  } finally {
    lock.releaseLock();
  }
}
