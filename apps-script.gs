/**
 * Biyomedikal Kalibrasyon — Google Apps Script Backend
 *
 * KURULUM ADIMLARI:
 *
 * 1) Yeni bir Google Sheets dosyası oluşturun. (örn: "BiyomedikalKalibrasyon")
 *
 * 2) İçinde 2 sayfa (sheet/sekme) oluşturun:
 *
 *    SAYFA 1: "Kullanicilar"
 *    Başlık satırı (1. satır):
 *      tc | pin | ad | unvan | roller
 *    Örnek satırlar:
 *      41128339146 | 1234 | Yönetici Admin    | Sistem Yöneticisi      | yonetici
 *      12345678901 | 5678 | Cem Özkaya        | Biyomedikal Mühendisi  | teknisyen
 *      11111111111 | 9999 | Müdür Demo        | Şube Müdürü            | mudur,yonetici
 *    NOT: roller virgülle ayrılır (yonetici, mudur, teknisyen).
 *    NOT: tc ve pin sütunları METİN olarak biçimlendirilmelidir
 *         (Format → Number → Plain text), aksi halde başlardaki sıfırlar kaybolur.
 *
 *    SAYFA 2: "Raporlar"
 *    Başlık satırı (1. satır):
 *      raporNo | gonderim_tarihi | tc | ad | onay_durumu | dev_type | saglik_tesisi | marka | model | seri_no | json_veri
 *
 * 3) Sheets URL'inden ID'yi alın. URL şu şekilde:
 *      https://docs.google.com/spreadsheets/d/AAAAA_BBBBB_CCCCC/edit
 *    "AAAAA_BBBBB_CCCCC" kısmı sizin ID'nizdir.
 *    Aşağıdaki SS_ID değerine bunu yazın.
 *
 * 4) Sheets'te Uzantılar (Extensions) → Apps Script
 *    Açılan editöre bu dosyanın TAMAMINI yapıştırın, kaydedin.
 *
 * 5) Apps Script editöründe sağ üstten "Dağıt" (Deploy) → "Yeni dağıtım" (New deployment)
 *    - Tür: "Web uygulaması" (Web app)
 *    - Şu kullanıcı olarak çalıştır: Ben (kendi hesabınız)
 *    - Erişim: "Herkes" (Anyone) — DİKKAT: bu seçeneği seçin, "anonim" değil
 *    - Dağıt'a basın, izinleri onaylayın.
 *    - Çıkan "Web uygulaması URL'sini" kopyalayın.
 *
 * 6) index.html dosyasındaki SHEETS_API_URL sabitine bu URL'i yapıştırın.
 *
 * 7) Kodda bir değişiklik yaparsanız "Dağıt → Dağıtımı yönet" → kalemden "Yeni sürüm"
 *    seçerek tekrar dağıtmayı unutmayın. URL aynı kalır.
 */

const SS_ID = 'BURAYA_SHEETS_ID_YAZIN';

// =====================================================
// HTTP GİRİŞ
// =====================================================
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'login')        return jsonOut(handleLogin(body));
    if (action === 'submitRapor')  return jsonOut(handleSubmitRapor(body));
    if (action === 'listRaporlar') return jsonOut(handleListRaporlar(body));

    return jsonOut({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput('Biyomedikal Kalibrasyon API çalışıyor. POST gerekiyor.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  return SpreadsheetApp.openById(SS_ID).getSheetByName(name);
}

// =====================================================
// AUTH — Kullanıcı doğrulama
// =====================================================
function findUser(tc, pin) {
  const sheet = getSheet('Kullanicilar');
  if (!sheet) throw new Error('Kullanicilar sayfası bulunamadı');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowTc  = String(data[i][0]).trim();
    const rowPin = String(data[i][1]).trim();
    if (rowTc === String(tc).trim() && rowPin === String(pin).trim()) {
      return {
        tc:     rowTc,
        ad:     String(data[i][2] || '').trim(),
        unvan:  String(data[i][3] || '').trim(),
        roller: String(data[i][4] || '').split(',').map(s => s.trim()).filter(Boolean)
      };
    }
  }
  return null;
}

function authUser(body) {
  return findUser(body.tc, body.pin);
}

function handleLogin(body) {
  const user = findUser(body.tc, body.pin);
  if (!user) return { ok: false, error: 'invalid_credentials' };
  return { ok: true, user: user };
}

// =====================================================
// RAPOR — Sheets'e gönderim
// =====================================================
function handleSubmitRapor(body) {
  const user = authUser(body);
  if (!user) return { ok: false, error: 'unauthorized' };

  const r = body.rapor || {};

  // Atomik rapor numarası üretimi (eş zamanlı gönderimlerde çakışmayı önler)
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet('Raporlar');
    if (!sheet) throw new Error('Raporlar sayfası bulunamadı');

    const finalRaporNo = generateRaporNo(sheet, r.devType);
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

    return { ok: true, raporNo: finalRaporNo };
  } finally {
    lock.releaseLock();
  }
}

function generateRaporNo(sheet, devType) {
  const prefixMap = {
    defibrillator: 'DEF',
    aspirator:     'ASP',
    infusion:      'INF',
    monitor:       'MON'
  };
  const prefix = prefixMap[devType] || 'XXX';

  const now = new Date();
  const dd  = pad2(now.getDate());
  const mm  = pad2(now.getMonth() + 1);
  const yy  = String(now.getFullYear()).slice(-2);
  const datePart = dd + mm + yy;

  // Aynı gün, aynı önekteki rapor sayısını bul
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const rno = String(data[i][0] || '');
    if (rno.indexOf(prefix + datePart) === 0) count++;
  }

  return prefix + datePart + pad3(count + 1);
}

function pad2(n) { return String(n).length < 2 ? '0' + n : String(n); }
function pad3(n) {
  const s = String(n);
  if (s.length >= 3) return s;
  return ('000' + s).slice(-3);
}

// =====================================================
// RAPOR LİSTELEME (ileride pgVeriler / pgOnay için)
// =====================================================
function handleListRaporlar(body) {
  const user = authUser(body);
  if (!user) return { ok: false, error: 'unauthorized' };

  const sheet = getSheet('Raporlar');
  const data  = sheet.getDataRange().getValues();
  const rows  = [];
  for (let i = 1; i < data.length; i++) {
    rows.push({
      raporNo:      String(data[i][0]),
      tarih:        data[i][1],
      tc:           String(data[i][2]),
      ad:           String(data[i][3]),
      onayDurumu:   String(data[i][4]),
      devType:      String(data[i][5]),
      saglikTesisi: String(data[i][6]),
      marka:        String(data[i][7]),
      model:        String(data[i][8]),
      seriNo:       String(data[i][9])
    });
  }
  return { ok: true, raporlar: rows };
}
