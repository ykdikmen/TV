/**
 * TV PANEL · sunucu (Google Apps Script)
 * "Tur Programi" takvimini okur; hava durumunu, hedefleri ve arka plan fotoğraflarını ekleyip TV'deki panele verir.
 * Hiçbir şeyi değiştirmez: takvimi ve Drive'ı sadece okur. Kendi ayarları için "TV Panel" adında ayrı bir tablo açar.
 *
 * Kurulum:  1) Bu dosyanın tamamını Code.gs'e yapıştır.
 *           2) Üstten "kurulum" fonksiyonunu seçip bir kez çalıştır, izinleri onayla.
 *           3) Dağıt › Yeni dağıtım › Web uygulaması (Farklı çalıştır: Ben · Erişim: Herkes).
 *
 * Güncelleme: 1) Bu dosyanın tamamını Code.gs'e yapıştır.  2) "kurulum"u bir kez daha çalıştır (tabloya eksik sayfaları ekler, TV kodunu değiştirmez).
 *             3) Dağıt › Dağıtımları yönet › kalem › Sürüm: Yeni sürüm › Dağıt (adres aynı kalır).
 */
var AYAR = {
  takvimAdi: "Tur Programi",   // turların durduğu Google Takvim
  tabloAdi: "TV Panel",        // hedefler ve ayarlar için açılacak tablo
  klasorAdi: "TV Panel",       // arka plan fotoğraflarının klasörü (Drive)
  sehir: "Roma", enlem: 41.9028, boylam: 12.4964,
  tz: "Europe/Rome"
};

/* ------------------------------------------------------------------ kurulum */

function kurulum() {
  var pr = PropertiesService.getScriptProperties();
  var tk = CalendarApp.getCalendarsByName(AYAR.takvimAdi);
  if (!tk.length) throw new Error('"' + AYAR.takvimAdi + '" adında bir takvim bulunamadı.');
  pr.setProperty("takvimId", tk[0].getId());

  var ss = null, id = pr.getProperty("tabloId");
  try { if (id) ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  if (!ss) { ss = SpreadsheetApp.create(AYAR.tabloAdi); pr.setProperty("tabloId", ss.getId()); tabloHazirla_(ss); }
  motorSayfasi_(ss);
  sayfaSuresiSatiri_(ss);

  var kl = null, kid = pr.getProperty("klasorId");
  try { if (kid) kl = DriveApp.getFolderById(kid); } catch (e) { kl = null; }
  if (!kl) {
    var it = DriveApp.getFoldersByName(AYAR.klasorAdi);
    kl = it.hasNext() ? it.next() : DriveApp.createFolder(AYAR.klasorAdi);
    pr.setProperty("klasorId", kl.getId());
  }

  CacheService.getScriptCache().removeAll(["veri", "ayar", "eski", "eski2", "hava", "hava2", "foto"]);
  var a = ayarOku_(), v = veriKur_(a), n = 0;
  v.hafta.forEach(function (g) { n += g.turlar.length; });
  Logger.log("KURULUM TAMAM");
  Logger.log("TV kodu: " + a.kod);
  Logger.log("Sürücü/rehber sayısı: " + v.ekip.liste.length + ". Motor sayısı: " + v.motorlar.length + (v.motorOrnek ? " (örnek satırlar)" : ""));
  Logger.log("Bu hafta takvimde " + n + " tur bulundu. Fotoğraf sayısı: " + v.fotoSayi + ". Hava: " + (v.hava ? v.hava.derece + "° " + v.hava.durum : "alınamadı"));
  Logger.log("Ayar tablosu: " + ss.getUrl());
  Logger.log("Fotoğraf klasörü: " + kl.getUrl());
}

function tabloHazirla_(ss) {
  var h = ss.getSheets()[0];
  h.setName("Hedefler");
  h.getRange(1, 1, 4, 4).setValues([
    ["Hedef", "Durum", "İlerleme %", "Not"],
    ["Şirket banka hesabını aç", "Bu hafta", 30, "örnek satır, kendi hedefinle değiştir"],
    ["Fiat 500 cabrio turlarını başlat", "Aralık", 55, "örnek satır, kendi hedefinle değiştir"],
    ["GetYourGuide puanı 4,9", "4,8", 80, "örnek satır, kendi hedefinle değiştir"]
  ]);
  h.getRange("A1:D1").setFontWeight("bold");
  h.setColumnWidth(1, 300); h.setColumnWidth(4, 320); h.setFrozenRows(1);

  var a = ss.insertSheet("Ayarlar"), kod = "", harf = "abcdefghjkmnpqrstuvwxyz23456789";
  for (var i = 0; i < 8; i++) kod += harf.charAt(Math.floor(Math.random() * harf.length));
  a.getRange(1, 1, 6, 3).setValues([
    ["Ayar", "Değer", "Açıklama"],
    ["İsim", "Kerem", "Karşılamada görünen isim"],
    ["Haftalık hedef (€ net)", "", "Boş bırakırsan geçen haftayla kıyaslar"],
    ["Aylık hedef (€ net)", "", "Boş bırakırsan geçen ayla kıyaslar"],
    ["Fotoğraf süresi (sn)", 45, "Arka plan kaç saniyede bir değişsin"],
    ["TV kodu", kod, "TV'de panel ilk açıldığında sorulan kod. Değiştirirsen TV'de yeniden girersin."]
  ]);
  a.getRange(7, 1, 1, 3).setValues([["Sayfa süresi (sn)", 60, "Paneldeki sayfalar kaç saniyede bir değişsin"]]);
  a.getRange("A1:C1").setFontWeight("bold");
  a.setColumnWidth(1, 220); a.setColumnWidth(2, 160); a.setColumnWidth(3, 460); a.setFrozenRows(1);
}

/* Motorlar sayfası yoksa örnek satırlarla ekler; varsa dokunmaz. */
function motorSayfasi_(ss) {
  if (ss.getSheetByName("Motorlar")) return;
  var m = ss.insertSheet("Motorlar"), y = new Date().getFullYear() + 1, ornek = "örnek satır, kendi motorunla değiştir";
  m.getRange(1, 1, 4, 8).setValues([
    ["Motor", "Plaka", "Durum", "Km", "Sigorta bitiş", "Muayene (revizyon)", "Son bakım", "Not"],
    ["Vespa 1 · Primavera 125", "AB 12345", "Hazır", 12400, new Date(y, 2, 15), new Date(y, 5, 30), new Date(y - 1, 9, 20), ornek],
    ["Vespa 2 · Primavera 125", "AB 12346", "Bakımda", 18950, new Date(y, 2, 15), new Date(y - 1, 11, 5), new Date(y - 1, 10, 2), ornek],
    ["Fiat 500 cabrio", "RM 500AB", "Hazır", 64200, new Date(y, 0, 10), new Date(y, 8, 1), new Date(y - 1, 8, 12), ornek]
  ]);
  m.getRange("A1:H1").setFontWeight("bold");
  m.getRange("E2:G200").setNumberFormat("dd.MM.yyyy");
  m.setColumnWidth(1, 240); m.setColumnWidth(8, 320); m.setFrozenRows(1);
}
/* Ayarlar sayfasında "Sayfa süresi" satırı yoksa en alta ekler. */
function sayfaSuresiSatiri_(ss) {
  var a = ss.getSheetByName("Ayarlar");
  if (!a) return;
  var var_ = a.getDataRange().getValues().some(function (r) { return sade_(r[0]).indexOf("sayfa") === 0; });
  if (!var_) a.appendRow(["Sayfa süresi (sn)", 60, "Paneldeki sayfalar kaç saniyede bir değişsin"]);
}

/* ------------------------------------------------------------------ web uç noktası */

function doGet(e) {
  var p = (e && e.parameter) || {}, cb = /^[A-Za-z0-9_]{1,40}$/.test(p.cb || "") ? p.cb : "", out;
  try {
    var c = CacheService.getScriptCache(), yanlis = +(c.get("yanlis") || 0);
    var ayar = ayarOku_();
    if (yanlis >= 30) out = { hata: "kilit" };
    else if (!p.k || !ayar.kod || String(p.k).trim().toLowerCase() !== String(ayar.kod).trim().toLowerCase()) { c.put("yanlis", String(yanlis + 1), 600); out = { hata: "kod" }; }
    else if (p.islem === "foto") out = fotoVer_(+p.no || 0);
    else { out = veriKur_(ayar); out.simdi = Utilities.formatDate(new Date(), AYAR.tz, "yyyy-MM-dd'T'HH:mm:ss"); }  /* veri önbellekten gelse de saat taze olsun */
  } catch (err) { out = { hata: "sunucu", ayrinti: String((err && err.message) || err) }; }
  var js = JSON.stringify(out);
  return cb
    ? ContentService.createTextOutput(cb + "(" + js + ");").setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(js).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ ayarlar ve hedefler */

function sayi_(x) {
  if (typeof x === "number") return x;
  var s = String(x || "").replace(/[^\d.,-]/g, "");
  if (!s) return 0;
  if (s.indexOf(",") >= 0) s = s.replace(/\./g, "").replace(",", ".");
  return parseFloat(s) || 0;
}
function sade_(s) {
  return String(s || "").toLowerCase().replace(/ı/g, "i").replace(/İ/g, "i").replace(/ş/g, "s").replace(/ğ/g, "g").replace(/ü/g, "u").replace(/ö/g, "o").replace(/ç/g, "c").replace(/[̀-ͯ]/g, "").trim();
}

function ayarOku_() {
  var c = CacheService.getScriptCache(), h = c.get("ayar");
  if (h) return JSON.parse(h);
  var a = { isim: "", hedefHafta: 0, hedefAy: 0, fotoSure: 45, sayfaSure: 60, kod: "", hedefler: [], hedefOrnek: false, motorlar: [], motorOrnek: false };
  var id = PropertiesService.getScriptProperties().getProperty("tabloId");
  if (!id) throw new Error("Önce kurulum fonksiyonunu çalıştır.");
  var ss = SpreadsheetApp.openById(id);
  var ay = ss.getSheetByName("Ayarlar");
  if (ay) ay.getDataRange().getValues().forEach(function (r) {
    var k = sade_(r[0]), v = r[1];
    if (k.indexOf("isim") === 0) a.isim = String(v || "").trim();
    else if (k.indexOf("haftalik") === 0) a.hedefHafta = sayi_(v);
    else if (k.indexOf("aylik") === 0) a.hedefAy = sayi_(v);
    else if (k.indexOf("fotograf") === 0) a.fotoSure = Math.max(10, sayi_(v) || 45);
    else if (k.indexOf("sayfa") === 0) a.sayfaSure = Math.max(15, sayi_(v) || 60);
    else if (k.indexOf("tv kodu") === 0) a.kod = String(v || "").trim();
  });
  var hs = ss.getSheetByName("Hedefler");
  if (hs) hs.getDataRange().getValues().slice(1).forEach(function (r) {
    var ad = String(r[0] || "").trim();
    if (!ad || a.hedefler.length >= 5) return;
    var o = sayi_(r[2]);
    if (typeof r[2] === "number" && o > 0 && o <= 1) o = o * 100;
    var not = String(r[3] || "").trim();
    if (/örnek/i.test(not)) a.hedefOrnek = true;
    a.hedefler.push({ ad: ad, durum: String(r[1] === 0 ? "0" : r[1] || "").trim(), oran: Math.round(o), alt: /örnek/i.test(not) ? "" : not });
  });
  var ms = ss.getSheetByName("Motorlar"), tz = ss.getSpreadsheetTimeZone() || AYAR.tz;
  if (ms) ms.getDataRange().getValues().slice(1).forEach(function (r) {
    var ad = String(r[0] || "").trim();
    if (!ad || a.motorlar.length >= 24) return;
    var not = String(r[7] || "").trim(), ornek = /örnek/i.test(not);
    if (ornek) a.motorOrnek = true;
    a.motorlar.push({ ad: ad, plaka: String(r[1] || "").trim(), durum: String(r[2] || "").trim(), km: r[3] === "" || r[3] == null ? "" : Math.round(sayi_(r[3])),
      sigorta: tarihYaz_(r[4], tz), muayene: tarihYaz_(r[5], tz), bakim: tarihYaz_(r[6], tz), not: ornek ? "" : not });
  });
  try { c.put("ayar", JSON.stringify(a), 300); } catch (e) {}
  return a;
}
/* Tablodaki tarih hücresi → 'yyyy-MM-dd'. Tarih değilse (ör. "bilinmiyor") yazıyı olduğu gibi verir. */
function tarihYaz_(x, tz) {
  if (Object.prototype.toString.call(x) === "[object Date]") return isNaN(x.getTime()) ? "" : Utilities.formatDate(x, tz, "yyyy-MM-dd");
  var s = String(x == null ? "" : x).trim(), m = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  return m ? m[3] + "-" + p2_(+m[2]) + "-" + p2_(+m[1]) : s;
}

/* ------------------------------------------------------------------ takvim */

function zaman_(tarih, saat) { return Utilities.parseDate(tarih + " " + (saat || "00:00"), AYAR.tz, "yyyy-MM-dd HH:mm"); }
function duzMetin_(s) {
  return String(s || "").replace(/<br\s*\/?>/gi, "\n").replace(/<\/(p|div|li)>/gi, "\n").replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&#39;/g, "'").replace(/&quot;/g, '"');
}

/* Takvim kartı örneği:  başlık "16:00  Ad Soyad  (2)  GYG"  ·  açıklamada "Urun: … / Kisi: … / Net: EUR 138,60 / Guide: … / Drivers: … / Kapanış: …" */
function olayCoz_(baslik, aciklama, yer, basZaman) {
  baslik = String(baslik || "");
  var m = baslik.match(/^\s*(IPTAL\s*-\s*)?(\d{1,2}[:.]\d{2})\s+(.*?)\s*\((\d+)\)\s*(\S*)\s*$/i), al = {};
  duzMetin_(aciklama).split("\n").forEach(function (l) {
    var k = l.match(/^\s*([A-Za-zÇĞİÖŞÜçğıöşü]+)\s*:\s*(.*)$/);
    if (!k) return;
    var ad = sade_(k[1]), d = k[2].trim();
    if (ad === "not") al.not = al.not ? al.not + " · " + d : d;
    else if (al[ad] === undefined) al[ad] = d;
  });
  if (!m && al.platform === undefined && al.urun === undefined) return null;
  var iptal = !!(m && m[1]) || /^\s*[İI]PTAL/i.test(baslik) || /iptal/i.test(al.kapanis || "");
  return {
    tarih: Utilities.formatDate(basZaman, AYAR.tz, "yyyy-MM-dd"), saat: Utilities.formatDate(basZaman, AYAR.tz, "HH:mm"),
    platform: al.platform || (m && m[5]) || "", tur: al.urun || "", kisi: sayi_(al.kisi) || (m ? +m[4] : 0),
    musteri: m ? m[3].trim() : baslik.trim(), dil: al.dil || "", bulusma: String(yer || ""), not: al.not || "",
    net: sayi_(al.net), durum: iptal ? "iptal" : "aktif", rehber: al.guide || al.rehber || "", suruculer: al.drivers || al.suruculer || ""
  };
}

function takvimOku_(bas, bit) {
  var id = PropertiesService.getScriptProperties().getProperty("takvimId");
  var cal = id ? CalendarApp.getCalendarById(id) : null;
  if (!cal) throw new Error("Takvim bulunamadı; kurulum fonksiyonunu yeniden çalıştır.");
  var out = [];
  cal.getEvents(bas, bit).forEach(function (ev) {
    if (ev.isAllDayEvent()) return;
    var r = olayCoz_(ev.getTitle(), ev.getDescription(), ev.getLocation(), ev.getStartTime());
    if (r) out.push(r);
  });
  return out;
}

/* ------------------------------------------------------------------ hava (Open-Meteo, ücretsiz, anahtarsız) */

var HAVA_KOD = { 0: "Açık", 1: "Az bulutlu", 2: "Parçalı bulut", 3: "Bulutlu", 45: "Sisli", 48: "Sisli", 51: "Çisenti", 53: "Çisenti", 55: "Çisenti", 56: "Çisenti", 57: "Çisenti", 61: "Hafif yağmur", 63: "Yağmur", 65: "Kuvvetli yağmur", 66: "Yağmur", 67: "Yağmur", 71: "Kar", 73: "Kar", 75: "Kar", 77: "Kar", 80: "Sağanak", 81: "Sağanak", 82: "Kuvvetli sağanak", 85: "Kar", 86: "Kar", 95: "Gök gürültülü", 96: "Gök gürültülü", 99: "Gök gürültülü" };
function havaIkon_(kod, gunduz) {
  if (kod >= 51) return "yagmur";
  if (kod === 3 || kod === 45 || kod === 48) return "bulut";
  if (!gunduz) return "gece";
  return kod === 0 ? "gunes" : "parcali";
}

function havaAl_(ileri, bugun, saat) {
  var c = CacheService.getScriptCache(), ham = c.get("hava2"), w;
  try {
    if (!ham) {
      ham = UrlFetchApp.fetch("https://api.open-meteo.com/v1/forecast?latitude=" + AYAR.enlem + "&longitude=" + AYAR.boylam +
        "&current=temperature_2m,apparent_temperature,weather_code,wind_speed_10m,is_day,relative_humidity_2m&hourly=temperature_2m,weather_code,precipitation_probability,is_day" +
        "&daily=weather_code,temperature_2m_max,temperature_2m_min,precipitation_probability_max,sunrise,sunset&timezone=" + encodeURIComponent(AYAR.tz) + "&forecast_days=7", { muteHttpExceptions: true }).getContentText();
      JSON.parse(ham).current.temperature_2m.toFixed(0);
      c.put("hava2", ham, 1200);
    }
    w = JSON.parse(ham);
  } catch (e) { return null; }
  var H = w.hourly, D = w.daily, C = w.current, i0 = H.time.indexOf(bugun + "T" + saat.slice(0, 2) + ":00");
  if (i0 < 0) i0 = 0;
  var KISA = ["Paz", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt"];
  var h = {
    sehir: AYAR.sehir, derece: Math.round(C.temperature_2m), hissedilen: Math.round(C.apparent_temperature), ruzgar: Math.round(C.wind_speed_10m),
    durum: HAVA_KOD[C.weather_code] || "", ikon: havaIkon_(C.weather_code, C.is_day), yagmur: H.precipitation_probability[i0] || 0, saatlik: [], gunler: [],
    nem: Math.round(C.relative_humidity_2m || 0), en: Math.round(D.temperature_2m_max[0]), min: Math.round(D.temperature_2m_min[0]),
    yagmurGun: D.precipitation_probability_max[0] || 0, dogus: String((D.sunrise || [""])[0]).slice(11, 16), batis: String((D.sunset || [""])[0]).slice(11, 16)
  };
  for (var i = i0 + 1; i <= i0 + 12 && i < H.time.length; i++) h.saatlik.push({ s: H.time[i].slice(11, 13), d: Math.round(H.temperature_2m[i]), i: havaIkon_(H.weather_code[i], H.is_day[i]), y: H.precipitation_probability[i] || 0 });
  for (var j = 1; j < D.time.length && h.gunler.length < 6; j++) h.gunler.push({ g: KISA[pTarih_(D.time[j]).getUTCDay()], i: havaIkon_(D.weather_code[j], 1), y: D.precipitation_probability_max[j] || 0, en: Math.round(D.temperature_2m_max[j]), min: Math.round(D.temperature_2m_min[j]) });

  /* turlarla ilgili kısa not */
  var simdiDk = pDk_(saat), kalan = ileri.filter(function (r) { return r.tarih === bugun && pDk_(r.saat) > simdiDk; }), not = "";
  if (kalan.length) {
    var enCok = 0, hangi = "";
    kalan.forEach(function (r) {
      var k = H.time.indexOf(bugun + "T" + r.saat.slice(0, 2) + ":00");
      for (var x = k; k >= 0 && x <= k + 1 && x < H.time.length; x++) if ((H.precipitation_probability[x] || 0) > enCok) { enCok = H.precipitation_probability[x]; hangi = r.saat; }
    });
    not = enCok >= 50 ? "Bugün " + hangi + " turunda yağmur ihtimali %" + enCok + "." : enCok >= 25 ? "Bugün kalan turlarda hafif yağmur ihtimali var (%" + enCok + ")." : "Bugün kalan turlarda yağmur beklenmiyor.";
  } else {
    var sonra = ileri.filter(function (r) { return r.tarih > bugun; })[0], di = sonra ? D.time.indexOf(sonra.tarih) : -1;
    if (di > 0) { var p = D.precipitation_probability_max[di] || 0, ga = di === 1 ? "Yarın" : P_GUN[pTarih_(sonra.tarih).getUTCDay()]; not = p >= 25 ? ga + " (sıradaki tur günü) yağmur ihtimali %" + p + "." : ga + " (sıradaki tur günü) için yağmur beklenmiyor."; }
    else not = "Bugün en yüksek " + Math.round(D.temperature_2m_max[0]) + "°, en düşük " + Math.round(D.temperature_2m_min[0]) + "°.";
  }
  h.not = not; h.cumle = not;
  return h;
}

/* ------------------------------------------------------------------ panel verisi */

function veriKur_(ayar) {
  var c = CacheService.getScriptCache(), h = c.get("veri");
  if (h) return JSON.parse(h);
  var simdi = new Date(), bugun = Utilities.formatDate(simdi, AYAR.tz, "yyyy-MM-dd"), saat = Utilities.formatDate(simdi, AYAR.tz, "HH:mm");
  var hg = pTarih_(bugun).getUTCDay(), pzt = pGunEkle_(bugun, -((hg + 6) % 7));
  var ayBas = bugun.slice(0, 8) + "01", d = pTarih_(ayBas), aySon = pYaz_(new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0)));
  var oncBas = pGunEkle_(ayBas, -1).slice(0, 8) + "01";
  var son = pGunEkle_(bugun, 14); if (aySon > son) son = aySon;

  /* bu haftadan öncesi sadece toplamlar için gerekir: sade tutulur ve 1 saat saklanır */
  var eski, eh = c.get("eski2");
  if (eh) { eh = JSON.parse(eh); if (eh.bas === oncBas && eh.bit === pzt) eski = eh.s; }
  if (!eski) {
    eski = takvimOku_(zaman_(oncBas), zaman_(pzt)).filter(function (r) { return r.durum !== "iptal"; }).map(function (r) { return [r.tarih, Math.round(r.net * 100) / 100, r.kisi, r.rehber, r.suruculer, pPlatform_(r.platform)]; });
    try { c.put("eski2", JSON.stringify({ bas: oncBas, bit: pzt, s: eski }), 3600); } catch (e) {}
  }
  var satirlar = eski.map(function (x) { return { tarih: x[0], saat: "", net: x[1], kisi: x[2], durum: "aktif", rehber: x[3] || "", suruculer: x[4] || "", platform: x[5] || "" }; });
  var yeni = takvimOku_(zaman_(pzt), zaman_(pGunEkle_(son, 1)));
  satirlar = satirlar.concat(yeni);

  var hava = havaAl_(yeni.filter(function (r) { return r.durum !== "iptal" && r.tarih >= bugun; }).sort(function (a, b) { return a.tarih < b.tarih ? -1 : a.tarih > b.tarih ? 1 : pDk_(a.saat) - pDk_(b.saat); }), bugun, saat);
  var v = panelVerisiKur(satirlar, bugun, saat, ayar, hava);
  v.hafta.forEach(function (g) { if (g.tarih !== bugun) g.turlar.forEach(function (t) { t.musteri = ""; }); });
  v.simdi = bugun + "T" + saat + ":" + Utilities.formatDate(simdi, AYAR.tz, "ss");
  v.fotoSayi = fotoListe_().length; v.fotoSure = ayar.fotoSure; v.sayfaSure = ayar.sayfaSure;
  v.motorlar = ayar.motorlar || []; v.motorOrnek = !!ayar.motorOrnek;
  try { c.put("veri", JSON.stringify(v), 240); } catch (e) {}
  return v;
}

/* ------------------------------------------------------------------ arka plan fotoğrafları */

function fotoListe_() {
  var c = CacheService.getScriptCache(), h = c.get("foto");
  if (h) return JSON.parse(h);
  var l = [], kid = PropertiesService.getScriptProperties().getProperty("klasorId");
  try {
    var it = DriveApp.getFolderById(kid).getFiles();
    while (it.hasNext()) { var f = it.next(); if (/^image\//.test(String(f.getMimeType()))) l.push([f.getName(), f.getId()]); }
  } catch (e) {}
  if (!l.length) {  /* klasör boşsa: Drive'daki "pexels" adlı fotoğraflar */
    try {
      var s = DriveApp.searchFiles("title contains 'pexels' and mimeType contains 'image/' and trashed = false");
      while (s.hasNext() && l.length < 12) { var g = s.next(); l.push([g.getName(), g.getId()]); }
    } catch (e) {}
  }
  l.sort();
  var ids = l.slice(0, 12).map(function (x) { return x[1]; });
  c.put("foto", JSON.stringify(ids), 1800);
  return ids;
}

function fotoVer_(no) {
  var ids = fotoListe_();
  if (!ids.length) return { n: 0, no: 0, u: "" };
  no = ((no % ids.length) + ids.length) % ids.length;
  var blob = null;
  try {  /* Drive'ın küçültülmüş hali (1920 px genişlik) */
    var r = UrlFetchApp.fetch("https://drive.google.com/thumbnail?sz=w1920&id=" + ids[no], { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (r.getResponseCode() === 200 && /^image\//.test(String(r.getBlob().getContentType()))) blob = r.getBlob();
  } catch (e) {}
  if (!blob) { try { var f = DriveApp.getFileById(ids[no]); if (f.getSize() <= 3 * 1024 * 1024) blob = f.getBlob(); } catch (e) {} }
  if (!blob) return { n: ids.length, no: no, u: "" };
  return { n: ids.length, no: no, u: "data:" + (blob.getContentType() || "image/jpeg") + ";base64," + Utilities.base64Encode(blob.getBytes()) };
}

/* Panel mantığı: Rezervasyonlar satırlarından panelin gösterdiği veriyi kurar.
   Aynı dosya hem Apps Script'te hem de yerel testte (Node) çalışır. */
var P_GUN = ["Pazar", "Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi"];
var P_AY = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"];

function p2_(n) { return (n < 10 ? "0" : "") + n; }
function pTarih_(s) { var p = s.split("-"); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])); }
function pYaz_(d) { return d.getUTCFullYear() + "-" + p2_(d.getUTCMonth() + 1) + "-" + p2_(d.getUTCDate()); }
function pGunEkle_(s, n) { var d = pTarih_(s); d.setUTCDate(d.getUTCDate() + n); return pYaz_(d); }
function pPara_(n) { var s = String(Math.round(n)), c = ""; while (s.length > 3) { c = "." + s.slice(-3) + c; s = s.slice(0, -3); } return "€" + s + c; }
function pDk_(saat) { var p = String(saat || "00:00").split(":"); return +p[0] * 60 + +p[1]; }

function pPlatform_(p) { p = String(p || ""); return /get|gyg/i.test(p) ? "gyg" : /viator|^via$/i.test(p) ? "viator" : "direkt"; }
/* "Ali, Mehmet ve Luca" → ["Ali", "Mehmet", "Luca"] */
function pKisiler_(s) {
  return String(s || "").split(/\s*[,;\/&+]\s*|\s+(?:ve|and|e)\s+/).map(function (x) { return x.replace(/\s+/g, " ").trim(); })
    .filter(function (x) { return x && !/^[-–—?.]+$/.test(x) && !/^(yok|none|tbd|belli de[gğ]il|atanmad[ıi])$/i.test(x); });
}

function pUrunKisa_(t) {
  t = String(t || "");
  if (/fiat/i.test(t)) return "Fiat 500";
  if (/golf/i.test(t)) return "Golf cart";
  if (/solo/i.test(t)) return "Solo Vespa";
  if (/lizzie/i.test(t)) return "Lizzie McGuire";
  if (/private|özel|ozel/i.test(t)) return "Özel Vespa";
  if (/vespa|guide|rehber/i.test(t)) return "Vespa turu";
  return t.length > 22 ? t.slice(0, 21) + "…" : (t || "Tur");
}

function pDilKodu_(d) {
  var m = [[/ispanyol|spanish|español/i, "ES"], [/frans|french|français/i, "FR"], [/portekiz|portug/i, "PT"], [/alman|german|deutsch/i, "DE"], [/ital/i, "IT"], [/leh[çc]e|polish/i, "PL"], [/t[üu]rk/i, "TR"], [/rus/i, "RU"]];
  for (var i = 0; i < m.length; i++) if (m[i][0].test(d)) return m[i][1];
  return "";
}

function pEtiketler_(r) {
  var e = [], not = String(r.not || ""), tur = String(r.tur || ""), bul = String(r.bulusma || ""), dil = String(r.dil || "");
  if (/adres yok/i.test(bul)) e.push("adres yok");
  else if (/not yet booked|henuz rezerve|henüz rezerve|not listed/i.test(bul)) e.push("otel belirsiz");
  if (/DEGISTI|DEĞİŞTİ/.test(not)) e.push("değişti");
  if (/son dakika/i.test(not)) e.push("son dakika");
  var kod = pDilKodu_(dil);
  if (kod) e.push((/sesli|audio|kulakl/i.test(dil) ? "sesli " : "canlı ") + kod);
  if (/drop.?off/i.test(tur + " " + not)) e.push("drop-off");
  if (/extended by|uzat[iı]lm[iı][sş]/i.test(tur + " " + not)) e.push("uzatılmış");
  if (/child|çocuk|cocuk|infant|bebek/i.test(not)) e.push("çocuk");
  return e;
}

/* satirlar: [{tarih:'yyyy-MM-dd', saat:'HH:mm', platform, tur, kisi, musteri, dil, bulusma, not, net, durum}]
   bugun: 'yyyy-MM-dd', saat: 'HH:mm', ayar: {isim, hedefHafta, hedefAy, hedefler[], brif[], brifSaat}, hava: {...} */
function panelVerisiKur(satirlar, bugun, saat, ayar, hava) {
  ayar = ayar || {};
  var aktif = satirlar.filter(function (r) { return r.tarih && String(r.durum || "aktif").toLowerCase() !== "iptal"; });
  aktif.sort(function (a, b) { return a.tarih < b.tarih ? -1 : a.tarih > b.tarih ? 1 : pDk_(a.saat) - pDk_(b.saat); });
  function aralik(a, b) { return aktif.filter(function (r) { return r.tarih >= a && r.tarih <= b; }); }
  function top(l) { var n = 0, k = 0; l.forEach(function (r) { n += +r.net || 0; k += +r.kisi || 0; }); return { net: n, kisi: k, tur: l.length }; }

  var hg = pTarih_(bugun).getUTCDay();
  var pzt = pGunEkle_(bugun, -((hg + 6) % 7)), paz = pGunEkle_(pzt, 6);
  var simdiDk = pDk_(saat);

  /* hafta */
  var hafta = [];
  for (var i = 0; i < 7; i++) {
    var t = pGunEkle_(pzt, i);
    hafta.push({ tarih: t, turlar: aralik(t, t).map(function (r) {
      var ekip = [r.rehber, r.suruculer].filter(function (x) { return x && String(x).trim(); }).join(" · ");
      var ileride = r.tarih > bugun || (r.tarih === bugun && pDk_(r.saat) > simdiDk);
      return { saat: r.saat || "--:--", kisi: +r.kisi || 0, platform: pPlatform_(r.platform), urun: pUrunKisa_(r.tur), musteri: String(r.musteri || ""), etiket: pEtiketler_(r), ekip: ekip, atanmadi: ileride && !ekip && r.rehber !== undefined };
    }) });
  }

  /* rakamlar */
  var bg = top(aralik(bugun, bugun)), gp = top(aralik(pGunEkle_(bugun, -7), pGunEkle_(bugun, -7)));
  var hf = top(aralik(pzt, paz)), hfYap = top(aralik(pzt, bugun)), gh = top(aralik(pGunEkle_(pzt, -7), pGunEkle_(pzt, -1)));
  var ayBas = bugun.slice(0, 8) + "01", d = pTarih_(ayBas);
  var aySon = pYaz_(new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0)));
  var oncSon = pGunEkle_(ayBas, -1), oncBas = oncSon.slice(0, 8) + "01";
  var ay = top(aralik(ayBas, aySon)), ayYap = top(aralik(ayBas, bugun)), oa = top(aralik(oncBas, oncSon));
  var ayAd = P_AY[d.getUTCMonth()], oncAd = P_AY[pTarih_(oncBas).getUTCMonth()];
  var rakamlar = [
    { ad: "Bugün", net: bg.net, tur: bg.tur, kisi: bg.kisi, kiyasAd: "Geçen " + P_GUN[hg], kiyas: gp.net, alt: "Geçen " + P_GUN[hg] + " " + gp.tur + " tur vardı" },
    { ad: "Bu hafta", net: hf.net, tur: hf.tur, kisi: hf.kisi, kiyasAd: ayar.hedefHafta ? "Hedef" : "Geçen hafta", kiyas: ayar.hedefHafta || gh.net, alt: "Yapılan " + pPara_(hfYap.net) + " · kalan rezerve " + pPara_(hf.net - hfYap.net) },
    { ad: ayAd, net: ay.net, tur: ay.tur, kisi: ay.kisi, kiyasAd: ayar.hedefAy ? "Hedef" : oncAd, kiyas: ayar.hedefAy || oa.net, alt: "Yapılan " + pPara_(ayYap.net) + " · kalan rezerve " + pPara_(ay.net - ayYap.net) }
  ];

  /* ay takvimi: ayın her günü için tur ve kişi sayısı */
  var ayGun = [];
  for (var tg = ayBas; tg <= aySon; tg = pGunEkle_(tg, 1)) {
    var gl = aralik(tg, tg), go = { t: tg, n: gl.length, k: 0, e: 0, v: 0, g: 0, d: 0 };
    gl.forEach(function (r) { go.k += +r.kisi || 0; go.e += +r.net || 0; var p = pPlatform_(r.platform); go[p === "gyg" ? "g" : p === "viator" ? "v" : "d"]++; });
    go.e = Math.round(go.e);
    ayGun.push(go);
  }

  /* ekip: sürücü ve rehberlerin kaç tura çıktığı */
  var kisiler = {}, sira = [], ekipsiz = 0;
  aktif.forEach(function (r) {
    var buAy = r.tarih >= ayBas && r.tarih <= aySon, onc = r.tarih >= oncBas && r.tarih <= oncSon, hft = r.tarih >= pzt && r.tarih <= paz;
    if (!buAy && !onc && !hft) return;
    var gecti = r.tarih < bugun || (r.tarih === bugun && pDk_(r.saat) <= simdiDk);
    var reh = pKisiler_(r.rehber), sur = pKisiler_(r.suruculer), gorulen = {};
    if (buAy && !gecti && !reh.length && !sur.length) ekipsiz++;
    reh.map(function (ad) { return [ad, 1]; }).concat(sur.map(function (ad) { return [ad, 0]; })).forEach(function (x) {
      var k = x[0].toLowerCase(), o = kisiler[k];
      if (!o) { o = kisiler[k] = { ad: x[0], hafta: 0, ay: 0, yap: 0, onceki: 0, rehber: 0, surucu: 0, kisi: 0 }; sira.push(o); }
      if (buAy) { if (x[1]) o.rehber++; else o.surucu++; }
      if (gorulen[k]) return;   /* aynı turda hem rehber hem sürücü yazılmışsa tur bir kez sayılır */
      gorulen[k] = 1;
      if (hft) o.hafta++;
      if (onc) o.onceki++;
      if (buAy) { o.ay++; o.kisi += +r.kisi || 0; if (gecti) o.yap++; }
    });
  });
  sira.sort(function (a, b) { return b.ay - a.ay || b.onceki - a.onceki || (a.ad < b.ad ? -1 : 1); });

  /* özet cümleleri */
  var c = [], bugunTur = hafta[(hg + 6) % 7].turlar;
  function gunAdi(t) { return t === bugun ? "Bugün" : t === pGunEkle_(bugun, 1) ? "Yarın" : P_GUN[pTarih_(t).getUTCDay()]; }
  if (bugunTur.length) {
    var kalan = bugunTur.filter(function (x) { return pDk_(x.saat) > simdiDk; });
    c.push("Bugün <b>" + bg.tur + " tur, " + bg.kisi + " kişi</b> var" + (kalan.length ? "; " + kalan.length + " tanesi daha çıkmadı." : "; hepsi çıktı."));
  } else c.push("Bugün <b>tur yok</b>.");
  var sonraki = aktif.filter(function (r) { return r.tarih > bugun || (r.tarih === bugun && pDk_(r.saat) > simdiDk); })[0];
  if (sonraki) c.push("Sıradaki çıkış <b>" + gunAdi(sonraki.tarih) + " " + sonraki.saat + "</b>: " + pUrunKisa_(sonraki.tur) + ", " + (+sonraki.kisi || 0) + " kişi.");
  /* önümüzdeki günlerde aynı saatte çakışan çıkışlar ve takip isteyenler */
  var cakisma = 0, takip = 0;
  hafta.forEach(function (g) {
    if (g.tarih < bugun) return;
    var say = {};
    g.turlar.forEach(function (x) {
      if (g.tarih === bugun && pDk_(x.saat) <= simdiDk) return;
      say[x.saat] = (say[x.saat] || 0) + 1;
      if (takip < 2) x.etiket.forEach(function (e) {
        if (takip < 2 && (e === "adres yok" || e === "otel belirsiz" || e === "değişti")) { takip++; c.push("<b>" + gunAdi(g.tarih) + " " + x.saat + "</b> turu: " + (e === "değişti" ? "rezervasyon değişmiş, kontrol et." : e === "adres yok" ? "mailde adres yok, panelden bak." : "müşterinin oteli belli değil.")); }
      });
    });
    for (var s in say) if (say[s] > 1 && cakisma < 2) { cakisma++; c.push("<b>" + gunAdi(g.tarih) + " " + s + "</b>: " + say[s] + " çıkış aynı anda, sürücüleri ona göre ayarla."); }
  });
  var ats = 0, atIlk = "";
  hafta.forEach(function (g) { g.turlar.forEach(function (x) { if (x.atanmadi) { if (!ats) atIlk = gunAdi(g.tarih) + " " + x.saat; ats++; } }); });
  if (ats) c.push("<b>" + ats + " turun ekibi henüz atanmadı</b>; ilki " + atIlk + ".");
  c.push("Bu hafta toplam <b>" + hf.tur + " tur, " + pPara_(hf.net) + "</b>." + (rakamlar[1].kiyas > 0 ? " " + rakamlar[1].kiyasAd + ": " + pPara_(rakamlar[1].kiyas) + " (%" + Math.round(hf.net / rakamlar[1].kiyas * 100) + ")." : ""));
  c.push(ayAd + ": şu ana kadar <b>" + ayYap.tur + " tur, " + pPara_(ayYap.net) + "</b>. Ay sonuna kadar rezerve toplam " + pPara_(ay.net) + ".");
  if (hava && hava.cumle) c.push(hava.cumle);
  var acik = (ayar.hedefler || []).filter(function (h) { return +h.oran < 100; })[0];
  if (acik && !ayar.hedefOrnek) c.push("Hello Rides'ta sıradaki hedef: <b>" + acik.ad + "</b>.");

  var claude = ayar.brif && ayar.brif.length;
  return {
    isim: ayar.isim || "", bugun: bugun, guncel: saat,
    hava: hava || null, hedefler: ayar.hedefler || [], hedefOrnek: !!ayar.hedefOrnek,
    hafta: hafta, rakamlar: rakamlar,
    ay: { ad: ayAd, yil: d.getUTCFullYear(), gunler: ayGun },
    ekip: { ay: ayAd, onceki: oncAd, ekipsiz: ekipsiz, liste: sira.slice(0, 14) },
    brif: { kim: claude ? "Claude" : "Günün özeti", alt: claude ? "Bugün " + (ayar.brifSaat || "") + "'de hazırlandı" : "Tur takviminden otomatik · " + saat, claude: !!claude, cumleler: claude ? ayar.brif : c }
  };
}

