# TV · Hello Rides ev paneli

Evdeki TV'nin tarayıcısında açılan bilgi ekranı. Sayfalara bölünmüştür; sayfalar dakikada bir sırayla değişir:

1. **Bu hafta** – haftanın turları, bugün geniş sütunda
2. **Ay** – aylık tur takvimi ve ay özeti
3. **Rakamlar** – bugün / hafta / ay net tutarı, gün gün grafik, platform dağılımı
4. **Hava** – şu an, önümüzdeki saatler ve günler
5. **Sürücüler** – sürücü ve rehberlerin kaç tura çıktığı (takvim kartlarındaki `Guide:` ve `Drivers:` satırlarından sayılır)
6. **Motorlar** – motorlar ve bilgileri (plaka, durum, km, sigorta, muayene, son bakım)
7. **Brifing** – günün özeti ve Hello Rides hedefleri

Kumandanın sağ/sol tuşu sayfayı elle değiştirir. Adresin sonuna `?sayfa=hava` (hafta, ay, rakam, hava, ekip, motor, brif) yazılırsa o sayfa sabit kalır.

- Sayfa: `index.html` (tek dosya, kurulum gerektirmez). Veri yoktur; her şey TV kodu girilince Google'dan gelir.
- Sunucu: `sunucu/Code.gs` (Google Apps Script). "Tur Programi" takvimini, "TV Panel" tablosunu ve "TV Panel" Drive klasörünü yalnızca okur.
- Sunucu adresi: `ayar.js`.

Ayarlar, Drive › "TV Panel" tablosunda durur:

- **Hedefler** sayfası: hedef, durum, ilerleme %, not
- **Motorlar** sayfası: motor, plaka, durum, km, sigorta bitiş, muayene, son bakım, not (her motor bir satır)
- **Ayarlar** sayfası: isim, haftalık/aylık hedef, fotoğraf süresi, sayfa süresi, TV kodu

Arka plan fotoğrafları: Drive › "TV Panel" klasörü.

## Sunucuyu güncelleme

`sunucu/Code.gs` değiştiğinde: dosyanın tamamını Apps Script'teki Code.gs'e yapıştır, `kurulum` fonksiyonunu bir kez çalıştır (tabloya eksik sayfaları ekler, TV kodunu değiştirmez), sonra Dağıt › Dağıtımları yönet › kalem › Sürüm: Yeni sürüm › Dağıt. Adres aynı kalır.

Sayfa eski sunucuyla da çalışır; yalnızca verisi gelmeyen sayfalar (Ay, Sürücüler, Motorlar) görünmez.
