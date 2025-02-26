# Excel Nöbet Listesi Otomasyon Scripti

Bu proje, Python kullanarak belirtilen bir Excel dosyasındaki nöbet listesi verilerini işleyen ve her bir kişi için nöbet tarihleri, günleri ile çalıştığı alanları içeren görsel tablolar (PNG formatında) oluşturan bir script içerir. Ayrıca telefonlara kolayca yüklenebilen iCalendar (.ics) formatında takvim dosyaları oluşturarak nöbet hatırlatıcıları ayarlar.

## Özellikler

- **Excel Dosyası Okuma:** Belirtilen Excel dosyasındaki verileri okur.
- **Veri İşleme:** 
  - A sütununda tarih bilgisi,
  - B-J sütunlarında nöbetçi isimleri ve bu sütunların ilk satırında (B1:J1) alan başlıkları bulunur.
- **Koşullu Arama:** Her satırda, belirli bir ismi (ör. "BURAK CAN") arar; eşleşen satırlarda, o günün tarihini, Türkçe gün adını ve ilgili alan başlığını kaydeder.
- **Görsel Çıktı:** Her kişi için modern görünümlü tablolar oluşturur ve PNG dosyası olarak belirtilen klasöre kaydeder.
- **Takvim Hatırlatıcıları:** Her kişi için nöbet tarihinden bir gün önce saat 14:00'da hatırlatıcı içeren iCalendar (.ics) dosyaları oluşturur. Bu dosyalar, telefonlardaki takvim uygulamalarına kolayca aktarılabilir.
- **Üretim Hazır Kod:** `main.py` içerisinde yer alan script, argümanlar sayesinde farklı Excel dosyaları veya veri aralıkları ile esnek biçimde kullanılabilir.

## Gereksinimler

- Python 3.8 veya daha yeni bir sürüm
- Aşağıdaki Python kütüphaneleri:
  - pandas
  - matplotlib
  - openpyxl
  - icalendar
  - pytz

## Kütüphaneleri yüklemek için:

```bash
pip install -r requirements.txt
```

## requirements.txt Örneği:

```bash
pandas
matplotlib
openpyxl
icalendar
pytz
```

## Kurulum

Bu repository'i klonlayın:

```bash
git clone https://github.com/drcan94/excel-nobet-otomasyon.git
cd excel-nobet-otomasyon
```

## (Opsiyonel) Sanal ortam oluşturun:

```bash
python -m venv venv
source venv/bin/activate  # MacOS/Linux
.\venv\Scripts\activate   # Windows
```

## Gerekli kütüphaneleri yükleyin:

```bash
pip install -r requirements.txt
```

## Kullanım

Script'i çalıştırmak için:

```bash
python main.py --excel "MART-2025-NOBET.xlsx" --output "nobetler"
```

Takvim dosyalarını da oluşturmak için:

```bash
python main.py --excel "MART-2025-NOBET.xlsx" --output "nobetler" --create_calendar
```

## Varsayılan ayarlar:

- Excel dosya yolu: MART-2025-NOBET.xlsx
- Çıktı klasörü: nobetler
- İşlenecek satır aralığı: Pandas indeksinde 0-32 (Excel'de 2-33. satırlar)
- İsimlerin bulunduğu sütun aralığı: B (index 1) ile J (index 10)

## Diğer parametreler için:

- --start_row: İşlenecek başlangıç satırı (varsayılan: 0)
- --end_row: İşlenecek bitiş satırı (varsayılan: 32)
- --name_start_col: İsimlerin bulunduğu başlangıç sütunu (varsayılan: 1)
- --name_end_col: İsimlerin bulunduğu bitiş sütunu (varsayılan: 10)
- --create_calendar: Takvim dosyaları (.ics) oluştur

## Çalışma Mantığı

1- Excel Dosyası Okuma:
  - Belirtilen Excel dosyasındaki veriler pandas kullanılarak okunur.
  - İlk sütun (A) tarih, sonraki sütunlar (B-J) nöbetçi isimleri ve ilgili alan başlıklarıdır.

2- Veri İşleme:
  - Her satır için tarih, Türkçe gün adı ve ilgili sütun başlığı (alan) alınır.
  - Hücre değerleri temizlenip, parantez içindeki ek bilgilerin kaldırılması sağlanır.
  - Her geçerli isim için, kayıtlar bir sözlükte toplanır.

3- Görsel Çıktı Oluşturma:
  - Her kişi için matplotlib kullanılarak modern bir tablo oluşturulur ve PNG olarak çıktı klasörüne kaydedilir.

4- Takvim Dosyaları Oluşturma (Opsiyonel):
  - Her kişi için nöbet zamanlarını içeren iCalendar (.ics) dosyaları oluşturulur.
  - Her nöbet için, nöbet gününden bir gün önce saat 14:00'da tetiklenecek bir hatırlatıcı eklenir.
  - Dosyalar, "nobetler/calendar" klasörüne kaydedilir.

## Takvim Dosyalarını Telefonunuza Aktarmak

Oluşturulan .ics dosyaları, şu adımlarla telefonunuza aktarılabilir:

1. Oluşturulan .ics dosyasını telefonunuza gönderin (e-posta, mesaj, vb.)
2. Telefonunuzda dosyayı açın
3. Takvim uygulamanız dosyayı açacak ve etkinlikleri içe aktarmak isteyip istemediğinizi soracaktır
4. İçe aktarmayı onaylayın
5. Tüm nöbet etkinlikleri takvime eklenir ve otomatik olarak nöbetten bir gün önce saat 14:00'da hatırlatıcı ayarlanır

Desteklenen takvim uygulamaları:
- Google Takvim
- Apple Takvim
- Outlook
- Ve diğer iCalendar (.ics) desteği olan takvim uygulamaları

## Katkıda Bulunma

- Hata bildirmek veya öneride bulunmak için GitHub Issues bölümünü kullanabilirsiniz.
- Pull Request göndererek projeye katkıda bulunabilirsiniz.

## Lisans
Bu proje MIT Lisansı altında lisanslanmıştır.

İletişim
Herhangi bir sorunuz veya işbirliği öneriniz için "drcan94@gmail.com" üzerinden bana ulaşabilirsiniz.