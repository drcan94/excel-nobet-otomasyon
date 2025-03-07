import os
import re
import argparse
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from icalendar import Calendar, Event, Alarm
import pytz


def turkish_day_name(date_obj):
    """
    Verilen datetime objesinden Türkçe gün adını döndürür.
    date_obj.weekday(): Pazartesi=0, Salı=1, Çarşamba=2, Perşembe=3, Cuma=4, Cumartesi=5, Pazar=6
    """
    gunler = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
    return gunler[date_obj.weekday()]


def slugify_name(name):
    """
    İsimdeki boşlukları tireye çevirir, küçük harfe dönüştürür.
    """
    return name.strip().lower().replace(" ", "-")


def process_excel(
    excel_file, start_row=0, end_row=32, name_start_col=1, name_end_col=10
):
    """
    Excel dosyasını işleyip, nöbet kayıtlarını içeren bir sözlük oluşturur.

    Parametreler:
      - excel_file: Excel dosyasının yolu.
      - start_row: İşlenecek başlangıç satırı indeksi (pandas indeks, varsayılan: 0)
      - end_row: İşlenecek bitiş satırı indeksi (varsayılan: 32, Excel'de 2-33. satırlar)
      - name_start_col: İsimlerin bulunduğu başlangıç sütunu indeksi (varsayılan: 1, B sütunu)
      - name_end_col: İsimlerin bulunduğu bitiş sütunu indeksi (varsayılan: 10, J sütunu)

    Dönüş:
      Sözlük, key olarak büyük harflerle isim, value olarak {"original": orijinal isim, "records": [[tarih - gün, alan], ...]} yapısında kayıtlar içerir.
    """
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        print(f"Excel dosyası okunamadı ({excel_file}): {e}")
        return {}

    # İlk sütun tarih, B-J sütunları ise nöbetçi isimleri ve alan başlıklarıdır.
    data_rows = df.iloc[start_row:end_row]
    schedule_dict = {}

    for idx, row in data_rows.iterrows():
        date_val = row.iloc[0]  # A sütunu: tarih
        if pd.notna(date_val):
            try:
                date_obj = pd.to_datetime(date_val)
            except Exception:
                continue

            date_str = date_obj.strftime("%d.%m.%Y")
            day_str = turkish_day_name(date_obj)
            combined_date = f"{date_str} - {day_str}"

            # İsimlerin bulunduğu sütunlar: B'den J'ye (indeks 1'den 10'a kadar)
            for col_idx in range(name_start_col, name_end_col):
                val = row.iloc[col_idx]
                if pd.notna(val):
                    # Sayı, tarih gibi türleri atla
                    if isinstance(val, (datetime, pd.Timestamp)) or isinstance(
                        val, (int, float)
                    ):
                        continue
                    name_raw = str(val).strip()
                    # Parantez içindeki ek bilgileri kaldır (örn. vardiya saatleri)
                    name_clean = re.split(r"\(", name_raw)[0].strip()
                    if not any(c.isalpha() for c in name_clean):
                        continue
                    # Sütun başlığı, ilgili alan bilgisi (ilk satırdaki başlıklar)
                    area_name = df.columns[col_idx]
                    record = [combined_date, area_name]

                    key = name_clean.upper()
                    if key not in schedule_dict:
                        schedule_dict[key] = {"original": name_clean, "records": []}
                    schedule_dict[key]["records"].append(record)
    return schedule_dict


def create_png_tables(schedule_dict, output_folder):
    """
    Her kişi için nöbet tablosunu oluşturur ve PNG olarak kaydeder.

    Parametreler:
      - schedule_dict: İşlenmiş nöbet kayıtlarını içeren sözlük.
      - output_folder: PNG dosyalarının kaydedileceği klasör.
    """
    os.makedirs(output_folder, exist_ok=True)

    for person_key, info in schedule_dict.items():
        original_name = info["original"]
        records = info["records"]

        # Tarih bilgisine göre sıralama (kayıtların ilk elemanı: "dd.mm.yyyy - Gün")
        def sort_key(item):
            try:
                return datetime.strptime(item[0].split(" - ")[0], "%d.%m.%Y")
            except Exception:
                return datetime.min

        records.sort(key=sort_key)
        if not records:
            records = [["Kayıt Yok", ""]]

        # Matplotlib ile tablo oluşturma
        fig, ax = plt.subplots(figsize=(12, len(records) * 0.6 + 1))
        ax.axis("off")
        plt.text(
            0.5,
            0.95,
            original_name,
            transform=fig.transFigure,
            fontsize=18,
            fontweight="bold",
            ha="center",
            va="top",
        )

        columns = ["Tarih", "Alan"]
        table = ax.table(
            cellText=records,
            colLabels=columns,
            cellLoc="center",
            loc="center",
            edges="closed",
        )
        table.auto_set_font_size(False)
        table.set_fontsize(14)
        table.scale(1.5, 1.5)

        for (r, c), cell in table.get_celld().items():
            cell.set_linewidth(1)
            cell.set_edgecolor("gray")
            cell.set_text_props(fontsize=14, ha="center")
            if r == 0:
                cell.set_text_props(weight="bold", color="white")
                cell.set_facecolor("#4F81BD")
            else:
                cell.set_facecolor("white")

        filename = slugify_name(original_name) + ".png"
        filepath = os.path.join(output_folder, filename)

        try:
            plt.savefig(
                filepath,
                dpi=300,
                bbox_inches="tight",
                pad_inches=0.2,
                facecolor="white",
            )
            print(f"{original_name} için {filepath} oluşturuldu.")
        except Exception as e:
            print(f"Hata oluştu: {e} -> {filepath}")
        finally:
            plt.close()


def create_ics_files(schedule_dict, output_folder):
    """
    Her kişi için nöbet hatırlatıcılarını içeren iCalendar (.ics) dosyaları oluşturur.

    Parametreler:
      - schedule_dict: İşlenmiş nöbet kayıtlarını içeren sözlük.
      - output_folder: ICS dosyalarının kaydedileceği klasör.
    """
    # Çıktı klasörü içinde "calendar" alt klasörü oluştur
    calendar_folder = os.path.join(output_folder, "calendar")
    os.makedirs(calendar_folder, exist_ok=True)

    # Türkiye zaman dilimi
    tz = pytz.timezone("Europe/Istanbul")

    for person_key, info in schedule_dict.items():
        original_name = info["original"]
        records = info["records"]

        # Takvim nesnesi oluştur
        cal = Calendar()
        cal.add("prodid", "-//Nöbet Takvimi//drcan94.github.io//")
        cal.add("version", "2.0")
        cal.add("x-wr-calname", f"{original_name} Nöbet Takvimi")

        # Kayıtları tarihe göre sırala
        def sort_key(item):
            try:
                return datetime.strptime(item[0].split(" - ")[0], "%d.%m.%Y")
            except Exception:
                return datetime.min

        records.sort(key=sort_key)

        # Her nöbet için etkinlik oluştur
        for record in records:
            try:
                date_str = record[0].split(" - ")[0]  # "dd.mm.yyyy" formatı
                area_name = record[1]

                # Nöbet tarihi
                duty_date = datetime.strptime(date_str, "%d.%m.%Y")

                # Nöbet başlangıç saati (sabah 08:00 olarak ayarla)
                duty_start = duty_date.replace(hour=8, minute=0, second=0)
                # Nöbet süresi 24 saat (ertesi gün aynı saatte bitiş)
                duty_end = duty_start + timedelta(days=1)

                # Etkinlik oluştur
                event = Event()
                event.add("summary", f"NÖBET: {area_name}")
                event.add(
                    "description", f"{original_name} için {area_name} alanında nöbet"
                )

                # Normal etkinlik (tüm gün değil, 24 saatlik)
                event.add("dtstart", duty_start)
                event.add("dtend", duty_end)

                # Benzersiz ID oluştur
                event_id = f"{slugify_name(original_name)}-{date_str}-{slugify_name(area_name)}"
                event.add("uid", event_id)

                # Bir gün önce saat 14:00'da hatırlatıcı ekle
                alarm = Alarm()
                alarm.add("action", "DISPLAY")
                alarm.add(
                    "description", f"HATIRLATMA: Yarın {area_name} alanda nöbetin var!"
                )

                # Hatırlatıcı zamanını hesapla: nöbetin bir gün öncesi saat 14:00
                reminder_date = duty_start - timedelta(days=1)
                reminder_time = reminder_date.replace(hour=14, minute=0, second=0)

                # Hatırlatıcı ayarları (hatırlatıcı zamanı ile etkinliğin başlangıç zamanı arasındaki fark)
                trigger_delta = reminder_time - duty_start
                alarm.add("trigger", trigger_delta)

                # Alarmı etkinliğe ekle
                event.add_component(alarm)

                # Etkinliği takvime ekle
                cal.add_component(event)

            except Exception as e:
                print(f"Takvim etkinliği oluştururken hata: {e} -> {record}")

        # Takvim dosyasını kaydet
        if len(cal.subcomponents) > 0:  # Etkinlik varsa
            filename = slugify_name(original_name) + ".ics"
            filepath = os.path.join(calendar_folder, filename)

            with open(filepath, "wb") as f:
                f.write(cal.to_ical())
            print(f"{original_name} için takvim dosyası oluşturuldu: {filepath}")
            print(
                f"  * Nöbet etkinliği 08:00'dan başlayıp 24 saat sürecek şekilde ayarlandı."
            )
            print(
                f"  * Hatırlatıcılar nöbetten 1 gün önce saat 14:00'da çalacak şekilde ayarlandı."
            )


def main():
    parser = argparse.ArgumentParser(
        description="Excel nöbet listesi scripti. Belirtilen Excel dosyasından nöbet bilgilerini okuyarak her kişi için PNG dosyası oluşturur."
    )
    parser.add_argument(
        "--excel",
        type=str,
        default="MART-2025-NOBET.xlsx",
        help="Excel dosyasının yolu (varsayılan: MART-2025-NOBET.xlsx)",
    )
    parser.add_argument(
        "--output",
        type=str,
        default="nobetler",
        help="Çıktı PNG dosyalarının kaydedileceği klasör (varsayılan: nobetler)",
    )
    parser.add_argument(
        "--start_row",
        type=int,
        default=0,
        help="İşlenecek başlangıç satırı indeksi (pandas indeksi, varsayılan: 0)",
    )
    parser.add_argument(
        "--end_row",
        type=int,
        default=32,
        help="İşlenecek bitiş satırı indeksi (varsayılan: 32, Excel'de 2-33. satırlar)",
    )
    parser.add_argument(
        "--name_start_col",
        type=int,
        default=1,
        help="İsimlerin bulunduğu başlangıç sütunu indeksi (varsayılan: 1, B sütunu)",
    )
    parser.add_argument(
        "--name_end_col",
        type=int,
        default=10,
        help="İsimlerin bulunduğu bitiş sütunu indeksi (varsayılan: 10, J sütunu)",
    )
    parser.add_argument(
        "--create_calendar",
        action="store_true",
        help="Nöbet hatırlatıcıları için iCalendar (.ics) dosyaları oluştur",
    )
    args = parser.parse_args()

    schedule_dict = process_excel(
        args.excel, args.start_row, args.end_row, args.name_start_col, args.name_end_col
    )
    if not schedule_dict:
        print("Herhangi bir kayıt bulunamadı.")
    else:
        create_png_tables(schedule_dict, args.output)
        print(f"\nTüm PNG dosyaları '{args.output}' klasöründe oluşturuldu.")

        # Takvim dosyaları oluştur
        if args.create_calendar:
            create_ics_files(schedule_dict, args.output)
            cal_folder = os.path.join(args.output, "calendar")
            print(f"\nTüm takvim dosyaları '{cal_folder}' klasöründe oluşturuldu.")
            print("\nKullanım: Bu .ics dosyaları telefonunuzun takvim uygulamasına")
            print("yüklenebilir (Google Takvim, Apple Takvim, Outlook vb.). ")
            print(
                "Böylece her nöbetten bir gün önce saat 14:00'da hatırlatıcı alacaksınız."
            )


if __name__ == "__main__":
    main()
