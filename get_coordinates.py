"""
Dorixonalar manzillaridan avtomatik koordinata olish (Yandex Geocoder)
Ishlatish: python get_coordinates.py
"""

import pandas as pd
import time
import requests

EXCEL_FILE = "dorixonalar.xlsx"
YANDEX_API_KEY = "0c14c63b-f533-492d-a442-e9f158eddc6f"

def get_coords(address):
    try:
        url = "https://geocode-maps.yandex.ru/1.x/"
        params = {
            "apikey": YANDEX_API_KEY,
            "geocode": address,
            "format": "json",
            "results": 1,
            "lang": "ru_RU",
        }
        r = requests.get(url, params=params, timeout=10)
        data = r.json()
        members = data["response"]["GeoObjectCollection"]["featureMember"]
        if members:
            pos = members[0]["GeoObject"]["Point"]["pos"]
            lon, lat = map(float, pos.split())
            return lat, lon
    except Exception as e:
        print(f"  Xato: {e}")
    return None, None

def main():
    print("📂 Excel fayl o'qilmoqda...")
    df = pd.read_excel(EXCEL_FILE)
    print(f"Jami {len(df)} ta dorixona topildi.")
    print()

    found = 0
    not_found = 0
    skipped = 0

    for idx, row in df.iterrows():
        lat = row.get("Latitude", None)
        lon = row.get("Longitude", None)
        if pd.notna(lat) and str(lat).strip() not in ["", "nan", "0"]:
            skipped += 1
            continue

        nomi = str(row.get("Nomi (RU)", "")).strip()
        manzil = str(row.get("Manzil (RU)", "")).strip()

        if not manzil or manzil == "nan":
            print(f"  [{idx+1}] {nomi[:40]} — manzil yo'q")
            not_found += 1
            continue

        print(f"  [{idx+1}] {nomi[:40]}")

        lat, lon = get_coords(manzil + ", Узбекистан")

        if lat and lon:
            df.at[idx, "Latitude"] = lat
            df.at[idx, "Longitude"] = lon
            print(f"       ✅ {lat:.4f}, {lon:.4f}")
            found += 1
        else:
            print(f"       ❌ Topilmadi")
            not_found += 1

        time.sleep(0.2)

    print()
    print("💾 Saqlanyapti...")
    df.to_excel(EXCEL_FILE, index=False)
    print()
    print("=" * 40)
    print(f"✅ Topildi:    {found} ta")
    print(f"⏭  O'tkazildi: {skipped} ta (allaqachon bor)")
    print(f"❌ Topilmadi:  {not_found} ta")
    print(f"💾 {EXCEL_FILE} yangilandi!")
    print("=" * 40)

if __name__ == "__main__":
    main()
