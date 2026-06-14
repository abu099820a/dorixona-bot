"""
Kanalga filiallarni yuborish skripti
Ishlatish: python post_to_channel.py
"""

import pandas as pd
import requests
import time
import json
import os

TOKEN = "8837024109:AAGFZP5akA2nPo0RugVCCbEl2wgoe9N5_Uo"
CHANNEL_ID = "-1003087072308"
EXCEL_FILE = "dorixonalar.xlsx"
SENT_FILE = "sent_filials.json"  # Yuborilganlarni eslab qoladi

def load_sent():
    """Avval yuborilganlarni yuklash"""
    if os.path.exists(SENT_FILE):
        with open(SENT_FILE, "r") as f:
            return set(json.load(f))
    return set()

def save_sent(sent):
    """Yuborilganlarni saqlash"""
    with open(SENT_FILE, "w") as f:
        json.dump(list(sent), f)

def send_text(text, parse_mode="Markdown", reply_to=None, lat=None, lon=None):
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    payload = {
        "chat_id": CHANNEL_ID,
        "text": text,
        "parse_mode": parse_mode,
        "disable_web_page_preview": True
    }
    if reply_to:
        payload["reply_to_message_id"] = reply_to
    if lat and lon:
        payload["reply_markup"] = {
            "inline_keyboard": [[
                {"text": "🗺 Yandex Maps", "url": f"https://maps.yandex.ru/?pt={lon},{lat}&z=17&l=map"},
                {"text": "🗺 Google Maps", "url": f"https://maps.google.com/?q={lat},{lon}"}
            ]]
        }
    for attempt in range(3):
        r = requests.post(url, json=payload)
        data = r.json()
        if data.get("ok"):
            return data
        if "retry_after" in str(data):
            wait = data.get("parameters", {}).get("retry_after", 30)
            print(f"    ⏳ {wait} sekund kutilmoqda...")
            time.sleep(wait + 2)
        else:
            return data
    return data

def send_location(lat, lon):
    url = f"https://api.telegram.org/bot{TOKEN}/sendLocation"
    for attempt in range(3):
        r = requests.post(url, json={
            "chat_id": CHANNEL_ID,
            "latitude": float(lat),
            "longitude": float(lon),
        })
        data = r.json()
        if data.get("ok"):
            return data
        if "retry_after" in str(data):
            wait = data.get("parameters", {}).get("retry_after", 30)
            print(f"    ⏳ Lokatsiya: {wait} sekund kutilmoqda...")
            time.sleep(wait + 2)
        else:
            return data
    return data

def format_post(row):
    nomi = row.get("Nomi (RU)", "") or row.get("Nomi (UZ)", "")
    filial = str(row.get("filial_no", "")).strip()
    hudud = row.get("Hudud (RU)", "")
    tuman = row.get("Tuman (RU)", "")
    manzil = row.get("Manzil (RU)", "")
    orientir = row.get("Orientir (RU)", "")
    hours = row.get("Ish vaqti (RU)", "")
    phone = row.get("Telefon", "")
    lat = str(row.get("Latitude", "")).strip()
    lon = str(row.get("Longitude", "")).strip()

    lines = [f"🏥 *{nomi}*"]
    if filial and filial not in ["nan", ""]:
        if filial == "АСОСИЙ":
            lines.append(f"🔢 Asosiy filial")
        else:
            lines.append(f"🔢 Filial: #{filial}")
    if hudud or tuman:
        lines.append(f"🗺 {hudud}" + (f", {tuman}" if tuman else ""))
    if manzil: lines.append(f"📍 {manzil}")
    if orientir: lines.append(f"🚩 {orientir}")
    if hours: lines.append(f"🕐 {hours}")
    if phone:
        clean_phone = str(phone).replace(" ","").replace("-","").replace("(","").replace(")","")
        lines.append(f"📞 [{phone}](tel:{clean_phone})")

    has_coords = lat not in ["","nan"] and lon not in ["","nan"]
    return "\n".join(str(l) for l in lines), lat, lon, has_coords

def main():
    print("📂 Excel o'qilmoqda...")
    df = pd.read_excel(EXCEL_FILE).fillna("")
    df["filial_no"] = df["Filial №"].astype(str).str.strip().str.replace(r"\.0$","",regex=True)
    df["_sort"] = pd.to_numeric(df["filial_no"], errors="coerce").fillna(9999)
    df = df.sort_values("_sort").reset_index(drop=True)

    # TEST: sent_filials ni tozalash
    sent = set()
    save_sent(sent)
    print("🧹 sent_filials tozalandi — hammasi qaytadan yuboriladi")

    to_send = df
    print(f"📤 Jami: {len(to_send)} ta | Faqat 5 ta test yuboriladi")
    print()
    to_send = to_send.head(5)  # FAQAT 5 TA

    sent_count = 0

    for _, row in to_send.iterrows():
        filial_no = row["filial_no"]
        nomi = row.get("Nomi (RU)", "")
        text, lat, lon, has_coords = format_post(row.to_dict())

        print(f"  Yuborilmoqda: #{filial_no} {nomi[:30]}...")

        # Avval lokatsiya (agar bor bo'lsa)
        location_msg_id = None
        if has_coords:
            r = send_location(lat, lon)
            if not r.get("ok"):
                print(f"    ❌ Lokatsiya xatosi: {r.get('description','')}")
                print(f"    🛑 TO'XTATILDI! #{filial_no} dan davom ettiring.")
                break
            else:
                location_msg_id = r["result"]["message_id"]

        # Keyin matn — lokatsiyaga reply sifatida + xarita tugmalari
        _lat = lat if has_coords else None
        _lon = lon if has_coords else None
        r = send_text(text, reply_to=location_msg_id, lat=_lat, lon=_lon)
        if r.get("ok"):
            sent.add(filial_no)
            save_sent(sent)
            sent_count += 1
            print(f"    ✅ Yuborildi!")
        else:
            print(f"    ❌ Xato: {r.get('description','')}")
            print(f"    🛑 TO'XTATILDI! #{filial_no} dan davom ettiring.")
            break

        # Telegram limit
        time.sleep(3)

    print()
    print("=" * 40)
    print(f"✅ Yuborildi: {sent_count} ta")
    print(f"📁 {SENT_FILE} yangilandi")
    print("=" * 40)
    print()
    print("Davom ettirish uchun qaytadan ishga tushiring.")

if __name__ == "__main__":
    main()
