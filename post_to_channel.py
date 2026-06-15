"""
Kanalga filiallarni yuborish skripti
Ishlatish: python post_to_channel.py
"""

import pandas as pd
import requests
import time
import json
import os
import re

TOKEN = "8837024109:AAGFZP5akA2nPo0RugVCCbEl2wgoe9N5_Uo"
CHANNEL_ID = "-1003087072308"
EXCEL_FILE = "dorixonalar.xlsx"
SENT_FILE = "sent_filials.json"

OFFICE_LAT = 41.219104
OFFICE_LON = 69.272889
OFFICE_TEXT = """🏢 *Vaksin Med — Bosh ofis / Sklad*
📍 Toshkent shahri
🗺 Yandex va Google Maps orqali ko'ring"""

def load_sent():
    if os.path.exists(SENT_FILE):
        with open(SENT_FILE, "r") as f:
            return set(json.load(f))
    return set()

def save_sent(sent):
    with open(SENT_FILE, "w") as f:
        json.dump(list(sent), f)

def phone_to_tg(phone):
    digits = re.sub(r'\D', '', str(phone))
    if not digits:
        return None
    if digits.startswith("998"):
        number = digits
    elif digits.startswith("0"):
        number = "998" + digits[1:]
    else:
        number = "998" + digits
    return f"https://t.me/+{number}"

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

def send_text(text, reply_to=None, lat=None, lon=None, phone=None):
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    
    # Inline tugmalar
    buttons = []
    if lat and lon:
        buttons.append([
            {"text": "🗺 Yandex Maps", "url": f"https://maps.yandex.ru/?pt={lon},{lat}&z=17&l=map"},
            {"text": "🗺 Google Maps", "url": f"https://maps.google.com/?q={lat},{lon}"}
        ])
    if phone:
        tg_url = phone_to_tg(phone)
        if tg_url:
            buttons.append([{"text": "💬 Telegram", "url": tg_url}])

    payload = {
        "chat_id": CHANNEL_ID,
        "text": text,
        "parse_mode": "Markdown",
        "disable_web_page_preview": True
    }
    if reply_to:
        payload["reply_to_message_id"] = reply_to
    if buttons:
        payload["reply_markup"] = {"inline_keyboard": buttons}

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

def clean_md(text):
    """Markdown xato belgilarini tozalash"""
    special = ['`', '_', '[', ']', '~', '>', '#', '+', '=', '|', '{', '}', '!']
    t = str(text)
    for ch in special:
        t = t.replace(ch, '')
    return t

def format_post(row):
    nomi = row.get("Nomi (RU)", "") or row.get("Nomi (UZ)", "")
    filial = str(row.get("filial_no", "")).strip()
    hudud = clean_md(row.get("Hudud (RU)", ""))
    tuman = clean_md(row.get("Tuman (RU)", ""))
    manzil = clean_md(row.get("Manzil (RU)", ""))
    orientir = clean_md(row.get("Orientir (RU)", ""))
    hours = clean_md(row.get("Ish vaqti (RU)", ""))
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
    return "\n".join(str(l) for l in lines), lat, lon, has_coords, str(phone)

def send_one(filial_no, text, lat, lon, has_coords, phone, sent):
    """Bitta filialni yuborish: avval lokatsiya, keyin matn"""
    location_msg_id = None
    if has_coords:
        r = send_location(lat, lon)
        if not r.get("ok"):
            print(f"    ❌ Lokatsiya xatosi: {r.get('description','')}")
            return False
        location_msg_id = r["result"]["message_id"]
        time.sleep(1)

    r = send_text(
        text,
        reply_to=location_msg_id,
        lat=lat if has_coords else None,
        lon=lon if has_coords else None,
        phone=phone
    )
    if r.get("ok"):
        sent.add(filial_no)
        save_sent(sent)
        print(f"    ✅ Yuborildi!")
        return True
    else:
        print(f"    ❌ Xato: {r.get('description','')}")
        return False

def main():
    print("📂 Excel o'qilmoqda...")
    df = pd.read_excel(EXCEL_FILE).fillna("")
    df["filial_no"] = df["Filial №"].astype(str).str.strip().str.replace(r"\.0$","",regex=True)
    df["_sort"] = pd.to_numeric(df["filial_no"], errors="coerce").fillna(9999)
    df = df.sort_values("_sort").reset_index(drop=True)

    sent = load_sent()

    # --- 1. RAQAMLI FILIALLAR (1 dan oxirigacha) ---
    # ASOSIY ni oxiriga qoldiramiz
    raqamli = df[df["filial_no"].apply(lambda x: x.isdigit())]
    asosiy = df[df["filial_no"] == "АСОСИЙ"]

    sent_count = 0
    for _, row in raqamli.iterrows():
        filial_no = row["filial_no"]
        if filial_no in sent:
            print(f"  ⏭ #{filial_no} allaqachon yuborilgan")
            continue

        nomi = row.get("Nomi (RU)", "")
        print(f"  Yuborilmoqda: #{filial_no} {nomi[:30]}...")
        text, lat, lon, has_coords, phone = format_post(row.to_dict())

        ok = send_one(filial_no, text, lat, lon, has_coords, phone, sent)
        if ok:
            sent_count += 1
        else:
            print(f"    🛑 TO'XTATILDI! #{filial_no} dan davom ettiring.")
            break

        time.sleep(3)

    # --- 2. ASOSIY (oxirida) ---
    for _, row in asosiy.iterrows():
        filial_no = row["filial_no"]
        if filial_no in sent:
            print(f"  ⏭ Asosiy allaqachon yuborilgan")
            continue
        nomi = row.get("Nomi (RU)", "")
        print(f"  Yuborilmoqda: Asosiy filial...")
        text, lat, lon, has_coords, phone = format_post(row.to_dict())
        ok = send_one(filial_no, text, lat, lon, has_coords, phone, sent)
        if ok:
            sent_count += 1
        time.sleep(3)

    # --- 3. OFIS (eng oxirida) ---
    if "OFIS" not in sent:
        print("  Yuborilmoqda: Bosh ofis...")
        r = send_location(OFFICE_LAT, OFFICE_LON)
        if r.get("ok"):
            loc_id = r["result"]["message_id"]
            time.sleep(1)
            r2 = send_text(
                OFFICE_TEXT,
                reply_to=loc_id,
                lat=str(OFFICE_LAT),
                lon=str(OFFICE_LON)
            )
            if r2.get("ok"):
                sent.add("OFIS")
                save_sent(sent)
                sent_count += 1
                print("    ✅ Ofis yuborildi!")
            else:
                print(f"    ❌ Xato: {r2.get('description','')}")
        else:
            print(f"    ❌ Lokatsiya xatosi: {r.get('description','')}")
    else:
        print("  ⏭ Ofis allaqachon yuborilgan")

    print()
    print("=" * 40)
    print(f"✅ Jami yuborildi: {sent_count} ta")
    print("=" * 40)

if __name__ == "__main__":
    main()
