"""
register_handlers.py — Ro'yxatdan o'tish moduli

TARTIB:
1. "📝 Ro'yxatdan o'tish" → telefon so'raladi
2. Telefon yuboriladi
   A) Sheets da bor + TelegramID bo'sh  → ID saqlaydi → tugaydi
   B) Sheets da bor + TelegramID to'lgan → "Allaqachon ro'yxatdansiz"
   C) Sheets da yo'q → Ismi → Filial raqami → avtomatik Lat/Lon/Lavozim
      → Shu filialning oxirgi xodimidan KEYIN qo'shiladi
      → TelegramID saqlanadi → Davomat jadvaliga ham qo'shiladi

Sheets ustunlari:
A=Filial | B=Ismi | C=Telefon | D=TelegramID | E=Lavozim | F=Lat | G=Lon
"""

import os
import re
import json
from telegram import (
    Update, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove,
)
from telegram.ext import (
    MessageHandler, ContextTypes, filters,
)
from google.oauth2.service_account import Credentials
import gspread
from attendance import run_read, run_write

# ─── Sozlamalar ───────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
PHARMACY_SHEET_ID   = os.getenv("PHARMACY_SHEET_ID", "")
ATTENDANCE_SHEET_ID = os.getenv("ATTENDANCE_SHEET_ID", "")
SALARY_SHEET_ID     = os.getenv("SALARY_SHEET_ID", "")
FILIALLAR_SHEET_ID  = os.getenv("FILIALLAR_SHEET_ID", "")
ADMIN_IDS = [709544046]

# Conversation states
REG_TYPE    = 399  # Dorixona / Firma tanlash
REG_PHONE   = 400
REG_NAME    = 401
REG_FILIAL  = 402
REG_LAVOZIM = 403
FIRM2_PHONE = 407
FIRM2_NAME  = 408

# Ustun raqamlari (1-indexed)
COL_FILIAL    = 1  # A
COL_ISMI      = 2  # B
COL_TELEFON   = 3  # C
COL_TELEGRAMID= 4  # D
COL_LAVOZIM   = 5  # E
COL_LAT       = 6  # F
COL_LON       = 7  # G

# ─── Yordamchi ───────────────────────────────────────────────────────────────

_CACHED_CLIENT = None

def _get_client():
    global _CACHED_CLIENT
    if _CACHED_CLIENT is not None:
        return _CACHED_CLIENT
    creds_json = os.getenv("GOOGLE_CREDENTIALS")
    if creds_json:
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    _CACHED_CLIENT = gspread.authorize(creds)
    return _CACHED_CLIENT


def normalize_phone(phone: str) -> str:
    digits = re.sub(r"\D", "", str(phone))
    if digits.startswith("998"):   return "+" + digits
    if digits.startswith("0"):     return "+998" + digits[1:]
    if len(digits) == 9:           return "+998" + digits
    return "+" + digits


def _filial_kod(filial: str) -> str:
    """'6 - ЮНУСАБАД 7' → '6'"""
    m = re.match(r"^(\d+)", str(filial).strip())
    return m.group(1) if m else ""


# ─── Google Sheets operatsiyalari ────────────────────────────────────────────

def find_by_phone(phone: str) -> dict | None:
    """Telefon raqami bo'yicha farmatsevtni topadi."""
    try:
        client = _get_client()
        ws = client.open_by_key(PHARMACY_SHEET_ID).worksheet("Farmatsevtlar")
        records = ws.get_all_records()
        norm = normalize_phone(phone)
        for i, row in enumerate(records):
            tel = str(row.get("Telefon", ""))
            if isinstance(tel, float):
                tel = str(int(float(tel)))
            if normalize_phone(tel) == norm:
                return {
                    "row_num":   i + 2,
                    "ismi":      str(row.get("Ismi", "")).strip(),
                    "filial":    str(row.get("Filial", "")).strip(),
                    "telefon":   tel,
                    "telegramid":str(row.get("TelegramID", "")).strip(),
                    "lavozim":   str(row.get("Lavozim", "")).strip(),
                    "lat":       str(row.get("Lat", "")).strip(),
                    "lon":       str(row.get("Lon", "")).strip(),
                }
        return None
    except Exception as e:
        print(f"[REG] Telefon qidirish xato: {e}")
        return None


def save_telegram_id(row_num: int, user_id: int) -> bool:
    """Mavjud xodimning TelegramID sini D ustuniga saqlaydi."""
    try:
        client = _get_client()
        ws = client.open_by_key(PHARMACY_SHEET_ID).worksheet("Farmatsevtlar")
        ws.update_cell(row_num, COL_TELEGRAMID, str(user_id))
        return True
    except Exception as e:
        print(f"[REG] TelegramID saqlash xato: {e}")
        return False


def get_filial_info(filial_kod: str) -> dict | None:
    """
    1. Filiallar Sheets dan (FILIALLAR_SHEET_ID) Lat/Lon oladi
       A=Filial №, M=Latitude, N=Longitude
    2. Farmatsevtlar Sheets dan filial nomini va oxirgi qatorni topadi
    """
    try:
        client = _get_client()

        # ── 1. Filiallar Sheets dan Lat/Lon olish ──
        lat = ""
        lon = ""
        filial_nomi_from_ph = None

        if FILIALLAR_SHEET_ID:
            try:
                fil_ws = client.open_by_key(FILIALLAR_SHEET_ID).sheet1
                fil_values = fil_ws.get_all_values()

                # Sarlavhadan ustun indekslarini topish
                if fil_values:
                    headers = [h.strip() for h in fil_values[0]]
                    # Filial № ustuni
                    try:
                        fil_no_idx = headers.index("Filial №")
                    except ValueError:
                        fil_no_idx = 0  # A ustun
                    # Latitude ustuni
                    try:
                        lat_idx = headers.index("Latitude")
                    except ValueError:
                        lat_idx = 12  # M ustun (0-indexed)
                    # Longitude ustuni
                    try:
                        lon_idx = headers.index("Longitude")
                    except ValueError:
                        lon_idx = 13  # N ustun (0-indexed)

                    for row in fil_values[1:]:
                        if not row or not row[fil_no_idx]:
                            continue
                        fil_no = str(row[fil_no_idx]).strip()
                        # "асосий" → "0", raqamli → raqam
                        if fil_no.lower() in ("асосий", "asosiy"):
                            fil_no = "0"
                        if fil_no == filial_kod.strip():
                            if len(row) > lat_idx:
                                v = str(row[lat_idx]).strip()
                                if v and v not in ("", "0", "nan"):
                                    lat = v
                            if len(row) > lon_idx:
                                v = str(row[lon_idx]).strip()
                                if v and v not in ("", "0", "nan"):
                                    lon = v
                            break
            except Exception as e:
                print(f"[REG] Filiallar Sheets xato: {e}")

        # ── 2. Farmatsevtlar Sheets dan filial nomi va last_row topish ──
        ph_ws = client.open_by_key(PHARMACY_SHEET_ID).worksheet("Farmatsevtlar")
        all_values = ph_ws.get_all_values()

        filial_nomi = None
        last_row = 1

        for i, row in enumerate(all_values):
            if i == 0:
                continue
            if not row or not row[0]:
                continue
            filial_cell = str(row[0]).strip()
            if _filial_kod(filial_cell) == filial_kod.strip():
                filial_nomi = filial_cell
                last_row = i + 1  # 1-indexed
                # Agar Filiallar Sheets da Lat/Lon topilmagan bo'lsa
                # — Farmatsevtlar Sheets dan ham qidiramiz
                if not lat and len(row) > 5:
                    v = str(row[5]).strip()
                    if v and v not in ("", "0", "nan"):
                        lat = v
                if not lon and len(row) > 6:
                    v = str(row[6]).strip()
                    if v and v not in ("", "0", "nan"):
                        lon = v

        if filial_nomi is None:
            return None

        print(f"[REG] Filial: {filial_nomi} | Lat={lat} | Lon={lon} | last_row={last_row}")
        return {
            "filial_nomi": filial_nomi,
            "lat":         lat,
            "lon":         lon,
            "last_row":    last_row,
        }
    except Exception as e:
        print(f"[REG] Filial info xato: {e}")
        return None


def _find_numeric_insert_position(rows: list, filial_kod: str) -> int:
    """
    Filial guruhi hali mavjud bo'lmagan holatda, uni RAQAMLI tartibda
    to'g'ri o'rniga (masalan 13 va 15 orasiga, agar 14 yetishmasa)
    joylashtirish uchun eng mos qator INDEKSINI (0-based, `rows`
    ro'yxati ichida) qaytaradi. Agar raqamlarni solishtirib bo'lmasa
    (masalan kod raqam emas) — ro'yxat oxirini qaytaradi.
    """
    try:
        target_num = int(filial_kod)
    except (ValueError, TypeError):
        return len(rows)

    last_smaller_end = 0
    for i, row in enumerate(rows):
        if not row or not row[0]:
            continue
        kod = _filial_kod(str(row[0]).strip())
        if not kod:
            continue
        try:
            num = int(kod)
        except ValueError:
            continue
        if num < target_num:
            last_smaller_end = i + 1
        elif num > target_num:
            return last_smaller_end
    return last_smaller_end


def _lavozim_priority(lavozim: str) -> int:
    """
    Xodim tartibini aniqlaydi: Dorixona mudiri (1) → Farmatsevt (2) →
    Stajyor (3). Har bir filial guruhi ichida xodimlar shu tartibda
    turishi kerak.
    """
    l = str(lavozim).strip().lower()
    if "mudir" in l:
        return 1
    if "stajyor" in l or "стажер" in l or "стажёр" in l:
        return 3
    return 2  # Farmatsevt va aniqlanmagan lavozimlar — o'rtada


def _get_group_lavozim_map(filial_kod: str) -> dict:
    """
    Farmatsevtlar jadvalidan berilgan filial guruhidagi xodimlarning
    Ismi -> Lavozim lug'atini qaytaradi (Davomat/Oylik/Aksiya
    varaqlarida tartibni aniqlash uchun ishlatiladi, chunki ular
    Lavozim ustunini o'z ichiga olmaydi).
    """
    try:
        client = _get_client()
        ws = client.open_by_key(PHARMACY_SHEET_ID).worksheet("Farmatsevtlar")
        all_values = ws.get_all_values()
        result = {}
        for row in all_values[1:]:
            if not row or not row[0]:
                continue
            if _filial_kod(str(row[0]).strip()) != filial_kod:
                continue
            ismi = str(row[1]).strip().upper() if len(row) > 1 else ""
            lavozim = str(row[4]).strip() if len(row) > 4 else ""
            if ismi:
                result[ismi] = lavozim
        return result
    except Exception as e:
        print(f"[REG] _get_group_lavozim_map xato: {e}")
        return {}


def add_new_farmatsevt(
    ismi: str, telefon: str, filial_nomi: str,
    lavozim: str, lat: str, lon: str,
    after_row: int, user_id: int
) -> bool:
    """
    Yangi farmatsevtni Sheets ga qo'shadi.
    MUHIM: after_row parametri endi faqat ZAXIRA sifatida ishlatiladi
    (agar biror sabab bilan tartiblab joylashtirish muvaffaqiyatsiz
    bo'lsa). Asosiy holatda, filial guruhi ICHIDA to'g'ri tartib bilan
    (Dorixona mudiri → Farmatsevt → Stajyor) joylashtiriladi — yangi
    xodim o'z lavozimiga mos joyga, tegishli guruhning oxiriga qo'shiladi.
    Davomat va Oylik/Aksiya Sheets ga ham avtomatik, xuddi shu tartibda
    qo'shiladi.
    """
    try:
        client = _get_client()
        ws = client.open_by_key(PHARMACY_SHEET_ID).worksheet("Farmatsevtlar")

        filial_kod = _filial_kod(filial_nomi)
        new_priority = _lavozim_priority(lavozim)

        insert_row = after_row + 1  # zaxira (agar pastdagi hisoblash ishlamasa)
        try:
            all_values = ws.get_all_values()
            target_row = None
            for i, row in enumerate(all_values):
                if i == 0 or not row or not row[0]:
                    continue
                if _filial_kod(str(row[0]).strip()) != filial_kod:
                    continue
                row_lavozim = str(row[4]).strip() if len(row) > 4 else ""
                row_priority = _lavozim_priority(row_lavozim)
                if row_priority <= new_priority:
                    target_row = i + 1  # shu qatordan keyin qo'shamiz (hozircha)
            if target_row is not None:
                insert_row = target_row + 1
        except Exception as e:
            print(f"[REG] Tartib hisoblashda xato, zaxira joy ishlatiladi: {e}")

        # Qator qo'shish (pastki qatorlarni surish)
        ws.insert_row(
            [filial_nomi, ismi, telefon, str(user_id), lavozim, lat, lon],
            index=insert_row,
            value_input_option="USER_ENTERED"
        )

        print(f"[REG] Yangi farmatsevt qo'shildi: {ismi} | {filial_nomi} | {lavozim} | qator {insert_row}")

        # Davomat Sheets ga ham qo'shish (tartib bilan)
        _add_to_attendance(ismi, filial_nomi, telefon, lavozim)

        # "Oylik" (Maosh) Sheets ga ham qo'shish (tartib bilan)
        _add_to_oylik(ismi, filial_nomi, telefon, lavozim)

        return True
    except Exception as e:
        print(f"[REG] Yangi farmatsevt qo'shish xato: {e}")
        return False


def _add_to_attendance(ismi: str, filial_nomi: str, telefon: str = "", lavozim: str = ""):
    """
    Davomat Sheets dagi joriy oy listiga yangi farmatsevtni qo'shadi.
    Jadval: A=Filial, B=Ismi
    Filial guruhi ICHIDA tartib bilan qo'shadi: Dorixona mudiri →
    Farmatsevt → Stajyor. Davomat jadvalining o'zida Lavozim ustuni
    bo'lmagani uchun, guruhdagi mavjud xodimlarning lavozimi
    Farmatsevtlar jadvalidan (Ismi orqali) aniqlanadi.
    """
    try:
        if not ATTENDANCE_SHEET_ID:
            return

        from datetime import datetime, timezone, timedelta

        UZ_TZ = timezone(timedelta(hours=5))
        now = datetime.now(UZ_TZ)

        OY_NOMLARI = {
            1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
            5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
            9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
        }
        sheet_name = f"{OY_NOMLARI[now.month]} {now.year}"

        client = _get_client()
        att_sh = client.open_by_key(ATTENDANCE_SHEET_ID)

        try:
            ws = att_sh.worksheet(sheet_name)
        except Exception:
            print(f"[REG] Davomat listi topilmadi: {sheet_name}")
            return

        all_values = ws.get_all_values()
        filial_kod = _filial_kod(filial_nomi)
        new_priority = _lavozim_priority(lavozim)
        lavozim_map = _get_group_lavozim_map(filial_kod)

        last_row = 2  # default
        target_row = None
        found_any = False

        # Jadval: A=Filial, B=Ismi
        for i, row in enumerate(all_values):
            if i < 2:
                continue
            if not row or not row[0]:
                continue
            row_filial = str(row[0]).strip()
            if _filial_kod(row_filial) != filial_kod:
                continue
            found_any = True
            last_row = i + 1  # zaxira (guruhning oxirgi qatori)

            row_ismi = str(row[1]).strip().upper() if len(row) > 1 else ""
            row_lavozim = lavozim_map.get(row_ismi, "")
            row_priority = _lavozim_priority(row_lavozim) if row_lavozim else 2
            if row_priority <= new_priority:
                target_row = i + 1

        if not found_any:
            # MUHIM: bu filial guruhi Davomat jadvalida UMUMAN topilmadi —
            # jadval oxiriga emas, RAQAMLI TARTIBDA to'g'ri o'rniga
            # (masalan 13 va 15 orasiga, agar 14 yetishmasa) qo'shamiz.
            pos = _find_numeric_insert_position(all_values[2:], filial_kod)
            insert_row = 2 + pos + 1
        else:
            insert_row = (target_row if target_row is not None else last_row) + 1
        # A=Filial, B=Ismi, C=Telefon tartibida qo'shish
        ws.insert_row([filial_nomi, ismi, telefon], index=insert_row)
        print(f"[REG] Davomat ga qo'shildi: {filial_nomi} | {ismi} | {lavozim} | qator {insert_row}")

        # Yangi qatorni oq rangga qaytarish (filial sarlavha rangini olmasa)
        try:
            att_sh_local = client.open_by_key(ATTENDANCE_SHEET_ID)
            ws_sheet = att_sh_local.worksheet(sheet_name)
            att_sh_local.batch_update({"requests": [{
                "repeatCell": {
                    "range": {
                        "sheetId": ws_sheet.id,
                        "startRowIndex": insert_row - 1,
                        "endRowIndex": insert_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": 3,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 1, "green": 1, "blue": 1},
                            "textFormat": {
                                "bold": False,
                                "foregroundColor": {"red": 0, "green": 0, "blue": 0}
                            },
                            "horizontalAlignment": "LEFT",
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
                }
            }]})
        except Exception as re:
            print(f"[REG] Rang tuzatish xato: {re}")

    except Exception as e:
        print(f"[REG] Davomat yangilash xato: {e}")


def _merge_filial_column(ws, first_row: int, last_row: int):
    """
    A ustunidagi [first_row, last_row] oralig'ini (1-indeksda) bitta
    katakka birlashtiradi (merge) — filial nomi guruh bo'ylab bitta
    katakda ko'rinishi uchun (Davomat jadvalidagi kabi). Avval o'sha
    diapazonni "unmerge" qilib, keyin qayta merge qiladi — bu eski
    merge holati qanday bo'lishidan qat'i nazar to'g'ri natija beradi.
    """
    if last_row <= first_row:
        return
    try:
        requests = [
            {"unmergeCells": {"range": {
                "sheetId": ws.id,
                "startRowIndex": first_row - 1, "endRowIndex": last_row,
                "startColumnIndex": 0, "endColumnIndex": 1,
            }}},
            {"mergeCells": {"range": {
                "sheetId": ws.id,
                "startRowIndex": first_row - 1, "endRowIndex": last_row,
                "startColumnIndex": 0, "endColumnIndex": 1,
            }, "mergeType": "MERGE_ALL"}},
        ]
        ws.spreadsheet.batch_update({"requests": requests})
    except Exception as e:
        print(f"[REG] Merge xato: {e}")


def _add_to_salary_sheet(ws_name: str, ismi: str, filial_nomi: str, telefon: str = "", lavozim: str = ""):
    """
    "Oylik" yoki "Aksiya" Sheets'ga (SALARY_SHEET_ID) yangi xodimni
    qo'shadi. Jadval endi Davomat bilan bir xil formatda:
        A=Filial (raqam bilan, masalan "1 - ТАШМИ-1") | B=Ismi | C=Telefon
    Filial guruhi ICHIDA tartib bilan qo'shiladi: Dorixona mudiri →
    Farmatsevt → Stajyor (Lavozim ma'lumoti Farmatsevtlar jadvalidan
    Ismi orqali aniqlanadi, chunki bu varaqda Lavozim ustuni yo'q).
    """
    try:
        if not SALARY_SHEET_ID:
            return
        client = _get_client()
        sh = client.open_by_key(SALARY_SHEET_ID)

        try:
            ws = sh.worksheet(ws_name)
        except Exception:
            print(f"[REG] '{ws_name}' varag'i topilmadi — o'tkazib yuborildi")
            return

        all_values = ws.get_all_values()
        filial_kod = _filial_kod(filial_nomi)
        new_priority = _lavozim_priority(lavozim)
        lavozim_map = _get_group_lavozim_map(filial_kod)

        last_row = 1  # sarlavha qatoridan keyin, default
        target_row = None
        found_any = False

        # MUHIM: A ustuni bo'sh bo'lgan qatorlarni ham hisobga olamiz
        # (masalan ilgari merge qilingan katakchalar tufayli) — yuqoridagi
        # oxirgi ko'rilgan filial nomini "forward-fill" qilib davom
        # ettiramiz, aks holda oxirgi qatorni noto'g'ri aniqlab, xodimni
        # boshqa filial guruhiga qo'shib qo'yishimiz mumkin.
        effective_filial = ""
        for i, row in enumerate(all_values):
            if i == 0:
                continue  # sarlavha qatori
            if row and row[0]:
                effective_filial = str(row[0]).strip()
            if not row or (len(row) < 2 or not row[1]):
                continue  # bo'sh qator yoki Ismi yo'q — xodim emas
            if _filial_kod(effective_filial) != filial_kod:
                continue
            found_any = True
            last_row = i + 1

            row_ismi = str(row[1]).strip().upper()
            row_lavozim = lavozim_map.get(row_ismi, "")
            row_priority = _lavozim_priority(row_lavozim) if row_lavozim else 2
            if row_priority <= new_priority:
                target_row = i + 1

        if not found_any:
            # MUHIM: bu filial guruhi jadvalda UMUMAN topilmadi (masalan
            # hali hech kim shu filialdan ro'yxatdan o'tmagan). Bunday
            # holda RAQAMLI TARTIBDA to'g'ri o'rniga (masalan 13 va 15
            # orasiga, agar 14 yetishmasa) qo'shamiz.
            pos = _find_numeric_insert_position(all_values[1:], filial_kod)
            insert_row = 1 + pos + 1
        else:
            insert_row = (target_row if target_row is not None else last_row) + 1
        ws.insert_row([filial_nomi, ismi, telefon], index=insert_row, value_input_option="USER_ENTERED")
        print(f"[REG] {ws_name} ga qo'shildi: {filial_nomi} | {ismi} | {lavozim} | qator {insert_row}")

        # MUHIM: endi merge QILMAYMIZ — Google Sheets merge qilinganda
        # guruhning birinchi qatoridan boshqa barcha qatorlardagi qiymatni
        # HAQIQATAN o'chirib tashlaydi (API orqali o'qiganda ham bo'sh
        # chiqadi) — bu boshqa xodimlarning "yo'qolib qolishiga" olib
        # kelgan edi.

    except Exception as e:
        print(f"[REG] {ws_name} yangilash xato: {e}")


def _add_to_oylik(ismi: str, filial_nomi: str, telefon: str = "", lavozim: str = ""):
    """Xodimni "Oylik" va "Aksiya" varaqlarining ikkalasiga ham, tartib bilan qo'shadi."""
    _add_to_salary_sheet("Oylik", ismi, filial_nomi, telefon, lavozim)
    _add_to_salary_sheet("Aksiya", ismi, filial_nomi, telefon, lavozim)


# ─── Klaviaturalar ────────────────────────────────────────────────────────────

def _phone_kb():
    return ReplyKeyboardMarkup([
        [KeyboardButton("📱 Telefon raqamimni yuborish", request_contact=True)],
        ["⬅️ Orqaga"],
    ], resize_keyboard=True)


def _back_kb():
    return ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True)


def _lavozim_kb():
    return ReplyKeyboardMarkup([
        ["👔 Farmatsevt"],
        ["👔 Dorixona mudiri"],
        ["👔 Stajyor"],
        ["⬅️ Orqaga"],
    ], resize_keyboard=True)


# ─── Handlerlar ───────────────────────────────────────────────────────────────

async def register_enter(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """'📝 Ro'yxatdan o'tish' bosilganda — turini so'raydi."""
    await update.message.reply_text(
        "📝 *Ro'yxatdan o'tish*\n\nKimsiz?",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup([
            ["💊 Dorixonadan ro'yxatdan o'tish"],
            ["🏢 Firma uchun ro'yxatdan o'tish"],
            ["⬅️ Orqaga"],
        ], resize_keyboard=True),
    )
    return REG_TYPE


async def reg_type_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Ro'yxatdan o'tish turi tanlanadi: Dorixona yoki Firma."""
    txt = update.message.text.strip() if update.message.text else ""

    if txt == "⬅️ Orqaga":
        from bot import main_keyboard, get_lang, MENU
        await update.message.reply_text("📋 Asosiy menyu", reply_markup=main_keyboard(get_lang(ctx)))
        return MENU

    if txt == "💊 Dorixonadan ro'yxatdan o'tish":
        ctx.user_data.pop("reg_phone", None)
        ctx.user_data.pop("reg_ismi", None)
        ctx.user_data.pop("reg_filial_info", None)
        await update.message.reply_text(
            "📝 *Dorixonadan ro'yxatdan o'tish*\n\n"
            "📱 Telefon raqamingizni yuboring:",
            parse_mode="Markdown",
            reply_markup=_phone_kb(),
        )
        return REG_PHONE

    if txt == "🏢 Firma uchun ro'yxatdan o'tish":
        await update.message.reply_text(
            "🏢 *Firma uchun ro'yxatdan o'tish*\n\n"
            "📱 Telefon raqamingizni yuboring:",
            parse_mode="Markdown",
            reply_markup=_phone_kb(),
        )
        return FIRM2_PHONE

    await update.message.reply_text("Iltimos, tugmalardan birini tanlang.")
    return REG_TYPE


async def firm2_phone_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Firma ro'yxatdan o'tishi: telefon qabul qilinadi."""
    if update.message.text == "⬅️ Orqaga":
        return await register_enter(update, ctx)

    contact = update.message.contact
    if not contact:
        await update.message.reply_text("❌ Iltimos, tugma orqali telefon yuboring.", reply_markup=_phone_kb())
        return FIRM2_PHONE

    ctx.user_data["firm2_phone"] = normalize_phone(contact.phone_number)
    await update.message.reply_text(
        "🏢 Firmangiz nomini kiriting:",
        reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
    )
    return FIRM2_NAME


async def firm2_name_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Firma ro'yxatdan o'tishi: firma nomi qabul qilinadi va saqlanadi."""
    from bot import main_keyboard, get_lang, MENU

    if update.message.text == "⬅️ Orqaga":
        await update.message.reply_text(
            "🏢 *Firma uchun ro'yxatdan o'tish*\n\n📱 Telefon raqamingizni yuboring:",
            parse_mode="Markdown", reply_markup=_phone_kb(),
        )
        return FIRM2_PHONE

    firma_nomi = update.message.text.strip()
    if len(firma_nomi) < 2:
        await update.message.reply_text("❌ Firma nomini to'liq kiriting:")
        return FIRM2_NAME

    user_id = update.effective_user.id
    phone = ctx.user_data.get("firm2_phone", "")

    result = save_firma(firma_nomi, phone, user_id)

    if result == "ok":
        await update.message.reply_text(
            f"🎉 *Muvaffaqiyatli ro'yxatdan o'tdingiz!*\n\n"
            f"🏢 {firma_nomi}\n📱 {phone}\n\n"
            f"Hisobotlaringiz shu hisobingizga yuboriladi.",
            parse_mode="Markdown",
        )
    elif result == "firma_taken":
        await update.message.reply_text(
            f"❌ *{firma_nomi}* firmasi allaqachon ro'yxatdan o'tgan.\n\n"
            "Agar bu xato deb hisoblasangiz, administratorga murojaat qiling.",
            parse_mode="Markdown",
        )
    elif result == "phone_taken":
        await update.message.reply_text(
            "❌ Bu telefon raqami allaqachon boshqa firma uchun ro'yxatdan o'tgan.\n\n"
            "Agar bu xato deb hisoblasangiz, administratorga murojaat qiling."
        )
    else:
        await update.message.reply_text("⚠️ Xatolik yuz berdi. Qayta urinib ko'ring.")

    await update.message.reply_text("📋 Asosiy menyu", reply_markup=main_keyboard(get_lang(ctx)))
    return MENU


FIRMALAR_WS_NAME = "Firmalar"


def _get_firmalar_ws():
    """PHARMACY_SHEET_ID ichidagi "Firmalar" varag'ini qaytaradi (bo'lmasa yaratadi)."""
    client = _get_client()
    sh = client.open_by_key(PHARMACY_SHEET_ID)
    try:
        return sh.worksheet(FIRMALAR_WS_NAME)
    except Exception:
        ws = sh.add_worksheet(title=FIRMALAR_WS_NAME, rows=200, cols=5)
        ws.update("A1:E1", [["Firma nomi", "Telefon", "TelegramID", "FileID", "FileName"]])
        return ws


def _norm_firma_nomi(s: str) -> str:
    """Firma nomini solishtirish uchun normallashtiradi (bo'shliq/ko'rinmas belgilar)."""
    s = re.sub(r"[\s\u00a0\u200b\u200c\u200d\ufeff]+", " ", str(s))
    return s.strip().upper()


def save_firma(firma_nomi: str, phone: str, user_id: int) -> str:
    """
    Firmani "Firmalar" varag'iga saqlaydi.

    XAVFSIZLIK QOIDASI: bitta firma nomi va bitta telefon raqami faqat
    BIR MARTA ro'yxatdan o'tishi mumkin. Agar firma nomi allaqachon
    (TelegramID bilan) ro'yxatdan o'tgan bo'lsa yoki shu telefon raqami
    boshqa firma uchun ishlatilgan bo'lsa — rad etiladi.

    Qaytaradi: "ok" | "firma_taken" | "phone_taken" | "error"
    """
    try:
        ws = _get_firmalar_ws()
        all_values = ws.get_all_values()
        target_firma = _norm_firma_nomi(firma_nomi)
        target_phone = normalize_phone(phone)

        empty_row_idx = None  # admin oldindan qo'shib qo'ygan, hali bo'sh qator

        for i, row in enumerate(all_values[1:], start=2):
            if not row:
                continue
            row_firma = _norm_firma_nomi(row[0]) if len(row) > 0 else ""
            row_phone = normalize_phone(row[1]) if len(row) > 1 else ""
            row_tid = str(row[2]).strip() if len(row) > 2 else ""

            if row_firma == target_firma:
                if row_tid:
                    return "firma_taken"
                empty_row_idx = i
                continue

            if row_phone and row_phone == target_phone and row_tid:
                return "phone_taken"

        if empty_row_idx:
            ws.update_cell(empty_row_idx, 2, phone)
            ws.update_cell(empty_row_idx, 3, str(user_id))
            return "ok"

        ws.append_row([firma_nomi, phone, str(user_id), "", ""])
        return "ok"
    except Exception as e:
        print(f"[FIRM2] Saqlash xato: {e}")
        return "error"


async def reg_phone_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Telefon qabul qilinadi."""

    if update.message.text == "⬅️ Orqaga":
        from bot import main_keyboard, get_lang, MENU
        await update.message.reply_text(
            "📋 Asosiy menyu",
            reply_markup=main_keyboard(get_lang(ctx))
        )
        return MENU

    # Kontakt tugmasi orqali yuborilgan
    if update.message.contact:
        phone = normalize_phone(update.message.contact.phone_number)
    elif update.message.text:
        phone = normalize_phone(update.message.text.strip())
    else:
        await update.message.reply_text(
            "❌ Iltimos, telefon raqamingizni yuboring.",
            reply_markup=_phone_kb(),
        )
        return REG_PHONE

    ctx.user_data["reg_phone"] = phone
    user_id = update.effective_user.id

    # Sheets da tekshirish
    farmatsevt = await run_read(find_by_phone, phone)

    if farmatsevt:
        # MAVJUD FARMATSEVT
        if farmatsevt["telegramid"] and farmatsevt["telegramid"] not in ["", "0"]:
            # TelegramID allaqachon bor
            await update.message.reply_text(
                f"✅ Siz allaqachon ro'yxatdan o'tgansiz!\n\n"
                f"👤 {farmatsevt['ismi']}\n"
                f"🏥 {farmatsevt['filial']}\n"
                f"👔 {farmatsevt['lavozim']}",
                parse_mode="Markdown",
                reply_markup=ReplyKeyboardRemove(),
            )
            from bot import main_keyboard, get_lang, MENU
            await update.message.reply_text(
                "📋 Asosiy menyu",
                reply_markup=main_keyboard(get_lang(ctx))
            )
            return MENU
        else:
            # TelegramID bo'sh — saqlaydi
            ok = await run_write(save_telegram_id, farmatsevt["row_num"], user_id)
            if ok:
                await update.message.reply_text(
                    f"🎉 *Ro'yxatdan o'tdingiz!*\n\n"
                    f"👤 {farmatsevt['ismi']}\n"
                    f"🏥 {farmatsevt['filial']}\n"
                    f"👔 {farmatsevt['lavozim']}",
                    parse_mode="Markdown",
                    reply_markup=ReplyKeyboardRemove(),
                )
            else:
                await update.message.reply_text(
                    "❌ Xatolik. Qayta urinib ko'ring.",
                    reply_markup=ReplyKeyboardRemove(),
                )
            from bot import main_keyboard, get_lang, MENU
            await update.message.reply_text(
                "📋 Asosiy menyu",
                reply_markup=main_keyboard(get_lang(ctx))
            )
            return MENU
    else:
        # YANGI FARMATSEVT — ismi so'raladi
        await update.message.reply_text(
            "👤 Ismingizni kiriting (To'liq ism va familiya):",
            reply_markup=_back_kb(),
        )
        return REG_NAME


async def reg_name_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Ismi qabul qilinadi."""
    txt = update.message.text.strip()

    if txt == "⬅️ Orqaga":
        await update.message.reply_text(
            "📱 Telefon raqamingizni yuboring:",
            reply_markup=_phone_kb(),
        )
        return REG_PHONE

    if len(txt) < 3:
        await update.message.reply_text(
            "❌ Ism juda qisqa. Iltimos to'liq ismingizni kiriting:",
            reply_markup=_back_kb(),
        )
        return REG_NAME

    ctx.user_data["reg_ismi"] = txt

    await update.message.reply_text(
        "🏥 Filial raqamini kiriting:\n_(masalan: 6)_",
        parse_mode="Markdown",
        reply_markup=_back_kb(),
    )
    return REG_FILIAL


async def reg_filial_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Filial raqami qabul qilinadi, tekshiriladi va Sheets ga yoziladi."""
    txt = update.message.text.strip()

    if txt == "⬅️ Orqaga":
        await update.message.reply_text(
            "👤 Ismingizni kiriting:",
            reply_markup=_back_kb(),
        )
        return REG_NAME

    # Filial ma'lumotlarini olish
    filial_info = await run_read(get_filial_info, txt)

    if not filial_info:
        await update.message.reply_text(
            f"❌ *{txt}* raqamli filial topilmadi.\n\n"
            "Filial raqamini qayta kiriting:",
            parse_mode="Markdown",
            reply_markup=_back_kb(),
        )
        return REG_FILIAL

    ismi   = ctx.user_data.get("reg_ismi", "")
    phone  = ctx.user_data.get("reg_phone", "")
    user_id = update.effective_user.id

    # Filial ma'lumotlarini saqlab, lavozim so'rash
    ctx.user_data["reg_filial_info"] = filial_info

    await update.message.reply_text(
        f"🏥 *{filial_info['filial_nomi']}*\n\n👔 Lavozimingizni tanlang:",
        parse_mode="Markdown",
        reply_markup=_lavozim_kb(),
    )
    return REG_LAVOZIM




async def reg_lavozim_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Lavozim tanlanadi va Sheets ga yoziladi."""
    txt = update.message.text.strip()

    if txt == "⬅️ Orqaga":
        await update.message.reply_text(
            "🏥 Filial raqamini kiriting:\n_(masalan: 6)_",
            parse_mode="Markdown",
            reply_markup=_back_kb(),
        )
        return REG_FILIAL

    lavozim_map = {
        "👔 Farmatsevt":       "Farmatsevt",
        "👔 Dorixona mudiri":  "Dorixona mudiri",
        "👔 Stajyor":          "Stajyor",
    }

    if txt not in lavozim_map:
        await update.message.reply_text(
            "❌ Iltimos, quyidagi tugmalardan birini tanlang:",
            reply_markup=_lavozim_kb(),
        )
        return REG_LAVOZIM

    lavozim = lavozim_map[txt]
    ismi        = ctx.user_data.get("reg_ismi", "")
    phone       = ctx.user_data.get("reg_phone", "")
    filial_info = ctx.user_data.get("reg_filial_info", {})
    user_id     = update.effective_user.id

    ok = await run_write(
        add_new_farmatsevt,
        ismi=ismi,
        telefon=phone,
        filial_nomi=filial_info.get("filial_nomi", ""),
        lavozim=lavozim,
        lat=filial_info.get("lat", ""),
        lon=filial_info.get("lon", ""),
        after_row=filial_info.get("last_row", 1),
        user_id=user_id,
    )

    if ok:
        await update.message.reply_text(
            f"🎉 *Ro'yxatdan o'tdingiz!*\n\n"
            f"👤 {ismi}\n"
            f"📱 {phone}\n"
            f"🏥 {filial_info.get('filial_nomi', '')}\n"
            f"👔 {lavozim}",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove(),
        )
    else:
        await update.message.reply_text(
            "❌ Xatolik yuz berdi. Admin bilan bog'laning.",
            reply_markup=ReplyKeyboardRemove(),
        )

    from bot import main_keyboard, get_lang, MENU
    await update.message.reply_text(
        "📋 Asosiy menyu",
        reply_markup=main_keyboard(get_lang(ctx))
    )
    return MENU

# ─── States ───────────────────────────────────────────────────────────────────

def get_reg_states():
    return {
        REG_TYPE: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, reg_type_handler),
        ],
        REG_PHONE: [
            MessageHandler(filters.CONTACT, reg_phone_handler),
            MessageHandler(filters.TEXT & ~filters.COMMAND, reg_phone_handler),
        ],
        REG_NAME: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, reg_name_handler),
        ],
        REG_FILIAL: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, reg_filial_handler),
        ],
        REG_LAVOZIM: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, reg_lavozim_handler),
        ],
        FIRM2_PHONE: [
            MessageHandler(filters.CONTACT, firm2_phone_handler),
            MessageHandler(filters.TEXT & ~filters.COMMAND, firm2_phone_handler),
        ],
        FIRM2_NAME: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, firm2_name_handler),
        ],
    }


def _col_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _add_filial_headers_to_ws(ws) -> dict:
    """
    Berilgan varaqda har bir filial guruhi uchun, agar hali bo'lmasa,
    sarlavha (header) qatorini qo'shadi — bu qatorda FAQAT filial nomi
    bo'ladi, hech qanday xodim (Ismi/Telefon) bo'lmaydi.

    MUHIM: avval A ustunidagi BARCHA eski "merge" (birlashtirilgan)
    katakchalar bekor qilinadi — chunki merge qilingan katakka yozilgan
    qiymat faqat birinchi (yuqori-chap) katakda saqlanib, qolganlari
    bo'sh ko'rinib qolgan edi. So'ng, filial qiymati yo'qolib qolgan
    (lekin Ismi bor) qatorlar yuqoridagi oxirgi ko'ringan filial nomi
    bilan "forward-fill" qilib tiklanadi. Faqat shundan keyin yangi
    sarlavha qatorlari qo'shiladi.

    Butun jadval xotirada qayta quriladi va yozish ham merge ham
    ma'lumot BITTA-BITTA so'rov bilan bajariladi (kvota va tartib
    buzilishining oldini olish uchun).
    """
    result = {"added": 0, "repaired": 0, "error": None}
    try:
        all_values = ws.get_all_values()
        if not all_values:
            result["error"] = "Jadval bo'sh"
            return result

        header = all_values[0]
        ncols = len(header)
        data_rows = [list(r) + [""] * (ncols - len(r)) for r in all_values[1:]]
        n = len(data_rows)

        # 1) A ustunidagi barcha eski merge'larni bekor qilamiz
        if n:
            try:
                ws.spreadsheet.batch_update({"requests": [{"unmergeCells": {"range": {
                    "sheetId": ws.id,
                    "startRowIndex": 1, "endRowIndex": 1 + n,
                    "startColumnIndex": 0, "endColumnIndex": 1,
                }}}]})
            except Exception as e:
                print(f"[FIX] Unmerge xato (davom etamiz): {e}")

        # 2) Filial qiymati yo'qolgan (lekin Ismi bor) qatorlarni tiklaymiz
        effective_filial = ""
        repaired = 0
        for row in data_rows:
            if row[0].strip():
                effective_filial = row[0].strip()
            elif len(row) > 1 and row[1].strip() and effective_filial:
                row[0] = effective_filial
                repaired += 1

        # 3) Har bir filial guruhi uchun sarlavha qatorini qo'shamiz
        new_rows = []
        last_kod = None
        added = 0

        for row in data_rows:
            filial_val = row[0].strip() if row[0] else ""
            ismi_val = row[1].strip() if len(row) > 1 and row[1] else ""
            cur_kod = _filial_kod(filial_val) if filial_val else None

            if cur_kod and cur_kod != last_kod:
                if ismi_val:
                    header_row = [""] * ncols
                    header_row[0] = filial_val
                    new_rows.append(header_row)
                    added += 1
                last_kod = cur_kod

            new_rows.append(row)

        if added or repaired:
            # XAVFSIZLIK TO'SIG'I: yangi qatorlar soni asl holatidan (yoki
            # qo'shilishi kutilgan sondan) kam bo'lib qolsa — yozishni
            # to'xtatamiz, ma'lumot yo'qolib ketmasligi uchun.
            expected_min = len(data_rows) + added
            if len(new_rows) < expected_min:
                result["error"] = (
                    f"XAVFSIZLIK: qayta qurilgan qatorlar soni ({len(new_rows)}) "
                    f"kutilganidan ({expected_min}) kam — yozish BEKOR QILINDI."
                )
                return result

            needed_rows = 1 + len(new_rows)
            if ws.row_count < needed_rows:
                ws.resize(rows=needed_rows)
            ws.update(f"A2:{_col_letter(ncols)}{needed_rows}", new_rows, value_input_option="USER_ENTERED")

        result["added"] = added
        result["repaired"] = repaired
        return result

    except Exception as e:
        print(f"[FIX] add_filial_headers xato: {e}")
        result["error"] = str(e)
        return result


def add_filial_headers() -> dict:
    """"Farmatsevtlar" jadvaliga (PHARMACY_SHEET_ID) sarlavha qatorlarini qo'shadi."""
    try:
        client = _get_client()
        sh = client.open_by_key(PHARMACY_SHEET_ID)
        ws = sh.worksheet("Farmatsevtlar")
        return _add_filial_headers_to_ws(ws)
    except Exception as e:
        return {"added": 0, "error": str(e)}


def add_filial_headers_by_gid(sheet_id: str, gid: int) -> dict:
    """Berilgan Google Sheet ID va gid (varaq ID) bo'yicha sarlavha qatorlarini qo'shadi."""
    try:
        client = _get_client()
        sh = client.open_by_key(sheet_id)
        ws = None
        for w in sh.worksheets():
            if w.id == gid:
                ws = w
                break
        if ws is None:
            return {"added": 0, "error": f"gid={gid} bo'lgan varaq topilmadi"}
        return _add_filial_headers_to_ws(ws)
    except Exception as e:
        return {"added": 0, "error": str(e)}



def _get_all_lavozim_map() -> dict:
    """Butun Farmatsevtlar jadvalidan Ismi -> Lavozim lug'atini (barcha filiallar) qaytaradi."""
    try:
        client = _get_client()
        ws = client.open_by_key(PHARMACY_SHEET_ID).worksheet("Farmatsevtlar")
        all_values = ws.get_all_values()
        result = {}
        for row in all_values[1:]:
            if not row or len(row) < 2 or not row[1]:
                continue
            ismi = str(row[1]).strip().upper()
            lavozim = str(row[4]).strip() if len(row) > 4 else ""
            result[ismi] = lavozim
        return result
    except Exception as e:
        print(f"[FIX] _get_all_lavozim_map xato: {e}")
        return {}


def _reorder_ws_by_lavozim(ws) -> dict:
    """
    Berilgan varaqdagi mavjud xodimlarni HAR BIR filial guruhi ichida
    Lavozim tartibi bo'yicha (Dorixona mudiri → Farmatsevt → Stajyor)
    qayta saralaydi. Sarlavha (bo'sh Ismi) qatorlari har doim guruh
    boshida qoladi. Filiallarning o'zaro tartibi (guruhlar ketma-ketligi)
    o'zgarmaydi — faqat har bir guruh ICHIDAGI xodimlar saralanadi.
    """
    result = {"reordered_groups": 0, "error": None}
    try:
        all_values = ws.get_all_values()
        if not all_values:
            result["error"] = "Jadval bo'sh"
            return result

        header = all_values[0]
        ncols = len(header)
        data_rows = [list(r) + [""] * (ncols - len(r)) for r in all_values[1:]]

        lavozim_map = _get_all_lavozim_map()

        # Guruhlarga ajratamiz (filial kodi bo'yicha, ketma-ketlikni saqlab)
        groups = []  # [(filial_kod, [rows])]
        cur_kod = None
        cur_group = []
        for row in data_rows:
            filial_val = row[0].strip() if row[0] else ""
            kod = _filial_kod(filial_val) if filial_val else cur_kod
            if kod != cur_kod and cur_group:
                groups.append((cur_kod, cur_group))
                cur_group = []
            cur_kod = kod
            cur_group.append(row)
        if cur_group:
            groups.append((cur_kod, cur_group))

        new_rows = []
        reordered_groups = 0

        for kod, group_rows in groups:
            headers_sub = [r for r in group_rows if not (len(r) > 1 and r[1].strip())]
            employees = [r for r in group_rows if len(r) > 1 and r[1].strip()]

            def _priority(row):
                ismi = row[1].strip().upper()
                lav = lavozim_map.get(ismi, "")
                return _lavozim_priority(lav) if lav else 2

            sorted_employees = sorted(employees, key=_priority)
            if [r[1] for r in sorted_employees] != [r[1] for r in employees]:
                reordered_groups += 1

            new_rows.extend(headers_sub)
            new_rows.extend(sorted_employees)

        if reordered_groups:
            # XAVFSIZLIK TO'SIG'I: agar biror sababdan qayta qurilgan
            # ma'lumot asl qatorlar sonidan KAM bo'lib qolsa — bu real
            # ma'lumot yo'qolganini bildiradi. Bunday holda YOZISHNI
            # BUTUNLAY TO'XTATAMIZ, xatolik qaytaramiz.
            if len(new_rows) < len(data_rows):
                result["error"] = (
                    f"XAVFSIZLIK: qayta qurilgan qatorlar soni ({len(new_rows)}) "
                    f"asl holatidan ({len(data_rows)}) kam — yozish BEKOR QILINDI, "
                    f"ma'lumot yo'qolmasligi uchun."
                )
                return result

            needed_rows = 1 + len(new_rows)
            if ws.row_count < needed_rows:
                ws.resize(rows=needed_rows)
            ws.update(f"A2:{_col_letter(ncols)}{needed_rows}", new_rows, value_input_option="USER_ENTERED")

        result["reordered_groups"] = reordered_groups
        return result

    except Exception as e:
        print(f"[FIX] _reorder_ws_by_lavozim xato: {e}")
        result["error"] = str(e)
        return result


def reorder_by_lavozim_by_gid(sheet_id: str, gid: int) -> dict:
    """Berilgan Google Sheet ID va gid bo'yicha xodimlarni Lavozim tartibida qayta saralaydi."""
    try:
        client = _get_client()
        sh = client.open_by_key(sheet_id)
        ws = None
        for w in sh.worksheets():
            if w.id == gid:
                ws = w
                break
        if ws is None:
            return {"reordered_groups": 0, "error": f"gid={gid} bo'lgan varaq topilmadi"}
        return _reorder_ws_by_lavozim(ws)
    except Exception as e:
        return {"reordered_groups": 0, "error": str(e)}


async def cmd_reorder_by_lavozim_salary(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/reorder_lavozim_salary — SALARY_SHEET_ID dagi gid=0 varag'ini Lavozim tartibida qayta saralaydi (admin)."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return
    msg = await update.message.reply_text("⏳ Xodimlar tartiblanmoqda...")
    try:
        from attendance import run_write
        res = await run_write(reorder_by_lavozim_by_gid, SALARY_SHEET_ID, 0)
        if res.get("error"):
            await msg.edit_text(f"❌ Xato: {res['error']}")
        else:
            await msg.edit_text(f"✅ Tayyor! {res['reordered_groups']} ta filial guruhi qayta tartiblandi.")
    except Exception as e:
        await msg.edit_text(f"❌ Xato: {e}")


def _get_all_filials_from_farmatsevtlar() -> list:
    """
    "Farmatsevtlar" jadvalidan (PHARMACY_SHEET_ID) BARCHA filiallarning
    to'liq ro'yxatini qaytaradi — xodimi bo'lsin-bo'lmasin, chunki bu
    jadvalda xodimsiz filiallar ham sarlavha (yoki hech bo'lmasa Lat/Lon
    bilan) qator sifatida mavjud.

    Qaytaradi: [(kod, to'liq_nom), ...] — jadvaldagi tartibda, dublikatsiz.
    """
    try:
        client = _get_client()
        ws = client.open_by_key(PHARMACY_SHEET_ID).worksheet("Farmatsevtlar")
        all_values = ws.get_all_values()
        seen = set()
        result = []
        for row in all_values[1:]:
            if not row or not row[0]:
                continue
            filial_val = str(row[0]).strip()
            kod = _filial_kod(filial_val)
            if not kod or kod in seen:
                continue
            seen.add(kod)
            result.append((kod, filial_val))
        return result
    except Exception as e:
        print(f"[FIX] _get_all_filials_from_farmatsevtlar xato: {e}")
        return []


def _sync_all_filials_to_ws(ws) -> dict:
    """
    Berilgan varaqni Farmatsevtlar jadvalidagi filiallar RAQAMLI
    TARTIBIGA mos ravishda qayta quradi: mavjud xodimlar o'z guruhida
    saqlanadi, yetishmayotgan filiallar esa o'zining TO'G'RI raqamli
    o'rniga (masalan 13 va 15 orasiga, agar 14 yetishmasa) bo'sh
    sarlavha qatori sifatida qo'shiladi.
    """
    result = {"added": 0, "error": None}
    try:
        all_filials = _get_all_filials_from_farmatsevtlar()  # (kod, nom), Farmatsevtlar tartibida
        if not all_filials:
            result["error"] = "Farmatsevtlar jadvalidan filiallar topilmadi"
            return result

        all_values = ws.get_all_values()
        if not all_values:
            result["error"] = "Jadval bo'sh"
            return result

        header = all_values[0]
        ncols = len(header)
        data_rows = [list(r) + [""] * (ncols - len(r)) for r in all_values[1:]]

        # Mavjud qatorlarni filial kodi bo'yicha guruhlaymiz (o'z ichidagi
        # tartibni saqlab), forward-fill bilan (bo'sh A ustuniga chidamli)
        groups_by_kod = {}
        unmatched_rows = []  # kod aniqlanmagan qatorlar (juda kam uchraydi)
        cur_kod = None
        for row in data_rows:
            filial_val = row[0].strip() if row[0] else ""
            kod = _filial_kod(filial_val) if filial_val else cur_kod
            if kod:
                cur_kod = kod
                groups_by_kod.setdefault(kod, []).append(row)
            else:
                unmatched_rows.append(row)

        new_rows = []
        added = 0

        # Farmatsevtlar tartibida (raqamli ketma-ketlikda) qayta quramiz
        for kod, nom in all_filials:
            if kod in groups_by_kod:
                new_rows.extend(groups_by_kod.pop(kod))
            else:
                header_row = [""] * ncols
                header_row[0] = nom
                new_rows.append(header_row)
                added += 1

        # Ehtiyot uchun: Farmatsevtlar ro'yxatida yo'q, lekin bu jadvalda
        # mavjud bo'lgan filiallar bo'lsa (masalan qo'lda kiritilgan) —
        # ularni OXIRIGA qo'shamiz, hech narsa yo'qotmaslik uchun
        for kod, rows in groups_by_kod.items():
            new_rows.extend(rows)
        new_rows.extend(unmatched_rows)

        # XAVFSIZLIK TO'SIG'I: umumiy ma'lumot qatorlari hech qachon
        # kamaymasligi kerak
        if len(new_rows) < len(data_rows):
            result["error"] = (
                f"XAVFSIZLIK: qayta qurilgan qatorlar soni ({len(new_rows)}) "
                f"asl holatidan ({len(data_rows)}) kam — yozish bekor qilindi."
            )
            return result

        needed_rows = 1 + len(new_rows)
        if ws.row_count < needed_rows:
            ws.resize(rows=needed_rows)
        ws.update(f"A2:{_col_letter(ncols)}{needed_rows}", new_rows, value_input_option="USER_ENTERED")

        result["added"] = added
        return result

    except Exception as e:
        print(f"[FIX] _sync_all_filials_to_ws xato: {e}")
        result["error"] = str(e)
        return result


def sync_all_filials_by_gid(sheet_id: str, gid: int) -> dict:
    """Berilgan Google Sheet ID va gid bo'yicha barcha filiallarni (Farmatsevtlar asosida) sinxronlaydi."""
    try:
        client = _get_client()
        sh = client.open_by_key(sheet_id)
        ws = None
        for w in sh.worksheets():
            if w.id == gid:
                ws = w
                break
        if ws is None:
            return {"added": 0, "error": f"gid={gid} bo'lgan varaq topilmadi"}
        return _sync_all_filials_to_ws(ws)
    except Exception as e:
        return {"added": 0, "error": str(e)}


async def cmd_sync_all_filials_salary(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/sync_all_filials_salary — SALARY_SHEET_ID gid=0 varag'iga yetishmayotgan barcha filiallarni qo'shadi (admin)."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return
    msg = await update.message.reply_text("⏳ Filiallar tekshirilmoqda...")
    try:
        from attendance import run_write
        res = await run_write(sync_all_filials_by_gid, SALARY_SHEET_ID, 0)
        if res.get("error"):
            await msg.edit_text(f"❌ Xato: {res['error']}")
        else:
            await msg.edit_text(f"✅ Tayyor! {res['added']} ta yetishmayotgan filial qo'shildi.")
    except Exception as e:
        await msg.edit_text(f"❌ Xato: {e}")


async def cmd_add_filial_headers(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/add_filial_headers — Farmatsevtlar jadvaliga filial sarlavha qatorlarini qo'shadi (admin)."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return
    msg = await update.message.reply_text("⏳ Sarlavha qatorlari tekshirilmoqda...")
    try:
        from attendance import run_write
        res = await run_write(add_filial_headers)
        if res.get("error"):
            await msg.edit_text(f"❌ Xato: {res['error']}")
        else:
            await msg.edit_text(f"✅ Tayyor! {res['added']} ta yangi sarlavha qatori qoshildi, {res.get('repaired',0)} ta yoqolgan filial qiymati tiklandi.")
    except Exception as e:
        await msg.edit_text(f"❌ Xato: {e}")


async def cmd_add_filial_headers_salary(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/add_filial_headers_salary — SALARY_SHEET_ID dagi gid=0 varag'iga sarlavha qatorlarini qo'shadi (admin)."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return
    msg = await update.message.reply_text("⏳ Sarlavha qatorlari tekshirilmoqda...")
    try:
        from attendance import run_write
        res = await run_write(add_filial_headers_by_gid, SALARY_SHEET_ID, 0)
        if res.get("error"):
            await msg.edit_text(f"❌ Xato: {res['error']}")
        else:
            await msg.edit_text(f"✅ Tayyor! {res['added']} ta yangi sarlavha qatori qoshildi, {res.get('repaired',0)} ta yoqolgan filial qiymati tiklandi.")
    except Exception as e:
        await msg.edit_text(f"❌ Xato: {e}")
