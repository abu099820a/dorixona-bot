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

# ─── Sozlamalar ───────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
PHARMACY_SHEET_ID   = os.getenv("PHARMACY_SHEET_ID", "")
ATTENDANCE_SHEET_ID = os.getenv("ATTENDANCE_SHEET_ID", "")
FILIALLAR_SHEET_ID  = os.getenv("FILIALLAR_SHEET_ID", "")

# Conversation states
REG_PHONE   = 400
REG_NAME    = 401
REG_FILIAL  = 402
REG_LAVOZIM = 403

# Ustun raqamlari (1-indexed)
COL_FILIAL    = 1  # A
COL_ISMI      = 2  # B
COL_TELEFON   = 3  # C
COL_TELEGRAMID= 4  # D
COL_LAVOZIM   = 5  # E
COL_LAT       = 6  # F
COL_LON       = 7  # G

# ─── Yordamchi ───────────────────────────────────────────────────────────────

def _get_client():
    creds_json = os.getenv("GOOGLE_CREDENTIALS")
    if creds_json:
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    return gspread.authorize(creds)


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
        ws = client.open_by_key(PHARMACY_SHEET_ID).sheet1
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
        ws = client.open_by_key(PHARMACY_SHEET_ID).sheet1
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
        ph_ws = client.open_by_key(PHARMACY_SHEET_ID).sheet1
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


def add_new_farmatsevt(
    ismi: str, telefon: str, filial_nomi: str,
    lavozim: str, lat: str, lon: str,
    after_row: int, user_id: int
) -> bool:
    """
    Yangi farmatsevtni Sheets ga qo'shadi.
    after_row: shu qatordan KEYIN (bir pastga) qo'shiladi.
    Davomat Sheets ga ham avtomatik qo'shiladi.
    """
    try:
        client = _get_client()
        ws = client.open_by_key(PHARMACY_SHEET_ID).sheet1

        # Qatorni after_row dan keyin qo'shish
        insert_row = after_row + 1

        # Qator qo'shish (pastki qatorlarni surish)
        ws.insert_row(
            [filial_nomi, ismi, telefon, str(user_id), lavozim, lat, lon],
            index=insert_row,
            value_input_option="USER_ENTERED"
        )

        print(f"[REG] Yangi farmatsevt qo'shildi: {ismi} | {filial_nomi} | qator {insert_row}")

        # Davomat Sheets ga ham qo'shish
        _add_to_attendance(ismi, filial_nomi, telefon)

        return True
    except Exception as e:
        print(f"[REG] Yangi farmatsevt qo'shish xato: {e}")
        return False


def _add_to_attendance(ismi: str, filial_nomi: str, telefon: str = ""):
    """
    Davomat Sheets dagi joriy oy listiga yangi farmatsevtni qo'shadi.
    Jadval: A=Filial, B=Ismi
    Shu filialdagi oxirgi xodimdan keyin qo'shadi.
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
        last_row = 2  # default

        # Jadval: A=Filial, B=Ismi
        # Shu filialdagi OXIRGI qatorni topish (sarlavha + xodimlar ichida)
        for i, row in enumerate(all_values):
            if i < 2:
                continue
            if not row or not row[0]:
                continue
            # A ustun = Filial
            row_filial = str(row[0]).strip()
            if _filial_kod(row_filial) == filial_kod:
                last_row = i + 1  # 1-indexed

        insert_row = last_row + 1
        # A=Filial, B=Ismi, C=Telefon tartibida qo'shish
        ws.insert_row([filial_nomi, ismi, telefon], index=insert_row)
        print(f"[REG] Davomat ga qo'shildi: {filial_nomi} | {ismi} | {telefon} | qator {insert_row}")

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
    """'📝 Ro'yxatdan o'tish' bosilganda."""
    ctx.user_data.pop("reg_phone", None)
    ctx.user_data.pop("reg_ismi", None)
    ctx.user_data.pop("reg_filial_info", None)

    await update.message.reply_text(
        "📝 *Ro'yxatdan o'tish*\n\n"
        "📱 Telefon raqamingizni yuboring:",
        parse_mode="Markdown",
        reply_markup=_phone_kb(),
    )
    return REG_PHONE


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
    farmatsevt = find_by_phone(phone)

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
            ok = save_telegram_id(farmatsevt["row_num"], user_id)
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
    filial_info = get_filial_info(txt)

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

    ok = add_new_farmatsevt(
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
    }
