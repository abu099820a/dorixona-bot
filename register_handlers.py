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

# Conversation states
REG_PHONE   = 400
REG_NAME    = 401
REG_FILIAL  = 402

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
    Filial kodiga mos birinchi xodimning Lat, Lon, Filial nomini qaytaradi.
    Shu filialning oxirgi qator raqamini ham qaytaradi (yangi xodimni keyin qo'shish uchun).
    """
    try:
        client = _get_client()
        ws = client.open_by_key(PHARMACY_SHEET_ID).sheet1
        all_values = ws.get_all_values()

        filial_nomi = None
        lat = ""
        lon = ""
        last_row = 1  # sarlavha

        for i, row in enumerate(all_values):
            if i == 0:
                continue  # sarlavha
            filial_cell = str(row[0]).strip() if len(row) > 0 else ""
            if _filial_kod(filial_cell) == filial_kod.strip():
                filial_nomi = filial_cell
                if not lat and len(row) > 5 and row[5]:
                    lat = row[5]
                if not lon and len(row) > 6 and row[6]:
                    lon = row[6]
                last_row = i + 1  # 1-indexed (sarlavha hisobga olingan)

        if filial_nomi is None:
            return None

        return {
            "filial_nomi": filial_nomi,
            "lat":         lat,
            "lon":         lon,
            "last_row":    last_row,  # shu filialning oxirgi qatori
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
        _add_to_attendance(ismi, filial_nomi)

        return True
    except Exception as e:
        print(f"[REG] Yangi farmatsevt qo'shish xato: {e}")
        return False


def _add_to_attendance(ismi: str, filial_nomi: str):
    """
    Davomat Sheets dagi joriy oy listiga yangi farmatsevtni qo'shadi.
    Shu filialdagi oxirgi xodimdan keyin qo'shadi.
    """
    try:
        if not ATTENDANCE_SHEET_ID:
            return

        from datetime import datetime, timezone, timedelta
        import calendar

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
        last_row = 2  # sarlavhadan keyin

        # Shu filialdagi oxirgi xodimni topish
        for i, row in enumerate(all_values):
            if i < 2:
                continue
            if not row or not row[0]:
                continue
            row_filial = str(row[1]).strip() if len(row) > 1 else ""
            if _filial_kod(row_filial) == filial_kod:
                last_row = i + 1

        insert_row = last_row + 1
        ws.insert_row([ismi, filial_nomi], index=insert_row)
        print(f"[REG] Davomat ga qo'shildi: {ismi} | {filial_nomi} | qator {insert_row}")

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

    # Yangi farmatsevtni qo'shish
    ok = add_new_farmatsevt(
        ismi=ismi,
        telefon=phone,
        filial_nomi=filial_info["filial_nomi"],
        lavozim="Farmatsevt",   # Yangi xodim uchun default lavozim
        lat=filial_info["lat"],
        lon=filial_info["lon"],
        after_row=filial_info["last_row"],
        user_id=user_id,
    )

    if ok:
        await update.message.reply_text(
            f"🎉 *Ro'yxatdan o'tdingiz!*\n\n"
            f"👤 {ismi}\n"
            f"📱 {phone}\n"
            f"🏥 {filial_info['filial_nomi']}\n"
            f"👔 Farmatsevt",
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
    }
