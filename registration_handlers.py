"""
registration_handlers.py — Ro'yxatdan o'tish moduli
"""

from telegram import Update, ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import MessageHandler, CallbackQueryHandler, ContextTypes, filters
import gspread
import os
import json
import re
from google.oauth2.service_account import Credentials

# ─── Sozlamalar ───────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

PHARMACY_SHEET_ID = os.getenv("PHARMACY_SHEET_ID", "")

# States
(
    REG_PHONE,
    REG_FILIAL,
    REG_CONFIRM,
    REG_LAVOZIM,
) = range(200, 204)

# ─── Google Sheets ─────────────────────────────────────────────────────────────

def get_client():
    creds_json = os.getenv("GOOGLE_CREDENTIALS")
    if creds_json:
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    return gspread.authorize(creds)


def normalize_phone(phone) -> str:
    if phone is None: return ""
    digits = re.sub(r'[^0-9]', '', str(phone))
    if not digits: return ""
    if digits.startswith("998"): return "+" + digits
    if digits.startswith("0"): return "+998" + digits[1:]
    if len(digits) == 9: return "+998" + digits
    return "+" + digits


def get_filiallar():
    """Farmatsevtlar Sheet A ustunidan filiallar ro'yxatini oladi"""
    try:
        client = get_client()
        sh = client.open_by_key(PHARMACY_SHEET_ID)
        ws = sh.sheet1
        all_values = ws.get_all_values()
        filiallar = {}
        import re
        for i, row in enumerate(all_values):
            if i == 0:
                continue  # sarlavha
            a_val = str(row[0]).strip() if row else ""
            if not a_val:
                continue
            # Filial qatori: B ustuni bosh (chodim emas)
            b_val = str(row[1]).strip() if len(row) > 1 else ""
            if not b_val:
                # Raqamni olish: "1 — ТАШМИ-1" -> no="1", nom="ТАШМИ-1"
                match = re.match(r'^(\d+)\s*[—-]\s*(.+)$', a_val)
                if match:
                    no = match.group(1)
                    nom = match.group(2).strip()
                    filiallar[no] = {"nom": nom, "full": a_val.replace(" — ", " - ")}
                elif a_val.startswith("Asosiy"):
                    filiallar["0"] = {"nom": a_val, "full": a_val.replace(" — ", " - ")}
        return filiallar
    except Exception as e:
        print(f"[REG] Filiallar xato: {e}")
        return {}


def is_already_registered(phone: str, user_id: int) -> bool:
    """Allaqachon ro'yxatdan o'tganmi tekshiradi"""
    try:
        client = get_client()
        sh = client.open_by_key(PHARMACY_SHEET_ID)
        ws = sh.sheet1
        records = ws.get_all_records()
        norm = normalize_phone(phone)
        for row in records:
            # Telefon tekshirish
            tel = row.get("Telefon", "")
            if isinstance(tel, float): tel = str(int(tel))
            if normalize_phone(str(tel)) == norm:
                return True
            # TelegramID tekshirish
            if str(row.get("TelegramID", "")).strip() == str(user_id):
                return True
        return False
    except Exception as e:
        print(f"[REG] Tekshirish xato: {e}")
        return False


def save_registration(user_id: int, ismi: str, phone: str, filial: str, lavozim: str) -> bool:
    """
    Farmatsevtni Sheet'ga saqlaydi.
    Filial A ustunida topiladi, B ustuniga ismi yoziladi.
    """
    try:
        client = get_client()
        sh = client.open_by_key(PHARMACY_SHEET_ID)
        ws = sh.sheet1
        all_values = ws.get_all_values()

        # Filial qatorini topish
        filial_row = None
        for i, row in enumerate(all_values):
            if i == 0:
                continue
            a_val = str(row[0]).strip() if row else ""
            if a_val == filial:
                filial_row = i + 1  # 1-indexed
                break

        if not filial_row:
            print(f"[REG] Filial topilmadi: {filial}")
            # Oxiriga qo'shish
            ws.append_row([filial, ismi, normalize_phone(phone), str(user_id), lavozim])
            return True

        # Filial qatoriga B-E ustunlariga yozish
        ws.update_cell(filial_row, 2, ismi)
        ws.update_cell(filial_row, 3, normalize_phone(phone))
        ws.update_cell(filial_row, 4, str(user_id))
        ws.update_cell(filial_row, 5, lavozim)

        return True
    except Exception as e:
        print(f"[REG] Saqlash xato: {e}")
        return False

# ─── Klaviaturalar ─────────────────────────────────────────────────────────────

def phone_keyboard():
    return ReplyKeyboardMarkup([
        [KeyboardButton("📱 Telefon raqamimni yuborish", request_contact=True)],
        ["⬅️ Orqaga"],
    ], resize_keyboard=True)


def lavozim_keyboard():
    return ReplyKeyboardMarkup([
        ["💊 Farmatsevt"],
        ["👑 Dorixona mudiri"],
        ["🎓 Stajyor"],
        ["⬅️ Orqaga"],
    ], resize_keyboard=True)


def confirm_keyboard():
    return ReplyKeyboardMarkup([
        ["✅ Ha, to'g'ri"],
        ["❌ Yo'q, qaytadan"],
    ], resize_keyboard=True)

# ─── Handlerlar ────────────────────────────────────────────────────────────────

async def reg_enter(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Ro'yxatdan o'tish boshlash"""
    user_id = update.effective_user.id

    # Allaqachon ro'yxatdan o'tganmi?
    if ctx.user_data.get("reg_done"):
        await update.message.reply_text(
            "✅ Siz allaqachon ro'yxatdan o'tgansiz!\n"
            f"👤 {ctx.user_data.get('reg_ismi', '')}\n"
            f"🏪 {ctx.user_data.get('reg_filial', '')}\n"
            f"👔 {ctx.user_data.get('reg_lavozim', '')}",
            reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True)
        )
        return None

    await update.message.reply_text(
        "📋 *Ro'yxatdan o'tish*\n\n"
        "Ismingizni kiriting (Ism Familiya):",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
    )
    return REG_PHONE


async def reg_ismi_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Ism qabul qilinadi, keyin telefon so'raladi"""
    if update.message.text == "⬅️ Orqaga":
        from bot import main_keyboard, get_lang, MENU
        await update.message.reply_text("📋 Asosiy menyu", reply_markup=main_keyboard(get_lang(ctx)))
        return MENU

    ismi = update.message.text.strip()
    if len(ismi) < 3:
        await update.message.reply_text("❌ Ismingizni to'liq kiriting (Ism Familiya):")
        return REG_PHONE

    ctx.user_data["reg_ismi"] = ismi
    await update.message.reply_text(
        f"👤 *{ismi}*\n\n📱 Telefon raqamingizni yuboring:",
        parse_mode="Markdown",
        reply_markup=phone_keyboard(),
    )
    return REG_FILIAL


async def reg_phone_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Telefon qabul qilinadi, filial so'raladi"""
    if update.message.text == "⬅️ Orqaga":
        await update.message.reply_text("Ismingizni kiriting:", reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True))
        return REG_PHONE

    contact = update.message.contact
    if not contact:
        await update.message.reply_text("❌ Iltimos, tugma orqali telefon yuboring.", reply_markup=phone_keyboard())
        return REG_FILIAL

    phone = normalize_phone(contact.phone_number)
    user_id = update.effective_user.id

    # Allaqachon ro'yxatdan o'tganmi?
    if is_already_registered(phone, user_id):
        await update.message.reply_text(
            f"⚠️ *{phone}* raqami allaqachon tizimda mavjud!\n"
            "Administratorga murojaat qiling.",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
        )
        return None

    ctx.user_data["reg_phone"] = phone

    await update.message.reply_text(
        f"✅ Telefon: *{phone}*\n\n"
        "🏪 Filial raqamini kiriting (masalan: *5*):",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
    )
    return REG_CONFIRM


async def reg_filial_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Filial raqami qabul qilinadi, tasdiqlanadi"""
    if update.message.text == "⬅️ Orqaga":
        await update.message.reply_text("📱 Telefon yuboring:", reply_markup=phone_keyboard())
        return REG_FILIAL

    filial_no = update.message.text.strip()
    if not filial_no.isdigit():
        await update.message.reply_text("❌ Faqat raqam kiriting (masalan: 5):")
        return REG_CONFIRM

    filiallar = get_filiallar()
    if filial_no not in filiallar:
        await update.message.reply_text(
            f"❌ *{filial_no}* raqamli filial topilmadi.\n"
            "To'g'ri raqam kiriting:",
            parse_mode="Markdown",
        )
        return REG_CONFIRM

    filial_info = filiallar[filial_no]
    filial_text = filial_info.get("full", f"{filial_no} - {filial_info['nom']}")
    ctx.user_data["reg_filial"] = filial_text

    await update.message.reply_text(
        f"🏪 *{filial_text}*\n\nShu filialda ishlaysizmi?",
        parse_mode="Markdown",
        reply_markup=confirm_keyboard(),
    )
    return REG_LAVOZIM


async def reg_confirm_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Filial tasdiqlanadi, lavozim so'raladi"""
    txt = update.message.text

    if txt == "❌ Yo'q, qaytadan":
        await update.message.reply_text(
            "🏪 Filial raqamini qayta kiriting:",
            reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
        )
        return REG_CONFIRM

    if txt != "✅ Ha, to'g'ri":
        await update.message.reply_text("Iltimos, tugmadan tanlang.", reply_markup=confirm_keyboard())
        return REG_LAVOZIM

    await update.message.reply_text(
        "👔 Lavozimingizni tanlang:",
        reply_markup=lavozim_keyboard(),
    )
    return REG_LAVOZIM + 1


async def reg_lavozim_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Lavozim qabul qilinadi, saqlanadi"""
    txt = update.message.text

    if txt == "⬅️ Orqaga":
        await update.message.reply_text(
            f"🏪 {ctx.user_data.get('reg_filial', '')} — shu filialda ishlaysizmi?",
            reply_markup=confirm_keyboard(),
        )
        return REG_LAVOZIM

    lavozim_map = {
        "💊 Farmatsevt": "Farmatsevt",
        "👑 Dorixona mudiri": "Dorixona mudiri",
        "🎓 Stajyor": "Stajyor",
    }

    if txt not in lavozim_map:
        await update.message.reply_text("Iltimos, tugmadan tanlang.", reply_markup=lavozim_keyboard())
        return REG_LAVOZIM + 1

    lavozim = lavozim_map[txt]
    ctx.user_data["reg_lavozim"] = lavozim

    # Saqlash
    ok = save_registration(
        user_id=update.effective_user.id,
        ismi=ctx.user_data.get("reg_ismi", ""),
        phone=ctx.user_data.get("reg_phone", ""),
        filial=ctx.user_data.get("reg_filial", ""),
        lavozim=lavozim,
    )

    if ok:
        ctx.user_data["reg_done"] = True
        await update.message.reply_text(
            f"🎉 *Ro'yxatdan o'tdingiz!*\n\n"
            f"👤 {ctx.user_data.get('reg_ismi', '')}\n"
            f"📱 {ctx.user_data.get('reg_phone', '')}\n"
            f"🏪 {ctx.user_data.get('reg_filial', '')}\n"
            f"👔 {lavozim}",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup([["⬅️ Asosiy menyu"]], resize_keyboard=True),
        )
    else:
        await update.message.reply_text(
            "⚠️ Xatolik yuz berdi. Qayta urinib ko'ring.",
            reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
        )

    from bot import main_keyboard, get_lang, MENU
    await update.message.reply_text("📋 Asosiy menyu", reply_markup=main_keyboard(get_lang(ctx)))
    return MENU


# ─── States ro'yxati ──────────────────────────────────────────────────────────

def get_reg_states():
    return {
        REG_PHONE: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, reg_ismi_handler),
        ],
        REG_FILIAL: [
            MessageHandler(filters.CONTACT, reg_phone_handler),
            MessageHandler(filters.TEXT & ~filters.COMMAND, reg_phone_handler),
        ],
        REG_CONFIRM: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, reg_filial_handler),
        ],
        REG_LAVOZIM: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, reg_confirm_handler),
        ],
        REG_LAVOZIM + 1: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, reg_lavozim_handler),
        ],
    }
