"""
salary_handlers.py — Oylik fayllarni mudirlariga yuborish moduli

Jarayon:
1. Admin /send_salaries yozadi
2. Bot ZIP faylni so'raydi
3. Admin ZIP yuboradi
4. Bot ZIP ochadi → har bir .xlsx faylni topadi
5. Fayl nomi → Filial nomi → Mudir TelegramID → yuboradi
6. Natija xabarini adminга qaytaradi

Fayl nomi formatlari qabul qilinadi:
- АСКАЯ.xlsx → "АСКАЯ" filiali
- 1-ГОР БОЛЬНИ....xlsx → "1-ГОР БОЛЬНИЦА" filiali
"""

import os
import io
import re
import json
import zipfile
import logging

from telegram import Update
from telegram.ext import (
    MessageHandler, CommandHandler,
    ContextTypes, filters, ConversationHandler,
)
from google.oauth2.service_account import Credentials
import gspread
from attendance import run_read, run_write

logger = logging.getLogger(__name__)

# ─── Sozlamalar ───────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

PHARMACY_SHEET_ID = os.getenv("PHARMACY_SHEET_ID", "")
SALARY_SHEET_ID = os.getenv("SALARY_SHEET_ID", "")
FIRMS_SHEET_ID = os.getenv("FIRMS_SHEET_ID", "")
ADMIN_IDS = [709544046]
PAYMENTS_PAROL = "офис"  # Davomat bo'limi bilan bir xil umumiy parol

# Conversation states
SAL_WAIT_ZIP = 500
REPORTS_MENU = 501
PAYMENTS_MENU = 502
PAYMENTS_PASSWORD = 503


# ─── Google Sheets ────────────────────────────────────────────────────────────

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


def get_mudir_map() -> dict:
    """
    Farmatsevtlar Sheets dan filial nomi → mudir TelegramID lug'atini qaytaradi.
    Faqat "Dorixona mudiri" lavozimli xodimlar.
    
    Returns: {"АСКАЯ": "123456789", "АНТЕЙ": "987654321", ...}
    """
    try:
        client = _get_client()
        ws = client.open_by_key(PHARMACY_SHEET_ID).worksheet("Farmatsevtlar")
        records = ws.get_all_records()

        mudir_map = {}
        for row in records:
            lavozim = str(row.get("Lavozim", "")).strip().lower()
            if "mudiri" not in lavozim and "директор" not in lavozim.lower():
                continue

            filial_full = str(row.get("Filial", "")).strip()
            telegram_id = str(row.get("TelegramID", "")).strip()

            if not filial_full or not telegram_id or telegram_id in ("", "0"):
                continue

            # Filial nomidan "0 - АСКАЯ" → "АСКАЯ" ajratish
            # Ham to'liq nom, ham qisqa nom bilan saqlash
            mudir_map[filial_full.upper()] = telegram_id

            # Raqamsiz nom ham (masalan: "АСКАЯ")
            short = re.sub(r"^\d+\s*[-–]\s*", "", filial_full).strip().upper()
            if short:
                mudir_map[short] = telegram_id

        logger.info(f"[SAL] {len(mudir_map)} ta mudir topildi")
        return mudir_map

    except Exception as e:
        logger.error(f"[SAL] Mudir map xato: {e}")
        return {}


def find_telegram_id(filename: str, mudir_map: dict) -> str | None:
    """
    Fayl nomidan filial nomini topadi va TelegramID ni qaytaradi.
    
    "АСКАЯ.xlsx" → "АСКАЯ" → TelegramID
    "1-ГОР БОЛЬНИ....xlsx" → shunga o'xshash nom qidiradi
    """
    # .xlsx kengaytmasini olib tashlash
    name = re.sub(r"\.xlsx?$", "", filename, flags=re.IGNORECASE).strip().upper()

    # To'g'ridan-to'g'ri mos kelsa
    if name in mudir_map:
        return mudir_map[name]

    # Qisman mos kelishni qidirish (fayl nomi qisqartirilgan bo'lishi mumkin)
    for key, tid in mudir_map.items():
        if name in key or key in name:
            return tid

    # Birinchi so'z bo'yicha qidirish
    first_word = name.split()[0] if name.split() else ""
    if first_word:
        for key, tid in mudir_map.items():
            if key.startswith(first_word):
                return tid

    return None


SALARY_WS_NAME = "Oylik"
AKSIYA_WS_NAME = "Aksiya"

# Maosh jadvalidagi ustunlar (pozitsiya bo'yicha, 1-indeksda):
# A=1 Filial/Ismi | B=2 Telefon | C=3 Reja(keyingi oy) | D=4 Reja(joriy oy)
# E=5 Savdo | F=6 Rejadan farq | G=7 Foiz | H=8 Oylik % (bonus)
# I=9 Fiksa | J=10 Reja bonusi | K=11 Avans | L=12 Pereuchyot shtraf
# M=13 Kech/erta shtraf | N=14 Srok shtraf | O=15 Umumiy summa
# P=16 Plastik kartaga tushadigan
SAL_COL_FILIAL_ISMI = 1
SAL_COL_TELEFON = 2
SAL_COL_REJA_KEYINGI = 3
SAL_COL_REJA_JORIY = 4
SAL_COL_SAVDO = 5
SAL_COL_FARQ = 6
SAL_COL_FOIZ = 7
SAL_COL_OYLIK_PERCENT = 8
SAL_COL_FIKSA = 9
SAL_COL_REJA_BONUS = 10
SAL_COL_AVANS = 11
SAL_COL_SHTRAF_PEREUCHYOT = 12
SAL_COL_SHTRAF_VAQT = 13
SAL_COL_SHTRAF_SROK = 14
SAL_COL_JAMI = 15
SAL_COL_KARTA = 16


def _sal_normalize_phone(phone) -> str:
    digits = re.sub(r"\D", "", str(phone))
    if digits.startswith("998"):
        return digits
    if digits.startswith("0"):
        return "998" + digits[1:]
    if len(digits) == 9:
        return "998" + digits
    return digits


def _get_phone_by_telegram_id(telegram_id) -> str:
    """Farmatsevtlar jadvalidan TelegramID bo'yicha telefon raqamini topadi."""
    try:
        client = _get_client()
        ws = client.open_by_key(PHARMACY_SHEET_ID).worksheet("Farmatsevtlar")
        records = ws.get_all_records()
        tid = str(telegram_id).strip()
        for row in records:
            if str(row.get("TelegramID", "")).strip() == tid:
                return str(row.get("Telefon", "")).strip()
        return ""
    except Exception as e:
        logger.error(f"[MAOSH] Telefon qidirish xato: {e}")
        return ""


def get_farmatsevt_salary(telegram_id) -> dict | None:
    """
    Xodimning TelegramID'si orqali telefon raqamini topadi, so'ng
    "Oylik" va "Aksiya" varaqlaridan (SALARY_SHEET_ID) o'sha telefon
    raqamiga mos qatorlarni qidiradi va ikkalasini birlashtirib to'liq
    hisobotni qaytaradi.

    Jadval tuzilishi (ikkala varaqda ham bir xil, filial sarlavha
    qatorlari + xodim qatorlari aralash holda, ustunlar POZITSIYA
    bo'yicha o'qiladi):
        A: Filial nomi (sarlavha qatorida) yoki Xodim ismi (xodim qatorida)
        B: Telefon (faqat xodim qatorida bo'ladi)
        C/D: Reja (faqat filial sarlavhasida)
        E: Bir oylik savdo
        G: Foiz | H: Oylik % (savdodan bonus)
        I: Fiksa (asosiy oylik) | J: Rejaga chiqqani uchun bonus
        K: Avans olganlar | L/M/N: turli shtraflar
        O: Umumiy summa (qo'lga tegadigan yakuniy summa)

    MUHIM: bu varaqlar oylik davomida QO'LDA tozalanib, keyingi oy
    ma'lumotlari bilan qayta to'ldiriladi — shuning uchun oy nomi bilan
    emas, doim bitta doimiy "Oylik"/"Aksiya" nomi bilan ochiladi.

    Qaytaradi to'liq breakdown dict yoki None (ikkalasida ham topilmasa).
    """
    def _find_row_by_phone(ws, target_phone):
        all_values = ws.get_all_values()
        current_filial = ""
        for row in all_values[1:]:
            if not row or not row[0]:
                continue

            def _cell(col, _row=row):
                idx = col - 1
                return _row[idx] if idx < len(_row) else ""

            telefon_cell = str(_cell(SAL_COL_TELEFON)).strip()
            if not telefon_cell:
                current_filial = str(_cell(SAL_COL_FILIAL_ISMI)).strip()
                continue
            if _sal_normalize_phone(telefon_cell) != target_phone:
                continue

            def _num(col, _row=row):
                v = _row[col - 1] if col - 1 < len(_row) else ""
                if v == "" or v is None:
                    return 0
                try:
                    return float(str(v).replace(",", ".").replace(" ", ""))
                except Exception:
                    return 0

            return {
                "ismi": str(_cell(SAL_COL_FILIAL_ISMI)).strip(),
                "filial": current_filial,
                "savdo": _num(SAL_COL_SAVDO),
                "foiz": _num(SAL_COL_FOIZ),
                "oylik_percent_bonus": _num(SAL_COL_OYLIK_PERCENT),
                "fiksa": _num(SAL_COL_FIKSA),
                "reja_bonus": _num(SAL_COL_REJA_BONUS),
                "avans": _num(SAL_COL_AVANS),
                "shtraf_pereuchyot": _num(SAL_COL_SHTRAF_PEREUCHYOT),
                "shtraf_vaqt": _num(SAL_COL_SHTRAF_VAQT),
                "shtraf_srok": _num(SAL_COL_SHTRAF_SROK),
                "jami": _num(SAL_COL_JAMI),
                "karta": _num(SAL_COL_KARTA),
            }
        return None

    try:
        phone = _get_phone_by_telegram_id(telegram_id)
        if not phone:
            return None
        target_phone = _sal_normalize_phone(phone)

        client = _get_client()
        sh = client.open_by_key(SALARY_SHEET_ID)

        oylik_data = None
        try:
            ws_oylik = sh.worksheet(SALARY_WS_NAME)
            oylik_data = _find_row_by_phone(ws_oylik, target_phone)
        except gspread.exceptions.WorksheetNotFound:
            pass

        aksiya_data = None
        try:
            ws_aksiya = sh.worksheet(AKSIYA_WS_NAME)
            aksiya_data = _find_row_by_phone(ws_aksiya, target_phone)
        except gspread.exceptions.WorksheetNotFound:
            pass

        if not oylik_data and not aksiya_data:
            return None

        base = oylik_data or {}
        result = dict(base)
        result["ismi"] = (oylik_data or aksiya_data).get("ismi", "")
        result["filial"] = (oylik_data or aksiya_data).get("filial", "")
        result["aksiya_jami"] = aksiya_data["jami"] if aksiya_data else 0
        result["oylik_jami"] = oylik_data["jami"] if oylik_data else 0
        result["yakuniy_jami"] = result["oylik_jami"] + result["aksiya_jami"]
        return result

    except Exception as e:
        logger.error(f"[MAOSH] get_farmatsevt_salary xato: {e}")
        return None


def get_firms_report() -> list:
    """
    "Firmalar to'lovlari" Google Sheets'idan (FIRMS_SHEET_ID) barcha
    firmalar ro'yxatini o'qiydi.

    Jadval tuzilishi (bitta oddiy varaq, oy bo'yicha bo'linmagan):
        Firma nomi | Summa | Holati

    Qaytaradi: [{"firma", "summa", "holati"}, ...]
    """
    try:
        client = _get_client()
        sh = client.open_by_key(FIRMS_SHEET_ID)
        ws = sh.sheet1
        records = ws.get_all_records()

        result = []
        for row in records:
            firma = str(row.get("Firma nomi", "")).strip()
            if not firma:
                continue
            result.append({
                "firma": firma,
                "summa": row.get("Summa", ""),
                "holati": str(row.get("Holati", "")).strip(),
            })
        return result

    except Exception as e:
        logger.error(f"[FIRMS] get_firms_report xato: {e}")
        return []


def payments_keyboard(language: str = "uz"):
    from telegram import ReplyKeyboardMarkup
    if language == "ru":
        return ReplyKeyboardMarkup([
            ["📦 Отправить ZIP файлы"],
            ["🏢 Отчёт по оплатам фирмам"],
            ["⬅️ Назад"],
        ], resize_keyboard=True)
    return ReplyKeyboardMarkup([
        ["📦 ZIP orqali yuborish"],
        ["🏢 Firmalarga to'lovlar hisoboti"],
        ["⬅️ Orqaga"],
    ], resize_keyboard=True)


async def payments_menu_enter(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    language = ctx.user_data.get("lang", "uz")
    text = (
        "📊 *Отчёт и оплаты*\n\nВыберите раздел:" if language == "ru"
        else "📊 *Отчёт va to'lovlar*\n\nBo'limni tanlang:"
    )
    await update.message.reply_text(
        text, parse_mode="Markdown", reply_markup=payments_keyboard(language)
    )
    return PAYMENTS_MENU


async def payments_menu_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    language = ctx.user_data.get("lang", "uz")
    txt = update.message.text.strip() if update.message and update.message.text else ""

    back_txt = "⬅️ Назад" if language == "ru" else "⬅️ Orqaga"
    zip_txt = "📦 Отправить ZIP файлы" if language == "ru" else "📦 ZIP orqali yuborish"
    firms_txt = "🏢 Отчёт по оплатам фирмам" if language == "ru" else "🏢 Firmalarga to'lovlar hisoboti"

    if txt == back_txt:
        await update.message.reply_text(
            "📊 *Hisobotlar va to'lovlar*\n\nBo'limni tanlang:" if language == "uz"
            else "📊 *Отчёты и оплаты*\n\nВыберите раздел:",
            parse_mode="Markdown",
            reply_markup=reports_keyboard(language),
        )
        return REPORTS_MENU

    elif txt == zip_txt:
        if update.effective_user.id not in ADMIN_IDS:
            await update.message.reply_text(
                "❌ Bu bo'lim faqat administrator uchun.",
                reply_markup=payments_keyboard(language),
            )
            return PAYMENTS_MENU
        return await cmd_send_salaries(update, ctx)

    elif txt == firms_txt:
        if ctx.user_data.get("payments_auth"):
            return await _show_firms_report(update, ctx)
        from telegram import ReplyKeyboardMarkup
        await update.message.reply_text(
            "🔐 Parolni kiriting:" if language == "uz" else "🔐 Введите пароль:",
            reply_markup=ReplyKeyboardMarkup([[back_txt]], resize_keyboard=True),
        )
        return PAYMENTS_PASSWORD

    # Tanilmagan matn
    await update.message.reply_text(
        "📊 *Отчёт va to'lovlar*\n\nBo'limni tanlang:" if language == "uz"
        else "📊 *Отчёт и оплаты*\n\nВыберите раздел:",
        parse_mode="Markdown",
        reply_markup=payments_keyboard(language),
    )
    return PAYMENTS_MENU


async def payments_password_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    language = ctx.user_data.get("lang", "uz")
    txt = update.message.text.strip() if update.message and update.message.text else ""
    back_txt = "⬅️ Назад" if language == "ru" else "⬅️ Orqaga"

    if txt == back_txt:
        await update.message.reply_text(
            "📊 *Отчёт va to'lovlar*\n\nBo'limni tanlang:" if language == "uz"
            else "📊 *Отчёт и оплаты*\n\nВыберите раздел:",
            parse_mode="Markdown",
            reply_markup=payments_keyboard(language),
        )
        return PAYMENTS_MENU

    if txt == PAYMENTS_PAROL:
        ctx.user_data["payments_auth"] = True
        return await _show_firms_report(update, ctx)
    else:
        await update.message.reply_text(
            "❌ Parol noto'g'ri. Qayta urinib ko'ring:" if language == "uz"
            else "❌ Неверный пароль. Попробуйте снова:",
        )
        return PAYMENTS_PASSWORD


async def _show_firms_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    language = ctx.user_data.get("lang", "uz")
    firms = await run_read(get_firms_report)

    if not firms:
        text = (
            "❌ Hozircha firmalar ro'yxati topilmadi." if language == "uz"
            else "❌ Список фирм пока не найден."
        )
        await update.message.reply_text(text, reply_markup=payments_keyboard(language))
        return PAYMENTS_MENU

    lines = ["🏢 *Firmalarga to'lovlar hisoboti*\n"] if language == "uz" \
        else ["🏢 *Отчёт по оплатам фирмам*\n"]

    for f in firms:
        holati = f["holati"].strip().lower()
        if holati in ("to'langan", "оплачено", "✅", "ha", "да"):
            belgi = "✅"
        elif holati in ("to'lanmagan", "не оплачено", "❌", "yo'q", "нет"):
            belgi = "❌"
        else:
            belgi = "⏳"
        lines.append(f"{belgi} *{f['firma']}* — {f['summa']} ({f['holati']})")

    await update.message.reply_text(
        "\n".join(lines), parse_mode="Markdown", reply_markup=payments_keyboard(language)
    )
    return PAYMENTS_MENU


def reports_keyboard(language: str = "uz"):
    """Hisobotlar va to'lovlar bo'limining pastki menyusi."""
    from telegram import ReplyKeyboardMarkup
    if language == "ru":
        return ReplyKeyboardMarkup([
            ["💰 Зарплата и бонусы"],
            ["📊 Отчёт и оплаты"],
            ["⬅️ Назад"],
        ], resize_keyboard=True)
    return ReplyKeyboardMarkup([
        ["💰 Maosh va aksiyalar"],
        ["📊 Отчёт va to'lovlar"],
        ["⬅️ Orqaga"],
    ], resize_keyboard=True)


async def reports_menu_enter(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """'📊 Hisobotlar va to'lovlar' tugmasi bosilganda chaqiriladi."""
    language = ctx.user_data.get("lang", "uz")
    text = (
        "📊 *Отчёты и оплаты*\n\nВыберите раздел:" if language == "ru"
        else "📊 *Hisobotlar va to'lovlar*\n\nBo'limni tanlang:"
    )
    await update.message.reply_text(
        text, parse_mode="Markdown", reply_markup=reports_keyboard(language)
    )
    return REPORTS_MENU


async def reports_menu_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Hisobotlar va to'lovlar pastki menyusidagi tugmalarni ishlaydi."""
    from bot import main_keyboard, T, MENU
    language = ctx.user_data.get("lang", "uz")
    txt = update.message.text.strip() if update.message and update.message.text else ""

    back_txt = "⬅️ Назад" if language == "ru" else "⬅️ Orqaga"
    salary_txt = "💰 Зарплата и бонусы" if language == "ru" else "💰 Maosh va aksiyalar"
    payments_txt = "📊 Отчёт и оплаты" if language == "ru" else "📊 Отчёт va to'lovlar"

    if txt == back_txt:
        await update.message.reply_text(
            T[language]["menu"], reply_markup=main_keyboard(language), parse_mode="Markdown"
        )
        return MENU

    elif txt == salary_txt:
        user_id = update.effective_user.id
        data = await run_read(get_farmatsevt_salary, user_id)

        if not data:
            msg = (
                "❌ Sizning telefon raqamingiz bo'yicha joriy oy uchun maosh "
                "ma'lumoti topilmadi.\n\nAgar siz ro'yxatdan o'tgan xodim "
                "bo'lsangiz, buxgalteriya hali ma'lumotni kiritmagan bo'lishi "
                "mumkin — birozdan so'ng qayta urinib ko'ring yoki "
                "administratorga murojaat qiling."
            )
            await update.message.reply_text(msg, reply_markup=reports_keyboard(language))
            return REPORTS_MENU

        def _fmt(n):
            try:
                n = float(n)
                if n == int(n):
                    return f"{int(n):,}".replace(",", " ")
                return f"{n:,.2f}".replace(",", " ")
            except Exception:
                return str(n)

        g = lambda k, d=0: data.get(k, d)

        lines = [
            f"💰 *Maosh va aksiyalar*\n",
            f"👤 {data['ismi']}",
            f"🏪 Filial: {data['filial']}\n",
            f"📊 Savdo: {_fmt(g('savdo'))} so'm",
        ]
        if g("foiz"):
            lines.append(f"📈 Foiz: {g('foiz') * 100:.1f}%")
        if g("oylik_percent_bonus"):
            lines.append(f"🎁 Savdodan bonus: {_fmt(g('oylik_percent_bonus'))} so'm")
        if g("fiksa"):
            lines.append(f"💵 Fiksa (asosiy oylik): {_fmt(g('fiksa'))} so'm")
        if g("reja_bonus"):
            lines.append(f"🏆 Rejaga chiqqani uchun bonus: {_fmt(g('reja_bonus'))} so'm")

        deductions = []
        if g("avans"):
            deductions.append(f"➖ Avans: {_fmt(g('avans'))} so'm")
        if g("shtraf_pereuchyot"):
            deductions.append(f"➖ Pereuchyot shtrafi: {_fmt(g('shtraf_pereuchyot'))} so'm")
        if g("shtraf_vaqt"):
            deductions.append(f"➖ Kech ochilgan/erta yopilgan shtrafi: {_fmt(g('shtraf_vaqt'))} so'm")
        if g("shtraf_srok"):
            deductions.append(f"➖ Srok shtrafi: {_fmt(g('shtraf_srok'))} so'm")
        if deductions:
            lines.append("")
            lines.extend(deductions)

        lines.append("")
        lines.append(f"💵 Oylik: {_fmt(g('oylik_jami'))} so'm")
        lines.append(f"🎁 Aksiya: {_fmt(g('aksiya_jami'))} so'm")
        lines.append(f"💰 *JAMI: {_fmt(g('yakuniy_jami'))} so'm*")
        if g("karta"):
            lines.append(f"💳 Plastik kartaga: {_fmt(g('karta'))} so'm")

        text = "\n".join(lines)
        await update.message.reply_text(
            text, parse_mode="Markdown", reply_markup=reports_keyboard(language)
        )
        return REPORTS_MENU

    elif txt == payments_txt:
        return await payments_menu_enter(update, ctx)

    # Tanilmagan matn — menyuni qayta ko'rsatamiz
    await update.message.reply_text(
        "📊 *Hisobotlar va to'lovlar*\n\nBo'limni tanlang:" if language == "uz"
        else "📊 *Отчёты и оплаты*\n\nВыберите раздел:",
        parse_mode="Markdown",
        reply_markup=reports_keyboard(language),
    )
    return REPORTS_MENU


# ─── Handlerlar ───────────────────────────────────────────────────────────────

async def cmd_send_salaries(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/send_salaries — Admin buyrug'i: ZIP faylni so'raydi."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return

    await update.message.reply_text(
        "📦 *Oylik fayllar ZIP arxivini yuboring*\n\n"
        "Fayl nomlari filial nomiga mos bo'lishi kerak:\n"
        "`АСКАЯ.xlsx`, `АНТЕЙ.xlsx` va h.k.",
        parse_mode="Markdown",
    )
    return SAL_WAIT_ZIP


async def sal_receive_zip(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """ZIP faylni qabul qiladi va mudirlariga yuboradi."""
    if update.effective_user.id not in ADMIN_IDS:
        return

    if not update.message.document:
        await update.message.reply_text("❌ Iltimos, ZIP fayl yuboring.")
        return SAL_WAIT_ZIP

    doc = update.message.document
    if not doc.file_name.lower().endswith(".zip"):
        await update.message.reply_text("❌ Faqat ZIP fayl qabul qilinadi.")
        return SAL_WAIT_ZIP

    msg = await update.message.reply_text("⏳ ZIP ochilmoqda va fayllar yuborilmoqda...")

    try:
        # ZIP faylni yuklab olish
        file = await doc.get_file()
        zip_bytes = await file.download_as_bytearray()

        # Mudirlar ro'yxatini olish
        mudir_map = await run_read(get_mudir_map)
        if not mudir_map:
            await msg.edit_text("❌ Mudirlar ro'yxati topilmadi. Farmatsevtlar Sheets ni tekshiring.")
            return

        # ZIP ochish va yuborish
        sent = 0
        not_found = 0
        errors = 0
        not_found_list = []

        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            xlsx_files = [
                name for name in zf.namelist()
                if name.lower().endswith(".xlsx") and not name.startswith("__")
            ]

            total = len(xlsx_files)
            await msg.edit_text(f"⏳ {total} ta fayl topildi. Yuborilmoqda...")

            for i, fname in enumerate(xlsx_files):
                # Fayl nomidan papka yo'lini olib tashlash
                base_name = os.path.basename(fname)

                # TelegramID topish
                tid = find_telegram_id(base_name, mudir_map)

                if not tid:
                    not_found += 1
                    not_found_list.append(base_name)
                    logger.warning(f"[SAL] Topilmadi: {base_name}")
                    continue

                try:
                    # Faylni o'qish va yuborish
                    file_data = zf.read(fname)
                    file_io = io.BytesIO(file_data)
                    file_io.name = base_name

                    await ctx.bot.send_document(
                        chat_id=int(tid),
                        document=file_io,
                        filename=base_name,
                        caption=f"📊 Oylik hisobot\n📁 {base_name}",
                    )
                    sent += 1
                    logger.info(f"[SAL] Yuborildi: {base_name} → {tid}")

                except Exception as e:
                    errors += 1
                    logger.error(f"[SAL] Yuborishda xato {base_name}: {e}")

                # Har 10 faylda progress yangilash
                if (i + 1) % 10 == 0:
                    await msg.edit_text(
                        f"⏳ Yuborilmoqda... {i+1}/{total}\n"
                        f"✅ Yuborildi: {sent} ta"
                    )

        # Yakuniy natija
        lines = [
            f"✅ *Yuborish tugadi!*\n",
            f"📤 Yuborildi: *{sent}* ta",
            f"❌ Xato: *{errors}* ta",
            f"🔍 Topilmadi: *{not_found}* ta",
        ]

        if not_found_list:
            lines.append(f"\n*Topilmagan fayllar:*")
            for name in not_found_list[:20]:
                lines.append(f"  • {name}")
            if len(not_found_list) > 20:
                lines.append(f"  ... va yana {len(not_found_list)-20} ta")

        await msg.edit_text("\n".join(lines), parse_mode="Markdown")

    except zipfile.BadZipFile:
        await msg.edit_text("❌ ZIP fayl buzilgan yoki noto'g'ri format.")
    except Exception as e:
        await msg.edit_text(f"❌ Xato: {e}")
        logger.error(f"[SAL] Umumiy xato: {e}")


async def sal_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Bekor qilish."""
    from bot import main_keyboard, get_lang, MENU
    await update.message.reply_text(
        "❌ Bekor qilindi.",
        reply_markup=main_keyboard(get_lang(ctx))
    )
    return MENU


# ─── States ───────────────────────────────────────────────────────────────────

def get_sal_states():
    return {
        SAL_WAIT_ZIP: [
            MessageHandler(filters.Document.ALL, sal_receive_zip),
            MessageHandler(filters.TEXT & ~filters.COMMAND, sal_cancel),
        ],
        PAYMENTS_MENU: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, payments_menu_handler),
        ],
        PAYMENTS_PASSWORD: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, payments_password_handler),
        ],
    }
