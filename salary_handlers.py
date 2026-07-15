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

logger = logging.getLogger(__name__)

# ─── Sozlamalar ───────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

PHARMACY_SHEET_ID = os.getenv("PHARMACY_SHEET_ID", "")
ADMIN_IDS = [709544046]

# Conversation states
SAL_WAIT_ZIP = 500


# ─── Google Sheets ────────────────────────────────────────────────────────────

def _get_client():
    creds_json = os.getenv("GOOGLE_CREDENTIALS")
    if creds_json:
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    return gspread.authorize(creds)


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
        mudir_map = get_mudir_map()
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
    }
