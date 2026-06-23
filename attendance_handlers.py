"""
attendance_handlers.py — Davomot uchun Telegram handlerlar
"""

from telegram import (
    Update, ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton,
)
from telegram.ext import (
    MessageHandler, CallbackQueryHandler,
    ContextTypes, filters,
)
from attendance import (
    ATT_PHONE, ATT_MENU, ATT_FILIAL_SELECT,
    ATT_LOCATION, ATT_ZAMENA_FILIAL, ATT_ZAMENA_LOCATION,
    get_farmatsevt, get_farmatsevt_by_userid, save_userid_to_sheet,
    write_attendance, get_filiallar_list,
    haversine_m, MAX_DISTANCE_KM, normalize_phone,
)

ATT_PASSWORD = 106   # Parol kutish state
ATT_PAROL = "офис"  # Universal parol

# ─── Klaviaturalar ────────────────────────────────────────────────────────────

def att_main_keyboard():
    return ReplyKeyboardMarkup([
        ["✅ Keldi", "🚪 Ketdi"],
        ["🔄 Zamena"],
        ["⬅️ Orqaga"],
    ], resize_keyboard=True)


def filial_inline_keyboard(filiallar: list):
    buttons = []
    row = []
    for f in filiallar:
        row.append(InlineKeyboardButton(f"#{f['filial']}", callback_data=f"att_fil_{f['filial']}"))
        if len(row) == 3:
            buttons.append(row)
            row = []
    if row:
        buttons.append(row)
    buttons.append([InlineKeyboardButton("⬅️ Orqaga", callback_data="att_back")])
    return InlineKeyboardMarkup(buttons)


def location_keyboard(btn_text="📍 Lokatsiyamni yuborish"):
    return ReplyKeyboardMarkup([
        [KeyboardButton(btn_text, request_location=True)],
        ["⬅️ Orqaga"],
    ], resize_keyboard=True)

def back_to_main_keyboard(language="uz"):
    return ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True)

# ─── 1. Davomot kirish — avval parol ─────────────────────────────────────────

async def att_enter(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Asosiy menyudan 'Davomot' bosilganda — parol so'raydi"""
    user_id = update.effective_user.id

    # Avval user_id bo'yicha tekshirish (oldin kirgan bo'lsa)
    if not ctx.user_data.get("att_farmatsevt"):
        farmatsevt = get_farmatsevt_by_userid(user_id)
        if farmatsevt:
            ctx.user_data["att_auth"] = True
            ctx.user_data["att_farmatsevt"] = farmatsevt
            ctx.user_data["att_phone"] = "saved"

    # Agar parol allaqachon tasdiqlangan bo'lsa
    if ctx.user_data.get("att_auth"):
        if ctx.user_data.get("att_farmatsevt"):
            return await _show_att_menu(update, ctx)
        return await _ask_phone(update, ctx)

    await update.message.reply_text(
        "🔐 Davomot tizimi\n\nParolni kiriting:",
        reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
    )
    return ATT_PASSWORD


async def att_password_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Parolni tekshiradi"""
    txt = update.message.text.strip()

    if txt == "⬅️ Orqaga":
        from bot import main_keyboard, get_lang, MENU
        language = get_lang(ctx)
        await update.message.reply_text("📋 Asosiy menyu", reply_markup=main_keyboard(language))
        return MENU

    if txt == ATT_PAROL:
        ctx.user_data["att_auth"] = True
        await update.message.reply_text("✅ Parol to'g'ri!")
        if ctx.user_data.get("att_phone"):
            return await _show_att_menu(update, ctx)
        return await _ask_phone(update, ctx)
    else:
        await update.message.reply_text(
            "❌ Parol noto'g'ri. Qayta urinib ko'ring:",
        )
        return ATT_PASSWORD


async def _ask_phone(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    kb = ReplyKeyboardMarkup([
        [KeyboardButton("📱 Telefon raqamimni yuborish", request_contact=True)],
        ["⬅️ Orqaga"],
    ], resize_keyboard=True)
    await update.message.reply_text(
        "👤 Telefon raqamingizni yuboring:",
        reply_markup=kb,
    )
    return ATT_PHONE

# ─── 2. Telefon ───────────────────────────────────────────────────────────────

async def att_phone_received(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.message.text == "⬅️ Orqaga":
        from bot import main_keyboard, get_lang, MENU
        language = get_lang(ctx)
        await update.message.reply_text("📋 Asosiy menyu", reply_markup=main_keyboard(language))
        return MENU

    contact = update.message.contact
    if not contact:
        await update.message.reply_text("❌ Iltimos, tugma orqali raqamingizni yuboring.")
        return ATT_PHONE

    phone = normalize_phone(contact.phone_number)
    farmatsevt = get_farmatsevt(phone)

    if not farmatsevt:
        await update.message.reply_text(
            f"❌ *{phone}* raqami tizimda topilmadi.\n"
            "Administratorga murojaat qiling.",
            parse_mode="Markdown",
        )
        return ATT_PHONE

    ctx.user_data["att_phone"] = phone
    ctx.user_data["att_farmatsevt"] = farmatsevt

    # TelegramID ni saqlash — keyingi safar telefon so'ralmaydi
    user_id = update.effective_user.id
    save_userid_to_sheet(user_id, phone)

    await update.message.reply_text(
        f"✅ Xush kelibsiz, *{farmatsevt['ismi']}*!\n"
        f"🏪 Filial: #{farmatsevt['filial']}",
        parse_mode="Markdown",
    )
    return await _show_att_menu(update, ctx)


async def _show_att_menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    farmatsevt = ctx.user_data.get("att_farmatsevt", {})
    await update.message.reply_text(
        f"📋 *Davomot menyu*\n"
        f"👤 {farmatsevt.get('ismi', '')}\n"
        f"🏪 Filial: #{farmatsevt.get('filial', '')}",
        reply_markup=att_main_keyboard(),
        parse_mode="Markdown",
    )
    return ATT_MENU

# ─── 3. Davomot menyusi ───────────────────────────────────────────────────────

async def att_menu_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text

    if txt == "⬅️ Orqaga":
        from bot import main_keyboard, get_lang, MENU
        language = get_lang(ctx)
        await update.message.reply_text("📋 Asosiy menyu", reply_markup=main_keyboard(language))
        return MENU

    elif txt in ["✅ Keldi", "🚪 Ketdi"]:
        ctx.user_data["att_action"] = "keldi" if txt == "✅ Keldi" else "ketdi"
        ctx.user_data["att_zamena"] = False
        await update.message.reply_text(
            "📍 Lokatsiyangizni yuboring:\n_(100 metr radiusda bo'lishingiz kerak)_",
            reply_markup=location_keyboard(),
            parse_mode="Markdown",
        )
        return ATT_LOCATION

    elif txt == "🔄 Zamena":
        ctx.user_data["att_zamena"] = True
        filiallar = get_filiallar_list()
        if not filiallar:
            await update.message.reply_text("❌ Filiallar ro'yxati yuklanmadi.")
            return ATT_MENU
        ctx.user_data["att_filiallar"] = filiallar
        await update.message.reply_text(
            "🔄 *Zamena rejimi*\nFilial raqamini tanlang:",
            reply_markup=filial_inline_keyboard(filiallar),
            parse_mode="Markdown",
        )
        return ATT_ZAMENA_FILIAL

    return ATT_MENU

# ─── 4. Lokatsiya tekshiruvi ──────────────────────────────────────────────────

async def att_location_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.message.text == "⬅️ Orqaga":
        return await _show_att_menu(update, ctx)

    if not update.message.location:
        await update.message.reply_text("❌ Iltimos, lokatsiya tugmasini bosing.")
        return ATT_LOCATION

    ulat = update.message.location.latitude
    ulon = update.message.location.longitude

    farmatsevt = ctx.user_data.get("att_farmatsevt", {})
    dist = haversine_m(ulat, ulon, farmatsevt.get("lat", 0), farmatsevt.get("lon", 0))

    if dist > MAX_DISTANCE_KM * 1000:
        await update.message.reply_text(
            f"❌ Siz filialdan *{dist:.0f} metr* uzoqdasiz.\n"
            f"Maksimal ruxsat: *100 metr*.",
            parse_mode="Markdown",
            reply_markup=location_keyboard(),
        )
        return ATT_LOCATION

    action = ctx.user_data.get("att_action", "keldi")
    ok = write_attendance(farmatsevt, action, zamena=False)
    now_str = __import__("datetime").datetime.now().strftime("%H:%M")
    emoji = "✅" if action == "keldi" else "🚪"

    if ok:
        await update.message.reply_text(
            f"{emoji} *{farmatsevt['ismi']}* — {action}!\n"
            f"🕐 Vaqt: {now_str}\n"
            f"🏪 Filial: #{farmatsevt['filial']}\n"
            f"📏 Masofa: {dist:.0f} m",
            parse_mode="Markdown",
            reply_markup=att_main_keyboard(),
        )
    else:
        await update.message.reply_text("⚠️ Xatolik. Qayta urinib ko'ring.", reply_markup=att_main_keyboard())
    return ATT_MENU

# ─── 5. Zamena ───────────────────────────────────────────────────────────────

async def att_zamena_filial_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()

    if q.data == "att_back":
        await q.message.reply_text("📋 Davomot menyu", reply_markup=att_main_keyboard())
        return ATT_MENU

    filial_no = q.data.replace("att_fil_", "")
    filiallar = ctx.user_data.get("att_filiallar", [])
    selected = next((f for f in filiallar if f["filial"] == filial_no), None)

    if not selected:
        await q.message.reply_text("❌ Filial topilmadi.")
        return ATT_ZAMENA_FILIAL

    ctx.user_data["att_zamena_filial"] = selected
    ctx.user_data["att_action"] = "keldi"

    await q.message.reply_text(
        f"🔄 *Zamena* — Filial #{filial_no}\n📍 Lokatsiyangizni yuboring:",
        reply_markup=location_keyboard("📍 Zamena lokatsiyamni yuborish"),
        parse_mode="Markdown",
    )
    return ATT_ZAMENA_LOCATION


async def att_zamena_location_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.message.text == "⬅️ Orqaga":
        return await _show_att_menu(update, ctx)

    if not update.message.location:
        await update.message.reply_text("❌ Iltimos, lokatsiya tugmasini bosing.")
        return ATT_ZAMENA_LOCATION

    ulat = update.message.location.latitude
    ulon = update.message.location.longitude

    zamena_filial = ctx.user_data.get("att_zamena_filial", {})
    dist = haversine_m(ulat, ulon, zamena_filial.get("lat", 0), zamena_filial.get("lon", 0))

    if dist > MAX_DISTANCE_KM * 1000:
        await update.message.reply_text(
            f"❌ Zamena filialidan *{dist:.0f} metr* uzoqdasiz.\nMaksimal: *100 metr*.",
            parse_mode="Markdown",
            reply_markup=location_keyboard("📍 Zamena lokatsiyamni yuborish"),
        )
        return ATT_ZAMENA_LOCATION

    farmatsevt = ctx.user_data.get("att_farmatsevt", {})
    zamena_info = {**farmatsevt, "filial": zamena_filial["filial"]}
    ok = write_attendance(zamena_info, "keldi", zamena=True)
    now_str = __import__("datetime").datetime.now().strftime("%H:%M")

    if ok:
        await update.message.reply_text(
            f"🔄 *Zamena tasdiqlandi!*\n"
            f"👤 {farmatsevt['ismi']}\n"
            f"🏪 Zamena filial: #{zamena_filial['filial']}\n"
            f"🕐 Vaqt: {now_str}\n"
            f"📏 Masofa: {dist:.0f} m\n\n"
            f"_Jadvalda sariq rangda ko'rinadi_",
            parse_mode="Markdown",
            reply_markup=att_main_keyboard(),
        )
    else:
        await update.message.reply_text("⚠️ Xatolik.", reply_markup=att_main_keyboard())

    return ATT_MENU

# ─── Handler ro'yxati ─────────────────────────────────────────────────────────

def get_att_states():
    return {
        ATT_PASSWORD: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_password_handler),
        ],
        ATT_PHONE: [
            MessageHandler(filters.CONTACT, att_phone_received),
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_phone_received),
        ],
        ATT_MENU: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_menu_handler),
        ],
        ATT_LOCATION: [
            MessageHandler(filters.LOCATION, att_location_handler),
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_location_handler),
        ],
        ATT_ZAMENA_FILIAL: [
            CallbackQueryHandler(att_zamena_filial_handler, pattern="^att_"),
        ],
        ATT_ZAMENA_LOCATION: [
            MessageHandler(filters.LOCATION, att_zamena_location_handler),
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_zamena_location_handler),
        ],
    }
