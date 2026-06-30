"""
attendance_handlers.py — Davomot uchun Telegram handlerlar

KIRISH TARTIBI:
  /start → "📋 Davomat" → Kod kiriting (masalan: F125678)
         → Topilsa: xush kelibsiz, menyu
         → Topilmasa: xato xabar

Kod formati: F{filial_raqami}{tel_oxirgi_4_raqam}
Misol: Filial #12, tel ...5678 → F125678

Ikkinchi marta kirishda Telegram ID saqlanadi → kod so'ralmaydi.
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
    get_farmatsevt_by_userid, get_farmatsevt_by_code,
    save_userid_by_code, write_attendance, get_filiallar_list,
    haversine_m, MAX_DISTANCE_KM,
    sync_pharmacists, fill_codes_in_sheet,
)

# ATT_PHONE state endi "kod kutish" uchun ishlatiladi
ATT_PASSWORD = 106
ATT_SELECT_WHO = 107  # Farmatsevt tanlash state

def generate_filial_kod(filial: str) -> str:
    """Filial nomidan raqamni ajratadi: '6 - ЮНУСАБАД 7' → '6'"""
    import re
    m = re.match(r"^(\d+)", str(filial).strip())
    return m.group(1) if m else re.sub(r"\D", "", str(filial))[:3]

# Admin Telegram ID lari
ADMIN_IDS = [709544046]


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
        row.append(InlineKeyboardButton(
            f"#{f['filial']}", callback_data=f"att_fil_{f['filial']}"
        ))
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


# ─── 1. Davomot kirish ────────────────────────────────────────────────────────

async def att_enter(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    "📋 Davomat" tugmasi bosilganda:
    1. TelegramID bo'yicha foydalanuvchini topadi (saqlangan bo'lsa)
    2. Ismini ko'rsatib, filial kodini so'raydi (har safar)
    3. Kod + ID mos kelsa → kiradi
    """
    # Har safar sessiyani tozalash
    ctx.user_data.pop("att_farmatsevt", None)
    ctx.user_data.pop("att_auth", None)
    ctx.user_data.pop("last_keldi_ts", None)
    ctx.user_data.pop("last_ketdi_ts", None)

    user_id = update.effective_user.id
    saved = get_farmatsevt_by_userid(user_id)

    if saved:
        # ID saqlangan — ismini ko'rsatib kod so'rash
        ctx.user_data["att_saved"] = saved
        await update.message.reply_text(
            f"👤 *{saved['ismi']}*\n🏪 {saved['filial']}\n\n🔐 Filial kodini kiriting:",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
        )
    else:
        # Yangi foydalanuvchi — oddiy kod so'rash
        ctx.user_data.pop("att_saved", None)
        await update.message.reply_text(
            "🔐 *Davomat tizimi*\n\nFilial *kodingizni* kiriting:\n_(Filial raqami, masalan: 6)_",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
        )
    return ATT_PASSWORD


async def att_code_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    Filial kodini qabul qiladi.

    HOLAT 1 — TelegramID saqlangan (qaytuvchi foydalanuvchi):
      Kiritilgan kod uning filial raqamiga mos kelishini tekshiradi.
      Mos kelsa → kiradi. Mos kelmasa → xato.

    HOLAT 2 — Yangi foydalanuvchi (ID yo'q):
      Filial raqami bo'yicha farmatsevtlar ro'yxatini ko'rsatadi → tanlaydi.
      Tanlangandan so'ng TelegramID saqlanadi.
    """
    txt = update.message.text.strip()

    if txt == "⬅️ Orqaga":
        from bot import main_keyboard, get_lang, MENU
        language = get_lang(ctx)
        await update.message.reply_text(
            "📋 Asosiy menyu", reply_markup=main_keyboard(language)
        )
        return MENU

    saved = ctx.user_data.get("att_saved")  # ID bo'yicha topilgan farmatsevt

    if saved:
        # ── HOLAT 1: ID saqlangan → kod filialiga mosligini tekshir ──
        filial_kod = generate_filial_kod(saved["filial"])
        if txt.strip() == filial_kod:
            # Mos keldi → kirish
            ctx.user_data["att_farmatsevt"] = saved
            ctx.user_data["att_auth"] = True
            await update.message.reply_text(
                f"✅ *{saved['ismi']}*, xush kelibsiz!\n🏪 {saved['filial']}",
                parse_mode="Markdown",
            )
            return await _show_att_menu(update, ctx)
        else:
            # Mos kelmadi → xato
            await update.message.reply_text(
                "❌ Kod noto'g'ri!\n\n"
                f"👤 *{saved['ismi']}* uchun to'g'ri kodni kiriting.",
                parse_mode="Markdown",
            )
            return ATT_PASSWORD

    # ── HOLAT 2: Yangi foydalanuvchi → filial ro'yxati ──
    farmatsevtlar = get_farmatsevt_by_code(txt)

    if not farmatsevtlar:
        await update.message.reply_text(
            f"❌ *{txt}* — filial topilmadi.\n\nFilial raqamini kiriting (masalan: 6)",
            parse_mode="Markdown",
        )
        return ATT_PASSWORD

    if len(farmatsevtlar) == 1:
        # Bitta farmatsevt — to'g'ri kirish + ID saqlash
        f = farmatsevtlar[0]
        ctx.user_data["att_farmatsevt"] = f
        ctx.user_data["att_auth"] = True
        save_userid_by_code(update.effective_user.id, txt, f["telefon"])
        await update.message.reply_text(
            f"✅ Xush kelibsiz, *{f['ismi']}*!\n🏪 {f['filial']}",
            parse_mode="Markdown",
        )
        return await _show_att_menu(update, ctx)

    # Bir nechta farmatsevt → tanlash
    ctx.user_data["att_filial_candidates"] = farmatsevtlar
    ctx.user_data["att_entered_code"] = txt
    buttons = []
    for i, f in enumerate(farmatsevtlar):
        lavozim = f.get("lavozim", "")
        label = f"👤 {f['ismi']}"
        if lavozim:
            label += f"  ({lavozim})"
        buttons.append([InlineKeyboardButton(label, callback_data=f"att_who_{i}")])
    buttons.append([InlineKeyboardButton("⬅️ Orqaga", callback_data="att_who_back")])

    await update.message.reply_text(
        f"🏪 *{farmatsevtlar[0]['filial']}*\n\nO'zingizni tanlang:",
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode="Markdown",
    )
    return ATT_SELECT_WHO



async def att_select_who_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Bir filialdagi bir nechta farmatsevtdan birini tanlash."""
    q = update.callback_query
    await q.answer()

    if q.data == "att_who_back":
        await q.message.reply_text(
            "🔐 *Davomat tizimi*\n\nFilial *kodingizni* kiriting:\n_(masalan: 6)_",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
        )
        return ATT_PASSWORD

    idx = int(q.data.replace("att_who_", ""))
    candidates = ctx.user_data.get("att_filial_candidates", [])

    if idx >= len(candidates):
        await q.message.reply_text("❌ Xato. Qayta urinib ko'ring.")
        return ATT_PASSWORD

    farmatsevt = candidates[idx]
    ctx.user_data["att_farmatsevt"] = farmatsevt
    ctx.user_data["att_auth"] = True

    # Yangi foydalanuvchi — TelegramID saqlash
    entered_code = ctx.user_data.get("att_entered_code", "")
    save_userid_by_code(q.from_user.id, entered_code, farmatsevt["telefon"])

    await q.message.reply_text(
        f"✅ Xush kelibsiz, *{farmatsevt['ismi']}*!\n🏪 {farmatsevt['filial']}",
        parse_mode="Markdown",
    )
    return await _show_att_menu(q.message, ctx)


async def _show_att_menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    farmatsevt = ctx.user_data.get("att_farmatsevt", {})
    await update.message.reply_text(
        f"📋 *Davomat menyu*\n"
        f"👤 {farmatsevt.get('ismi', '')}\n"
        f"🏪 Filial: #{farmatsevt.get('filial', '')}",
        reply_markup=att_main_keyboard(),
        parse_mode="Markdown",
    )
    return ATT_MENU


# ─── 2. Davomat menyusi ───────────────────────────────────────────────────────

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


# ─── 3. Lokatsiya tekshiruvi ──────────────────────────────────────────────────

async def att_location_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    from datetime import datetime, timezone, timedelta
    UZ_TZ = timezone(timedelta(hours=5))

    if update.message.text == "⬅️ Orqaga":
        return await _show_att_menu(update, ctx)

    if not update.message.location:
        await update.message.reply_text("❌ Iltimos, lokatsiya tugmasini bosing.")
        return ATT_LOCATION

    # Jonli lokatsiya tekshiruvi
    loc = update.message.location
    if not loc.live_period:
        await update.message.reply_text(
            "❌ Faqat *jonli lokatsiya* qabul qilinadi!\n\n"
            "📍 Lokatsiya yuborish tugmasini bosing →\n"
            "*Jonli lokatsiya ulashish* tanlang.",
            parse_mode="Markdown",
            reply_markup=location_keyboard(),
        )
        return ATT_LOCATION

    ulat = loc.latitude
    ulon = loc.longitude

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
    now = datetime.now(UZ_TZ)
    now_str = now.strftime("%H:%M")
    now_ts = now.timestamp()

    # Vaqt cheklovi
    last_keldi = ctx.user_data.get("last_keldi_ts")
    last_ketdi = ctx.user_data.get("last_ketdi_ts")

    if action == "keldi":
        if last_ketdi and (now_ts - last_ketdi) < 7 * 3600:
            qolgan = int((7 * 3600 - (now_ts - last_ketdi)) / 60)
            await update.message.reply_text(
                f"⏳ Ketdidan keyin *7 soat* kutish kerak.\n"
                f"Qolgan: *{qolgan // 60} soat {qolgan % 60} daqiqa*",
                parse_mode="Markdown",
                reply_markup=att_main_keyboard(),
            )
            return ATT_MENU
        if last_keldi and not last_ketdi:
            await update.message.reply_text(
                f"❌ Avval *Ketdi* ni bosing!\n"
                f"Keldi vaqti: *{ctx.user_data.get('last_keldi_str', '')}*",
                parse_mode="Markdown",
                reply_markup=att_main_keyboard(),
            )
            return ATT_MENU

    elif action == "ketdi":
        if not last_keldi:
            await update.message.reply_text(
                "❌ Avval *Keldi* ni bosing!",
                parse_mode="Markdown",
                reply_markup=att_main_keyboard(),
            )
            return ATT_MENU

    ok = write_attendance(farmatsevt, action, zamena=False)

    if ok:
        if action == "keldi":
            ctx.user_data["last_keldi_ts"] = now_ts
            ctx.user_data["last_keldi_str"] = now_str
            ctx.user_data["last_ketdi_ts"] = None
        else:
            ctx.user_data["last_ketdi_ts"] = now_ts

        emoji = "✅" if action == "keldi" else "🚪"
        await update.message.reply_text(
            f"{emoji} *{farmatsevt['ismi']}* — {action.upper()}!\n"
            f"🕐 Vaqt: {now_str}\n"
            f"🏪 Filial: #{farmatsevt['filial']}\n"
            f"📏 Masofa: {dist:.0f} m",
            parse_mode="Markdown",
            reply_markup=att_main_keyboard(),
        )
    else:
        await update.message.reply_text(
            "⚠️ Xatolik. Qayta urinib ko'ring.",
            reply_markup=att_main_keyboard()
        )
    return ATT_MENU


# ─── 4. Zamena ───────────────────────────────────────────────────────────────

async def att_zamena_filial_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()

    if q.data == "att_back":
        await q.message.reply_text("📋 Davomat menyu", reply_markup=att_main_keyboard())
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

    if not update.message.location.live_period:
        await update.message.reply_text(
            "❌ Faqat *jonli lokatsiya* qabul qilinadi!\n\n"
            "📍 Lokatsiya → *Jonli lokatsiya ulashish* tanlang.",
            parse_mode="Markdown",
            reply_markup=location_keyboard("📍 Zamena lokatsiyamni yuborish"),
        )
        return ATT_ZAMENA_LOCATION

    ulat = update.message.location.latitude
    ulon = update.message.location.longitude

    zamena_filial = ctx.user_data.get("att_zamena_filial", {})
    dist = haversine_m(ulat, ulon, zamena_filial.get("lat", 0), zamena_filial.get("lon", 0))

    if dist > MAX_DISTANCE_KM * 1000:
        await update.message.reply_text(
            f"❌ Zamena filialidan *{dist:.0f} metr* uzoqdasiz.\n"
            f"Maksimal: *100 metr*.",
            parse_mode="Markdown",
            reply_markup=location_keyboard("📍 Zamena lokatsiyamni yuborish"),
        )
        return ATT_ZAMENA_LOCATION

    farmatsevt = ctx.user_data.get("att_farmatsevt", {})
    zamena_info = {**farmatsevt, "filial": zamena_filial["filial"]}
    ok = write_attendance(zamena_info, "keldi", zamena=True)

    from datetime import datetime, timezone, timedelta
    now_str = datetime.now(timezone(timedelta(hours=5))).strftime("%H:%M")

    if ok:
        await update.message.reply_text(
            f"🔄 *Zamena tasdiqlandi!*\n"
            f"👤 {farmatsevt['ismi']}\n"
            f"🏪 Zamena filial: #{zamena_filial['filial']}\n"
            f"🕐 Vaqt: {now_str}\n"
            f"📏 Masofa: {dist:.0f} m\n\n"
            f"_Jadvalda sariq rangda ko'rinadi_ 🟡",
            parse_mode="Markdown",
            reply_markup=att_main_keyboard(),
        )
    else:
        await update.message.reply_text("⚠️ Xatolik.", reply_markup=att_main_keyboard())

    return ATT_MENU


# ─── 5. Admin buyruqlari ──────────────────────────────────────────────────────

async def cmd_fill_codes(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    /fill_codes — Farmatsevtlar Sheets dagi barcha qatorlarga
    avtomatik kod yaratib 'Kod' (G) ustuniga yozadi.
    Sheets da G ustuni sarlavhasi: Kod
    """
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return

    msg = await update.message.reply_text("⏳ Kodlar yaratilmoqda...")
    codes = fill_codes_in_sheet()

    if not codes:
        await msg.edit_text("ℹ️ Barcha farmatsevtlarda kod allaqachon bor yoki ro'yxat bo'sh.")
        return

    lines = [f"✅ *{len(codes)} ta farmatsevtga kod yozildi:*\n"]
    for c in codes[:30]:   # Juda ko'p bo'lsa qisqartir
        lines.append(f"  • {c}")
    if len(codes) > 30:
        lines.append(f"  ... va yana {len(codes)-30} ta")

    await msg.edit_text("\n".join(lines), parse_mode="Markdown")


async def cmd_sync_pharmacists(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/sync_pharmacists — Farmatsevtlar ro'yxatini davomat jadvali bilan sinxronlashtiradi."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return

    msg = await update.message.reply_text("⏳ Sinxronizatsiya boshlanmoqda...")
    results = sync_pharmacists()

    if "error" in results:
        await msg.edit_text(f"❌ Xato: {results['error']}")
        return

    lines = ["✅ *Sinxronizatsiya tugadi!*\n"]
    if results["added"]:
        lines.append(f"🆕 *Yangi ({len(results['added'])} ta):*")
        for name in results["added"]:
            lines.append(f"  • {name}")
    if results["updated"]:
        lines.append(f"\n✏️ *O'zgardi ({len(results['updated'])} ta):*")
        for info in results["updated"]:
            lines.append(f"  • {info}")
    if results["removed"]:
        lines.append(f"\n🚫 *O'chirildi — kulrang ({len(results['removed'])} ta):*")
        for name in results["removed"]:
            lines.append(f"  • {name}")
    lines.append(f"\n⚪ O'zgarishsiz: {results['unchanged']} ta")

    await msg.edit_text("\n".join(lines), parse_mode="Markdown")


async def cmd_init_month(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/init_month — Oy boshida farmatsevtlarni Sheet ga yozadi."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return
    await update.message.reply_text("⏳ Oy listi tayyorlanmoqda...")
    try:
        from attendance import init_month_sheet
        init_month_sheet()
        await update.message.reply_text("✅ Farmatsevtlar ro'yxati Sheet ga yozildi!")
    except Exception as e:
        await update.message.reply_text(f"❌ Xato: {e}")


async def cmd_calc_hours(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/calc_hours — Oy oxirida ish soatlarini hisoblaydi."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return
    await update.message.reply_text("⏳ Ish soatlari hisoblanmoqda...")
    try:
        from attendance import calculate_monthly_hours
        count = calculate_monthly_hours()
        await update.message.reply_text(f"✅ {count} ta farmatsevt ish soati hisoblandi!")
    except Exception as e:
        await update.message.reply_text(f"❌ Xato: {e}")


# ─── Handler ro'yxati ─────────────────────────────────────────────────────────

def get_att_states():
    return {
        ATT_PASSWORD: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_code_handler),
        ],
        ATT_PHONE: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_code_handler),
        ],
        ATT_SELECT_WHO: [
            CallbackQueryHandler(att_select_who_handler, pattern="^att_who_"),
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
