import os
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
    init_month_sheet, calculate_monthly_hours,
    sync_pharmacists, fill_codes_in_sheet,
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
    from datetime import datetime, timezone, timedelta
    UZ_TZ = timezone(timedelta(hours=5))

    if update.message.text == "⬅️ Orqaga":
        return await _show_att_menu(update, ctx)

    if not update.message.location:
        await update.message.reply_text("❌ Iltimos, lokatsiya tugmasini bosing.")
        return ATT_LOCATION

    # 🔴 Jonli lokatsiya tekshiruvi
    loc = update.message.location
    if not loc.live_period:
        await update.message.reply_text(
            "❌ Faqat *jonli lokatsiya* qabul qilinadi!\n\n"
            "📍 Lokatsiya yuborish tugmasini bosing → "
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

    # ⏰ Vaqt cheklovi tekshiruvi
    last_keldi = ctx.user_data.get("last_keldi_ts")
    last_ketdi = ctx.user_data.get("last_ketdi_ts")

    if action == "keldi":
        # Ketdidan keyin 7 soat o'tganmi?
        if last_ketdi and (now_ts - last_ketdi) < 7 * 3600:
            qolgan_min = int((7 * 3600 - (now_ts - last_ketdi)) / 60)
            soat = qolgan_min // 60
            daqiqa = qolgan_min % 60
            await update.message.reply_text(
                f"⏳ Ketdidan keyin *7 soat* kutish kerak.\n"
                f"Qolgan vaqt: *{soat} soat {daqiqa} daqiqa*",
                parse_mode="Markdown",
                reply_markup=att_main_keyboard(),
            )
            return ATT_MENU
        # Ketdi bosilmay yana keldi bosyaptimi?
        if last_keldi and not last_ketdi:
            await update.message.reply_text(
                f"❌ Avval *Ketdi* ni bosing!\n"
                f"Keldi vaqti: *{ctx.user_data.get('last_keldi_str', '')}*",
                parse_mode="Markdown",
                reply_markup=att_main_keyboard(),
            )
            return ATT_MENU
        if last_keldi and last_ketdi and last_keldi > last_ketdi:
            await update.message.reply_text(
                f"❌ Avval *Ketdi* ni bosing!",
                parse_mode="Markdown",
                reply_markup=att_main_keyboard(),
            )
            return ATT_MENU

    elif action == "ketdi":
        # Keldi bosilmay ketdi bosyaptimi?
        if not last_keldi:
            await update.message.reply_text(
                "❌ Avval *Keldi* ni bosing!",
                parse_mode="Markdown",
                reply_markup=att_main_keyboard(),
            )
            return ATT_MENU

    ok = write_attendance(farmatsevt, action, zamena=False)

    if ok:
        # Vaqtni saqlash
        if action == "keldi":
            ctx.user_data["last_keldi_ts"] = now_ts
            ctx.user_data["last_keldi_str"] = now_str
        else:
            ctx.user_data["last_ketdi_ts"] = now_ts

        emoji = "✅" if action == "keldi" else "🚪"
        await update.message.reply_text(
            f"{emoji} *{farmatsevt['ismi']}* — {action}!\n"
            f"🕐 Vaqt: {now_str}\n"
            f"🏪 Filial: {farmatsevt['filial']}\n"
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

    # Jonli lokatsiya tekshiruvi
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



async def cmd_fix_latlon(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    /fix_latlon — Farmatsevtlar Sheets dagi Lat/Lon bo'sh qatorlarni
    Filiallar Sheets dan avtomatik to'ldiradi.
    """
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return

    msg = await update.message.reply_text("⏳ Lat/Lon to'ldirilmoqda...")

    try:
        import json, re
        from google.oauth2.service_account import Credentials
        import gspread

        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        PHARMACY_SHEET_ID  = os.getenv("PHARMACY_SHEET_ID", "")
        FILIALLAR_SHEET_ID = os.getenv("FILIALLAR_SHEET_ID", "")

        creds_json = os.getenv("GOOGLE_CREDENTIALS")
        if creds_json:
            info = json.loads(creds_json)
            creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        else:
            creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
        client = gspread.authorize(creds)

        # 1. Filiallar Sheets dan barcha koordinatalarni olish
        fil_ws = client.open_by_key(FILIALLAR_SHEET_ID).sheet1
        fil_values = fil_ws.get_all_values()

        # Sarlavhadan indekslar
        headers = [h.strip() for h in fil_values[0]] if fil_values else []
        try:
            fil_no_idx = headers.index("Filial №")
        except ValueError:
            fil_no_idx = 0
        try:
            lat_idx = headers.index("Latitude")
        except ValueError:
            lat_idx = 12
        try:
            lon_idx = headers.index("Longitude")
        except ValueError:
            lon_idx = 13

        # Filial raqami → {lat, lon} lug'at
        filial_coords = {}
        for row in fil_values[1:]:
            if not row or not row[fil_no_idx]:
                continue
            fil_no = str(row[fil_no_idx]).strip()
            if fil_no.lower() in ("асосий", "asosiy"):
                fil_no = "0"
            lat = str(row[lat_idx]).strip() if len(row) > lat_idx else ""
            lon = str(row[lon_idx]).strip() if len(row) > lon_idx else ""
            if lat and lon and lat not in ("0", "nan") and lon not in ("0", "nan"):
                filial_coords[fil_no] = {"lat": lat, "lon": lon}

        # 2. Farmatsevtlar Sheets ni olish
        ph_ws = client.open_by_key(PHARMACY_SHEET_ID).sheet1
        ph_values = ph_ws.get_all_values()

        updated = 0
        not_found = 0
        already_has = 0

        updates = []
        for i, row in enumerate(ph_values):
            if i == 0:
                continue  # sarlavha
            if not row or not row[0]:
                continue

            # Lat/Lon tekshirish (F=5, G=6, 0-indexed)
            lat_val = str(row[5]).strip() if len(row) > 5 else ""
            lon_val = str(row[6]).strip() if len(row) > 6 else ""

            if lat_val and lon_val and lat_val not in ("", "0", "nan"):
                already_has += 1
                continue  # allaqachon bor

            # Filial raqamini ajratish
            filial_cell = str(row[0]).strip()
            m = re.match(r"^(\d+)", filial_cell)
            fil_no = m.group(1) if m else ""

            if not fil_no or fil_no not in filial_coords:
                not_found += 1
                continue

            # Yangilash
            row_num = i + 1  # 1-indexed
            coords = filial_coords[fil_no]
            updates.append({
                "row": row_num,
                "lat": coords["lat"],
                "lon": coords["lon"],
                "ismi": str(row[1]).strip() if len(row) > 1 else f"qator {row_num}",
            })

        # Batch yangilash — bir so'rovda hammasi
        if updates:
            batch_data = []
            for upd in updates:
                # F ustun (Lat)
                from gspread.utils import rowcol_to_a1
                batch_data.append({
                    "range": f"F{upd['row']}:G{upd['row']}",
                    "values": [[upd["lat"], upd["lon"]]]
                })
                updated += 1
                print(f"[FIX] {upd['ismi']} → Lat={upd['lat']}, Lon={upd['lon']}")

            ph_ws.batch_update(batch_data, value_input_option="USER_ENTERED")

        lines = [
            f"✅ *Lat/Lon yangilandi!*\n",
            f"✅ To'ldirildi: *{updated}* ta",
            f"⚪ Allaqachon bor: *{already_has}* ta",
            f"❌ Filial topilmadi: *{not_found}* ta",
        ]
        await msg.edit_text("\n".join(lines), parse_mode="Markdown")

    except Exception as e:
        await msg.edit_text(f"❌ Xato: {e}")

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


# ─── Admin buyruqlari ─────────────────────────────────────────────────────────

# Admin Telegram ID larini shu yerga qo'shing
ADMIN_IDS = [709544046]  # Admin: Abdulaziz



async def cmd_sync_pharmacists(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/sync_pharmacists — Farmatsevtlar ro'yxatini davomat jadvali bilan sinxronlashtiradi."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return
    msg = await update.message.reply_text("⏳ Sinxronizatsiya boshlanmoqda...")
    try:
        results = sync_pharmacists()
        if "error" in results:
            await msg.edit_text(f"❌ Xato: {results['error']}")
            return
        lines = ["✅ *Sinxronizatsiya tugadi!*\n"]
        if results.get("added"):
            lines.append(f"🆕 *Yangi ({len(results['added'])} ta):*")
            for name in results["added"]:
                lines.append(f"  • {name}")
        if results.get("updated"):
            lines.append(f"\n✏️ *O'zgardi ({len(results['updated'])} ta):*")
            for info in results["updated"]:
                lines.append(f"  • {info}")
        if results.get("removed"):
            lines.append(f"\n🚫 *O'chirildi ({len(results['removed'])} ta):*")
            for name in results["removed"]:
                lines.append(f"  • {name}")
        lines.append(f"\n⚪ O'zgarishsiz: {results.get('unchanged', 0)} ta")
        await msg.edit_text("\n".join(lines), parse_mode="Markdown")
    except Exception as e:
        await msg.edit_text(f"❌ Xato: {e}")


async def cmd_fill_codes(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/fill_codes — Farmatsevtlar Sheets ga kod yozadi."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return
    msg = await update.message.reply_text("⏳ Kodlar yaratilmoqda...")
    try:
        codes = fill_codes_in_sheet()
        if not codes:
            await msg.edit_text("ℹ️ Barcha farmatsevtlarda kod allaqachon bor.")
            return
        lines = [f"✅ *{len(codes)} ta farmatsevtga kod yozildi:*\n"]
        for c in codes[:30]:
            lines.append(f"  • {c}")
        if len(codes) > 30:
            lines.append(f"  ... va yana {len(codes)-30} ta")
        await msg.edit_text("\n".join(lines), parse_mode="Markdown")
    except Exception as e:
        await msg.edit_text(f"❌ Xato: {e}")

async def cmd_init_month(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Oy boshida farmatsevtlarni Sheet ga yozadi. /init_month"""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return
    await update.message.reply_text("⏳ Oy listi tayyorlanmoqda...")
    try:
        init_month_sheet()
        await update.message.reply_text("✅ Farmatsevtlar ro'yxati Sheet ga yozildi!")
    except Exception as e:
        await update.message.reply_text(f"❌ Xato: {e}")


async def cmd_calc_hours(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Oy oxirida ish soatlarini hisoblaydi. /calc_hours"""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return
    await update.message.reply_text("⏳ Ish soatlari hisoblanmoqda...")
    try:
        count = calculate_monthly_hours()
        await update.message.reply_text(f"✅ {count} ta farmatsevt ish soati hisoblandi!")
    except Exception as e:
        await update.message.reply_text(f"❌ Xato: {e}")
