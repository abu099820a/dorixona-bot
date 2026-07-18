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
    run_read, run_write,
)

ATT_PASSWORD = 106   # Parol kutish state
ATT_CHANGE_FILIAL = 108
ATT_CHANGE_LAVOZIM = 109
ATT_DAYMARK_START = 110   # Dam/Javob olish: boshlanish sanasi kutiladi
ATT_DAYMARK_END = 111     # Dam/Javob olish: tugash sanasi kutiladi
ATT_PAROL = "офис"  # Universal parol

# ─── Klaviaturalar ────────────────────────────────────────────────────────────

def att_main_keyboard():
    return ReplyKeyboardMarkup([
        ["✅ Keldi", "🚪 Ketdi"],
        ["🔄 Zamena"],
        ["🛌 Dam olish kuni", "📄 Javob olish kuni"],
        ["🏥 Filial/Lavozim o'zgartirish"],
        ["⬅️ Orqaga"],
    ], resize_keyboard=True)


def filial_inline_keyboard(filiallar: list):
    buttons = []
    row = []
    for f in filiallar:
        name = f.get("filial_name", f["filial"])
        row.append(InlineKeyboardButton(f"#{f['filial']} {name}", callback_data=f"att_fil_{f['filial']}"))
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
    """
    Asosiy menyudan 'Davomot' bosilganda:
    - TelegramID saqlangan bo'lsa → to'g'ri menyuga
    - Saqlangan bo'lmasa → Ro'yxatdan o'tishga yo'naltiradi
    """
    user_id = update.effective_user.id

    # TelegramID bo'yicha tekshirish
    if not ctx.user_data.get("att_farmatsevt"):
        farmatsevt = await run_read(get_farmatsevt_by_userid, user_id)
        if farmatsevt:
            ctx.user_data["att_auth"] = True
            ctx.user_data["att_farmatsevt"] = farmatsevt

    if ctx.user_data.get("att_auth") and ctx.user_data.get("att_farmatsevt"):
        return await _show_att_menu(update, ctx)

    # Ro'yxatdan o'tmagan — yo'naltirish
    from bot import main_keyboard, get_lang, MENU
    await update.message.reply_text(
        "❌ Siz hali ro'yxatdan o'tmagansiz!\n\n"
        "📝 Iltimos, avval *Ro'yxatdan o'tish* tugmasini bosing.",
        parse_mode="Markdown",
        reply_markup=main_keyboard(get_lang(ctx)),
    )
    return MENU


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
    farmatsevt = await run_read(get_farmatsevt, phone)

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
    await run_write(save_userid_to_sheet, user_id, phone)

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

    elif txt in ["🛌 Dam olish kuni", "📄 Javob olish kuni"]:
        kind = "dam" if txt == "🛌 Dam olish kuni" else "javob"
        label = "Dam olish" if kind == "dam" else "Javob olish"
        ctx.user_data["daymark_kind"] = kind
        await update.message.reply_text(
            f"📅 *{label} kunlari*\n\n"
            f"Boshlanish sanasini kiriting (kun raqami, masalan: 20):",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
        )
        return ATT_DAYMARK_START

    elif txt == "🏥 Filial/Lavozim o'zgartirish":
        farmatsevt = ctx.user_data.get("att_farmatsevt", {})
        ctx.user_data["change_farmatsevt"] = farmatsevt
        await update.message.reply_text(
            f"🏥 *Filial/Lavozim o'zgartirish*\n\n"
            f"👤 {farmatsevt.get('ismi', '')}\n"
            f"🏪 Hozirgi filial: {farmatsevt.get('filial', '')}\n\n"
            f"Yangi filial *raqamini* kiriting:\n_(masalan: 6)_",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
        )
        return ATT_CHANGE_FILIAL

    elif txt == "🔄 Zamena":
        ctx.user_data["att_zamena"] = True
        await update.message.reply_text(
            "🔄 *Zamena rejimi*\n\n"
            "Boradigan filial *raqamini* yozing:\n"
            "_(masalan: 6)_",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
        )
        return ATT_ZAMENA_FILIAL

    return ATT_MENU

# ─── 4. Lokatsiya tekshiruvi ──────────────────────────────────────────────────

async def att_location_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    from datetime import datetime, timezone, timedelta
    UZ_TZ = timezone(timedelta(hours=5))

    # update.message None bo'lishi mumkin (inline update)
    if not update.message:
        return ATT_LOCATION

    if update.message.text and update.message.text == "⬅️ Orqaga":
        return await _show_att_menu(update, ctx)

    if not update.message.location:
        await update.message.reply_text("❌ Iltimos, lokatsiya tugmasini bosing.")
        return ATT_LOCATION

    loc = update.message.location

    ulat = loc.latitude
    ulon = loc.longitude

    farmatsevt = ctx.user_data.get("att_farmatsevt", {})

    fil_lat = farmatsevt.get("lat", 0)
    fil_lon = farmatsevt.get("lon", 0)
    if not fil_lat or not fil_lon:
        # Filialning koordinatasi jadvalda kiritilmagan yoki 0/bo'sh —
        # bunday holatda (0,0) nuqtasidan hisoblash mantiqsiz masofa
        # (~8000+ km) beradi. Aniq xabar bilan to'xtatamiz.
        await update.message.reply_text(
            f"⚠️ *{farmatsevt.get('filial', '')}* filiali uchun "
            f"koordinata (Lat/Lon) sozlanmagan.\n\n"
            f"Iltimos, administratorga murojaat qiling — "
            f"Farmatsevtlar jadvalida Lat/Lon ustunlarini to'ldirish kerak.",
            parse_mode="Markdown",
            reply_markup=location_keyboard(),
        )
        return ATT_LOCATION

    dist = haversine_m(ulat, ulon, fil_lat, fil_lon)

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

    # Tungi smena: 00:00-05:00 da ketdi bosilsa, bu kechagi ish kuniga tegishli
    check_date = now
    if action == "ketdi" and now.hour < 5:
        check_date = now - timedelta(days=1)

    # ⏰ JADVALDAN tekshirish — ISHONCHLI MANBA.
    # MUHIM: avvalgi versiyada bu tekshiruv faqat ctx.user_data (bot
    # xotirasi) ga tayanar edi. Bot qayta ishga tushganda (Railway
    # redeploy, xato bo'lib qayta ko'tarilishi va h.k.) xotiradagi bu
    # ma'lumot yo'qolib ketadi — natijada xodim "Keldi"ni allaqachon
    # bosgan (jadvalda yozilgan) bo'lsa ham, bot "Avval Keldi ni bosing!"
    # deb NOTO'G'RI xabar berardi (aynan shu — Nigora Eshmirzayevada
    # ko'rilgan muammo). Endi tekshiruv har doim to'g'ridan-to'g'ri
    # Google Sheets'dagi haqiqiy Keldi/Ketdi qiymatidan olinadi.
    from attendance import get_today_status
    sheet_today = await run_read(
        get_today_status, farmatsevt.get("ismi", ""), farmatsevt.get("filial", ""), check_date
    )
    sheet_keldi_time = sheet_today.get("keldi")   # "HH:MM" | None
    sheet_ketdi_time = sheet_today.get("ketdi")   # "HH:MM" | None

    last_keldi_ts = ctx.user_data.get("last_keldi_ts")
    last_ketdi_ts = ctx.user_data.get("last_ketdi_ts")

    keldi_done = bool(sheet_keldi_time) or bool(last_keldi_ts)
    ketdi_done = bool(sheet_ketdi_time) or bool(last_ketdi_ts)
    is_late_checkout = False

    if action == "keldi":
        # Ketdidan keyin 7 soat o'tganmi? Iloji boricha jadvaldagi haqiqiy
        # ketdi vaqtidan hisoblaymiz (xotiradagi vaqt bo'lmasa ham ishlaydi).
        ketdi_ts_for_cooldown = last_ketdi_ts
        if not ketdi_ts_for_cooldown and sheet_ketdi_time:
            try:
                h, m = map(int, sheet_ketdi_time.split(":"))
                ketdi_dt = check_date.replace(hour=h, minute=m, second=0, microsecond=0)
                ketdi_ts_for_cooldown = ketdi_dt.timestamp()
            except Exception:
                ketdi_ts_for_cooldown = None

        if ketdi_ts_for_cooldown and (now_ts - ketdi_ts_for_cooldown) < 7 * 3600:
            qolgan_min = int((7 * 3600 - (now_ts - ketdi_ts_for_cooldown)) / 60)
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
        if keldi_done and not ketdi_done:
            await update.message.reply_text(
                f"❌ Avval *Ketdi* ni bosing!\n"
                f"Keldi vaqti: *{sheet_keldi_time or ctx.user_data.get('last_keldi_str', '')}*",
                parse_mode="Markdown",
                reply_markup=att_main_keyboard(),
            )
            return ATT_MENU

    elif action == "ketdi":
        # Keldi bosilmay ketdi bosyaptimi?
        if not keldi_done:
            # Ehtimol xodim KECHA Keldi bosgan-u, Ketdi bosishni kechiktirib
            # yuborgan (masalan ertalab soat 5 dan keyin bosayotgan bo'lsa,
            # tungi smena qoidasi ishlamaydi). Shu sababli KECHAGI kunni ham
            # tekshiramiz — agar u yerda "ochiq" (Keldi bor, Ketdi yo'q)
            # qator topilsa, Ketdi o'sha kunga yoziladi.
            fallback_date = check_date - timedelta(days=1)
            fallback_status = await run_read(
                get_today_status, farmatsevt.get("ismi", ""), farmatsevt.get("filial", ""), fallback_date
            )
            if fallback_status.get("keldi") and not fallback_status.get("ketdi"):
                check_date = fallback_date
                sheet_keldi_time = fallback_status.get("keldi")
                sheet_ketdi_time = fallback_status.get("ketdi")
                keldi_done = True
                ketdi_done = False
                is_late_checkout = True
            else:
                await update.message.reply_text(
                    "❌ Avval *Keldi* ni bosing!",
                    parse_mode="Markdown",
                    reply_markup=att_main_keyboard(),
                )
                return ATT_MENU

    # write_time sifatida check_date ishlatiladi — bu tungi smena qoidasini
    # ("00:00-05:00 da ketdi → kechagi kun") va yuqoridagi kechikkan-ketdi
    # holatini (ochiq qolgan kechagi Keldi topilsa) ikkalasini ham to'g'ri
    # hisobga oladi.
    write_now = check_date

    ok = await run_write(write_attendance, farmatsevt, action, False, write_now)

    if ok:
        # Vaqtni saqlash
        if action == "keldi":
            ctx.user_data["last_keldi_ts"] = now_ts
            ctx.user_data["last_keldi_str"] = now_str
        else:
            ctx.user_data["last_ketdi_ts"] = now_ts

        emoji = "✅" if action == "keldi" else "🚪"
        late_note = ""
        if is_late_checkout:
            late_note = (
                f"\n⚠️ Bu *{write_now.strftime('%d.%m')}* kunidagi ochiq qolgan "
                f"Keldi uchun Ketdi sifatida belgilandi (kechikkan ketdi)."
            )
        await update.message.reply_text(
            f"{emoji} *{farmatsevt['ismi']}* — {action}!\n"
            f"🕐 Vaqt: {now_str}\n"
            f"🏪 Filial: {farmatsevt['filial']}\n"
            f"📏 Masofa: {dist:.0f} m"
            f"{late_note}",
            parse_mode="Markdown",
            reply_markup=att_main_keyboard(),
        )
    else:
        await update.message.reply_text("⚠️ Xatolik. Qayta urinib ko'ring.", reply_markup=att_main_keyboard())
    return ATT_MENU

# ─── 5. Zamena ───────────────────────────────────────────────────────────────

async def att_zamena_filial_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Zamena uchun filial raqamini matn orqali qabul qiladi."""
    txt = update.message.text.strip() if update.message and update.message.text else ""

    if txt == "⬅️ Orqaga":
        return await _show_att_menu(update, ctx)

    filiallar = await run_read(get_filiallar_list)
    selected = None
    for f in filiallar:
        fil_no = str(f["filial"]).strip()
        import re as _re
        m = _re.match(r"^(\d+)", fil_no)
        if m and m.group(1) == txt.strip():
            selected = f
            break
        if fil_no == txt.strip():
            selected = f
            break

    if not selected:
        await update.message.reply_text(
            f"❌ *{txt}* raqamli filial topilmadi.\n\nFilial raqamini qayta kiriting:",
            parse_mode="Markdown",
        )
        return ATT_ZAMENA_FILIAL

    ctx.user_data["att_zamena_filial"] = selected
    ctx.user_data["att_action"] = "keldi"

    filial_name = selected.get("filial_name", txt)
    await update.message.reply_text(
        f"🔄 *Zamena* — #{txt} {filial_name}\n\n📍 Lokatsiyangizni yuboring:",
        reply_markup=location_keyboard("📍 Zamena lokatsiyamni yuborish"),
        parse_mode="Markdown",
    )
    return ATT_ZAMENA_LOCATION


async def att_zamena_location_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.message.text and update.message.text == "⬅️ Orqaga":
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

    zf_lat = zamena_filial.get("lat", 0)
    zf_lon = zamena_filial.get("lon", 0)
    if not zf_lat or not zf_lon:
        await update.message.reply_text(
            f"⚠️ *{zamena_filial.get('filial_name', '')}* filiali uchun "
            f"koordinata (Lat/Lon) sozlanmagan.\n\n"
            f"Iltimos, administratorga murojaat qiling.",
            parse_mode="Markdown",
            reply_markup=location_keyboard("📍 Zamena lokatsiyamni yuborish"),
        )
        return ATT_ZAMENA_LOCATION

    dist = haversine_m(ulat, ulon, zf_lat, zf_lon)

    if dist > MAX_DISTANCE_KM * 1000:
        await update.message.reply_text(
            f"❌ Zamena filialidan *{dist:.0f} metr* uzoqdasiz.\nMaksimal: *100 metr*.",
            parse_mode="Markdown",
            reply_markup=location_keyboard("📍 Zamena lokatsiyamni yuborish"),
        )
        return ATT_ZAMENA_LOCATION

    farmatsevt = ctx.user_data.get("att_farmatsevt", {})
    zamena_info = {**farmatsevt, "filial": farmatsevt["filial"], "zamena_filial": zamena_filial["filial"]}
    ok = await run_write(write_attendance, zamena_info, "keldi", True)
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




async def att_daymark_start_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Dam olish / Javob olish: boshlanish sanasini qabul qiladi."""
    txt = update.message.text.strip() if update.message and update.message.text else ""

    if txt == "⬅️ Orqaga":
        return await _show_att_menu(update, ctx)

    if not txt.isdigit():
        await update.message.reply_text(
            "❌ Iltimos, kun raqamini kiriting (masalan: 20):",
        )
        return ATT_DAYMARK_START

    ctx.user_data["daymark_start"] = int(txt)
    await update.message.reply_text(
        "📅 Tugash sanasini kiriting (kun raqami, masalan: 23):",
    )
    return ATT_DAYMARK_END


async def att_daymark_end_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Dam olish / Javob olish: tugash sanasini qabul qiladi va jadvalga yozadi."""
    txt = update.message.text.strip() if update.message and update.message.text else ""

    if txt == "⬅️ Orqaga":
        return await _show_att_menu(update, ctx)

    if not txt.isdigit():
        await update.message.reply_text(
            "❌ Iltimos, kun raqamini kiriting (masalan: 23):",
        )
        return ATT_DAYMARK_END

    start_day = ctx.user_data.get("daymark_start")
    end_day = int(txt)
    kind = ctx.user_data.get("daymark_kind", "dam")
    label = "Dam olish" if kind == "dam" else "Javob olish"
    farmatsevt = ctx.user_data.get("att_farmatsevt", {})

    from attendance import mark_rest_days
    res = await run_write(
        mark_rest_days,
        farmatsevt.get("ismi", ""),
        farmatsevt.get("filial", ""),
        start_day,
        end_day,
        kind,
    )

    if res.get("error"):
        await update.message.reply_text(
            f"❌ {res['error']}",
            reply_markup=att_main_keyboard(),
        )
    elif res.get("ok"):
        lines = [f"✅ *{label} belgilandi!*\n"]
        if res["marked"]:
            lines.append(f"📌 Belgilangan kunlar: {', '.join(map(str, res['marked']))}")
        if res["skipped"]:
            lines.append(
                f"⚠️ O'tkazib yuborilgan kunlar (allaqachon to'ldirilgan): "
                f"{', '.join(map(str, res['skipped']))}"
            )
        await update.message.reply_text(
            "\n".join(lines),
            parse_mode="Markdown",
            reply_markup=att_main_keyboard(),
        )
    else:
        await update.message.reply_text(
            "❌ Xatolik yuz berdi. Administratorga murojaat qiling.",
            reply_markup=att_main_keyboard(),
        )

    ctx.user_data.pop("daymark_start", None)
    ctx.user_data.pop("daymark_kind", None)
    return ATT_MENU


def change_lavozim_keyboard():
    return ReplyKeyboardMarkup([
        ["👔 Farmatsevt"],
        ["👔 Dorixona mudiri"],
        ["👔 Stajyor"],
        ["⬅️ Orqaga"],
    ], resize_keyboard=True)

async def att_change_filial_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Yangi filial raqamini qabul qiladi."""
    txt = update.message.text.strip() if update.message and update.message.text else ""

    if txt == "⬅️ Orqaga":
        return await _show_att_menu(update, ctx)

    # Filial mavjudligini tekshirish
    # MUHIM: get_filiallar_list() o'rniga get_filial_info() ishlatiladi, chunki u
    # Farmatsevtlar jadvalida allaqachon mavjud bo'lgan "RAQAM - NOM" formatidagi
    # to'liq filial nomini qaytaradi (masalan "1 - ТАШМИ-1"). get_filiallar_list()
    # esa faqat raqamsiz nom qaytargani uchun keyingi bosqichda filial guruhini
    # aniqlash (raqam solishtirish) ishlamay qolib, xodim doim jadval oxiriga
    # tushib ketishi yoki eski joyida qolib ketishiga sabab bo'lgan.
    from register_handlers import get_filial_info
    selected = await run_read(get_filial_info, txt.strip())

    if not selected:
        await update.message.reply_text(
            f"❌ *{txt}* raqamli filial topilmadi.\nQayta kiriting:",
            parse_mode="Markdown",
        )
        return ATT_CHANGE_FILIAL

    ctx.user_data["change_new_filial"] = selected
    await update.message.reply_text(
        f"🏪 *{selected.get('filial_nomi', txt)}*\n\n👔 Yangi lavozimingizni tanlang:",
        parse_mode="Markdown",
        reply_markup=change_lavozim_keyboard(),
    )
    return ATT_CHANGE_LAVOZIM


async def att_change_lavozim_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Yangi lavozimni qabul qiladi va ma'lumotni yangilaydi."""
    txt = update.message.text.strip() if update.message and update.message.text else ""

    if txt == "⬅️ Orqaga":
        farmatsevt = ctx.user_data.get("change_farmatsevt", {})
        await update.message.reply_text(
            f"Yangi filial raqamini kiriting:",
            reply_markup=ReplyKeyboardMarkup([["⬅️ Orqaga"]], resize_keyboard=True),
        )
        return ATT_CHANGE_FILIAL

    lavozim_map = {
        "👔 Farmatsevt": "Farmatsevt",
        "👔 Dorixona mudiri": "Dorixona mudiri",
        "👔 Stajyor": "Stajyor",
    }

    if txt not in lavozim_map:
        await update.message.reply_text(
            "❌ Tugmalardan birini tanlang:",
            reply_markup=change_lavozim_keyboard(),
        )
        return ATT_CHANGE_LAVOZIM

    lavozim = lavozim_map[txt]
    farmatsevt = ctx.user_data.get("change_farmatsevt", {})
    new_filial = ctx.user_data.get("change_new_filial", {})
    user_id = update.effective_user.id

    def _safe_float(v):
        try:
            v = str(v).strip().replace(",", ".")
            return float(v) if v else 0.0
        except Exception:
            return 0.0

    new_lat = _safe_float(new_filial.get("lat", 0))
    new_lon = _safe_float(new_filial.get("lon", 0))

    # Sheets da yangilash
    from attendance import update_farmatsevt_filial_lavozim
    ok = await run_write(
        update_farmatsevt_filial_lavozim,
        ismi=farmatsevt.get("ismi", ""),
        old_filial=farmatsevt.get("filial", ""),
        new_filial=new_filial.get("filial_nomi", ""),
        new_lavozim=lavozim,
        lat=new_lat,
        lon=new_lon,
        telegram_id=user_id,
    )

    if ok:
        # Sessiyani yangilash
        ctx.user_data["att_farmatsevt"] = {
            **farmatsevt,
            "filial": new_filial.get("filial_nomi", ""),
            "lat": new_lat,
            "lon": new_lon,
        }
        if not new_lat or not new_lon:
            await update.message.reply_text(
                f"⚠️ Diqqat: *{new_filial.get('filial_nomi', '')}* filiali uchun "
                f"koordinata (Lat/Lon) topilmadi — administratorga xabar bering, "
                f"aks holda Keldi/Ketdi bosilganda xatolik chiqishi mumkin.",
                parse_mode="Markdown",
            )
        await update.message.reply_text(
            f"✅ *Ma'lumotlar yangilandi!*\n\n"
            f"👤 {farmatsevt.get('ismi', '')}\n"
            f"🏪 Yangi filial: {new_filial.get('filial_nomi', '')}\n"
            f"👔 Lavozim: {lavozim}",
            parse_mode="Markdown",
            reply_markup=att_main_keyboard(),
        )
    else:
        await update.message.reply_text(
            "❌ Xatolik. Admin bilan bog'laning.",
            reply_markup=att_main_keyboard(),
        )
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
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_zamena_filial_handler),
        ],
        ATT_CHANGE_FILIAL: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_change_filial_handler),
        ],
        ATT_CHANGE_LAVOZIM: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_change_lavozim_handler),
        ],
        ATT_ZAMENA_LOCATION: [
            MessageHandler(filters.LOCATION, att_zamena_location_handler),
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_zamena_location_handler),
        ],
        ATT_DAYMARK_START: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_daymark_start_handler),
        ],
        ATT_DAYMARK_END: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, att_daymark_end_handler),
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
        results = await run_write(sync_pharmacists)
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
        codes = await run_write(fill_codes_in_sheet)
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



async def cmd_fill_phones(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    /fill_phones — Davomat jadvalidagi barcha farmatsevtlarning
    telefon raqamlarini C ustuniga Farmatsevtlar Sheets dan olib yozadi.
    """
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return

    msg = await update.message.reply_text("⏳ Telefon raqamlari to'ldirilmoqda...")

    try:
        import json, re
        from google.oauth2.service_account import Credentials
        import gspread

        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        PHARMACY_SHEET_ID  = os.getenv("PHARMACY_SHEET_ID", "")
        ATTENDANCE_SHEET_ID_local = os.getenv("ATTENDANCE_SHEET_ID", "")

        creds_json = os.getenv("GOOGLE_CREDENTIALS")
        if creds_json:
            info = json.loads(creds_json)
            creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        else:
            creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
        client = gspread.authorize(creds)

        # 1. Farmatsevtlar Sheets dan ismi → telefon lug'at
        ph_ws = client.open_by_key(PHARMACY_SHEET_ID).sheet1
        ph_records = ph_ws.get_all_records()

        phone_dict = {}
        for row in ph_records:
            ismi = str(row.get("Ismi", "")).strip()
            tel = row.get("Telefon", "")
            if isinstance(tel, float):
                tel = str(int(tel))
            else:
                tel = str(tel).strip()
            if ismi and tel:
                phone_dict[ismi] = tel

        # 2. Davomat jadvalini olish
        from datetime import datetime, timezone, timedelta
        UZ_TZ = timezone(timedelta(hours=5))
        now = datetime.now(UZ_TZ)
        OY_NOMLARI = {
            1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
            5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
            9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
        }
        sheet_name = f"{OY_NOMLARI[now.month]} {now.year}"
        att_sh = client.open_by_key(ATTENDANCE_SHEET_ID_local)

        try:
            ws = att_sh.worksheet(sheet_name)
        except Exception:
            await msg.edit_text(f"❌ '{sheet_name}' listi topilmadi.")
            return

        all_values = ws.get_all_values()

        # 3. C ustuniga telefon yozish (batch)
        updates = []
        filled = 0
        skipped = 0

        for i, row in enumerate(all_values):
            if i < 2:
                continue
            if not row:
                continue
            # B ustun = Ismi
            ismi = str(row[1]).strip() if len(row) > 1 else ""
            if not ismi:
                continue  # filial sarlavha qatori

            # C ustun = Telefon
            existing_tel = str(row[2]).strip() if len(row) > 2 else ""
            if existing_tel and existing_tel not in ("", "0"):
                skipped += 1
                continue  # allaqachon bor

            if ismi in phone_dict:
                row_num = i + 1
                updates.append({
                    "range": f"C{row_num}",
                    "values": [[phone_dict[ismi]]]
                })
                filled += 1

        if updates:
            ws.batch_update(updates)

        lines = [
            f"✅ *Telefon raqamlari to'ldirildi!*\n",
            f"📱 To'ldirildi: *{filled}* ta",
            f"⚪ Allaqachon bor: *{skipped}* ta",
        ]
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
        await run_write(init_month_sheet)
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
        count = await run_write(calculate_monthly_hours)
        await update.message.reply_text(f"✅ {count} ta farmatsevt ish soati hisoblandi!")
    except Exception as e:
        await update.message.reply_text(f"❌ Xato: {e}")
