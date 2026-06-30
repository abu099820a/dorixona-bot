import math, io, re
import pandas as pd
from attendance_handlers import (
    att_enter, get_att_states,
    ATT_PHONE, ATT_MENU, ATT_LOCATION,
    ATT_ZAMENA_FILIAL, ATT_ZAMENA_LOCATION,
    cmd_init_month, cmd_calc_hours, cmd_sync_pharmacists, cmd_fill_codes,
)
from thefuzz import process as fuzz_process
from telegram import (
    Update, ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters, ConversationHandler
)
import os
from dotenv import load_dotenv
load_dotenv()

TOKEN = os.getenv("TOKEN", "8837024109:AAGFZP5akA2nPo0RugVCCbEl2wgoe9N5_Uo")
EXCEL_FILE = "dorixonalar.xlsx"
FILIALLAR_FILE = "filiallar.xlsx"
SHEETS_ID = os.getenv("SHEETS_ID", "1CfuogH-yY--y5kiBK0qXsl5AFi_Hmzj_onWcA-Qyvco")
MY_MAPS_URL = "https://www.google.com/maps/d/viewer?mid=1VxLvn1YnqjWs-PJTnwBR3dFaMVn3JyA"

TELEGRAM_CHAT_LINK = "https://t.me/+gDbA_KTD5fdjOGE6"
TELEGRAM_CHANNEL_LINK = "https://t.me/Vaksina_med_axborot"

OFFICE_LAT = 41.219104
OFFICE_LON = 69.272889

LANG, MENU, SEARCH_MENU, SEARCH_INPUT, LOCATION_WAIT, \
SELECT_RESULT, SELECT_REGION, SELECT_DISTRICT, LIST_PAGE = range(9)

PAGE_SIZE = 10

VILOYATLAR = {
    "Toshkent shahri": [
        "Алмазарский район","Мирабадский район","Мирзо Улугбекский район",
        "Сергелийский район","Учтепинский район","Чиланзарский район",
        "Шайхантахурский район","шайхантахурский район","Юнусабадский район",
        "Яккасарайский район","Янгихаётский район","Яшнободский район"
    ],
    "Toshkent viloyati": ["Зангиота","Кибрай","Куйичирчик","Паркент","Тошкент тумани"],
    "Samarqand viloyati": ["Самарканд"],
    "Surxondaryo viloyati": ["Сурхондарё"],
    "Qashqadaryo viloyati": ["Кашкадаре"],
    "Sirdaryo viloyati": ["Сирдарё"],
    "Farg'ona viloyati": ["Фаргона"],
    "Namangan viloyati": ["Наманган"],
    "Andijon viloyati": ["Андижон"],
    "Qoraqalpog'iston": ["Каракалпакстан"],
}

KIR_TO_LAT = {
    "а":"a","б":"b","в":"v","г":"g","д":"d","е":"e","ё":"yo","ж":"j","з":"z",
    "и":"i","й":"y","к":"k","л":"l","м":"m","н":"n","о":"o","п":"p","р":"r",
    "с":"s","т":"t","у":"u","ф":"f","х":"x","ц":"ts","ч":"ch","ш":"sh",
    "э":"e","ю":"yu","я":"ya","ў":"o","қ":"q","ғ":"g","ҳ":"h",
}
LAT_TO_KIR_MULTI = [("ch","ч"),("sh","ш"),("yo","ё"),("yu","ю"),("ya","я"),("ts","ц")]
LAT_TO_KIR_SINGLE = {
    "a":"а","b":"б","v":"в","g":"г","d":"д","e":"е","f":"ф","j":"ж",
    "z":"з","i":"и","y":"й","k":"к","l":"л","m":"м","n":"н","o":"о",
    "p":"п","r":"р","s":"с","t":"т","u":"у","x":"х","q":"к","h":"х",
}

def to_kirill(text):
    t = text.lower()
    for lat, kir in LAT_TO_KIR_MULTI:
        t = t.replace(lat, kir)
    for lat, kir in LAT_TO_KIR_SINGLE.items():
        t = t.replace(lat, kir)
    return t

def to_latin(text):
    t = text.lower()
    for kir, lat in KIR_TO_LAT.items():
        t = t.replace(kir, lat)
    return t

def search_variants(text):
    return list(set([text, text.lower(), to_kirill(text), to_latin(text)]))

T = {
    "uz": {
        "welcome": "👋 *Vaksin Med Dorixonalari* botiga xush kelibsiz!\n\n🏥 100+ filial, butun O'zbekiston bo'ylab\n\nTilni tanlang:",
        "menu": "📋 *Asosiy menyu*",
        "search_btn": "🔍 Filiallarni qidirish",
        "chat_btn": "💬 Telegram chat",
        "channel_btn": "📢 Kanal",
        "lang_btn": "🌐 Til o'zgartirish",
        "search_menu": "🔍 *Qidirish turi:*",
        "by_number": "🔢 Filial raqami",
        "by_name": "🔤 Nomi bo'yicha",
        "by_region": "🗺 Hudud bo'yicha",
        "nearest": "📍 Eng yaqin filial",
        "office_loc": "🏢 Ofis/sklad lokatsiyasi",
        "excel_btn": "📊 Excel olish",
        "map_btn": "🗺 Barcha filiallar kartada",
        "attendance_btn": "📋 Davomat",
        "back": "⬅️ Orqaga",
        "enter_number": "🔢 Filial raqamini kiriting:\n_(masalan: 1, 5, 23)_",
        "enter_name": "🔤 Dorixona nomini kiriting:",
        "select_viloyat": "🗺 Viloyatni tanlang:",
        "select_tuman": "📍 Tumanni tanlang:",
        "send_loc": "📍 Joylashuvingizni yuboring:",
        "loc_btn": "📍 Joylashuvimni yuborish",
        "not_found": "❌ Hech narsa topilmadi. Qaytadan urinib ko'ring.",
        "similar": "🔍 Siz quyidagilarni nazarda tutdingizmi?",
        "found_many": "✅ *{n} ta topildi.* Birini tanlang:",
        "excel_cap": "📊 Vaksin Med — barcha filiallar ro'yxati",
        "km": "km uzoqlikda",
        "all_districts": "📋 Barcha filiallar",
        "list_title": "📋 *Filiallar ro'yxati* ({start}-{end} / {total}):",
        "prev": "⬅️ Oldingi",
        "next": "Keyingi ➡️",
        "map_choose": "Xaritani tanlang:",
        "yandex": "🗺 Yandex Maps",
        "google": "🗺 Google Maps",
        "office_title": "🏢 *Ofis/Sklad manzili*",
        "office_address": "📍 Toshkent shahri, Yunusobod tumani",
    },
    "ru": {
        "welcome": "👋 Добро пожаловать в бот *Vaksin Med*!\n\n🏥 100+ филиалов по всему Узбекистану\n\nВыберите язык:",
        "menu": "📋 *Главное меню*",
        "search_btn": "🔍 Найти филиал",
        "chat_btn": "💬 Telegram чат",
        "channel_btn": "📢 Канал",
        "lang_btn": "🌐 Сменить язык",
        "search_menu": "🔍 *Тип поиска:*",
        "by_number": "🔢 По номеру филиала",
        "by_name": "🔤 По названию",
        "by_region": "🗺 По региону",
        "nearest": "📍 Ближайший филиал",
        "office_loc": "🏢 Офис/склад локация",
        "excel_btn": "📊 Скачать Excel",
        "map_btn": "🗺 Все филиалы на карте",
        "attendance_btn": "📋 Davomat",
        "back": "⬅️ Назад",
        "enter_number": "🔢 Введите номер филиала:\n_(например: 1, 5, 23)_",
        "enter_name": "🔤 Введите название аптеки:",
        "select_viloyat": "🗺 Выберите регион:",
        "select_tuman": "📍 Выберите район:",
        "send_loc": "📍 Отправьте ваше местоположение:",
        "loc_btn": "📍 Отправить местоположение",
        "not_found": "❌ Ничего не найдено. Попробуйте снова.",
        "similar": "🔍 Возможно вы имели в виду:",
        "found_many": "✅ *Найдено {n}.* Выберите один:",
        "excel_cap": "📊 Vaksin Med — список всех филиалов",
        "km": "км от вас",
        "all_districts": "📋 Все филиалы",
        "list_title": "📋 *Список филиалов* ({start}-{end} / {total}):",
        "prev": "⬅️ Назад",
        "next": "Вперёд ➡️",
        "map_choose": "Выберите карту:",
        "yandex": "🗺 Yandex Maps",
        "google": "🗺 Google Maps",
        "office_title": "🏢 *Адрес офиса/склада*",
        "office_address": "📍 г. Ташкент, Юнусабадский район",
    }
}

def get_lang(ctx): return ctx.user_data.get("lang", "uz")

def load_df():
    try:
        # Google Sheets dan o'qish
        url = f"https://docs.google.com/spreadsheets/d/{SHEETS_ID}/export?format=xlsx"
        import urllib.request
        with urllib.request.urlopen(url) as response:
            data = response.read()
        import io as _io
        df = pd.read_excel(_io.BytesIO(data)).fillna("")
        df["filial_no"] = df["Filial №"].astype(str).str.strip().str.replace(r"\.0$","",regex=True)
        df["_sort"] = pd.to_numeric(df["filial_no"], errors="coerce").fillna(9999)
        df = df.sort_values("_sort").reset_index(drop=True)
        print(f"Google Sheets dan {len(df)} ta filial yuklandi")
        return df
    except Exception as e:
        print(f"Google Sheets xato: {e}, local fayldan o'qilmoqda...")
        try:
            df = pd.read_excel(EXCEL_FILE).fillna("")
            df["filial_no"] = df["Filial №"].astype(str).str.strip().str.replace(r"\.0$","",regex=True)
            df["_sort"] = pd.to_numeric(df["filial_no"], errors="coerce").fillna(9999)
            df = df.sort_values("_sort").reset_index(drop=True)
            return df
        except Exception as e2:
            print(f"Local fayl xato: {e2}")
            return pd.DataFrame()

def clean_number(text):
    text = str(text).strip()
    m = re.fullmatch(r"0*(\d+)", text)
    if m: return m.group(1)
    m = re.match(r"0*(\d+)\s*[-]", text)
    if m: return m.group(1)
    m = re.search(r"(\d+)", text)
    if m: return m.group(1)
    return text

def format_card(row, language):
    lg_up = "UZ" if language == "uz" else "RU"
    nomi = row.get(f"Nomi ({lg_up})", "") or row.get("Nomi (RU)", "")
    filial = str(row.get("filial_no", "")).strip()
    hudud = row.get(f"Hudud ({lg_up})", "") or row.get("Hudud (RU)", "")
    tuman = row.get(f"Tuman ({lg_up})", "") or row.get("Tuman (RU)", "")
    manzil = row.get(f"Manzil ({lg_up})", "") or row.get("Manzil (RU)", "")
    orientir = row.get(f"Orientir ({lg_up})", "") or row.get("Orientir (RU)", "")
    hours = row.get(f"Ish vaqti ({lg_up})", "") or row.get("Ish vaqti (RU)", "")
    phone = row.get("Telefon", "")

    koordinator_ismi = row.get("Koordinator ismi", "")
    koordinator_tel = row.get("Koordinator tel", "")
    mudiri_tel = row.get("Dorixona mudiri tel", "")

    lines = [f"🏥 *{nomi}*"]
    if filial and filial not in ["nan", ""]:
        lines.append(f"🔢 Filial: #{filial}")
    if hudud or tuman:
        lines.append(f"🗺 {hudud}" + (f", {tuman}" if tuman else ""))
    if manzil: lines.append(f"📍 {manzil}")
    if orientir: lines.append(f"🚩 {orientir}")
    if hours: lines.append(f"🕐 {hours}")
    if phone:
        clean_phone = str(phone).replace(" ","").replace("-","").replace("(","").replace(")","")
        lines.append(f"📞 [{phone}](tel:{clean_phone})")
    if koordinator_ismi and str(koordinator_ismi) not in ["", "nan"]:
        if koordinator_tel and str(koordinator_tel) not in ["", "nan"]:
            digits_k = re.sub(r"\D", "", str(koordinator_tel))
            tg_k = f"https://t.me/+{digits_k}"
            lines.append(f"👤 Koordinator: [{koordinator_ismi}]({tg_k})")
        else:
            lines.append(f"👤 Koordinator: {koordinator_ismi}")
    if mudiri_tel and str(mudiri_tel) not in ["", "nan"]:
        digits_m = re.sub(r"\D", "", str(mudiri_tel))
        tg_m = f"https://t.me/+{digits_m}"
        lines.append(f"💊 Dorixona mudiri: [{mudiri_tel}]({tg_m})")
    return "\n".join(str(l) for l in lines)

def phone_to_tg(phone):
    digits = re.sub(r"\D", "", str(phone))
    if not digits:
        return None
    if digits.startswith("998"):
        number = digits
    elif digits.startswith("0"):
        number = "998" + digits[1:]
    else:
        number = "998" + digits
    return f"https://t.me/+{number}"

def get_map_buttons(lat, lon, language, phone=None):
    yandex_url = f"https://maps.yandex.ru/?pt={lon},{lat}&z=17&l=map"
    google_url = f"https://maps.google.com/?q={lat},{lon}"
    buttons = [[
        InlineKeyboardButton(T[language]["yandex"], url=yandex_url),
        InlineKeyboardButton(T[language]["google"], url=google_url),
    ]]
    if phone:
        tg_url = phone_to_tg(phone)
        if tg_url:
            buttons.append([InlineKeyboardButton("💬 Telegram", url=tg_url)])
    return InlineKeyboardMarkup(buttons)

def get_buttons_no_map(language, phone=None):
    if phone:
        tg_url = phone_to_tg(phone)
        if tg_url:
            return InlineKeyboardMarkup([[InlineKeyboardButton("💬 Telegram", url=tg_url)]])
    return None

def haversine(lat1, lon1, lat2, lon2):
    R = 6371
    dlat, dlon = math.radians(lat2-lat1), math.radians(lon2-lon1)
    a = math.sin(dlat/2)**2+math.cos(math.radians(lat1))*math.cos(math.radians(lat2))*math.sin(dlon/2)**2
    return R*2*math.asin(math.sqrt(a))

def main_keyboard(language):
    return ReplyKeyboardMarkup([
        [T[language]["search_btn"]],
        [T[language]["attendance_btn"]],
        [T[language]["chat_btn"], T[language]["channel_btn"]],
        [T[language]["lang_btn"]],
    ], resize_keyboard=True)

def search_keyboard(language):
    return ReplyKeyboardMarkup([
        [T[language]["by_number"], T[language]["by_name"]],
        [T[language]["by_region"], T[language]["nearest"]],
        [T[language]["office_loc"], T[language]["excel_btn"]],
        [T[language]["map_btn"]],
        [T[language]["back"]],
    ], resize_keyboard=True)

def back_keyboard(language):
    return ReplyKeyboardMarkup([[T[language]["back"]]], resize_keyboard=True)

async def send_card(msg, row, language):
    text = format_card(row, language)
    lat = str(row.get("Latitude","")).strip()
    lon = str(row.get("Longitude","")).strip()
    phone = str(row.get("Telefon","")).strip()
    has_coords = lat not in ["","nan"] and lon not in ["","nan"]

    if has_coords:
        loc_msg = await msg.reply_location(latitude=float(lat), longitude=float(lon))
        kb = get_map_buttons(lat, lon, language, phone)
        await loc_msg.reply_text(text, parse_mode="Markdown", reply_markup=kb)
    else:
        kb = get_buttons_no_map(language, phone)
        if kb:
            await msg.reply_text(text, parse_mode="Markdown", reply_markup=kb)
        else:
            await msg.reply_text(text, parse_mode="Markdown")

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("🇺🇿 O'zbek", callback_data="lang_uz"),
        InlineKeyboardButton("🇷🇺 Русский", callback_data="lang_ru")
    ]])
    await update.message.reply_text(T["uz"]["welcome"], reply_markup=kb, parse_mode="Markdown")
    return LANG

async def set_lang(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    language = q.data.split("_")[1]
    ctx.user_data["lang"] = language
    await q.message.reply_text(T[language]["menu"], reply_markup=main_keyboard(language), parse_mode="Markdown")
    return MENU

async def menu_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    language = get_lang(ctx)
    txt = update.message.text

    if txt == T[language]["lang_btn"]:
        kb = InlineKeyboardMarkup([[
            InlineKeyboardButton("🇺🇿 O'zbek", callback_data="lang_uz"),
            InlineKeyboardButton("🇷🇺 Русский", callback_data="lang_ru")
        ]])
        await update.message.reply_text(T[language]["welcome"], reply_markup=kb, parse_mode="Markdown")
        return LANG

    elif txt == T[language]["search_btn"]:
        await update.message.reply_text(T[language]["search_menu"],
                                         reply_markup=search_keyboard(language), parse_mode="Markdown")
        return SEARCH_MENU

    elif txt == T[language]["attendance_btn"]:
        return await att_enter(update, ctx)

    elif txt == T[language]["chat_btn"]:
        kb = InlineKeyboardMarkup([[InlineKeyboardButton("💬 Chat", url=TELEGRAM_CHAT_LINK)]])
        await update.message.reply_text("💬 Telegram chatimizga xush kelibsiz!" if language == "uz" else "💬 Добро пожаловать в наш Telegram чат!", reply_markup=kb)
        return MENU

    elif txt == T[language]["channel_btn"]:
        kb = InlineKeyboardMarkup([[InlineKeyboardButton("📢 Kanal" if language == "uz" else "📢 Канал", url=TELEGRAM_CHANNEL_LINK)]])
        await update.message.reply_text("📢 Rasmiy kanalimiz:" if language == "uz" else "📢 Наш официальный канал:", reply_markup=kb)
        return MENU

    return MENU

async def search_menu_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    language = get_lang(ctx)
    txt = update.message.text

    if txt == T[language]["back"]:
        await update.message.reply_text(T[language]["menu"], reply_markup=main_keyboard(language), parse_mode="Markdown")
        return MENU

    elif txt == T[language]["nearest"]:
        kb = ReplyKeyboardMarkup([
            [KeyboardButton(T[language]["loc_btn"], request_location=True)],
            [T[language]["back"]]
        ], resize_keyboard=True)
        await update.message.reply_text(T[language]["send_loc"], reply_markup=kb)
        return LOCATION_WAIT

    elif txt == T[language]["by_region"]:
        buttons = []
        row_btns = []
        for i, vil in enumerate(VILOYATLAR.keys()):
            row_btns.append(InlineKeyboardButton(vil, callback_data=f"vil_{i}"))
            if len(row_btns) == 2:
                buttons.append(row_btns)
                row_btns = []
        if row_btns: buttons.append(row_btns)
        await update.message.reply_text(T[language]["select_viloyat"],
                                         reply_markup=InlineKeyboardMarkup(buttons))
        return SELECT_REGION

    elif txt in [T[language]["by_number"], T[language]["by_name"]]:
        ctx.user_data["stype"] = "number" if txt == T[language]["by_number"] else "name"
        prompt = T[language]["enter_number"] if ctx.user_data["stype"] == "number" else T[language]["enter_name"]
        await update.message.reply_text(prompt, reply_markup=back_keyboard(language), parse_mode="Markdown")
        return SEARCH_INPUT

    elif txt == T[language]["map_btn"]:
        kb = InlineKeyboardMarkup([[
            InlineKeyboardButton("🗺 Google My Maps", url=MY_MAPS_URL)
        ]])
        await update.message.reply_text(
            "🗺 Barcha filiallarni kartada ko'rish:" if language == "uz" else "🗺 Все филиалы на карте:",
            reply_markup=kb
        )
        return SEARCH_MENU

    elif txt == T[language]["office_loc"]:
        kb = get_map_buttons(str(OFFICE_LAT), str(OFFICE_LON), language)
        await update.message.reply_location(latitude=OFFICE_LAT, longitude=OFFICE_LON)
        await update.message.reply_text(T[language]["office_title"], parse_mode="Markdown", reply_markup=kb)
        return SEARCH_MENU

    elif txt == T[language]["excel_btn"]:
        try:
            import urllib.request as _ur
            import io as _io
            url = f"https://docs.google.com/spreadsheets/d/{SHEETS_ID}/export?format=xlsx"
            data = _ur.urlopen(url).read()
            df_all = pd.read_excel(_io.BytesIO(data)).fillna("")
            # Faqat kerakli ustunlar
            cols = ["Filial №", "Nomi (RU)", "Manzil (RU)", "Telefon"]
            cols_exist = [c for c in cols if c in df_all.columns]
            df_excel = df_all[cols_exist].copy()
            df_excel["Filial №"] = df_excel["Filial №"].astype(str).str.replace(r"\.0$","",regex=True)
            # Excel faylga yozish
            output = _io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_excel.to_excel(writer, index=False, sheet_name="Filiallar")
            output.seek(0)
            await update.message.reply_document(
                output,
                filename="filiallar.xlsx",
                caption=T[language]["excel_cap"]
            )
        except Exception as e:
            print(f"Excel xato: {e}")
            await update.message.reply_text("❌ Xatolik yuz berdi!" if language == "uz" else "❌ Произошла ошибка!")
        return SEARCH_MENU

    return SEARCH_MENU

async def send_list_page(msg, ctx, language):
    records = ctx.user_data.get("list_df", [])
    page = ctx.user_data.get("list_page", 0)
    total = len(records)
    start = page * PAGE_SIZE
    end = min(start + PAGE_SIZE, total)
    chunk = records[start:end]

    col = "Nomi (UZ)" if language == "uz" else "Nomi (RU)"
    text = T[language]["list_title"].format(start=start+1, end=end, total=total) + "\n\n"
    buttons = []
    for i, row in enumerate(chunk):
        nomi = str(row.get(col, ""))[:25]
        filial = str(row.get("filial_no",""))
        label = f"#{filial} — {nomi}" if filial not in ["","nan"] else nomi
        buttons.append([InlineKeyboardButton(label, callback_data=f"list_{start+i}")])

    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton(T[language]["prev"], callback_data="listpage_prev"))
    if end < total:
        nav.append(InlineKeyboardButton(T[language]["next"], callback_data="listpage_next"))
    if nav: buttons.append(nav)
    buttons.append([InlineKeyboardButton(T[language]["back"], callback_data="listpage_back")])

    await msg.reply_text(text, reply_markup=InlineKeyboardMarkup(buttons), parse_mode="Markdown")

async def list_page_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    language = get_lang(ctx)

    if q.data == "listpage_prev":
        ctx.user_data["list_page"] = max(0, ctx.user_data.get("list_page",0) - 1)
        await q.message.delete()
        await send_list_page(q.message, ctx, language)
        return LIST_PAGE
    elif q.data == "listpage_next":
        ctx.user_data["list_page"] = ctx.user_data.get("list_page",0) + 1
        await q.message.delete()
        await send_list_page(q.message, ctx, language)
        return LIST_PAGE
    elif q.data == "listpage_back":
        await q.message.reply_text(T[language]["menu"], reply_markup=main_keyboard(language), parse_mode="Markdown")
        return MENU
    elif q.data.startswith("list_"):
        idx = int(q.data.split("_")[1])
        records = ctx.user_data.get("list_df", [])
        if idx < len(records):
            await send_card(q.message, records[idx], language)
        return LIST_PAGE

    return LIST_PAGE

async def search_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    language = get_lang(ctx)
    txt = update.message.text.strip()
    if txt == T[language]["back"]:
        await update.message.reply_text(T[language]["search_menu"],
                                         reply_markup=search_keyboard(language), parse_mode="Markdown")
        return SEARCH_MENU

    df = load_df()
    if df.empty:
        await update.message.reply_text("❌ Ma'lumotlar bazasi topilmadi!")
        return SEARCH_MENU

    # Avtomatik aniqlash: raqam yoki nom
    stype = ctx.user_data.get("stype","name")
    if txt.strip().isdigit():
        stype = "number"

    if stype == "number":
        query_no = clean_number(txt)
        results = df[df["filial_no"] == query_no]
        if results.empty:
            results = df[df["filial_no"].str.contains(query_no, na=False)]
    else:
        variants = search_variants(txt)
        mask = pd.Series([False]*len(df), index=df.index)
        for v in variants:
            mask = mask | df["Nomi (UZ)"].str.contains(v, case=False, na=False)
            mask = mask | df["Nomi (RU)"].str.contains(v, case=False, na=False)
        results = df[mask]
        if results.empty:
            all_names = df["Nomi (UZ)"].tolist() + df["Nomi (RU)"].tolist()
            matches = fuzz_process.extract(txt, all_names, limit=5)
            good = [m[0] for m in matches if m[1] >= 45]
            if good:
                await update.message.reply_text(T[language]["similar"])
                results = df[df["Nomi (UZ)"].isin(good) | df["Nomi (RU)"].isin(good)]

    if results.empty:
        await update.message.reply_text(T[language]["not_found"])
        return SEARCH_INPUT

    if len(results) == 1:
        await send_card(update.message, results.iloc[0].to_dict(), language)
        await update.message.reply_text(
            "🔍 Keyingi qidiruv uchun yozing:" if language == "uz" else "🔍 Введите следующий запрос:",
            reply_markup=back_keyboard(language)
        )
        return SEARCH_INPUT

    ctx.user_data["results"] = results.to_dict("records")
    col = "Nomi (UZ)" if language == "uz" else "Nomi (RU)"
    buttons = []
    for i, (_, row) in enumerate(results.head(10).iterrows()):
        nomi = str(row.get(col,""))[:25]
        filial = str(row.get("filial_no",""))
        label = f"#{filial} — {nomi}" if filial not in ["","nan"] else nomi
        buttons.append([InlineKeyboardButton(label, callback_data=f"sel_{i}")])
    await update.message.reply_text(T[language]["found_many"].format(n=len(results)),
                                     reply_markup=InlineKeyboardMarkup(buttons), parse_mode="Markdown")
    return SELECT_RESULT

async def select_result(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    language = get_lang(ctx)
    idx = int(q.data.split("_")[1])
    results = ctx.user_data.get("results",[])
    if idx < len(results):
        await send_card(q.message, results[idx], language)
    await q.message.reply_text(
        "🔍 Keyingi qidiruv uchun yozing:" if language == "uz" else "🔍 Введите следующий запрос:",
        reply_markup=back_keyboard(language)
    )
    return SEARCH_INPUT

async def location_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    language = get_lang(ctx)
    if update.message.text == T[language]["back"]:
        await update.message.reply_text(T[language]["search_menu"],
                                         reply_markup=search_keyboard(language), parse_mode="Markdown")
        return SEARCH_MENU
    if not update.message.location: return LOCATION_WAIT

    ulat = update.message.location.latitude
    ulon = update.message.location.longitude
    df = load_df()
    df["_lat"] = pd.to_numeric(df.get("Latitude", pd.Series(dtype=float)), errors="coerce")
    df["_lon"] = pd.to_numeric(df.get("Longitude", pd.Series(dtype=float)), errors="coerce")
    valid = df.dropna(subset=["_lat","_lon"]).copy()

    if valid.empty:
        await update.message.reply_text("❌ Koordinatalar hali kiritilmagan!")
        await update.message.reply_text(T[language]["search_menu"], reply_markup=search_keyboard(language), parse_mode="Markdown")
        return SEARCH_MENU

    valid["_dist"] = valid.apply(lambda r: haversine(ulat,ulon,r["_lat"],r["_lon"]), axis=1)
    for _, row in valid.nsmallest(3,"_dist").iterrows():
        r = row.to_dict()
        text = format_card(r, language) + f"\n📏 {row['_dist']:.1f} {T[language]['km']}"
        lat, lon = str(row["_lat"]), str(row["_lon"])
        kb = get_map_buttons(lat, lon, language)
        await update.message.reply_text(text, parse_mode="Markdown", reply_markup=kb)
        await update.message.reply_location(latitude=row["_lat"], longitude=row["_lon"])

    await update.message.reply_text(T[language]["search_menu"], reply_markup=search_keyboard(language), parse_mode="Markdown")
    return SEARCH_MENU

async def select_region(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    language = get_lang(ctx)
    vil_idx = int(q.data.split("_")[1])
    vil_name = list(VILOYATLAR.keys())[vil_idx]
    ctx.user_data["viloyat"] = vil_name
    tumanlar = VILOYATLAR[vil_name]

    buttons = [[InlineKeyboardButton(T[language]["all_districts"], callback_data="tuman_all")]]
    row_btns = []
    for i, tuman in enumerate(tumanlar):
        row_btns.append(InlineKeyboardButton(tuman, callback_data=f"tuman_{i}"))
        if len(row_btns) == 2:
            buttons.append(row_btns)
            row_btns = []
    if row_btns: buttons.append(row_btns)
    await q.message.reply_text(f"📍 *{vil_name}*\n{T[language]['select_tuman']}",
                                 reply_markup=InlineKeyboardMarkup(buttons), parse_mode="Markdown")
    return SELECT_DISTRICT

async def select_district(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    language = get_lang(ctx)
    vil_name = ctx.user_data.get("viloyat","")
    tumanlar = VILOYATLAR.get(vil_name, [])
    df = load_df()

    if q.data == "tuman_all":
        lower_t = [t.lower() for t in tumanlar]
        results = df[df["Hudud (RU)"].str.lower().isin(lower_t)]
    else:
        tuman_idx = int(q.data.split("_")[1])
        tuman_name = tumanlar[tuman_idx]
        results = df[df["Hudud (RU)"].str.lower() == tuman_name.lower()]

    if results.empty:
        await q.message.reply_text(T[language]["not_found"])
        await q.message.reply_text(T[language]["search_menu"], reply_markup=search_keyboard(language), parse_mode="Markdown")
        return SEARCH_MENU

    if len(results) <= 3:
        for _, row in results.iterrows():
            await send_card(q.message, row.to_dict(), language)
        await q.message.reply_text(T[language]["search_menu"], reply_markup=search_keyboard(language), parse_mode="Markdown")
        return SEARCH_MENU

    ctx.user_data["results"] = results.to_dict("records")
    col = "Nomi (UZ)" if language == "uz" else "Nomi (RU)"
    buttons = []
    for i, (_, row) in enumerate(results.head(15).iterrows()):
        nomi = str(row.get(col,""))[:25]
        filial = str(row.get("filial_no",""))
        label = f"#{filial} — {nomi}" if filial not in ["","nan"] else nomi
        buttons.append([InlineKeyboardButton(label, callback_data=f"sel_{i}")])
    await q.message.reply_text(T[language]["found_many"].format(n=len(results)),
                                 reply_markup=InlineKeyboardMarkup(buttons), parse_mode="Markdown")
    return SELECT_RESULT

def main():
    app = Application.builder().token(TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            LANG: [CallbackQueryHandler(set_lang, pattern="^lang_")],
            MENU: [MessageHandler(filters.TEXT & ~filters.COMMAND, menu_handler)],
            SEARCH_MENU: [MessageHandler(filters.TEXT & ~filters.COMMAND, search_menu_handler)],
            SEARCH_INPUT: [MessageHandler(filters.TEXT & ~filters.COMMAND, search_handler)],
            SELECT_RESULT: [CallbackQueryHandler(select_result, pattern="^sel_")],
            SELECT_REGION: [CallbackQueryHandler(select_region, pattern="^vil_")],
            SELECT_DISTRICT: [
                CallbackQueryHandler(select_district, pattern="^tuman_"),
                CallbackQueryHandler(select_result, pattern="^sel_"),
            ],
            LIST_PAGE: [
                CallbackQueryHandler(list_page_handler, pattern="^list"),
            ],
            LOCATION_WAIT: [
                MessageHandler(filters.LOCATION, location_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, location_handler),
            ],
            **get_att_states(),
        },
        fallbacks=[CommandHandler("start", start)],
    )
    app.add_handler(conv)
    print("✅ Vaksin Med bot ishga tushdi!")
    from telegram.ext import CommandHandler as CmdHandler
    app.add_handler(CmdHandler("init_month", cmd_init_month))
    app.add_handler(CmdHandler("calc_hours", cmd_calc_hours))
    app.add_handler(CmdHandler("sync_pharmacists", cmd_sync_pharmacists))
    app.add_handler(CmdHandler("fill_codes", cmd_fill_codes))
    app.run_polling()

if __name__ == "__main__":
    main()
