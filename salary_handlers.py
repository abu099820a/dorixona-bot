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
SALARY_PHONE_WAIT = 504


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

# Maosh jadvalidagi ustunlar (pozitsiya bo'yicha, 1-indeksda) — YANGI format
# (Davomat jadvaliga o'xshab, alohida Filial/Ismi/Telefon ustunlari bilan):
# A=1 Filial (raqam bilan, "1 - ТАШМИ-1") | B=2 Ismi | C=3 Telefon
# D=4 Reja(keyingi oy) | E=5 Reja(joriy oy) | F=6 Savdo | G=7 Rejadan farq
# H=8 Foiz | I=9 Oylik % (bonus) | J=10 Fiksa | K=11 Reja bonusi
# L=12 Avans | M=13 Pereuchyot shtraf | N=14 Kech/erta shtraf
# O=15 Srok shtraf | P=16 Umumiy summa | Q=17 Plastik kartaga tushadigan
SAL_COL_FILIAL = 1
SAL_COL_ISMI = 2
SAL_COL_TELEFON = 3
SAL_COL_REJA_KEYINGI = 4
SAL_COL_REJA_JORIY = 5
SAL_COL_SAVDO = 6
SAL_COL_FARQ = 7
SAL_COL_FOIZ = 8
SAL_COL_OYLIK_PERCENT = 9
SAL_COL_FIKSA = 10
SAL_COL_REJA_BONUS = 11
SAL_COL_AVANS = 12
SAL_COL_SHTRAF_PEREUCHYOT = 13
SAL_COL_SHTRAF_VAQT = 14
SAL_COL_SHTRAF_SROK = 15
SAL_COL_JAMI = 16
SAL_COL_KARTA = 17


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


def _filial_kod(filial: str) -> str:
    """'6 - ЮНУСАБАД 7' → '6'"""
    m = re.match(r"^(\d+)", str(filial).strip())
    return m.group(1) if m else ""


def _sync_one_sheet(ws, farmatsevtlar: list) -> dict:
    """
    Berilgan varaqqa (Oylik yoki Aksiya) Farmatsevtlar ro'yxatidagi
    xodimlardan hali mavjud bo'lmaganlarini qo'shadi.

    MUHIM: bu funksiya BUTUN jadval mazmunini XOTIRADA (Python ichida)
    qayta quradi va faqat OXIRIDA 1-2 ta Google Sheets so'rovi yuboradi
    (bitta "yozish" + bitta "merge"). Avvalgi versiya har bir xodim
    uchun alohida-alohida so'rov yuborar edi — bu ko'p xodimda Google'ning
    "daqiqasiga so'rov" kvotasidan (429 xato) oshib ketishga va, jarayon
    yarim yo'lda uzilib qolsa, jadval tartibi buzilib qolishiga sabab
    bo'lgan edi.

    Qaytaradi: {"added": [...], "skipped": int}
    """
    all_values = ws.get_all_values()
    header = all_values[0] if all_values else []
    data_rows = [list(r) for r in all_values[1:]]
    ncols = max(len(header), 17)

    def _pad(row):
        row = list(row) + [""] * (ncols - len(row))
        return row[:ncols]

    data_rows = [_pad(r) for r in data_rows]

    # Mavjud telefon raqamlari
    existing_phones = set()
    for row in data_rows:
        if row[2]:
            existing_phones.add(_sal_normalize_phone(row[2]))

    # Filial bo'yicha guruhlab, qo'shilishi kerak bo'lganlarni ajratamiz
    to_add_by_filial = {}
    added = []
    for f in farmatsevtlar:
        phone_norm = _sal_normalize_phone(f["telefon"])
        if not phone_norm or phone_norm in existing_phones:
            continue
        kod = _filial_kod(f["filial"])
        to_add_by_filial.setdefault(kod, []).append(f)
        existing_phones.add(phone_norm)
        added.append(f["ismi"])

    # DIQQAT: bu yerda "if not added: return" QILMAYMIZ — chunki merge
    # (birlashtirish) qadami har doim ishlashi kerak, hatto yangi xodim
    # qo'shilmasa ham (masalan qayta ishga tushirilganda). Avvalgi
    # versiyada shu yerda erta chiqib ketilardi va shu sabab merge hech
    # qachon bajarilmay qolgan edi.

    # Yangi ro'yxatni original tartibni SAQLAGAN holda quramiz: mavjud
    # qatorlar orasidan, har bir filial guruhi TUGAGAN joyda, o'sha
    # filialga tegishli yangi xodimlarni qo'shib chiqamiz
    new_rows = []
    n = len(data_rows)
    i = 0
    while i < n:
        row = data_rows[i]
        new_rows.append(row)
        cur_kod = _filial_kod(row[0]) if row[0] else None
        next_kod = None
        if i + 1 < n and data_rows[i + 1][0]:
            next_kod = _filial_kod(data_rows[i + 1][0])
        if cur_kod and cur_kod != next_kod and cur_kod in to_add_by_filial:
            for f in to_add_by_filial.pop(cur_kod):
                new_rows.append(_pad([f["filial"], f["ismi"], f["telefon"]]))
        i += 1

    # Jadvalda umuman topilmagan filiallar (masalan yangi filial) — oxiriga
    for kod, flist in to_add_by_filial.items():
        for f in flist:
            new_rows.append(_pad([f["filial"], f["ismi"], f["telefon"]]))

    # 1) Agar yangi xodim qo'shilgan bo'lsa — butun ma'lumotni BITTA
    #    so'rov bilan yozib qo'yamiz (aks holda yozish shart emas —
    #    kvota tejash uchun)
    if added:
        needed_rows = 1 + len(new_rows)
        if ws.row_count < needed_rows:
            ws.resize(rows=needed_rows)
        ws.update(
            f"A2:{col_letter(ncols)}{needed_rows}",
            new_rows, value_input_option="USER_ENTERED"
        )

    # 2) Filial ustunini (A) guruhlar bo'ylab BITTA so'rovda qayta merge
    #    qilamiz — bu HAR DOIM bajariladi, yangi xodim bo'lsin-bo'lmasin
    _merge_all_filial_groups(ws, new_rows)

    return {"added": added, "skipped": len(farmatsevtlar) - len(added)}


def col_letter(n: int) -> str:
    """1 -> A, 2 -> B, 27 -> AA ..."""
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _merge_all_filial_groups(ws, data_rows: list):
    """
    A ustunidagi filial guruhlarini (yonma-yon bir xil kodli qatorlarni)
    BITTA batch_update so'rovi bilan qayta merge qiladi (avval butun
    ustunni unmerge qilib, so'ng har bir guruhni alohida merge qilish
    so'rovlarini bitta so'rovga jamlab yuboradi).
    """
    n = len(data_rows)
    if n == 0:
        return
    requests = [{"unmergeCells": {"range": {
        "sheetId": ws.id,
        "startRowIndex": 1, "endRowIndex": 1 + n,
        "startColumnIndex": 0, "endColumnIndex": 1,
    }}}]

    i = 0
    while i < n:
        kod = _filial_kod(data_rows[i][0]) if data_rows[i][0] else None
        j = i
        while j + 1 < n:
            next_kod = _filial_kod(data_rows[j + 1][0]) if data_rows[j + 1][0] else None
            if next_kod != kod:
                break
            j += 1
        if kod and j > i:
            requests.append({"mergeCells": {"range": {
                "sheetId": ws.id,
                "startRowIndex": 1 + i, "endRowIndex": 1 + j + 1,
                "startColumnIndex": 0, "endColumnIndex": 1,
            }, "mergeType": "MERGE_ALL"}})
        i = j + 1

    try:
        ws.spreadsheet.batch_update({"requests": requests})
    except Exception as e:
        logger.error(f"[MAOSH] Merge xato: {e}")


def sync_oylik_sheet() -> dict:
    """
    Farmatsevtlar jadvalidagi BARCHA xodimlarni "Oylik" va "Aksiya"
    varaqlariga (hali mavjud bo'lmaganlarini, telefon raqami bo'yicha)
    qo'shib chiqadi. Admin buyrug'i orqali (masalan bir martalik
    "orqaga qoldirilgan" ro'yxatdan o'tganlarni to'ldirish uchun) ishga
    tushiriladi.
    """
    result = {"oylik": {"added": [], "skipped": 0}, "aksiya": {"added": [], "skipped": 0}, "error": None}
    try:
        client = _get_client()

        ph_ws = client.open_by_key(PHARMACY_SHEET_ID).worksheet("Farmatsevtlar")
        records = ph_ws.get_all_records()
        farmatsevtlar = []
        for row in records:
            ismi = str(row.get("Ismi", "")).strip()
            filial = str(row.get("Filial", "")).strip()
            telefon = str(row.get("Telefon", "")).strip()
            if ismi and filial and telefon:
                farmatsevtlar.append({"ismi": ismi, "filial": filial, "telefon": telefon})

        sh = client.open_by_key(SALARY_SHEET_ID)
        mavjud_varaqlar = [w.title for w in sh.worksheets()]

        try:
            ws_oylik = sh.worksheet(SALARY_WS_NAME)
            result["oylik"] = _sync_one_sheet(ws_oylik, farmatsevtlar)
        except gspread.exceptions.WorksheetNotFound:
            result["error"] = (
                f"'{SALARY_WS_NAME}' varag'i topilmadi. "
                f"Mavjud varaqlar: {', '.join(mavjud_varaqlar)}"
            )

        try:
            ws_aksiya = sh.worksheet(AKSIYA_WS_NAME)
            result["aksiya"] = _sync_one_sheet(ws_aksiya, farmatsevtlar)
        except gspread.exceptions.WorksheetNotFound:
            result["aksiya_error"] = (
                f"'{AKSIYA_WS_NAME}' varag'i topilmadi. "
                f"Mavjud varaqlar: {', '.join(mavjud_varaqlar)}"
            )

        return result

    except Exception as e:
        logger.error(f"[MAOSH] sync_oylik_sheet xato: {e}")
        result["error"] = str(e)
        return result


async def cmd_sync_oylik(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/sync_oylik — Farmatsevtlar ro'yxatidagi hammani Oylik/Aksiya varaqlariga qo'shadi (admin)."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return
    msg = await update.message.reply_text("⏳ Sinxronizatsiya boshlanmoqda...")
    try:
        res = await run_write(sync_oylik_sheet)
        if res.get("error"):
            await msg.edit_text(f"❌ Xato: {res['error']}")
            return
        lines = ["✅ *Sinxronizatsiya tugadi!*\n"]
        lines.append(f"📄 Oylik: +{len(res['oylik']['added'])} ta qo'shildi")
        if res.get("aksiya_error"):
            lines.append(f"⚠️ Aksiya: {res['aksiya_error']}")
        else:
            lines.append(f"🎁 Aksiya: +{len(res['aksiya']['added'])} ta qo'shildi")
        await msg.edit_text("\n".join(lines), parse_mode="Markdown")
    except Exception as e:
        await msg.edit_text(f"❌ Xato: {e}")


def get_farmatsevt_salary_by_phone(phone: str) -> dict | None:
    """
    Berilgan TELEFON RAQAMI orqali (TelegramID orqali EMAS) "Oylik" va
    "Aksiya" varaqlaridan (SALARY_SHEET_ID) mos qatorlarni qidiradi va
    ikkalasini birlashtirib to'liq hisobotni qaytaradi.

    MUHIM: bu funksiya Farmatsevtlar jadvalidagi TelegramID ustuniga
    umuman bog'liq emas — faqat berilgan telefon raqamini Oylik/Aksiya
    jadvalidagi C ustuni bilan solishtiradi. Shuning uchun, agar
    foydalanuvchining TelegramID'si Farmatsevtlar jadvalida to'g'ri
    saqlanmagan bo'lsa ham (masalan qo'lda qo'shilgan xodim), telefon
    raqami to'g'ri bo'lsa — natija topiladi.

    Jadval tuzilishi (ikkala varaqda ham bir xil, Davomat kabi):
        A: Filial (raqam bilan) | B: Ismi | C: Telefon
        D: Reja(keyingi oy) | E: Reja(joriy oy) | F: Savdo | G: Farq
        H: Foiz | I: Oylik % (bonus) | J: Fiksa | K: Reja bonusi
        L: Avans | M/N/O: shtraflar | P: Umumiy summa | Q: Karta

    Qaytaradi to'liq breakdown dict yoki None (topilmasa).
    """
    def _find_row_by_phone(ws, target_phone):
        all_values = ws.get_all_values()
        seen_phones = []
        for row in all_values[1:]:
            if not row or not row[0]:
                continue

            def _cell(col, _row=row):
                idx = col - 1
                return _row[idx] if idx < len(_row) else ""

            telefon_cell = str(_cell(SAL_COL_TELEFON)).strip()
            if not telefon_cell:
                continue  # bo'sh (sarlavha) qator — xodim emas
            norm = _sal_normalize_phone(telefon_cell)
            seen_phones.append((telefon_cell, norm))
            if norm != target_phone:
                continue
            print(f"[MAOSH] MOS TOPILDI: xom='{telefon_cell}' normalized='{norm}'")

            def _num(col, _row=row):
                v = _row[col - 1] if col - 1 < len(_row) else ""
                if v == "" or v is None:
                    return 0
                try:
                    return float(str(v).replace(",", ".").replace(" ", ""))
                except Exception:
                    return 0

            return {
                "ismi": str(_cell(SAL_COL_ISMI)).strip(),
                "filial": str(_cell(SAL_COL_FILIAL)).strip(),
                "reja_keyingi": _num(SAL_COL_REJA_KEYINGI),
                "reja_joriy": _num(SAL_COL_REJA_JORIY),
                "savdo": _num(SAL_COL_SAVDO),
                "reja_farq": _num(SAL_COL_FARQ),
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
        print(f"[MAOSH] '{ws.title}' da mos topilmadi. Ko'rilgan telefon(lar) "
              f"(xom -> normalized), birinchi 15 tasi: {seen_phones[:15]}")
        return None

    try:
        if not phone:
            return None
        target_phone = _sal_normalize_phone(phone)
        print(f"[MAOSH] Qidirilayotgan telefon: '{target_phone}'")

        client = _get_client()
        sh = client.open_by_key(SALARY_SHEET_ID)
        mavjud_varaqlar = [w.title for w in sh.worksheets()]
        print(f"[MAOSH] SALARY_SHEET_ID dagi varaqlar: {mavjud_varaqlar}")

        oylik_data = None
        try:
            ws_oylik = sh.worksheet(SALARY_WS_NAME)
            oylik_data = _find_row_by_phone(ws_oylik, target_phone)
            print(f"[MAOSH] '{SALARY_WS_NAME}' da topildimi: {oylik_data is not None}")
        except gspread.exceptions.WorksheetNotFound:
            print(f"[MAOSH] '{SALARY_WS_NAME}' varag'i topilmadi. Mavjudlar: {mavjud_varaqlar}")

        aksiya_data = None
        try:
            ws_aksiya = sh.worksheet(AKSIYA_WS_NAME)
            aksiya_data = _find_row_by_phone(ws_aksiya, target_phone)
            print(f"[MAOSH] '{AKSIYA_WS_NAME}' da topildimi: {aksiya_data is not None}")
        except gspread.exceptions.WorksheetNotFound:
            print(f"[MAOSH] '{AKSIYA_WS_NAME}' varag'i topilmadi. Mavjudlar: {mavjud_varaqlar}")

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
        logger.error(f"[MAOSH] get_farmatsevt_salary_by_phone xato: {e}")
        return None


def get_farmatsevt_salary(telegram_id) -> dict | None:
    """
    ESKI (zaxira) yo'l: TelegramID orqali Farmatsevtlar jadvalidan
    telefon raqamini topib, so'ng get_farmatsevt_salary_by_phone() ni
    chaqiradi. Endi asosiy oqim buni ishlatmaydi (o'rniga foydalanuvchi
    telefon raqamini to'g'ridan-to'g'ri kiritadi/ulashadi) — lekin
    boshqa joylarda kerak bo'lib qolishi mumkinligi uchun saqlanmoqda.
    """
    phone = _get_phone_by_telegram_id(telegram_id)
    print(f"[MAOSH] (zaxira yo'l) TelegramID={telegram_id} -> telefon='{phone}'")
    if not phone:
        return None
    return get_farmatsevt_salary_by_phone(phone)


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
        # 1) Avval sessiyada bormi tekshiramiz (tezroq)
        phone = ctx.user_data.get("salary_phone") or ctx.user_data.get("att_phone")

        # 2) Bo'lmasa — Farmatsevtlar jadvalidan TelegramID orqali avtomatik
        #    topishga harakat qilamiz (bu DOIMIY manba, bot qayta ishga
        #    tushsa ham yo'qolmaydi — foydalanuvchi har safar qayta telefon
        #    yubormasligi uchun)
        if not phone:
            user_id = update.effective_user.id
            phone = await run_read(_get_phone_by_telegram_id, user_id)
            if phone:
                ctx.user_data["salary_phone"] = phone

        if phone:
            return await _show_salary_report(update, ctx, phone)

        # 3) Faqat shu ikkalasi ham muvaffaqiyatsiz bo'lsa — telefon so'raymiz
        #    (masalan ro'yxatdan o'tish yozuvi buzilgan/topilmagan hollarda)
        from telegram import ReplyKeyboardMarkup, KeyboardButton
        kb = ReplyKeyboardMarkup([
            [KeyboardButton("📱 Telefon raqamimni yuborish", request_contact=True)],
            [back_txt],
        ], resize_keyboard=True)
        await update.message.reply_text(
            "📱 Maosh ma'lumotingizni ko'rish uchun telefon raqamingizni yuboring:"
            if language == "uz" else
            "📱 Отправьте номер телефона, чтобы посмотреть данные о зарплате:",
            reply_markup=kb,
        )
        return SALARY_PHONE_WAIT

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


async def _show_salary_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE, phone: str):
    """Berilgan telefon raqami bo'yicha maosh hisobotini topib ko'rsatadi."""
    language = ctx.user_data.get("lang", "uz")
    data = await run_read(get_farmatsevt_salary_by_phone, phone)

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
    ]

    reja_joriy = g("reja_joriy")
    savdo = g("savdo")
    if reja_joriy:
        bajarildi_foiz = (savdo / reja_joriy) * 100
        lines.append(f"🎯 Reja (joriy oy): {_fmt(reja_joriy)} so'm")
        lines.append(f"📊 Savdo: {_fmt(savdo)} so'm")
        lines.append(f"✅ Bajarildi: *{bajarildi_foiz:.1f}%*")
        if g("reja_farq"):
            farq = g("reja_farq")
            belgi = "🔺" if farq >= 0 else "🔻"
            lines.append(f"{belgi} Rejadan farq: {_fmt(farq)} so'm")
    else:
        lines.append(f"📊 Savdo: {_fmt(savdo)} so'm")

    if g("reja_keyingi"):
        lines.append(f"📅 Reja (keyingi oy): {_fmt(g('reja_keyingi'))} so'm")

    lines.append("")
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


async def salary_phone_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """SALARY_PHONE_WAIT holatida: telefon raqamini (kontakt orqali) qabul qiladi."""
    language = ctx.user_data.get("lang", "uz")
    back_txt = "⬅️ Назад" if language == "ru" else "⬅️ Orqaga"

    if update.message.text == back_txt:
        await update.message.reply_text(
            "📊 *Отчёт va to'lovlar*\n\nBo'limni tanlang:" if language == "uz"
            else "📊 *Отчёт и оплаты*\n\nВыберите раздел:",
            parse_mode="Markdown",
            reply_markup=reports_keyboard(language),
        )
        return REPORTS_MENU

    contact = update.message.contact
    if not contact:
        await update.message.reply_text(
            "❌ Iltimos, tugma orqali telefon raqamingizni yuboring."
            if language == "uz" else
            "❌ Пожалуйста, отправьте номер телефона через кнопку."
        )
        return SALARY_PHONE_WAIT

    phone = contact.phone_number
    ctx.user_data["salary_phone"] = phone
    return await _show_salary_report(update, ctx, phone)


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
        SALARY_PHONE_WAIT: [
            MessageHandler(filters.CONTACT, salary_phone_handler),
            MessageHandler(filters.TEXT & ~filters.COMMAND, salary_phone_handler),
        ],
    }
