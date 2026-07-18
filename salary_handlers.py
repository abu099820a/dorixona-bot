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
FIRM_WAIT_ZIP = 505
FIRM_REG_NAME = 506
FIRM_REG_PHONE = 507
ADMIN_FIRM_LOOKUP = 508


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
SAL_COL_AKSIYA = 18  # Yangi: Aksiya endi alohida varaq emas, shu ustunda


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

    # TA'MIRLASH: A ustuni (Filial) bo'sh qolgan qatorlarni (avvalgi
    # buzilgan "merge" operatsiyasi tufayli) yuqoridagi oxirgi ko'ringan
    # filial nomi bilan to'ldiramiz ("forward-fill"). Bu shu safargi
    # sinxronizatsiya davomida barcha eski yo'qolgan xodimlarni ham
    # tiklab beradi, chunki keyinroq bu ma'lumot qaytadan yozib qo'yiladi.
    effective_filial = ""
    repaired_count = 0
    for row in data_rows:
        if row[0]:
            effective_filial = row[0]
        elif row[1] and effective_filial:
            row[0] = effective_filial
            repaired_count += 1
    if repaired_count:
        print(f"[MAOSH] {ws.title}: {repaired_count} ta qatorda bo'sh Filial katagi tiklandi")

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

    # Butun ma'lumotni (ta'mirlangan + yangi qo'shilganlar bilan) BITTA
    # so'rov bilan yozib qo'yamiz — bu HAR DOIM bajariladi (yangi xodim
    # bo'lsin-bo'lmasin), chunki ta'mirlash har safar kerak bo'lishi mumkin
    if new_rows:
        needed_rows = 1 + len(new_rows)
        if ws.row_count < needed_rows:
            ws.resize(rows=needed_rows)
        ws.update(
            f"A2:{col_letter(ncols)}{needed_rows}",
            new_rows, value_input_option="USER_ENTERED"
        )

    # MUHIM: endi katakchalarni MERGE qilmaymiz — chunki Google Sheets
    # merge qilinganda guruhning birinchi qatoridan boshqa barcha
    # qatorlardagi qiymatni HAQIQATAN o'chirib tashlaydi (API orqali
    # o'qiganda ham bo'sh chiqadi), bu esa filialning boshqa xodimlari
    # "yo'qolib qolishiga" olib kelgan edi.

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
        effective_filial = ""
        for row in all_values[1:]:
            if row and row[0]:
                effective_filial = str(row[0]).strip()
            if not row or len(row) < 2 or not row[1]:
                continue  # bo'sh qator yoki Ismi yo'q — xodim emas

            def _cell(col, _row=row):
                idx = col - 1
                return _row[idx] if idx < len(_row) else ""

            telefon_cell = str(_cell(SAL_COL_TELEFON)).strip()
            if not telefon_cell:
                continue
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
                    return float(re.sub(r"\s+", "", str(v)).replace(",", "."))
                except Exception:
                    return 0

            return {
                "ismi": str(_cell(SAL_COL_ISMI)).strip(),
                "filial": effective_filial,
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
                "aksiya": _num(SAL_COL_AKSIYA),
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

        if not oylik_data:
            return None

        result = dict(oylik_data)
        result["aksiya_jami"] = oylik_data.get("aksiya", 0)
        result["oylik_jami"] = oylik_data.get("jami", 0)
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


def get_firma_map() -> dict:
    """
    PHARMACY_SHEET_ID ichidagi "Firmalar" varag'idan (register_handlers.py
    orqali to'ldiriladi) firma nomi -> TelegramID lug'atini qaytaradi.
    Faqat TelegramID to'ldirilgan (ro'yxatdan o'tgan) firmalar kiradi.

    Jadval tuzilishi: Firma nomi | Telefon | TelegramID | FileID | FileName
    """
    try:
        from register_handlers import _get_firmalar_ws
        ws = _get_firmalar_ws()
        records = ws.get_all_records()

        firma_map = {}
        for row in records:
            firma = str(row.get("Firma nomi", "")).strip()
            tid = str(row.get("TelegramID", "")).strip()
            if firma and tid and tid not in ("", "0"):
                firma_map[firma.upper()] = tid

        logger.info(f"[FIRMS] {len(firma_map)} ta ro'yxatdan o'tgan firma topildi")
        return firma_map
    except Exception as e:
        logger.error(f"[FIRMS] get_firma_map xato: {e}")
        return {}


def save_firma_file(firma_nomi: str, file_id: str, file_name: str) -> bool:
    """
    Berilgan firma qatoriga fayl ID va nomini saqlaydi — bu keyinchalik
    firma o'zi "Hisobotimni olish" tugmasini bosganda shu faylni qayta
    yuborish uchun ishlatiladi (Telegram file_id doimiy amal qiladi,
    faylni qayta serverga yuklash shart emas).
    """
    try:
        from register_handlers import _get_firmalar_ws
        ws = _get_firmalar_ws()
        all_values = ws.get_all_values()

        for i, row in enumerate(all_values[1:], start=2):
            if row and _norm_firma_nomi(row[0]) == _norm_firma_nomi(firma_nomi):
                ws.update_cell(i, 4, file_id)
                ws.update_cell(i, 5, file_name)
                return True
        return False
    except Exception as e:
        logger.error(f"[FIRMS] save_firma_file xato: {e}")
        return False


def get_firma_file_by_telegram_id(telegram_id) -> dict | None:
    """
    Berilgan TelegramID bo'yicha "Firmalar" varag'idan firma qatorini
    topib, saqlangan fayl ma'lumotini (file_id, file_name, firma_nomi)
    qaytaradi. Fayl hali yuklanmagan bo'lsa — file_id bo'sh bo'ladi.
    """
    try:
        from register_handlers import _get_firmalar_ws
        ws = _get_firmalar_ws()
        records = ws.get_all_records()
        tid = str(telegram_id).strip()

        for row in records:
            if str(row.get("TelegramID", "")).strip() == tid:
                return {
                    "firma_nomi": str(row.get("Firma nomi", "")).strip(),
                    "file_id": str(row.get("FileID", "")).strip(),
                    "file_name": str(row.get("FileName", "")).strip(),
                }
        return None
    except Exception as e:
        logger.error(f"[FIRMS] get_firma_file_by_telegram_id xato: {e}")
        return None


def _norm_firma_nomi(s: str) -> str:
    """
    Firma nomini solishtirish uchun "normallashtiradi": bosh/oxiridagi
    va ortiqcha ichki bo'shliqlarni, ko'rinmas belgilarni (NBSP va h.k.)
    olib tashlaydi, katta harfga o'tkazadi. Ikkita jadvaldagi bir xil
    nom turlicha bo'shliq/formatlanish tufayli mos kelmay qolishining
    oldini oladi.
    """
    s = str(s)
    s = re.sub(r"[\s\u00a0\u200b\u200c\u200d\ufeff]+", " ", s)
    return s.strip().upper()


def get_firma_file_by_name(firma_nomi: str) -> dict | None:
    """Berilgan firma NOMI bo'yicha "Firmalar" varag'idan saqlangan faylni topadi (admin uchun)."""
    try:
        from register_handlers import _get_firmalar_ws
        ws = _get_firmalar_ws()
        records = ws.get_all_records()
        target = _norm_firma_nomi(firma_nomi)

        for row in records:
            if _norm_firma_nomi(row.get("Firma nomi", "")) == target:
                return {
                    "firma_nomi": str(row.get("Firma nomi", "")).strip(),
                    "file_id": str(row.get("FileID", "")).strip(),
                    "file_name": str(row.get("FileName", "")).strip(),
                }
        return None
    except Exception as e:
        logger.error(f"[FIRMS] get_firma_file_by_name xato: {e}")
        return None


TOLOVLAR_WS_NAME = "To'lovlar"


def _find_worksheet_flexible(sh, target_name: str):
    """
    Varaqni nomi bo'yicha topadi — apostrof belgisi turlicha yozilishi
    ('/'/ʻ/`) va katta-kichik harfga sezgir emas holda qidiradi.
    Bevosita mos kelmasa, barcha varaq nomlarini "normallashtirib"
    solishtiradi.
    """
    def _norm(s):
        s = str(s).strip().lower()
        for ch in ["'", "’", "ʻ", "`", "‘"]:
            s = s.replace(ch, "'")
        return s

    target_norm = _norm(target_name)
    for ws in sh.worksheets():
        if _norm(ws.title) == target_norm:
            return ws
    raise gspread.exceptions.WorksheetNotFound(target_name)


def get_firm_summa(firma_nomi: str) -> dict | None:
    """SALARY_SHEET_ID ichidagi "To'lovlar" varag'idan berilgan firma uchun ma'lumotlarini topadi."""
    try:
        client = _get_client()
        sh = client.open_by_key(SALARY_SHEET_ID)
        ws = _find_worksheet_flexible(sh, TOLOVLAR_WS_NAME)
        print(f"[FIRMS] '{TOLOVLAR_WS_NAME}' varag'i topildi: '{ws.title}'")
        records = ws.get_all_records()
        print(f"[FIRMS] Qidirilayotgan firma: '{firma_nomi}' | Jadvaldagi firmalar: "
              f"{[str(r.get('Firma nomi','')) for r in records]}")
        target = _norm_firma_nomi(firma_nomi)
        for row in records:
            firma = str(row.get("Firma nomi", "")).strip()
            if _norm_firma_nomi(firma) == target:
                print(f"[FIRMS] MOS TOPILDI: qator={row}")
                return {
                    "shartnoma": str(row.get("Shartnoma raqami", "")).strip(),
                    "inn": str(row.get("INN", "")).strip(),
                    "summa": row.get("Summa", ""),
                    "holati": str(row.get("Holati", "")).strip(),
                }
        print(f"[FIRMS] Mos firma topilmadi: '{firma_nomi}'")
        return None
    except gspread.exceptions.WorksheetNotFound:
        try:
            titles = [w.title for w in client.open_by_key(SALARY_SHEET_ID).worksheets()]
        except Exception:
            titles = ["(SALARY_SHEET_ID o'zi ochilmadi)"]
        logger.error(f"[FIRMS] '{TOLOVLAR_WS_NAME}' varag'i topilmadi. Mavjud varaqlar: {titles}")
        return None
    except Exception as e:
        logger.error(f"[FIRMS] get_firm_summa xato: {e}")
        return None


def get_firms_report() -> list:
    """
    "Firmalar to'lovlari" Google Sheets'idan (FIRMS_SHEET_ID) barcha
    firmalar ro'yxatini o'qiydi.

    Jadval tuzilishi (SALARY_SHEET_ID ichidagi "To'lovlar" varag'i):
        Firma nomi | Summa | Holati

    Qaytaradi: [{"firma", "summa", "holati"}, ...]
    """
    try:
        client = _get_client()
        sh = client.open_by_key(SALARY_SHEET_ID)
        ws = _find_worksheet_flexible(sh, TOLOVLAR_WS_NAME)
        records = ws.get_all_records()

        result = []
        for row in records:
            firma = str(row.get("Firma nomi", "")).strip()
            if not firma:
                continue
            result.append({
                "firma": firma,
                "shartnoma": str(row.get("Shartnoma raqami", "")).strip(),
                "inn": str(row.get("INN", "")).strip(),
                "summa": row.get("Summa", ""),
                "holati": str(row.get("Holati", "")).strip(),
            })
        return result

    except gspread.exceptions.WorksheetNotFound:
        logger.error(f"[FIRMS] '{TOLOVLAR_WS_NAME}' varag'i topilmadi (SALARY_SHEET_ID)")
        return []
    except Exception as e:
        logger.error(f"[FIRMS] get_firms_report xato: {e}")
        return []


def payments_keyboard(language: str = "uz"):
    from telegram import ReplyKeyboardMarkup
    if language == "ru":
        return ReplyKeyboardMarkup([
            ["📤 Отправить файлы фирмам"],
            ["📊 Получить мой отчёт"],
            ["⬅️ Назад"],
        ], resize_keyboard=True)
    return ReplyKeyboardMarkup([
        ["📤 Firmalarga fayl yuborish"],
        ["📊 Hisobotimni olish"],
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
    firm_send_txt = "📤 Отправить файлы фирмам" if language == "ru" else "📤 Firmalarga fayl yuborish"
    my_report_txt = "📊 Получить мой отчёт" if language == "ru" else "📊 Hisobotimni olish"

    if txt == back_txt:
        await update.message.reply_text(
            "📊 *Hisobotlar va to'lovlar*\n\nBo'limni tanlang:" if language == "uz"
            else "📊 *Отчёты и оплаты*\n\nВыберите раздел:",
            parse_mode="Markdown",
            reply_markup=reports_keyboard(language),
        )
        return REPORTS_MENU

    elif txt == firm_send_txt:
        if update.effective_user.id not in ADMIN_IDS:
            await update.message.reply_text(
                "❌ Bu bo'lim faqat administrator uchun.",
                reply_markup=payments_keyboard(language),
            )
            return PAYMENTS_MENU
        return await cmd_send_firm_files(update, ctx)

    elif txt == my_report_txt:
        if update.effective_user.id in ADMIN_IDS:
            # Admin uchun: firma nomini so'raymiz, so'ng shu firmaning
            # saqlangan faylini va to'lovini topib beramiz
            from telegram import ReplyKeyboardMarkup
            await update.message.reply_text(
                "🏢 Qaysi firma? Firma nomini kiriting:",
                reply_markup=ReplyKeyboardMarkup([[back_txt]], resize_keyboard=True),
            )
            return ADMIN_FIRM_LOOKUP
        await get_my_report_handler(update, ctx)
        return PAYMENTS_MENU

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
        user_id = update.effective_user.id
        firm_info = await run_read(get_firma_file_by_telegram_id, user_id)
        if firm_info:
            return await _send_firm_direct_report(update, ctx, firm_info)
        if user_id in ADMIN_IDS:
            return await payments_menu_enter(update, ctx)
        await update.message.reply_text(
            "❌ Siz ro'yxatdan o'tmagansiz.\n\n"
            "Agar siz firma vakili bo'lsangiz, avval \"📝 Ro'yxatdan o'tish\" → "
            "\"🏢 Firma uchun ro'yxatdan o'tish\" orqali ro'yxatdan o'ting."
            if language == "uz" else
            "❌ Вы не зарегистрированы.\n\n"
            "Если вы представитель фирмы, сначала зарегистрируйтесь через "
            "\"📝 Регистрация\" → \"🏢 Регистрация как фирма\".",
            reply_markup=reports_keyboard(language),
        )
        return REPORTS_MENU

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


# ─── Firmalarga fayl yuborish (ZIP) ────────────────────────────────────────────

async def cmd_send_firm_files(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Admin: firmalarga yuboriladigan hisobot fayllari ZIP arxivini so'raydi."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return
    await update.message.reply_text(
        "📦 *Firmalarga yuboriladigan hisobot fayllari ZIP arxivini yuboring*\n\n"
        "Fayl nomlari firma nomiga mos bo'lishi kerak:\n"
        "`ООО Фарм Импекс.xlsx`, `Дори-Дармон.xlsx` va h.k.\n\n"
        "Faqat *ro'yxatdan o'tgan* (botga TelegramID orqali ulangan) "
        "firmalarga yuboriladi.",
        parse_mode="Markdown",
    )
    return FIRM_WAIT_ZIP


async def firm_receive_zip(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """ZIP yoki bitta .xlsx faylni qabul qiladi va ro'yxatdan o'tgan firmalarga yuboradi."""
    if update.effective_user.id not in ADMIN_IDS:
        return

    if not update.message.document:
        await update.message.reply_text("❌ Iltimos, ZIP yoki .xlsx fayl yuboring.")
        return FIRM_WAIT_ZIP

    doc = update.message.document
    fname_lower = doc.file_name.lower()

    if fname_lower.endswith(".xlsx"):
        return await _firm_receive_single_xlsx(update, ctx, doc)

    if not fname_lower.endswith(".zip"):
        await update.message.reply_text("❌ Faqat ZIP yoki .xlsx fayl qabul qilinadi.")
        return FIRM_WAIT_ZIP

    from bot import main_keyboard, get_lang, MENU
    msg = await update.message.reply_text("⏳ ZIP ochilmoqda va fayllar yuborilmoqda...")

    try:
        file = await doc.get_file()
        zip_bytes = await file.download_as_bytearray()

        firma_map = await run_read(get_firma_map)
        if not firma_map:
            await msg.edit_text(
                "❌ Ro'yxatdan o'tgan firma topilmadi. Firmalar avval botda "
                "\"Ro'yxatdan o'tish\" → \"Firma uchun ro'yxatdan o'tish\" orqali ro'yxatdan o'tishi kerak."
            )
            await update.message.reply_text("📋 Asosiy menyu", reply_markup=main_keyboard(get_lang(ctx)))
            return MENU

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
                base_name = os.path.basename(fname)
                tid = find_telegram_id(base_name, firma_map)

                if not tid:
                    not_found += 1
                    not_found_list.append(base_name)
                    logger.warning(f"[FIRMS] Topilmadi: {base_name}")
                    continue

                try:
                    file_data = zf.read(fname)
                    file_io = io.BytesIO(file_data)
                    file_io.name = base_name

                    sent_msg = await ctx.bot.send_document(
                        chat_id=int(tid),
                        document=file_io,
                        filename=base_name,
                        caption=f"📊 Hisobot\n📁 {base_name}",
                    )
                    sent += 1
                    logger.info(f"[FIRMS] Yuborildi: {base_name} → {tid}")

                    # Kelajakda firma o'zi so'rasa qayta yuborish uchun
                    # file_id'ni saqlab qo'yamiz
                    try:
                        matched_firma = next(
                            (k for k, v in firma_map.items() if v == tid), None
                        )
                        if matched_firma and sent_msg.document:
                            await run_write(
                                save_firma_file, matched_firma,
                                sent_msg.document.file_id, base_name,
                            )
                    except Exception as e:
                        logger.error(f"[FIRMS] file_id saqlash xato: {e}")

                except Exception as e:
                    errors += 1
                    logger.error(f"[FIRMS] Yuborishda xato {base_name}: {e}")

                if (i + 1) % 10 == 0:
                    await msg.edit_text(f"⏳ Yuborilmoqda... {i+1}/{total}\n✅ Yuborildi: {sent} ta")

        lines = [
            f"✅ *Yuborish tugadi!*\n",
            f"📤 Yuborildi: *{sent}* ta",
            f"❌ Xato: *{errors}* ta",
            f"🔍 Topilmadi (ro'yxatdan o'tmagan): *{not_found}* ta",
        ]
        if not_found_list:
            lines.append(f"\n*Topilmagan/ro'yxatdan o'tmagan firmalar:*")
            for name in not_found_list[:20]:
                lines.append(f"  • {name}")
            if len(not_found_list) > 20:
                lines.append(f"  ... va yana {len(not_found_list)-20} ta")

        await msg.edit_text("\n".join(lines), parse_mode="Markdown")

    except zipfile.BadZipFile:
        await msg.edit_text("❌ ZIP fayl buzilgan yoki noto'g'ri format.")
    except Exception as e:
        await msg.edit_text(f"❌ Xato: {e}")
        logger.error(f"[FIRMS] Umumiy xato: {e}")

    await update.message.reply_text("📋 Asosiy menyu", reply_markup=main_keyboard(get_lang(ctx)))
    return MENU


async def _firm_receive_single_xlsx(update: Update, ctx: ContextTypes.DEFAULT_TYPE, doc):
    """
    Bitta .xlsx fayl to'g'ridan-to'g'ri yuborilganda (ZIP'siz) — fayl
    nomidan firma nomini topib, o'sha firmaga darhol yuboradi va
    kelajakda o'zi olishi uchun file_id'ni saqlaydi.
    """
    from bot import main_keyboard, get_lang, MENU
    msg = await update.message.reply_text("⏳ Fayl qayta ishlanmoqda...")

    try:
        base_name = doc.file_name
        firma_map = await run_read(get_firma_map)
        tid = find_telegram_id(base_name, firma_map) if firma_map else None

        if not tid:
            await msg.edit_text(
                f"❌ *{base_name}* nomiga mos, ro'yxatdan o'tgan firma topilmadi.\n\n"
                "Fayl nomi firma nomiga mos bo'lishi va firma avval botda "
                "ro'yxatdan o'tgan bo'lishi kerak.",
                parse_mode="Markdown",
            )
            await update.message.reply_text("📋 Asosiy menyu", reply_markup=main_keyboard(get_lang(ctx)))
            return MENU

        file = await doc.get_file()
        file_bytes = await file.download_as_bytearray()
        file_io = io.BytesIO(file_bytes)
        file_io.name = base_name

        sent_msg = await ctx.bot.send_document(
            chat_id=int(tid),
            document=file_io,
            filename=base_name,
            caption=f"📊 Hisobot\n📁 {base_name}",
        )

        matched_firma = next((k for k, v in firma_map.items() if v == tid), None)
        if matched_firma and sent_msg.document:
            await run_write(save_firma_file, matched_firma, sent_msg.document.file_id, base_name)

        await msg.edit_text(f"✅ *{matched_firma or base_name}* ga yuborildi va saqlandi!", parse_mode="Markdown")

    except Exception as e:
        logger.error(f"[FIRMS] Bitta fayl yuborish xato: {e}")
        await msg.edit_text(f"❌ Xato: {e}")

    await update.message.reply_text("📋 Asosiy menyu", reply_markup=main_keyboard(get_lang(ctx)))
    return MENU


async def firm_zip_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Bekor qilish."""
    from bot import main_keyboard, get_lang, MENU
    await update.message.reply_text(
        "❌ Bekor qilindi.",
        reply_markup=main_keyboard(get_lang(ctx))
    )
    return MENU


# ─── Firma o'zi hisobotini olishi (self-service) ───────────────────────────────

async def _send_firm_direct_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE, firm_info: dict):
    """
    Firma vakili "Отчёт va to'lovlar" bosganda — to'g'ridan-to'g'ri
    (submenyusiz) faylini va shu oylik to'lov summasini yuboradi.
    """
    language = ctx.user_data.get("lang", "uz")
    firma_nomi = firm_info.get("firma_nomi", "")

    # Fayl (agar yuklangan bo'lsa)
    if firm_info.get("file_id"):
        try:
            await ctx.bot.send_document(
                chat_id=update.effective_chat.id,
                document=firm_info["file_id"],
                filename=firm_info.get("file_name") or "hisobot.xlsx",
            )
        except Exception as e:
            logger.error(f"[FIRMS] Fayl yuborish xato: {e}")
            await update.message.reply_text("❌ Faylni yuborishda xatolik yuz berdi.")
    else:
        await update.message.reply_text(
            "📭 Hisobot fayli hali yuklanmagan." if language == "uz"
            else "📭 Файл отчёта ещё не загружен."
        )

    # Shu oylik to'lov summasi
    summa_info = await run_read(get_firm_summa, firma_nomi)
    if summa_info:
        holati = summa_info.get("holati", "").strip().lower()
        if holati in ("to'langan", "оплачено", "✅"):
            belgi = "✅"
        elif holati in ("to'lanmagan", "не оплачено", "❌"):
            belgi = "❌"
        else:
            belgi = "⏳"

        lines = [f"{belgi} *{firma_nomi}*"]
        if summa_info.get("shartnoma"):
            lines.append(f"📄 Shartnoma: {summa_info['shartnoma']}")
        if summa_info.get("inn"):
            lines.append(f"🆔 INN: {summa_info['inn']}")
        lines.append(
            f"💰 To'lov: {summa_info.get('summa', '')} ({summa_info.get('holati', '')})"
            if language == "uz" else
            f"💰 Оплата: {summa_info.get('summa', '')} ({summa_info.get('holati', '')})"
        )
        await update.message.reply_text("\n".join(lines), parse_mode="Markdown")
    else:
        await update.message.reply_text(
            f"ℹ️ *{firma_nomi}* uchun joriy oy to'lov ma'lumoti hali kiritilmagan."
            if language == "uz" else
            f"ℹ️ Данные об оплате для *{firma_nomi}* ещё не внесены.",
            parse_mode="Markdown",
        )

    return REPORTS_MENU


async def admin_firm_lookup_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Admin firma nomini kiritadi — bot shu firmaning faylini va to'lovini topib beradi."""
    language = ctx.user_data.get("lang", "uz")
    back_txt = "⬅️ Назад" if language == "ru" else "⬅️ Orqaga"
    txt = update.message.text.strip() if update.message.text else ""

    if txt == back_txt:
        await update.message.reply_text(
            "📊 *Отчёт va to'lovlar*\n\nBo'limni tanlang:" if language == "uz"
            else "📊 *Отчёт и оплаты*\n\nВыберите раздел:",
            parse_mode="Markdown",
            reply_markup=payments_keyboard(language),
        )
        return PAYMENTS_MENU

    info = await run_read(get_firma_file_by_name, txt)
    if not info:
        await update.message.reply_text(
            f"❌ *{txt}* nomli firma \"Firmalar\" ro'yxatida topilmadi.\n\nQayta kiriting:",
            parse_mode="Markdown",
        )
        return ADMIN_FIRM_LOOKUP

    await _send_firm_direct_report(update, ctx, info)
    return PAYMENTS_MENU


async def get_my_report_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    Firma vakili "📊 Hisobotimni olish" tugmasini bosganda — uning
    TelegramID'si bo'yicha faylni VA shu oylik to'lov summasini yuboradi
    (xuddi "Отчёт va to'lovlar" to'g'ridan-to'g'ri bosilganidek).
    """
    language = ctx.user_data.get("lang", "uz")
    user_id = update.effective_user.id

    info = await run_read(get_firma_file_by_telegram_id, user_id)

    if not info:
        await update.message.reply_text(
            "❌ Siz firma sifatida ro'yxatdan o'tmagansiz.\n\n"
            "Avval \"Ro'yxatdan o'tish\" → \"Firma uchun ro'yxatdan o'tish\" "
            "orqali ro'yxatdan o'ting."
            if language == "uz" else
            "❌ Вы не зарегистрированы как фирма."
        )
        return

    await _send_firm_direct_report(update, ctx, info)


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
        FIRM_WAIT_ZIP: [
            MessageHandler(filters.Document.ALL, firm_receive_zip),
            MessageHandler(filters.TEXT & ~filters.COMMAND, firm_zip_cancel),
        ],
        ADMIN_FIRM_LOOKUP: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, admin_firm_lookup_handler),
        ],
    }
