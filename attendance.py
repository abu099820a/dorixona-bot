"""
attendance.py — Davomiylik moduli (HR uslubi)
Har oy yangi lист, har sana uchun 2 ustun (Keldi / Ketdi)
Sana birlashtirilgan (merged) katakda
"""

import math
import re
import os
import json
from datetime import datetime, date, timezone, timedelta
UZ_TZ = timezone(timedelta(hours=5))
import gspread
from google.oauth2.service_account import Credentials

# ─── Sozlamalar ──────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

ATTENDANCE_SHEET_ID = os.getenv("ATTENDANCE_SHEET_ID", "BU_YERGA_DAVOMАТ_SHEET_ID")
PHARMACY_SHEET_ID   = os.getenv("PHARMACY_SHEET_ID",   "BU_YERGA_FARMATSEVTLAR_SHEET_ID")

MAX_DISTANCE_KM = 0.1

OY_NOMLARI = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
    5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
}

# Ranglar
COLOR_GREEN  = {"red": 0.7,  "green": 0.93, "blue": 0.7}   # keldi+ketdi
COLOR_ORANGE = {"red": 1.0,  "green": 0.8,  "blue": 0.4}   # faqat keldi
COLOR_YELLOW = {"red": 1.0,  "green": 0.95, "blue": 0.0}   # zamena
COLOR_RED    = {"red": 0.95, "green": 0.6,  "blue": 0.6}   # kelmagan
COLOR_HEADER = {"red": 0.27, "green": 0.51, "blue": 0.71}  # sarlavha (ko'k)
COLOR_DATE   = {"red": 0.18, "green": 0.33, "blue": 0.55}  # sana satri

# ─── ConversationHandler holatlari ───────────────────────────────────────────

(
    ATT_PHONE,
    ATT_MENU,
    ATT_FILIAL_SELECT,
    ATT_LOCATION,
    ATT_ZAMENA_FILIAL,
    ATT_ZAMENA_LOCATION,
) = range(100, 106)

# ─── Yordamchi funksiyalar ───────────────────────────────────────────────────

def haversine_m(lat1, lon1, lat2, lon2):
    R = 6371000
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat/2)**2
         + math.cos(math.radians(lat1))
         * math.cos(math.radians(lat2))
         * math.sin(dlon/2)**2)
    return R * 2 * math.asin(math.sqrt(a))


def normalize_phone(phone: str) -> str:
    digits = re.sub(r"\D", "", str(phone))
    if digits.startswith("998"):   return "+" + digits
    if digits.startswith("0"):     return "+998" + digits[1:]
    if len(digits) == 9:           return "+998" + digits
    return "+" + digits


def get_sheets_client():
    creds_json = os.getenv("GOOGLE_CREDENTIALS")
    if creds_json:
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    return gspread.authorize(creds)


def col_letter(n):
    """1 → A, 2 → B, 27 → AA ..."""
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def date_to_col(day: int) -> int:
    """
    1-kun → 3-ustun (C), chunki A=Ismi, B=Filial
    Har kun 2 ustun: keldi va ketdi
    1-kun keldi → ustun 3, ketdi → 4
    2-kun keldi → 5, ketdi → 6
    ...
    """
    return 3 + (day - 1) * 2


# ─── Sheet tuzilishi ─────────────────────────────────────────────────────────

def _get_or_create_month_sheet(sh):
    """
    Joriy oy uchun list topadi yoki yaratadi.
    Qator 1: Ismi | Filial | 01.06 (merged 2 ustun) | 02.06 | ...
    Qator 2:       |        | Keldi | Ketdi | Keldi | Ketdi | ...
    Qator 3+: farmatsevtlar
    """
    import calendar
    now = datetime.now(UZ_TZ)
    sheet_name = f"{OY_NOMLARI[now.month]} {now.year}"
    existing = [ws.title for ws in sh.worksheets()]

    if sheet_name in existing:
        return sh.worksheet(sheet_name)

    days_in_month = calendar.monthrange(now.year, now.month)[1]
    total_cols = 2 + days_in_month * 2

    ws = sh.add_worksheet(title=sheet_name, rows=400, cols=total_cols)

    # 1-qator: Ismi, Filial, sanalar
    row1 = ["Ismi", "Filial"]
    for d in range(1, days_in_month + 1):
        row1.append(f"{d:02d}.{now.month:02d}")
        row1.append("")
    ws.update("A1", [row1])

    # 2-qator: Keldi/Ketdi
    row2 = ["", ""]
    for _ in range(days_in_month):
        row2.extend(["Keldi", "Ketdi"])
    ws.update("A2", [row2])

    # Merge so'rovlari
    requests = []

    # A1:A2 merge (Ismi)
    requests.append({"mergeCells": {"range": {
        "sheetId": ws.id,
        "startRowIndex": 0, "endRowIndex": 2,
        "startColumnIndex": 0, "endColumnIndex": 1
    }, "mergeType": "MERGE_ALL"}})

    # B1:B2 merge (Filial)
    requests.append({"mergeCells": {"range": {
        "sheetId": ws.id,
        "startRowIndex": 0, "endRowIndex": 2,
        "startColumnIndex": 1, "endColumnIndex": 2
    }, "mergeType": "MERGE_ALL"}})

    # Har kun uchun merge
    for d in range(days_in_month):
        col_start = 2 + d * 2
        requests.append({"mergeCells": {"range": {
            "sheetId": ws.id,
            "startRowIndex": 0, "endRowIndex": 1,
            "startColumnIndex": col_start, "endColumnIndex": col_start + 2
        }, "mergeType": "MERGE_ALL"}})

    sh.batch_update({"requests": requests})

    # Format
    last_col = col_letter(total_cols)
    ws.format("A1:B2", {
        "backgroundColor": COLOR_HEADER,
        "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE",
    })
    ws.format(f"C1:{last_col}1", {
        "backgroundColor": COLOR_DATE,
        "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}},
        "horizontalAlignment": "CENTER",
    })
    ws.format(f"C2:{last_col}2", {
        "backgroundColor": COLOR_HEADER,
        "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}},
        "horizontalAlignment": "CENTER",
    })

    # Ustun kengligi
    sh.batch_update({"requests": [
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 180}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2},
            "properties": {"pixelSize": 130}, "fields": "pixelSize"
        }},
    ]})

    return ws

def _get_farmatsevt_row(ws, ismi: str) -> int:
    """
    Farmatsevtning qator raqamini topadi.
    Topilmasa — yangi qator qo'shadi va raqamini qaytaradi.
    """
    all_values = ws.get_all_values()
    for i, row in enumerate(all_values):
        if i < 2:
            continue   # sarlavha qatorlari
        if row and row[0] == ismi:
            return i + 1   # 1-indexed

    # Yangi qator qo'shish
    next_row = len(all_values) + 1
    return next_row


# ─── Asosiy funksiyalar ───────────────────────────────────────────────────────

def get_farmatsevt(phone: str):
    try:
        client = get_sheets_client()
        sh = client.open_by_key(PHARMACY_SHEET_ID)
        ws = sh.sheet1
        records = ws.get_all_records()
        norm = normalize_phone(phone)
        for row in records:
            tel_raw = row.get("Telefon", "")
            if isinstance(tel_raw, float):
                tel_raw = str(int(tel_raw))
            else:
                tel_raw = str(tel_raw)
            if normalize_phone(tel_raw) == norm:
                return {
                    "ismi":   str(row.get("Ismi", "")).strip(),
                    "filial": str(row.get("Filial", "")).strip(),
                    "lat":    float(str(row.get("Lat", 0)).replace(",", ".")),
                    "lon":    float(str(row.get("Lon", 0)).replace(",", ".")),
                }
        return None
    except Exception as e:
        print(f"[ATT] Farmatsevt xato: {e}")
        return None


def write_attendance(farmatsevt: dict, action: str, zamena: bool = False):
    """
    Joriy oy listiga, farmatsevt qatoriga, bugungi ustunga vaqt yozadi.
    action: 'keldi' | 'ketdi'
    """
    try:
        client = get_sheets_client()
        sh = client.open_by_key(ATTENDANCE_SHEET_ID)
        ws = _get_or_create_month_sheet(sh)

        now = datetime.now(UZ_TZ)
        day = now.day
        time_str = now.strftime("%H:%M")

        # Ustun raqami
        if action == "keldi":
            col_num = date_to_col(day)        # juft: keldi
        else:
            col_num = date_to_col(day) + 1    # toq: ketdi

        col_ltr = col_letter(col_num)

        # Farmatsevt qatorini topish
        row_num = _get_farmatsevt_row(ws, farmatsevt["ismi"])

        # Agar yangi qator bo'lsa — ismi va filialini yozish
        existing = ws.cell(row_num, 1).value
        if not existing:
            ws.update_cell(row_num, 1, farmatsevt["ismi"])
            ws.update_cell(row_num, 2, farmatsevt["filial"])

        # Vaqtni yozish
        ws.update_cell(row_num, col_num, time_str)

        # Rang belgilash
        cell_range = f"{col_ltr}{row_num}"
        if zamena:
            color = COLOR_YELLOW
        else:
            # Keldi va ketdi ikkalasi to'lganmi?
            keldi_col = col_letter(date_to_col(day))
            ketdi_col = col_letter(date_to_col(day) + 1)
            keldi_val = ws.acell(f"{keldi_col}{row_num}").value
            ketdi_val = ws.acell(f"{ketdi_col}{row_num}").value

            if keldi_val and ketdi_val:
                # Ikkalasi to'liq — ikkalasini ham yashil qilish
                ws.format(f"{keldi_col}{row_num}:{ketdi_col}{row_num}",
                          {"backgroundColor": COLOR_GREEN})
                return True
            else:
                color = COLOR_ORANGE   # Faqat keldi

        ws.format(cell_range, {"backgroundColor": color})
        return True

    except Exception as e:
        print(f"[ATT] Yozish xato: {e}")
        return False


def mark_absent_today():
    """
    Bugun kelmagan farmatsevtlarni qizil rang bilan belgilaydi.
    Bu funksiya har kun kechqurun (21:00) ishga tushirilishi kerak.
    """
    try:
        client = get_sheets_client()
        sh = client.open_by_key(ATTENDANCE_SHEET_ID)
        ws = _get_or_create_month_sheet(sh)

        now = datetime.now(UZ_TZ)
        day = now.day
        keldi_col_num = date_to_col(day)
        ketdi_col_num = date_to_col(day) + 1
        keldi_col = col_letter(keldi_col_num)
        ketdi_col = col_letter(ketdi_col_num)

        all_values = ws.get_all_values()
        for i, row in enumerate(all_values):
            if i < 2: continue
            if not row or not row[0]: continue

            keldi_val = row[keldi_col_num - 1] if len(row) >= keldi_col_num else ""
            ketdi_val = row[ketdi_col_num - 1] if len(row) >= ketdi_col_num else ""

            if not keldi_val and not ketdi_val:
                row_num = i + 1
                ws.format(f"{keldi_col}{row_num}:{ketdi_col}{row_num}",
                          {"backgroundColor": COLOR_RED})
    except Exception as e:
        print(f"[ATT] Kelmagan belgilash xato: {e}")




def get_farmatsevt_by_userid(user_id: int):
    """Telegram user_id bo'yicha farmatsevtni topadi"""
    try:
        client = get_sheets_client()
        sh = client.open_by_key(PHARMACY_SHEET_ID)
        ws = sh.sheet1
        records = ws.get_all_records()
        uid = str(user_id)
        for i, row in enumerate(records):
            if str(row.get("TelegramID", "")).strip() == uid:
                return {
                    "ismi":   str(row.get("Ismi", "")).strip(),
                    "filial": str(row.get("Filial", "")).strip(),
                    "lat":    float(str(row.get("Lat", 0)).replace(",", ".")),
                    "lon":    float(str(row.get("Lon", 0)).replace(",", ".")),
                }
        return None
    except Exception as e:
        print(f"[ATT] UserID qidirish xato: {e}")
        return None


def save_userid_to_sheet(user_id: int, phone: str):
    """Farmatsevtning TelegramID sini saqlaydi"""
    try:
        client = get_sheets_client()
        sh = client.open_by_key(PHARMACY_SHEET_ID)
        ws = sh.sheet1
        records = ws.get_all_records()
        norm = normalize_phone(phone)

        for i, row in enumerate(records):
            tel_raw = row.get("Telefon", "")
            if isinstance(tel_raw, float):
                tel_raw = str(int(tel_raw))
            else:
                tel_raw = str(tel_raw)
            if normalize_phone(tel_raw) == norm:
                row_num = i + 2  # 1-indexed + sarlavha
                # TelegramID ustuni F (6-ustun) bo'lsin
                ws.update_cell(row_num, 6, str(user_id))
                return True
        return False
    except Exception as e:
        print(f"[ATT] UserID saqlash xato: {e}")
        return False

def get_filiallar_list():
    try:
        client = get_sheets_client()
        sh = client.open_by_key(PHARMACY_SHEET_ID)
        ws = sh.sheet1
        records = ws.get_all_records()
        seen = {}
        for row in records:
            f = str(row.get("Filial", "")).strip()
            if f and f not in seen:
                try:
                    seen[f] = {
                        "filial": f,
                        "lat": float(str(row.get("Lat", 0)).replace(",", ".")),
                        "lon": float(str(row.get("Lon", 0)).replace(",", ".")),
                    }
                except Exception:
                    pass
        return sorted(seen.values(), key=lambda x: int(x["filial"]) if x["filial"].isdigit() else 9999)
    except Exception as e:
        print(f"[ATT] Filiallar xato: {e}")
        return []
