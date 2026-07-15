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
    return 4 + (day - 1) * 2  # A=Filial, B=Ismi, C=Telefon, D=01 Keldi...


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

    ws = sh.add_worksheet(title=sheet_name, rows=400, cols=total_cols + 2)  # +2: Jami soat va zaxira

    # 1-qator: Ismi, Filial, sanalar
    row1 = ["Filial", "Ismi"]
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

    # A1:A2 merge (Filial)
    requests.append({"mergeCells": {"range": {
        "sheetId": ws.id,
        "startRowIndex": 0, "endRowIndex": 2,
        "startColumnIndex": 0, "endColumnIndex": 1
    }, "mergeType": "MERGE_ALL"}})

    # B1:B2 merge (Ismi)
    requests.append({"mergeCells": {"range": {
        "sheetId": ws.id,
        "startRowIndex": 0, "endRowIndex": 2,
        "startColumnIndex": 1, "endColumnIndex": 2
    }, "mergeType": "MERGE_ALL"}})



    # Har kun uchun merge
    for d in range(days_in_month):
        col_start = 3 + d * 2  # A=Filial, B=Ismi, C=Telefon dan keyin
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
            "properties": {"pixelSize": 160}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2},
            "properties": {"pixelSize": 140}, "fields": "pixelSize"
        }},

    ]})

    return ws

def _get_farmatsevt_row(ws, ismi: str, filial: str = "") -> int:
    """
    Farmatsevtning qator raqamini topadi.
    Jadval: A=Filial, B=Ismi
    Topilmasa — -1 qaytaradi (yangi qator qo'shilmaydi, init_month kerak)
    """
    all_values = ws.get_all_values()
    for i, row in enumerate(all_values):
        if i < 2:
            continue   # sarlavha qatorlari
        # B ustun (index 1) = Ismi
        if len(row) > 1 and row[1].strip() == ismi.strip():
            return i + 1   # 1-indexed
    return -1


# ─── Asosiy funksiyalar ───────────────────────────────────────────────────────

def get_farmatsevt(phone: str):
    try:
        client = get_sheets_client()
        sh = client.open_by_key(PHARMACY_SHEET_ID)
        ws = sh.worksheet("Farmatsevtlar")
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


def write_attendance(farmatsevt: dict, action: str, zamena: bool = False, write_time=None):
    """
    Joriy oy listiga, farmatsevt qatoriga, bugungi ustunga vaqt yozadi.
    action: 'keldi' | 'ketdi'
    """
    try:
        client = get_sheets_client()
        sh = client.open_by_key(ATTENDANCE_SHEET_ID)
        ws = _get_or_create_month_sheet(sh)

        now = datetime.now(UZ_TZ)
        if write_time is not None:
            day = write_time.day
        else:
            day = now.day
        time_str = now.strftime("%H:%M")

        # Ustun raqami
        if action == "keldi":
            col_num = date_to_col(day)        # juft: keldi
        else:
            col_num = date_to_col(day) + 1    # toq: ketdi

        col_ltr = col_letter(col_num)

        # Farmatsevt qatorini topish (B ustun = Ismi)
        row_num = _get_farmatsevt_row(ws, farmatsevt["ismi"], farmatsevt.get("filial", ""))

        if row_num == -1:
            # Topilmadi — jadval to'liq emas, oxiriga qo'shamiz
            all_values = ws.get_all_values()
            row_num = len(all_values) + 1
            # A=Filial, B=Ismi
            ws.update_cell(row_num, 1, farmatsevt.get("filial", ""))
            ws.update_cell(row_num, 2, farmatsevt["ismi"])

        # Vaqtni yozish
        if zamena and farmatsevt.get("zamena_filial"):
            # Zamena: "09:00 (6)" formatida
            fil = farmatsevt["zamena_filial"]
            import re
            fil_no = re.match(r"^(\d+)", str(fil).strip())
            fil_label = fil_no.group(1) if fil_no else fil
            write_val = f"{time_str} ({fil_label})"
        else:
            write_val = time_str
        ws.update_cell(row_num, col_num, write_val)

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
        ws = sh.worksheet("Farmatsevtlar")
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
        ws = sh.worksheet("Farmatsevtlar")
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



def init_month_sheet(sh=None):
    """
    Oy boshida chaqiriladi.
    1. Yangi oy listi yaratadi (yoki mavjudini tozalab qayta to'ldiradi)
    2. A=Filial, B=Ismi tartibida yozadi
    3. Har filial uchun sarlavha qatori (faqat filial nomi, ko'k rang)
    4. Xodimi yo'q filiallar ham sarlavha qatori sifatida yoziladi
    """
    import calendar
    now = datetime.now(UZ_TZ)

    if sh is None:
        client = get_sheets_client()
        sh = client.open_by_key(ATTENDANCE_SHEET_ID)

    ws = _get_or_create_month_sheet(sh)

    # Farmatsevtlar Sheets dan ma'lumot olish
    ph_client = get_sheets_client()
    ph_ws = ph_client.open_by_key(PHARMACY_SHEET_ID).worksheet("Farmatsevtlar")
    all_ph = ph_ws.get_all_values()

    # Barcha filiallar va ularning xodimlari — tartibli lug'at
    # {filial_nomi: [ismi1, ismi2, ...]}
    from collections import OrderedDict
    filial_dict = OrderedDict()

    for i, row in enumerate(all_ph):
        if i == 0:
            continue  # sarlavha
        if not row or not row[0]:
            continue
        filial = str(row[0]).strip()
        ismi   = str(row[1]).strip() if len(row) > 1 else ""
        tel    = str(row[2]).strip() if len(row) > 2 else ""
        if isinstance(row[2] if len(row) > 2 else "", float):
            tel = str(int(float(tel))) if tel else ""
        if filial not in filial_dict:
            filial_dict[filial] = []
        if ismi:
            filial_dict[filial].append((ismi, tel))

    # Qatorlarni tuzish: filial sarlavhasi + xodimlar
    rows_to_write = []   # [(filial, ismi), ...]
    for filial, xodimlar in filial_dict.items():
        rows_to_write.append((filial, "", ""))  # filial sarlavha qatori
        for ismi, tel in xodimlar:
            rows_to_write.append((filial, ismi, tel))

    # Jadvalga yozish (3-qatordan boshlanadi)
    updates = []
    filial_header_rows = []  # rang berish uchun

    for idx, (filial, ismi, tel) in enumerate(rows_to_write):
        row_num = idx + 3  # 1=sarlavha, 2=Keldi/Ketdi
        updates.append({
            "range": f"A{row_num}:C{row_num}",
            "values": [[filial, ismi, tel]]
        })
        if ismi == "":
            filial_header_rows.append(row_num)

    if updates:
        ws.batch_update(updates)

    # Filial sarlavha qatorlariga rang berish (to'q ko'k)
    batch_requests = []
    for row_num in filial_header_rows:
        batch_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": row_num - 1,
                    "endRowIndex": row_num,
                    "startColumnIndex": 0,
                    "endColumnIndex": 2,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": COLOR_HEADER,
                        "textFormat": {
                            "bold": True,
                            "foregroundColor": {"red": 1, "green": 1, "blue": 1}
                        },
                        "horizontalAlignment": "CENTER",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
            }
        })

    if batch_requests:
        sh.batch_update({"requests": batch_requests})

    print(f"[ATT] {len(rows_to_write)} ta qator yozildi ({len(filial_header_rows)} ta filial)")
    return ws


def calculate_monthly_hours():
    """
    Joriy oy uchun har bir farmatsevtning ish soatini hisoblaydi.
    Keldi va Ketdi vaqtlari farqidan hisoblanadi.
    Oxirgi ustunda ko'rsatiladi.
    """
    import calendar
    now = datetime.now(UZ_TZ)

    client = get_sheets_client()
    sh = client.open_by_key(ATTENDANCE_SHEET_ID)
    ws = _get_or_create_month_sheet(sh)

    days_in_month = calendar.monthrange(now.year, now.month)[1]
    total_cols = 2 + days_in_month * 2
    jami_col_num = total_cols + 1
    jami_col = col_letter(jami_col_num)

    all_values = ws.get_all_values()
    updates = []

    for i, row in enumerate(all_values):
        if i < 2: continue  # sarlavhalar
        if not row or not row[0]: continue

        total_minutes = 0
        for d in range(1, days_in_month + 1):
            keldi_idx = date_to_col(d) - 1      # 0-indexed
            ketdi_idx = date_to_col(d)           # 0-indexed

            keldi_val = row[keldi_idx] if len(row) > keldi_idx else ""
            ketdi_val = row[ketdi_idx] if len(row) > ketdi_idx else ""

            if keldi_val and ketdi_val:
                try:
                    # HH:MM formatida
                    kh, km = map(int, keldi_val.split(":"))
                    th, tm = map(int, ketdi_val.split(":"))
                    diff = (th * 60 + tm) - (kh * 60 + km)
                    if diff > 0:
                        total_minutes += diff
                except Exception:
                    pass

        if total_minutes > 0:
            soat = total_minutes / 60
            row_num = i + 1
            updates.append({
                "range": f"{jami_col}{row_num}",
                "values": [[f"{soat:.1f} soat"]]
            })

    if updates:
        ws.batch_update(updates)
        # Yashil rang
        for upd in updates:
            ws.format(upd["range"], {
                "backgroundColor": {"red": 0.7, "green": 0.93, "blue": 0.7},
                "textFormat": {"bold": True},
                "horizontalAlignment": "CENTER",
            })

    print(f"[ATT] Ish soatlari hisoblandi: {len(updates)} ta")
    return len(updates)

def get_filiallar_list():
    """
    Filiallar ro'yxatini qaytaradi.
    Avval FILIALLAR_SHEET_ID dan, bo'lmasa Farmatsevtlar Sheets dan oladi.
    """
    try:
        client = get_sheets_client()
        filiallar_sheet_id = os.getenv("FILIALLAR_SHEET_ID", "")

        if filiallar_sheet_id:
            # Filiallar Sheets dan (A=Filial №, M=Latitude, N=Longitude)
            sh = client.open_by_key(filiallar_sheet_id)
            ws = sh.sheet1
            all_values = ws.get_all_values()
            if not all_values:
                return []

            headers = [h.strip() for h in all_values[0]]
            try:
                lat_idx = headers.index("Latitude")
            except ValueError:
                lat_idx = 12
            try:
                lon_idx = headers.index("Longitude")
            except ValueError:
                lon_idx = 13

            result = []
            for row in all_values[1:]:
                if not row or not row[0]:
                    continue
                fil_no = str(row[0]).strip()
                fil_name = str(row[1]).strip() if len(row) > 1 else fil_no
                lat_val = str(row[lat_idx]).strip() if len(row) > lat_idx else ""
                lon_val = str(row[lon_idx]).strip() if len(row) > lon_idx else ""

                if not lat_val or not lon_val or lat_val in ("0", "nan", ""):
                    continue

                try:
                    result.append({
                        "filial": fil_no,
                        "filial_name": fil_name,
                        "lat": float(lat_val.replace(",", ".")),
                        "lon": float(lon_val.replace(",", ".")),
                    })
                except Exception:
                    pass

            # Raqam bo'yicha tartiblash
            result.sort(key=lambda x: int(x["filial"]) if str(x["filial"]).isdigit() else 9999)
            print(f"[ATT] Filiallar Sheets dan {len(result)} ta filial yuklandi")
            return result

        else:
            # Farmatsevtlar Sheets dan (zaxira)
            sh = client.open_by_key(PHARMACY_SHEET_ID)
            ws = sh.worksheet("Farmatsevtlar")
            records = ws.get_all_records()
            seen = {}
            for row in records:
                f = str(row.get("Filial", "")).strip()
                if f and f not in seen:
                    try:
                        lat = float(str(row.get("Lat", 0)).replace(",", "."))
                        lon = float(str(row.get("Lon", 0)).replace(",", "."))
                        if lat and lon:
                            seen[f] = {"filial": f, "filial_name": f, "lat": lat, "lon": lon}
                    except Exception:
                        pass
            return sorted(seen.values(), key=lambda x: 9999)

    except Exception as e:
        print(f"[ATT] Filiallar xato: {e}")
        return []


def generate_code(filial: str, phone: str = "") -> str:
    """Filial nomidan raqamni ajratadi: '6 - ЮНУСАБАД 7' → '6'"""
    m = re.match(r"^(\d+)", str(filial).strip())
    return m.group(1) if m else re.sub(r"\D", "", str(filial))[:3]


def fill_codes_in_sheet():
    """Farmatsevtlar Sheets ga kod yozadi."""
    try:
        client = get_sheets_client()
        sh = client.open_by_key(PHARMACY_SHEET_ID)
        ws = sh.worksheet("Farmatsevtlar")
        headers = ws.row_values(1)
        if "Kod" in headers:
            kod_col_num = headers.index("Kod") + 1
            kod_col = col_letter(kod_col_num)
        else:
            kod_col_num = len(headers) + 1
            kod_col = col_letter(kod_col_num)
            ws.update_cell(1, kod_col_num, "Kod")

        records = ws.get_all_records()
        updates = []
        codes_written = []

        for i, row in enumerate(records):
            ismi = str(row.get("Ismi", "")).strip()
            filial = str(row.get("Filial", "")).strip()
            tel_raw = row.get("Telefon", "")
            if isinstance(tel_raw, float):
                tel_raw = str(int(tel_raw))
            else:
                tel_raw = str(tel_raw)
            if not ismi or not filial:
                continue
            existing_code = str(row.get("Kod", "")).strip()
            if existing_code:
                continue
            code = generate_code(filial)
            row_num = i + 2
            updates.append({"range": f"{kod_col}{row_num}", "values": [[code]]})
            codes_written.append(f"{ismi} (#{filial}) → {code}")

        if updates:
            ws.batch_update(updates)
        return codes_written
    except Exception as e:
        print(f"[ATT] fill_codes xato: {e}")
        return []


def sync_pharmacists():
    """Farmatsevtlar ro'yxatini davomat jadvali bilan sinxronlashtiradi."""
    results = {"added": [], "updated": [], "removed": [], "unchanged": 0}
    try:
        client = get_sheets_client()

        ph_sh = client.open_by_key(PHARMACY_SHEET_ID)
        ph_ws = ph_sh.worksheet("Farmatsevtlar")
        ph_records = ph_ws.get_all_records()

        ph_dict = {}
        for row in ph_records:
            ismi = str(row.get("Ismi", "")).strip()
            filial = str(row.get("Filial", "")).strip()
            tel = str(row.get("Telefon", "")).strip()
            if isinstance(row.get("Telefon", ""), float):
                tel = str(int(float(tel))) if tel else ""
            if ismi:
                ph_dict[ismi] = {"filial": filial, "tel": tel}

        att_sh = client.open_by_key(ATTENDANCE_SHEET_ID)
        ws = _get_or_create_month_sheet(att_sh)
        all_values = ws.get_all_values()

        att_dict = {}
        for i, row in enumerate(all_values):
            if i < 2:
                continue
            if not row:
                continue
            # A=Filial, B=Ismi
            filial = str(row[0]).strip() if len(row) > 0 else ""
            ismi   = str(row[1]).strip() if len(row) > 1 else ""
            if not ismi:
                continue  # Bo'sh B ustun = filial sarlavha qatori, o'tkazib yuborish
            att_dict[ismi] = {"row_num": i + 1, "filial": filial}

        batch_requests = []

        for ismi, ph_info in ph_dict.items():
            filial = ph_info["filial"]
            tel = ph_info.get("tel", "")
            if ismi not in att_dict:
                # Shu filialdagi oxirgi qatorni topish
                filial_last_row = 2  # default: sarlavhadan keyin
                for i, row in enumerate(all_values):
                    if i < 2:
                        continue
                    if not row:
                        continue
                    # A=Filial ustun
                    row_filial = str(row[0]).strip() if len(row) > 0 else ""
                    if row_filial == filial:
                        filial_last_row = i + 1  # 1-indexed (filial sarlavha ham, xodim ham)

                # Shu filialdan keyin qo'shish
                insert_row = filial_last_row + 1
                ws.insert_row([filial, ismi, tel], index=insert_row)
                # all_values ni yangilash (keyingi iteratsiya uchun)
                all_values.insert(insert_row - 1, [filial, ismi, tel])
                results["added"].append(ismi)

        COLOR_HIDDEN = {"red": 0.85, "green": 0.85, "blue": 0.85}
        for ismi, info in att_dict.items():
            if ismi not in ph_dict:
                row_num = info["row_num"]
                batch_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": row_num - 1,
                            "endRowIndex": row_num,
                            "startColumnIndex": 0,
                            "endColumnIndex": 2,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": COLOR_HIDDEN,
                                "textFormat": {
                                    "strikethrough": True,
                                    "foregroundColor": {"red": 0.5, "green": 0.5, "blue": 0.5}
                                }
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat)",
                    }
                })
                results["removed"].append(ismi)
            else:
                new_filial = ph_dict[ismi]["filial"]
                if att_dict[ismi]["filial"] != new_filial:
                    row_num = att_dict[ismi]["row_num"]
                    ws.update_cell(row_num, 1, new_filial)
                    results["updated"].append(f"{ismi}: {att_dict[ismi]['filial']} → {new_filial}")
                else:
                    results["unchanged"] += 1

        if batch_requests:
            att_sh.batch_update({"requests": batch_requests})

    except Exception as e:
        print(f"[SYNC] Xato: {e}")
        results["error"] = str(e)

    return results


def check_today_keldi(ismi: str) -> bool:
    """
    Bugun farmatsevtning keldi vaqti jadvalda yozilganmi tekshiradi.
    Bot qayta ishga tushganda sessiya yo'qolsa — jadvaldan tekshiradi.
    """
    try:
        client = get_sheets_client()
        sh = client.open_by_key(ATTENDANCE_SHEET_ID)
        ws = _get_or_create_month_sheet(sh)

        now = datetime.now(UZ_TZ)
        day = now.day
        keldi_col_num = date_to_col(day)

        all_values = ws.get_all_values()
        for i, row in enumerate(all_values):
            if i < 2:
                continue
            if not row or len(row) < 2:
                continue
            # B ustun = Ismi
            row_ismi = str(row[1]).strip()
            if row_ismi == ismi.strip():
                keldi_val = row[keldi_col_num - 1] if len(row) >= keldi_col_num else ""
                return bool(keldi_val and keldi_val.strip())
        return False
    except Exception as e:
        print(f"[ATT] check_today_keldi xato: {e}")
        return True  # Xato bo'lsa — ruxsat berish (bloklamaslik)


def check_today_ketdi(ismi: str, now=None) -> bool:
    """
    Bugun farmatsevtning ketdi vaqti jadvalda yozilganmi tekshiradi.
    """
    try:
        if now is None:
            now = datetime.now(UZ_TZ)
        client = get_sheets_client()
        sh = client.open_by_key(ATTENDANCE_SHEET_ID)
        ws = _get_or_create_month_sheet(sh)

        day = now.day
        ketdi_col_num = date_to_col(day) + 1  # Ketdi = Keldi + 1

        all_values = ws.get_all_values()
        for i, row in enumerate(all_values):
            if i < 2:
                continue
            if not row or len(row) < 2:
                continue
            row_ismi = str(row[1]).strip()
            if row_ismi == ismi.strip():
                ketdi_val = row[ketdi_col_num - 1] if len(row) >= ketdi_col_num else ""
                return bool(ketdi_val and ketdi_val.strip())
        return False
    except Exception as e:
        print(f"[ATT] check_today_ketdi xato: {e}")
        return False


def get_today_times(ismi: str, now=None) -> dict:
    """
    Bugungi keldi va ketdi vaqtlarini qaytaradi.
    {"keldi": "09:15", "ketdi": "18:00"} yoki bo'sh string
    """
    try:
        if now is None:
            now = datetime.now(UZ_TZ)
        client = get_sheets_client()
        sh = client.open_by_key(ATTENDANCE_SHEET_ID)
        ws = _get_or_create_month_sheet(sh)

        day = now.day
        keldi_col = date_to_col(day)
        ketdi_col = date_to_col(day) + 1

        all_values = ws.get_all_values()
        for i, row in enumerate(all_values):
            if i < 2:
                continue
            if not row or len(row) < 2:
                continue
            if str(row[1]).strip() == ismi.strip():
                keldi = row[keldi_col - 1] if len(row) >= keldi_col else ""
                ketdi = row[ketdi_col - 1] if len(row) >= ketdi_col else ""
                return {"keldi": str(keldi).strip(), "ketdi": str(ketdi).strip()}
        return {"keldi": "", "ketdi": ""}
    except Exception as e:
        print(f"[ATT] get_today_times xato: {e}")
        return {"keldi": "", "ketdi": ""}


def update_farmatsevt_filial_lavozim(
    ismi: str, old_filial: str, new_filial: str,
    new_lavozim: str, lat: float, lon: float, telegram_id: int
) -> bool:
    """
    Farmatsevtning filial va lavozimini yangilaydi:
    1. Eski qatorni o'chiradi
    2. Yangi filialning oxirgi xodimidan KEYIN qo'shadi
    3. Davomat jadvalida B ustunidagi filial nomini yangilaydi
    """
    try:
        client = get_sheets_client()
        sh = client.open_by_key(PHARMACY_SHEET_ID)
        ws = sh.worksheet("Farmatsevtlar")
        all_values = ws.get_all_values()

        # Eski qatorni topish (B ustun = Ismi)
        old_row_num = None
        tel = ""
        for i, row in enumerate(all_values):
            if i == 0:
                continue
            if len(row) > 1 and str(row[1]).strip() == ismi.strip():
                old_row_num = i + 1  # 1-indexed
                tel = str(row[2]).strip() if len(row) > 2 else ""
                break

        if not old_row_num:
            print(f"[CHG] Farmatsevt topilmadi: {ismi}")
            return False

        # Eski qatorni o'chirish
        ws.delete_rows(old_row_num)
        print(f"[CHG] Eski qator o'chirildi: {ismi} | {old_filial} | qator {old_row_num}")

        # Yangi filialning oxirgi qatorini topish
        all_values = ws.get_all_values()
        last_row = len(all_values)  # default: eng oxiri
        found_filial = False

        for i, row in enumerate(all_values):
            if i == 0:
                continue
            if not row or not row[0]:
                continue
            row_filial = str(row[0]).strip()
            # Filial raqamini solishtirish
            m1 = re.match(r"^(\d+)", row_filial)
            m2 = re.match(r"^(\d+)", new_filial)
            if m1 and m2 and m1.group(1) == m2.group(1):
                last_row = i + 1  # 1-indexed
                found_filial = True

        insert_row = last_row + 1
        ws.insert_row(
            [new_filial, ismi, tel, str(telegram_id), new_lavozim, str(lat), str(lon)],
            index=insert_row,
            value_input_option="USER_ENTERED"
        )
        print(f"[CHG] Yangi qator qo'shildi: {ismi} | {new_filial} | qator {insert_row}")

        # Davomat jadvalida xodimning qatori topilib, ESKI joydan olib
        # tashlanadi va YANGI filial guruhi ostiga ko'chiriladi (nafaqat A
        # ustunidagi matn almashtiriladi — aks holda xodim eski filial
        # bo'limida qolib, faqat nomi o'zgarib ko'rinadi).
        try:
            att_sh = client.open_by_key(ATTENDANCE_SHEET_ID)
            ws_att = _get_or_create_month_sheet(att_sh)
            att_values = ws_att.get_all_values()

            old_att_row_num = None
            row_data = None
            for i, row in enumerate(att_values):
                if i < 2:
                    continue
                if not row or len(row) < 2:
                    continue
                if str(row[1]).strip() == ismi.strip():
                    old_att_row_num = i + 1  # 1-indexed
                    row_data = list(row)
                    break

            if old_att_row_num and row_data is not None:
                # Filial ustunini (A) yangilash
                row_data[0] = new_filial

                # Eski qatorni o'chirish
                ws_att.delete_rows(old_att_row_num)
                print(f"[CHG] Davomat: eski qator o'chirildi | {ismi} | qator {old_att_row_num}")

                # Yangi filial guruhining oxirgi qatorini topish (raqam bo'yicha)
                att_values = ws_att.get_all_values()
                new_last_row = len(att_values)
                m_new = re.match(r"^(\d+)", new_filial)

                for i, row in enumerate(att_values):
                    if i < 2:
                        continue
                    if not row or not row[0]:
                        continue
                    m_row = re.match(r"^(\d+)", str(row[0]).strip())
                    if m_row and m_new and m_row.group(1) == m_new.group(1):
                        new_last_row = i + 1

                insert_row = new_last_row + 1
                ws_att.insert_row(row_data, index=insert_row, value_input_option="USER_ENTERED")
                print(f"[CHG] Davomat: yangi qator qo'shildi | {ismi} | {new_filial} | qator {insert_row}")
            else:
                print(f"[CHG] Davomat: {ismi} topilmadi")
        except Exception as e:
            print(f"[CHG] Davomat yangilash xato: {e}")

        return True

    except Exception as e:
        print(f"[CHG] Yangilash xato: {e}")
        return False

