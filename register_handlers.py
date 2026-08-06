"""
register_handlers.py — Umumiy Google Sheets yordamchi funksiyalar.

MUHIM (tarix uchun izoh): bu faylning ilgarigi versiyasi salary_handlers.py
bilan deyarli bir xil (adashib nusxalangan) edi va bot.py/salary_handlers.py/
attendance_handlers.py talab qilayotgan bir qancha funksiyalar (register_enter,
get_reg_states, _get_firmalar_ws, get_filial_info va h.k.) unda umuman yo'q edi
— shu sababli bot ishga tushganda ImportError bilan yiqilib qolardi.

Endi bu fayl faqat IKKI kichik, umumiy Google Sheets yordamchisini saqlaydi
(_get_firmalar_ws, get_filial_info) — bular salary_handlers.py va
attendance_handlers.py tomonidan ishlatiladi. Ro'yxatdan o'tish (registratsiya)
mantiqi to'liq registration_handlers.py da joylashgan — bot.py endi
register_enter/get_reg_states'ni to'g'ridan-to'g'ri O'SHA yerdan import qiladi.

cmd_add_filial_headers / cmd_add_filial_headers_salary /
cmd_reorder_by_lavozim_salary / cmd_sync_all_filials_salary — bular bir
martalik, eski migratsiya/admin buyruqlari bo'lib, ularning asl kodi loyihada
topilmadi. Xavfli (Google Sheets tuzilishini o'zgartiruvchi) buyruqlarni asl
koddan aniq bilmasdan qayta yozish xato qilib qo'yish xavfi yuqori bo'lgani
uchun bu 4 buyruq botga QAYTA QO'SHILMADI. Agar sizga ular kerak bo'lsa —
asl kodni topib bering, men ularni qayta ulab beraman.
"""

import os
import json
import re

import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

PHARMACY_SHEET_ID = os.getenv("PHARMACY_SHEET_ID", "")

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


def _get_firmalar_ws():
    """PHARMACY_SHEET_ID ichidagi "Firmalar" varag'ini qaytaradi."""
    client = _get_client()
    sh = client.open_by_key(PHARMACY_SHEET_ID)
    return sh.worksheet("Firmalar")


def get_filial_info(filial_no) -> dict | None:
    """
    Berilgan filial RAQAMI bo'yicha Farmatsevtlar jadvalidagi (A ustun)
    "RAQAM - NOM" formatidagi TO'LIQ filial nomini qaytaradi.

    Masalan: get_filial_info("1") -> {"filial_no": "1", "filial_nomi": "1 - ТАШМИ-1"}

    MUHIM: registration_handlers.py dagi get_filiallar() bilan BIR XIL
    ustun formatiga (A ustunda "RAQAM - NOM") tayanadi — shu jadval
    tuzilishi o'zgarsa, ikkalasini birga yangilash kerak.
    """
    filial_no = str(filial_no).strip()
    if not filial_no:
        return None
    try:
        client = _get_client()
        ws = client.open_by_key(PHARMACY_SHEET_ID).worksheet("Farmatsevtlar")
        all_values = ws.get_all_values()
        seen = set()
        for row in all_values[1:]:
            a_val = str(row[0]).strip() if row else ""
            if not a_val or a_val in seen:
                continue
            seen.add(a_val)
            m = re.match(r"^(\d+)\s*[-—]\s*(.+)$", a_val)
            if m and m.group(1) == filial_no:
                return {"filial_no": filial_no, "filial_nomi": a_val}
        return None
    except Exception as e:
        print(f"[REG_UTILS] get_filial_info xato: {e}")
        return None
