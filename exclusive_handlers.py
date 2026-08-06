"""
exclusive_handlers.py — Eksklyuziv dori: filiallar bo'yicha "yaxshi ketmoqda /
ketmagan" tahlili va yaqin (dori hali yo'q) aptekalar ro'yxatini chiqarish.

Ishlash tartibi (admin uchun):
1. Admin botga /eksklyuziv (kerak bo'lsa radius bilan: /eksklyuziv 1.5) yozadi.
2. Bot shu dorining "aylanma" (turnover) hisobotini .xlsx fayl sifatida
   yuborishni so'raydi (1C/boshqa tizimdan olingan, "Оборот по сети" formatidagi
   hisobot — har bir filial uchun: qoldiq boshi, sotib olindi, SOTUV, ko'chirildi,
   qoldiq oxiri ustunlari bilan).
3. Admin faylni yuboradi.
4. Bot:
   - Filiallarni SOTILGAN miqdoriga qarab toifalarga ajratadi
     (Yaxshi ketmoqda / Sekin ketmoqda / Ketmagan / Yangi berilgan).
   - Filial nomlarini joylashuv ma'lumotlari bilan (Sheet1 — bot.py dagi
     load_df() ishlatadigan bir xil manba, Google Sheets export orqali,
     gspread credential SHART EMAS) mos keladi.
   - "Yaxshi ketmoqda" filiallarga berilgan radius (km) ichida joylashgan,
     bu dori HALI YO'Q aptekalarni topadi.
   - Natijani xabar + tayyor .xlsx fayl sifatida qaytaradi.

MUHIM: bu modul PHARMACY_SHEET_ID/SHEETS_ID dagi "Sheet1" varag'ini FAQAT
o'qish uchun ishlatadi (yozish YO'Q) — shuning uchun registratsiya/davomat
modullaridagi qaysi varaq nima uchun ishlatilishi bilan hech qanday
to'qnashuv yo'q.
"""

import os
import io
import re
import math
import logging

import requests
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from thefuzz import fuzz, process as fuzz_process

from telegram import Update
from telegram.ext import MessageHandler, CommandHandler, ContextTypes, filters

from attendance import run_read
from salary_handlers import ADMIN_IDS

logger = logging.getLogger(__name__)

SHEETS_ID = os.getenv("SHEETS_ID", "1CfuogH-yY--y5kiBK0qXsl5AFi_Hmzj_onWcA-Qyvco")
LOCATION_SHEET_NAME = "Sheet1"
DEFAULT_RADIUS_KM = 2.0

# Suhbat holati (bot.py dagi boshqa state raqamlari bilan TO'QNASHMASLIGI
# uchun 700 dan boshlangan; agar bot.py da allaqachon 700+ band bo'lsa,
# shu raqamni bot.py bilan birga o'zgartiring).
EKS_WAIT_XLSX = 700

# Ba'zi filiallar nomi aylanma hisobotida va Sheet1'da tizimli ravishda
# har xil yozilishi ma'lum bo'lgan holatlar uchun qo'lda mos yozuv.
# thefuzz bularni ishonchli aniqlay olmaydi (masalan bir nechta "гор
# больница" nuqtasi bo'lgani uchun).
KNOWN_ALIASES = {
    "1 гор больница (1)": "(1) ГОР БОЛЬНИЦА",
    "1 гор больница (2)": "(2) ГОР БОЛЬНИЦА",
    "1 гор больница фарм люкс": "ФАРМ ЛЮКС",
    "ташсельмаш": "ТАШСЕЛМАШ",
    "ахмад дониш": "ЮНУСОБОД АХМАД ДОНИШ",
    "кадышева базар": "КАДЕШВА БОЗОР",
}
FUZZY_MATCH_THRESHOLD = 80  # shundan past bo'lsa "tekshirish kerak" deb belgilanadi
SKIP_FILIAL_NAMES = {"по сети", "склад"}


def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlmb = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlmb / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def _load_location_df():
    """
    Filiallar joylashuvini (Filial №, Nomi, Latitude, Longitude) o'qiydi.
    bot.py dagi load_df() bilan AYNAN bir xil manba/usul: Google Sheets
    export (HTTP), gspread credential shart emas. Bu funksiya BLOKLOVCHI
    (tarmoq so'rovi) — faqat run_read() orqali chaqirilishi kerak.
    """
    url = f"https://docs.google.com/spreadsheets/d/{SHEETS_ID}/export?format=xlsx&sheet={LOCATION_SHEET_NAME}"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    df = pd.read_excel(io.BytesIO(resp.content)).fillna("")
    required = ["Filial №", "Nomi (UZ)", "Latitude", "Longitude"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise RuntimeError(f"Sheet1 da kerakli ustunlar topilmadi: {missing}")
    return df


def classify(sold: int, ost_boshi, peremeshenie) -> str:
    sold = abs(sold or 0)
    ost_boshi = ost_boshi or 0
    peremeshenie = peremeshenie or 0
    if sold >= 3:
        return "Yaxshi ketmoqda"
    if 1 <= sold <= 2:
        return "Sekin ketmoqda"
    if sold == 0 and ost_boshi > 0:
        return "Ketmagan (zaxira bor, sotuv yo'q)"
    if sold == 0 and ost_boshi == 0 and peremeshenie > 0:
        return "Yangi berilgan (hali erta baholash uchun)"
    return "Aniqlanmagan"


def _parse_turnover_xlsx(xlsx_bytes: bytes):
    """Turnover_report.xlsx dan (birinchi varaq) filiallar bo'yicha qatorlarni o'qiydi."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    # Birinchi ustundagi qiymat dori nomi bo'lgan qatorni topamiz (2-3 qatorlarda
    # sarlavha bo'lishi mumkin, shuning uchun filial nomi bo'sh bo'lmagan va
    # "по сети"/"склад" bo'lmagan birinchi qatordan boshlaymiz).
    product_name = None
    data_rows = []
    for r in rows:
        if r is None or len(r) < 12:
            continue
        filial = r[2]
        if not filial:
            continue
        ost_boshi, sotuv, peremeshenie, ost_oxiri = r[3] or 0, r[7] or 0, r[9] or 0, r[11] or 0
        # Faqat RAQAMLI qiymatga ega qatorlar haqiqiy ma'lumot qatori
        # hisoblanadi — shu tekshiruv orqali sarlavha qatorlarini (ustun
        # nomlari matn bo'lgani uchun) chetlab o'tamiz. MUHIM: mahsulot
        # nomini ('product_name') ham FAQAT shu tasdiqlangan qatordan
        # olamiz — aks holda 1-qatordagi ustun sarlavhasi ("наименование")
        # xato ravishda mahsulot nomi sifatida qabul qilinib qolardi.
        if not isinstance(ost_boshi, (int, float)):
            continue
        if product_name is None and r[0]:
            product_name = r[0]
        if filial in SKIP_FILIAL_NAMES:
            continue
        data_rows.append({
            "filial_raw": str(filial).strip(),
            "ost_boshi": ost_boshi, "sotuv": sotuv,
            "peremeshenie": peremeshenie, "ost_oxiri": ost_oxiri,
        })
    return product_name or "Nomalum dori", data_rows


def _match_location(filial_raw: str, loc_df, name_choices, name_to_row):
    key = filial_raw.strip().lower()
    if key in KNOWN_ALIASES:
        target = KNOWN_ALIASES[key].strip().upper()
        row = name_to_row.get(target)
        if row is not None:
            return row, "alias", 100
    best = fuzz_process.extractOne(filial_raw, name_choices, scorer=fuzz.token_sort_ratio)
    if best is None:
        return None, "no_match", 0
    name, score = best[0], best[1]
    row = name_to_row.get(name.strip().upper())
    how = "fuzzy" if score >= FUZZY_MATCH_THRESHOLD else "low_confidence"
    return row, how, score


def _process_eksklyuziv(xlsx_bytes: bytes, radius_km: float):
    """
    HAMMA blokловchi ish (tarmoq + fayl parsing) shu yerda — run_read() orqali
    alohida oqimda ishlaydi, botning asosiy event loop'ini to'xtatib qo'ymaydi.
    Qaytaradi: dict(product_name, matched, unmatched, good, candidates, xlsx_bytes)
    """
    product_name, turnover_rows = _parse_turnover_xlsx(xlsx_bytes)
    loc_df = _load_location_df()

    name_choices = loc_df["Nomi (UZ)"].astype(str).tolist()
    name_to_row = {}
    for _, row in loc_df.iterrows():
        nom = str(row["Nomi (UZ)"]).strip().upper()
        if nom and nom not in name_to_row:
            name_to_row[nom] = row

    matched, unmatched = [], []
    for t in turnover_rows:
        row, how, score = _match_location(t["filial_raw"], loc_df, name_choices, name_to_row)
        if row is None or how == "low_confidence" or how == "no_match":
            t["best_guess"] = row["Nomi (UZ)"] if row is not None else ""
            t["match_score"] = score
            unmatched.append(t)
            continue
        t["category"] = classify(t["sotuv"], t["ost_boshi"], t["peremeshenie"])
        t["loc_nom"] = row["Nomi (UZ)"]
        t["loc_no"] = row["Filial №"]
        t["lat"] = float(row["Latitude"])
        t["lon"] = float(row["Longitude"])
        matched.append(t)

    has_drug_nos = set(t["loc_no"] for t in matched)
    good = [t for t in matched if t["category"] == "Yaxshi ketmoqda"]

    candidates_by_no = {}
    for t in good:
        for _, cd in loc_df.iterrows():
            cd_no = cd["Filial №"]
            if cd_no in has_drug_nos:
                continue
            try:
                clat, clon = float(cd["Latitude"]), float(cd["Longitude"])
            except (TypeError, ValueError):
                continue
            dist = haversine_km(t["lat"], t["lon"], clat, clon)
            if dist <= radius_km:
                key = cd_no
                cand = {
                    "candidate_nom": cd["Nomi (UZ)"], "candidate_no": cd_no,
                    "candidate_lat": clat, "candidate_lon": clon,
                    "good_filial": t["filial_raw"], "good_sold": abs(t["sotuv"]),
                    "distance_km": round(dist, 2),
                }
                if key not in candidates_by_no or dist < candidates_by_no[key]["distance_km"]:
                    candidates_by_no[key] = cand

    dedup_candidates = sorted(candidates_by_no.values(), key=lambda x: x["distance_km"])

    xlsx_out = _build_report_xlsx(product_name, matched, unmatched, radius_km, dedup_candidates)

    return {
        "product_name": product_name,
        "matched": matched, "unmatched": unmatched,
        "good": good, "candidates": dedup_candidates,
        "xlsx_bytes": xlsx_out,
    }


def _build_report_xlsx(product_name, matched, unmatched, radius_km, dedup_candidates):
    FONT = "Arial"
    HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
    HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=11)
    TITLE_FONT = Font(name=FONT, bold=True, size=14)
    NORMAL_FONT = Font(name=FONT, size=10)
    THIN = Side(style="thin", color="D0D0D0")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    CAT_COLORS = {
        "Yaxshi ketmoqda": "C6EFCE",
        "Ketmagan (zaxira bor, sotuv yo'q)": "FFC7CE",
        "Sekin ketmoqda": "FFEB9C",
        "Yangi berilgan (hali erta baholash uchun)": "D9D9D9",
    }

    wb = openpyxl.Workbook()
    s2 = wb.active
    s2.title = "Filiallar holati"
    headers2 = ["Filial", "Davr boshi qoldiq", "Sotildi (dona)", "Ko'chirildi",
                "Davr oxiri qoldiq", "Toifa", "Lokatsiya (Sheet1)", "Latitude", "Longitude"]
    for i, h in enumerate(headers2, 1):
        c = s2.cell(row=1, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    s2.freeze_panes = "A2"
    row_i = 2
    for t in sorted(matched, key=lambda x: -abs(x["sotuv"])):
        vals = [t["filial_raw"], t["ost_boshi"], abs(t["sotuv"]), t["peremeshenie"],
                t["ost_oxiri"], t["category"], t["loc_nom"], t["lat"], t["lon"]]
        for i, v in enumerate(vals, 1):
            c = s2.cell(row=row_i, column=i, value=v)
            c.font = NORMAL_FONT
            c.border = BORDER
            c.fill = PatternFill("solid", fgColor=CAT_COLORS.get(t["category"], "FFFFFF"))
        row_i += 1
    for t in unmatched:
        vals = [t["filial_raw"], t["ost_boshi"], abs(t["sotuv"]), t["peremeshenie"], t["ost_oxiri"],
                f"Lokatsiya aniq topilmadi (eng yaqin taxmin: {t.get('best_guess','')}, {t.get('match_score',0)}%)",
                "", "", ""]
        for i, v in enumerate(vals, 1):
            c = s2.cell(row=row_i, column=i, value=v)
            c.font = NORMAL_FONT
            c.border = BORDER
            c.fill = PatternFill("solid", fgColor="BFBFBF")
        row_i += 1
    widths2 = [28, 16, 14, 12, 16, 40, 34, 12, 12]
    for i, w in enumerate(widths2, 1):
        s2.column_dimensions[get_column_letter(i)].width = w

    s3 = wb.create_sheet("Yaqin nomzod aptekalar")
    s3["A1"] = f"«{product_name}» hali YO'Q, lekin yaxshi ketayotgan filialga {radius_km} km ichida joylashgan aptekalar"
    s3["A1"].font = TITLE_FONT
    s3.merge_cells("A1:F1")
    headers3 = ["Nomzod apteka (dori yo'q)", "Masofa (km)", "Eng yaqin \"yaxshi ketayotgan\" filial",
                "O'sha filialda sotilgan (dona)", "Nomzod Latitude", "Nomzod Longitude"]
    for i, h in enumerate(headers3, 1):
        c = s3.cell(row=3, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    s3.freeze_panes = "A4"
    row_i = 4
    for c in dedup_candidates:
        vals = [c["candidate_nom"], c["distance_km"], c["good_filial"], c["good_sold"],
                c["candidate_lat"], c["candidate_lon"]]
        for i, v in enumerate(vals, 1):
            cell = s3.cell(row=row_i, column=i, value=v)
            cell.font = NORMAL_FONT
            cell.border = BORDER
        row_i += 1
    widths3 = [34, 12, 34, 20, 14, 14]
    for i, w in enumerate(widths3, 1):
        s3.column_dimensions[get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─── Telegram handlerlar ───────────────────────────────────────────────────

async def cmd_eksklyuziv(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """/eksklyuziv [radius_km] — admin buyrug'i: turnover xlsx faylni so'raydi."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Ruxsat yo'q.")
        return None

    radius = DEFAULT_RADIUS_KM
    if ctx.args:
        try:
            radius = float(ctx.args[0].replace(",", "."))
        except (ValueError, IndexError):
            pass
    ctx.user_data["eks_radius"] = radius

    await update.message.reply_text(
        f"📦 *Eksklyuziv dori tahlili*\n\n"
        f"Ushbu dorining aylanma (turnover) hisobotini *.xlsx* fayl sifatida yuboring "
        f"(\"Оборот по сети\" formatida — filial, qoldiq, sotuv ustunlari bilan).\n\n"
        f"📍 Radius: *{radius} km* (o'zgartirish uchun: `/eksklyuziv 1.5`)",
        parse_mode="Markdown",
    )
    return EKS_WAIT_XLSX


async def eks_receive_xlsx(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Turnover xlsx faylni qabul qiladi, tahlil qiladi va natijani qaytaradi."""
    if update.effective_user.id not in ADMIN_IDS:
        return None

    if not update.message.document or not update.message.document.file_name.lower().endswith(".xlsx"):
        await update.message.reply_text("❌ Iltimos, .xlsx fayl yuboring.")
        return EKS_WAIT_XLSX

    radius = ctx.user_data.get("eks_radius", DEFAULT_RADIUS_KM)
    msg = await update.message.reply_text("⏳ Tahlil qilinmoqda (bir necha soniya)...")

    try:
        file = await update.message.document.get_file()
        xlsx_bytes = bytes(await file.download_as_bytearray())
        result = await run_read(_process_eksklyuziv, xlsx_bytes, radius)
    except Exception as e:
        logger.exception("[EKS] Tahlil xatosi")
        await msg.edit_text(f"❌ Xatolik yuz berdi: {e}")
        return None

    n_matched = len(result["matched"])
    n_unmatched = len(result["unmatched"])
    n_good = len(result["good"])
    n_cand = len(result["candidates"])

    lines = [
        f"✅ *«{result['product_name']}»* tahlili tayyor",
        f"",
        f"🏪 Jami filial (dori bor): {n_matched + n_unmatched}",
        f"📍 Lokatsiyasi topildi: {n_matched}" + (f" (⚠️ {n_unmatched} tasi topilmadi)" if n_unmatched else ""),
        f"🟢 Yaxshi ketmoqda: {n_good}",
        f"",
        f"🎯 {radius} km radiusda topilgan, dori HALI YO'Q nomzod aptekalar: *{n_cand}*",
    ]
    await msg.edit_text("\n".join(lines), parse_mode="Markdown")

    out_name = f"eksklyuziv_{result['product_name'][:20]}.xlsx".replace(" ", "_")
    file_io = io.BytesIO(result["xlsx_bytes"])
    file_io.name = out_name
    await ctx.bot.send_document(chat_id=update.effective_chat.id, document=file_io, filename=out_name)

    return None


def get_eks_states():
    return {
        EKS_WAIT_XLSX: [
            MessageHandler(filters.Document.ALL, eks_receive_xlsx),
        ],
    }
