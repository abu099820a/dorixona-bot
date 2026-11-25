"""
vozvrat_handlers.py — Murojaat bo'limining kengaytmasi:

  💊 Dorixonalarga → 📝 Oddiy murojaat     (eski oqim, salary_handlers.py da)
                   → 📦 Skladga vozvrat    (skrinshot / Excel → har filialga o'z ro'yxati)
                   → 🔄 Peremeshenie       (tayyor Excel → beruvchi va qabul qiluvchi zavlarga)

Skrinshotlarni o'qish uchun Claude (Anthropic) vision API ishlatiladi:
  ANTHROPIC_API_KEY  — Railway Variables ga qo'shiladi (majburiy, rasm uchun)
  ANTHROPIC_MODEL    — ixtiyoriy (standart: claude-sonnet-4-5)
Excel fayl yuborilsa — API kerak emas.

1C dagi filial nomlari bazadagidan farq qilsa (masalan "ШАХРИСТАНСКАЯ" ↔
"ШАХРИСТАНСКИЙ"), bot o'xshashlik bo'yicha topadi. Topa olmasa — admin
"1C nomi = Filial №" ko'rinishida yozib bog'laydi va bu "1C nomlari"
varag'iga saqlanadi (keyingi safar avtomatik topiladi).
"""

import os
import io
import re
import json
import base64
import asyncio
import logging
import difflib
from collections import OrderedDict, defaultdict

from telegram import (
    Update, ReplyKeyboardMarkup, InlineKeyboardMarkup, InlineKeyboardButton, InputFile,
)
from telegram.ext import (
    ContextTypes, MessageHandler, CallbackQueryHandler, filters, ApplicationHandlerStop,
)

logger = logging.getLogger(__name__)

# ─── Holatlar (ConversationHandler) ──────────────────────────────────────────
APPEAL_DORIXONA_TYPE = 590   # Oddiy / Vozvrat / Peremeshenie tanlash
VOZVRAT_UPLOAD       = 591   # Skrinshot(lar) yoki Excel qabul qilish
VOZVRAT_CONFIRM      = 592   # Tahlil natijasi → Yuborish / Bekor / alias
PEREMESH_UPLOAD      = 593   # Peremeshenie Excel faylini qabul qilish
PEREMESH_CONFIRM     = 594   # Tahlil natijasi → Yuborish / Bekor / alias

ALIAS_WS_NAME = "1C nomlari"   # PHARMACY_SHEET_ID ichidagi varaq: "1C nomi" | "Filial №"

BTN_ODDIY    = "📝 Oddiy murojaat"
BTN_VOZVRAT  = "📦 Skladga vozvrat"
BTN_PEREMESH = "🔄 Peremeshenie"
BTN_TAYYOR   = "✅ Tayyor"
BTN_BEKOR    = "❌ Bekor qilish"
BTN_TASKS    = "📥 Tayyor topshiriqlar"


def dorixona_type_keyboard(language: str = "uz"):
    back_txt = "⬅️ Назад" if language == "ru" else "⬅️ Orqaga"
    return ReplyKeyboardMarkup(
        [[BTN_ODDIY], [BTN_VOZVRAT], [BTN_PEREMESH], [BTN_TASKS], [back_txt]],
        resize_keyboard=True,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# 1) FILIAL NOMLARINI MOSLASHTIRISH
# ═══════════════════════════════════════════════════════════════════════════════

_CHAR_MAP = str.maketrans({
    "Ё": "Е", "Ў": "У", "Қ": "К", "Ғ": "Г", "Ҳ": "Х", "Й": "И", "Ъ": "", "Ь": "",
})


def _nm(s) -> str:
    """Nomni solishtirish uchun normallashtirish: katta harf, Ё→Е, belgilarsiz."""
    s = str(s or "").upper().translate(_CHAR_MAP)
    s = re.sub(r"[^0-9A-ZА-Я]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _tok(s: str) -> str:
    return " ".join(sorted(_nm(s).split()))


def _digits(s: str) -> set:
    return set(re.findall(r"\d+", _nm(s)))


def _load_alias_map() -> dict:
    """'1C nomlari' varag'idan {_nm(1C nomi): filial_no} qaytaradi."""
    from salary_handlers import _get_client, PHARMACY_SHEET_ID
    try:
        ws = _get_client().open_by_key(PHARMACY_SHEET_ID).worksheet(ALIAS_WS_NAME)
        out = {}
        for row in ws.get_all_values()[1:]:
            if len(row) >= 2 and row[0].strip() and row[1].strip():
                out[_nm(row[0])] = row[1].strip().replace(".0", "")
        return out
    except Exception:
        return {}


def _save_aliases(pairs: list) -> bool:
    """[(1C nomi, filial_no), ...] ni '1C nomlari' varag'iga yozadi (yo'q bo'lsa yaratadi)."""
    from salary_handlers import _get_client, PHARMACY_SHEET_ID
    try:
        sh = _get_client().open_by_key(PHARMACY_SHEET_ID)
        try:
            ws = sh.worksheet(ALIAS_WS_NAME)
        except Exception:
            ws = sh.add_worksheet(title=ALIAS_WS_NAME, rows=500, cols=3)
            ws.append_row(["1C nomi", "Filial №", "Bazadagi nomi"])
        ws.append_rows([[a, f, n] for a, f, n in pairs], value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        logger.error(f"[ALIAS] Saqlashda xato: {e}")
        return False


def _load_pharmacies() -> dict:
    """Sheet1 dagi barcha filiallar: {filial_no: {filial_no, nomi, tid, ...}} + alias xaritasi."""
    from salary_handlers import get_all_dorixonalar_for_broadcast
    ph = OrderedDict()
    for d in get_all_dorixonalar_for_broadcast(None):
        if d.get("filial_no"):
            ph[d["filial_no"]] = d
    return {"ph": ph, "alias": _load_alias_map()}


def match_filial(raw: str, base: dict):
    """
    1C dagi nomni bazadagi filialga moslaydi.
    Qaytaradi: (filial_no | None, izoh)
      - alias → aniq
      - normallashtirilgan nom teng → aniq
      - o'xshashlik ≥ 0.80 va ikkinchisidan aniq ustun → topildi
      - aks holda None (topilmadi yoki noaniq)
    """
    ph, alias = base["ph"], base["alias"]
    n = _nm(raw)
    if not n:
        return None, "bo'sh"
    # Filial raqami bilan yozilgan bo'lsa: "48" yoki "48 - ШАХРИСТАНСКИЙ"
    mnum = re.match(r"^\s*(\d+)\s*(-|$)", str(raw))
    if mnum and mnum.group(1) in ph:
        return mnum.group(1), "aniq"
    if n in alias and alias[n] in ph:
        return alias[n], "alias"

    t = _tok(raw)
    for fno, d in ph.items():
        if n == _nm(d["nomi"]) or t == _tok(d["nomi"]):
            return fno, "aniq"

    rd = _digits(raw)
    scored = []
    for fno, d in ph.items():
        cd = _digits(d["nomi"])
        # Raqamlar zid bo'lsa (ЮНУСАБАД 15 ↔ ЮНУСАБАД 18) — nomzod emas
        if rd and cd and not (rd <= cd or cd <= rd):
            continue
        cn = _nm(d["nomi"])
        # Qisqa nom bazadagi to'liq nom ichida bo'lsa ham topilsin:
        # "БЕШКАЙРАГАЧ" ↔ "УЧТЕПА БЕШКАЙРАГОЧ" (har bir so'z eng o'xshash so'zga solishtiriladi)
        rt = [w for w in n.split() if not w.isdigit()]
        ct = cn.split()
        cover = (sum(max(difflib.SequenceMatcher(None, w, c).ratio() for c in ct) for w in rt) / len(rt)
                 if rt and ct and sum(len(w) for w in rt) >= 6 else 0.0)
        s = max(
            difflib.SequenceMatcher(None, n, cn).ratio(),
            difflib.SequenceMatcher(None, t, _tok(d["nomi"])).ratio(),
            cover * 0.97,
        )
        scored.append((s, fno))
    scored.sort(reverse=True)
    if not scored:
        return None, "topilmadi"
    best_s, best_f = scored[0]
    second = scored[1][0] if len(scored) > 1 else 0.0
    if best_s >= 0.80 and best_s - second >= 0.05:
        return best_f, f"o'xshash {best_s:.0%}"
    if best_s >= 0.80:
        return None, f"noaniq: {ph[best_f]['nomi']} / {ph[scored[1][1]]['nomi']}"
    return None, "topilmadi"


def _label(d: dict) -> str:
    return f"{d['filial_no']} - {d['nomi']}" if d.get("filial_no") else d.get("nomi", "")


def _fmt_qty(q) -> str:
    try:
        q = float(q)
    except Exception:
        return str(q)
    return str(int(q)) if q == int(q) else f"{q:g}".replace(".", ",")


def _to_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# 2) FAYL / RASMNI O'QISH
# ═══════════════════════════════════════════════════════════════════════════════

def _find_header(rows: list, groups: dict, optional: tuple = (), scan: int = 15):
    """
    Sarlavha qatorini topadi.
    groups = {"kalit": [regex1, regex2, ...]} — regexlar ustuvorlik tartibida
    (avval 1-regex barcha ustunlarda qidiriladi, topilmasa 2-si va h.k.).
    optional — topilmasa ham bo'ladigan kalitlar.
    Qaytaradi: (sarlavha_qator_indeksi, {kalit: ustun_indeksi}) yoki (None, None)
    """
    need = [k for k in groups if k not in optional]
    for ri, row in enumerate(rows[:scan]):
        cells = [str(c or "").strip().lower() for c in row]
        found, used = {}, set()
        for key, pats in groups.items():
            for pat in pats:
                ci = next((i for i, c in enumerate(cells)
                           if i not in used and c and re.search(pat, c)), None)
                if ci is not None:
                    found[key] = ci
                    used.add(ci)
                    break
        if all(k in found for k in need):
            return ri, found
    return None, None


def _read_xlsx_sheets(data: bytes) -> list:
    """Excel (xlsx) ning BARCHA varaqlari: [[qator, ...], ...]"""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    return [[list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets]


def _cell(r, ci):
    return r[ci] if ci is not None and ci < len(r) else None


def _is_total_row(r) -> bool:
    return any(re.match(r"^\s*(итого|всего|jami)", str(c or ""), re.I) for c in r)


def parse_vozvrat_excel(data: bytes) -> list:
    """Vozvrat Excel: Филиал | Наименование | Кол-во → [{filial, dori, soni}]"""
    out, found_any = [], False
    for rows in _read_xlsx_sheets(data):
        hi, cols = _find_header(rows, {
            "filial": [r"филиал", r"аптека", r"filial"],
            "dori":   [r"наименование", r"препарат", r"товар", r"dori"],
            "soni":   [r"^кол", r"кол-во", r"soni"],
        })
        if hi is None:
            continue
        found_any = True
        for r in rows[hi + 1:]:
            if _is_total_row(r):
                break
            f = str(_cell(r, cols["filial"]) or "").strip()
            d = str(_cell(r, cols["dori"]) or "").strip()
            q = _to_num(_cell(r, cols["soni"]))
            if f and d and q:
                out.append({"filial": f, "dori": d, "soni": q})
    if not found_any:
        raise ValueError("Sarlavha topilmadi (kerakli ustunlar: Филиал, Наименование, Кол-во)")
    return out


def _drug_from_title(rows: list, hi: int) -> str:
    """Dori ustuni yo'q bo'lsa — sarlavhadan yuqoridagi nomdan olinadi:
    'Перемещение товара: АМИКОС ФЛАК 250МЛ №1 / FUREN ...' → 'АМИКОС ФЛАК 250МЛ №1 / FUREN ...'"""
    texts = [str(c).strip() for r in rows[:hi] for c in r if isinstance(c, str) and c.strip()]
    for t in texts:
        m = re.search(r"(?:товар[а-я]*|препарат[а-я]*|dori)\s*[:\-–]\s*(.+)$", t, re.I)
        if m:
            return m.group(1).strip()
    for t in texts:
        if ":" in t:
            return t.split(":", 1)[1].strip()
    return texts[0] if texts else ""


def parse_peremesh_excel(data: bytes) -> list:
    """
    Peremeshenie Excel → [{dan, ga, dori, soni}]
    Qo'llab-quvvatlanadigan ko'rinishlar:
      • Откуда | Кол-во | Препарат | Куда                (bir nechta dori bitta jadvalda)
      • Откуда | Куда | Остаток | Кол-во к перемещению   (dori nomi jadval tepasidagi sarlavhada)
    Har bir varaq alohida o'qiladi (har varaqda boshqa dori bo'lishi mumkin).
    """
    out, found_any = [], False
    for rows in _read_xlsx_sheets(data):
        hi, cols = _find_header(rows, {
            "dan":  [r"^откуда", r"отправител", r"qayerdan", r"beruvchi"],
            "ga":   [r"^куда", r"получател", r"qayerga", r"oluvchi"],
            "dori": [r"препарат", r"наименование", r"товар", r"dori"],
            "soni": [r"к перемещ", r"перемещ", r"^кол", r"кол-во", r"soni"],
        }, optional=("dori",))
        if hi is None:
            continue
        found_any = True
        title_drug = _drug_from_title(rows, hi) if "dori" not in cols else ""
        for r in rows[hi + 1:]:
            if _is_total_row(r):
                break
            dan = str(_cell(r, cols["dan"]) or "").strip()
            ga = str(_cell(r, cols["ga"]) or "").strip()
            dori = str(_cell(r, cols.get("dori")) or "").strip() or title_drug
            q = _to_num(_cell(r, cols["soni"]))
            if dan and ga and dori and q:
                out.append({"dan": dan, "ga": ga, "dori": dori, "soni": q})
    if not found_any:
        raise ValueError("Sarlavha topilmadi (kerakli ustunlar: Откуда, Куда, Кол-во, Препарат)")
    return out


_VISION_PROMPT = (
    "Bu dorixona dasturidagi jadval skrinshoti. Ustunlar: Ид, Филиал, Наименование, Кол-во.\n"
    "Har bir ma'lumot qatorini (sarlavhasiz) aynan yozilganidek ko'chir — kirill harflarini, "
    "raqamlarni va belgilarni o'zgartirma, hech bir qatorni tashlab ketma, takroriy qatorlarni "
    "ham alohida yoz. Qatordagi rang (yashil/qizil) ahamiyatsiz.\n"
    "Javobni FAQAT JSON massiv ko'rinishida ber, boshqa hech narsa yozma:\n"
    '[{"filial": "ТАШМИ-1", "dori": "ФОРЛАК ФОРТЕ КАПЛИ 30МЛ\\\\FARMFORMAT, ООО (УЗБЕКИСТАН)", "soni": 2}]'
)


async def vision_extract_rows(image_bytes: bytes, mime: str = "image/jpeg") -> list:
    """Skrinshotdagi jadvalni Claude vision orqali o'qiydi → [{filial, dori, soni}]"""
    api_key = os.getenv("ANTHROPIC_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("NO_API_KEY")
    import httpx
    payload = {
        "model": os.getenv("ANTHROPIC_MODEL", "claude-sonnet-4-5"),
        "max_tokens": 8000,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image", "source": {
                    "type": "base64", "media_type": mime,
                    "data": base64.b64encode(image_bytes).decode(),
                }},
                {"type": "text", "text": _VISION_PROMPT},
            ],
        }],
    }
    async with httpx.AsyncClient(timeout=180) as client:
        r = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json=payload,
        )
    if r.status_code != 200:
        raise RuntimeError(f"API {r.status_code}: {r.text[:200]}")
    text = "".join(b.get("text", "") for b in r.json().get("content", []) if b.get("type") == "text")
    a, b = text.find("["), text.rfind("]")
    if a < 0 or b < 0:
        raise RuntimeError("Javobda JSON topilmadi")
    out = []
    for it in json.loads(text[a:b + 1]):
        f = str(it.get("filial", "")).strip()
        d = str(it.get("dori", "")).strip()
        q = _to_num(it.get("soni"))
        if f and d and q:
            out.append({"filial": f, "dori": d, "soni": q})
    return out


# ═══════════════════════════════════════════════════════════════════════════════
# 3) TAHLIL (filiallarga ajratish)
# ═══════════════════════════════════════════════════════════════════════════════

def build_vozvrat_plan(rows: list, base: dict) -> dict:
    """
    rows → har filial uchun dorilar ro'yxati.
    Qaytaradi: {
      "items":     OrderedDict{filial_no: OrderedDict{dori: soni}},
      "raw_names": {filial_no: set(1C nomlari)},
      "fuzzy":     [(1C nomi, filial_no, izoh)],   # o'xshashlik bo'yicha topilganlar
      "unmatched": OrderedDict{1C nomi: (izoh, OrderedDict{dori: soni})},
    }
    """
    items = OrderedDict()
    raw_names = defaultdict(set)
    fuzzy, unmatched, cache = [], OrderedDict(), {}
    for r in rows:
        raw = r["filial"]
        if raw not in cache:
            cache[raw] = match_filial(raw, base)
        fno, why = cache[raw]
        dori = re.sub(r"\s+", " ", r["dori"]).strip()
        if fno:
            raw_names[fno].add(raw)
            items.setdefault(fno, OrderedDict())
            items[fno][dori] = items[fno].get(dori, 0) + r["soni"]
        else:
            u = unmatched.setdefault(raw, (why, OrderedDict()))
            u[1][dori] = u[1].get(dori, 0) + r["soni"]
    for raw, (fno, why) in cache.items():
        if fno and why.startswith("o'xshash"):
            fuzzy.append((raw, fno, why))
    return {"items": items, "raw_names": raw_names, "fuzzy": fuzzy, "unmatched": unmatched}


def build_peremesh_plan(rows: list, base: dict) -> dict:
    """
    rows → beruvchi va qabul qiluvchi filiallar uchun ro'yxatlar.
    out[filial_no] = [(dori, soni, qabul_qiluvchi_label)]
    inc[filial_no] = [(dori, soni, beruvchi_label)]
    """
    ph = base["ph"]
    out, inc = OrderedDict(), OrderedDict()
    fuzzy, unmatched, cache = [], OrderedDict(), {}

    def m(raw):
        if raw not in cache:
            cache[raw] = match_filial(raw, base)
        return cache[raw]

    for r in rows:
        f_dan, why_dan = m(r["dan"])
        f_ga, why_ga = m(r["ga"])
        lbl_dan = _label(ph[f_dan]) if f_dan else r["dan"]
        lbl_ga = _label(ph[f_ga]) if f_ga else r["ga"]
        if f_dan:
            out.setdefault(f_dan, []).append((r["dori"], r["soni"], lbl_ga))
        else:
            unmatched.setdefault(r["dan"], why_dan)
        if f_ga:
            inc.setdefault(f_ga, []).append((r["dori"], r["soni"], lbl_dan))
        else:
            unmatched.setdefault(r["ga"], why_ga)
    for raw, (fno, why) in cache.items():
        if fno and why.startswith("o'xshash"):
            fuzzy.append((raw, fno, why))
    return {"out": out, "inc": inc, "fuzzy": fuzzy, "unmatched": unmatched, "rows": rows}


# ═══════════════════════════════════════════════════════════════════════════════
# 4) YORDAMCHI: yuborish, bo'lish, hisobot
# ═══════════════════════════════════════════════════════════════════════════════

def _chunks(header: str, lines: list, footer: str, limit: int = 3800) -> list:
    """Uzun ro'yxatni Telegram chegarasidan (4096) oshmaydigan bo'laklarga bo'ladi."""
    parts, cur = [], header
    for ln in lines:
        if len(cur) + len(ln) + 1 > limit:
            parts.append(cur)
            cur = "(davomi)\n\n"
        cur += ln + "\n"
    cur += footer
    parts.append(cur)
    return parts


def _sender_name(user) -> str:
    n = " ".join(filter(None, [user.first_name, user.last_name])).strip()
    return n or (f"@{user.username}" if user.username else "Administrator")


def log_murojaat_batch(entries: list) -> dict:
    """
    Murojaatlar jurnaliga bir nechta qatorni BITTA so'rov bilan yozadi.
    (Avval har bir xabar uchun alohida yozilardi — 3-4 ta API so'rovi × o'nlab xabar
    Google'ning daqiqalik limitiga urilib, jurnal yozuvlari yo'qolardi.)
    entries: [(telegram_id, matn)] → {telegram_id: qator_raqami}
    """
    from datetime import datetime
    from attendance import UZ_TZ
    from salary_handlers import _get_murojaatlar_ws
    ws = _get_murojaatlar_ws()
    now_str = datetime.now(UZ_TZ).strftime("%d.%m.%Y %H:%M")
    rows = [[now_str, str(t), str(x)[:300], "Kutilmoqda", "", "", ""] for t, x in entries]
    res = ws.append_rows(rows, value_input_option="USER_ENTERED")
    rng = ((res or {}).get("updates") or {}).get("updatedRange", "")
    m = re.search(r"![A-Z]+(\d+)", rng)
    if not m:
        return {}
    start = int(m.group(1))
    return {str(t): start + i for i, (t, _) in enumerate(entries)}


async def _send_messages(ctx, admin_chat_id: int, messages: list, progress_msg=None) -> list:
    """
    messages: [{filial_no, label, rol, tid, header, lines, footer, log_text}]
    Har biriga "✅ Bajarildi / ❌ Bekor qilish" tugmasi bilan yuboradi.
    Qaytaradi: [(filial_no, label, rol, emoji, sabab, n_items)]
    """
    from attendance import run_write
    from salary_handlers import _tg_error_reason

    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Bajarildi", callback_data=f"appealresp:done:{admin_chat_id}"),
        InlineKeyboardButton("❌ Bekor qilish", callback_data=f"appealresp:cancel:{admin_chat_id}"),
    ]])
    log_rows = ctx.bot_data.setdefault("appeal_log_rows", {})
    tid_names = ctx.bot_data.setdefault("appeal_tid_names", {})

    results = []
    to_log = []   # jurnalga bitta so'rov bilan yoziladi (Google API limitiga tushmaslik uchun)
    total = len(messages)
    for i, m in enumerate(messages, 1):
        n = len(m["lines"])
        if not m["tid"]:
            results.append((m["filial_no"], m["label"], m["rol"], "⚠️", "TID yo'q", n))
            continue
        try:
            parts = _chunks(m["header"], m["lines"], m["footer"])
            for pi, part in enumerate(parts):
                await ctx.bot.send_message(
                    chat_id=int(m["tid"]),
                    text=part,
                    reply_markup=kb if pi == len(parts) - 1 else None,
                )
            results.append((m["filial_no"], m["label"], m["rol"], "✅", "Yuborildi", n))
            tid_names[int(m["tid"])] = m["label"]
            to_log.append((m["tid"], m["log_text"]))
        except Exception as exc:
            results.append((m["filial_no"], m["label"], m["rol"], "❌", _tg_error_reason(exc), n))
            logger.warning(f"[VOZVRAT/PEREMESH] {m['label']} ga yuborilmadi: {exc}")
        if progress_msg and (i % 10 == 0 or i == total):
            try:
                await progress_msg.edit_text(f"⏳ Yuborilmoqda... {i} / {total}")
            except Exception:
                pass
        await asyncio.sleep(0.05)

    if to_log:
        try:
            rows_map = await run_write(log_murojaat_batch, to_log)
            for tid, row_idx in (rows_map or {}).items():
                log_rows[int(tid)] = row_idx
        except Exception as e:
            logger.warning(f"[VOZVRAT/PEREMESH] Jurnalga yozishda xato: {e}")
    return results


def _report_xlsx(results: list, detail_header: list, detail_rows: list, title: str, extra=None) -> io.BytesIO:
    """Natija + batafsil ro'yxat — ikki varaqli Excel hisobot."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = openpyxl.Workbook()
    hdr_font = Font(bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    fills = {
        "✅": PatternFill("solid", fgColor="E2EFDA"),
        "❌": PatternFill("solid", fgColor="FCE4D6"),
        "⚠️": PatternFill("solid", fgColor="FFEB9C"),
    }

    def sheet(ws, header, rows, widths, emoji_col=None):
        ws.append(header)
        for c in ws[1]:
            c.font, c.fill = hdr_font, hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in rows:
            ws.append(list(r))
            if emoji_col is not None and r[emoji_col] in fills:
                for c in ws[ws.max_row]:
                    c.fill = fills[r[emoji_col]]
        for i, w in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = w
        ws.freeze_panes = "A2"

    ws1 = wb.active
    if results:
        ws1.title = "Natija"
        sheet(ws1, ["Filial №", "Dorixona", "Kimga", "Holat", "Sabab", "Pozitsiyalar"],
              results, [10, 34, 18, 8, 32, 12], emoji_col=3)
        ws2 = wb.create_sheet("Ro'yxat")
    else:
        ws2 = ws1
        ws2.title = "Ro'yxat"
    sheet(ws2, detail_header, detail_rows, [10, 34, 60, 10, 34, 14])
    if extra:  # (varaq nomi, sarlavha, qatorlar)
        sheet(wb.create_sheet(extra[0]), extra[1], extra[2], [34, 60, 12, 10, 16, 28])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


async def _back_to_main(ctx, chat_id: int, text: str = "📋 Asosiy menyu"):
    from bot import main_keyboard, get_lang
    for k in ("vz_rows", "vz_plan", "vz_base", "pm_rows", "pm_plan", "pm_base", "vz_files",
              "vz_title", "vz_sklad", "vz_src", "vz_src_opts", "vz_src_pick", "cv_task_id", "cv_tasks", "pm_title"):
        ctx.user_data.pop(k, None)
    await ctx.bot.send_message(chat_id=chat_id, text=text, reply_markup=main_keyboard(get_lang(ctx)))


def _parse_alias_lines(text: str) -> list:
    """'АСКИЯ 1 = 0' ko'rinishidagi qatorlar → [(1C nomi, filial_no)]"""
    out = []
    for ln in text.splitlines():
        m = re.match(r"^\s*(.+?)\s*=\s*(\d+)\s*$", ln)
        if m:
            out.append((m.group(1).strip(), m.group(2)))
    return out


async def _apply_aliases(update: Update, ctx, base_key: str) -> bool:
    """Admin yozgan aliaslarni saqlaydi va bazani yangilaydi. True — biror narsa saqlandi."""
    from attendance import run_write
    pairs = _parse_alias_lines(update.message.text or "")
    if not pairs:
        return False
    base = ctx.user_data.get(base_key) or {"ph": {}, "alias": {}}
    ok_pairs, bad = [], []
    for raw, fno in pairs:
        if fno in base["ph"]:
            ok_pairs.append((raw, fno, base["ph"][fno]["nomi"]))
            base["alias"][_nm(raw)] = fno
        else:
            bad.append(f"{raw} = {fno}")
    if ok_pairs:
        await run_write(_save_aliases, ok_pairs)
    ctx.user_data[base_key] = base
    msg = "🔗 Bog'landi:\n" + "\n".join(f"• {r} → {f} - {n}" for r, f, n in ok_pairs) if ok_pairs else ""
    if bad:
        msg += ("\n\n" if msg else "") + "⚠️ Bunday filial raqami yo'q:\n" + "\n".join(bad)
    await update.message.reply_text(msg)
    return True


# ═══════════════════════════════════════════════════════════════════════════════
# 5) TUR TANLASH: Oddiy / Vozvrat / Peremeshenie
# ═══════════════════════════════════════════════════════════════════════════════

async def show_dorixona_type_menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    language = ctx.user_data.get("lang", "uz")
    # DIQQAT: Markdown matnida "_" ishlatmang (masalan Claude_vazifa) — Telegram
    # xabarni rad etadi va bot javob bermay qoladi.
    text = (
        "💊 *Dorixonalarga murojaat* — turini tanlang:\n\n"
        f"{BTN_ODDIY} — matnli xabar\n"
        f"{BTN_VOZVRAT} — skrinshot/Excel → har filialga o'z ro'yxati\n"
        f"{BTN_PEREMESH} — tayyor Excel → beruvchi va qabul qiluvchiga\n"
        f"{BTN_TASKS} — Claude tayyorlagan ro'yxatlar"
    )
    try:
        await update.message.reply_text(text, parse_mode="Markdown",
                                        reply_markup=dorixona_type_keyboard(language))
    except Exception:
        await update.message.reply_text(text.replace("*", ""),
                                        reply_markup=dorixona_type_keyboard(language))
    return APPEAL_DORIXONA_TYPE


async def dorixona_type_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    from salary_handlers import (
        appeal_keyboard, appeal_target_keyboard, APPEAL_MENU, APPEAL_DORIXONA_TARGET,
    )
    language = ctx.user_data.get("lang", "uz")
    back_txt = "⬅️ Назад" if language == "ru" else "⬅️ Orqaga"
    txt = (update.message.text or "").strip()

    if txt == back_txt:
        await update.message.reply_text(
            "📩 *Murojaat*\n\nKimga xabar yubormoqchisiz?",
            parse_mode="Markdown",
            reply_markup=appeal_keyboard(language),
        )
        return APPEAL_MENU

    if txt == BTN_ODDIY:
        ctx.user_data["appeal_kind"] = "dorixona"
        await update.message.reply_text(
            "💊 Dorixonalarga qanday yubormoqchisiz?",
            reply_markup=appeal_target_keyboard(language),
        )
        return APPEAL_DORIXONA_TARGET

    if txt == BTN_VOZVRAT:
        for k in ("vz_title", "vz_sklad", "vz_src", "vz_src_opts", "vz_src_pick"):
            ctx.user_data.pop(k, None)
        ctx.user_data["vz_rows"] = []
        ctx.user_data["vz_files"] = 0
        await update.message.reply_text(
            "📦 *Skladga vozvrat*\n\n"
            "1️⃣ Dasturdagi jadval *skrinshot(lar)ini* yoki *Excel* faylini yuboring "
            "(ustunlar: Филиал, Наименование, Кол-во), so'ng "
            f"*{BTN_TAYYOR}* ni bosing.\n\n"
            f"2️⃣ Yoki ro'yxatni OSTATOK'dan avtomatik oling:\n"
            f"*{BTN_BY_POST}* — partiya qaysi postavshikdan kelgan bo'lsa\n"
            f"*{BTN_BY_PRED}* — spravochnikdagi predstavitelstvo bo'yicha\n\n"
            "📚 Spravochnikni yangilash uchun «итоговый справочник» (.csv yoki .xlsx) faylini shu yerga yuboring.",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup([[BTN_BY_POST, BTN_BY_PRED], [BTN_TAYYOR], [BTN_BEKOR]],
                                             resize_keyboard=True),
        )
        return VOZVRAT_UPLOAD

    if txt == BTN_TASKS:
        return await cv_list_tasks(update, ctx)

    if txt == BTN_PEREMESH:
        await update.message.reply_text(
            "🔄 *Peremeshenie*\n\n"
            "Tayyor tahlil qilingan *Excel* faylni yuboring "
            "(ustunlar: Откуда, Кол-во, Препарат, Куда).",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup([[BTN_BEKOR]], resize_keyboard=True),
        )
        return PEREMESH_UPLOAD

    await update.message.reply_text("Iltimos, tugmalardan birini tanlang.",
                                    reply_markup=dorixona_type_keyboard(language))
    return APPEAL_DORIXONA_TYPE


# ═══════════════════════════════════════════════════════════════════════════════
# 6) 📦 SKLADGA VOZVRAT
# ═══════════════════════════════════════════════════════════════════════════════

async def vozvrat_upload_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Rasm, Excel yoki "✅ Tayyor" / "❌ Bekor qilish" ni qabul qiladi."""
    msg = update.message
    rows = ctx.user_data.setdefault("vz_rows", [])

    # ── Matn ──
    if msg.text:
        txt = msg.text.strip()
        if txt == BTN_BEKOR:
            await _back_to_main(ctx, msg.chat_id, "❌ Bekor qilindi.")
            from bot import MENU
            return MENU
        if txt in (BTN_BY_POST, BTN_BY_PRED):
            return await vz_src_start(update, ctx, "post" if txt == BTN_BY_POST else "pred")
        if txt == BTN_TAYYOR:
            if not rows:
                await msg.reply_text("⚠️ Hali hech qanday ma'lumot yuklanmadi. Skrinshot yoki Excel yuboring.")
                return VOZVRAT_UPLOAD
            return await _vozvrat_analyze_and_preview(update, ctx)
        await msg.reply_text(f"📎 Skrinshot yoki Excel yuboring, tugatgach {BTN_TAYYOR} ni bosing.")
        return VOZVRAT_UPLOAD

    # ── Excel / CSV ──
    doc = msg.document
    if doc and (doc.file_name or "").lower().endswith((".xlsx", ".xlsm", ".csv")):
        try:
            f = await ctx.bot.get_file(doc.file_id)
            data = bytes(await f.download_as_bytearray())
            sprav = parse_spravochnik(data, doc.file_name or "")
        except Exception:
            sprav = None
        if sprav is not None:   # bu — spravochnik fayli
            from attendance import run_write
            wait = await msg.reply_text(f"⏳ Spravochnik saqlanmoqda ({len(sprav)} ta tovar)...")
            ok = await run_write(save_spravochnik, sprav)
            await wait.edit_text(
                f"✅ Spravochnik yangilandi: {len(sprav)} ta tovar, "
                f"{sum(1 for r in sprav if r[3])} tasida predstavitelstvo bor."
                if ok else "❌ Spravochnikni saqlashda xato (logni tekshiring)."
            )
            return VOZVRAT_UPLOAD
        if (doc.file_name or "").lower().endswith(".csv"):
            await msg.reply_text("⚠️ CSV faqat spravochnik uchun qabul qilinadi. Vozvrat ro'yxati uchun Excel yuboring.")
            return VOZVRAT_UPLOAD
        try:
            new = parse_vozvrat_excel(data)
        except Exception as e:
            await msg.reply_text(f"❌ Excel o'qilmadi: {e}")
            return VOZVRAT_UPLOAD
        rows.extend(new)
        ctx.user_data["vz_files"] = ctx.user_data.get("vz_files", 0) + 1
        await msg.reply_text(
            f"✅ Excel: {len(new)} ta qator o'qildi (jami {len(rows)}).\n"
            f"Yana yuboring yoki {BTN_TAYYOR} ni bosing."
        )
        return VOZVRAT_UPLOAD

    # ── Rasm (photo yoki rasm-fayl) ──
    file_id, mime = None, "image/jpeg"
    if msg.photo:
        file_id = msg.photo[-1].file_id
    elif doc and (doc.mime_type or "").startswith("image/"):
        file_id, mime = doc.file_id, doc.mime_type
    if not file_id:
        await msg.reply_text("⚠️ Faqat skrinshot (rasm) yoki Excel (.xlsx) qabul qilinadi.")
        return VOZVRAT_UPLOAD

    wait = await msg.reply_text("🔍 Skrinshot tahlil qilinmoqda...")
    try:
        f = await ctx.bot.get_file(file_id)
        img = bytes(await f.download_as_bytearray())
        if mime == "image/jpg":
            mime = "image/jpeg"
        new = await vision_extract_rows(img, mime)
    except RuntimeError as e:
        if str(e) == "NO_API_KEY":
            await wait.edit_text(
                "⚠️ Skrinshotni o'qish uchun ANTHROPIC_API_KEY sozlanmagan.\n"
                "Hozircha Excel fayl yuboring."
            )
        else:
            await wait.edit_text(f"❌ Tahlil xatosi: {e}")
        raise ApplicationHandlerStop(VOZVRAT_UPLOAD)
    except Exception as e:
        await wait.edit_text(f"❌ Tahlil xatosi: {e}")
        raise ApplicationHandlerStop(VOZVRAT_UPLOAD)

    rows.extend(new)
    ctx.user_data["vz_files"] = ctx.user_data.get("vz_files", 0) + 1
    filials = len({r["filial"] for r in new})
    await wait.edit_text(
        f"✅ Skrinshot #{ctx.user_data['vz_files']}: {len(new)} ta qator, {filials} ta filial "
        f"(jami {len(rows)} qator).\nYana yuboring yoki {BTN_TAYYOR} ni bosing."
    )
    # Boshqa (OCR) rasm handleri ishlamasligi uchun to'xtatamiz
    raise ApplicationHandlerStop(VOZVRAT_UPLOAD)


async def _vozvrat_analyze_and_preview(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    from attendance import run_read
    chat_id = update.effective_chat.id
    wait = await ctx.bot.send_message(chat_id, "⏳ Filiallar bo'yicha ajratilmoqda...")

    base = ctx.user_data.get("vz_base")
    if not base:
        base = await run_read(_load_pharmacies)
        ctx.user_data["vz_base"] = base
    plan = build_vozvrat_plan(ctx.user_data.get("vz_rows", []), base)
    ctx.user_data["vz_plan"] = plan
    ph = base["ph"]

    n_pos = sum(len(v) for v in plan["items"].values())
    no_tid = [f for f in plan["items"] if not ph[f].get("tid")]
    title = ctx.user_data.get("vz_title")
    lines = [
        "📦 *Vozvrat — tahlil natijasi*" + (f"\n🏢 {title}" if title else "") + "\n",
        f"🏥 Filiallar: *{len(plan['items'])}* ta",
        f"💊 Pozitsiyalar: *{n_pos}* ta",
    ]
    sklad = ctx.user_data.get("vz_sklad") or []
    if sklad:
        lines.append(f"🏬 Skladda: {_fmt_qty(sum(r[2] for r in sklad))} dona, "
                     f"{len(sklad)} ta partiya — xabar yuborilmaydi (faylda alohida varaqda)")
    if no_tid:
        lines.append(f"⚠️ TID yo'q (yuborilmaydi): {len(no_tid)} ta")
    lines.append("")
    for fno, drugs in list(plan["items"].items())[:40]:
        lines.append(f"• {_label(ph[fno])} — {len(drugs)} ta dori")
    if len(plan["items"]) > 40:
        lines.append(f"... va yana {len(plan['items']) - 40} ta (to'liq ro'yxat faylda)")

    if plan["fuzzy"]:
        lines.append("\n🔎 *O'xshashlik bo'yicha topildi — tekshiring:*")
        for raw, fno, why in plan["fuzzy"]:
            lines.append(f"• {raw} → {_label(ph[fno])} ({why})")
    if plan["unmatched"]:
        lines.append("\n❗ *Bazada topilmadi — yuborilmaydi:*")
        for raw, (why, drugs) in plan["unmatched"].items():
            lines.append(f"• {raw} — {len(drugs)} ta dori ({why})")
    if plan["fuzzy"] or plan["unmatched"]:
        lines.append(
            "\n✏️ Bog'lash yoki to'g'irlash uchun shunday yozing (har biri yangi qatordan):\n"
            "`АСКИЯ 1 = 0`\n`1 ГОР БОЛЬНИЦА (2) = 17`\n"
            "Bu eslab qolinadi — keyingi safar avtomatik topiladi."
        )

    # Batafsil ro'yxat fayli
    detail = []
    for fno, drugs in plan["items"].items():
        raw = ", ".join(sorted(plan["raw_names"].get(fno, [])))
        for dori, q in drugs.items():
            detail.append([fno, ph[fno]["nomi"], dori, q, raw, "✅ topildi"])
    for raw, (why, drugs) in plan["unmatched"].items():
        for dori, q in drugs.items():
            detail.append(["", "", dori, q, raw, f"⚠️ {why}"])
    extra = ("SKLAD", ["Dori", "Srok / seriya", "Soni", "", "Summa (prixod)", "Postavshik"],
             [[r[0], r[1], r[2], "", r[3], r[4]] for r in sklad]) if sklad else None
    buf = _report_xlsx([], ["Filial №", "Dorixona", "Dori", "Soni", "1C dagi nomi", "Holat"],
                       detail, "Vozvrat", extra=extra)

    text = "\n".join(lines)
    if len(text) > 4000:
        text = text[:3990] + "\n..."
    await wait.delete()
    from datetime import datetime as _dt
    await ctx.bot.send_document(
        chat_id, document=InputFile(buf, filename=f"vozvrat_tahlil_{_dt.now():%Y%m%d_%H%M}.xlsx"),
        caption="📋 Har bir filial bo'yicha to'liq ro'yxat (yuborishdan oldin tekshiring)",
    )
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton(f"✅ Yuborish ({len(plan['items']) - len(no_tid)})", callback_data="vz:send"),
        InlineKeyboardButton("❌ Bekor", callback_data="vz:cancel"),
    ]])
    try:
        await ctx.bot.send_message(chat_id, text, parse_mode="Markdown", reply_markup=kb)
    except Exception:
        await ctx.bot.send_message(chat_id, text.replace("*", "").replace("`", ""), reply_markup=kb)
    return VOZVRAT_CONFIRM


async def vozvrat_confirm_text_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Tasdiqlash bosqichida: alias yozilsa — qayta tahlil; Bekor — chiqish."""
    txt = (update.message.text or "").strip()
    if txt == BTN_BEKOR:
        await _back_to_main(ctx, update.effective_chat.id, "❌ Bekor qilindi.")
        from bot import MENU
        return MENU
    if await _apply_aliases(update, ctx, "vz_base"):
        return await _vozvrat_analyze_and_preview(update, ctx)
    await update.message.reply_text(
        "Yuqoridagi ✅ Yuborish / ❌ Bekor tugmasini bosing yoki `1C nomi = Filial №` ko'rinishida yozing.",
        parse_mode="Markdown",
    )
    return VOZVRAT_CONFIRM


async def vozvrat_confirm_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat_id
    from bot import MENU
    try:
        await query.edit_message_reply_markup(reply_markup=None)
    except Exception:
        pass

    if query.data == "vz:cancel":
        await _back_to_main(ctx, chat_id, "❌ Bekor qilindi.")
        return MENU

    plan = ctx.user_data.get("vz_plan")
    base = ctx.user_data.get("vz_base")
    if not plan or not base:
        await _back_to_main(ctx, chat_id, "⚠️ Ma'lumot topilmadi, qaytadan boshlang.")
        return MENU
    ph = base["ph"]

    # 1-qadam: "✅ Yuborish" → koordinatorga ham yuborilsinmi?
    if query.data == "vz:send":
        koords = {ph[f].get("koord_tid") for f in plan["items"] if ph[f].get("koord_tid")}
        await ctx.bot.send_message(
            chat_id,
            f"👤 Koordinatorlarga ham ushbu vozvrat ro'yxati yuborilsinmi?\n"
            f"(tanlangan filiallarda {len(koords)} ta koordinator bor)",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("✅ Ha", callback_data="vz:koord:ha"),
                InlineKeyboardButton("❌ Yo'q", callback_data="vz:koord:yoq"),
            ], [
                InlineKeyboardButton("🚫 Bekor qilish", callback_data="vz:cancel"),
            ]]),
        )
        return VOZVRAT_CONFIRM

    # 2-qadam: vz:koord:ha / vz:koord:yoq → yuborish
    koord_yuborish = query.data == "vz:koord:ha"
    try:
        await query.edit_message_text(
            query.message.text + ("\n\n✅ Ha — koordinatorlarga ham yuboriladi."
                                  if koord_yuborish else "\n\n❌ Faqat dorixonalarga yuboriladi.")
        )
    except Exception:
        pass
    sender = _sender_name(update.effective_user)
    _tline = f" — {ctx.user_data['vz_title']}" if ctx.user_data.get("vz_title") else ""

    messages, detail = [], []
    for fno, drugs in plan["items"].items():
        d = ph[fno]
        label = _label(d)
        dlines = [f"{i}. {dori} — {_fmt_qty(q)} ta" for i, (dori, q) in enumerate(drugs.items(), 1)]
        messages.append({
            "filial_no": fno, "label": label, "rol": "Vozvrat", "tid": d.get("tid", ""),
            "header": (f"📦 Skladga vozvrat{_tline}\n👤 Kimdan: {sender}\n\n🏥 {label}\n\n"
                       f"Quyidagi dorilarni skladga vozvrat qiling:\n\n"),
            "lines": dlines,
            "footer": "\n✅ Vozvrat qilingach «Bajarildi» tugmasini bosing.",
            "log_text": f"📦 Vozvrat ({len(drugs)} ta dori): " + "; ".join(
                f"{k} — {_fmt_qty(v)}" for k, v in drugs.items()),
        })
        for dori, q in drugs.items():
            detail.append([fno, d["nomi"], dori, q, "", ""])

    progress = await ctx.bot.send_message(chat_id, f"⏳ Yuborilmoqda... 0 / {len(messages)}")
    results = await _send_messages(ctx, chat_id, messages, progress)

    # ── Koordinatorlarga: har biriga faqat o'ziga tegishli filiallar ro'yxati ──
    if koord_yuborish:
        status = {r[0]: r[3] for r in results}   # filial_no → ✅/❌/⚠️
        by_koord = OrderedDict()
        for fno in plan["items"]:
            ktid = ph[fno].get("koord_tid")
            if ktid:
                by_koord.setdefault(ktid, []).append(fno)
        sent_k = failed_k = 0
        for ktid, fnos in by_koord.items():
            ismi = next((ph[f].get("koord_ismi") for f in fnos if ph[f].get("koord_ismi")), "") or "Koordinator"
            klines = []
            for fno in fnos:
                st = status.get(fno, "")
                note = "" if st == "✅" else ("  ⚠️ zavga yetmadi (TID yo'q)" if st == "⚠️"
                                              else "  ❌ zavga yuborilmadi")
                klines.append(f"\n🏥 {_label(ph[fno])}{note}")
                for i, (dori, q) in enumerate(plan["items"][fno].items(), 1):
                    klines.append(f"  {i}. {dori} — {_fmt_qty(q)} ta")
            header = (f"📦 Skladga vozvrat{_tline}\n👤 Kimdan: {sender}\n\n"
                      f"{ismi}, sizga tegishli {len(fnos)} ta dorixonaga vozvrat ro'yxati yuborildi:\n")
            try:
                for part in _chunks(header, klines, "\nℹ️ Nazorat qilib boring."):
                    await ctx.bot.send_message(chat_id=int(ktid), text=part)
                sent_k += 1
            except Exception as e:
                failed_k += 1
                logger.warning(f"[VOZVRAT KOORD] {ktid} ga yuborilmadi: {e}")
            await asyncio.sleep(0.05)
        await ctx.bot.send_message(
            chat_id,
            (f"👤 Koordinatorlarga: ✅ {sent_k} ta yuborildi" + (f", ❌ {failed_k} ta yuborilmadi" if failed_k else ""))
            if by_koord else "⚠️ Tanlangan dorixonalarda koordinator TID topilmadi.",
        )

    for raw, (why, drugs) in plan["unmatched"].items():
        results.append(("", raw, "Vozvrat", "⚠️", f"Bazada topilmadi ({why})", len(drugs)))
        for dori, q in drugs.items():
            detail.append(["", raw, dori, q, "", "⚠️ topilmadi"])

    ok = sum(1 for r in results if r[3] == "✅")
    bad = sum(1 for r in results if r[3] == "❌")
    warn = sum(1 for r in results if r[3] == "⚠️")
    summary = (f"📊 *Vozvrat yuborildi*\n\n✅ Yuborildi: *{ok}* ta\n"
               f"❌ Yuborilmadi: *{bad}* ta\n⚠️ TID yo'q / topilmadi: *{warn}* ta")
    try:
        await progress.edit_text(summary, parse_mode="Markdown")
    except Exception:
        pass
    from datetime import datetime as _dt
    buf = _report_xlsx(results, ["Filial №", "Dorixona", "Dori", "Soni", "", ""], detail, "Vozvrat")
    await ctx.bot.send_document(
        chat_id, document=InputFile(buf, filename=f"vozvrat_hisobot_{_dt.now():%Y%m%d_%H%M}.xlsx"),
        caption=summary, parse_mode="Markdown",
    )
    await _cv_mark_done(ctx, chat_id, summary)
    await _back_to_main(ctx, chat_id)
    return MENU


# ═══════════════════════════════════════════════════════════════════════════════
# 6b) 📦 VOZVRAT — OSTATOK'dan avtomatik (postavshik / predstavitelstvo bo'yicha)
# ═══════════════════════════════════════════════════════════════════════════════
#
# OSTATOK varag'i (Excel makrosi yuklaydi) ustunlari:
#   A Поставщик | B Филиал | C Наименование | D Производитель | E Остаток | F Серия |
#   G Тип наценки | H Срок годности | ... | N Сумма приходная
# Varaq juda katta (~340 ming qator), shuning uchun to'liq o'qilmaydi — Google
# Visualization (gviz) so'rovi bilan faqat kerakli qatorlar olinadi.
#
# Spravochnik varag'i (admin «итоговый справочник» faylini botga yuborganda yoziladi):
#   A Код | B Наименование | C Производитель | D Представительство |
#   E Ответственный снабженец | F Снабженцы представительства

VZ_SRC_SEARCH = 595   # postavshik/predstavitelstvo nomini qidirish
VZ_SRC_PICK   = 596   # topilganlardan tanlash + partiya filtri
VZ_SRC_DATE   = 597   # "muddati yaqin" uchun sana

BTN_BY_POST = "🏢 Поставщик бўйича"
BTN_BY_PRED = "🏷 Представительство бўйича"
SPRAV_WS = "Spravochnik"
OSTATOK_WS = "OSTATOK"
SPRAV_HEADERS = ["Код", "Наименование", "Производитель", "Представительство",
                 "Ответственный снабженец", "Снабженцы представительства"]


def _ostatok_sheet_id() -> str:
    from salary_handlers import PHARMACY_SHEET_ID
    return os.getenv("OSTATOK_SHEET_ID", "").strip() or PHARMACY_SHEET_ID


_GV_SESSION = None


def _gv_session():
    """Service account bilan avtorizatsiyalangan HTTP sessiya (gviz so'rovlari uchun)."""
    global _GV_SESSION
    if _GV_SESSION is None:
        from google.auth.transport.requests import AuthorizedSession
        from google.oauth2.service_account import Credentials
        from salary_handlers import SCOPES
        info = os.getenv("GOOGLE_CREDENTIALS")
        creds = (Credentials.from_service_account_info(json.loads(info), scopes=SCOPES) if info
                 else Credentials.from_service_account_file("credentials.json", scopes=SCOPES))
        _GV_SESSION = AuthorizedSession(creds)
    return _GV_SESSION


def _q(s: str) -> str:
    """gviz so'rovi uchun matnli literal."""
    s = str(s)
    return f'"{s}"' if "'" in s else f"'{s}'"


def gviz_query(sheet: str, rng: str, tq: str) -> list:
    """Google Sheets'dan gviz so'rovi bilan qatorlarni oladi (sarlavhasiz, matn ko'rinishida)."""
    import csv as _csv
    url = f"https://docs.google.com/spreadsheets/d/{_ostatok_sheet_id()}/gviz/tq"
    r = _gv_session().get(url, params={"sheet": sheet, "range": rng, "headers": "1",
                                       "tqx": "out:csv", "tq": tq}, timeout=180)
    r.raise_for_status()
    text = r.content.decode("utf-8-sig", errors="replace")
    if text.lstrip().startswith("<"):
        raise RuntimeError("gviz so'rovi bajarilmadi (varaq yoki ustun topilmadi)")
    rows = list(_csv.reader(io.StringIO(text)))
    return rows[1:] if rows else []


def _parse_date(s):
    import datetime as _d
    s = str(s or "").strip()
    m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})", s)
    if m:
        return _d.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return _d.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def _num(s) -> float:
    v = _to_num(str(s or "").replace("\xa0", "").replace(" ", ""))
    return v or 0.0


# ── Spravochnik ─────────────────────────────────────────────────────────────

def parse_spravochnik(data: bytes, fname: str):
    """«итоговый справочник» (.csv/.xlsx) → [[kod, nomi, ishlab chiqaruvchi, predst., snabjenets, predst. snabjenetslari]].
    Fayl spravochnik bo'lmasa None qaytaradi."""
    import csv as _csv
    if fname.lower().endswith(".csv"):
        rows = list(_csv.reader(io.StringIO(data.decode("utf-8-sig", errors="replace"))))
    else:
        rows = _read_xlsx_sheets(data)[0]
    hi, cols = _find_header(rows, {
        "kod": [r"^код"], "nomi": [r"^наименование"], "proizv": [r"производител"],
        "pred": [r"^представительство"], "otv": [r"ответственный"], "snab": [r"снабженцы представ"],
    }, optional=("kod", "proizv", "otv", "snab"), scan=5)
    if hi is None:
        return None
    out = []
    for r in rows[hi + 1:]:
        g = lambda k: str(_cell(r, cols.get(k)) or "").strip()
        if g("nomi"):
            out.append([g("kod"), g("nomi"), g("proizv"), g("pred"), g("otv"), g("snab")])
    return out


def save_spravochnik(rows: list) -> bool:
    """Spravochnikni OSTATOK turgan jadvaldagi "Spravochnik" varag'iga yozadi (eskisini almashtiradi)."""
    from salary_handlers import _get_client
    try:
        sh = _get_client().open_by_key(_ostatok_sheet_id())
        try:
            ws = sh.worksheet(SPRAV_WS)
        except Exception:
            ws = sh.add_worksheet(title=SPRAV_WS, rows=len(rows) + 10, cols=len(SPRAV_HEADERS))
        ws.clear()
        ws.resize(rows=len(rows) + 10, cols=len(SPRAV_HEADERS))
        data = [SPRAV_HEADERS] + rows
        step = 5000
        for i in range(0, len(data), step):
            ws.update(range_name=f"A{i + 1}", values=data[i:i + step], value_input_option="RAW")
        return True
    except Exception as e:
        logger.error(f"[SPRAV] Saqlashda xato: {e}")
        return False


# ── Qidirish va ma'lumot olish ──────────────────────────────────────────────

def search_sources(mode: str, term: str) -> list:
    """Nomi bo'yicha postavshik (OSTATOK!A) yoki predstavitelstvo (Spravochnik!D) qidiradi.
    Qaytaradi: [(nomi, soni), ...] — eng ko'p qatorlilari oldin."""
    t = term.strip().lower().replace("'", "").replace('"', "")
    if mode == "post":
        rows = gviz_query(OSTATOK_WS, "A2:P",
                          f"select A, count(C) where E > 0 and lower(A) contains {_q(t)} "
                          f"group by A order by count(C) desc")
    else:
        rows = gviz_query(SPRAV_WS, "A1:F",
                          f"select D, count(B) where lower(D) contains {_q(t)} "
                          f"group by D order by count(B) desc")
    return [(r[0], int(_num(r[1]))) for r in rows if r and r[0].strip()][:20]


def fetch_source_rows(mode: str, name: str) -> list:
    """Tanlangan postavshik/predstavitelstvo bo'yicha OSTATOK qatorlari:
    [{filial, dori, seriya, srok(date|None), srok_txt, soni, summa, post}]"""
    cols = "B, C, F, H, E, N, A"
    if mode == "post":
        raw = gviz_query(OSTATOK_WS, "A2:P", f"select {cols} where E > 0 and A = {_q(name)}")
    else:
        names = [r[0] for r in gviz_query(SPRAV_WS, "A1:F", f"select B where D = {_q(name)}") if r and r[0]]
        raw = []
        for i in range(0, len(names), 20):
            cond = " or ".join(f"C = {_q(n)}" for n in names[i:i + 20])
            raw += gviz_query(OSTATOK_WS, "A2:P", f"select {cols} where E > 0 and ({cond})")
    out = []
    for r in raw:
        if len(r) < 7:
            continue
        q = _num(r[4])
        if q < 0.01:
            continue
        out.append({"filial": r[0].strip(), "dori": re.sub(r"\s+", " ", r[1]).strip(),
                    "seriya": r[2].strip(), "srok": _parse_date(r[3]), "srok_txt": r[3].strip(),
                    "soni": q, "summa": _num(r[5]), "post": r[6].strip()})
    return out


def _ostatok_date_text() -> str:
    try:
        rows = gviz_query(OSTATOK_WS, "A1:A1", "select A")
        return rows[0][0] if rows and rows[0] else ""
    except Exception:
        return ""


# ── Handlerlar ──────────────────────────────────────────────────────────────

async def vz_src_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE, mode: str):
    ctx.user_data["vz_src"] = mode
    what = "postavshik" if mode == "post" else "predstavitelstvo"
    await update.message.reply_text(
        f"🔎 {what.capitalize()} nomining bir qismini yozing (masalan: `эмир`):",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup([[BTN_BEKOR]], resize_keyboard=True),
    )
    return VZ_SRC_SEARCH


async def vz_src_search_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    from attendance import run_read
    txt = (update.message.text or "").strip()
    if txt == BTN_BEKOR:
        await _back_to_main(ctx, update.effective_chat.id, "❌ Bekor qilindi.")
        from bot import MENU
        return MENU
    if len(txt) < 2:
        await update.message.reply_text("Kamida 2 ta harf yozing.")
        return VZ_SRC_SEARCH
    mode = ctx.user_data.get("vz_src", "post")
    wait = await update.message.reply_text("⏳ Qidirilmoqda...")
    try:
        found = await run_read(search_sources, mode, txt)
    except Exception as e:
        logger.warning(f"[VZ_SRC] qidiruv xatosi: {e}")
        if mode == "pred":
            await wait.edit_text("⚠️ Spravochnik topilmadi yoki o'qib bo'lmadi.\n"
                                 "«итоговый справочник» faylini (.csv) shu yerga yuboring, so'ng qayta urinib ko'ring.")
            return VOZVRAT_UPLOAD
        await wait.edit_text(f"❌ OSTATOK'dan o'qishda xato: {e}")
        return VZ_SRC_SEARCH
    if not found:
        await wait.edit_text("🤷 Hech narsa topilmadi. Boshqacha yozib ko'ring:")
        return VZ_SRC_SEARCH
    ctx.user_data["vz_src_opts"] = found
    unit = "partiya" if mode == "post" else "tovar"
    kb = [[InlineKeyboardButton(f"{n} ({c} {unit})", callback_data=f"vzs:{i}")] for i, (n, c) in enumerate(found)]
    await wait.edit_text("👇 Keraklisini tanlang (yoki boshqa nom yozing):", reply_markup=InlineKeyboardMarkup(kb))
    return VZ_SRC_PICK


def _filter_kb():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📦 Hammasi", callback_data="vzf:all")],
        [InlineKeyboardButton("⛔ Muddati o'tganlar", callback_data="vzf:exp")],
        [InlineKeyboardButton("⏳ Muddati yaqinlar (sana kiritiladi)", callback_data="vzf:near")],
        [InlineKeyboardButton("❌ Bekor", callback_data="vzf:cancel")],
    ])


async def vz_src_pick_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat_id
    data = query.data

    if data.startswith("vzs:"):
        opts = ctx.user_data.get("vz_src_opts") or []
        i = int(data.split(":")[1])
        if i >= len(opts):
            return VZ_SRC_PICK
        ctx.user_data["vz_src_pick"] = opts[i][0]
        try:
            await query.edit_message_text(
                f"✅ Tanlandi: *{opts[i][0]}*\n\nQaysi partiyalar qaytarilsin?",
                parse_mode="Markdown", reply_markup=_filter_kb(),
            )
        except Exception:
            await query.edit_message_text(
                f"✅ Tanlandi: {opts[i][0]}\n\nQaysi partiyalar qaytarilsin?", reply_markup=_filter_kb(),
            )
        return VZ_SRC_PICK

    kind = data.split(":")[1]
    if kind == "cancel":
        await query.edit_message_reply_markup(reply_markup=None)
        await _back_to_main(ctx, chat_id, "❌ Bekor qilindi.")
        from bot import MENU
        return MENU
    if kind == "near":
        await query.edit_message_reply_markup(reply_markup=None)
        await ctx.bot.send_message(
            chat_id, "📅 Qaysi sanagacha tugaydiganlar olinsin?\n"
                     "Sana yozing (`01.03.2027`) yoki oylar sonini (`3` — 3 oy ichida):",
            parse_mode="Markdown",
        )
        return VZ_SRC_DATE
    await query.edit_message_reply_markup(reply_markup=None)
    return await _vz_src_build(update, ctx, kind, None)


async def vz_src_date_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    import datetime as _d
    txt = (update.message.text or "").strip()
    if txt == BTN_BEKOR:
        await _back_to_main(ctx, update.effective_chat.id, "❌ Bekor qilindi.")
        from bot import MENU
        return MENU
    until = _parse_date(txt)
    if until is None and txt.isdigit() and 0 < int(txt) <= 36:
        # N oy ichida: masalan 24.11.2026 + 3 oy → 24.02.2027 gacha
        t = _d.date.today()
        m0 = t.month - 1 + int(txt)
        y, m = t.year + m0 // 12, m0 % 12 + 1
        import calendar
        until = _d.date(y, m, min(t.day, calendar.monthrange(y, m)[1]))
    if until is None:
        await update.message.reply_text("⚠️ Tushunmadim. Masalan: `01.03.2027` yoki `3`", parse_mode="Markdown")
        return VZ_SRC_DATE
    return await _vz_src_build(update, ctx, "near", until)


async def _vz_src_build(update: Update, ctx: ContextTypes.DEFAULT_TYPE, kind: str, until):
    """OSTATOK'dan qatorlarni olib, vozvrat rejasiga aylantiradi va oldindan ko'rsatadi."""
    import datetime as _d
    from attendance import run_read
    chat_id = update.effective_chat.id
    mode = ctx.user_data.get("vz_src", "post")
    name = ctx.user_data.get("vz_src_pick", "")
    wait = await ctx.bot.send_message(chat_id, f"⏳ OSTATOK'dan «{name}» qoldiqlari olinmoqda...")
    try:
        rows = await run_read(fetch_source_rows, mode, name)
        ost_date = await run_read(_ostatok_date_text)
    except Exception as e:
        await wait.edit_text(f"❌ OSTATOK'dan o'qishda xato: {e}")
        return VZ_SRC_PICK

    today = _d.date.today()
    if kind == "exp":
        rows = [r for r in rows if r["srok"] and r["srok"] < today]
        ftxt = "muddati o'tganlar"
    elif kind == "near":
        rows = [r for r in rows if r["srok"] and r["srok"] <= until]
        ftxt = f"{until:%d.%m.%Y} gacha tugaydiganlar"
    else:
        ftxt = "hammasi"

    vz_rows, sklad = [], []
    for r in rows:
        srok = r["srok"].strftime("%d.%m.%Y") if r["srok"] else r["srok_txt"]
        dori = f"{r['dori']} (срок {srok})" if srok else r["dori"]
        if _nm(r["filial"]).startswith("СКЛАД"):
            sklad.append([r["dori"], f"{srok} / {r['seriya']}".strip(" /"), r["soni"], r["summa"], r["post"]])
        else:
            vz_rows.append({"filial": r["filial"], "dori": dori, "soni": r["soni"]})

    label = "Postavshik" if mode == "post" else "Predstavitelstvo"
    ctx.user_data["vz_title"] = f"{label}: {name} ({ftxt})"
    ctx.user_data["vz_sklad"] = sklad
    ctx.user_data["vz_rows"] = vz_rows
    try:
        await wait.edit_text(
            f"📊 {ost_date or 'OSTATOK'}\n"
            f"Topildi: aptekalarda {len(vz_rows)} ta qator, skladda {len(sklad)} ta partiya."
        )
    except Exception:
        pass
    if not vz_rows:
        await ctx.bot.send_message(
            chat_id, "🤷 Aptekalarda bu shartga mos qoldiq yo'q — yuboriladigan narsa yo'q."
            + (f"\n🏬 Skladda: {_fmt_qty(sum(r[2] for r in sklad))} dona." if sklad else "")
        )
        await _back_to_main(ctx, chat_id)
        from bot import MENU
        return MENU
    return await _vozvrat_analyze_and_preview(update, ctx)


# ═══════════════════════════════════════════════════════════════════════════════
# 7) 🔄 PEREMESHENIE
# ═══════════════════════════════════════════════════════════════════════════════

async def peremesh_upload_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    from bot import MENU
    if msg.text:
        if msg.text.strip() == BTN_BEKOR:
            await _back_to_main(ctx, msg.chat_id, "❌ Bekor qilindi.")
            return MENU
        await msg.reply_text("📎 Peremeshenie Excel faylini (.xlsx) yuboring.")
        return PEREMESH_UPLOAD

    doc = msg.document
    if not doc or not (doc.file_name or "").lower().endswith((".xlsx", ".xlsm")):
        await msg.reply_text("⚠️ Faqat Excel (.xlsx) fayl qabul qilinadi.")
        if msg.photo:
            raise ApplicationHandlerStop(PEREMESH_UPLOAD)
        return PEREMESH_UPLOAD
    try:
        f = await ctx.bot.get_file(doc.file_id)
        rows = parse_peremesh_excel(bytes(await f.download_as_bytearray()))
    except Exception as e:
        await msg.reply_text(f"❌ Excel o'qilmadi: {e}")
        return PEREMESH_UPLOAD
    if not rows:
        await msg.reply_text("⚠️ Faylda peremeshenie qatorlari topilmadi.")
        return PEREMESH_UPLOAD
    ctx.user_data["pm_rows"] = rows
    return await _peremesh_analyze_and_preview(update, ctx)


async def _peremesh_analyze_and_preview(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    from attendance import run_read
    chat_id = update.effective_chat.id
    base = ctx.user_data.get("pm_base")
    if not base:
        base = await run_read(_load_pharmacies)
        ctx.user_data["pm_base"] = base
    plan = build_peremesh_plan(ctx.user_data.get("pm_rows", []), base)
    ctx.user_data["pm_plan"] = plan
    ph = base["ph"]

    lines = [
        "🔄 *Peremeshenie — tahlil natijasi*" + (f"\n🏢 {ctx.user_data['pm_title']}" if ctx.user_data.get("pm_title") else "") + "\n",
        f"📄 Qatorlar: *{len(plan['rows'])}* ta",
        f"📤 Beruvchi filiallar: *{len(plan['out'])}* ta",
        f"📥 Qabul qiluvchi filiallar: *{len(plan['inc'])}* ta\n",
        "📤 *Beruvchilar:*",
    ]
    for fno, lst in list(plan["out"].items())[:30]:
        lines.append(f"• {_label(ph[fno])} — {len(lst)} ta pozitsiya")
    lines.append("\n📥 *Qabul qiluvchilar:*")
    for fno, lst in list(plan["inc"].items())[:30]:
        lines.append(f"• {_label(ph[fno])} — {len(lst)} ta pozitsiya")
    if plan["fuzzy"]:
        lines.append("\n🔎 *O'xshashlik bo'yicha topildi — tekshiring:*")
        for raw, fno, why in plan["fuzzy"]:
            lines.append(f"• {raw} → {_label(ph[fno])} ({why})")
    if plan["unmatched"]:
        lines.append("\n❗ *Bazada topilmadi — bu filialga xabar bormaydi:*")
        for raw, why in plan["unmatched"].items():
            lines.append(f"• {raw} ({why})")
    if plan["fuzzy"] or plan["unmatched"]:
        lines.append("\n✏️ To'g'irlash uchun yozing: `1C nomi = Filial №`")

    text = "\n".join(lines)
    if len(text) > 4000:
        text = text[:3990] + "\n..."
    n_send = len(plan["out"]) + len(plan["inc"])
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton(f"✅ Yuborish ({n_send} xabar)", callback_data="pm:send"),
        InlineKeyboardButton("❌ Bekor", callback_data="pm:cancel"),
    ]])
    try:
        await ctx.bot.send_message(chat_id, text, parse_mode="Markdown", reply_markup=kb)
    except Exception:
        await ctx.bot.send_message(chat_id, text.replace("*", "").replace("`", ""), reply_markup=kb)
    return PEREMESH_CONFIRM


async def peremesh_confirm_text_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    txt = (update.message.text or "").strip()
    if txt == BTN_BEKOR:
        await _back_to_main(ctx, update.effective_chat.id, "❌ Bekor qilindi.")
        from bot import MENU
        return MENU
    if await _apply_aliases(update, ctx, "pm_base"):
        return await _peremesh_analyze_and_preview(update, ctx)
    await update.message.reply_text(
        "Yuqoridagi ✅ Yuborish / ❌ Bekor tugmasini bosing yoki `1C nomi = Filial №` ko'rinishida yozing.",
        parse_mode="Markdown",
    )
    return PEREMESH_CONFIRM


async def peremesh_confirm_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat_id
    from bot import MENU
    try:
        await query.edit_message_reply_markup(reply_markup=None)
    except Exception:
        pass

    if query.data == "pm:cancel":
        await _back_to_main(ctx, chat_id, "❌ Bekor qilindi.")
        return MENU

    plan = ctx.user_data.get("pm_plan")
    base = ctx.user_data.get("pm_base")
    if not plan or not base:
        await _back_to_main(ctx, chat_id, "⚠️ Ma'lumot topilmadi, qaytadan boshlang.")
        return MENU
    ph = base["ph"]
    sender = _sender_name(update.effective_user)
    _ptl = f" — {ctx.user_data['pm_title']}" if ctx.user_data.get("pm_title") else ""

    messages, detail = [], []
    # Beruvchi filiallar: "shu dorini shu aptekaga peremeshenie qiling"
    for fno, lst in plan["out"].items():
        d = ph[fno]
        label = _label(d)
        dlines = []
        for i, (dori, q, to_lbl) in enumerate(lst, 1):
            dlines.append(f"{i}. {dori} — {_fmt_qty(q)} up.\n    ➡️ {to_lbl} ga")
            detail.append([fno, d["nomi"], dori, q, to_lbl, "Beradi"])
        messages.append({
            "filial_no": fno, "label": label, "rol": "Beruvchi 📤", "tid": d.get("tid", ""),
            "header": (f"🔄 Peremeshenie{_ptl}\n👤 Kimdan: {sender}\n\n🏥 {label}\n\n"
                       f"Quyidagi dorilarni ko'rsatilgan filiallarga peremeshenie qiling:\n\n"),
            "lines": dlines,
            "footer": "\n✅ Jo'natgach «Bajarildi» tugmasini bosing.",
            "log_text": "🔄 Peremeshenie (beradi): " + "; ".join(
                f"{a} — {_fmt_qty(b)} → {c}" for a, b, c in lst),
        })
    # Qabul qiluvchi filiallar: "shu dori sizga N-filialdan kelyapti"
    for fno, lst in plan["inc"].items():
        d = ph[fno]
        label = _label(d)
        dlines = []
        for i, (dori, q, from_lbl) in enumerate(lst, 1):
            dlines.append(f"{i}. {dori} — {_fmt_qty(q)} up.\n    ⬅️ {from_lbl} dan")
            detail.append([fno, d["nomi"], dori, q, from_lbl, "Qabul qiladi"])
        messages.append({
            "filial_no": fno, "label": label, "rol": "Qabul qiluvchi 📥", "tid": d.get("tid", ""),
            "header": (f"🔄 Peremeshenie{_ptl}\n👤 Kimdan: {sender}\n\n🏥 {label}\n\n"
                       f"Sizga quyidagi dorilar peremeshenie orqali keladi:\n\n"),
            "lines": dlines,
            "footer": "\n✅ Qabul qilgach «Bajarildi» tugmasini bosing.",
            "log_text": "🔄 Peremeshenie (qabul): " + "; ".join(
                f"{a} — {_fmt_qty(b)} ← {c}" for a, b, c in lst),
        })

    progress = await ctx.bot.send_message(chat_id, f"⏳ Yuborilmoqda... 0 / {len(messages)}")
    results = await _send_messages(ctx, chat_id, messages, progress)
    for raw, why in plan["unmatched"].items():
        results.append(("", raw, "—", "⚠️", f"Bazada topilmadi ({why})", 0))

    ok = sum(1 for r in results if r[3] == "✅")
    bad = sum(1 for r in results if r[3] == "❌")
    warn = sum(1 for r in results if r[3] == "⚠️")
    summary = (f"📊 *Peremeshenie yuborildi*\n\n✅ Yuborildi: *{ok}* ta\n"
               f"❌ Yuborilmadi: *{bad}* ta\n⚠️ TID yo'q / topilmadi: *{warn}* ta")
    try:
        await progress.edit_text(summary, parse_mode="Markdown")
    except Exception:
        pass
    from datetime import datetime as _dt
    buf = _report_xlsx(results, ["Filial №", "Dorixona", "Dori", "Soni", "Qayerga / Qayerdan", "Rol"],
                       detail, "Peremeshenie")
    await ctx.bot.send_document(
        chat_id, document=InputFile(buf, filename=f"peremeshenie_hisobot_{_dt.now():%Y%m%d_%H%M}.xlsx"),
        caption=summary, parse_mode="Markdown",
    )
    await _cv_mark_done(ctx, chat_id, summary)
    await _back_to_main(ctx, chat_id)
    return MENU


# ═══════════════════════════════════════════════════════════════════════════════
# 7b) 📥 TAYYOR TOPSHIRIQLAR — Claude tayyorlagan ro'yxatlar
# ═══════════════════════════════════════════════════════════════════════════════
#
# "Claude_vazifa" varag'i (OSTATOK turgan jadvalda). Har bir qator — bitta pozitsiya:
#   A Topshiriq ID | B Turi (vozvrat / peremesh) | C Sarlavha | D Filial (qayerdan) |
#   E Dori | F Soni | G Qayerga (faqat peremesh) | H Izoh | I Holat | J Yaratildi
# Holat bo'sh yoki "yangi" — hali yuborilmagan. Bot yuborgach "yuborildi ..." yozadi.
# Filial nomi 1C dagidek yoki filial raqami bilan ("48" / "48 - ШАХРИСТАНСКИЙ") yozilishi mumkin.

CV_PICK = 598
CV_WS = "Claude_vazifa"
CV_HEADERS = ["Topshiriq ID", "Turi", "Sarlavha", "Filial (qayerdan)", "Dori", "Soni",
              "Qayerga (peremesh)", "Izoh", "Holat", "Yaratildi"]


def _cv_ws():
    from salary_handlers import _get_client
    sh = _get_client().open_by_key(_ostatok_sheet_id())
    try:
        return sh.worksheet(CV_WS)
    except Exception:
        ws = sh.add_worksheet(title=CV_WS, rows=1000, cols=len(CV_HEADERS))
        ws.append_row(CV_HEADERS)
        return ws


def cv_load_tasks() -> "OrderedDict":
    """Yuborilmagan topshiriqlar: {id: {"turi", "sarlavha", "rows": [(row_no, filial, dori, soni, ga, izoh)]}}"""
    vals = _cv_ws().get_all_values()
    tasks = OrderedDict()
    for i, r in enumerate(vals[1:], start=2):
        r = (r + [""] * len(CV_HEADERS))[:len(CV_HEADERS)]
        tid, turi, sarl, fil, dori, soni, ga, izoh, holat = [x.strip() for x in r[:9]]
        if not tid or not dori or (holat and holat.lower() != "yangi"):
            continue
        q = _to_num(soni)
        if not q:
            continue
        t = tasks.setdefault(tid, {"turi": turi.lower(), "sarlavha": sarl, "rows": []})
        if sarl and not t["sarlavha"]:
            t["sarlavha"] = sarl
        t["rows"].append((i, fil, dori, q, ga, izoh))
    return tasks


def cv_mark(row_nos: list, text: str) -> bool:
    try:
        ws = _cv_ws()
        ws.batch_update([{"range": f"I{n}", "values": [[text]]} for n in row_nos])
        return True
    except Exception as e:
        logger.error(f"[CV] Holatni yozishda xato: {e}")
        return False


async def cv_list_tasks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    from attendance import run_read
    wait = await update.message.reply_text("⏳ Tayyor topshiriqlar o'qilmoqda...")
    try:
        tasks = await run_read(cv_load_tasks)
    except Exception as e:
        await wait.edit_text(f"❌ «{CV_WS}» varag'ini o'qib bo'lmadi: {e}")
        return APPEAL_DORIXONA_TYPE
    if not tasks:
        await wait.edit_text("📭 Hozircha yangi topshiriq yo'q.\n"
                             "Claude'dan so'rang — u ro'yxatni «Claude_vazifa» varag'iga yozadi.")
        return APPEAL_DORIXONA_TYPE
    ctx.user_data["cv_tasks"] = tasks
    kb, lines = [], ["📥 *Tayyor topshiriqlar:*\n"]
    for n, (tid, t) in enumerate(tasks.items()):
        icon = "🔄" if t["turi"].startswith("peremesh") else "📦"
        fil = len({r[1] for r in t["rows"]})
        lines.append(f"{icon} *{tid}* — {t['sarlavha'] or t['turi']}\n    {len(t['rows'])} ta pozitsiya, {fil} ta filial")
        kb.append([InlineKeyboardButton(f"{icon} {tid}", callback_data=f"cvt:{n}")])
    kb.append([InlineKeyboardButton("❌ Bekor", callback_data="cvt:x")])
    text = "\n".join(lines)
    try:
        await wait.edit_text(text[:4000], parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(kb))
    except Exception:
        await wait.edit_text(text[:4000].replace("*", ""), reply_markup=InlineKeyboardMarkup(kb))
    return CV_PICK


async def cv_pick_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat_id
    try:
        await query.edit_message_reply_markup(reply_markup=None)
    except Exception:
        pass
    key = query.data.split(":")[1]
    tasks = ctx.user_data.get("cv_tasks") or OrderedDict()
    if key == "x" or not tasks:
        await _back_to_main(ctx, chat_id, "❌ Bekor qilindi.")
        from bot import MENU
        return MENU
    tid = list(tasks.keys())[int(key)]
    t = tasks[tid]
    ctx.user_data["cv_task_id"] = (tid, [r[0] for r in t["rows"]])
    title = f"{tid}: {t['sarlavha']}" if t["sarlavha"] else tid

    if t["turi"].startswith("peremesh"):
        bad = [r for r in t["rows"] if not r[4]]
        if bad:
            await ctx.bot.send_message(chat_id, f"⚠️ {len(bad)} ta qatorda «Qayerga» bo'sh — ular tashlab ketildi.")
        ctx.user_data["pm_rows"] = [{"dan": r[1], "ga": r[4], "dori": r[2], "soni": r[3]}
                                    for r in t["rows"] if r[4]]
        ctx.user_data["pm_title"] = title
        ctx.user_data.pop("pm_base", None)
        return await _peremesh_analyze_and_preview(update, ctx)

    ctx.user_data["vz_rows"] = [{"filial": r[1], "dori": r[2], "soni": r[3]} for r in t["rows"]]
    ctx.user_data["vz_title"] = title
    ctx.user_data.pop("vz_sklad", None)
    ctx.user_data.pop("vz_base", None)
    return await _vozvrat_analyze_and_preview(update, ctx)


async def _cv_mark_done(ctx, chat_id, summary: str = ""):
    """Topshiriq yuborilgach — Claude_vazifa varag'ida "yuborildi" deb belgilaydi."""
    task = ctx.user_data.get("cv_task_id")
    if not task:
        return
    from attendance import run_write
    import datetime as _d
    tid, row_nos = task
    ok = await run_write(cv_mark, row_nos, f"yuborildi {_d.datetime.now():%d.%m.%Y %H:%M}")
    if not ok:
        await ctx.bot.send_message(chat_id, f"⚠️ «{tid}» holatini varaqqa yozib bo'lmadi — qayta yuborilmasligi uchun "
                                            f"{CV_WS} varag'ida Holat ustuniga qo'lda «yuborildi» yozing.")


# ═══════════════════════════════════════════════════════════════════════════════
# 8) ConversationHandler holatlari (salary_handlers.get_sal_states ga qo'shiladi)
# ═══════════════════════════════════════════════════════════════════════════════

def get_vozvrat_states() -> dict:
    return {
        APPEAL_DORIXONA_TYPE: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, dorixona_type_handler),
        ],
        VOZVRAT_UPLOAD: [
            MessageHandler(
                (filters.PHOTO | filters.Document.ALL | filters.TEXT) & ~filters.COMMAND,
                vozvrat_upload_handler,
            ),
        ],
        VOZVRAT_CONFIRM: [
            CallbackQueryHandler(vozvrat_confirm_callback, pattern="^vz:"),
            MessageHandler(filters.TEXT & ~filters.COMMAND, vozvrat_confirm_text_handler),
        ],
        VZ_SRC_SEARCH: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, vz_src_search_handler),
        ],
        VZ_SRC_PICK: [
            CallbackQueryHandler(vz_src_pick_callback, pattern="^vz[sf]:"),
            MessageHandler(filters.TEXT & ~filters.COMMAND, vz_src_search_handler),
        ],
        VZ_SRC_DATE: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, vz_src_date_handler),
        ],
        CV_PICK: [
            CallbackQueryHandler(cv_pick_callback, pattern="^cvt:"),
        ],
        PEREMESH_UPLOAD: [
            MessageHandler(
                (filters.PHOTO | filters.Document.ALL | filters.TEXT) & ~filters.COMMAND,
                peremesh_upload_handler,
            ),
        ],
        PEREMESH_CONFIRM: [
            CallbackQueryHandler(peremesh_confirm_callback, pattern="^pm:"),
            MessageHandler(filters.TEXT & ~filters.COMMAND, peremesh_confirm_text_handler),
        ],
    }
