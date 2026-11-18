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


def dorixona_type_keyboard(language: str = "uz"):
    back_txt = "⬅️ Назад" if language == "ru" else "⬅️ Orqaga"
    return ReplyKeyboardMarkup(
        [[BTN_ODDIY], [BTN_VOZVRAT], [BTN_PEREMESH], [back_txt]],
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
        s = max(
            difflib.SequenceMatcher(None, n, cn).ratio(),
            difflib.SequenceMatcher(None, t, _tok(d["nomi"])).ratio(),
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

def _find_header(rows: list, groups: dict, scan: int = 15):
    """
    Sarlavha qatorini topadi. groups = {"kalit": ["so'z1", "so'z2"], ...}
    Qaytaradi: (sarlavha_qator_indeksi, {kalit: ustun_indeksi}) yoki (None, None)
    """
    for ri, row in enumerate(rows[:scan]):
        cells = [str(c or "").strip().lower() for c in row]
        found = {}
        for key, words in groups.items():
            for ci, c in enumerate(cells):
                if ci in found.values():
                    continue
                if any(w in c for w in words):
                    found[key] = ci
                    break
        if len(found) == len(groups):
            return ri, found
    return None, None


def _read_xlsx_rows(data: bytes) -> list:
    """Excel (xlsx) ning birinchi varag'ini qiymatlar ro'yxati sifatida qaytaradi."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    ws = wb.worksheets[0]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def parse_vozvrat_excel(data: bytes) -> list:
    """Vozvrat Excel: Филиал | Наименование | Кол-во → [{filial, dori, soni}]"""
    rows = _read_xlsx_rows(data)
    hi, cols = _find_header(rows, {
        "filial": ["филиал", "аптека", "filial"],
        "dori":   ["наименование", "препарат", "товар", "dori"],
        "soni":   ["кол", "soni", "кол-во"],
    })
    if hi is None:
        raise ValueError("Sarlavha topilmadi (kerakli ustunlar: Филиал, Наименование, Кол-во)")
    out = []
    for r in rows[hi + 1:]:
        f = str(r[cols["filial"]] or "").strip() if cols["filial"] < len(r) else ""
        d = str(r[cols["dori"]] or "").strip() if cols["dori"] < len(r) else ""
        q = _to_num(r[cols["soni"]]) if cols["soni"] < len(r) else None
        if f and d and q:
            out.append({"filial": f, "dori": d, "soni": q})
    return out


def parse_peremesh_excel(data: bytes) -> list:
    """Peremeshenie Excel: Откуда | Кол-во | Препарат | Куда → [{dan, ga, dori, soni}]"""
    rows = _read_xlsx_rows(data)
    hi, cols = _find_header(rows, {
        "dan":  ["откуда", "отправител", "qayerdan", "beruvchi"],
        "ga":   ["куда", "получател", "qayerga", "oluvchi"],
        "dori": ["препарат", "наименование", "товар", "dori"],
        "soni": ["кол", "soni"],
    })
    if hi is None:
        raise ValueError("Sarlavha topilmadi (kerakli ustunlar: Откуда, Куда, Препарат, Кол-во)")
    out = []
    for r in rows[hi + 1:]:
        g = lambda k: (r[cols[k]] if cols[k] < len(r) else None)
        dan, ga, dori = str(g("dan") or "").strip(), str(g("ga") or "").strip(), str(g("dori") or "").strip()
        q = _to_num(g("soni"))
        if dan and ga and dori and q:
            out.append({"dan": dan, "ga": ga, "dori": dori, "soni": q})
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


async def _send_messages(ctx, admin_chat_id: int, messages: list, progress_msg=None) -> list:
    """
    messages: [{filial_no, label, rol, tid, header, lines, footer, log_text}]
    Har biriga "✅ Bajarildi / ❌ Bekor qilish" tugmasi bilan yuboradi.
    Qaytaradi: [(filial_no, label, rol, emoji, sabab, n_items)]
    """
    from attendance import run_write
    from salary_handlers import log_murojaat, _tg_error_reason

    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Bajarildi", callback_data=f"appealresp:done:{admin_chat_id}"),
        InlineKeyboardButton("❌ Bekor qilish", callback_data=f"appealresp:cancel:{admin_chat_id}"),
    ]])
    log_rows = ctx.bot_data.setdefault("appeal_log_rows", {})
    tid_names = ctx.bot_data.setdefault("appeal_tid_names", {})

    results = []
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
            try:
                row_idx = await run_write(log_murojaat, m["tid"], m["log_text"])
                if row_idx and row_idx > 0:
                    log_rows[int(m["tid"])] = row_idx
            except Exception:
                pass
        except Exception as exc:
            results.append((m["filial_no"], m["label"], m["rol"], "❌", _tg_error_reason(exc), n))
            logger.warning(f"[VOZVRAT/PEREMESH] {m['label']} ga yuborilmadi: {exc}")
        if progress_msg and (i % 10 == 0 or i == total):
            try:
                await progress_msg.edit_text(f"⏳ Yuborilmoqda... {i} / {total}")
            except Exception:
                pass
        await asyncio.sleep(0.05)
    return results


def _report_xlsx(results: list, detail_header: list, detail_rows: list, title: str) -> io.BytesIO:
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

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


async def _back_to_main(ctx, chat_id: int, text: str = "📋 Asosiy menyu"):
    from bot import main_keyboard, get_lang
    for k in ("vz_rows", "vz_plan", "vz_base", "pm_rows", "pm_plan", "pm_base", "vz_files"):
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
    await update.message.reply_text(
        "💊 *Dorixonalarga murojaat* — turini tanlang:\n\n"
        f"{BTN_ODDIY} — matnli xabar\n"
        f"{BTN_VOZVRAT} — skrinshot/Excel → har filialga o'z ro'yxati\n"
        f"{BTN_PEREMESH} — tayyor Excel → beruvchi va qabul qiluvchiga",
        parse_mode="Markdown",
        reply_markup=dorixona_type_keyboard(language),
    )
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
        ctx.user_data["vz_rows"] = []
        ctx.user_data["vz_files"] = 0
        await update.message.reply_text(
            "📦 *Skladga vozvrat*\n\n"
            "Dasturdagi jadval *skrinshot(lar)ini* yoki *Excel* faylini yuboring "
            "(ustunlar: Филиал, Наименование, Кол-во).\n\n"
            "Bir nechta skrinshot bo'lsa — hammasini ketma-ket yuboring, "
            f"so'ng *{BTN_TAYYOR}* tugmasini bosing.",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup([[BTN_TAYYOR], [BTN_BEKOR]], resize_keyboard=True),
        )
        return VOZVRAT_UPLOAD

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
        if txt == BTN_TAYYOR:
            if not rows:
                await msg.reply_text("⚠️ Hali hech qanday ma'lumot yuklanmadi. Skrinshot yoki Excel yuboring.")
                return VOZVRAT_UPLOAD
            return await _vozvrat_analyze_and_preview(update, ctx)
        await msg.reply_text(f"📎 Skrinshot yoki Excel yuboring, tugatgach {BTN_TAYYOR} ni bosing.")
        return VOZVRAT_UPLOAD

    # ── Excel ──
    doc = msg.document
    if doc and (doc.file_name or "").lower().endswith((".xlsx", ".xlsm")):
        try:
            f = await ctx.bot.get_file(doc.file_id)
            data = bytes(await f.download_as_bytearray())
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
    lines = [
        "📦 *Vozvrat — tahlil natijasi*\n",
        f"🏥 Filiallar: *{len(plan['items'])}* ta",
        f"💊 Pozitsiyalar: *{n_pos}* ta",
    ]
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
    buf = _report_xlsx([], ["Filial №", "Dorixona", "Dori", "Soni", "1C dagi nomi", "Holat"],
                       detail, "Vozvrat")

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
    sender = _sender_name(update.effective_user)

    messages, detail = [], []
    for fno, drugs in plan["items"].items():
        d = ph[fno]
        label = _label(d)
        dlines = [f"{i}. {dori} — {_fmt_qty(q)} ta" for i, (dori, q) in enumerate(drugs.items(), 1)]
        messages.append({
            "filial_no": fno, "label": label, "rol": "Vozvrat", "tid": d.get("tid", ""),
            "header": (f"📦 Skladga vozvrat\n👤 Kimdan: {sender}\n\n🏥 {label}\n\n"
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
    await _back_to_main(ctx, chat_id)
    return MENU


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
        "🔄 *Peremeshenie — tahlil natijasi*\n",
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
            "header": (f"🔄 Peremeshenie\n👤 Kimdan: {sender}\n\n🏥 {label}\n\n"
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
            "header": (f"🔄 Peremeshenie\n👤 Kimdan: {sender}\n\n🏥 {label}\n\n"
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
    await _back_to_main(ctx, chat_id)
    return MENU


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
