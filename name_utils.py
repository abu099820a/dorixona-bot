"""
name_utils.py — Ism-familyalarni krilcha "Familya Ism" formatiga
avtomatik o'tkazish uchun yordamchi modul.

Lotin yozuvidagi (yoki lotin-kirill aralash) va Ism/Familya tartibi
har xil yozilgan ismlarni:
  1) o'zbekcha krill yozuviga o'tkazadi (transliteratsiya),
  2) Familya-Ism tartibiga solib beradi (odatiy o'zbek familiya
     qo'shimchalari: -ov/-ova/-ev/-eva/-yov/-yova, shuningdek
     -zoda/-zade orqali aniqlanadi).

salary_handlers.py dagi Oylik/Aksiya varaqlariga xodim ismini
yozishdan oldin shu modulning normalize_name() funksiyasi orqali
o'tkaziladi.
"""

import re

APOST_CHARS = ["'", "’", "‘", "ʻ", "ʼ", "`", "´", "ʹ"]
def normalize_apostrophes(s):
    for a in APOST_CHARS:
        s = s.replace(a, "'")
    s = s.replace("ò", "o'").replace("ó", "o'").replace("Ò", "O'").replace("Ó", "O'")
    return s

CYR_RE = re.compile(r'[А-Яа-яЁёҚқҲҳҲЎўҒғ]')
LAT_RE = re.compile(r'[A-Za-z]')

def has_cyrillic(s):
    return bool(CYR_RE.search(s))

PATRONYM_MAP = {
    "qizi": "қизи", "kizi": "қизи", "qiz": "қизи",
    "ogli": "ўғли", "og'li": "ўғли", "ug'li": "ўғли", "o'g'li": "ўғли", "oglu": "ўғли",
    "угли": "ўғли", "кизи": "қизи", "қизи": "қизи", "ўғли": "ўғли",
}

def strip_punct(tok):
    return tok.strip(".,;:!?()[]\"")

VOWELS_LAT = set("aeiou")

def translit_latin_token(tok):
    """Transliterate a single Latin-script Uzbek token to Cyrillic, letter by letter with digraph handling."""
    s = normalize_apostrophes(tok)
    out = []
    i = 0
    n = len(s)
    is_word_start = True
    while i < n:
        ch = s[i]
        low = ch.lower()
        rest = s[i:i+4].lower()

        # o' -> ў ; g' -> ғ
        if low == "o" and i+1 < n and s[i+1] == "'":
            out.append("Ў" if ch.isupper() else "ў")
            i += 2
            is_word_start = False
            continue
        if low == "g" and i+1 < n and s[i+1] == "'":
            out.append("Ғ" if ch.isupper() else "ғ")
            i += 2
            is_word_start = False
            continue
        # bare tutuq belgisi (glottal) not after o/g -> ъ
        if ch == "'":
            out.append("ъ")
            i += 1
            is_word_start = False
            continue
        # digraphs sh, ch
        if rest[:2] == "sh":
            out.append("Ш" if ch.isupper() else "ш")
            i += 2
            is_word_start = False
            continue
        if rest[:2] == "ch":
            out.append("Ч" if ch.isupper() else "ч")
            i += 2
            is_word_start = False
            continue
        # y-combinations
        if low == "y":
            nxt = s[i+1].lower() if i+1 < n else ""
            if nxt == "a":
                out.append("Я" if ch.isupper() else "я"); i += 2; is_word_start = False; continue
            if nxt == "o":
                out.append("Ё" if ch.isupper() else "ё"); i += 2; is_word_start = False; continue
            if nxt == "u":
                out.append("Ю" if ch.isupper() else "ю"); i += 2; is_word_start = False; continue
            if nxt == "e":
                out.append("Е" if ch.isupper() else "е"); i += 2; is_word_start = False; continue
            # bare y -> й
            out.append("Й" if ch.isupper() else "й"); i += 1; is_word_start = False; continue
        if low == "e":
            if is_word_start:
                out.append("Э" if ch.isupper() else "э")
            else:
                out.append("Е" if ch.isupper() else "е")
            i += 1; is_word_start = False; continue

        single_map = {
            "a": "а", "b": "б", "d": "д", "f": "ф", "g": "г", "h": "ҳ",
            "i": "и", "j": "ж", "k": "к", "l": "л", "m": "м", "n": "н",
            "o": "о", "p": "п", "q": "қ", "r": "р", "s": "с", "t": "т",
            "u": "у", "v": "в", "x": "х", "z": "з", "c": "ц", "w": "в",
        }
        if low in single_map:
            mapped = single_map[low]
            out.append(mapped.upper() if ch.isupper() else mapped)
            i += 1
            is_word_start = False
            continue
        # anything else (digits, hyphen, apostrophe leftover) pass through
        out.append(ch)
        if ch == "-" or ch == " ":
            is_word_start = True
        else:
            is_word_start = False
        i += 1
    return "".join(out)

def translit_token_mixed(tok):
    """Apply latin transliteration to latin runs, leave cyrillic runs untouched, for a token possibly hyphenated."""
    # split on hyphen but keep hyphens
    parts = re.split(r'(-)', tok)
    res = []
    for p in parts:
        if p == "-" or p == "":
            res.append(p)
            continue
        if has_cyrillic(p):
            res.append(p)
        elif LAT_RE.search(p):
            res.append(translit_latin_token(p))
        else:
            res.append(p)
    return "".join(res)

def title_case_cyr(tok):
    # capitalize each hyphen-separated part
    parts = tok.split("-")
    out = []
    for p in parts:
        if not p:
            out.append(p)
            continue
        out.append(p[0].upper() + p[1:].lower())
    return "-".join(out)

STRONG_SUFFIXES = ("ова", "ева", "ёва", "ов", "ев", "ёв")
WEAK_SUFFIXES = ("зода", "заде")
SURNAME_SUFFIXES = STRONG_SUFFIXES + WEAK_SUFFIXES

def is_surname_token(cyr_tok):
    low = cyr_tok.lower().strip(".,")
    return low.endswith(SURNAME_SUFFIXES)

def surname_index(cyr_tokens):
    """Return the index of the token most likely to be the surname, or None if ambiguous."""
    lows = [t.lower().strip(".,") for t in cyr_tokens]
    strong = [i for i, t in enumerate(lows) if t.endswith(STRONG_SUFFIXES)]
    if len(strong) == 1:
        return strong[0]
    if len(strong) == 0:
        weak = [i for i, t in enumerate(lows) if t.endswith(WEAK_SUFFIXES)]
        if len(weak) == 1:
            return weak[0]
    return None

def normalize_name(raw):
    if raw is None:
        return ""
    if not isinstance(raw, str):
        return ""  # e.g. phone numbers etc.
    s = raw.strip()
    if not s:
        return ""
    s = normalize_apostrophes(s)
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r'\s+', ' ', s).strip()
    s = s.strip(". ").strip()
    if not s:
        return ""

    raw_tokens = [t for t in s.split(" ") if t]
    cyr_tokens = []
    for t in raw_tokens:
        core = strip_punct(t)
        if not core:
            continue
        norm_check = normalize_apostrophes(core).lower()
        if norm_check in PATRONYM_MAP:
            cyr_tokens.append(PATRONYM_MAP[norm_check])
            continue
        cyr_tokens.append(translit_token_mixed(core))

    if not cyr_tokens:
        return ""

    # find surname token index
    idx = surname_index(cyr_tokens)
    if idx is not None and idx != 0:
        cyr_tokens = [cyr_tokens[idx]] + cyr_tokens[:idx] + cyr_tokens[idx+1:]
    # else: already first, or ambiguous -> leave order as-is

    titled = [title_case_cyr(t) if t not in PATRONYM_MAP.values() else t for t in cyr_tokens]
    # patronym markers should stay lowercase (қизи/ўғли), already lowercase in map
    return " ".join(titled)
