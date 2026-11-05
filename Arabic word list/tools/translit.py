"""
Transliterasi Latin otomatis (gaya Indonesia, mendekati pedoman Kemenag) untuk SATU kata Arab berharakat.

Catatan penting:
- Hasil dibuat mesin dari harakat teks Tanzil, BELUM ditinjau pembaca Al-Quran yang kompeten.
- Kata diproses sendiri-sendiri (bentuk tersambung di dalam kata, bukan antar kata).
- Panjang bunyi: ā ī ū. Tanda ' = hamzah, ʿ = 'ain.
"""
import re
import unicodedata

CONS = {
    "ب": "b", "ت": "t", "ث": "ts", "ج": "j", "ح": "h", "خ": "kh", "د": "d", "ذ": "dz", "ر": "r", "ز": "z",
    "س": "s", "ش": "sy", "ص": "sh", "ض": "dh", "ط": "th", "ظ": "zh", "ع": "ʿ", "غ": "gh", "ف": "f", "ق": "q",
    "ك": "k", "ل": "l", "م": "m", "ن": "n", "ه": "h", "و": "w", "ي": "y",
}
SUN = set("تثدذرزسشصضطظلن")
HAMZA_CARRIERS = set("ءأإؤئ")
ALIF = ("ا", "ٱ")
FATHA, DAMMA, KASRA = "َ", "ُ", "ِ"
TANF, TAND, TANK = "ً", "ٌ", "ٍ"
SHADDA, SUKUN, MADDA, DAGGER = "ّ", "ْ", "ٓ", "ٰ"
VOWEL_ROMAN = {FATHA: "a", DAMMA: "u", KASRA: "i", TANF: "an", TAND: "un", TANK: "in"}
# tanda baca kecil Al-Quran (waqaf, dsb.) dan tatwil: diabaikan
SKIP = {chr(c) for c in list(range(0x06D6, 0x06DD)) + list(range(0x06DF, 0x06EE))} | {"ـ", "ۥ", "ۦ"}
SKIP -= {"ۡ"}


def _letters(word):
    """Pecah kata menjadi [[huruf, set_harakat]]."""
    out = []
    for ch in unicodedata.normalize("NFC", word):
        if ch in SKIP:
            continue
        if unicodedata.category(ch).startswith("M") or ch == "ۡ":
            if ch == "ۡ":
                ch = SUKUN
            if out:
                out[-1][1].add(ch)
            continue
        out.append([ch, set()])
    return out


def _vowel(marks):
    for m in (FATHA, DAMMA, KASRA, TANF, TAND, TANK):
        if m in marks:
            return m
    return None


def _lengthen(prev, short, long_):
    if prev.endswith(short):
        return prev[: -len(short)] + long_
    return prev


def translit_word(word):
    L = _letters(word)
    n = len(L)
    out = []
    i = 0
    prev_v = None          # harakat huruf sebelumnya
    al_pending = False     # baru melewati "al" (alif wasal + lam)
    while i < n:
        ch, mk = L[i]
        nxt = L[i + 1] if i + 1 < n else None
        v = _vowel(mk)
        sh = SHADDA in mk
        # --- lam "li" + "al" melebur: لِّلْ -> lil- ---
        if i == 0 and ch == "ل" and sh and nxt and nxt[0] == "ل" and SUKUN in nxt[1]:
            out.append("lil-")
            i += 2
            prev_v = None
            continue
        # --- alif madd (آ) ---
        if ch == "آ":
            out.append(("'" if i else "") + "ā")
            prev_v = None
            i += 1
            continue
        # --- alif ---
        if ch in ALIF and v is None:
            if i == 0 or (i == 1 and L[0][0] in "وفبكل" and _vowel(L[0][1])):
                nl = nxt is not None and nxt[0] == "ل" and (_vowel(nxt[1]) is None or SHADDA in nxt[1])
                if nl and ch in ALIF:
                    if i == 0:
                        out.append("a")
                    al_pending = True
                    prev_v = None
                    i += 1
                    continue
                third = L[i + 2] if i + 2 < n else None
                if i == 0:
                    out.append("u" if (third and DAMMA in third[1]) else "i")
                prev_v = None
                i += 1
                continue
            if prev_v == FATHA and out:       # alif panjang: ā
                out[-1] = _lengthen(out[-1], "a", "ā")
            prev_v = None                      # selain itu alif tidak dibaca
            i += 1
            continue
        # --- hamzah ---
        if ch in HAMZA_CARRIERS or (ch in ALIF and v):
            out.append(("" if i == 0 else "'") + (VOWEL_ROMAN[v] if v else ""))
            prev_v = v
            if v == TANF and nxt and nxt[0] in ALIF and _vowel(nxt[1]) is None:
                i += 1
            i += 1
            continue
        # --- alif maqsurah ---
        if ch == "ى" and v is None:
            if prev_v == FATHA and out:
                out[-1] = _lengthen(out[-1], "a", "ā")
            prev_v = None
            i += 1
            continue
        # --- ta marbutah ---
        if ch == "ة":
            out.append(("t" + VOWEL_ROMAN[v]) if v else "h")
            prev_v = v
            if v == TANF and nxt and nxt[0] in ALIF and _vowel(nxt[1]) is None:
                i += 1
            i += 1
            continue
        # --- waw / ya sebagai huruf madd atau diftong ---
        if ch in ("و", "ي") and v is None and not sh and i > 0 and out:
            pv = L[i - 1][1]
            if ch == "و" and DAMMA in pv:
                out[-1] = _lengthen(out[-1], "u", "ū")
                prev_v = None
                if nxt and nxt[0] in ALIF and _vowel(nxt[1]) is None and i + 2 >= n:
                    i += 1                     # alif waw jamak: tidak dibaca
                i += 1
                continue
            if ch == "ي" and KASRA in pv:
                out[-1] = _lengthen(out[-1], "i", "ī")
                prev_v = None
                i += 1
                continue
            if FATHA in pv:                    # aw / ay
                out.append("w" if ch == "و" else "y")
                prev_v = None
                i += 1
                continue
        # --- konsonan ---
        r = CONS.get(ch)
        if r is None:
            i += 1
            continue
        if al_pending:
            al_pending = False
            if ch == "ل" and nxt is not None and SHADDA in nxt[1] and nxt[0] in SUN and v is None:
                i += 1                         # lam melebur ke huruf syamsiyah berikutnya
                al_pending = "sun"
                continue
            if ch == "ل" and v is None and not sh:
                out.append("l-")               # lam qamariyah: al-
                prev_v = None
                i += 1
                continue
            if sh and ch in SUN:
                out.append(r + "-" + r)        # huruf syamsiyah: asy-syams
            else:
                out.append(r if not sh else (r + r if len(r) == 1 else r[0] + r))
        elif sh:
            out.append(r + r if len(r) == 1 else r[0] + r)
        else:
            out.append(r)
        if al_pending == "sun":
            al_pending = False
        if v:
            rv = VOWEL_ROMAN[v]
            if DAGGER in mk and v == FATHA:
                rv = "ā"
            out.append(rv)
        elif DAGGER in mk:
            out.append("ā")
        prev_v = v
        if v == TANF and nxt and nxt[0] in ALIF and _vowel(nxt[1]) is None:
            i += 1
        i += 1
    s = "".join(out)
    s = re.sub(r"al-lah", "allāh", s)
    s = re.sub(r"illah", "illāh", s)
    return s


if __name__ == "__main__":
    import sys
    for w in sys.argv[1:]:
        print(w, "->", translit_word(w))
