"""
Menyusun Kosakata_Arab.xlsx dari sumber data (sekali jalan; hasilnya yang diedit dan di-build).

Sumber:
  - Teks Arab        : Tanzil Project (quran-simple.txt)        https://tanzil.net/download/
  - Terjemah ayat    : Kemenag RI via Tanzil (id.indonesian.txt) https://tanzil.net/trans/
  - Arti per kata    : SadaqaWorks/Word-By-Word-Quran-Android (MIT), tabel `bywords`, dipetakan ke kata Tanzil
  - Nama surah       : tabel `surah_name` pada database yang sama
  - Transliterasi    : dibuat mesin oleh tools/translit.py (belum ditinjau pembaca Al-Quran)

Pakai:
  python tools/prep_data.py quran-simple.txt id.indonesian.txt wordbyword.db Kosakata_Arab.xlsx

Aturan dedup (urutan mushaf, surah 1 -> 114):
  - Kata  : bentuk persis (dengan harakat) yang sudah muncul lebih dulu tidak ditulis lagi.
  - Ayat  : teks Arab persis yang sudah muncul lebih dulu (surah sama atau sebelumnya) tidak ditulis lagi.
  - Klausa: ayat dipecah di tanda waqaf kecil; klausa dengan teks Arab persis sama tidak ditulis lagi.
  - Basmalah di awal ayat 1 (kecuali surah 1 dan 9) dibuang; basmalah hanya dihitung pada 1:1.
"""
import collections
import re
import sqlite3
import sys
import unicodedata
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
from translit import translit_word  # noqa: E402

WAQF = {chr(c) for c in range(0x06D6, 0x06DD)}
MAX_EX_WORDS = 25
N_EX = 3
# arti untuk kata yang tidak ada padanannya di sumber arti per kata (ditulis Claude, PERLU DICEK)
MANUAL = {
    "يَا": "wahai", "اللَّيْلِ": "malam", "حَيَاةٌ": "kehidupan", "وَيَبْسُطُ": "dan Dia melapangkan",
    "بِاللَّيْلِ": "pada malam hari", "تَسْأَمُوا": "kamu bosan", "اللَّيْلَ": "malam", "اصْطَفَاكِ": "Dia telah memilihmu",
    "وَاصْطَفَاكِ": "dan telah memilihmu", "هَا": "inilah", "وَاللَّاتِي": "dan (perempuan-perempuan) yang",
    "وَاللَّذَانِ": "dan dua orang yang", "اللَّاتِي": "(perempuan-perempuan) yang", "تَلْوُوا": "kamu memutarbalikkan",
    "أَلْقَاهَا": "dia melemparkannya", "وَيْلَتَىٰ": "celaka aku", "لَيْتَنَا": "sekiranya kami", "حَسْرَتَنَا": "penyesalan kami",
    "اللَّيْلُ": "malam", "مَعْشَرَ": "wahai golongan", "مَثْوَاكُمْ": "tempat tinggalmu", "وَيَا": "dan wahai",
    "يَرَاكُمْ": "dia melihat kamu", "أَنجَاهُمْ": "Dia menyelamatkan mereka", "وَآتَانِي": "dan Dia memberiku",
    "تَسْأَلْنِ": "kamu bertanya kepadaku", "أَتَنْهَانَا": "apakah kamu melarang kami", "أَصَلَاتُكَ": "apakah shalatmu",
    "أَنْهَاكُمْ": "aku melarang kamu", "أَبَتِ": "wahai ayahku", "رُؤْيَاكَ": "mimpimu", "صَاحِبَيِ": "wahai kedua temanku",
    "رُؤْيَايَ": "mimpiku", "لِلرُّؤْيَا": "tentang mimpi", "أَسَفَىٰ": "duhai sedihnya", "تَيْأَسُوا": "kamu berputus asa",
    "يَيْأَسُ": "berputus asa", "مُّزْجَاةٍ": "yang sedikit nilainya", "وَيْلَتَنَا": "celaka kami", "أَحْصَاهَا": "dia menghitungnya",
    "تُعَذِّبَ": "kamu menyiksa", "سَامِرِيُّ": "Samiri", "وَيْلَنَا": "celaka kami", "لِيَسْتَأْذِنكُمُ": "hendaklah meminta izin kepadamu",
    "صَلَاةِ": "shalat", "أَقْصَى": "paling jauh", "اسْتَأْجِرْهُ": "upahlah dia", "اسْتَأْجَرْتَ": "kamu mengupah",
    "هَامَانُ": "Haman", "لَيْتَ": "sekiranya", "اللَّائِي": "(perempuan-perempuan) yang", "وَيَسْتَأْذِنُ": "dan meminta izin",
    "إِنَاهُ": "waktu masaknya", "مُسْتَأْنِسِينَ": "asyik bercakap-cakap", "فَيَسْتَحْيِي": "maka dia malu",
    "فَاسْأَلُوهُنَّ": "maka mintalah kepada mereka", "وَبِاللَّيْلِ": "dan pada malam hari", "حَسْرَتَىٰ": "duhai penyesalan",
    "النَّجَاةِ": "keselamatan", "اللَّذَيْنِ": "dua orang yang", "أَتْقَاكُمْ": "yang paling bertakwa di antara kamu",
    "وَلْيَسْأَلُوا": "dan hendaklah mereka meminta", "لِلْحَوَارِيِّينَ": "kepada para pengikut setia", "لِلصَّلَاةِ": "untuk shalat",
    "وَاللَّائِي": "dan (perempuan-perempuan) yang", "لَيْتَهَا": "sekiranya ia", "وَاللَّيْلِ": "demi malam",
}


def norm(s):
    s = unicodedata.normalize("NFD", s).replace("ٰ", "ا")
    s = "".join(c for c in s if not unicodedata.category(c).startswith("M") and unicodedata.category(c) != "Lm")
    for a, b in (("ٱ", "ا"), ("آ", "ا"), ("أ", "ا"), ("إ", "ا"), ("ى", "ي"),
                 ("ة", "ه"), ("ؤ", "و"), ("ئ", "ي"), ("ء", "")):
        s = s.replace(a, b)
    return re.sub("ا+", "ا", s)


def skel(s):
    return re.sub("[اوي]", "", norm(s))


def is_word(tok):
    return any(unicodedata.category(c)[0] == "L" and unicodedata.category(c) != "Lm" for c in tok)


def read_pipe(path):
    rows = []
    for ln in open(path, encoding="utf-8"):
        if ln.startswith("#") or not ln.strip():
            continue
        s, a, t = ln.rstrip("\n").split("|", 2)
        rows.append((int(s), int(a), unicodedata.normalize("NFC", t.strip())))
    return rows


def main(ar_path, id_path, db_path, out_path):
    ar_rows, id_rows = read_pipe(ar_path), read_pipe(id_path)
    assert [(s, a) for s, a, _ in ar_rows] == [(s, a) for s, a, _ in id_rows], "kunci ayat Arab/Indonesia tidak sama"
    indo = {(s, a): t for s, a, t in id_rows}
    bism = ar_rows[0][2]

    # --- ayat bersih (tanpa basmalah) -> token & segmen waqaf
    verses = []
    for s, a, t in ar_rows:
        w = t.split()
        if a == 1 and s not in (1, 9):
            w = w[4:]
        text = " ".join(w)
        # segmen (klausa): pecah di tanda waqaf
        segs, cur = [], []
        for x in w:
            if all(c in WAQF for c in x):
                if cur:
                    segs.append(cur)
                    cur = []
            elif is_word(x):
                cur.append(x)
        if cur:
            segs.append(cur)
        verses.append({"s": s, "a": a, "text": text, "toks": [x for x in w if is_word(x)], "segs": segs})

    # --- arti per kata (SadaqaWorks)
    db = sqlite3.connect("file:%s?mode=ro" % db_path, uri=True)
    bw = collections.defaultdict(list)
    for s, v, i, ar, g in db.execute("select surah_id,verse_id,words_id,words_ar,translate_indo from bywords order by surah_id,verse_id,words_id"):
        g = (g or "").strip().replace("sholat", "shalat").replace("Sholat", "Shalat")   # samakan dengan ejaan Kemenag
        bw[(s, v)].append((ar, g))
    byform = collections.defaultdict(collections.Counter)
    for lst in bw.values():
        for ar, g in lst:
            if g:
                byform[norm(ar)][g] += 1
    snames = {r[0]: r[1:] for r in db.execute("select surah_no,name_arabic,name_english,type,arti_nama,ayah_number from surah_name")}

    aligned, aligned2 = {}, {}
    ctx_gloss = {}          # (s,a,idx) -> arti pada konteks (bila selaras)
    first, freq = collections.OrderedDict(), collections.Counter()
    for v in verses:
        b = bw[(v["s"], v["a"])]
        same_len = len(b) == len(v["toks"])
        for i, x in enumerate(v["toks"]):
            freq[x] += 1
            if x not in first:
                first[x] = (v["s"], v["a"], i + 1)
            if same_len and b[i][1]:
                if norm(b[i][0]) == norm(x):
                    ctx_gloss[(v["s"], v["a"], i)] = b[i][1]
                    aligned.setdefault(x, b[i][1])
                elif skel(b[i][0]) == skel(x):
                    ctx_gloss[(v["s"], v["a"], i)] = b[i][1]
                    aligned2.setdefault(x, b[i][1])

    words = []             # urut kemunculan pertama
    for x, (s, a, pos) in first.items():
        nf = norm(x)
        alts = []
        if x in aligned:
            g, src = aligned[x], "selaras"
        elif nf in byform:
            g, src = byform[nf].most_common(1)[0][0], "bentuk-sama"
        elif x in aligned2:
            g, src = aligned2[x], "ejaan-beda (cek)"
        elif x in MANUAL:
            g, src = MANUAL[x], "manual (cek)"
        else:
            g, src = "", "belum-ada"
        if nf in byform:
            tot = sum(byform[nf].values())
            # buang arti yang jarang (kemungkinan salah segmentasi di sumber): minimal 2 kali dan 3% dari kemunculan
            alts = [k for k, c in byform[nf].most_common(6) if k != g and c >= 2 and c >= 0.03 * tot][:3]
        words.append({"ar": x, "s": s, "a": a, "pos": pos, "id": g, "src": src, "freq": freq[x], "alt": " | ".join(alts),
                      "tl": translit_word(x)})
    wmap = {w["ar"]: w for w in words}
    print("kata unik:", len(words), dict(collections.Counter(w["src"] for w in words)))
    missing = [w["ar"] for w in words if w["src"] == "belum-ada"]
    if missing:
        print("KATA BELUM ADA ARTI (%d):" % len(missing), " ".join(missing))

    def gloss_of(v, i):
        g = ctx_gloss.get((v["s"], v["a"], i))
        return g if g else wmap[v["toks"][i]]["id"]

    # --- ayat unik
    seen_a, ayat = {}, []
    for v in verses:
        if v["text"] in seen_a:
            seen_a[v["text"]]["dups"].append("%d:%d" % (v["s"], v["a"]))
            continue
        row = {"s": v["s"], "a": v["a"], "ar": v["text"], "n": len(v["toks"]), "dups": [],
               "tl": " ".join(wmap[x]["tl"] for x in v["toks"]), "id": indo[(v["s"], v["a"])], "toks": v["toks"]}
        seen_a[v["text"]] = row
        ayat.append(row)
    print("ayat unik:", len(ayat), "| duplikat:", len(verses) - len(ayat))

    # --- contoh kalimat per kata (ayat unik terpendek yang memuat kata)
    occ = collections.defaultdict(list)
    for r in ayat:
        for x in set(r["toks"]):
            occ[x].append(r)
    for w in words:
        cands = sorted([r for r in occ[w["ar"]] if r["n"] <= MAX_EX_WORDS] or occ[w["ar"]],
                       key=lambda r: (r["n"], r["s"], r["a"]))
        w["ex"] = ["%d:%d" % (r["s"], r["a"]) for r in cands[:N_EX]]

    # --- klausa unik
    seen_k, klausa, kidx = set(), [], 0
    for v in verses:
        pos = 0
        for k, seg in enumerate(v["segs"], 1):
            ar = " ".join(seg)
            idxs = list(range(pos, pos + len(seg)))
            pos += len(seg)
            if ar in seen_k:
                continue
            seen_k.add(ar)
            kidx += 1
            klausa.append({"s": v["s"], "a": v["a"], "k": k, "kn": len(v["segs"]), "ar": ar, "n": len(seg),
                           "tl": " ".join(wmap[v["toks"][i]]["tl"] for i in idxs),
                           "id": " ".join(g for g in (gloss_of(v, i) for i in idxs) if g)})
    print("klausa unik:", len(klausa))

    # --- tulis Excel
    wb = openpyxl.Workbook()
    head_fill = PatternFill("solid", fgColor="4285F4")
    head_font = Font(bold=True, color="FFFFFF", name="Calibri")
    ar_font = Font(name="Arial", size=14)

    def sheet(title, headers, rows, widths, ar_cols=(), first=False):
        ws = wb.active if first else wb.create_sheet()
        ws.title = title
        ws.append(headers)
        for r in rows:
            ws.append(r)
        for c in range(1, len(headers) + 1):
            h = ws.cell(row=1, column=c)
            h.fill, h.font = head_fill, head_font
            h.alignment = Alignment(vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(c)].width = widths[c - 1]
        for c in ar_cols:
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=c)
                cell.font = ar_font
                cell.alignment = Alignment(horizontal="right", readingOrder=2, wrap_text=True, vertical="top")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        return ws

    def sur(s):
        n = snames[s]
        return "%d. %s" % (s, n[1].replace("Surah ", ""))

    kos_rows = []
    for i, w in enumerate(words, 1):
        ex = w["ex"] + [""] * (N_EX - len(w["ex"]))
        kos_rows.append([i, w["s"], sur(w["s"]), w["id"], w["tl"], w["ar"], w["freq"], "%d:%d" % (w["s"], w["a"]), w["alt"],
                         w["src"], "", "Belum", ""] + ex)
    ws = sheet("Kosakata",
               ["No", "No Surah", "Surah", "Indonesia", "Transliterasi", "Arab", "Frekuensi", "Ref pertama", "Arti alternatif",
                "Sumber arti", "Keterangan", "Status", "Catatan edit", "Contoh 1 Ref", "Contoh 2 Ref", "Contoh 3 Ref"],
               kos_rows, [7, 9, 22, 30, 24, 22, 10, 11, 34, 16, 28, 10, 26, 12, 12, 12], ar_cols=(6,), first=True)
    dv = DataValidation(type="list", formula1='"Belum,Sedang,Hafal"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("L2:L%d" % ws.max_row)

    ay_rows = [[i, r["s"], sur(r["s"]), r["a"], "%d:%d" % (r["s"], r["a"]), r["id"], r["tl"], r["ar"], r["n"],
                ", ".join(r["dups"]), "Belum", ""] for i, r in enumerate(ayat, 1)]
    ws2 = sheet("Ayat", ["No", "No Surah", "Surah", "Ayat", "Ref", "Indonesia", "Transliterasi", "Arab", "Jumlah kata",
                         "Juga muncul di", "Status", "Catatan edit"],
                ay_rows, [7, 9, 22, 7, 9, 48, 44, 48, 11, 26, 10, 26], ar_cols=(8,))
    dv2 = DataValidation(type="list", formula1='"Belum,Sedang,Hafal"', allow_blank=True)
    ws2.add_data_validation(dv2)
    dv2.add("K2:K%d" % ws2.max_row)

    kl_rows = [[i, r["s"], sur(r["s"]), r["a"], r["k"], "%d:%d" % (r["s"], r["a"]), r["id"], r["tl"], r["ar"], r["n"],
                "Belum", ""] for i, r in enumerate(klausa, 1)]
    ws3 = sheet("Klausa", ["No", "No Surah", "Surah", "Ayat", "Klausa ke", "Ref", "Indonesia (gabungan arti per kata)",
                           "Transliterasi", "Arab", "Jumlah kata", "Status", "Catatan edit"],
                kl_rows, [7, 9, 22, 7, 9, 9, 48, 40, 44, 11, 10, 26], ar_cols=(9,))
    dv3 = DataValidation(type="list", formula1='"Belum,Sedang,Hafal"', allow_blank=True)
    ws3.add_data_validation(dv3)
    dv3.add("K2:K%d" % ws3.max_row)

    cnt_w = collections.Counter(w["s"] for w in words)
    cnt_a = collections.Counter(r["s"] for r in ayat)
    cnt_k = collections.Counter(r["s"] for r in klausa)
    # snames[s] = (name_arabic, name_english, type, arti_nama, ayah_number)
    sr_rows = []
    for s in range(1, 115):
        n = snames[s]
        sr_rows.append([s, sur(s), n[0], n[2], n[3], n[4], cnt_w[s], cnt_a[s], cnt_k[s]])
    sheet("Surah", ["No", "Surah", "Nama Arab", "Tempat turun", "Arti nama", "Jumlah ayat (mushaf)", "Kata baru", "Ayat unik baru", "Klausa baru"],
          sr_rows, [6, 24, 18, 13, 20, 12, 10, 12, 12], ar_cols=(3,))

    wp = wb.create_sheet("Progress")
    wp.append(["No", "Surah", "Kata baru", "Hafal", "Sedang", "Belum", "% hafal"])
    for c in range(1, 8):
        h = wp.cell(row=1, column=c)
        h.fill, h.font = head_fill, head_font
    for s in range(1, 115):
        r = s + 1
        wp.append([s, sur(s), "=COUNTIFS(Kosakata!$B:$B,A%d)" % r, '=COUNTIFS(Kosakata!$B:$B,A%d,Kosakata!$L:$L,"Hafal")' % r,
                   '=COUNTIFS(Kosakata!$B:$B,A%d,Kosakata!$L:$L,"Sedang")' % r, "=C%d-D%d-E%d" % (r, r, r),
                   '=IF(C%d=0,0,D%d/C%d)' % (r, r, r)])
        wp.cell(row=r, column=7).number_format = "0%"
    wp.append(["", "Total", "=SUM(C2:C115)", "=SUM(D2:D115)", "=SUM(E2:E115)", "=SUM(F2:F115)", "=IF(C116=0,0,D116/C116)"])
    wp.cell(row=116, column=7).number_format = "0%"
    for c, wd in enumerate([6, 24, 11, 9, 9, 9, 9], 1):
        wp.column_dimensions[get_column_letter(c)].width = wd
    wp.freeze_panes = "A2"

    wb.save(out_path)
    print("OK ->", out_path)


if __name__ == "__main__":
    if len(sys.argv) != 5:
        sys.exit(__doc__)
    main(*sys.argv[1:])
