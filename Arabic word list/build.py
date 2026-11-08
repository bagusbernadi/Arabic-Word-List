"""
Build website kamus Arab Al-Quran dari satu file Excel.

Input : Arabic word list/Kosakata_Arab.xlsx
          - sheet "Kosakata" : satu baris per kata unik (urut kemunculan pertama di mushaf)
          - sheet "Ayat"     : satu baris per ayat unik
          - sheet "Klausa"   : ayat dipecah per klausa (di tanda waqaf), klausa unik
          - sheet "Surah"    : 114 surah (nama, tempat turun, jumlah)
          - sheet "Progress" : rekap hafalan di Excel (tidak dipakai website)
Output: docs/index.html (kamus), docs/hijaiyah.html (Huruf Hijaiyah, data huruf ada di template),
        docs/deck.html (Belajar Surah: kuis per surah/bagian),
        docs/data.json (kata), docs/ayat.json, docs/klausa.json

Kolom dicari berdasarkan nama header, bukan posisi. Kolom Status dan Catatan edit hanya untuk Excel.
Contoh ayat di kamus diambil dari sheet Ayat lewat kolom "Contoh 1 Ref" .. "Contoh 3 Ref" (mis. 2:255).
Website memuat file JSON lewat fetch: buka lewat GitHub Pages atau server lokal
(python -m http.server --directory docs), bukan dengan klik dua kali file index.html.
"""
import json
import re
import sys
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent
XLSX = BASE / "Kosakata_Arab.xlsx"
TEMPLATE = BASE / "template.html"
DECK_TEMPLATE = BASE / "deck_template.html"
HIJ_TEMPLATE = BASE / "hijaiyah_template.html"
OUT_DIR = BASE.parent / "docs"

NOTICE = ("Teks Arab: Tanzil Project (https://tanzil.net), tidak diubah; terjemah ayat: Kementerian Agama RI melalui Tanzil; "
          "arti per kata: SadaqaWorks/Word-By-Word-Quran-Android (MIT); transliterasi dibuat mesin dan belum ditinjau.")
NEEDS = {
    "Kosakata": ["No Surah", "Indonesia", "Transliterasi", "Arab", "Frekuensi", "Ref pertama", "Arti alternatif", "Sumber arti",
                 "Keterangan", "Contoh 1 Ref", "Contoh 2 Ref", "Contoh 3 Ref"],
    "Ayat": ["No Surah", "Ayat", "Indonesia", "Transliterasi", "Arab", "Juga muncul di"],
    "Klausa": ["No Surah", "Ayat", "Klausa ke", "Ref", "Indonesia (gabungan arti per kata)", "Transliterasi", "Arab"],
    "Surah": ["No", "Surah", "Tempat turun", "Arti nama", "Jumlah ayat (mushaf)", "Kata baru", "Ayat unik baru", "Klausa baru"],
}


def clean(v):
    return "" if v is None else str(v).strip()


def table(wb, name):
    if name not in wb.sheetnames:
        sys.exit(f'Sheet "{name}" tidak ditemukan di {XLSX.name}')
    ws = wb[name]
    headers = {clean(c.value): i for i, c in enumerate(ws[1])}
    missing = [h for h in NEEDS[name] if h not in headers]
    if missing:
        sys.exit(f"Header sheet {name} tidak ditemukan:\n  - " + "\n  - ".join(missing))
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(c is not None and str(c).strip() for c in row):
            continue
        rows.append(lambda h, row=row: clean(row[headers[h]]))
    return rows


def dump(path, rows):
    path.write_text(json.dumps({"sumber": NOTICE, "rows": rows}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")


def main():
    if not XLSX.exists():
        sys.exit(f"File tidak ditemukan: {XLSX}")
    wb = openpyxl.load_workbook(XLSX, data_only=True, read_only=False)

    surah = []
    for g in table(wb, "Surah"):
        n = int(g("No"))
        name = re.sub(r"^\d+\.\s*", "", g("Surah"))
        surah.append({"n": n, "name": name, "arti": g("Arti nama"), "type": g("Tempat turun"),
                      "ayat_mushaf": int(g("Jumlah ayat (mushaf)") or 0), "kata": int(g("Kata baru") or 0),
                      "ayat": int(g("Ayat unik baru") or 0), "klausa": int(g("Klausa baru") or 0)})
    if [s["n"] for s in surah] != list(range(1, 115)):
        sys.exit("Sheet Surah harus berisi surah 1 sampai 114 berurutan")

    ayat_refs, ayat = set(), []
    for g in table(wb, "Ayat"):
        s, a = int(g("No Surah")), int(g("Ayat"))
        ayat_refs.add(f"{s}:{a}")
        ayat.append({"s": s, "n": a, "ar": g("Arab"), "tl": g("Transliterasi"), "id": g("Indonesia"), "d": g("Juga muncul di")})

    kata, bad_ref = [], 0
    for g in table(wb, "Kosakata"):
        s = int(g("No Surah"))
        exs = [g(f"Contoh {k} Ref") for k in (1, 2, 3)]
        exs = [e for e in exs if e]
        for e in exs:
            if e not in ayat_refs:
                bad_ref += 1
        o = {"s": s, "i": g("Indonesia"), "t": g("Transliterasi"), "a": g("Arab"), "f": int(g("Frekuensi") or 0),
             "r": g("Ref pertama"), "e": [e for e in exs if e in ayat_refs]}
        if g("Arti alternatif"):
            o["x"] = g("Arti alternatif")
        if "cek" in g("Sumber arti") or not g("Indonesia"):
            o["c"] = 1
        if g("Keterangan"):
            o["n"] = g("Keterangan")
        kata.append(o)
    if bad_ref:
        print(f"PERINGATAN: {bad_ref} referensi contoh tidak ditemukan di sheet Ayat (dilewati)")

    klausa = [{"s": int(g("No Surah")), "r": g("Ref"), "k": int(g("Klausa ke")), "ar": g("Arab"), "tl": g("Transliterasi"),
               "id": g("Indonesia (gabungan arti per kata)")} for g in table(wb, "Klausa")]

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    dump(OUT_DIR / "data.json", kata)
    dump(OUT_DIR / "ayat.json", ayat)
    dump(OUT_DIR / "klausa.json", klausa)
    (OUT_DIR / ".nojekyll").write_text("", encoding="utf-8")
    html = TEMPLATE.read_text(encoding="utf-8")
    marker = "var SURAH=[];"
    if html.count(marker) != 1:
        sys.exit(f"template.html harus memuat tepat satu baris '{marker}'")
    html = html.replace(marker, "var SURAH=" + json.dumps(surah, ensure_ascii=False) + ";")
    (OUT_DIR / "index.html").write_text(html, encoding="utf-8")

    # Belajar Surah (deck.html): memakai CSS yang sama dengan index.html
    dt = DECK_TEMPLATE.read_text(encoding="utf-8")
    style = re.search(r"<style>.*?</style>", html, re.S).group(0)
    if dt.count("<!--STYLE-->") != 1:
        sys.exit("deck_template.html harus memuat tepat satu penanda <!--STYLE-->")
    if dt.count(marker) != 1:
        sys.exit(f"deck_template.html harus memuat tepat satu baris '{marker}'")
    dt = dt.replace("<!--STYLE-->", style).replace(marker, "var SURAH=" + json.dumps(surah, ensure_ascii=False) + ";")
    (OUT_DIR / "deck.html").write_text(dt, encoding="utf-8")
    # Huruf Hijaiyah (hijaiyah.html): halaman statis, CSS sama dengan index.html
    ht = HIJ_TEMPLATE.read_text(encoding="utf-8")
    if ht.count("<!--STYLE-->") != 1:
        sys.exit("hijaiyah_template.html harus memuat tepat satu penanda <!--STYLE-->")
    (OUT_DIR / "hijaiyah.html").write_text(ht.replace("<!--STYLE-->", style), encoding="utf-8")
    print(f"OK: {len(kata)} kata, {len(ayat)} ayat, {len(klausa)} klausa, 114 surah -> {OUT_DIR}")


if __name__ == "__main__":
    main()
