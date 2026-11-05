# Arabic Word List

🔗 **Website kamus:** [https://bagusbernadi.github.io/Arabic-Word-List/](https://bagusbernadi.github.io/Arabic-Word-List/)

Kamus kosakata Al-Quran (Arab – transliterasi – Indonesia), dibagi per 114 surah. Sistemnya sama dengan [Mandarin-Word-List](https://github.com/bagusbernadi/Mandarin-Word-List): website dibuat otomatis dari satu file Excel dan ditayangkan lewat GitHub Pages.

## Isi website

- **Kamus** (`index.html`): pencarian Indonesia, transliterasi, atau Arab (tanpa harakat juga bisa). Tiga jenis entri, dipilih lewat tombol di bawah kolom pencarian:
  - **Kata** (18.242): setiap bentuk kata (dengan harakat) hanya ditulis sekali, di surah tempat ia **pertama kali muncul**. Klik arti Indonesia untuk melihat 3 contoh ayat.
  - **Ayat** (6.059): ayat dengan teks Arab persis sama yang sudah muncul sebelumnya tidak ditulis lagi (ayat yang sama ditulis di kolom "Juga muncul di").
  - **Klausa** (9.655): ayat dipecah di tanda waqaf kecil (ۖ ۗ ۚ ۛ dst.); klausa yang sama tidak ditulis lagi. Artinya adalah gabungan arti per kata (terjemah resmi hanya tersedia per ayat).
- **Surah** (tombol "Semua surah ▾"): memilih satu dari 114 surah. Kamus dan Latihan mengikuti surah yang dipilih (kata baru di surah itu).
- **Latihan** (tab): kuis 10 soal, 3 jenis. **A:** Indonesia → Arab + transliterasi. **B:** Arab → Indonesia. **C:** Transliterasi → Arab. Pilihan jawaban dipilih yang mirip huruf dan artinya.
- **Belajar Surah** (`deck.html`, tab di sebelah "Latihan"; padanan "30H Belajar" di repo Mandarin): pilih salah satu dari 114 surah (biru = Makkiyah, hijau = Madaniyah), lalu pilih **bagian**. Satu bagian berisi maksimal 20 kata baru surah itu, urut kemunculan di mushaf (Al-Baqarah punya 2.627 kata baru, jadi 132 bagian). Setiap kata diuji satu kali, Indonesia → Arab + transliterasi, dengan pilihan jawaban hanya dari bagian yang sama. Kotak pilihan surah dan bagian disembunyikan secara bawaan; buka/tutup dengan tombol **Pilih surah / bagian**. Setelah menjawab ada tombol **Contoh ayat**, soal bisa digeser kiri/kanan, dan di akhir ada "Ulangi bagian ini" serta "Bagian/Surah berikutnya". Posisi terakhir diingat di browser.
- **Suara**: tombol speaker memakai suara bawaan browser (Web Speech API, `ar-SA`). Belum ada file audio di versi 1.

## Aturan dedup (urutan mushaf, surah 1 → 114)

- **Kata:** bentuk persis (dengan harakat) yang sudah muncul lebih dulu tidak ditulis lagi.
- **Ayat:** teks Arab persis yang sudah muncul lebih dulu (di surah sama atau sebelumnya) tidak ditulis lagi.
- **Klausa:** teks Arab persis yang sama tidak ditulis lagi.
- Basmalah di awal ayat 1 (kecuali surah 1 dan 9) dibuang; basmalah hanya dihitung di 1:1.

| | Jumlah |
|---|---|
| Ayat (mushaf Hafs) | 6.236 |
| Kata (token) | 77.800 |
| Kata unik (dengan harakat) | 18.242 |
| Ayat unik | 6.059 (177 duplikat) |
| Klausa unik | 9.655 |

## Satu sumber data

Semua isi website berasal dari `Arabic word list/Kosakata_Arab.xlsx`.

| Sheet | Dipakai untuk |
|---|---|
| `Kosakata` | Tab Kata: satu baris per kata unik |
| `Ayat` | Tab Ayat dan contoh ayat di Kata |
| `Klausa` | Tab Klausa |
| `Surah` | Daftar 114 surah (nama, tempat turun, jumlah) |
| `Progress` | Rekap hafalan di Excel (rumus dari kolom Status). Tidak dipakai website |

Kolom penting sheet `Kosakata`: `No Surah`, `Indonesia`, `Transliterasi`, `Arab`, `Frekuensi`, `Ref pertama`, `Arti alternatif`, `Sumber arti`, `Keterangan`, `Contoh 1..3 Ref` (mis. `2:255`, diambil dari sheet `Ayat`), `Status` (Belum / Sedang / Hafal, hanya Excel), `Catatan edit` (hanya Excel). Kolom dicari berdasarkan nama header, bukan posisi.

**Kolom `Sumber arti`:** `selaras` dan `bentuk-sama` = arti dari sumber, cocok dengan kata Tanzil. `ejaan-beda (cek)` dan `manual (cek)` = perlu Anda periksa; di website tampil sebagai "Arti ini perlu dicek".

## Struktur

```
.
├── README.md
├── .github/workflows/build.yml     ← GitHub Action: build otomatis saat isi folder "Arabic word list" berubah
├── docs/                           ← hasil build (jangan diedit manual)
│   ├── index.html                  ← kamus + latihan
│   ├── deck.html                   ← Belajar Surah (latihan per surah/bagian)
│   ├── data.json                   ← kata
│   ├── ayat.json                   ← ayat unik
│   └── klausa.json                 ← klausa unik
└── Arabic word list/
    ├── Kosakata_Arab.xlsx          ← sumber data (edit file ini)
    ├── build.py                    ← program: Excel → website
    ├── template.html               ← kerangka halaman kamus + latihan
    ├── deck_template.html          ← kerangka halaman Belajar Surah (CSS-nya mengikuti template.html)
    ├── sumber/                     ← salinan asli sumber teks (Tanzil), jangan diubah
    │   ├── quran-simple.txt
    │   └── id.indonesian.txt
    └── tools/
        ├── prep_data.py            ← menyusun Excel dari sumber (sekali jalan)
        └── translit.py             ← transliterasi Latin otomatis per kata
```

## Cara update

1. Edit `Arabic word list/Kosakata_Arab.xlsx` (mis. koreksi arti kata bertanda "cek").
2. Upload file itu ke repo (**Add file → Upload files** di GitHub, atau `git push`).
3. Tunggu sekitar 1–2 menit. GitHub Action membuat ulang `docs/` dan meng-commit hasilnya.
4. Buka halaman GitHub Pages, lalu refresh. Cek status build di tab **Actions**.

## Menjalankan di komputer sendiri

```
pip install openpyxl
python "Arabic word list/build.py"
python -m http.server --directory docs 8000
```

Lalu buka `http://localhost:8000`. Website memuat data lewat `fetch`, jadi **tidak bisa** dibuka dengan klik dua kali `index.html`.

Menyusun ulang Excel dari sumber (hanya bila sumber berubah; butuh database `wordbyword.db` dari repo SadaqaWorks):

```
python "Arabic word list/tools/prep_data.py" "Arabic word list/sumber/quran-simple.txt" "Arabic word list/sumber/id.indonesian.txt" wordbyword.db "Arabic word list/Kosakata_Arab.xlsx"
```

Peringatan: menjalankan `prep_data.py` menimpa semua edit manual di Excel.

## Mengaktifkan GitHub Pages (sekali saja)

1. Repo → **Settings → Pages**.
2. **Source: Deploy from a branch**, branch `main`, folder `/docs`, lalu **Save**.
3. Di **Settings → Actions → General**, pastikan **Workflow permissions** = **Read and write permissions**.

## Sumber data dan atribusi

| Data | Sumber | Ketentuan |
|---|---|---|
| Teks Arab | [Tanzil Project](https://tanzil.net/download/) (`quran-simple`) | Salinan boleh disebarkan **tanpa diubah**, wajib mencantumkan sumber dan tautan ke tanzil.net. Catatan hak cipta Tanzil ada di awal file `sumber/*.txt` dan tidak boleh dihapus |
| Terjemah ayat | Kementerian Agama RI (edisi 2010) lewat [Tanzil](https://tanzil.net/trans/) | Sama seperti di atas |
| Arti per kata | [SadaqaWorks/Word-By-Word-Quran-Android](https://github.com/SadaqaWorks/Word-By-Word-Quran-Android) (tabel `bywords`) | Lisensi repo MIT. Dipetakan ke kata Tanzil dengan pencocokan otomatis, sebagian kecil dilengkapi manual (ditandai "cek") |
| Nama surah | Tabel `surah_name` di repo yang sama | MIT |
| Transliterasi | `tools/translit.py` | Dibuat mesin, **belum ditinjau** pembaca Al-Quran yang kompeten. Tidak mencerminkan hukum bacaan antar kata (idgham, iqlab, ikhfa, wasal) |

Daftar kata, ayat, dan klausa di sini adalah **turunan** dari teks Tanzil (disaring dan diurutkan), bukan perubahan isi ayat. Atribusi Tanzil tampil di footer website dan di setiap file JSON (`sumber`).

## Catatan

- Repo ini publik agar GitHub Pages gratis.
- Proyek belajar pribadi. Untuk menyebarkan teks Al-Quran di Indonesia secara luas, perlu dicek ketentuan tashih Lajnah Pentashihan Mushaf Al-Qur'an (Kemenag), yang terpisah dari urusan hak cipta.
- Jangan mengedit `docs/` secara manual; file itu ditimpa setiap build.
- Untuk mengubah tampilan kamus dan latihan, edit `template.html`; untuk Belajar Surah, edit `deck_template.html`. Tombol menu ("Kamus", "Latihan", "Belajar Surah") ada di kedua file, jadi ubah keduanya bila menu berubah. Ukuran bagian diatur lewat `var PART=20;` di `deck_template.html`.
- Riwayat perubahan Excel tersimpan di git.
