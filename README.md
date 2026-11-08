# Skill UT — Makalah, Tugas Wajib & Diskusi Tuton

Kumpulan skill [opencode](https://opencode.ai) untuk membantu mahasiswa
Universitas Terbuka (UT) mengerjakan **makalah**, **tugas wajib**, dan
**diskusi tuton**.

## Daftar skill

| Skill | Fungsi | Pemicu (kata kunci) |
| --- | --- | --- |
| [`dy-makalah`](dy-makalah/SKILL.md) | Membuat makalah dari nol dan merapikan `.docx` (APA 7, daftar isi, daftar pustaka) | DY Makalah, buat makalah, bikin makalah, tulis makalah, tugas makalah |
| [`ut-tugas-wajib`](ut-tugas-wajib/SKILL.md) | Mengerjakan TMK / Tugas Tutorial / tugas mata kuliah, output `.docx`/`.pdf` | tugas wajib UT, TMK, Tugas Tutorial, silayar, e-learning UT |
| [`ut-diskusi-tuton`](ut-diskusi-tuton/SKILL.md) | Menulis inisiasi/jawaban awal dan tanggapan forum diskusi tuton | diskusi tuton, forum tuton, inisiasi, jawaban diskusi UT |

## Isi setiap skill

Setiap folder berisi satu berkas `SKILL.md` dengan frontmatter (`name`,
`description`) dan panduan kerja lengkap:

- **Format wajib** — struktur jawaban sesuai acuan.
- **Prinsip** — bahasa formal, APA 7, sumber nyata (bukan karangan).
- **Langkah kerja** — dari mengumpulkan bahan sampai menghasilkan dokumen.
- **Checklist** — pemeriksaan sebelum diserahkan.

### `dy-makalah`
Membuat makalah umum dari topik menjadi `.docx` siap kumpul, serta merapikan
makalah yang sudah jadi tanpa mengubah isinya. Dilengkapi dua skrip Python:

- [`scripts/rapikan.py`](dy-makalah/scripts/rapikan.py) — perapi tata letak,
  penomoran halaman, daftar isi/pustaka, dan audit kata per kata.
- [`scripts/sitasi.py`](dy-makalah/scripts/sitasi.py) — pembuat footnote nyata
  dari placeholder `[[FN:n]]`.

Dokumentasi perintah, `config.json`, peran blok, dan status keluaran ada di
[`dy-makalah/scripts/README.md`](dy-makalah/scripts/README.md).

### `ut-tugas-wajib`
Menekankan identitas lengkap (nama, NIM, UPBJJ, kode & nama MK, tutorial ke-,
tutor), penomoran soal yang rapi, gaya penyajian narasi, serta daftar sumber
terpercaya. Format default: A4, margin 2,5 cm, Times New Roman 12 pt, spasi 1,5.

### `ut-diskusi-tuton`
Khusus forum diskusi: diawali kalimat izin menjawab, isi berupa narasi 2–4
paragraf, minimal **3 sumber terpercaya** (jurnal/berita kredibel/buku), dan
sumber di bagian bawah. Membedakan **inisiasi** dan **tanggapan** ke mahasiswa
lain.

## Cara pasang

Salin folder skill ke direktori skill opencode:

```bash
# global (semua proyek)
cp -r dy-makalah ut-tugas-wajib ut-diskusi-tuton ~/.config/opencode/skills/

# atau khusus satu proyek
cp -r dy-makalah ut-tugas-wajib ut-diskusi-tuton .opencode/skills/
```

Restart sesi opencode agar skill terbaca. Setelah itu cukup sebut kata kunci
pemicunya, misalnya:

> "Buatkan makalah tentang literasi digital, 15 halaman."
>
> "Buatkan jawaban tugas wajib UT mata kuliah X"

Untuk `dy-makalah`, pasang juga dependensi skrip sesuai kebutuhan:

```bash
# rapikan.py hanya butuh Python 3.8+ (stdlib)
# sitasi.py butuh lxml
pip install lxml
```

## Struktur repositori

```
.
├── README.md
├── LICENSE
├── CREDITS.md
├── dy-makalah/
│   ├── SKILL.md
│   ├── README.md
│   └── scripts/
│       ├── rapikan.py
│       ├── sitasi.py
│       └── README.md
├── ut-tugas-wajib/
│   └── SKILL.md
└── ut-diskusi-tuton/
    └── SKILL.md
```

## Lisensi

MIT — lihat [`LICENSE`](LICENSE). Kode perapi makalah berasal dari YZ PROJECT
(@yzproject24) lalu dimodifikasi; atribusi lengkap ada di
[`CREDITS.md`](CREDITS.md).
