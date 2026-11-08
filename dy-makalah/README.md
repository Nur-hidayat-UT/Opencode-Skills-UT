# DY Makalah — Pembuat & Perapi Makalah `.docx`

Skill [opencode](https://opencode.ai) untuk membuat makalah dari nol (riset,
kerangka, isi, sitasi APA 7, DOCX rapi) sekaligus merapikan makalah `.docx`
yang sudah jadi **tanpa mengubah kata-kata naskah**.

- `SKILL.md` — panduan kerja untuk AI (alur pembuatan, format default, checklist).
- `scripts/rapikan.py` — perapi tata letak DOCX + pemeriksa salah ketik.
- `scripts/sitasi.py` — pembuat footnote nyata dari placeholder `[[FN:n]]`.

## Asal & kredit

Skrip perapi berasal dari **YZMakalah V2 — YZ PROJECT (@yzproject24)** yang
kemudian dimodifikasi, dibersihkan, dan didokumentasikan ulang sebagai proyek
sumber terbuka di repositori ini. Tautan proyek asal:
<https://lynk.id/yzproject>. Lihat juga [`CREDITS.md`](../CREDITS.md).

Prinsip utama skrip: **tata letak dibangun ulang, isi naskah tidak disentuh.**
Satu-satunya perubahan teks adalah daftar `koreksi` yang ditulis eksplisit di
`config.json`, lalu hasilnya diaudit kata per kata terhadap file asli.

## Fitur

- Merapikan margin, font, spasi, indent, dan penomoran halaman.
- Menyusun ulang Daftar Isi, Daftar Tabel, dan Daftar Gambar.
- Mengurutkan Daftar Pustaka dan menyeragamkan huruf miring istilah asing.
- Menggabung/memecah paragraf dan memaksa peran blok (judul, badan, tabel, dll.).
- Membetulkan salah ketik yang pasti lewat daftar `koreksi`.
- Audit kata per kata: file **tidak disimpan** jika isi berubah di luar koreksi.
- Membuat footnote DOCX nyata dari placeholder `[[FN:n]]` (`sitasi.py`).

## Kebutuhan

| Skrip | Kebutuhan |
| --- | --- |
| `rapikan.py` | Python 3.8+ (hanya pustaka standar) |
| `sitasi.py` | Python 3.8+ dan `lxml` (`pip install lxml`) |
| Opsional | `soffice`/`libreoffice` + `pdftotext` untuk mengisi & memverifikasi nomor halaman Daftar Isi |

Tanpa LibreOffice/`pdftotext`, perapi tetap berjalan, tetapi status nomor
halaman menjadi `TANPA_RENDER`. Buka file di Word lalu klik **Yes/Ya** saat
diminta memperbarui field agar nomor Daftar Isi menyesuaikan.

## Instalasi

```bash
# global (semua proyek)
cp -r dy-makalah ~/.config/opencode/skills/

# khusus satu proyek
cp -r dy-makalah .opencode/skills/
```

Restart sesi opencode. Jika dipasang khusus proyek, sesuaikan path skrip pada
`SKILL.md` menjadi `.opencode/skills/dy-makalah`.

## Cara pakai lewat AI

Cukup sebut pemicunya, misalnya:

> "Buatkan makalah tentang literasi digital, 15 halaman, format kampus saya."
>
> "Rapikan makalah ini." (lampirkan `.docx`)

AI akan mengumpulkan spesifikasi, menyusun kerangka, menulis isi, membuat
DOCX, lalu menjalankan perapi dan memverifikasi hasilnya.

## Cara pakai skrip langsung

Dokumentasi lengkap tiap perintah, `config.json`, peran blok, dan status ada di
[`scripts/README.md`](scripts/README.md).

```bash
# periksa naskah + deteksi salah ketik
python3 scripts/rapikan.py periksa "makalah.docx"

# lihat konteks satu kata
python3 scripts/rapikan.py konteks "disini" "makalah.docx"

# rapikan (config.json harus disiapkan lebih dulu)
python3 scripts/rapikan.py rapikan "makalah.docx" config.json --out "folder_asal"

# buat footnote nyata dari placeholder [[FN:1]]
python3 scripts/sitasi.py footnote "makalah.docx" catatan.json --out "folder_asal"
```

Contoh `config.json` minimal:

```json
{
  "margin": "4,3,3,3",
  "daftar_tabel_gambar": "tidak",
  "posisi_nomor": "5",
  "koreksi": [{"cari": "gerjea", "ganti": "gereja"}],
  "miring": ["network"],
  "miring_pustaka": ["Journal of Media and Religion"]
}
```

## Struktur

```
dy-makalah/
├── SKILL.md
├── README.md
└── scripts/
    ├── rapikan.py
    ├── sitasi.py
    └── README.md
```

## Lisensi

MIT — lihat [`LICENSE`](../LICENSE). Kredit kode asal ada di
[`CREDITS.md`](../CREDITS.md).
