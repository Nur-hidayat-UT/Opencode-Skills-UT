# Skill UT — Tugas Wajib & Diskusi Tuton

Kumpulan skill [opencode](https://opencode.ai) untuk membantu mahasiswa Universitas Terbuka (UT) mengerjakan **tugas wajib** dan **diskusi tuton** di e-learning/silayar.

## Daftar skill

| Skill | Fungsi | Pemicu (kata kunci) |
| --- | --- | --- |
| [`ut-tugas-wajib`](ut-tugas-wajib/SKILL.md) | Mengerjakan TMK / Tugas Tutorial / tugas mata kuliah, output `.docx`/`.pdf` | tugas wajib UT, TMK, Tugas Tutorial, silayar, e-learning UT |
| [`ut-diskusi-tuton`](ut-diskusi-tuton/SKILL.md) | Menulis inisiasi/jawaban awal dan tanggapan forum diskusi tuton | diskusi tuton, forum tuton, inisiasi, jawaban diskusi UT |

## Isi setiap skill

Setiap folder berisi satu berkas `SKILL.md` dengan frontmatter (`name`, `description`) dan panduan kerja lengkap:

- **Format wajib** — struktur jawaban sesuai acuan.
- **Prinsip** — bahasa formal, APA 7, sumber nyata (bukan karangan).
- **Langkah kerja** — dari mengumpulkan bahan sampai menghasilkan dokumen.
- **Checklist** — pemeriksaan sebelum diserahkan.

### `ut-tugas-wajib`
Menekankan identitas lengkap (nama, NIM, UPBJJ, kode & nama MK, tutorial ke-, tutor), penomoran soal yang rapi, gaya penyajian narasi, serta daftar sumber terpercaya. Format default: A4, margin 2,5 cm, Times New Roman 12 pt, spasi 1,5.

### `ut-diskusi-tuton`
Khusus forum diskusi: diawali kalimat izin menjawab, isi berupa narasi 2–4 paragraf, minimal **3 sumber terpercaya** (jurnal/berita kredibel/buku), dan sumber di bagian bawah. Membedakan **inisiasi** dan **tanggapan** ke mahasiswa lain.

## Cara pasang

Salin folder skill ke direktori skill opencode:

```bash
# global (semua proyek)
cp -r ut-tugas-wajib ut-diskusi-tuton ~/.config/opencode/skills/

# atau khusus satu proyek
cp -r ut-tugas-wajib ut-diskusi-tuton .opencode/skills/
```

Restart sesi opencode agar skill terbaca. Setelah itu cukup sebut kata kunci pemicunya, misalnya:

> "Buatkan jawaban tugas wajib UT mata kuliah X"

## Struktur repositori

```
.
├── README.md
├── ut-tugas-wajib/
│   └── SKILL.md
└── ut-diskusi-tuton/
    └── SKILL.md
```

## Lisensi

Bebas dipakai dan dimodifikasi untuk keperluan belajar.
