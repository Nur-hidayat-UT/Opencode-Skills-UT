---
name: ut-tugas-wajib
description: Mengerjakan tugas wajib Universitas Terbuka (UT) di e-learning/silayar — TMK, Tugas Tutorial, tugas mata kuliah. Gunakan saat pengguna menyebut tugas wajib UT, TMK, Tugas Tutorial, e-learning UT, silayar, atau meminta jawaban tugas mata kuliah UT dalam format Word/PDF.
---

# Tugas Wajib UT (TMK / e-learning / silayar)

Skill ini memandu pengerjaan **tugas wajib** UT, bukan diskusi tuton.
Untuk forum diskusi pakai skill `ut-diskusi-tuton`.

## Format wajib (sesuai acuan pengguna)

1. **Identitas di awal.** Cantumkan nama, NIM, UPBJJ, kode & nama mata
   kuliah, tutorial ke-, dan nama tutor. Bisa pada halaman judul dan/atau
   kop jawaban.
2. **Jawab tugas dengan baik dan benar**, langsung menjawab pertanyaan.
3. **Penomoran.** Jika soal lebih dari satu, beri nomor yang rapi
   (`1.`, `2.`, … lalu `a.`, `b.` bila ada sub-soal). Satu nomor = satu
   bagian berjudul.
4. **Bentuk penyajian.** Utamakan **narasi/paragraf**. Poin diperbolehkan
   tetapi tidak disarankan; gabungan narasi + poin juga boleh. **Jika
   pengguna belum menentukan, tanyakan dulu** pilihan: narasi, poin, atau
   gabungan.
5. **Sumber di akhir.** Wajib ada daftar sumber terpercaya (jurnal, buku,
   modul BMP, dll.) dengan APA 7. **Tanyakan apakah sumber dicarikan oleh
   saya atau disediakan/diunggah oleh pengguna**, agar sumber sinkron dengan
   tugas.

## Prinsip

- Bahasa Indonesia formal sesuai PUEBI/EYD V.
- Sitasi APA edisi ke-7. Rujukan utama = modul BMP yang diberikan pengguna
  (tulis nama modul, penulis, tahun, dan **halaman**).
- Jangan mengarang sumber, data, atau nomor halaman. Jika tidak ada di modul
  dan pengguna tidak menyediakan, katakan belum tersedia.
- Jawaban harus orisinal, terstruktur, dan langsung menjawab pertanyaan.
- Format akhir default: `.docx` (dan `.pdf` bila diminta).

## Langkah kerja

1. **Kumpulkan bahan dan identitas.** Pastikan ada: soal/tugas lengkap,
   rubrik penilaian, modul/BMP terkait, serta nama, NIM, UPBJJ, kode & nama
   MK, tutorial ke-, dan nama tutor. Jika ada yang kurang, tanyakan dulu.
   **Tanyakan mata kuliahnya** bila belum jelas dari soal; jika dari soal
   sudah jelas, tidak perlu bertanya.
2. **Periksa jumlah soal.** Jika lebih dari satu, susun penomoran yang rapi
   sesuai urutan soal dan sub-soal.
3. **Tanyakan gaya penyajian** (bila belum ditentukan): narasi/paragraf,
   poin, atau gabungan. Default: narasi.
4. **Tanyakan sumber**: "Apakah sumber saya carikan, atau Anda yang
   menyediakan/mengunggah?" Tunggu jawaban agar sumber sinkron dengan tugas.
   Jika pengguna meminta dicarikan, gunakan `websearch`/`webfetch` dan
   verifikasi tiap sumber nyata (DOI/URL, penulis, tahun).
5. **Klarifikasi format.** Jika belum ada acuan: jumlah halaman/paragraf,
   bobot per soal, font & spasi, batas kata, nama file, dan perlu/tidaknya
   halaman judul. Jika pengguna memilih "ikuti standar", pakai **Format
   Default** di bawah.
6. **Bedah soal.** Uraikan tiap nomor: kata kerja perintah (jelaskan,
   analisis, bandingkan, hitung), data yang diberikan, dan bobot nilai.
   Tentukan bagian modul yang relevan.
7. **Susun kerangka.** Rencana per nomor sebelum menulis penuh. Untuk soal
   hitungan: rumus → substitusi → hasil → satuan → interpretasi.
8. **Tulis jawaban.** Ikuti penomoran dan gaya penyajian yang dipilih.
   Narasi memakai pola pembuka-pengembangan-penutup. Sitasi setiap klaim
   dari modul/sumber, contoh: (Wardani, 2021, hlm. 45).
9. **Periksa.** Semua nomor terjawab, tidak keluar topik, ejaan PUEBI,
   konsistensi istilah, sumber lengkap di daftar rujukan, tidak ada sumber
   fiktif.
10. **Hasilkan dokumen.** Simpan di folder kerja dengan nama
    `TMK-<KodeMK>-<NIM>-<Nama>.docx`.

## Format Default (jika belum ada acuan)

- Kertas A4, margin 2,5 cm, font Times New Roman 12 pt, spasi 1,5.
- Halaman judul: judul tugas, kode & nama MK, nama, NIM, UPBJJ, tutorial ke-,
  nama tutor, tahun.
- Bagian: **Identitas → Jawaban bernomor → Sumber**.
- Daftar sumber gaya APA 7, urut abjad, menjorok gantung.

## Output dokumen (alat sudah terpasang)

Gunakan perangkat yang tersedia di mesin ini:

- **DOCX** → `python-docx` 1.1.2 (sudah terpasang). Atur A4, margin 2,5 cm,
  Times New Roman 12 pt, spasi 1,5, halaman judul, penomoran, dan daftar
  sumber menggantung lewat `python-docx`.
- **Konversi lain** → converter bawaan OnlyOffice:
  ```bash
  cd /opt/onlyoffice/desktopeditors/converter
  ./x2t input.docx output.odt   # docx↔odt/rtf/txt, html→docx, dll.
  ```
- **PDF** → render HTML ke PDF dengan Brave headless (x2t ke PDF gagal saat
  headless):
  ```bash
  brave-browser --headless --disable-gpu --no-sandbox \
    --print-to-pdf="$(pwd)/hasil.pdf" "file://$(pwd)/hasil.html"
  ```
  Sertakan CSS `@page{size:A4;margin:2.5cm}` dan font Times New Roman pada
  HTML agar hasil PDF sesuai format tugas.
- Jika PDF diminta dari dokumen Word yang sudah ada, buka OnlyOffice
  (GUI) dan ekspor PDF, atau minta pengguna melakukannya.

Jangan mengarang hasil konversi: pastikan file keluaran benar-benar ada
(`ls -la`) sebelum menyatakan selesai.

## Checklist sebelum menyerahkan

- [ ] Identitas (nama, NIM, UPBJJ, MK, tutorial, tutor) lengkap.
- [ ] Soal bernomor rapi bila lebih dari satu.
- [ ] Gaya penyajian sesuai pilihan pengguna (default narasi).
- [ ] Semua nomor soal terjawab dan sesuai bobot.
- [ ] Ada daftar sumber terpercaya di akhir (APA 7).
- [ ] Sumber nyata/terverifikasi, bukan karangan.
- [ ] Tidak ada plagiarisme; kutipan diberi tanda petik dan sitasi.
- [ ] Ejaan dan istilah konsisten (PUEBI).
- [ ] Nama file & identitas benar.
- [ ] Format sesuai rubrik/petunjuk tutor.
