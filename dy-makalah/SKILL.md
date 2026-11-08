---
name: dy-makalah
description: Membuat makalah secara umum dari nol (riset, kerangka, isi, sitasi APA 7, DOCX rapi) dan merapikan makalah .docx yang sudah jadi. Gunakan saat pengguna menyebut DY Makalah, buat makalah, bikin makalah, tulis makalah, makalah umum, tugas makalah, atau meminta makalah Word lengkap dengan daftar isi dan daftar pustaka.
---

# DY Makalah — Pembuat & Perapi Makalah

Skill untuk membuat makalah secara umum: dari topik menjadi dokumen `.docx` siap
kumpul. Bisa juga merapikan makalah yang sudah jadi tanpa mengubah isinya.

## ATURAN UNTUK AI

1. Jangan mengarang sumber, data, kutipan, atau nomor halaman. Sumber harus nyata dan terverifikasi.
2. Jangan menambah bagian yang tidak diminta. Ikuti template kampus bila pengguna memberikannya.
3. Satu pertanyaan per pesan. Jangan menampilkan kode, config, atau log ke pengguna.
4. Tanyakan hal yang belum jelas sebelum menulis; jangan berasumsi soal topik, jumlah halaman, atau format.
5. Jangan menulis atau mengubah skrip perapi sendiri. Panggil `scripts/rapikan.py` apa adanya.

## ALUR PEMBUATAN MAKALAH

**1. Kumpulkan spesifikasi.** Tanyakan yang belum ada, satu per satu:
- topik/judul dan mata kuliah;
- jenis karya (makalah biasa, artikel ilmiah, laporan, review);
- target panjang (jumlah halaman atau kata) dan jumlah bab;
- acuan format kampus (template/juknis bila ada);
- sumber: dicarikan oleh AI atau disediakan pengguna;
- identitas penulis (nama, NIM, prodi, kampus, dosen, tahun) bila diperlukan pada sampul.

Jika pengguna memilih "terserah" atau "standar", pakai **Format Default** di bawah.

**2. Susun kerangka.** Tampilkan kerangka bab dan subbab, lalu minta persetujuan
sebelum menulis penuh. Sesuaikan kedalaman dengan panjang yang diminta.

**3. Riset sumber.** Bila sumber dicarikan, gunakan `websearch`/`webfetch`.
Verifikasi tiap sumber: penulis, tahun, judul, penerbit/jurnal, DOI/URL.
Utamakan jurnal, buku, modul/BMP, dan sumber resmi. Jangan mencantumkan sumber
yang tidak benar-benar dibaca/ditemukan.

**4. Tulis isi.** Bahasa Indonesia formal sesuai PUEBI/EYD V. Paragraf mengalir
(pembuka–pengembangan–penutup), bukan sekadar daftar poin. Setiap klaim dari
sumber diberi sitasi APA 7, contoh: (Santoso, 2020, hlm. 45). Kutipan langsung
diberi tanda petik. Hindari plagiarisme dan parafrasa berlebihan yang mengaburkan makna.

**5. Susun DOCX** (lihat bagian Alat). Struktur default:
- Halaman sampul
- Kata Pengantar
- Daftar Isi
- Daftar Tabel / Daftar Gambar (hanya jika ada tabel/gambar)
- Bab I Pendahuluan: Latar Belakang, Rumusan Masalah, Tujuan, Manfaat
- Bab II Pembahasan: beberapa subbab sesuai kerangka
- Bab III Penutup: Kesimpulan, Saran
- Daftar Pustaka

**6. Rapikan dan verifikasi.** Setelah DOCX dibuat, jalankan perapi (bagian
Merapikan) agar margin, penomoran halaman, daftar isi, dan daftar pustaka
konsisten. Lalu pastikan file benar-benar ada dan bisa dibuka (`ls -la`,
`unzip -t`), jumlah kata/halaman wajar, dan semua bab terisi.

**7. Serahkan.** Sebutkan path file hasil, ringkasan isi, dan daftar sumber.
Jika diminta PDF, konversi sesuai bagian Alat.

## FORMAT DEFAULT (bila tidak ada acuan kampus)

- A4, margin Kiri 4 cm; Atas/Kanan/Bawah 3 cm (atau 4-3-3-3 sesuai permintaan).
- Times New Roman 12 pt, spasi 1,5, rata kiri-kanan, paragraf menjorok 1,25 cm.
- Judul bab: kapital, tebal, rata tengah; subbab: tebal, rata kiri.
- Nomor halaman: kanan bawah (default), angka Arab mulai dari Pendahuluan.
- Daftar Pustaka: APA 7, urut abjad, menjorok gantung 1,25 cm.
- Tabel/gambar diberi nomor dan judul, sumber ditulis di bawahnya.

## ALAT YANG TERSEDIA

- **DOCX** → `python-docx` 1.1.2 (terpasang). Atur A4, margin, font, spasi,
  heading, nomor halaman, dan hanging indent lewat Python.
- **Perapi DOCX** → `scripts/rapikan.py` (pustaka standar Python; lihat bagian Merapikan).
- **Konversi format** → converter OnlyOffice:
  ```bash
  cd /opt/onlyoffice/desktopeditors/converter
  ./x2t input.docx output.odt   # docx↔odt/rtf/txt, html→docx, dll.
  ```
- **PDF** → render HTML ke PDF dengan Brave headless:
  ```bash
  brave-browser --headless --disable-gpu --no-sandbox \
    --print-to-pdf="$(pwd)/hasil.pdf" "file://$(pwd)/hasil.html"
  ```
  Sertakan CSS `@page{size:A4;margin:3cm}` dan Times New Roman pada HTML.
  Untuk DOCX yang sudah ada, ekspor PDF lewat OnlyOffice (GUI) atau minta pengguna.

Jangan mengarang hasil konversi: pastikan file keluaran benar-benar ada sebelum menyatakan selesai.

## MERAPKAN MAKALAH YANG SUDAH JADI

Perapi mempertahankan isi kata per kata; perubahan teks hanya dari daftar
`koreksi` yang pasti salah ketik.

**A. Tanya 3 hal** (gunakan tool `question` untuk pilihan, satu per pesan):
1. 📐 Margin: `4,3,3,3` atau `4,4,3,3` (Kiri, Atas, Kanan, Bawah).
2. 🖼️ Halaman Daftar Tabel/Gambar: `Tidak` atau `Ya`.
3. 🔢 Posisi nomor halaman: Kiri/Tengah/Kanan × Atas/Bawah (1–6).
   "default"/"lewati" = `4,3,3,3`, Tidak, nomor 5.

**B. Periksa naskah.**

```bash
mkdir -p /tmp/opencode/dy-makalah
python3 ~/.config/opencode/skills/dy-makalah/scripts/rapikan.py periksa "PATH.docx"
```

Jalankan dengan `workdir` = `/tmp/opencode/dy-makalah`. Bila skill dipasang
khusus proyek, ganti path dengan `.opencode/skills/dy-makalah`.
Untuk melihat konteks kata:

```bash
python3 ~/.config/opencode/skills/dy-makalah/scripts/rapikan.py konteks KATA "PATH.docx"
```

Betulkan hanya salah ketik yang pasti (huruf tertukar/kurang/lebih, kata ganda,
ketikan acak, ejaan baku seperti `disini` → `di sini`). Jangan ubah nama,
istilah asing, NIM, penanda, pilihan kata, atau isi Daftar Pustaka.

**C. Tulis `config.json` di `/tmp/opencode/dy-makalah`, lalu jalankan.**

```json
{"margin":"4,3,3,3","daftar_tabel_gambar":"tidak","posisi_nomor":"5",
 "koreksi":[{"cari":"gerjea","ganti":"gereja"}],
 "miring":["network"],"miring_pustaka":["Journal of Media and Religion"]}
```

```bash
python3 ~/.config/opencode/skills/dy-makalah/scripts/rapikan.py rapikan "PATH.docx" config.json --out "FOLDER_ASAL"
```

Hasil disimpan di samping file asli dengan akhiran `_Rapi`. Opsional:
`"pecah":[{"blok":50,"setelah":"A. Latar Belakang","peran":["h2","badan"]}]`,
`"gabung":[[51,52]]`, `"peran":{"54":"h2"}`. Peran: h1, h2, h3, h4, badan,
daftar, kutipan, pustaka, cap_tabel, cap_gambar, sumber.

**D. Ikuti `STATUS`.**
- `SIAP_KIRIM`: salin teks antara `---BALASAN---` dan `---AKHIR---` persis, lalu sebutkan path `FILE`.
- `PERBAIKI_CONFIG`: buang/perbaiki koreksi yang disebut, ulangi.
- `JANGAN_KIRIM`: hasil berbeda dari aslinya; perbaiki koreksi, ulangi sekali. Jika tetap gagal, jelaskan dan jangan kirim file.

Catatan: nomor halaman Daftar Isi terisi otomatis bila `soffice`/`libreoffice` dan
`pdftotext` tersedia. Di mesin ini `soffice` tidak ada, jadi status bisa
`TANPA_RENDER`; minta pengguna memperbarui field di Word (`Yes/Ya`).

## CHECKLIST SEBELUM MENYERAHKAN

- [ ] Topik, struktur, dan panjang sesuai permintaan.
- [ ] Semua bab terisi; tidak ada bagian kosong atau "TODO".
- [ ] Sitasi APA 7 di teks dan daftar pustaka konsisten.
- [ ] Sumber nyata dan terverifikasi, bukan karangan.
- [ ] Ejaan dan istilah konsisten (PUEBI).
- [ ] Margin, font, spasi, dan nomor halaman sesuai format.
- [ ] File `.docx` benar-benar ada dan bisa dibuka.
