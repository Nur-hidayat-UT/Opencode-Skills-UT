# Dokumentasi Skrip DY Makalah

Dua skrip mandiri untuk mengolah `.docx`. Keduanya tidak membutuhkan Microsoft
Word dan dapat dijalankan dari terminal mana pun.

| Skrip | Fungsi | Dependensi |
| --- | --- | --- |
| [`rapikan.py`](#rapikanpy) | Periksa dan rapikan tata letak makalah `.docx` | Python 3.8+ (stdlib) |
| [`sitasi.py`](#sitasipy) | Ubah placeholder `[[FN:n]]` menjadi footnote DOCX | Python 3.8+, `lxml` |

Asal kode dan atribusi: lihat [`CREDITS.md`](../../CREDITS.md).

---

## `rapikan.py`

Perapi makalah `.docx` yang sudah jadi. Tata letak dibangun ulang, sedangkan
kata-kata naskah dipertahankan. Perubahan teks hanya berasal dari daftar
`koreksi`, dan hasil akhirnya **diaudit kata per kata** terhadap file asli.

### Prinsip keselamatan

- Jika hasil audit berbeda dari naskah asli di luar koreksi, file **tidak
  disimpan** (`STATUS=JANGAN_KIRIM`).
- Koreksi yang terlalu jauh dari kata aslinya ditolak (`PERBAIKI_CONFIG`),
  sehingga tidak bisa dipakai untuk menulis ulang isi.
- Skrip dianjurkan dipanggil apa adanya. Semua pengaturan lewat `config.json`;
  tidak perlu mengedit kode.

### Kebutuhan

- Python 3.8+ — hanya pustaka standar.
- Opsional: `soffice`/`libreoffice` dan `pdftotext`. Jika tersedia, skrip
  merender PDF untuk mengisi dan memverifikasi nomor halaman Daftar Isi serta
  memastikan sampul tidak lebih dari satu halaman.

### Perintah

Tanda `[...]` berarti opsional.

#### 1. `periksa` — bedah naskah & deteksi salah ketik

```bash
python3 rapikan.py periksa "makalah.docx"
python3 rapikan.py periksa "makalah.docx" --penuh
```

Menulis `naskah.txt` di direktori kerja, lalu menampilkan:

- jumlah tabel dan gambar;
- isi blok sampul dan daftar judul (h1–h4, caption);
- blok yang perlu dilihat;
- kandidat salah ketik (kata ganda, ketikan acak, ejaan tidak baku, spasi/tanda
  baca, kalimat tanpa titik akhir, kata mirip);
- kata yang hanya muncul sekali;
- istilah di dalam tanda kurung;
- entri Pustaka tanpa huruf miring.

`--penuh` mencetak seluruh isi `naskah.txt` ke layar.

#### 2. `konteks` — lihat sekitar satu kata

```bash
python3 rapikan.py konteks "gerjea" "makalah.docx"
```

Menampilkan maksimal 10 kemunculan beserta 10 kata di kiri dan kanannya. Cocok
untuk memastikan kata yang akan dikoreksi.

#### 3. `baca` — ekspor daftar blok

```bash
python3 rapikan.py baca "makalah.docx"
python3 rapikan.py baca "makalah.docx" --naskah blok.txt
```

Menulis daftar blok berurutan (ID, peran, isi). Berguna untuk mencari ID blok
yang dipakai pada `peran`, `pecah`, atau `gabung`. Output contoh:

```
[12|h1] BAB I PENDAHULUAN
[13|badan] Latar belakang ...
```

Ringkasan yang dicetak: `NASKAH=`, `BLOK=`, `TABEL=`, `GAMBAR=`,
`CATATAN_KAKI=`.

#### 4. `rapikan` — jalankan perapian

```bash
# urutan bebas: config dulu atau docx dulu
python3 rapikan.py rapikan "makalah.docx" config.json --out "folder_asal"
python3 rapikan.py rapikan config.json "makalah.docx" --out "folder_asal"
```

- `--out` = folder keluaran. Default: `/mnt/user-data/outputs` bila ada,
  selain itu direktori kerja saat ini.
- Nama hasil: `<nama_asli>_Rapi.docx`, kecuali diatur lewat `nama_file`.
- `hasil.json` ditulis di direktori kerja saat ini.
- Bila tidak ada path `.docx`, skrip mencari `/mnt/user-data/uploads`.

### `config.json`

Semua kunci opsional. Contoh lengkap:

```json
{
  "margin": "4,3,3,3",
  "daftar_tabel_gambar": "tidak",
  "posisi_nomor": "5",
  "nama_file": "Makalah_Final",
  "koreksi": [
    {"cari": "gerjea", "ganti": "gereja"},
    {"cari": "disini", "ganti": "di sini", "blok": "42", "alasan": "ejaan baku"}
  ],
  "miring": ["network", "machine learning"],
  "miring_pustaka": ["Journal of Media and Religion"],
  "gabung": [[51, 52]],
  "pecah": [
    {"blok": 50, "setelah": "A. Latar Belakang", "peran": ["h2", "badan"]}
  ],
  "peran": {"54": "h2"}
}
```

| Kunci | Tipe | Arti |
| --- | --- | --- |
| `margin` / `margin_cm` | string `"kiri,atas,kanan,bawah"` atau objek | Margin halaman dalam cm. Default `4,3,3,3`. Nilai harus 1–6 cm. |
| `daftar_tabel_gambar` | `"ya"`/`"tidak"`/boolean | Buat halaman Daftar Tabel & Daftar Gambar. Default tidak. |
| `posisi_nomor` | `1`–`6` atau teks | Posisi nomor halaman: 1 kiri atas, 2 tengah atas, 3 kanan atas, 4 kiri bawah, 5 tengah bawah, 6 kanan bawah. Default `5`. |
| `nama_file` | string | Nama file keluaran (akhiran `.docx` ditambahkan otomatis). |
| `koreksi` | array objek | Daftar penggantian teks. Lihat di bawah. |
| `miring` | array string | Istilah yang dimiringkan di badan teks. |
| `miring_pustaka` | array string | Istilah yang dimiringkan khusus di Daftar Pustaka. |
| `gabung` | array pasangan ID | Gabungkan beberapa blok menjadi satu. Contoh `[[51,52]]`. |
| `pecah` | array objek | Pecah satu blok pada teks tertentu. |
| `peran` | objek `ID → peran` | Paksa peran sebuah blok. |

#### Detail `koreksi`

```json
{"cari": "teks_lama", "ganti": "teks_baru", "blok": "42", "alasan": "opsional"}
```

- `cari` dicocokkan sebagai kata utuh (batas kata), bukan potongan kata.
- `blok` opsional; tanpa itu koreksi berlaku di seluruh dokumen.
- `alasan` hanya keterangan, tidak memengaruhi hasil.
- Pengaman: koreksi ditolak bila kemiripan teks < 55% atau selisih jumlah kata
  > 1. Tujuannya mencegah penulisan ulang isi. Koreksi yang ditolak
  dilaporkan sebagai `KOREKSI_DITOLAK`.

#### Detail `pecah`

```json
{"blok": 50, "setelah": "A. Latar Belakang", "peran": ["h2", "badan"]}
```

- `blok` = ID blok asal (angka sebelum titik, mis. `50` untuk `50.1`).
- `setelah` = teks penanda; blok dipotong tepat setelah teks ini.
- `peran` = peran untuk blok kiri dan blok kanan hasil pemecahan.
- Blok baru mendapat ID `<blok>.1`.

#### Peran blok yang dikenal

`h1`, `h2`, `h3`, `h4`, `badan`, `daftar`, `kutipan`, `pustaka`, `sumber`,
`cap_tabel`, `cap_gambar`, `cap_lampiran`, `gambar`, `sampul`, `buang`,
`judul_daftar`.

### Status keluaran

| Status | Arti | Tindakan |
| --- | --- | --- |
| `SIAP_KIRIM` | Audit cocok, file tersimpan. | Ambil teks di antara `---BALASAN---` dan `---AKHIR---`, sebutkan path `FILE=`. |
| `PERBAIKI_CONFIG` | Ada `KOREKSI_DITOLAK` atau `KOREKSI_TIDAK_DITEMUKAN`. | Hapus/perbaiki koreksi terkait, jalankan ulang. |
| `JANGAN_KIRIM` | Isi berbeda dari aslinya (`SELISIH=`). | File tidak disimpan. Periksa koreksi. |
| `CONFIG_ERROR` | `config.json` tidak valid. | Perbaiki JSON. |
| `FILE_TIDAK_ADA` | Tidak ada `.docx` yang bisa diproses. | Berikan path dokumen. |

Catatan: `status_nomor_halaman` di `hasil.json` bisa bernilai `TANPA_RENDER`
bila tidak ada LibreOffice/`pdftotext`. Nomor Daftar Isi baru pasti setelah
field diperbarui di Word.

### `hasil.json`

Berisi ringkasan eksekusi, antara lain:

- `file`, `status_nomor_halaman`, `jumlah_halaman`, `sampul_satu_halaman`;
- `judul_tidak_terdeteksi`, `margin_cm`, `posisi_nomor`;
- `audit_isi` (jumlah kata dan kecocokan), `koreksi_diterapkan`,
  `koreksi_tidak_ditemukan`;
- `entri_daftar_isi_lama_tanpa_halaman`, `halaman_judul_di_daftar_isi`;
- `jumlah` (bab, subbab, subsubbab, tabel, gambar), `daftar_tabel`,
  `daftar_gambar`;
- `judul_kosong_disisipkan`, `pustaka_diurutkan_ulang`, `keterangan_dipindah`,
  `judul_dipisah_otomatis`, `peringatan`.

---

## `sitasi.py`

Membuat **footnote DOCX nyata** dari placeholder `[[FN:n]]`. Cocok untuk
makalah yang mewajibkan catatan kaki, bukan sitasi dalam kurung.

### Kebutuhan

```bash
pip install lxml
```

### Pemakaian

```bash
python3 sitasi.py footnote "makalah.docx" catatan.json --out "folder_asal"
```

- `makalah.docx` — dokumen sumber (tidak diubah).
- `catatan.json` — teks catatan kaki.
- `--out` — folder keluaran; default folder file asli.
- Hasil: `<nama>_Sitasi.docx`.

### Format `catatan.json`

```json
{
  "catatan": {
    "1": "Santoso, A. (2020). Media dan Masyarakat. Jakarta: Prenada.",
    "2": "Wawancara dengan narasumber, 12 Maret 2024."
  }
}
```

Bentuk ringkas tanpa pembungkus `catatan` juga diterima:

```json
{
  "1": "Teks catatan kaki pertama.",
  "2": "Teks catatan kaki kedua."
}
```

### Cara kerja

1. Setiap `[[FN:n]]` di dokumen diganti nomor footnote superscript.
2. Teks catatan ditambahkan ke `word/footnotes.xml` dengan ID berikutnya.
3. Skrip memastikan `[Content_Types].xml`, relasi dokumen, styles footnote, dan
   settings footnote tersedia.
4. Bila tidak ada placeholder yang ditemukan, skrip berhenti dengan pesan
   `Tidak ada placeholder [[FN:...]] yang ditemukan di dokumen`.

Placeholder boleh memakai kunci bebas, misalnya `[[FN:daftar-pustaka]]`,
selama kuncinya ada di `catatan.json`.

### Contoh

Isi dokumen:

```
Pendidikan adalah proses sepanjang hayat.[[FN:1]]
```

Setelah dijalankan:

```
Pendidikan adalah proses sepanjang hayat.¹
```

Footnote 1 berisi teks dari `catatan.json`.

---

## Tips

- Selalu jalankan `periksa` sebelum `rapikan` agar salah ketik ditemukan lebih
  dulu.
- Gunakan `baca` untuk mendapatkan ID blok sebelum memakai `peran`, `pecah`,
  atau `gabung`.
- Uji `koreksi` satu per satu lewat `konteks` bila ragu.
- Bila status `PERBAIKI_CONFIG`, jangan menambah koreksi baru; betulkan yang
  dilaporkan dahulu.
- Simpan hasil di folder yang sama dengan file asli agar mudah dibandingkan.
