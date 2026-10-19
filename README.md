# MDs - Ultimate Markdown Converter 🚀

**MDs** adalah alat baris perintah (CLI) berbasis Python untuk:

- mengonversi dokumen (PDF, DOCX) dan file kode ke Markdown (`.md`),
- menggabungkan (*bundle*) seluruh isi folder proyek menjadi satu file Markdown,
- memulihkan file Markdown kembali ke file aslinya.

Ringan dan responsif, dibuat untuk lingkungan terbatas seperti **Termux (Android)** maupun terminal desktop (**Windows CMD / PowerShell**, Linux, macOS).

---

## ✨ Fitur

- 📜 **File → Markdown**
  Konversi satu file. DOCX diubah menjadi Markdown sungguhan (heading, daftar, tabel), PDF diekstrak per halaman, file kode/teks disimpan apa adanya dengan metadata nama asli.
- 📁 **Folder → Markdown (Bundle)**
  Menggabungkan seluruh file teks/kode dalam folder ke satu file `.md`, lengkap dengan:
  - pohon struktur direktori di bagian atas,
  - header path per file dan *syntax highlighting*,
  - filter ekstensi (mis. hanya `.py,.js`),
  - estimasi jumlah token (berguna untuk dikirim ke AI),
  - daftar file yang dilewati beserta alasannya.
- 🔄 **Markdown → File Asli (Unbundle)**
  Memulihkan bundle menjadi struktur folder semula, atau memulihkan file tunggal dengan nama asli dari metadata. Mendukung pilihan lokasi simpan dan konfirmasi sebelum menimpa file.
- 📂 **Navigasi Folder Interaktif**
  Jelajah direktori dengan nomor, tanpa mengetik path. Dilengkapi halaman (`n`/`p`), tombol kembali (`0`), dan pilih folder (`s`).
- ⌨️ **Mode CLI**
  Bisa dipakai lewat perintah (tanpa menu) untuk skrip dan otomasi.
- 🛡️ **Aman**
  Unbundle menolak path berbahaya (`../`, path absolut), folder sampah dan file rahasia tidak ikut dibundle.
- 🎨 **Tampilan terminal berwarna** dengan layar bersih di setiap menu dan keluar cepat (`exit`).

---

## 🛠️ Format yang Didukung

| Kategori | Ekstensi |
| --- | --- |
| Dokumen | `.pdf`, `.docx` |
| Teks & data | `.txt`, `.md`, `.json`, `.yaml`, `.yml`, `.toml`, `.ini`, `.cfg`, `.xml`, `.csv` |
| Web | `.html`, `.css`, `.js`, `.ts`, `.jsx`, `.tsx`, `.php` |
| Bahasa pemrograman | `.py`, `.c`, `.h`, `.cpp`, `.hpp`, `.java`, `.cs`, `.go`, `.rs`, `.rb`, `.sh`, `.sql`, `.kt`, `.swift` |

> **Catatan:** format `.doc` (Word lama) tidak didukung. Simpan ulang sebagai `.docx` terlebih dahulu.

---

## 💻 Prasyarat

- **Python 3.8+**
- Pustaka (lihat `requirements.txt`):
  - `pypdf` untuk ekstraksi teks PDF
  - `python-docx` untuk dokumen Word (butuh `lxml`, terpasang otomatis di desktop)

File teks dan kode tetap bisa dikonversi walau kedua pustaka belum terpasang.

---

## 📥 Instalasi

### Termux (Android)

```bash
pkg update && pkg upgrade -y
pkg install python git python-lxml -y
termux-setup-storage        # izinkan akses ke ~/storage/downloads

git clone https://github.com/ShimaYuu205/md-converter.git
cd md-converter
pip install -r requirements.txt
```

> `python-lxml` dipasang lewat `pkg` karena kompilasi `lxml` lewat `pip` sering gagal di Termux.

### Windows (CMD / PowerShell)

```powershell
git clone https://github.com/ShimaYuu205/md-converter.git
cd md-converter
pip install -r requirements.txt
```

### Linux / macOS

```bash
git clone https://github.com/ShimaYuu205/md-converter.git
cd md-converter
pip install -r requirements.txt
```

---

## 📕 Cara Penggunaan

### Mode interaktif

```bash
python md.py
```

Menu utama:

| Menu | Fungsi |
| --- | --- |
| `1` | Convert file → Markdown |
| `2` | Convert isi folder → Markdown (bundle) |
| `3` | Convert Markdown → file asli / unbundle |
| `0` / `exit` | Keluar |

Navigasi folder: ketik **nomor** untuk masuk folder atau memilih file, `0` untuk kembali, `n` / `p` untuk halaman berikutnya / sebelumnya, `s` untuk memilih direktori saat ini (mode folder), `q` untuk batal.

### Mode CLI

```bash
# Satu file -> Markdown
python md.py file laporan.docx
python md.py file kode.py -o hasil.md

# Folder -> satu file Markdown (bundle)
python md.py bundle ./proyek
python md.py bundle ./proyek -o proyek.md --ext .py,.js --max-kb 500

# Markdown -> file asli (bundle maupun file tunggal)
python md.py restore proyek.md -d ./hasil
python md.py restore proyek.md -d ./hasil --force   # timpa file yang ada
```

Tambahkan `--force` untuk menimpa file hasil yang sudah ada.

---

## 🔁 Alur Kerja

### 1. Convert File → Markdown

1. Pilih menu **1**, lalu pilih file lewat navigasi angka.
2. Isi nama file Markdown (kosongkan untuk default). Ekstensi `.md` ditambahkan otomatis.
3. Konfirmasi `y`.

### 2. Convert Folder → Markdown (Bundle)

1. Pilih menu **2**, masuk ke folder proyek, ketik `s` untuk memilihnya.
2. Isi nama file bundle (default: `nama_folder.md`) dan, bila perlu, filter ekstensi.
3. Seluruh file teks/kode digabung ke satu file Markdown. Ringkasan jumlah file, ukuran, dan estimasi token ditampilkan.

Yang **tidak** ikut dibundle:

- folder tersembunyi dan folder umum yang berat: `.git`, `node_modules`, `venv`, `__pycache__`, `dist`, `build`, dll.
- file tersembunyi (mis. `.env`) agar rahasia tidak ikut tertulis,
- file kunci seperti `package-lock.json` dan `yarn.lock`,
- file biner, file di atas 1 MB, dan bundle lama,
- PDF/DOCX di dalam folder (gunakan menu **1** untuk itu).

### 3. Convert Markdown → File Asli

1. Pilih menu **3**, lalu pilih file `.md`.
2. **Bundle** terdeteksi otomatis: pilih lokasi tujuan, lalu seluruh struktur folder dipulihkan.
3. **File tunggal**: bila ada metadata nama asli, Anda ditanya apakah ingin memakainya.

Pilihan lokasi simpan:

- direktori saat ini,
- direktori lain (dipilih interaktif),
- folder baru di direktori saat ini.

---

## ⚠️ Batasan

- **PDF hasil scan** tidak bisa diekstrak (tidak ada lapisan teks, butuh OCR).
- **Konversi PDF/DOCX bersifat satu arah.** Saat dipulihkan, isinya hanya teks dan disimpan sebagai `.txt`, bukan PDF/DOCX asli.
- Format DOCX yang didukung: heading, daftar, dan tabel. Gambar dan *styling* lanjutan diabaikan.
- Token pada ringkasan hanyalah perkiraan kasar (~4 karakter per token).

---

## ❓ Pemecahan Masalah

| Masalah | Solusi |
| --- | --- |
| `pypdf belum terpasang` / `python-docx belum terpasang` | Jalankan `pip install -r requirements.txt` |
| Folder `~/storage/downloads` tidak ada di Termux | Jalankan `termux-setup-storage`, atau jalankan skrip dari folder tujuan |
| Banner tampil ringkas | Normal, lebar terminal di bawah 80 kolom. Putar layar ke landscape untuk banner penuh |
| Warna tidak muncul di Windows lama | Gunakan Windows Terminal atau PowerShell versi terbaru |

---

## 👤 Penulis

Dibuat dengan ❤️ oleh **ShimaYuu**.

## 📄 Lisensi

Proyek ini dilisensikan di bawah **MIT License**.
