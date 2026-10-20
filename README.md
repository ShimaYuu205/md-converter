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

MDs bisa dipakai dengan dua cara: **menu interaktif** (cocok untuk pemula dan layar sentuh) atau **perintah CLI** (cocok untuk skrip dan otomasi).

### Memulai

```bash
python md.py
```

Menu interaktif langsung terbuka. Navigasi folder dimulai dari `~/storage/downloads` bila ada (Termux), atau dari folder tempat Anda menjalankan skrip.

Ketik `exit` di prompt mana pun untuk keluar, dan `q` di navigasi folder untuk membatalkan pilihan.

### Menu utama

| Menu | Fungsi | Hasil |
| --- | --- | --- |
| `1` | Convert file → Markdown | satu file `.md` di folder yang sama dengan file sumber |
| `2` | Convert isi folder → Markdown (bundle) | satu file `.md` berisi seluruh proyek |
| `3` | Convert Markdown → file asli / unbundle | file atau struktur folder yang dipulihkan |
| `0` | Keluar | |

### Navigasi folder

Pada menu 1, 2, dan 3 Anda memilih file atau folder lewat **nomor**, tanpa mengetik path.

```
[Direktori Saat Ini]: /storage/emulated/0/Download
==================================================
0. [Kembali ke direktori sebelumnya]
1. [DIR] proyek-web
2. [DIR] tugas
3. laporan.docx (.docx)
4. catatan.pdf (.pdf)
--------------------------------------------------
Pilih nomor ('q' batal, 'exit' keluar):
```

| Ketik | Fungsi |
| --- | --- |
| nomor | Masuk ke folder, atau memilih file |
| `0` | Naik ke folder induk |
| `s` (atau `99`) | Memilih folder yang sedang dibuka (hanya saat memilih folder) |
| `n` / `p` | Halaman berikutnya / sebelumnya (20 item per halaman) |
| `q` | Membatalkan dan kembali ke menu utama |
| `exit` | Keluar dari program |

Folder dan file tersembunyi (diawali titik) tidak ditampilkan. Pada menu 1 hanya file yang didukung yang muncul, dan pada menu 3 hanya file `.md`.

---

### Menu 1: File → Markdown

**Langkah:**

1. Pilih menu `1`, lalu cari dan pilih file.
2. Masukkan nama file Markdown, atau tekan Enter untuk memakai nama default (nama file sumber + `.md`). Jika Anda tidak menulis `.md`, ekstensi ditambahkan otomatis.
3. Jika nama hasil sudah ada, Anda ditanya apakah ingin menimpanya.
4. Konfirmasi dengan `y`.

Hasil disimpan **di folder yang sama dengan file sumber**. Ringkasan ukuran dan estimasi token ditampilkan setelah selesai.

**Apa yang terjadi pada tiap jenis file:**

| Jenis file | Hasil di Markdown |
| --- | --- |
| Kode / teks (`.py`, `.html`, `.txt`, dll.) | Isi disalin apa adanya |
| `.docx` | Heading menjadi `#`, daftar menjadi `-` / `1.`, tabel menjadi tabel Markdown |
| `.pdf` | Teks per halaman, diberi penanda `<!-- Halaman N -->` |
| `.doc` | Ditolak. Simpan ulang sebagai `.docx` dulu |

**Metadata nama asli.** Baris pertama hasil selalu berisi nama file sumber:

```
<!-- META_NAME: laporan.docx -->
```

Baris ini dipakai menu 3 untuk menyarankan nama asli saat file dipulihkan. Anda boleh menghapusnya bila tidak diperlukan.

> Jika file sumber sendiri berakhiran `.md`, hasil disimpan sebagai `namafile_converted.md` agar file asli tidak tertimpa.

---

### Menu 2: Folder → Markdown (Bundle)

**Langkah:**

1. Pilih menu `2`, masuk ke folder proyek, lalu ketik `s` untuk memilihnya.
2. Masukkan nama file bundle, atau tekan Enter untuk default `nama_folder.md`.
3. (Opsional) Isi **filter ekstensi**, misalnya `.py,.js` untuk hanya menggabungkan file Python dan JavaScript. Kosongkan untuk menggabungkan semua jenis file yang didukung.
4. Konfirmasi dengan `y`.

File bundle disimpan **di dalam folder yang dipilih**. Saat dijalankan ulang, bundle lama otomatis dilewati sehingga tidak ikut tergabung.

**Contoh isi file bundle:**

````markdown
# Bundle Proyek: proyek-web

> Dibuat oleh MDs v2.0 pada 2026-10-20 10:30 | 3 file

## Struktur Direktori

```text
proyek-web/
├── css/
│   └── style.css
├── index.html
└── script.js
```

---

## File: `css/style.css`

```css
body { margin: 0; }
```

---

## File: `index.html`

```html
<h1>Halo</h1>
```
````

Jika sebuah file berisi blok kode (``` ) di dalamnya, MDs otomatis memakai pagar yang lebih panjang (misalnya ```` ) sehingga bundle tetap valid dan bisa dipulihkan dengan benar.

**Yang dilewati otomatis:**

- folder tersembunyi dan folder berat: `.git`, `node_modules`, `venv`, `__pycache__`, `dist`, `build`, dll.
- file tersembunyi seperti `.env`, supaya data rahasia tidak ikut tertulis,
- file kunci: `package-lock.json`, `yarn.lock`, `poetry.lock`, dll.,
- file biner dan file teks di atas 1 MB,
- PDF/DOCX di dalam folder (konversi lewat menu 1).

Setelah selesai, MDs menampilkan jumlah file, ukuran, **estimasi token**, dan daftar file yang dilewati beserta alasannya. Di Termux, Anda juga ditawari untuk menyalin hasil ke clipboard.

---

### Menu 3: Markdown → File Asli

**Langkah:**

1. Pilih menu `3`, lalu pilih file `.md`.
2. MDs mengenali jenis file secara otomatis (bundle atau file tunggal).
3. Pilih **lokasi penyimpanan**:
   - `1` direktori saat ini (folder tempat file `.md` berada),
   - `2` direktori lain, dipilih lewat navigasi folder,
   - `3` buat folder baru di direktori saat ini.
4. Konfirmasi dengan `y`.

**Jika file adalah bundle proyek**, seluruh struktur folder dipulihkan, termasuk subfolder. Jika ada file yang sudah ada di tujuan, Anda diberi pilihan:

- `y` untuk menimpa semuanya,
- `n` untuk melewati file yang sudah ada dan hanya menulis yang baru.

Setiap file ditampilkan statusnya: `DIPULIHKAN`, `DILEWATI`, atau `DITOLAK`.

**Jika file adalah file tunggal**, MDs membaca `META_NAME` dan bertanya apakah Anda ingin memakai nama aslinya. Jika tidak ada metadata, Anda mengetik nama file sendiri (misalnya `index.html`).

**Keamanan:** entri bundle dengan path berbahaya (`../../sesuatu` atau path absolut seperti `/etc/x`) akan **ditolak**, sehingga file tidak bisa ditulis keluar dari folder tujuan.

> Hasil konversi PDF/DOCX tidak bisa dikembalikan menjadi PDF/DOCX. Saat dipulihkan, isinya disimpan sebagai `.txt`.

---

### Mode CLI (tanpa menu)

```bash
python md.py file    <path>   [-o OUTPUT] [--force]
python md.py bundle  <folder> [-o OUTPUT] [--ext EKSTENSI] [--max-kb N] [--force]
python md.py restore <file.md> [-d TUJUAN] [--force]
```

| Opsi | Berlaku untuk | Fungsi |
| --- | --- | --- |
| `-o`, `--output` | `file`, `bundle` | Lokasi dan nama file `.md` hasil |
| `--ext` | `bundle` | Filter ekstensi, dipisah koma, mis. `.py,.js` |
| `--max-kb` | `bundle` | Lewati file teks lebih besar dari N KB (default 1000) |
| `-d`, `--dest` | `restore` | Folder tujuan (default: folder file `.md`) |
| `--force` | semua | Timpa file yang sudah ada tanpa bertanya |

Tanpa `--force`, perintah akan berhenti dengan pesan error jika file hasil sudah ada. Kode keluar `0` berarti berhasil dan `1` berarti gagal, sehingga aman dipakai di skrip.

**Contoh:**

```bash
# Ubah dokumen Word menjadi Markdown
python md.py file laporan.docx

# Bundle proyek Python saja, maksimal 500 KB per file
python md.py bundle ./proyek -o proyek.md --ext .py --max-kb 500

# Pulihkan bundle ke folder baru, timpa bila ada
python md.py restore proyek.md -d ./hasil --force
```

---

## 💡 Contoh Kasus Penggunaan

| Tujuan | Cara |
| --- | --- |
| Mengirim seluruh kode proyek ke AI dalam satu pesan | Menu 2 atau `bundle`, lalu salin / unggah file `.md`. Perhatikan estimasi token |
| Memindahkan proyek lewat satu file (chat, email, catatan) | Bundle di perangkat A, restore di perangkat B |
| Menyalin kode yang dihasilkan AI dalam format bundle ke folder nyata | Simpan jawaban sebagai `.md`, lalu menu 3 atau `restore` |
| Mengubah dokumen Word atau PDF menjadi teks Markdown | Menu 1 atau `file` |
| Mencadangkan proyek kecil sebagai satu dokumen yang mudah dibaca | Menu 2 |

## 🧭 Tips

- Gunakan **filter ekstensi** agar bundle lebih kecil dan relevan (misalnya hanya `.py`).
- Cek **estimasi token** sebelum mengirim bundle ke AI. Angkanya hanya perkiraan kasar.
- Bundle bisa dibuka dan dibaca langsung di editor Markdown, karena strukturnya rapi dan berpohon.
- Jangan menaruh data rahasia di file biasa (bukan tersembunyi) jika bundle akan dibagikan. Periksa isinya dulu.
- Di layar sempit (Termux portrait), banner tampil ringkas. Putar ke landscape untuk banner penuh.

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
