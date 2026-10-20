# MDs - Ultimate Markdown Converter 🚀

**MDs** adalah alat baris perintah (CLI) berbasis Python untuk:

- mengonversi dokumen (PDF, DOCX) dan file kode ke Markdown (`.md`),
- menggabungkan (*bundle*) seluruh isi folder proyek menjadi satu file Markdown,
- memulihkan file Markdown kembali ke file aslinya.

Ringan dan responsif, dibuat untuk lingkungan terbatas seperti **Termux (Android)** maupun terminal desktop (**Windows CMD / PowerShell**, Linux, macOS). Dijalankan lewat perintah **`MDc`**, baik sebagai menu interaktif maupun perintah CLI.

---

## 📑 Daftar Isi

- [Fitur](#-fitur)
- [Format yang Didukung](#️-format-yang-didukung)
- [Instalasi](#-instalasi)
- [Cara Memulai](#-cara-memulai)
- [Referensi Perintah Lengkap](#-referensi-perintah-lengkap)
- [Pola --ignore](#-pola---ignore)
- [Contoh Kasus Penggunaan](#-contoh-kasus-penggunaan)
- [Memperbarui ke Versi Terbaru](#-memperbarui-ke-versi-terbaru)
- [Batasan](#️-batasan)
- [Pemecahan Masalah](#-pemecahan-masalah)
- [Riwayat Versi](#-riwayat-versi)

---

## ✨ Fitur

- 📜 **File → Markdown**
  DOCX diubah menjadi Markdown sungguhan (heading, daftar, tabel), PDF diekstrak per halaman, file kode/teks disimpan apa adanya dengan metadata nama asli.
- 📁 **Folder → Markdown (Bundle)**
  Menggabungkan seluruh file teks/kode dalam folder ke satu file `.md`, lengkap dengan pohon struktur direktori, header path per file, dan *syntax highlighting*.
- 🔄 **Markdown → File Asli (Unbundle)**
  Memulihkan bundle menjadi struktur folder semula, atau memulihkan file tunggal dengan nama aslinya.
- 👀 **Pratinjau sebelum konfirmasi**
  Sebelum bundle atau restore dijalankan, daftar file yang akan diproses ditampilkan lebih dulu. Ketik `l` untuk melihat daftar lengkap.
- 🚫 **Pola abaikan (`--ignore`)**
  Kecualikan file atau folder tertentu saat bundle, misalnya `tests` atau `*.min.js`.
- 🧪 **Mode dry-run (`-n`)** untuk CLI: lihat hasilnya saja tanpa menulis apa pun.
- 🔢 **Filter ekstensi** dan **estimasi token** (berguna saat bundle dikirim ke AI).
- 📂 **Navigasi folder interaktif** dengan nomor, halaman, dan tombol kembali.
- 🛡️ **Aman**: unbundle menolak path berbahaya (`../`, path absolut), dan file rahasia (`.env`) tidak ikut dibundle.
- 🎨 **Tampilan terminal berwarna**, layar bersih di setiap menu, keluar cepat dengan `exit`.

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

## 📥 Instalasi

**Prasyarat:** Python 3.8+ dan Git. Pustaka `pypdf` dan `python-docx` dipasang otomatis. File teks dan kode tetap bisa dikonversi walau kedua pustaka itu belum ada.

### Termux (Android)

```bash
pkg update && pkg upgrade -y
pkg install python git python-lxml -y
termux-setup-storage        # izinkan akses ke ~/storage/downloads

git clone https://github.com/ShimaYuu205/md-converter.git
cd md-converter
pip install -e .            # memasang perintah MDc
```

> `python-lxml` dipasang lewat `pkg` karena kompilasi `lxml` lewat `pip` sering gagal di Termux.

### Windows (CMD / PowerShell)

```powershell
git clone https://github.com/ShimaYuu205/md-converter.git
cd md-converter
pip install -e .
```

Jika perintah `MDc` tidak dikenali, folder `Scripts` milik Python belum ada di PATH. Gunakan `python md.py` sebagai gantinya.

### Linux / macOS

```bash
git clone https://github.com/ShimaYuu205/md-converter.git
cd md-converter
pip install -e .
```

### Pilihan lain: tanpa memasang perintah

Kalau tidak ingin memakai `pip install -e .`, cukup pasang pustakanya lalu jalankan skripnya langsung:

```bash
pip install -r requirements.txt
python md.py
```

Di Termux, agar tetap bisa mengetik `MDc`, buat alias:

```bash
echo "alias MDc='python ~/md-converter/md.py'" >> ~/.bashrc
source ~/.bashrc
```

### Memeriksa dan mencopot

```bash
MDc --help                    # cek instalasi
pip uninstall md-converter    # mencopot perintah MDc
```

> Nama perintahnya **`MDc`** (huruf M dan D besar, c kecil). Di Termux dan Linux, huruf besar-kecil berpengaruh.

---

## 🚀 Cara Memulai

```bash
MDc                 # menu interaktif
MDc file dok.docx   # atau langsung lewat perintah CLI
```

Tanpa memasang perintah, ganti `MDc` dengan `python md.py`. Semua contoh di bawah berlaku untuk keduanya.

---

## 📕 Referensi Perintah Lengkap

### Ringkasan: dua cara memakai MDc

| Tujuan | Menu interaktif (`MDc`) | Perintah CLI |
| --- | --- | --- |
| Satu file → Markdown | Menu **1** | `MDc file <path>` |
| Folder → satu file Markdown | Menu **2** | `MDc bundle <folder>` |
| Markdown → file asli | Menu **3** | `MDc restore <file.md>` |
| Lihat bantuan | | `MDc --help` |
| Bantuan satu perintah | | `MDc bundle --help` |
| Keluar | `0` atau `exit` | |

Gunakan **menu interaktif** kalau ingin memilih file lewat nomor tanpa mengetik path (nyaman di layar sentuh). Gunakan **CLI** untuk skrip, otomasi, atau kalau path sudah diketahui.

---

### A. Menu interaktif

Jalankan `MDc` tanpa argumen. Ketik `exit` di prompt mana pun untuk keluar.

#### Navigasi folder

Pada menu 1, 2, dan 3, file dan folder dipilih lewat **nomor**:

```
[Direktori Saat Ini]: /storage/emulated/0/Download
==================================================
0. [Kembali ke direktori sebelumnya]
1. [DIR] proyek-web
2. laporan.docx (.docx)
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

Folder dan file tersembunyi (diawali titik) tidak ditampilkan. Menu 1 hanya menampilkan file yang didukung, dan menu 3 hanya file `.md`.

#### Menu 1: File → Markdown

1. Pilih file lewat navigasi.
2. Isi nama file Markdown, atau tekan Enter untuk default (nama sumber + `.md`).
3. Jika nama hasil sudah ada, Anda ditanya apakah ingin menimpanya.
4. Konfirmasi `y`.

Hasil disimpan di folder yang sama dengan file sumber. Jika sumbernya sendiri berakhiran `.md`, hasil disimpan sebagai `namafile_converted.md`.

| Jenis file | Hasil di Markdown |
| --- | --- |
| Kode / teks (`.py`, `.html`, `.txt`, dll.) | Isi disalin apa adanya |
| `.docx` | Heading menjadi `#`, daftar menjadi `-` / `1.`, tabel menjadi tabel Markdown |
| `.pdf` | Teks per halaman, diberi penanda `<!-- Halaman N -->` |
| `.doc` | Ditolak. Simpan ulang sebagai `.docx` dulu |

Baris pertama hasil berisi metadata nama asli, dipakai menu 3 untuk menyarankan nama saat file dipulihkan:

```
<!-- META_NAME: laporan.docx -->
```

#### Menu 2: Folder → Markdown (Bundle)

Urutan pertanyaan:

1. Pilih folder proyek, lalu ketik `s`.
2. **Nama file bundle** (default: `nama_folder.md`).
3. **Filter ekstensi** (kosong = semua, contoh: `.py,.js`).
4. **Abaikan file/folder** (kosong = tidak ada, contoh: `tests,*.min.js`). Aturannya ada di bagian [Pola --ignore](#-pola---ignore).
5. **Pratinjau daftar file**, lalu konfirmasi:

```
6 file akan digabungkan (~2.4 KB, ~610 token)
  + main.py (312 B)
  + src/a.py (1.1 KB)
  + src/b.js (905 B)
  ...
2 file diabaikan oleh pola abaikan
1 file dilewati:
  - logo.txt: file biner
Convert isi folder 'proyek' ke 'proyek.md'? (y/n, l = lihat semua):
```

- `y`: buat bundle.
- `n`: batal, tidak ada file yang ditulis.
- `l`: tampilkan daftar lengkap (pratinjau awal dibatasi 30 file), lalu bertanya lagi.

File bundle disimpan **di dalam folder yang dipilih**. Saat dijalankan ulang, bundle lama otomatis dilewati. Setelah selesai, ukuran dan estimasi token ditampilkan. Di Termux dengan paket `termux-api`, Anda juga ditawari menyalin hasil ke clipboard.

#### Menu 3: Markdown → File Asli

1. Pilih file `.md` yang akan dipulihkan.
2. MDs mengenali jenisnya otomatis: **bundle proyek** atau **file tunggal**.
3. Pilih **lokasi penyimpanan**:
   - `1` direktori saat ini (folder tempat file `.md` berada),
   - `2` direktori lain, dipilih lewat navigasi folder,
   - `3` buat folder baru di direktori saat ini.
4. **Bundle proyek**: pratinjau daftar file ditampilkan sebelum konfirmasi.

```
5 file akan diekstrak ke: /storage/emulated/0/Download/hasil
  baru: 4 | sudah ada: 1 | ditolak: 0
  ~ main.py (sudah ada)
  + src/a.py
  + src/b.js
Ekstrak 5 file ke '.../hasil'? (y/n, l = lihat semua):
```

| Tanda | Arti |
| --- | --- |
| `+` | File baru, akan dibuat |
| `~` | Sudah ada di tujuan |
| `x` | Ditolak karena path tidak aman |

Jika ada file yang sudah ada, Anda ditanya: `y` untuk menimpa semuanya, atau `n` untuk melewati yang sudah ada dan hanya menulis yang baru.

5. **File tunggal**: MDs membaca `META_NAME` dan bertanya apakah ingin memakai nama aslinya. Tanpa metadata, Anda mengetik nama sendiri (misalnya `index.html`). Jika sumber aslinya PDF atau DOCX, isinya disimpan sebagai `.txt`.

---

### B. Perintah CLI

```
MDc file    <path>    [-o OUTPUT] [--force]
MDc bundle  <folder>  [-o OUTPUT] [--ext EKSTENSI] [--ignore POLA] [--max-kb N] [-n] [--force]
MDc restore <file.md> [-d TUJUAN] [-n] [--force]
```

#### `MDc file`: satu file → Markdown

| Opsi | Fungsi |
| --- | --- |
| `<path>` | File yang dikonversi (wajib) |
| `-o`, `--output` | Lokasi dan nama file `.md` hasil (default: di samping sumber) |
| `--force` | Timpa file hasil jika sudah ada |

```bash
MDc file laporan.docx
MDc file kode.py -o hasil/kode.md
MDc file buku.pdf --force
```

#### `MDc bundle`: folder → satu file Markdown

| Opsi | Fungsi |
| --- | --- |
| `<folder>` | Folder proyek (wajib) |
| `-o`, `--output` | File bundle hasil (default: `<folder>/<nama_folder>.md`) |
| `--ext` | Filter ekstensi, dipisah koma, mis. `.py,.js` |
| `--ignore` | Abaikan file/folder sesuai pola, dipisah koma, boleh diulang |
| `--max-kb` | Lewati file teks yang lebih besar dari N KB (default 1000) |
| `-n`, `--dry-run` | Tampilkan daftar file saja, tanpa menulis apa pun |
| `--force` | Timpa file bundle jika sudah ada |

```bash
MDc bundle ./proyek
MDc bundle ./proyek -o proyek.md --ext .py,.js
MDc bundle ./proyek --ignore tests,docs --ignore "*.min.js"
MDc bundle ./proyek --max-kb 500
MDc bundle ./proyek --ignore tests -n      # lihat dulu daftar filenya
```

#### `MDc restore`: Markdown → file asli

| Opsi | Fungsi |
| --- | --- |
| `<file.md>` | File Markdown yang dipulihkan (wajib) |
| `-d`, `--dest` | Folder tujuan (default: folder tempat file `.md` berada) |
| `-n`, `--dry-run` | Tampilkan daftar file saja, tanpa menulis apa pun |
| `--force` | Timpa file yang sudah ada |

```bash
MDc restore proyek.md
MDc restore proyek.md -d ./hasil
MDc restore proyek.md -d ./hasil -n        # lihat dulu apa yang akan ditulis
MDc restore proyek.md -d ./hasil --force   # timpa file yang sudah ada
```

Tanpa `--force`, file yang sudah ada **dilewati** saat restore bundle, dan perintah `file` serta `bundle` berhenti dengan pesan error jika file hasil sudah ada.

#### Kode keluar

| Kode | Arti |
| --- | --- |
| `0` | Berhasil |
| `1` | Gagal (file tidak ditemukan, hasil sudah ada tanpa `--force`, tidak ada file yang cocok, dll.) |
| `130` | Dibatalkan dengan Ctrl+C |

---

## 🚫 Pola `--ignore`

Pola dipakai di CLI (`--ignore`) dan di menu 2 (pertanyaan "Abaikan file/folder"). Pisahkan beberapa pola dengan koma. Di CLI, opsi juga boleh diulang.

Pola memakai wildcard gaya shell (`*`, `?`, `[abc]`) dan cocok bila mengenai salah satu dari:

| Yang dicocokkan | Contoh pola | Efek |
| --- | --- | --- |
| Nama folder (di level mana pun) | `tests` | Semua file di dalam folder `tests` diabaikan |
| Nama file | `*.min.js`, `config.json` | File dengan nama itu diabaikan |
| Path lengkap | `docs/*` | Semua isi `docs/` diabaikan |
| Path folder bertingkat | `src/legacy` | Hanya folder `src/legacy` yang diabaikan |

Selalu dilewati otomatis, tanpa perlu pola:

- folder dan file tersembunyi (diawali titik, mis. `.git`, `.env`),
- folder berat: `node_modules`, `venv`, `env`, `__pycache__`, `dist`, `build`, `target`, `vendor`, `site-packages`,
- file kunci: `package-lock.json`, `yarn.lock`, `pnpm-lock.yaml`, `poetry.lock`, `composer.lock`, `Cargo.lock`,
- file biner, file teks di atas `--max-kb` (default 1 MB), bundle lama, dan file hasil bundle itu sendiri,
- PDF/DOCX di dalam folder (konversi lewat `MDc file`).

---

## 💡 Contoh Kasus Penggunaan

| Tujuan | Cara |
| --- | --- |
| Mengirim seluruh kode proyek ke AI dalam satu pesan | `MDc bundle ./proyek --ext .py --ignore tests`, lalu cek estimasi token |
| Memeriksa dulu isi bundle sebelum dibuat | `MDc bundle ./proyek -n` atau lihat pratinjau di menu 2 |
| Memindahkan proyek lewat satu file (chat, email, catatan) | Bundle di perangkat A, restore di perangkat B |
| Menyalin kode hasil AI berformat bundle ke folder nyata | Simpan jawaban sebagai `.md`, lalu `MDc restore jawaban.md -d ./hasil` |
| Mengubah dokumen Word atau PDF menjadi Markdown | `MDc file dokumen.docx` |

**Tips**
- Gunakan `--ext` dan `--ignore` agar bundle lebih kecil dan relevan.
- Estimasi token hanya perkiraan kasar (~4 karakter per token).
- Periksa isi bundle sebelum dibagikan. File biasa yang berisi data rahasia tetap ikut tergabung.
- Penyalinan ke clipboard di Termux butuh `pkg install termux-api` dan aplikasi **Termux:API**.
- Di layar sempit (Termux portrait), banner tampil ringkas. Putar ke landscape untuk banner penuh.

---

## 🔄 Memperbarui ke Versi Terbaru

```bash
cd ~/md-converter
git pull
```

Jika dipasang dengan `pip install -e .`, tidak ada langkah lain. Jika dipasang tanpa `-e` (`pip install .`), jalankan ulang `pip install .` setelah `git pull`.

---

## ⚠️ Batasan

- **PDF hasil scan** tidak bisa diekstrak (tidak ada lapisan teks, butuh OCR).
- **Konversi PDF/DOCX bersifat satu arah.** Saat dipulihkan, isinya hanya teks dan disimpan sebagai `.txt`, bukan PDF/DOCX asli.
- Format DOCX yang didukung: heading, daftar, dan tabel. Gambar dan *styling* lanjutan diabaikan.
- Estimasi token hanya perkiraan kasar.

---

## ❓ Pemecahan Masalah

| Masalah | Solusi |
| --- | --- |
| `MDc: command not found` | Pastikan sudah menjalankan `pip install -e .` dan mengetik `MDc` (huruf besar-kecil berpengaruh). Atau pakai `python md.py` |
| `pypdf belum terpasang` / `python-docx belum terpasang` | Jalankan `pip install -r requirements.txt` |
| `pip install -e .` gagal membangun `lxml` di Termux | Jalankan `pkg install python-lxml` lebih dulu, lalu ulangi |
| Folder `~/storage/downloads` tidak ada di Termux | Jalankan `termux-setup-storage`, atau jalankan dari folder tujuan |
| `tidak ada file yang cocok untuk digabungkan` | Periksa filter `--ext` dan pola `--ignore`. Coba `-n` untuk melihat hasilnya |
| Banner tampil ringkas | Normal, lebar terminal di bawah 80 kolom |
| Warna tidak muncul di Windows lama | Gunakan Windows Terminal atau PowerShell versi terbaru |

---

## 📝 Riwayat Versi

**v2.1**
- Opsi `--ignore` untuk bundle (CLI dan menu interaktif).
- Pratinjau daftar file saat konfirmasi bundle dan restore, dengan `l` untuk melihat semua.
- Mode `-n` / `--dry-run` untuk `bundle` dan `restore`.
- Dapat dipasang sebagai perintah `MDc` (`pip install -e .`).

**v2.0**
- Ditulis ulang: DOCX → Markdown dengan heading, daftar, dan tabel.
- Bundle dengan pohon direktori, filter ekstensi, dan estimasi token.
- Unbundle aman (menolak path berbahaya) dan pilihan lokasi simpan.
- Mode CLI (`file`, `bundle`, `restore`) dan navigasi folder berhalaman.

---

## 👤 Penulis

Dibuat dengan ❤️ oleh **ShimaYuu**.

## 📄 Lisensi

Proyek ini dilisensikan di bawah **MIT License**. Lihat file `LICENSE`.
