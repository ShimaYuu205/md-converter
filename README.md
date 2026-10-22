# MDs - Ultimate Markdown Converter 🚀

**MDs** adalah alat baris perintah (CLI) berbasis Python untuk:

- mengonversi dokumen (PDF, DOCX) dan file kode ke Markdown (`.md`),
- menggabungkan (*bundle*) seluruh isi folder proyek menjadi satu file Markdown,
- memulihkan file Markdown kembali ke file aslinya, termasuk **kembali menjadi DOCX atau PDF**.

Ringan dan responsif, dibuat untuk lingkungan terbatas seperti **Termux (Android)** maupun terminal desktop (**Windows CMD / PowerShell**, Linux, macOS). Dijalankan lewat perintah **`MDc`**, baik sebagai menu interaktif, perintah CLI, maupun **seret dan lepas** file/folder.

---

## 📑 Daftar Isi

- [Fitur](#-fitur)
- [Format yang Didukung](#️-format-yang-didukung)
- [Instalasi](#-instalasi)
- [Cara Memulai](#-cara-memulai)
- [Referensi Perintah Lengkap](#-referensi-perintah-lengkap)
- [PDF dan DOCX Dua Arah](#-pdf-dan-docx-dua-arah)
- [Pola --ignore](#-pola---ignore)
- [Contoh Kasus Penggunaan](#-contoh-kasus-penggunaan)
- [Memperbarui ke Versi Terbaru](#-memperbarui-ke-versi-terbaru)
- [Batasan](#️-batasan)
- [Pemecahan Masalah](#-pemecahan-masalah)
- [Riwayat Versi](#-riwayat-versi)

---

## ✨ Fitur

- 📜 **File → Markdown**
  DOCX diubah menjadi Markdown sungguhan (heading, daftar bertingkat, tabel, **tebal**, *miring*, tautan), PDF diekstrak per halaman, file kode/teks disimpan apa adanya dengan metadata nama asli.
- 📁 **Folder → Markdown (Bundle)**
  Menggabungkan seluruh file teks/kode dalam folder ke satu file `.md`, lengkap dengan pohon struktur direktori, header path per file, dan *syntax highlighting*.
- 🔄 **Markdown → File Asli (Unbundle)**
  Memulihkan bundle menjadi struktur folder semula, atau memulihkan file tunggal. **Dokumen PDF/DOCX bisa dipulihkan kembali sebagai PDF/DOCX** (lihat [PDF dan DOCX Dua Arah](#-pdf-dan-docx-dua-arah)).
- 🖱 **Seret dan lepas / tempel path**
  Seret file atau folder ke `MDc.bat` (Windows) atau tempel path di prompt mana pun. Aksi dipilih otomatis.
- 👀 **Pratinjau sebelum konfirmasi**
  Sebelum bundle atau restore dijalankan, daftar file yang akan diproses ditampilkan lebih dulu. Ketik `l` untuk melihat daftar lengkap.
- 🚫 **Pola abaikan (`--ignore`)**
  Kecualikan file atau folder tertentu saat bundle, misalnya `tests` atau `*.min.js`.
- 🧪 **Mode dry-run (`-n`)** untuk CLI: lihat hasilnya saja tanpa menulis apa pun.
- 🔢 **Filter ekstensi** dan **estimasi token** (berguna saat bundle dikirim ke AI).
- 📂 **Navigasi folder interaktif** dengan nomor, halaman, dan tombol kembali.
- 🩺 **`MDc doctor`** memeriksa Python, pustaka, git, dan PATH, berguna saat `pip`/`MDc` tidak dikenali.
- 🛡️ **Aman**: unbundle menolak path berbahaya (`../`, path absolut), file rahasia (`.env`) tidak ikut dibundle, dan symlink ke luar folder dilewati.
- 🎨 **Tampilan terminal berwarna**, layar bersih di setiap menu, keluar cepat dengan `exit`.

---

## 🛠️ Format yang Didukung

| Kategori | Ekstensi |
| --- | --- |
| Dokumen (dua arah) | `.pdf`, `.docx` |
| Teks & dokumentasi | `.txt`, `.md`, `.markdown`, `.mdx`, `.rst`, `.adoc`, `.tex`, `.bib`, `.org`, `.csv`, `.tsv`, `.diff`, `.patch` |
| Konfigurasi & data | `.json`, `.jsonc`, `.json5`, `.yaml`, `.yml`, `.toml`, `.ini`, `.cfg`, `.conf`, `.properties`, `.xml`, `.plist`, `.csproj`, `.service`, `.tf`, `.tfvars`, `.nix`, `.proto`, `.graphql`, `.gql`, `.cmake` |
| Web | `.html`, `.htm`, `.xhtml`, `.css`, `.scss`, `.sass`, `.less`, `.js`, `.mjs`, `.cjs`, `.jsx`, `.ts`, `.tsx`, `.vue`, `.svelte`, `.php` |
| Bahasa pemrograman | `.py`, `.c`, `.h`, `.cpp`, `.cc`, `.cxx`, `.hpp`, `.hh`, `.hxx`, `.ino`, `.java`, `.cs`, `.go`, `.rs`, `.rb`, `.kt`, `.kts`, `.swift`, `.dart`, `.scala`, `.lua`, `.r`, `.pl`, `.pm`, `.hs`, `.ex`, `.exs`, `.erl`, `.clj`, `.jl`, `.zig`, `.sol`, `.gd`, `.vb`, `.fs`, `.groovy`, `.gradle`, `.sql`, `.asm` |
| Shell & skrip | `.sh`, `.bash`, `.zsh`, `.fish`, `.bat`, `.cmd`, `.ps1`, `.psm1` |
| Tanpa ekstensi | `Dockerfile`, `Makefile`, `LICENSE`, `COPYING`, `NOTICE`, `AUTHORS`, `CHANGELOG`, `README`, `Procfile`, `Gemfile`, `Rakefile`, `Vagrantfile`, `Jenkinsfile`, `Brewfile` |
| File tersembunyi yang aman | `.gitignore`, `.gitattributes`, `.dockerignore`, `.editorconfig`, `.prettierrc`, `.eslintrc`, `.babelrc`, `.nvmrc`, `.python-version`, `.flake8` |

> **Catatan:** format `.doc` (Word lama) tidak didukung. Simpan ulang sebagai `.docx` terlebih dahulu. File rahasia seperti `.env` sengaja **tidak** pernah ikut dibundle.

Untuk memfilter file tanpa ekstensi, tulis namanya: `--ext .py,Dockerfile`.

---

## 📥 Instalasi

Ada tiga langkah: **(1) siapkan Python (dan Git, opsional)**, **(2) ambil kodenya**, **(3) pasang atau langsung jalankan**.

### 1. Siapkan Python dan Git

**Python 3.8+ wajib.** **Git opsional**: hanya dipakai untuk mengunduh dan memperbarui. Tanpa Git, cukup unduh ZIP (lihat langkah 2).

| Sistem | Python | Git |
| --- | --- | --- |
| **Windows** | Unduh dari <https://www.python.org/downloads/>. Saat instalasi **centang "Add python.exe to PATH"**. Atau: `winget install Python.Python.3.12` | <https://git-scm.com/download/win>, atau: `winget install Git.Git` |
| **Termux** | `pkg install python` | `pkg install git` |
| **Debian / Ubuntu** | `sudo apt install python3 python3-pip` | `sudo apt install git` |
| **macOS** | `brew install python` | `brew install git` |

Setelah memasang di Windows, **tutup lalu buka CMD baru**, kemudian periksa:

```bat
python --version
git --version
```

- Jika `python` tidak dikenali tetapi `py --version` berhasil, pakai **`py`** sebagai pengganti `python` di semua perintah di bawah.
- Jika mengetik `python` malah membuka Microsoft Store: pasang Python dari python.org (langkah di atas), atau matikan *App execution aliases* untuk `python.exe` di **Pengaturan → Aplikasi → Pengaturan aplikasi lanjutan → Alias eksekusi aplikasi**.
- Di Linux/macOS, gunakan `python3` bila `python` tidak ada.

### 2. Ambil kodenya

**Dengan Git:**

```bash
git clone https://github.com/ShimaYuu205/md-converter.git
cd md-converter
```

**Tanpa Git (unduh ZIP):** buka <https://github.com/ShimaYuu205/md-converter>, klik **Code → Download ZIP**, lalu ekstrak. Folder hasilnya bernama `md-converter-main`. Lewat terminal:

```bat
:: Windows 10/11 (curl dan tar sudah bawaan)
curl -L -o md-converter.zip https://github.com/ShimaYuu205/md-converter/archive/refs/heads/main.zip
tar -xf md-converter.zip
cd md-converter-main
```

```bash
# Termux / Linux / macOS
curl -L -o md-converter.zip https://github.com/ShimaYuu205/md-converter/archive/refs/heads/main.zip
unzip md-converter.zip && cd md-converter-main     # Termux: pkg install unzip curl
```

### 3. Pasang atau langsung jalankan

#### Termux (Android)

```bash
pkg update && pkg upgrade -y
pkg install python git python-lxml -y
termux-setup-storage        # izinkan akses ke ~/storage/downloads

cd md-converter
python -m pip install -e .  # memasang perintah MDc
```

> `python-lxml` dipasang lewat `pkg` karena kompilasi `lxml` lewat `pip` sering gagal di Termux.

#### Windows (CMD / PowerShell)

Di Windows, `pip` sering tidak dikenali. **Selalu panggil lewat Python**:

```bat
cd md-converter
python -m pip install -e .
MDc --version
```

(atau `py -m pip install -e .` bila hanya `py` yang tersedia.)

Jika `MDc` **tidak dikenali** (folder `Scripts` Python belum ada di PATH), pilih salah satu:

1. **Tanpa mengubah apa pun** (di folder proyek):
   ```bat
   python md.py
   python -m md file dok.docx
   ```
2. **Pakai `MDc.bat`** (sudah ada di folder proyek). Ia mencari `python`/`py`, memasang pustaka yang kurang, lalu menjalankan MDs. Tambahkan folder proyek ke PATH agar `MDc` bisa diketik dari mana saja, atau cukup **klik dua kali / seret file ke `MDc.bat`**.
3. **Tambahkan folder `Scripts` ke PATH.** Jalankan `python md.py doctor`; ia mencetak folder yang perlu ditambahkan. Lalu: tekan `Win + R` → ketik `sysdm.cpl` → tab **Advanced** → **Environment Variables** → pilih `Path` → **Edit** → **New** → tempel foldernya → OK. **Buka CMD baru.**

#### Linux / macOS

```bash
cd md-converter
python3 -m pip install -e .
```

Jika muncul `error: externally-managed-environment`, pakai lingkungan virtual atau `pipx`:

```bash
python3 -m venv .venv && source .venv/bin/activate
python -m pip install -e .
# atau:  pipx install .
```

#### Tanpa memasang perintah `MDc`

Cukup pasang pustakanya lalu jalankan skripnya langsung:

```bash
python -m pip install -r requirements.txt
python md.py
```

Di Termux/Linux, agar tetap bisa mengetik `MDc`, buat alias:

```bash
echo "alias MDc='python ~/md-converter/md.py'" >> ~/.bashrc
source ~/.bashrc
```

> File teks dan kode tetap bisa dikonversi walau `pypdf` dan `python-docx` belum terpasang, dan **membangun PDF tidak butuh pustaka tambahan**.

### Memeriksa dan mencopot

```bash
MDc --version                 # versi
MDc doctor                    # cek Python, pustaka, git, dan PATH
python -m pip uninstall md-converter    # mencopot perintah MDc
```

> Nama perintahnya **`MDc`** (huruf M dan D besar, c kecil). Di Termux dan Linux, huruf besar-kecil berpengaruh.

---

## 🚀 Cara Memulai

```bash
MDc                 # menu interaktif
MDc file dok.docx   # atau langsung lewat perintah CLI
MDc dok.docx        # seret-lepas: aksi dipilih otomatis
```

Tanpa memasang perintah, ganti `MDc` dengan `python md.py`. Semua contoh di bawah berlaku untuk keduanya.

### Seret dan lepas / tempel path

| Anda memberi | Yang terjadi |
| --- | --- |
| **Folder** | Dibundle menjadi satu file Markdown |
| **File `.md`** | Dipulihkan (bundle → folder, file tunggal → file asli, bisa DOCX/PDF) |
| **File lain** (`.docx`, `.pdf`, kode, teks) | Dikonversi ke Markdown |

Caranya:

- **Windows:** seret file/folder dan lepas di ikon **`MDc.bat`**. Atau seret ke jendela CMD untuk menempelkan path-nya, misalnya `MDc "C:\Users\Ani\Dokumen\laporan.docx"`.
- **Menu interaktif:** di menu utama maupun di layar pemilihan file/folder, **tempel atau seret path** sebagai pengganti nomor. Tanda kutip, `file://`, dan spasi yang di-*escape* dibersihkan otomatis.
- **Termux:** Android tidak punya seret-lepas ke terminal, jadi **tekan lama → Paste** path-nya.
- **Linux/macOS:** seret file ke jendela terminal untuk menempelkan path.

Setiap aksi tetap menampilkan pertanyaan konfirmasi (nama hasil, timpa/tidak) sebelum menulis apa pun.

---

## 📕 Referensi Perintah Lengkap

### Ringkasan: tiga cara memakai MDc

| Tujuan | Menu interaktif (`MDc`) | Perintah CLI |
| --- | --- | --- |
| Satu file → Markdown | Menu **1** | `MDc file <path>` |
| Folder → satu file Markdown | Menu **2** | `MDc bundle <folder>` |
| Markdown → file asli (TXT/DOCX/PDF) / bundle → folder | Menu **3** | `MDc restore <file.md>` |
| Otomatis (seret-lepas) | Tempel path di menu utama | `MDc <file/folder>` |
| Periksa instalasi | | `MDc doctor` |
| Lihat bantuan / versi | | `MDc --help` / `MDc --version` |
| Keluar | `0` atau `exit` | |

Gunakan **menu interaktif** kalau ingin memilih file lewat nomor tanpa mengetik path (nyaman di layar sentuh). Gunakan **CLI** untuk skrip, otomasi, atau kalau path sudah diketahui.

---

### A. Menu interaktif

Jalankan `MDc` tanpa argumen. Ketik `exit` di prompt mana pun untuk keluar.

#### Navigasi folder

Pada menu 1, 2, dan 3, file dan folder dipilih lewat **nomor** (atau path yang ditempel):

```
[Direktori Saat Ini]: /storage/emulated/0/Download
==================================================
0. [Kembali ke direktori sebelumnya]
1. [DIR] proyek-web
2. laporan.docx (.docx)
--------------------------------------------------
Pilih nomor (q = batal, exit = keluar) atau tempel/seret path:
```

| Ketik | Fungsi |
| --- | --- |
| nomor | Masuk ke folder, atau memilih file |
| path | Folder: dipilih langsung (menu 2) atau dibuka (menu 1/3). File: langsung dipilih |
| `0` | Naik ke folder induk |
| `s` (atau `99`) | Memilih folder yang sedang dibuka (hanya saat memilih folder) |
| `n` / `p` | Halaman berikutnya / sebelumnya (20 item per halaman) |
| `q` | Membatalkan dan kembali ke menu utama |
| `exit` | Keluar dari program |

Folder dan file tersembunyi (diawali titik) tidak ditampilkan. Menu 1 hanya menampilkan file yang didukung, dan menu 3 hanya file `.md`/`.markdown`.

#### Menu 1: File → Markdown

1. Pilih file lewat navigasi.
2. Isi nama file Markdown, atau tekan Enter untuk default (nama sumber + `.md`).
3. Untuk PDF/DOCX, ditanya apakah **menyematkan salinan file asli** (lihat [PDF dan DOCX Dua Arah](#-pdf-dan-docx-dua-arah)).
4. Jika nama hasil sudah ada, Anda ditanya apakah ingin menimpanya.
5. Konfirmasi `y`.

Hasil disimpan di folder yang sama dengan file sumber. Jika sumbernya sendiri berakhiran `.md`, hasil disimpan sebagai `namafile_converted.md`.

| Jenis file | Hasil di Markdown |
| --- | --- |
| Kode / teks (`.py`, `.html`, `.txt`, `Dockerfile`, dll.) | Isi disalin apa adanya |
| `.docx` | Heading menjadi `#`, daftar menjadi `-` / `1.` (bertingkat dengan indentasi), tabel menjadi tabel Markdown, `**tebal**`, `*miring*`, `[tautan](url)` |
| `.pdf` | Teks per halaman, diberi penanda `<!-- Halaman N -->` |
| `.doc` | Ditolak. Simpan ulang sebagai `.docx` dulu |

Baris pertama hasil berisi metadata nama asli, dipakai menu 3 untuk menyarankan nama saat file dipulihkan:

```
<!-- META_NAME: laporan.docx -->
```

#### Menu 2: Folder → Markdown (Bundle)

Urutan pertanyaan:

1. Pilih folder proyek, lalu ketik `s` (atau tempel path foldernya).
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

5. **File tunggal**: MDs membaca `META_NAME` dan bertanya apakah ingin memakai nama aslinya. Jika sumber aslinya PDF/DOCX, ia memberi tahu apakah file akan **dipulihkan persis** (ada salinan asli tertanam) atau **dibangun ulang dari teks**. Tanpa metadata (Markdown biasa), Anda mengetik nama sendiri: `index.html`, `catatan.txt`, atau **`laporan.docx` / `laporan.pdf`** untuk mengubah Markdown menjadi dokumen.

---

### B. Perintah CLI

```
MDc file    <path>    [-o OUTPUT] [--embed] [--force]
MDc bundle  <folder>  [-o OUTPUT] [--ext EKSTENSI] [--ignore POLA] [--max-kb N] [-n] [--force]
MDc restore <file.md> [-d TUJUAN] [--name NAMA] [--text] [--rebuild] [-n] [--force]
MDc doctor
MDc <file atau folder>            # seret-lepas: aksi otomatis
```

#### `MDc file`: satu file → Markdown

| Opsi | Fungsi |
| --- | --- |
| `<path>` | File yang dikonversi (wajib) |
| `-o`, `--output` | Lokasi dan nama file `.md` hasil (default: di samping sumber) |
| `--embed` | PDF/DOCX: sematkan salinan file asli di dalam `.md` agar bisa dipulihkan **persis** |
| `--force` | Timpa file hasil jika sudah ada |

```bash
MDc file laporan.docx
MDc file laporan.docx --embed
MDc file kode.py -o hasil/kode.md
MDc file buku.pdf --force
```

#### `MDc bundle`: folder → satu file Markdown

| Opsi | Fungsi |
| --- | --- |
| `<folder>` | Folder proyek (wajib) |
| `-o`, `--output` | File bundle hasil (default: `<folder>/<nama_folder>.md`) |
| `--ext` | Filter ekstensi atau nama file, dipisah koma, mis. `.py,.js,Dockerfile` |
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
| `--name` | File tunggal: nama hasil. **Ekstensinya menentukan format**: `.docx`, `.pdf`, atau lainnya (teks) |
| `--text` | PDF/DOCX: simpan sebagai `.txt` (perilaku versi lama) |
| `--rebuild` | Abaikan salinan asli tertanam, bangun ulang dari teks Markdown |
| `-n`, `--dry-run` | Tampilkan daftar file saja, tanpa menulis apa pun |
| `--force` | Timpa file yang sudah ada |

```bash
MDc restore proyek.md
MDc restore proyek.md -d ./hasil
MDc restore proyek.md -d ./hasil -n        # lihat dulu apa yang akan ditulis
MDc restore proyek.md -d ./hasil --force   # timpa file yang sudah ada
MDc restore laporan.md                     # laporan.docx / .pdf sesuai META_NAME
MDc restore catatan.md --name catatan.pdf  # Markdown biasa -> PDF
MDc restore catatan.md --name catatan.docx # Markdown biasa -> Word
```

Tanpa `--force`, file yang sudah ada **dilewati** saat restore bundle, dan perintah `file`, `bundle`, serta restore file tunggal berhenti dengan pesan error jika file hasil sudah ada.

#### `MDc doctor`

Menampilkan versi Python, status `pypdf`/`python-docx`/`git`, apakah `MDc` ada di PATH, dan folder `Scripts` yang perlu ditambahkan ke PATH bila belum.

#### Kode keluar

| Kode | Arti |
| --- | --- |
| `0` | Berhasil |
| `1` | Gagal (file tidak ditemukan, hasil sudah ada tanpa `--force`, tidak ada file yang cocok, dll.) |
| `130` | Dibatalkan dengan Ctrl+C |

---

## 🔁 PDF dan DOCX Dua Arah

Markdown hanya berisi teks, jadi PDF/DOCX asli tidak bisa disusun ulang **persis** dari teksnya saja. MDs menyediakan dua cara:

| Cara | Perintah | Hasil saat dipulihkan |
| --- | --- | --- |
| **Salinan asli tertanam** (`--embed`) | `MDc file laporan.docx --embed` lalu `MDc restore laporan.md` | File **identik byte-per-byte** dengan aslinya (checksum SHA-256 diverifikasi) |
| **Bangun ulang dari teks** (default) | `MDc file laporan.docx` lalu `MDc restore laporan.md` | Dokumen **baru** dengan isi yang sama, struktur Markdown dipertahankan |

**Bangun ulang ke DOCX** mempertahankan: heading, paragraf, **tebal**/*miring*/`kode`, tautan, daftar bullet dan berurut (nomor mulai dari 1 di tiap daftar, bertingkat sampai 3 level), tabel, blok kode, kutipan, garis pemisah, dan pemisah halaman (`<!-- Halaman N -->`).

**Bangun ulang ke PDF** dibuat langsung oleh MDs tanpa pustaka tambahan, memakai font standar PDF (Helvetica dan Courier): heading, paragraf, daftar, tabel, kode, kutipan, tautan yang bisa diklik, dan pemisah halaman. Karakter di luar Latin (mis. huruf Arab/Kanji) diganti `?`, dan emoji/simbol dekoratif dihapus; jumlahnya dilaporkan setelah proses.

**Yang tidak ikut kembali tanpa `--embed`:** gambar, font dan warna asli, margin, header/footer, penggabungan sel tabel, komentar, serta tata letak PDF asli.

**Markdown biasa → dokumen:** file `.md` apa pun (mis. jawaban AI) bisa diubah jadi dokumen:

```bash
MDc restore jawaban.md --name jawaban.docx
MDc restore jawaban.md --name jawaban.pdf
```

**Perhatian untuk `--embed`:**
- Ukuran file `.md` bertambah sekitar 33% dari ukuran file asli, dan batasnya 25 MB per file.
- Salinan asli ditaruh di akhir file sebagai komentar HTML (`<!-- META_ORIGINAL_BASE64 ... -->`). Estimasi token tidak menghitungnya, tetapi jika `.md` dikirim ke AI, blok itu ikut terkirim, jadi **jangan pakai `--embed` untuk file yang akan dikirim ke AI**.
- Jika blok rusak (checksum tidak cocok), MDs memberi peringatan dan membangun ulang dari teks.
- Jangan mengedit isi teks lalu berharap hasilnya ikut berubah saat dipulihkan dengan salinan asli: gunakan `--rebuild` bila Anda sudah mengedit teksnya.

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

- folder dan file tersembunyi (diawali titik, mis. `.git`, `.env`), **kecuali** daftar file tersembunyi aman di bagian [Format yang Didukung](#️-format-yang-didukung),
- folder berat: `node_modules`, `venv`, `env`, `__pycache__`, `dist`, `build`, `target`, `vendor`, `site-packages`,
- file kunci: `package-lock.json`, `yarn.lock`, `pnpm-lock.yaml`, `poetry.lock`, `composer.lock`, `Cargo.lock`,
- file biner, file teks di atas `--max-kb` (default 1 MB), symlink yang menunjuk ke luar folder, bundle lama, dan file hasil bundle itu sendiri,
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
| Mengedit Word lewat AI lalu mengembalikannya ke Word | `MDc file laporan.docx`, edit `.md` dengan AI, lalu `MDc restore laporan.md` |
| Mengubah jawaban AI (Markdown) menjadi PDF/Word | `MDc restore jawaban.md --name jawaban.pdf` |
| Mengarsipkan dokumen agar bisa dikembalikan persis | `MDc file kontrak.pdf --embed` |
| Cara tercepat di Windows | Seret file/folder ke `MDc.bat` |

**Tips**
- Gunakan `--ext` dan `--ignore` agar bundle lebih kecil dan relevan.
- Estimasi token hanya perkiraan kasar (~4 karakter per token).
- Periksa isi bundle sebelum dibagikan. File biasa yang berisi data rahasia tetap ikut tergabung.
- Penyalinan ke clipboard di Termux butuh `pkg install termux-api` dan aplikasi **Termux:API**.
- Di layar sempit (Termux portrait), banner tampil ringkas. Putar ke landscape untuk banner penuh.

---

## 🔄 Memperbarui ke Versi Terbaru

**Dengan Git:**

```bash
cd md-converter
git pull
```

Jika dipasang dengan `python -m pip install -e .`, tidak ada langkah lain. Jika dipasang tanpa `-e` (`python -m pip install .`), jalankan ulang perintah itu setelah `git pull`.

**Tanpa Git:** unduh ZIP terbaru (lihat [Instalasi](#-instalasi)), lalu timpa isi folder lama dengan isi folder hasil ekstrak. Jika sebelumnya dipasang tanpa `-e`, jalankan ulang `python -m pip install .`.

---

## ⚠️ Batasan

- **PDF hasil scan** tidak bisa diekstrak (tidak ada lapisan teks, butuh OCR).
- **PDF/DOCX dua arah hanya persis bila memakai `--embed`.** Tanpa itu, hasil pemulihan adalah dokumen baru dari teks: gambar, font/warna asli, header/footer, dan tata letak tidak kembali. Lihat [PDF dan DOCX Dua Arah](#-pdf-dan-docx-dua-arah).
- PDF buatan MDs memakai font standar (Helvetica/Courier): huruf non-Latin diganti `?` dan emoji dihapus.
- Daftar berurut bertingkat di DOCX hasil bangun ulang tidak memulai nomor dari 1 lagi untuk sub-daftar.
- Akhir baris Windows (CRLF) dinormalisasi menjadi LF saat bundle dan restore, jadi file Windows tidak identik byte-per-byte setelah bundle → restore. Isi teksnya tetap sama.
- Estimasi token hanya perkiraan kasar.

---

## ❓ Pemecahan Masalah

| Masalah | Solusi |
| --- | --- |
| `'pip' is not recognized` (Windows) | Pakai `python -m pip install -e .` (atau `py -m pip ...`) |
| `'python' is not recognized` / membuka Microsoft Store | Pasang Python dari python.org dan **centang "Add python.exe to PATH"**, lalu buka CMD baru. Atau pakai `py`. Atau matikan *App execution aliases* untuk python |
| `'MDc' is not recognized` | Pakai `python md.py`, `python -m md` (di folder proyek), atau `MDc.bat`. Atau tambahkan folder `Scripts` ke PATH (`python md.py doctor` menampilkannya) lalu buka CMD baru |
| `'git' is not recognized` | Pasang Git for Windows lalu buka CMD baru, atau lewati Git dan unduh ZIP (lihat [Instalasi](#-instalasi)) |
| `No module named md` saat `python -m md` | Jalankan dari dalam folder proyek, atau pasang dulu dengan `python -m pip install -e .` |
| `MDc: command not found` (Termux/Linux) | Pastikan sudah menjalankan `python -m pip install -e .` dan mengetik `MDc` (huruf besar-kecil berpengaruh). Atau pakai `python md.py` |
| `pypdf belum terpasang` / `python-docx belum terpasang` | Jalankan `python -m pip install -r requirements.txt` |
| `error: externally-managed-environment` | Pakai `python3 -m venv .venv` lalu aktifkan, atau `pipx install .` |
| `pip install -e .` gagal membangun `lxml` di Termux | Jalankan `pkg install python-lxml` lebih dulu, lalu ulangi |
| Folder `~/storage/downloads` tidak ada di Termux | Jalankan `termux-setup-storage`, atau jalankan dari folder tujuan |
| `tidak ada file yang cocok untuk digabungkan` | Periksa filter `--ext` dan pola `--ignore`. Coba `-n` untuk melihat hasilnya |
| PDF hasil pemulihan berisi tanda `?` | Karakter non-Latin tidak bisa ditulis dengan font standar PDF. Pulihkan sebagai `.docx`, atau pakai `--embed` agar PDF asli kembali persis |
| Hasil pemulihan tidak sama dengan file asli | Normal tanpa `--embed` (dibangun ulang dari teks). Gunakan `--embed` saat konversi |
| Banner tampil ringkas | Normal, lebar terminal di bawah 80 kolom |
| Warna tidak muncul di Windows lama | Gunakan Windows Terminal atau PowerShell versi terbaru |

Jalankan **`MDc doctor`** (atau `python md.py doctor`) untuk memeriksa Python, pustaka, git, dan PATH sekaligus.

---

## 📝 Riwayat Versi

**v2.2**
- **PDF dan DOCX dua arah:** `restore` kini bisa menghasilkan `.docx` dan `.pdf` (dibangun ulang dari teks), atau mengembalikan file **persis** lewat `--embed`. Opsi baru: `file --embed`, `restore --name / --text / --rebuild`. Markdown biasa pun bisa diubah ke DOCX/PDF.
- DOCX → Markdown lebih lengkap: **tebal**, *miring*, tautan, daftar bawaan Word (tombol bullet), daftar bertingkat, sel gabungan tidak lagi terduplikasi, dan karakter khusus Markdown di-*escape*.
- **Seret dan lepas** file/folder (`MDc <path>`, `MDc.bat`) dan tempel path di semua pemilihan file/folder.
- `MDc.bat` untuk Windows, perintah `MDc doctor` dan `MDc --version`, serta panduan instalasi Python/Git/tanpa Git; semua petunjuk pip memakai `python -m pip`.
- **97 ekstensi** (naik dari 31), file tanpa ekstensi (`Dockerfile`, `Makefile`, `LICENSE`, dst.), dan dotfile aman (`.gitignore`, `.editorconfig`, dst.). `--ext` bisa memfilter nama file.
- Keamanan: symlink yang menunjuk ke luar folder dilewati saat bundle.
- `pyproject.toml`: `setuptools>=64` (diperlukan untuk `pip install -e`), lisensi MIT dicantumkan.

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
