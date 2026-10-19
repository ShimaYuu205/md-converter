# MDs - Ultimate Markdown Converter 🚀

**MDs** (Markdown Converter) adalah alat baris perintah (*Command Line Interface / CLI*) interaktif berbasis Python murni yang dirancang untuk mengonversi dokumen dan proyek kode ke dalam format Markdown (`.md`), menggabungkan (*bundle*) direktori proyek web, serta memulihkan file Markdown kembali ke bentuk file aslinya.

Dikembangkan secara khusus agar ringan, responsif, dan fleksibel dijalankan di lingkungan dengan sumber daya terbatas seperti **Termux (Android)** maupun terminal desktop modern (**Windows CMD / PowerShell**).

---

## ✨ Fitur Unggulan

- 📜 **Single File Conversion (`File -> Markdown`)**
  Mengonversi dokumen atau kode individual ke Markdown dengan deteksi *metadata* nama asli otomatis.
- 📁 **Project Folder Bundling (`Folder -> Markdown`)**
  Memindai seluruh struktur folder proyek (HTML, CSS, JS, JSON, TXT) dan menggabungkannya ke dalam satu file Markdown terstruktur lengkap dengan *syntax highlighting* dan *header* path file yang jelas.
- 🔄 **Restoration (`Markdown -> File Asli`)**
  Mengekstrak dan memulihkan kembali file Markdown ke format file aslinya. Memiliki fitur deteksi otomatis nama file asli dari *metadata tag*.
- 📂 **Navigasi Folder Interaktif**
  Penjelajahan direktori internal berbasis angka tanpa perlu mengetik *path* manual. Dilengkapi fungsi kembali (`0. Back`) dan pemilihan folder kerja cepat.
- 🎯 **Opsi Penyimpanan Fleksibel**
  Pilih lokasi penyimpanan hasil ekstraksi: di direktori saat ini, direktori lain, atau buat folder baru secara otomatis.
- 🎨 **Text-Based GUI & Clear Screen**
  Antarmuka terminal berbasis warna (ANSI Escape Colors) dengan tampilan layar bersih (*clean screen*) otomatis di setiap menu serta fitur keluar cepat (`exit`).

---

## 🛠️ Format File yang Didukung

- **Dokumen Teks / Kode:** `.html`, `.css`, `.js`, `.json`, `.txt`, `.md`
- **Dokumen Kantor / PDF:** `.pdf`, `.docx`, `.doc`

---

## 💻 Prasyarat Sistem

1. **Python 3.x**
2. Pustaka Python:
   - `pypdf` (ekstraksi teks PDF)
   - `python-docx` (ekstraksi dokumen Word)
   - `lxml` (pemrosesan dokumen XML/Word)

---

## 📥 Panduan Instalasi & Setup

### 1. Di Termux (Android)

Jalankan perintah berikut di terminal Termux Anda:


### Update paket & instal Python + Git
```
pkg update && pkg upgrade -y
pkg install python git python-lxml -y
```
### Clone repository ini (atau unduh kodenya)
```
git clone https://github.com/ShimaYuu205/md-converter.git
cd md-converter
```
### Instal dependensi Python
```
pip install pypdf python-docx
```

### 2. Di Windows (CMD / PowerShell)
​Buka Command Prompt atau PowerShell, lalu jalankan:

### Clone repository

```
git clone https://github.com/ShimaYuu205/md-converter.git
cd md-converter
```
### Instal dependensi Python
```
pip install pypdf python-docx lxml

```

# 📕 Cara Penggunaan
​Jalankan skrip utama menggunakan perintah berikut:
```
python md.py
```

# Panduan Alur Kerja (Workflow)
- **Convert File -> Markdown**

  *Pilih menu 1.*

  Gunakan navigasi angka untuk memilih file
  yang ingin dikonversi. ​Masukkan nama file
  Markdown kustom (jika tidak mengetik
  ekstensi .md, sistem akan otomatis
  menambahkan .md).
  Konfirmasi konversi (y/n).
- **Convert Isi Folder -> Markdown (Project
  Bundle)**

  *Pilih menu 2.*
  
  ​Gunakan navigasi folder untuk masuk ke
  direktori proyek web Anda.Ketik 99 untuk
  memilih direktori saat ini.
  ​Masukkan nama file bundle yang diinginkan
  (default: nama_folder.md).
  ​Seluruh file HTML, CSS, JS, dll. di dalam
  folder akan dirangkum rapi ke dalam satu
  file Markdown.
- **Convert Markdown -> File Asli**
  
   *​Pilih menu 3.*
  
  ​Pilih file .md yang ingin dipulihkan.
  ​Jika file .md memiliki metadata tag nama
  asli, sistem akan menanyakan apakah ingin
  menggunakan nama asli tersebut secara
  otomatis.

**Pilih lokasi direktori tujuan penyimpanan:**
- Simpan di direktori saat ini.
- Pilih direktori lain secara interaktif.
- Buat folder baru di direktori saat ini.

---
👤 Penulis / Author
​Dibuat dengan ❤️ oleh **ShimaYuu**.
---
​📄 Lisensi
​Proyek ini terlisensi di bawah MIT License.
