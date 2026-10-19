#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MDs - Ultimate Markdown Converter

Mengonversi dokumen/kode ke Markdown, membundel folder proyek menjadi satu
file .md, dan memulihkannya kembali ke file asli.

Mode interaktif : python md.py
Mode CLI        : python md.py file  <path> [-o OUT]
                  python md.py bundle <folder> [-o OUT] [--ext .py,.js]
                  python md.py restore <file.md> [-d TUJUAN] [--force]
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime

# Pustaka opsional: skrip tetap jalan (untuk file teks/kode) bila belum terpasang.
try:
    import pypdf
except ImportError:
    pypdf = None

try:
    from docx import Document as DocxDocument
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph as DocxParagraph
except ImportError:
    DocxDocument = None

VERSION = "2.0"

# --------------------------------------------------------------------------
# Konfigurasi
# --------------------------------------------------------------------------

# Ekstensi teks/kode -> tag syntax highlighting Markdown.
LANG_MAP = {
    ".txt": "text", ".md": "markdown", ".json": "json",
    ".yaml": "yaml", ".yml": "yaml", ".toml": "toml", ".ini": "ini",
    ".cfg": "ini", ".xml": "xml", ".csv": "csv",
    ".html": "html", ".css": "css", ".js": "javascript", ".ts": "typescript",
    ".jsx": "jsx", ".tsx": "tsx", ".php": "php",
    ".py": "python", ".c": "c", ".h": "c", ".cpp": "cpp", ".hpp": "cpp",
    ".java": "java", ".cs": "csharp", ".go": "go", ".rs": "rust",
    ".rb": "ruby", ".sh": "bash", ".sql": "sql", ".kt": "kotlin",
    ".swift": "swift",
}
TEXT_EXTS = set(LANG_MAP)
DOC_EXTS = {".pdf", ".docx"}
SUPPORTED_EXTS = TEXT_EXTS | DOC_EXTS

# Folder & file yang dilewati saat bundle (noise / besar / berisiko).
IGNORE_DIRS = {
    "node_modules", "venv", "env", "__pycache__", "dist", "build",
    "target", "vendor", "site-packages",
}
IGNORE_FILES = {
    "package-lock.json", "yarn.lock", "pnpm-lock.yaml", "poetry.lock",
    "composer.lock", "Cargo.lock",
}
MAX_FILE_BYTES = 1_000_000  # file teks di atas ini dilewati saat bundle
PAGE_SIZE = 20             # jumlah item per halaman di navigasi folder
BUNDLE_TITLE = "# Bundle Proyek:"

# --------------------------------------------------------------------------
# Tampilan terminal
# --------------------------------------------------------------------------

if os.name == "nt":
    os.system("")  # aktifkan ANSI escape di Windows 10+

RESET = "\033[0m"
RED = "\033[1;31m"
GREEN = "\033[1;32m"
YELLOW = "\033[1;33m"
BLUE = "\033[1;34m"
CYAN = "\033[1;36m"
MAGENTA = "\033[0;35m"

BANNER = r"""
                   /$$      /$$ /$$$$$$$
                  | $$$    /$$$| $$__  $$
                  | $$$$  /$$$$| $$  \ $$  /$$$$$$$
                  | $$ $$/$$ $$| $$  | $$ /$$_____/
                  | $$  $$$| $$| $$  | $$| $$
                  | $$\  $ | $$| $$  | $$| $$
                  | $$ \/  | $$| $$$$$$$/|  $$$$$$$
                  |__/     |__/|_______/  \_______/



  /$$$$$$  /$$       /$$                         /$$     /$$
 /$$__  $$| $$      |__/                        |  $$   /$$/
| $$  \__/| $$$$$$$  /$$ /$$$$$$/$$$$   /$$$$$$  \  $$ /$$/  /$$   /$$ /$$   /$$
|  $$$$$$ | $$__  $$| $$| $$_  $$_  $$ |____  $$  \  $$$$/  | $$  | $$| $$  | $$
 \____  $$| $$  \ $$| $$| $$ \ $$ \ $$  /$$$$$$$   \  $$/   | $$  | $$| $$  | $$
 /$$  \ $$| $$  | $$| $$| $$ | $$ | $$ /$$__  $$    | $$    | $$  | $$| $$  | $$
|  $$$$$$/| $$  | $$| $$| $$ | $$ | $$|  $$$$$$$    | $$    |  $$$$$$/|  $$$$$$/
 \______/ |__/  |__/|__/|__/ |__/ |__/ \_______/    |__/     \______/  \______/
"""


def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")


def print_banner():
    clear_screen()
    width = shutil.get_terminal_size((80, 24)).columns
    if width >= 80:
        art = "\n".join(line.rstrip() for line in BANNER.strip("\n").split("\n"))
        print(CYAN + art + RESET)
    else:  # layar sempit (mis. Termux portrait): banner ringkas
        print(CYAN + "=== MDs | ShimaYuu ===" + RESET)
    print(f"{YELLOW}=== ULTIMATE MARKDOWN CONVERTER v{VERSION} ==={RESET}")
    print(f"{GREEN}Format:{RESET} PDF, DOCX, HTML, CSS, JS, TS, PY, C/C++, JAVA, JSON, TXT, dll.")
    print("-" * min(width, 50))


def pause(msg="\nTekan [Enter] untuk lanjut..."):
    input(msg)


def ask(prompt):
    """input() yang mengenali kata 'exit' di mana pun."""
    value = input(prompt).strip()
    if value.lower() == "exit":
        print("\nKeluar dari skrip. Sampai jumpa!")
        sys.exit(0)
    return value


def confirm(prompt):
    return ask(f"{prompt} (y/n): ").lower() in ("y", "ya")


def format_size(num_bytes):
    return f"{num_bytes / 1024:.1f} KB"


def estimate_tokens(text):
    """Perkiraan kasar (~4 karakter per token)."""
    return max(1, len(text) // 4)


# --------------------------------------------------------------------------
# Utilitas file
# --------------------------------------------------------------------------

def read_text_file(path):
    """Baca file teks dengan fallback encoding; tolak file biner."""
    with open(path, "rb") as f:
        raw = f.read()
    if b"\0" in raw[:8192]:
        raise ValueError("file biner")
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp1252", errors="replace")
    return text.replace("\r\n", "\n").replace("\r", "\n")


def write_text_file(path, text):
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def sanitize_filename(name):
    """Ambil nama file saja (buang path) dan tolak nama berbahaya."""
    name = os.path.basename(name.replace("\\", "/")).strip()
    return "" if name in ("", ".", "..") else name


def safe_join(base_dir, rel_path):
    """Gabungkan path dan pastikan hasilnya tetap di dalam base_dir."""
    rel_path = rel_path.replace("\\", "/")
    base_real = os.path.realpath(base_dir)
    target = os.path.realpath(os.path.join(base_real, rel_path))
    try:
        if os.path.commonpath([base_real, target]) != base_real:
            return None
    except ValueError:  # beda drive di Windows
        return None
    return target if target != base_real else None


def default_start_dir():
    start = os.path.expanduser("~/storage/downloads")  # Termux
    return start if os.path.isdir(start) else os.getcwd()


def copy_to_clipboard(text):
    """Salin ke clipboard (hanya Termux). Return True bila berhasil."""
    exe = shutil.which("termux-clipboard-set")
    if not exe:
        return False
    try:
        subprocess.run([exe], input=text.encode("utf-8"), check=True, timeout=15)
        return True
    except Exception:
        return False


# --------------------------------------------------------------------------
# Ekstraksi dokumen (PDF / DOCX)
# --------------------------------------------------------------------------

def pdf_to_text(path):
    if pypdf is None:
        raise ValueError("pypdf belum terpasang (pip install pypdf)")
    reader = pypdf.PdfReader(path)
    if reader.is_encrypted:
        try:
            ok = reader.decrypt("")
        except Exception:
            ok = 0
        if not ok:
            raise ValueError("PDF terenkripsi / butuh password")
    parts = []
    for number, page in enumerate(reader.pages, 1):
        text = (page.extract_text() or "").strip()
        if text:
            parts.append(f"<!-- Halaman {number} -->\n\n{text}")
    if not parts:
        raise ValueError("tidak ada teks yang bisa diekstrak (mungkin PDF hasil scan, butuh OCR)")
    return "\n\n".join(parts) + "\n"


def _docx_paragraph(par):
    text = par.text.strip()
    if not text:
        return ("p", "")
    style = (par.style.name if par.style is not None else "") or ""
    low = style.lower()
    if low == "title":
        return ("h", "# " + text)
    match = re.match(r"heading (\d)", low)
    if match:
        return ("h", "#" * min(int(match.group(1)), 6) + " " + text)
    if "list bullet" in low:
        return ("li", "- " + text)
    if "list number" in low:
        return ("li", "1. " + text)
    return ("p", text)


def _docx_table(table):
    rows = []
    for row in table.rows:
        rows.append([c.text.strip().replace("\n", " ").replace("|", "\\|") for c in row.cells])
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    lines = ["| " + " | ".join(rows[0]) + " |", "| " + " | ".join(["---"] * width) + " |"]
    lines += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join(lines)


def docx_to_markdown(path):
    if DocxDocument is None:
        raise ValueError("python-docx belum terpasang (pip install python-docx)")
    doc = DocxDocument(path)
    blocks = []  # (jenis, teks)
    for child in doc.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            kind, text = _docx_paragraph(DocxParagraph(child, doc))
        elif tag == "tbl":
            kind, text = "t", _docx_table(DocxTable(child, doc))
        else:
            continue
        if text.strip():
            blocks.append((kind, text))

    out = ""
    prev = None
    for kind, text in blocks:
        if prev is not None:
            out += "\n" if (kind == "li" and prev == "li") else "\n\n"
        out += text
        prev = kind
    return out + "\n"


# --------------------------------------------------------------------------
# Inti: File -> Markdown
# --------------------------------------------------------------------------

def convert_file_to_md(path):
    ext = os.path.splitext(path)[1].lower()
    name = os.path.basename(path)
    if ext == ".pdf":
        body = pdf_to_text(path)
    elif ext == ".docx":
        body = docx_to_markdown(path)
    elif ext == ".doc":
        raise ValueError("format .doc lama tidak didukung; simpan ulang sebagai .docx")
    elif ext in TEXT_EXTS:
        body = read_text_file(path)
    else:
        raise ValueError(f"format '{ext}' tidak didukung")
    return f"<!-- META_NAME: {name} -->\n\n{body}"


META_RE = re.compile(r"<!--\s*META_NAME:\s*(.+?)\s*-->")


def split_meta(content):
    """Pisahkan header META_NAME. Return (nama_asli | None, isi)."""
    match = META_RE.match(content)
    if not match:
        return None, content
    rest = content[match.end():]
    if rest.startswith("\n\n"):
        rest = rest[2:]
    elif rest.startswith("\n"):
        rest = rest[1:]
    return sanitize_filename(match.group(1)) or None, rest


# --------------------------------------------------------------------------
# Inti: Folder -> Markdown (bundle)
# --------------------------------------------------------------------------

def make_fence(text):
    """Pagar code block yang lebih panjang dari backtick mana pun di dalam teks."""
    longest = max((len(m) for m in re.findall(r"`{3,}", text)), default=2)
    return "`" * max(3, longest + 1)


def build_tree(rel_paths, root_name):
    tree = {}
    for rel in rel_paths:
        node = tree
        for part in rel.split("/"):
            node = node.setdefault(part, {})

    lines = [root_name + "/"]

    def walk(node, prefix):
        items = sorted(node.items(), key=lambda kv: (not kv[1], kv[0].lower()))
        for index, (name, child) in enumerate(items):
            last = index == len(items) - 1
            lines.append(prefix + ("└── " if last else "├── ") + name + ("/" if child else ""))
            if child:
                walk(child, prefix + ("    " if last else "│   "))

    walk(tree, "")
    return "\n".join(lines)


def is_prior_bundle(path):
    if not path.lower().endswith(".md"):
        return False
    try:
        with open(path, "rb") as f:
            return f.read(64).startswith(BUNDLE_TITLE.encode("utf-8"))
    except OSError:
        return False


def bundle_folder(target_dir, output_md, exts=None, max_bytes=MAX_FILE_BYTES):
    """Gabungkan file teks/kode dalam folder ke satu file Markdown.

    Return dict: count, skipped [(path, alasan)], chars, tokens.
    """
    out_abs = os.path.abspath(output_md)
    folder_name = os.path.basename(os.path.abspath(target_dir)) or "project_bundle"
    included, skipped = [], []

    for root, dirs, files in os.walk(target_dir):
        dirs[:] = sorted(d for d in dirs if d not in IGNORE_DIRS and not d.startswith("."))
        for name in sorted(files):
            if name.startswith(".") or name in IGNORE_FILES:
                continue
            ext = os.path.splitext(name)[1].lower()
            if ext not in TEXT_EXTS or (exts and ext not in exts):
                continue
            path = os.path.join(root, name)
            if os.path.abspath(path) == out_abs or is_prior_bundle(path):
                continue
            rel = os.path.relpath(path, target_dir).replace(os.sep, "/")
            try:
                size = os.path.getsize(path)
                if size > max_bytes:
                    skipped.append((rel, f"terlalu besar ({format_size(size)})"))
                    continue
                included.append((rel, ext, read_text_file(path)))
            except ValueError as e:
                skipped.append((rel, str(e)))
            except OSError as e:
                skipped.append((rel, f"gagal dibaca: {e}"))

    if not included:
        return {"count": 0, "skipped": skipped, "chars": 0, "tokens": 0}

    tree = build_tree([rel for rel, _, _ in included], folder_name)
    tree_fence = make_fence(tree)
    parts = [
        f"{BUNDLE_TITLE} {folder_name}\n",
        f"> Dibuat oleh MDs v{VERSION} pada {datetime.now():%Y-%m-%d %H:%M} "
        f"| {len(included)} file\n",
        f"## Struktur Direktori\n\n{tree_fence}text\n{tree}\n{tree_fence}\n\n---\n",
    ]
    for rel, ext, text in included:
        fence = make_fence(text)
        lang = LANG_MAP.get(ext, "text")
        parts.append(f"## File: `{rel}`\n\n{fence}{lang}\n{text}\n{fence}\n\n---\n")

    result = "\n".join(parts)
    write_text_file(output_md, result)
    return {
        "count": len(included),
        "skipped": skipped,
        "chars": len(result),
        "tokens": estimate_tokens(result),
        "text": result,
    }


# --------------------------------------------------------------------------
# Inti: Markdown -> File asli (unbundle)
# --------------------------------------------------------------------------

FILE_HDR_RE = re.compile(r"^## File: `(.+)`\s*$")
OPEN_FENCE_RE = re.compile(r"^(`{3,})[^`]*$")


def parse_bundle(content):
    """Return daftar [(rel_path, isi_file)] dari file bundle."""
    lines = content.split("\n")
    entries = []
    i = 0
    while i < len(lines):
        header = FILE_HDR_RE.match(lines[i])
        if header:
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            fence_match = OPEN_FENCE_RE.match(lines[j]) if j < len(lines) else None
            if fence_match:
                fence = fence_match.group(1)
                k = j + 1
                while k < len(lines) and lines[k].rstrip() != fence:
                    k += 1
                if k < len(lines):
                    entries.append((header.group(1), "\n".join(lines[j + 1:k])))
                    i = k + 1
                    continue
        i += 1
    return entries


def existing_targets(entries, dest_dir):
    found = []
    for rel, _ in entries:
        target = safe_join(dest_dir, rel)
        if target and os.path.exists(target):
            found.append(rel)
    return found


def extract_entries(entries, dest_dir, overwrite=False):
    """Tulis semua entri bundle. Return [(rel, status, pesan)]."""
    results = []
    for rel, body in entries:
        target = safe_join(dest_dir, rel)
        if target is None:
            results.append((rel, "ditolak", "path tidak aman"))
        elif os.path.exists(target) and not overwrite:
            results.append((rel, "dilewati", "sudah ada"))
        else:
            try:
                write_text_file(target, body)
                results.append((rel, "ok", ""))
            except OSError as e:
                results.append((rel, "gagal", str(e)))
    return results


# --------------------------------------------------------------------------
# Navigasi folder interaktif
# --------------------------------------------------------------------------

def browse_and_select(current_dir, target_type="file", allowed_exts=None):
    """Pilih file atau folder lewat nomor. Return path, atau None bila batal."""
    page = 0
    while True:
        current_dir = os.path.abspath(current_dir)
        try:
            names = sorted(os.listdir(current_dir), key=str.lower)
        except OSError as e:
            print(f"{RED}Error membaca direktori: {e}{RESET}")
            parent = os.path.dirname(current_dir)
            if parent == current_dir:
                return None
            current_dir, page = parent, 0
            pause()
            continue

        folders = [n for n in names
                   if not n.startswith(".") and os.path.isdir(os.path.join(current_dir, n))]
        files = []
        if target_type == "file":
            files = [n for n in names
                     if not n.startswith(".")
                     and os.path.isfile(os.path.join(current_dir, n))
                     and (allowed_exts is None
                          or os.path.splitext(n)[1].lower() in allowed_exts)]
        entries = [(n, True) for n in folders] + [(n, False) for n in files]
        total_pages = max(1, -(-len(entries) // PAGE_SIZE))
        page = min(page, total_pages - 1)
        select_keys = {"s"} | ({"99"} if len(entries) < 99 else set())

        print_banner()
        print(f"{BLUE}[Direktori Saat Ini]: {current_dir}{RESET}")
        print("=" * 50)
        print(f"0. {YELLOW}[Kembali ke direktori sebelumnya]{RESET}")
        if target_type == "folder":
            print(f"s. {GREEN}[PILIH DIREKTORI INI]{RESET}")

        start = page * PAGE_SIZE
        for number, (name, is_dir) in enumerate(entries[start:start + PAGE_SIZE], start + 1):
            if is_dir:
                print(f"{number}. {CYAN}[DIR] {name}{RESET}")
            else:
                print(f"{number}. {name} {MAGENTA}({os.path.splitext(name)[1].lower()}){RESET}")
        if not entries:
            print("(kosong)")
        if total_pages > 1:
            print(f"-- Halaman {page + 1}/{total_pages}  (n = berikutnya, p = sebelumnya) --")
        print("-" * 50)

        choice = ask("Pilih nomor ('q' batal, 'exit' keluar): ").lower()

        if choice == "q":
            return None
        if choice == "n":
            page = min(page + 1, total_pages - 1)
            continue
        if choice == "p":
            page = max(page - 1, 0)
            continue
        if choice == "0":
            parent = os.path.dirname(current_dir)
            if parent == current_dir:
                print(f"{RED}Sudah berada di root direktori!{RESET}")
                pause()
            else:
                current_dir, page = parent, 0
            continue
        if target_type == "folder" and choice in select_keys:
            return current_dir

        try:
            number = int(choice)
        except ValueError:
            print(f"{RED}Masukkan nomor yang valid!{RESET}")
            pause()
            continue
        if 1 <= number <= len(entries):
            name, is_dir = entries[number - 1]
            if is_dir:
                current_dir, page = os.path.join(current_dir, name), 0
            else:
                return os.path.join(current_dir, name)
        else:
            print(f"{RED}Pilihan tidak valid!{RESET}")
            pause()


def choose_destination(default_dir):
    """Pilih lokasi simpan hasil ekstraksi. Return path, atau None bila batal."""
    print("\nLokasi penyimpanan hasil:")
    print(f"1. Direktori saat ini ({default_dir})")
    print("2. Pilih direktori lain")
    print("3. Buat folder baru di direktori saat ini")
    choice = ask("Pilihan [1]: ") or "1"

    if choice == "2":
        return browse_and_select(default_dir, target_type="folder")
    if choice == "3":
        name = ask("Nama folder baru: ")
        if not name or sanitize_filename(name) != name:
            print(f"{RED}Nama folder tidak valid.{RESET}")
            return None
        path = os.path.join(default_dir, name)
        try:
            os.makedirs(path, exist_ok=True)
        except OSError as e:
            print(f"{RED}Gagal membuat folder: {e}{RESET}")
            return None
        return path
    return default_dir


# --------------------------------------------------------------------------
# Menu interaktif
# --------------------------------------------------------------------------

def menu_convert_file():
    print_banner()
    print("--- 1. Convert File ke Markdown ---")
    selected = browse_and_select(default_start_dir(), "file", SUPPORTED_EXTS | {".doc"})
    if not selected:
        return

    default_out = os.path.splitext(selected)[0] + ".md"
    print_banner()
    print(f"File terpilih: {selected}")
    custom = ask(f"Nama file hasil Markdown [{os.path.basename(default_out)}]: ")
    output = default_out
    if custom:
        name = sanitize_filename(custom)
        if name:
            if not name.lower().endswith(".md"):
                name += ".md"
            output = os.path.join(os.path.dirname(selected), name)

    if os.path.abspath(output) == os.path.abspath(selected):
        stem = os.path.splitext(selected)[0]
        output = f"{stem}_converted.md"
        print(f"{YELLOW}Nama sama dengan file sumber, hasil disimpan sebagai "
              f"'{os.path.basename(output)}'.{RESET}")

    if os.path.exists(output) and not confirm(f"'{os.path.basename(output)}' sudah ada. Timpa?"):
        print("Dibatalkan.")
        pause()
        return
    if not confirm(f"Konversi '{os.path.basename(selected)}' menjadi '{os.path.basename(output)}'?"):
        print("Konversi dibatalkan.")
        pause()
        return

    try:
        text = convert_file_to_md(selected)
        write_text_file(output, text)
        print(f"{GREEN}Berhasil! Tersimpan di: {output}{RESET}")
        print(f"Ukuran {format_size(len(text.encode('utf-8')))} | ~{estimate_tokens(text):,} token")
    except Exception as e:
        print(f"{RED}Terjadi kesalahan: {e}{RESET}")
    pause("\nTekan [Enter] untuk kembali ke menu utama...")


def parse_exts(raw):
    if not raw:
        return None
    result = set()
    for part in re.split(r"[,\s]+", raw.strip().lower()):
        if part:
            result.add(part if part.startswith(".") else "." + part)
    return result or None


def menu_convert_folder():
    print_banner()
    print("--- 2. Convert Isi Folder ke Markdown (Bundle) ---")
    target_dir = browse_and_select(default_start_dir(), target_type="folder")
    if not target_dir:
        return

    folder_name = os.path.basename(target_dir) or "project_bundle"
    default_name = f"{folder_name}.md"
    print_banner()
    print(f"Direktori terpilih: {target_dir}")
    custom = ask(f"Nama file bundle Markdown [{default_name}]: ")
    name = sanitize_filename(custom) if custom else default_name
    name = name or default_name
    if not name.lower().endswith(".md"):
        name += ".md"
    output = os.path.join(target_dir, name)

    exts = parse_exts(ask("Filter ekstensi (kosong = semua, contoh: .py,.js): "))

    if os.path.exists(output) and not confirm(f"'{name}' sudah ada. Timpa?"):
        print("Dibatalkan.")
        pause()
        return
    if not confirm(f"Convert isi folder '{folder_name}' ke '{name}'?"):
        print("Konversi folder dibatalkan.")
        pause()
        return

    try:
        stats = bundle_folder(target_dir, output, exts)
    except OSError as e:
        print(f"{RED}Gagal membuat bundle: {e}{RESET}")
        pause()
        return

    if stats["count"] == 0:
        print(f"{YELLOW}Tidak ada file yang cocok untuk digabungkan.{RESET}")
    else:
        print(f"{GREEN}Berhasil menggabungkan {stats['count']} file ke '{name}'!{RESET}")
        print(f"Ukuran {format_size(stats['chars'])} | ~{stats['tokens']:,} token (perkiraan kasar)")
    if stats["skipped"]:
        print(f"{YELLOW}{len(stats['skipped'])} file dilewati:{RESET}")
        for rel, reason in stats["skipped"][:10]:
            print(f"  - {rel}: {reason}")
        if len(stats["skipped"]) > 10:
            print(f"  ... dan {len(stats['skipped']) - 10} lainnya")

    if stats["count"] and shutil.which("termux-clipboard-set"):
        if confirm("Salin hasil ke clipboard?"):
            print(f"{GREEN}Tersalin!{RESET}" if copy_to_clipboard(stats["text"])
                  else f"{RED}Gagal menyalin (mungkin terlalu besar).{RESET}")
    pause("\nTekan [Enter] untuk kembali ke menu utama...")


def menu_restore():
    print_banner()
    print("--- 3. Convert Markdown ke File Asli / Unbundle ---")
    selected = browse_and_select(default_start_dir(), "file", {".md"})
    if not selected:
        return

    try:
        content = read_text_file(selected)
    except (OSError, ValueError) as e:
        print(f"{RED}Gagal membaca file: {e}{RESET}")
        pause()
        return

    base_dir = os.path.dirname(selected)
    entries = parse_bundle(content)

    if entries:  # --- bundle proyek ---
        print_banner()
        print(f"{GREEN}[INFO] Terdeteksi Bundle Proyek: {len(entries)} file{RESET}")
        dest = choose_destination(base_dir)
        if not dest:
            print("Dibatalkan.")
            pause()
            return
        existing = existing_targets(entries, dest)
        overwrite = False
        if existing:
            print(f"{YELLOW}{len(existing)} file sudah ada di tujuan.{RESET}")
            overwrite = confirm("Timpa file yang sudah ada? (n = lewati saja)")
        if not confirm(f"Ekstrak {len(entries)} file ke '{dest}'?"):
            print("Ekstraksi dibatalkan.")
            pause()
            return
        ok = 0
        for rel, status, msg in extract_entries(entries, dest, overwrite):
            if status == "ok":
                ok += 1
                print(f"{GREEN}[DIPULIHKAN] {rel}{RESET}")
            else:
                print(f"{RED}[{status.upper()}] {rel}: {msg}{RESET}")
        print(f"\n{GREEN}Berhasil mengekstrak {ok} dari {len(entries)} file ke {dest}{RESET}")
        pause("\nTekan [Enter] untuk kembali ke menu utama...")
        return

    # --- file tunggal ---
    detected, body = split_meta(content)
    print_banner()
    print(f"File Markdown terpilih: {selected}")

    output_name = ""
    if detected:
        suggested = detected
        if os.path.splitext(detected)[1].lower() in DOC_EXTS | {".doc"}:
            suggested = os.path.splitext(detected)[0] + ".txt"
            print(f"{YELLOW}[INFO] Sumber asli '{detected}' adalah dokumen biner; isi Markdown "
                  f"hanya teks, jadi disarankan disimpan sebagai '{suggested}'.{RESET}")
        else:
            print(f"{GREEN}[INFO] Ditemukan nama asli: {detected}{RESET}")
        if confirm(f"Gunakan nama '{suggested}'?"):
            output_name = suggested
    else:
        print(f"{YELLOW}[INFO] Tidak ada metadata nama asli di file ini.{RESET}")
    if not output_name:
        output_name = sanitize_filename(
            ask("Masukkan nama file hasil (contoh: index.html / catatan.txt): "))
    if not output_name:
        print("Nama file tidak valid / kosong.")
        pause()
        return

    dest = choose_destination(base_dir)
    if not dest:
        print("Dibatalkan.")
        pause()
        return
    output_path = os.path.join(dest, output_name)
    if os.path.exists(output_path) and not confirm(f"'{output_path}' sudah ada. Timpa?"):
        print("Dibatalkan.")
        pause()
        return
    if not confirm(f"Simpan ke '{output_path}'?"):
        print("Dibatalkan.")
        pause()
        return
    try:
        write_text_file(output_path, body)
        print(f"{GREEN}Berhasil memulihkan file ke: {output_path}{RESET}")
    except OSError as e:
        print(f"{RED}Terjadi kesalahan: {e}{RESET}")
    pause("\nTekan [Enter] untuk kembali ke menu utama...")


def interactive_menu():
    while True:
        print_banner()
        print("Pilih Menu Utama:")
        print("1. Convert file -> Markdown")
        print("2. Convert isi folder -> Markdown (bundle)")
        print("3. Convert Markdown -> File asli / Unbundle")
        print("0. Keluar / ketik 'exit'")
        choice = ask("\nMasukkan pilihan (0-3): ")
        if choice == "0":
            print("\nTerima kasih telah menggunakan MDs!")
            return
        elif choice == "1":
            menu_convert_file()
        elif choice == "2":
            menu_convert_folder()
        elif choice == "3":
            menu_restore()
        else:
            print(f"{RED}Pilihan tidak valid!{RESET}")
            pause("\nTekan [Enter] untuk mencoba lagi...")


# --------------------------------------------------------------------------
# Mode CLI
# --------------------------------------------------------------------------

def build_parser():
    parser = argparse.ArgumentParser(
        prog="md.py",
        description="MDs - Ultimate Markdown Converter (tanpa argumen = mode interaktif)")
    sub = parser.add_subparsers(dest="command")

    p_file = sub.add_parser("file", help="konversi satu file ke Markdown")
    p_file.add_argument("path")
    p_file.add_argument("-o", "--output", help="file .md hasil (default: di samping sumber)")
    p_file.add_argument("--force", action="store_true", help="timpa file hasil bila sudah ada")

    p_bundle = sub.add_parser("bundle", help="gabungkan isi folder ke satu file Markdown")
    p_bundle.add_argument("folder")
    p_bundle.add_argument("-o", "--output", help="file .md hasil (default: <folder>/<nama>.md)")
    p_bundle.add_argument("--ext", help="filter ekstensi, mis. .py,.js")
    p_bundle.add_argument("--max-kb", type=int, default=MAX_FILE_BYTES // 1000,
                          help="lewati file teks lebih besar dari ini (KB)")
    p_bundle.add_argument("--force", action="store_true")

    p_restore = sub.add_parser("restore", help="pulihkan file/bundle Markdown ke file asli")
    p_restore.add_argument("markdown")
    p_restore.add_argument("-d", "--dest", help="folder tujuan (default: folder file .md)")
    p_restore.add_argument("--force", action="store_true", help="timpa file yang sudah ada")
    return parser


def run_cli(args):
    def fail(msg):
        print(f"Error: {msg}", file=sys.stderr)
        return 1

    if args.command == "file":
        if not os.path.isfile(args.path):
            return fail(f"file tidak ditemukan: {args.path}")
        output = args.output or os.path.splitext(args.path)[0] + ".md"
        if os.path.abspath(output) == os.path.abspath(args.path):
            output = os.path.splitext(args.path)[0] + "_converted.md"
        if os.path.exists(output) and not args.force:
            return fail(f"'{output}' sudah ada (gunakan --force untuk menimpa)")
        try:
            text = convert_file_to_md(args.path)
            write_text_file(output, text)
        except Exception as e:
            return fail(e)
        print(f"OK: {output} (~{estimate_tokens(text):,} token)")
        return 0

    if args.command == "bundle":
        if not os.path.isdir(args.folder):
            return fail(f"folder tidak ditemukan: {args.folder}")
        folder = os.path.abspath(args.folder)
        output = args.output or os.path.join(folder, os.path.basename(folder) + ".md")
        if os.path.exists(output) and not args.force:
            return fail(f"'{output}' sudah ada (gunakan --force untuk menimpa)")
        try:
            stats = bundle_folder(folder, output, parse_exts(args.ext), args.max_kb * 1000)
        except OSError as e:
            return fail(e)
        if stats["count"] == 0:
            return fail("tidak ada file yang cocok untuk digabungkan")
        print(f"OK: {stats['count']} file -> {output} (~{stats['tokens']:,} token)")
        for rel, reason in stats["skipped"]:
            print(f"  dilewati: {rel} ({reason})")
        return 0

    if args.command == "restore":
        if not os.path.isfile(args.markdown):
            return fail(f"file tidak ditemukan: {args.markdown}")
        try:
            content = read_text_file(args.markdown)
        except (OSError, ValueError) as e:
            return fail(e)
        dest = args.dest or os.path.dirname(os.path.abspath(args.markdown))
        entries = parse_bundle(content)
        if entries:
            bad = 0
            for rel, status, msg in extract_entries(entries, dest, args.force):
                print(f"  [{status}] {rel} {msg}".rstrip())
                bad += status not in ("ok", "dilewati")
            return 1 if bad else 0
        name, body = split_meta(content)
        if not name:
            return fail("tidak ada metadata nama asli; bukan bundle dan nama tidak diketahui")
        if os.path.splitext(name)[1].lower() in DOC_EXTS | {".doc"}:
            name = os.path.splitext(name)[0] + ".txt"
        target = safe_join(dest, name)
        if target is None:
            return fail("nama file tidak aman")
        if os.path.exists(target) and not args.force:
            return fail(f"'{target}' sudah ada (gunakan --force untuk menimpa)")
        try:
            write_text_file(target, body)
        except OSError as e:
            return fail(e)
        print(f"OK: {target}")
        return 0

    return 0


def main(argv=None):
    args = build_parser().parse_args(argv)
    try:
        if args.command is None:
            interactive_menu()
            return 0
        return run_cli(args)
    except (KeyboardInterrupt, EOFError):
        print("\nDibatalkan. Sampai jumpa!")
        return 130


if __name__ == "__main__":
    sys.exit(main())
