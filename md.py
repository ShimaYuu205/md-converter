import os
import sys
import pypdf
from docx import Document as DocxDocument

# Diperluas untuk mendukung berbagai bahasa pemrograman dan format dokumen
SUPPORTED_EXTS = {
    # Dokumen & Teks
    ".pdf", ".docx", ".doc", ".txt", ".md", ".json", ".yaml", ".yml",
    # Web & Script
    ".html", ".css", ".js", ".ts", ".jsx", ".tsx", ".php",
    # Bahasa Pemrograman Lainnya
    ".py", ".cpp", ".c", ".h", ".java", ".cs", ".go", ".rs", ".rb", ".sh", ".sql"
}

def clear_screen():
    os.system('clear' if os.name == 'posix' else 'cls')

def print_banner():
    clear_screen()
    banner = "\033[1;36m" + r"""
                   /$$      /$$/$$$$$$$\vert{}$$$/$$$\vert{}$$__  $$\vert{}$$$$/$$$$\vert{}$$  \ $$/$$$$$$$                             
                  | $$ $$/$$ $$| $$\vert{}$$ /$$_____/ \vert{} $$  $$$\vert{}$$| $$\vert{}$$| $$\vert{}$$\  $\vert{}$$| $$\vert{}$$| $$\vert{}$$ \/  | $$\vert{}$$$$$$$/\vert{}$$$$$$$ \vert{}__/     \vert{}__/\vert{}_______/  \_______/    /$$$$$$/$$       /$$/$$     /$$/$$__  $$\vert{}$$      |__/                        |  $$/$$/                    
| $$\__/\vert{}$$$$$$$  /$$/$$$$$$/$$$$/$$$$$$\ $$ /$$/  /$$   /$$/$$   /$$\vert{}$$$$$$ | $$__$$| $$\vert{}$$_  $$_$$ |____  $$\ $$$$/  | $$\vert{}$$| $$\vert{}$$
 \____  $$\vert{}$$  \ $$\vert{}$$| $$\ $$ \ $$/$$$$$$$\ $$/   | $$\vert{}$$| $$\vert{}$$
 /$$\ $$| $$\vert{}$$| $$\vert{}$$ | $$\vert{}$$ /$$__$$    | $$\vert{}$$  | $$\vert{}$$  | $$\vert{}$$$$$$/\vert{}$$  | $$\vert{}$$| $$\vert{}$$ | $$\vert{}$$$$$$$    | $$\vert{}$$$$$$/\vert{}$$$$$$/
 \______/ |__/  |__/|__/|__/ |__/ |__/ \_______/    |__/     \______/  \______/ 
                                                                                
                                                                                
                                                                                

    """ + "\033[0m"
    print(banner)
    print("\033[1;33m=== ULTIMATE MARKDOWN CONVERTER ===\033[0m")
    print("\033[0;32mSupport Format:\033[0m PDF, DOCX, DOC, HTML, CSS, JS, TS, PY, C++, JAVA, JSON, TXT, dll.")
    print("-" * 45)

def browse_and_select(current_dir, target_type="file"):
    while True:
        current_dir = os.path.abspath(current_dir)
        try:
            items = sorted(os.listdir(current_dir))
        except Exception as e:
            print(f"\033[1;31mError membaca direktori: {e}\033[0m")
            current_dir = os.path.dirname(current_dir)
            continue

        print_banner()
        print(f"\033[1;34m[Direktori Saat Ini]: {current_dir}\033[0m")
        print("="*50)
        
        folders = [d for d in items if os.path.isdir(os.path.join(current_dir, d)) and not d.startswith('.')]
        files = [f for f in items if os.path.isfile(os.path.join(current_dir, f)) and not f.startswith('.')] if target_type == "file" else []

        idx = 1
        print("0. \033[1;33m[Back / Kembali ke Direktori Sebelumnya]\033[0m")
        
        if target_type == "folder":
            print("99. \033[1;32m[PILIH DIREKTORI INI UNTUK DI-CONVERT]\033[0m")

        folder_map = {}
        for folder in folders:
            print(f"{idx}. \033[1;36m[DIR] {folder}\033[0m")
            folder_map[idx] = os.path.join(current_dir, folder)
            idx += 1

        file_map = {}
        if target_type == "file":
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                print(f"{idx}. {file} \033[0;35m({ext})\033[0m")
                file_map[idx] = os.path.join(current_dir, file)
                idx += 1

        print("-" * 50)
        choice = input("Pilih nomor (ketik 'exit' untuk keluar, atau 'q' untuk batal): ").strip()

        if choice.lower() == 'exit':
            print("\nKeluar dari skrip. Sampai jumpa!")
            sys.exit(0)
        if choice.lower() == 'q':
            return None
        
        if choice == '0':
            parent_dir = os.path.dirname(current_dir)
            if parent_dir == current_dir:
                print("\033[1;31mSudah berada di root direktori!\033[0m")
                input("\nTekan [Enter] untuk lanjut...")
            else:
                current_dir = parent_dir
            continue

        if target_type == "folder" and choice == '99':
            return current_dir

        try:
            choice_num = int(choice)
            if choice_num in folder_map:
                current_dir = folder_map[choice_num]
            elif target_type == "file" and choice_num in file_map:
                return file_map[choice_num]
            else:
                print("\033[1;31mPilihan tidak valid!\033[0m")
                input("\nTekan [Enter] untuk lanjut...")
        except ValueError:
            print("\033[1;31mMasukkan angka yang valid atau 'exit'!\033[0m")
            input("\nTekan [Enter] untuk lanjut...")

def convert_single_file():
    print_banner()
    print("--- 1. Convert File ke Markdown ---")
    start_dir = os.path.expanduser("~/storage/downloads")
    if not os.path.exists(start_dir):
        start_dir = os.getcwd()
        
    selected_file = browse_and_select(start_dir, target_type="file")
    if not selected_file:
        return

    ext = os.path.splitext(selected_file)[1].lower()
    if ext not in SUPPORTED_EXTS and ext not in [".pdf", ".docx", ".doc"]:
        print(f"\033[1;31mFormat '{ext}' tidak didukung!\033[0m")
        input("\nTekan [Enter] untuk kembali...")
        return

    default_out = os.path.splitext(selected_file)[0] + ".md"
    print_banner()
    print(f"File terpilih: {selected_file}")
    custom_out = input(f"Nama file hasil Markdown [{os.path.basename(default_out)}]: ").strip()
    
    if custom_out.lower() == 'exit':
        sys.exit(0)

    if custom_out:
        if not custom_out.endswith(".md"):
            custom_out += ".md"
        output_md = os.path.join(os.path.dirname(selected_file), custom_out)
    else:
        output_md = default_out

    confirm = input(f"Konversi '{os.path.basename(selected_file)}' menjadi '{os.path.basename(output_md)}'? (y/n): ").strip().lower()
    if confirm != 'y':
        print("Konversi dibatalkan.")
        input("\nTekan [Enter] untuk kembali...")
        return

    file_original_name = os.path.basename(selected_file)
    text = f"<!-- META_NAME: {file_original_name} -->\n\n"
    
    try:
        if ext == ".pdf":
            reader = pypdf.PdfReader(selected_file)
            for page in reader.pages:
                t = page.extract_text()
                if t: text += t + "\n"
        elif ext in [".doc", ".docx"]:
            doc = DocxDocument(selected_file)
            for para in doc.paragraphs:
                text += para.text + "\n"
        else:
            with open(selected_file, "r", encoding="utf-8", errors="ignore") as f:
                text += f.read()

        with open(output_md, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"\033[1;32mBerhasil! Tersimpan di: {output_md}\033[0m")
    except Exception as e:
        print(f"\033[1;31mTerjadi kesalahan saat konversi: {e}\033[0m")
    input("\nTekan [Enter] untuk kembali ke menu utama...")

def convert_folder():
    print_banner()
    print("--- 2. Convert Isi Folder ke Markdown (Bundle) ---")
    start_dir = os.path.expanduser("~/storage/downloads")
    if not os.path.exists(start_dir):
        start_dir = os.getcwd()
        
    target_dir = browse_and_select(start_dir, target_type="folder")
    if not target_dir:
        return

    folder_name = os.path.basename(target_dir) or "project_bundle"
    default_out_name = f"{folder_name}.md"
    
    print_banner()
    print(f"Direktori terpilih: {target_dir}")
    custom_out = input(f"Nama file bundle Markdown [{default_out_name}]: ").strip()
    if custom_out.lower() == 'exit':
        sys.exit(0)

    if custom_out:
        if not custom_out.endswith(".md"):
            custom_out += ".md"
        output_md = os.path.join(target_dir, custom_out)
    else:
        output_md = os.path.join(target_dir, default_out_name)

    confirm = input(f"Konfirmasi convert seluruh isi folder '{folder_name}' ke '{os.path.basename(output_md)}'? (y/n): ").strip().lower()
    if confirm != 'y':
        print("Konversi folder dibatalkan.")
        input("\nTekan [Enter] untuk kembali...")
        return

    combined_text = f"# Bundle Proyek: {folder_name}\n\n"
    count = 0

    # Pemeta ekstensi ke tag syntax highlighting Markdown
    ext_map = {
        ".js": "javascript",
        ".ts": "typescript",
        ".jsx": "jsx",
        ".tsx": "tsx",
        ".py": "python",
        ".cpp": "cpp",
        ".c": "c",
        ".java": "java",
        ".cs": "csharp",
        ".go": "go",
        ".rs": "rust",
        ".rb": "ruby",
        ".php": "php",
        ".sh": "bash",
        ".sql": "sql",
        ".yaml": "yaml",
        ".yml": "yaml",
        ".html": "html",
        ".css": "css",
        ".json": "json"
    }

    for root, dirs, files in os.walk(target_dir):
        for file in sorted(files):
            if file == os.path.basename(output_md) or file == "md.py":
                continue
            ext = os.path.splitext(file)[1].lower()
            if ext in SUPPORTED_EXTS:
                file_path = os.path.join(root, file)
                rel_path = os.path.relpath(file_path, target_dir)
                
                combined_text += f"## File: `{rel_path}`\n\n"
                lang_tag = ext_map.get(ext, "text")
                combined_text += f"```{lang_tag}\n"
                
                try:
                    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                        combined_text += f.read()
                except Exception as e:
                    combined_text += f"// Gagal membaca file: {e}\n"
                
                combined_text += "\n```\n\n---\n\n"
                count += 1

    with open(output_md, "w", encoding="utf-8") as f:
        f.write(combined_text)
    
    print(f"\033[1;32mBerhasil menggabungkan {count} file ke dalam '{os.path.basename(output_md)}'!\033[0m")
    input("\nTekan [Enter] untuk kembali ke menu utama...")

def convert_md_to_file():
    print_banner()
    print("--- 3. Convert Markdown ke File Asli / Unbundle ---")
    start_dir = os.path.expanduser("~/storage/downloads")
    if not os.path.exists(start_dir):
        start_dir = os.getcwd()
        
    selected_md = browse_and_select(start_dir, target_type="file")
    if not selected_md or not selected_md.endswith('.md'):
        print("\033[1;31mHarap pilih file berformat .md (Markdown)!\033[0m")
        input("\nTekan [Enter] untuk kembali...")
        return

    try:
        with open(selected_md, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
    except Exception as e:
        print(f"\033[1;31mGagal membaca file: {e}\033[0m")
        input("\nTekan [Enter] untuk kembali...")
        return

    base_dir = os.path.dirname(selected_md)

    # Deteksi apakah ini file bundle (memiliki pola '## File: ')
    if "## File: `" in content:
        print_banner()
        print(f"\033[1;32m[INFO] Terdeteksi sebagai file Bundle Proyek!\033[0m")
        confirm = input(f"Ekstrak seluruh file di dalam bundle ini ke folder saat ini? (y/n): ").strip().lower()
        if confirm != 'y':
            print("Ekstraksi dibatalkan.")
            input("\nTekan [Enter] untuk kembali...")
            return

        parts = content.split("## File: `")
        extracted_count = 0

        for part in parts[1:]:
            lines = part.split("\n")
            rel_path = lines[0].split("`")[0].strip()
            
            code_start_idx = -1
            code_end_idx = -1
            
            for i, line in enumerate(lines):
                if line.startswith("```") and code_start_idx == -1:
                    code_start_idx = i
                elif line.startswith("```") and code_start_idx != -1:
                    code_end_idx = i
                    break

            if code_start_idx != -1 and code_end_idx != -1:
                file_content = "\n".join(lines[code_start_idx + 1:code_end_idx])
                target_path = os.path.join(base_dir, rel_path)
                
                try:
                    os.makedirs(os.path.dirname(target_path), exist_ok=True)
                    with open(target_path, "w", encoding="utf-8") as f:
                        f.write(file_content)
                    print(f"\033[0;32m[DIPULIHKAN] {rel_path}\033[0m")
                    extracted_count += 1
                except Exception as e:
                    print(f"\033[1;31m[GAGAL] {rel_path}: {e}\033[0m")

        print(f"\n\033[1;32mBerhasil mengekstrak {extracted_count} file dari bundle!\033[0m")
        input("\nTekan [Enter] untuk kembali ke menu utama...")
        return

    # Logika untuk Single File
    detected_name = None
    if content.startswith("<!-- META_NAME:"):
        end_idx = content.find("-->")
        if end_idx != -1:
            detected_name = content[14:end_idx].strip()
            content = content[end_idx + 3:].lstrip()

    print_banner()
    print(f"File Markdown terpilih: {selected_md}")

    if detected_name:
        print(f"\033[1;32m[INFO] Ditemukan keterangan nama & ekstensi asli: {detected_name}\033[0m")
        use_detected = input("Gunakan nama asli tersebut? (y/n): ").strip().lower()
        if use_detected == 'y':
            output_name = detected_name
        else:
            output_name = input("Masukkan nama file hasil ekstraksi kustom: ").strip()
    else:
        print("\033[1;33m[INFO] Tidak ada keterangan nama/ekstensi otomatis di file ini.\033[0m")
        output_name = input("Masukkan nama file hasil ekstraksi (contoh: index.html / catatan.txt): ").strip()

    if output_name.lower() == 'exit':
        sys.exit(0)
    if not output_name:
        print("Nama file tidak boleh kosong.")
        input("\nTekan [Enter] untuk kembali...")
        return

    output_path = os.path.join(base_dir, output_name)
    confirm = input(f"Konversi kembali ke '{output_path}'? (y/n): ").strip().lower()
    if confirm != 'y':
        print("Konversi dibatalkan.")
        input("\nTekan [Enter] untuk kembali...")
        return

    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"\033[1;32mBerhasil memulihkan file ke: {output_path}\033[0m")
    except Exception as e:
        print(f"\033[1;31mTerjadi kesalahan: {e}\033[0m")
    input("\nTekan [Enter] untuk kembali ke menu utama...")

def main():
    while True:
        print_banner()
        print("Pilih Menu Utama:")
        print("1. Convert file -> Markdown")
        print("2. Convert isi folder -> Markdown")
        print("3. Convert Markdown -> File asli / Unbundle")
        print("0. Keluar / Ketik 'exit'")
        
        choice = input("\nMasukkan pilihan (0-3): ").strip()
        
        if choice.lower() == 'exit' or choice == '0':
            print("\nTerima kasih telah menggunakan MD Converter!")
            break
        elif choice == '1':
            convert_single_file()
        elif choice == '2':
            convert_folder()
        elif choice == '3':
            convert_md_to_file()
        else:
            print("\033[1;31mPilihan tidak valid!\033[0m")
            input("\nTekan [Enter] untuk mencoba lagi...")

if __name__ == "__main__":
    main()
