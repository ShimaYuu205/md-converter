@echo off
rem ==========================================================================
rem  MDc.bat - peluncur MDs untuk Windows (CMD / PowerShell / seret-lepas)
rem
rem  * Klik dua kali           : membuka menu interaktif
rem  * Seret file/folder ke sini: aksi dipilih otomatis (bundle / pulihkan / konversi)
rem  * Dari CMD                : MDc file dok.docx   (bila folder ini ada di PATH)
rem ==========================================================================
setlocal
set "DBL="
echo %cmdcmdline% | find /i "/c" >nul && set "DBL=1"

set "PY="
python --version >nul 2>nul && set "PY=python"
if not defined PY (
    py --version >nul 2>nul && set "PY=py"
)
if not defined PY (
    echo Python tidak ditemukan.
    echo Pasang dari https://www.python.org/downloads/ ^(centang "Add python.exe to PATH"^)
    echo atau jalankan: winget install Python.Python.3.12
    pause
    exit /b 1
)

%PY% -c "import pypdf, docx" >nul 2>nul
if errorlevel 1 (
    echo Memasang pustaka yang dibutuhkan ^(pypdf, python-docx^)...
    %PY% -m pip install --user pypdf python-docx
)

%PY% "%~dp0md.py" %*
set "RC=%errorlevel%"
if not "%RC%"=="0" if defined DBL pause
endlocal & exit /b %RC%
