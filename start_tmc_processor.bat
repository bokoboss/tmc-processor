@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul

cd /d "%~dp0"

echo.
echo TMC Processor
echo =============
echo.

if exist ".venv" (
    echo พบโฟลเดอร์ .venv แล้ว ข้ามขั้นตอนติดตั้ง
) else (
    echo ไม่พบโฟลเดอร์ .venv
    echo กำลังติดตั้งครั้งแรก อาจใช้เวลาหลายนาที...
    echo.

    where python >nul 2>nul
    if errorlevel 1 (
        where py >nul 2>nul
        if errorlevel 1 (
            echo ไม่พบ Python
            echo กรุณาติดตั้ง Python 3.10 หรือใหม่กว่า แล้วเปิดไฟล์นี้ใหม่อีกครั้ง
            echo ดาวน์โหลดได้ที่ https://www.python.org/downloads/windows/
            echo.
            pause
            exit /b 1
        )
        set "PYTHON_CMD=py -3"
    ) else (
        set "PYTHON_CMD=python"
    )

    !PYTHON_CMD! --version
    if errorlevel 1 (
        echo.
        echo เปิด Python ไม่สำเร็จ
        pause
        exit /b 1
    )

    echo.
    echo กำลังสร้างสภาพแวดล้อม Python ใน .venv ...
    !PYTHON_CMD! -m venv .venv
    if errorlevel 1 (
        echo.
        echo สร้าง .venv ไม่สำเร็จ
        pause
        exit /b 1
    )

    call ".venv\Scripts\activate.bat"
    if errorlevel 1 (
        echo.
        echo เปิดใช้งาน .venv ไม่สำเร็จ
        pause
        exit /b 1
    )

    echo.
    echo กำลังอัปเดต pip ...
    python -m pip install --upgrade pip
    if errorlevel 1 (
        echo.
        echo อัปเดต pip ไม่สำเร็จ
        pause
        exit /b 1
    )

    echo.
    echo กำลังติดตั้งแพ็กเกจที่จำเป็น ...
    python -m pip install -e .
    if errorlevel 1 (
        echo.
        echo ติดตั้งแพ็กเกจของ TMC Processor ไม่สำเร็จ
        pause
        exit /b 1
    )

    echo.
    echo กำลังติดตั้ง pywin32 สำหรับ Excel COM ...
    python -m pip install pywin32
    if errorlevel 1 (
        echo.
        echo ติดตั้ง pywin32 ไม่สำเร็จ
        pause
        exit /b 1
    )

    echo.
    echo ทดสอบ Excel COM แบบไม่บังคับ ...
    python -c "import win32com.client as client; excel = client.Dispatch('Excel.Application'); print('Excel COM พร้อมใช้งาน เวอร์ชัน', excel.Version); excel.Quit()"
    if errorlevel 1 (
        echo Excel COM ยังไม่พร้อมใช้งานหรือไม่มี Microsoft Excel
        echo โปรแกรมยังเปิดได้ และจะใช้โหมดสำรองเมื่อจำเป็น
    )

    echo.
    echo ติดตั้งครั้งแรกเสร็จแล้ว
)

if not exist ".venv\Scripts\activate.bat" (
    echo.
    echo พบ .venv แต่ไม่พบไฟล์เปิดใช้งาน
    echo กรุณาลบโฟลเดอร์ .venv แล้วเปิดไฟล์นี้ใหม่อีกครั้ง
    echo.
    pause
    exit /b 1
)

call ".venv\Scripts\activate.bat"
if errorlevel 1 (
    echo.
    echo เปิดใช้งาน .venv ไม่สำเร็จ
    pause
    exit /b 1
)

echo.
echo กำลังเปิด TMC Processor ...
echo หน้าต่างเว็บเบราว์เซอร์จะเปิดขึ้นอัตโนมัติ
echo หากต้องการปิดโปรแกรม ให้กลับมาที่หน้าต่างนี้แล้วกด Ctrl+C
echo.

python -m streamlit run app.py
if errorlevel 1 (
    echo.
    echo เปิดโปรแกรมไม่สำเร็จ
    pause
    exit /b 1
)
