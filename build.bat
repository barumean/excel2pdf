@echo off
:: Excel → PDF 일괄 변환기  PyInstaller 빌드 스크립트
:: 실행: build.bat
:: 결과물: dist\excel2pdf\excel2pdf.exe

setlocal
set APP_NAME=excel2pdf
set ENTRY=main.py

echo [1/3] 의존 패키지 설치 중...
pip install -r requirements.txt
if errorlevel 1 (
    echo 패키지 설치 실패. 로그를 확인해 주세요.
    exit /b 1
)

echo [2/3] PyInstaller 빌드 중...
pyinstaller ^
    --name "%APP_NAME%" ^
    --onedir ^
    --windowed ^
    --clean ^
    --noconfirm ^
    --hidden-import win32com ^
    --hidden-import win32com.client ^
    --hidden-import win32com.server ^
    --hidden-import pywintypes ^
    --hidden-import pythoncom ^
    --hidden-import tkinterdnd2 ^
    --collect-all tkinterdnd2 ^
    "%ENTRY%"

if errorlevel 1 (
    echo 빌드 실패. 오류 로그를 확인해 주세요.
    exit /b 1
)

echo [3/3] 빌드 완료!
echo 실행 파일 위치: dist\%APP_NAME%\%APP_NAME%.exe
endlocal
