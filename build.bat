@echo off
chcp 65001 >nul
echo ========================================
echo  EML Importer - .exe 빌드 스크립트
echo ========================================
echo.

:: Python 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo https://www.python.org 에서 Python 3.8 이상을 설치하세요.
    pause
    exit /b 1
)

:: 의존성 설치
echo [1/3] 필요 패키지 설치 중...
pip install -r requirements.txt
if errorlevel 1 (
    echo [오류] 패키지 설치 실패
    pause
    exit /b 1
)

:: pywin32 post-install (필수)
echo.
echo [2/3] pywin32 초기화 중...
python -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    python Scripts\pywin32_postinstall.py -install >nul 2>&1
)

:: PyInstaller로 exe 빌드
echo.
echo [3/3] .exe 파일 빌드 중...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "EML_Importer" ^
    --icon NONE ^
    --add-data "." ^
    eml_importer.py

if errorlevel 1 (
    echo.
    echo [오류] 빌드 실패. 위 로그를 확인하세요.
    pause
    exit /b 1
)

echo.
echo ========================================
echo  빌드 완료!
echo  실행 파일: dist\EML_Importer.exe
echo ========================================
echo.
echo EML_Importer.exe 를 더블클릭하면 GUI가 실행됩니다.
echo CLI 사용법: EML_Importer.exe C:\eml폴더 --folder "보낸 편지함"
echo.
pause
