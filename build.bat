@echo off
chcp 65001 >nul
echo ========================================
echo  EML Importer v2 - .exe 빌드
echo ========================================
pip install pyinstaller
pyinstaller --onefile --windowed --name "EML_Importer_v2" eml_importer.py
echo.
echo 완료: dist\EML_Importer_v2.exe
pause
