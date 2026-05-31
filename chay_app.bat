@echo off
chcp 65001 > nul
echo ==========================================
echo  Ung dung Tong hop Bao cao INV
echo  Cong ty Dien luc Vung Tau
echo ==========================================
echo.
echo Dang khoi dong ung dung...
echo Vui long cho...
echo.
cd /d "%~dp0"
python -m streamlit run app.py --server.port 8501
pause
