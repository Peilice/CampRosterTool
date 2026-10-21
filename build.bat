@echo off
chcp 65001 >nul
echo 安裝所需套件...
pip install -r requirements.txt
if errorlevel 1 goto :error

echo 打包成單一執行檔...
pyinstaller --onefile --console --name 報名資料轉換 main.py
if errorlevel 1 goto :error

echo.
echo 完成！執行檔位於 dist\報名資料轉換.exe
echo 請將該 exe 複製到與原始 Excel 檔同一資料夾使用。
pause
exit /b 0

:error
echo 發生錯誤，請檢查上方訊息。
pause
exit /b 1
