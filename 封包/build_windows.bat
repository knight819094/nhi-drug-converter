@echo off
chcp 65001 >nul
echo ============================================
echo  健保用藥品項轉換工具 - Windows 打包腳本
echo ============================================
echo.

:: 切換到專案根目錄（此腳本位於 封包\ 子資料夾）
cd /d "%~dp0\.."

:: 安裝依賴
echo [1/3] 安裝 Python 套件...
pip install -r requirements.txt
pip install pyinstaller
if %errorlevel% neq 0 (
    echo 錯誤：pip 安裝失敗，請確認 Python 已安裝並加入 PATH
    pause
    exit /b 1
)

:: PyInstaller 打包
echo.
echo [2/3] 打包執行檔...
pyinstaller --clean 封包\健保用藥轉換工具.spec
if %errorlevel% neq 0 (
    echo 錯誤：PyInstaller 打包失敗
    pause
    exit /b 1
)

:: 建立預設資料夾（Inno Setup 需要）
if not exist "健保藥檔載入" mkdir 健保藥檔載入
if not exist "匯出結果"    mkdir 匯出結果

echo.
echo [3/3] 完成！
echo.
echo 執行檔位置： dist\健保用藥品項轉換工具\健保用藥品項轉換工具.exe
echo.
echo 如需打包成安裝程式：
echo   1. 安裝 Inno Setup (https://jrsoftware.org/isdl.php)
echo   2. 用 Inno Setup 開啟 封包\installer.iss
echo   3. 點選 Build → Compile
echo   4. 安裝檔會輸出到 installer_output\ 資料夾
echo.
pause
