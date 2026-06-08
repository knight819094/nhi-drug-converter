# 封包說明

本資料夾收錄將程式打包成 Windows 安裝檔所需的全部檔案。

| 檔案 | 用途 |
|------|------|
| `健保用藥轉換工具.spec` | PyInstaller 設定（把 main.py 打包成 .exe） |
| `installer.iss` | Inno Setup 腳本（把 .exe 包成安裝程式） |
| `build_windows.bat` | Windows 一鍵打包腳本 |
| `installer_output/` | 安裝檔輸出位置（執行後自動產生） |

> GitHub Actions 設定檔位於專案根目錄 `.github/workflows/build-windows.yml`，
> 必須留在根目錄，GitHub 才偵測得到。

---

## 方式一：GitHub Actions（推薦，全自動，免 Windows 電腦）

1. 將整個專案推上 GitHub repository
2. 網頁 → **Actions** → **Build Windows Installer** → **Run workflow**
3. 約 5 分鐘後在 **Artifacts** 下載 `健保用藥品項轉換工具_安裝程式.exe`

---

## 方式二：在 Windows 電腦手動打包

1. 安裝 [Python 3.11](https://python.org)（勾選 Add to PATH）
2. 把整個專案複製到 Windows
3. 雙擊 `封包\build_windows.bat`
4. 安裝 [Inno Setup](https://jrsoftware.org/isdl.php)，
   開啟 `封包\installer.iss` → Build → Compile
5. 安裝檔輸出於 `封包\installer_output\`

---

## 注意事項

- 所有路徑均以**專案根目錄**為基準執行，請勿單獨把腳本搬離資料夾結構。
- 若要加入應用程式圖示，放一個 `icon.ico` 在根目錄，
  並取消 `.spec` 中 `# icon='icon.ico'` 的註解。
