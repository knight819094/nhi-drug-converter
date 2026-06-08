[Setup]
AppName=健保用藥品項轉換工具
AppVersion=1.0
AppPublisher=藥局工具
DefaultDirName={autopf}\健保用藥品項轉換工具
DefaultGroupName=健保用藥品項轉換工具
; 此腳本位於 封包\ 子資料夾，將基準目錄設回專案根目錄
SourceDir=..
OutputDir=封包\installer_output
OutputBaseFilename=健保用藥品項轉換工具_安裝程式
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; 讓安裝程式顯示中文
LanguageDetectionMethod=locale

[Languages]
Name: "tradchinese"; MessagesFile: "compiler:Languages\ChineseTrad.isl"

[Tasks]
Name: "desktopicon"; Description: "建立桌面捷徑"; GroupDescription: "其他選項:"

[Files]
; 打包後的整個資料夾
Source: "dist\健保用藥品項轉換工具\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; 建立預設的輸入/輸出資料夾
Source: "健保藥檔載入\.gitkeep"; DestDir: "{app}\健保藥檔載入"; Flags: ignoreversion skipifsourcedoesntexist
Source: "匯出結果\.gitkeep";    DestDir: "{app}\匯出結果";    Flags: ignoreversion skipifsourcedoesntexist

[Dirs]
Name: "{app}\健保藥檔載入"
Name: "{app}\匯出結果"

[Icons]
Name: "{group}\健保用藥品項轉換工具"; Filename: "{app}\健保用藥品項轉換工具.exe"
Name: "{group}\解除安裝";            Filename: "{uninstallexe}"
Name: "{autodesktop}\健保用藥品項轉換工具"; Filename: "{app}\健保用藥品項轉換工具.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\健保用藥品項轉換工具.exe"; Description: "立即開啟工具"; Flags: nowait postinstall skipifsilent
