; FaithSlide 安裝程式
; 由 build.py 呼叫：ISCC /DAppVersion=1.3.0 FaithSlide_Installer.iss
; 路徑皆相對於本檔所在資料夾

#ifndef AppVersion
  #define AppVersion "0.0.0-dev"
#endif

[Setup]
; 舊版未指定 AppId，Inno Setup 預設使用 AppName；沿用相同值，升級時才會辨認為同一個程式
AppId=FaithSlide
AppName=FaithSlide
AppVersion={#AppVersion}
AppVerName=FaithSlide {#AppVersion}
DefaultDirName={autopf}\FaithSlide
DefaultGroupName=FaithSlide
OutputDir=Output
OutputBaseFilename=FaithSlide-Setup-{#AppVersion}
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\FaithSlide.exe
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
; 預設英文確保一定能編譯；繁體中文語系檔需另外放在 Inno Setup 的 Languages 資料夾
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinesetraditional"; MessagesFile: "compiler:Languages\ChineseTraditional.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[InstallDelete]
; 1.2.x 以前的執行檔名稱含版本號（如 FaithSlide 1.2.2-beta.exe），升級時移除
Type: files; Name: "{app}\FaithSlide *.exe"

[Files]
Source: "dist\FaithSlide.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\FaithSlide"; Filename: "{app}\FaithSlide.exe"
Name: "{autodesktop}\FaithSlide"; Filename: "{app}\FaithSlide.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\FaithSlide.exe"; Description: "{cm:LaunchProgram,FaithSlide}"; Flags: nowait postinstall
