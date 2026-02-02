[Setup]
AppName=FaithSlide
AppVersion=1.2.1-beta
DefaultDirName={autopf}\FaithSlide
DefaultGroupName=FaithSlide
OutputDir=D:\python\FaitSlide\Output
OutputBaseFilename=FaithSlide
SetupIconFile=D:\python\FaitSlide\icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
; 改用預設英文，確保編譯成功
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; 請確認此路徑下的檔案名稱與實際檔案完全一致
Source: "D:\python\FaitSlide\dist\FaithSlide 1.2.1-beta.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\FaithSlide"; Filename: "{app}\FaithSlide 1.2.1-beta.exe"
Name: "{autodesktop}\FaithSlide"; Filename: "{app}\FaithSlide 1.2.1-beta.exe"; Tasks: desktopicon

[Run]
; 移除了可能報錯的 skipfsynccheck，保留最穩定的標籤
Filename: "{app}\FaithSlide 1.2.1-beta.exe"; Description: "{cm:LaunchProgram,FaithSlide}"; Flags: nowait postinstall