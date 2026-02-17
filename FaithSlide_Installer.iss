[Setup]
AppName=FaithSlide
AppVersion=1.2.2-beta
DefaultDirName={autopf}\FaithSlide
DefaultGroupName=FaithSlide
OutputDir=D:\python\FaitSlide\Output
OutputBaseFilename=FaithSlide
SetupIconFile=D:\python\FaitSlide\icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
; ��ιw�]�^��A�T�O�sĶ���\
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinesetraditional"; MessagesFile: "compiler:Languages\ChineseTraditional.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; �нT�{�����|�U���ɮצW�ٻP����ɮק����@�P
Source: "D:\python\FaitSlide\dist\FaithSlide 1.2.2-beta.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\FaithSlide"; Filename: "{app}\FaithSlide 1.2.2-beta.exe"
Name: "{autodesktop}\FaithSlide"; Filename: "{app}\FaithSlide 1.2.2-beta.exe"; Tasks: desktopicon

[Run]
; �����F�i������� skipfsynccheck�A�O�d��í�w������
Filename: "{app}\FaithSlide 1.2.2-beta.exe"; Description: "{cm:LaunchProgram,FaithSlide}"; Flags: nowait postinstall