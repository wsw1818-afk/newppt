; DocumentExtractor v3 설치형 (Inno Setup)
; 목적: PyInstaller onedir(폴더배포)를 setup.exe 1개로 패키징.
;       설치 후 설치폴더의 exe를 직접 실행 → %TEMP% 자가추출 없음 → Smart App Control 통과.
;       (onefile은 실행 시 %TEMP% 자가추출로 SAC에 차단되므로 반드시 onedir를 담는다.)

#define AppName "Document Extractor v3"
#define AppVersion "3.0.0"
#define AppPublisher "DocExtractor"
#define AppExeName "DocumentExtractor_v3.exe"
#define SourceDir "H:\Claude_work\newppt\dist\DocumentExtractor_v3_folder"
#define OutputDir "H:\Claude_work\newppt\installer_out"

[Setup]
AppId={{D0C4E172-4E58-4F2D-9A3C-1B2E3F4A5B6C}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\DocumentExtractor_v3
DefaultGroupName=Document Extractor v3
AllowNoIcons=yes
OutputDir={#OutputDir}
OutputBaseFilename=DocumentExtractor_v3_Setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; 관리자 권한 없이 사용자 폴더(LocalAppData)에 설치 가능 → 회사 PC 대응
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName}
MinVersion=10.0.19041

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면에 바로가기 만들기"; GroupDescription: "추가 아이콘:"; Flags: unchecked

[Files]
; onedir 폴더 전체(exe + _internal)를 설치 폴더에 담는다.
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\{#AppName} 제거"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{#AppName} 실행"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 제거 시 설치 폴더 전체(런타임 생성 파일 포함)를 깨끗이 삭제
Type: filesandordirs; Name: "{app}"
