; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "NetComBridge"
#define MyAppVersion GetFileVersion(".\NetComBridge\bin\Release\NetComBridgeLib.dll")
#define MyAppPublisher "Florent BREHERET"
#define MyAppURL "http://code.google.com/p/net-com-bridge/"

[Setup]
AppId={{9E85CFED-63CA-4F45-AA81-2B49D69C7642}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
VersionInfoVersion={#MyAppVersion}
VersionInfoTextVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=.\NetComBridge\bin\Release\License.txt
;InfoBeforeFile=.\ClassLibrary1\bin\Release\Info.txt
OutputDir="."
OutputBaseFilename=NetComBridgeSetup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: ".\NetComBridge\bin\Release\NetComBridgeLib.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: ".\NetComBridge\bin\Release\License.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: ".\NetComBridge\bin\Release\Readme.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: ".\NetComBridgeAPI.chm"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Readme"; Filename: "{app}\Readme.txt"; WorkingDir: "{app}";
Name: "{group}\NetComBridge API"; Filename: "{app}\NetComBridgeAPI.chm"; WorkingDir: "{app}";
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

[Registry]
Root: HKLM; Subkey: "SOFTWARE\Microsoft\.NETFramework\Policy\AppPatch\v2.0.50727.00000\excel.exe\{{2CCAA9FE-6884-4AF2-99DD-5217B94115DF}}"; ValueType: string; ValueName: "Target Version"; ValueData: "v2.0.50727"

[Run]
Filename:"{reg:HKLM\SOFTWARE\Microsoft\.NETFramework,InstallRoot}\{reg:HKCR\CLSID\{{61b3e12b-3586-3a58-a497-7ed7c4c794b9%7D\InprocServer32\2.0.0.0,RuntimeVersion}\RegAsm.exe"; Parameters: NetComBridgeLib.dll /unregister /tlb:NetComBridgeLib.tlb; WorkingDir: {app}; StatusMsg: "Unregistering NetComBridge dll ..."; Flags: runhidden;
Filename:"{reg:HKLM\SOFTWARE\Microsoft\.NETFramework,InstallRoot}\{reg:HKCR\CLSID\{{61b3e12b-3586-3a58-a497-7ed7c4c794b9%7D\InprocServer32\2.0.0.0,RuntimeVersion}\RegAsm.exe"; Parameters: NetComBridgeLib.dll /tlb:NetComBridgeLib.tlb  /codebase;WorkingDir: {app}; StatusMsg: "Registering NetComBridge dll ..."; Flags: runhidden;

[UninstallRun]
Filename:"{reg:HKLM\SOFTWARE\Microsoft\.NETFramework,InstallRoot}\{reg:HKCR\CLSID\{{61b3e12b-3586-3a58-a497-7ed7c4c794b9%7D\InprocServer32\2.0.0.0,RuntimeVersion}\RegAsm.exe"; Parameters: NetComBridgeLib.dll /unregister /tlb:NetComBridgeLib.tlb; WorkingDir: {app}; StatusMsg: "Unregistering NetComBridge dll ..."; Flags: runhidden;
