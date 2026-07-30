
[Setup]
AppName=Applied Genetics
AppVerName=Applied Genetics V1.2
DefaultDirName={pf}\AppliedGenetics
DefaultGroupName=AG
UninstallDisplayIcon={app}\AG.exe
OutputDir=userdocs:AP
SetupIconFile=ext.ico
LicenseFile=licenta.txt

[Messages]
ButtonNext=&Continuare >

[Languages]
Name: ro; MessagesFile: "compiler:ro.isl"

[Files]
Source: "AG.exe"; DestDir: "{app}";
Source: "mscomctl.ocx"; DestDir: "{app}"
Source: "GdiPlus.dll"; DestDir: "{app}"
Source: "msvbvm60.DLL"; DestDir: "{app}"
Source: "msvcrt.dll"; DestDir: "{app}"
Source: "Rsrc32.dll"; DestDir: "{app}"
Source: "VB6STKIT.DLL"; DestDir: "{app}"
Source: "winmm.dll"; DestDir: "{app}"
Source: "comdlg32.ocx"; DestDir: "{app}"
Source: "MSFLXGRD.OCX"; DestDir: "{app}"
Source: "richtx32.ocx"; DestDir: "{app}"
Source: "TABCTL32.OCX"; DestDir: "{app}"
Source: "vbrun60sp5.exe"; DestDir: "{app}"
Source: "advapi32.dll"; DestDir: "{app}"
Source: "comdlg32.dll"; DestDir: "{app}"
Source: "gdi32.dll"; DestDir: "{app}"
Source: "msimg32.dll"; DestDir: "{app}"
Source: "msvbvm50.dll"; DestDir: "{app}"
Source: "msvfw32.dll"; DestDir: "{app}"
Source: "ole32.dll"; DestDir: "{app}"
Source: "olepro32.dll"; DestDir: "{app}"
Source: "Rsrc32.dll"; DestDir: "{app}"
Source: "user32.dll"; DestDir: "{app}"
Source: "shell32.dll"; DestDir: "{app}"
Source: "uxtheme.dll"; DestDir: "{app}"
Source: "VB6STKIT.DLL"; DestDir: "{app}"
Source: "VBA6.DLL"; DestDir: "{app}"
Source: "winmm.dll"; DestDir: "{app}"
Source: "zlib1.dll"; DestDir: "{app}"
Source: "zlib.dll"; DestDir: "{app}"
Source: "sit.dat"; DestDir: "{app}"
Source: "AG_DB.db"; DestDir: "{app}"
Source: "scr.db"; DestDir: "{app}"
Source: "MS.exe"; DestDir: "{app}"
Source: "omg.js"; DestDir: "{app}"
Source: "vbrun60sp5.exe"; DestDir: "{app}"
Source: "jos_verde.png"; DestDir: "{app}"
Source: "cutie_verde.png"; DestDir: "{app}"
Source: "st.wav"; DestDir: "{app}"
Source: "licenta.txt"; DestDir: "{app}"
Source: "ico.jpg"; DestDir: "{app}"
Source: "licenta.txt"; DestDir: "{app}"
Source: "comdlg32.oca"; DestDir: "{app}"
Source: "MSFLXGRD.oca"; DestDir: "{app}"
Source: "richtx32.oca"; DestDir: "{app}"
Source: "TABCTL32.oca"; DestDir: "{app}"
Source: "COMDLG32.DEP"; DestDir: "{app}"
Source: "MSFLXGRD.DEP"; DestDir: "{app}"
Source: "RICHTX32.DEP"; DestDir: "{app}"
Source: "TABCTL32.DEP"; DestDir: "{app}"
Source: "fundal.gif"; DestDir: "{app}"
Source: "ajutor.htm"; DestDir: "{app}"
Source: "video.avi"; DestDir: "{app}"
Source: "probleme_cunoscute.htm"; DestDir: "{app}"
Source: "comdlg32.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "mscomctl.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "MSFLXGRD.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "richtx32.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "TABCTL32.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "GdiPlus.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "msvbvm60.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "msimg32.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "msvbvm50.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "msvcrt.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "msvfw32.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "ole32.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "Rsrc32.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "VB6STKIT.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "VBA6.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "winmm.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "zlib.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "zlib1.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "comdlg32.oca"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "MSFLXGRD.oca"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "richtx32.oca"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "TABCTL32.oca"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "COMDLG32.DEP"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "MSFLXGRD.DEP"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "RICHTX32.DEP"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall
Source: "TABCTL32.DEP"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: onlyifdoesntexist uninsneveruninstall

[Icons]
Name: "{group}\AppliedGenetics"; Filename: "{app}\AG.exe"


[Code]

procedure AfterMyProgInstall(S: String);
begin
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\comdlg32.ocx'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\mscomctl.ocx'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\MSFLXGRD.OCX'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\richtx32.ocx'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\TABCTL32.OCX'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\comdlg32.dll'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\GdiPlus.dll'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\msvbvm60.DLL'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\gdi32.dll'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\msimg32.dll'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\msvbvm50.dll'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\msvcrt.dll'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\msvfw32.dll'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\ole32.dll'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\Rsrc32.dll'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\uxtheme.dll'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\VB6STKIT.DLL'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\VBA6.DLL'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\winmm.dll'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\zlib.dll'), False);
  RegisterServer(Is64BitInstallMode, ExpandConstant('{sys}\zlib1.dll'), False);
end;

[Registry]
Root: HKCU; Subkey: "Software\AppliedGenetics"; Flags: uninsdeletekeyifempty
Root: HKCU; Subkey: "Software\AppliedGenetics\AG"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\AppliedGenetics"; Flags: uninsdeletekeyifempty
Root: HKLM; Subkey: "Software\AppliedGenetics\AG"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\AppliedGenetics\AG\Settings"; ValueType: string; ValueName: "Path"; ValueData: "{app}"

[Run]
Filename: "{app}\AG.exe"; Description: "Lansati aplicatia"; Flags: postinstall nowait skipifsilent unchecked
