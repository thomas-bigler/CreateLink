{$A+,B-,C+,D-,H+,I-,J+,M-,O-,P-,Q-,R-,S-,T-,U-,V+,W-,X+,Z1}
{$APPTYPE CONSOLE}
{$MODE DELPHI}
{$warnings on}
{$Define UNICODE}
{$R version.res}

// Free Pascal 3.2.2

Program CreateLink;
Uses Windows;

function CommandLineToArgvW(lpCmdLine: PWideChar;out pNumArgs: Integer): PPWideChar; stdcall; external 'shell32.dll';
const 
AppName='CreateLink v1.0.0.0';
Copyright='(c) 2024-2026 bigler.thomas@gmail.com';
CR=#13#10;

Var
  Buf:array[0..255] of char;
  i,IcnNum,SW,e:integer;
  arg:array[1..8] of array[0..2*max_Path] of WideChar;
  

function ParamStrW(Number: Integer): UnicodeString;
var
  argc: Integer;
  argv: PPWideChar;
begin
  Result := '';
  argv := CommandLineToArgvW(GetCommandLineW, argc);
  if argv = nil then Exit;
  try
    if (Number >= 0) and (Number < argc) then
      Result := argv[Number];
  finally
    LocalFree(HLOCAL(argv));
  end;
end;


function ParamCountW: Integer;
var
  argc: Integer;
  argv: PPWideChar;
begin
  Result := 0;
  argv := CommandLineToArgvW(GetCommandLineW, argc);
  if argv = nil then Exit;
  try
    if argc > 0 then
      Result := argc - 1;
  finally
    LocalFree(HLOCAL(argv));
  end;
end;


Function FileExistsW(const s:PWidechar):Boolean;
begin
  // 1. FASTER than FindFirstFile(.. - version
  // 2. Can check for Directorys too! 
  // 3. Works over Network
  result:=GetFileAttributesW(s)<>$FFFFFFFF; // = 0xFFFFFFFF = FAIL...
end;

Function GetFileSizeW(const s:PWideChar):Dword;
var
  sr:TWin32FindDataW;
  h:Thandle;
begin
  h:=Windows.FindFirstFileW(s,sr);
  if (h<>INVALID_HANDLE_VALUE)
   then result:=sr.nFileSizeLow else result:=0;
  windows.FindClose(h);
end;

Procedure ExpandEnvStrW(S:PWideChar);
var 
  t:array[0..512] of WideChar;
begin
  if ExpandEnvironmentStringsW(s,t,sizeof(t)-1)<>0 then lstrcpyW(s,t);
end;


////////////
const
  CLSID_ShellLink: TGUID = (D1: $00021401; D2: $0000; D3: $0000; D4: ($C0, $00, $00, $00, $00, $00, $00, $46));
type
  _SHITEMID = record
    cb: Word; { Size of the ID (including cb itself) }
    abID: array[0..0] of Byte; { The item ID (variable length) }
  end;
  TSHItemID = _SHITEMID;
  SHITEMID = _SHITEMID;
  PItemIDList = ^TItemIDList;
  _ITEMIDLIST = record
    mkid: TSHItemID;
  end;
  TItemIDList = _ITEMIDLIST;
  
////////////////////////////////////////////////////////////////////////////

IShellLinkW = interface(IUnknown) { sl }
    ['{000214F9-0000-0000-C000-000000000046}']  {[SID_IShellLinkW]}
    function GetPath(pszFile: PWideChar; cchMaxPath: Integer; var pfd: TWin32FindDataW; fFlags: DWORD): HResult; stdcall;
    function GetIDList(var ppidl: PItemIDList): HResult; stdcall;
    function SetIDList(pidl: PItemIDList): HResult; stdcall;
    function GetDescription(pszName: PWideChar; cchMaxName: Integer): HResult; stdcall;
    function SetDescription(pszName: PWideChar): HResult; stdcall;
    function GetWorkingDirectory(pszDir: PWideChar; cchMaxPath: Integer): HResult; stdcall;
    function SetWorkingDirectory(pszDir: PWideChar): HResult; stdcall;
    function GetArguments(pszArgs: PWideChar; cchMaxPath: Integer): HResult; stdcall;
    function SetArguments(pszArgs: PWideChar): HResult; stdcall;
    function GetHotkey(var pwHotkey: Word): HResult; stdcall;
    function SetHotkey(wHotkey: Word): HResult; stdcall;
    function GetShowCmd(out piShowCmd: Integer): HResult; stdcall;
    function SetShowCmd(iShowCmd: Integer): HResult; stdcall;
    function GetIconLocation(pszIconPath: PWideChar; cchIconPath: Integer; out piIcon: Integer): HResult; stdcall;
    function SetIconLocation(pszIconPath: PWideChar; iIcon: Integer): HResult; stdcall;
    function SetRelativePath(pszPathRel: PWideChar; dwReserved: DWORD): HResult; stdcall;
    function Resolve(Wnd: HWND; fFlags: DWORD): HResult; stdcall;
    function SetPath(pszFile: PWideChar): HResult; stdcall;
  end;
/////////////////////////////////// End of cuts from ShlObj  

  IShellLink = IShellLinkW;
  /////////////////////////////////// End of cuts from ShlObj

  /////////////////////////////////// cuts from ActiveX
type
  IPersist = interface(IUnknown)
    ['{0000010C-0000-0000-C000-000000000046}']
    function GetClassID(out classID: TGUID): HResult; stdcall;
  end;
  IPersistFile = interface(IPersist)
    ['{0000010B-0000-0000-C000-000000000046}']
    function IsDirty: HResult; stdcall;
    function Load(pszFileName: PWideChar; dwMode: Longint): HResult;
      stdcall;
    function Save(pszFileName: PWideChar; fRemember: BOOL): HResult;
      stdcall;
    function SaveCompleted(pszFileName: PWideChar): HResult;
      stdcall;
    function GetCurFile(out pszFileName: PWideChar): HResult;
      stdcall;
  end;

function CoInitialize(pvReserved: Pointer): HResult; stdcall; external 'ole32.dll' name 'CoInitialize';
procedure CoUninitialize; stdcall; external 'ole32.dll' name 'CoUninitialize';
function CoCreateInstance(const clsid: TGUID; unkOuter: IUnknown; dwClsContext: Longint; const iid: TGUID; out pv): HResult; stdcall; external 'ole32.dll' name 'CoCreateInstance';

   
function CreateLinkExW(const FileName, RunParams, WorkDir, LinkSelfFileName, Description, IconFile: PWidechar; IconNumber, WndParams: cardinal):boolean;
var
  IObject: IUnknown;
  Wdir:Array[0..512] OF WideChar;
  i:Integer;
begin
  if FileExistsW(LinkSelfFileName) then begin
    SetFileAttributesW(LinkSelfFileName,FILE_ATTRIBUTE_NORMAL);
    DeleteFileW(LinkSelfFileName);
  end;

  
  lStrCpyW(wDir,WorkDir);
  if wDir[0]='"' then lstrcpyW(wDir,@wDir[1]); // prevent double quotes
  if lstrlenW(wDir)<1 then begin
    lStrCpyW(wDir,FileName);
    i:=lStrLenW(wDir);
    WHILE (wDir[i]<>'\') AND (i>0) DO Dec(i);    
    IF i>0 THEN wDir[i]:=#0; // ExtractDir
  end;
  
  Coinitialize(nil);
  if CoCreateInstance(CLSID_ShellLink, nil, 1 or 4, IUnknown, IObject) <> 0 then  begin
    CoUninitialize;result:=false;exit;
  end;
  with (IObject as IShellLink) do
  begin        
    SetPath(FileName);
    SetArguments(RunParams);
    SetWorkingDirectory(wDir);
    SetDescription(Description);
    SetIconLocation(IconFile, IconNumber);
    SetShowCmd(WndParams);  // only allowed:   SW_SHOWNORMAL (1), SW_SHOWMAXIMIZED (3), SW_SHOWMINNOACTIVE (7)    
  end;
  (IObject as IPersistFile).Save(PWChar(LinkSelfFileName), false);
    
  CoUninitialize;  
  result:=true;
  
  if not FileExistsW(LinkSelfFileName) then result:=false;  
  if GetFileSizeW(LinkSelfFileName)<256 then result:=false;
end;

////////
    
Function WideToANSI(const Text:PWidechar):PChar;
var
  ByteLen:Integer;
begin  
  ByteLen:=WideCharToMultiByte(CP_ACP, 0,Text, -1,nil, 0,nil, nil)-1;
  WideCharToMultiByte(CP_ACP, 0, Text, -1, Buf, ByteLen,NIL,NIL);
  result:=@Buf;
end;
  
function FileWriteW(fh: THandle; s: UnicodeString): DWORD;
var
  utf8: UTF8String;
  Written: DWORD;
begin
  utf8 := UTF8String(s);
  WriteFile(fh, utf8[1], Length(utf8), Written, nil);
  Result := Written;
end;

function FileWriteLnW(fh: THandle; s: UnicodeString): DWORD;
var
  utf8: UTF8String;
  Written: DWORD;
begin
  utf8 := UTF8String(s + #13#10);
  WriteFile(fh, utf8[1], Length(utf8), Written, nil);
  Result := Written;
end;

function IsConsoleHandle(h: THandle): Boolean;
var
  mode: DWORD;
begin
  Result := GetConsoleMode(h, mode);
end;

procedure WriteW(const s: UnicodeString);
var
  h: THandle;
  written: DWORD;
  utf8: UTF8String;
begin
  h := GetStdHandle(STD_OUTPUT_HANDLE);

  if IsConsoleHandle(h) then begin
    WriteConsoleW(h, PWideChar(s), Length(s), written, nil);
  end else begin
    utf8 := UTF8String(s);
    WriteFile(h, utf8[1], Length(utf8), written, nil);
  end;
end;


procedure WriteLnW(const s: UnicodeString);
begin
  WriteW(UnicodeString(s+#13#10));
end;


function writeExample(fn:PWideChar):Integer;
var 
  f:THandle;
  r:Integer;
begin  
  {$i-}
  r:=1;
  f := CreateFileW(fn,GENERIC_WRITE,0,NIL,CREATE_ALWAYS,FILE_ATTRIBUTE_NORMAL,0);    
  if (f<>INVALID_HANDLE_VALUE) then begin    
    FileWriteLnW(f,
    '@echo off'+CR+
    'ECHO CreateLink examples'+CR+
    'ECHO.'+CR+CR+
    'MD "%USERPROFILE%\Desktop\CreateLink" 2>NUL'+CR+CR+
    'CreateLink "%USERPROFILE%\Desktop\CreateLink\Notepad.lnk" "%SYSTEMROOT%\Notepad.exe"'+CR+CR+
    
    'CreateLink "%USERPROFILE%\Desktop\CreateLink\On-Screen Keyboard.lnk"^'+CR+
    ' "%SYSTEMROOT%\System32\osk.exe"^'+CR+
    ' ""^'+CR+
    ' "%SYSTEMROOT%\System32"^'+CR+
    ' "On-Screen Keyboard can be used instead of a physical keyboard"^'+CR+
    ' "%SYSTEMROOT%\System32\osk.exe"^'+CR+
    ' ""^'+CR+
    ' 1'+CR+CR+
    
    'CreateLink "%USERPROFILE%\Desktop\CreateLink\Character Map.lnk"^'+CR+
    ' "%SYSTEMROOT%\System32\charmap.exe"^'+CR+
    ' ""^'+CR+
    ' "%SYSTEMROOT%\System32"^'+CR+
    ' "View all characters in any installed font"^'+CR+
    ' "%SYSTEMROOT%\System32\charmap.exe"^'+CR+
    ' ""^'+CR+
    ' 1'+CR+CR+
    
    'CreateLink "%USERPROFILE%\Desktop\CreateLink\Far Manager.lnk"^'+CR+
    ' "c:\Program Files\Far Manager\Far.exe"^'+CR+
    ' ""^'+CR+
    ' "c:\Program Files\Far Manager"^'+CR+
    ' "Classical File Manager"^'+CR+
    ' "c:\Program Files\Far Manager\Far.exe"^'+CR+
    ' 2^'+CR+
    ' 3'+CR+CR+
    'PAUSE');
    r:=0;
  end;
  SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$07);
  CloseHandle(f);{$i+}
  result:=ioresult+r;
end;

begin
  ZeroMemory(@arg,sizeOf(arg));
  
  for i:=1 to ParamCountW do 
    lstrcpyW(arg[i],PWideChar(ParamStrW(i)));    
  
  val(WideToANSI(@arg[7][0]),IcnNum,e);if e<>0 then IcnNum:=0;
  val(WideToANSI(@arg[8][0]),SW,e);if e<>0 then SW:=1;
  
  if paramCountW<2 then begin
    
    if lstrcmpiW(arg[1],'x')=0 then begin      // Example
      SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$07);
      lstrcpyW(arg[1],'%USERPROFILE%\Desktop\Example.bat');ExpandEnvStrW(arg[1]);      
      if writeExample(arg[1])=0 then begin  
        WriteW(' Written: "');WriteW(PWideChar(arg[1]));WriteLnW('"');
        halt(0);  
      end else begin
        WriteW(' Error writing "');WriteW(PWideChar(arg[1]));WriteLnW('"'#13#10);  
        halt(1);  
      end;
    end;    
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$0a);    
    WriteW(PWideChar(#13+AppName+#32));WriteLnW(PWideChar(Copyright));
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$0b);
    WriteLnW('CREATELINK Link Target [Argument(s)] [WorkingDirectory] [Description] [IconPath] [IconNumber] [WndParams]'#13#10);    
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$07);  
    WriteW('Example (batch, using line-break ^):'#9#9#9);
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$0b);  
    WriteLnW('CREATELINK x   => EXAMPLE.BAT on user desktop'#13#10);    
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$07);  
    WriteLnW(
    ' CreateLink "%USERPROFILE%\Desktop\Far Manager.lnk"^'#9'Link'+CR+
    '  "c:\Program Files\Far Manager\Far.exe"^'#9#9'Executable'+CR+
    '  ""^'#9#9#9#9#9#9#9'Argument(s)'+CR+
    '  "c:\Program Files\Far Manager"^'#9#9#9'Working Directory'+CR+
    '  "Classical File Manager"^'#9#9#9#9'Description'+CR+
    '  "c:\Program Files\Far Manager\Far.exe"^'#9#9'Icon Path'+CR+
    '  2^'#9#9#9#9#9#9#9'Icon Number'+CR+
    '  3'#9#9#9#9#9#9#9'Show Command [1 {Normal}|3 {Maximized}|7 {MinNoActive}]'+CR+CR+
    'Existing links will be overwritten without prompting.');
    halt(0);
  end;    
  
 if not CreateLinkExW(PWideChar(arg[2]),arg[3],arg[4],arg[1],arg[5],arg[6],IcnNum,SW) then begin    
    WriteW(' Error creating: "');WriteW(PWideChar(ParamStrW(1)));WriteLnW('"'#13#10);    
    halt(1);
  end else begin
    WriteW(' Link created: "');WriteW(PWideChar(ParamStrW(1)));WriteLnW('"');
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$07);
    WriteLnW('');
    halt(0);  
  end;  
end.
