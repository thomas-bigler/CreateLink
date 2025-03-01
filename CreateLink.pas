{$A+,B-,C+,D-,H+,I-,J+,M-,O-,P-,Q-,R-,S-,T-,U-,V+,W-,X+,Z1}
{$APPTYPE CONSOLE}
{$MODE DELPHI}
{$warnings on}
{$Define UNICODE}
{$R version.res}

// Free Pascal 3.2.2

Program CreateLink;
Uses Windows;

const 
AppName='CreateLink v1.0.0.0';
Copyright='(c) 2024-2025 bigler.thomas@gmail.com';

Function FileExistsW(const s:PWidechar):Boolean;
begin
  // 1. FASTER than FindFirstFile(.. - version
  // 2. Can check for Directorys too! 
  // 3. Works over Network
  result:=GetFileAttributesW(s)<>$FFFFFFFF; // = 0xFFFFFFFF = FAIL...
end;

Function GetFileSizeX(const s:PWideChar):Dword;
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
var t:array[0..512] of WideChar;
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
  if GetFileSizeX(LinkSelfFileName)<256 then result:=false;
end;

////////


Function ParamStrW(Index: LongInt):WideString;
Var 
  P, P2: PWideChar;
  i: Integer;
  B: Boolean;
Begin
  If Index = 0 Then Begin
    SetLength(Result, Max_Path);
    SetLength(Result, GetModuleFileNameW(0, Pointer(Result), Max_Path));
    Exit;
  End Else If Index < 0 Then Inc(Index);
  P := GetCommandLineW;
  If P <> nil Then
    While P^ <> #0 do Begin
      While (P^ <> #0) and (P^ <= ' ') do Inc(P);
      P2 := P; i := 0; B := False;
      While (P2^ <> #0) and ((P2^ > ' ') or B) do Begin
        If P2^ <> '"' Then Inc(i) Else B := not B;
        Inc(P2);
      End;
      If (Index = 0) and (P <> P2) Then Begin
        SetLength(Result, i);
        i := 1;
        Repeat
          If P^ <> '"' Then Begin Result[i] := P^; Inc(i); End;
          Inc(P);
        Until P >= P2;
        Exit;
      End;
      P := P2;
      Dec(Index);
    End;
  Result := '';
End;

    
Function w2p(const s:PWidechar):PChar;
var 
  p:array[0..16384] of Char;  
begin  
  WideCharToMultiByte(CP_ACP, 0, s, -1, p, sizeof(p),NIL,NIL);
  w2p:=@p;
end;

      
function WL(fh:THandle;s:Pchar):DWORD;
// like WriteLN(...
var 
  toWrite,Written:DWORD;
begin  
  toWrite:=lstrlen(s);
  WriteFile(fh,s^,toWrite,Written,NIL);    
  WriteFile(fh,#13#10,2,Written,NIL);
  result:=Written;
end;


function writeExample(fn:PWideChar):Integer;
var f:THandle;r:Integer;
begin  
  {$i-}
  r:=1;
  f := CreateFileW(fn,GENERIC_WRITE,0,NIL,CREATE_ALWAYS,FILE_ATTRIBUTE_NORMAL,0);    
  if (f<>INVALID_HANDLE_VALUE) then begin    
    WL(f,'@echo off');
    WL(f,'ECHO CreateLink examples');
    WL(f,'ECHO.'#13#10);
    WL(f,'MD "%USERPROFILE%\Desktop\CreateLink" 2>NUL'#13#10);
    
    WL(f,'CreateLink "%USERPROFILE%\Desktop\CreateLink\Notepad.lnk" "%SYSTEMROOT%\Notepad.exe"'#13#10);        
    
    WL(f,'CreateLink "%USERPROFILE%\Desktop\CreateLink\On-Screen Keyboard.lnk"^');
    WL(f,' "%SYSTEMROOT%\System32\osk.exe"^');
    //~ WL(f,' "/dockbottom"^');
    WL(f,' ""^');
    WL(f,' "%SYSTEMROOT%\System32"^');
    WL(f,' "On-Screen Keyboard can be used instead of a physical keyboard"^');
    WL(f,' "%SYSTEMROOT%\System32\osk.exe"^');
    WL(f,' ""^');
    WL(f,' 1'#13#10);
    
    WL(f,'CreateLink "%USERPROFILE%\Desktop\CreateLink\Character Map.lnk"^');
    WL(f,' "%SYSTEMROOT%\System32\charmap.exe"^');
    //~ WL(f,' "/dockbottom"^');
    WL(f,' ""^');
    WL(f,' "%SYSTEMROOT%\System32"^');
    WL(f,' "View all characters in any installed font"^');
    WL(f,' "%SYSTEMROOT%\System32\charmap.exe"^');
    WL(f,' ""^');
    WL(f,' 1'#13#10);    
    
    WL(f,'CreateLink "%USERPROFILE%\Desktop\CreateLink\Far Manager.lnk"^');
    WL(f,' "c:\Program Files\Far Manager\Far.exe"^');
    WL(f,' ""^');
    WL(f,' "c:\Program Files\Far Manager"^');
    WL(f,' "Classical File Manager"^');
    WL(f,' "c:\Program Files\Far Manager\Far.exe"^');
    WL(f,' 2^');
    WL(f,' 3'#13#10);    
    WL(f,'PAUSE');
    r:=0;
  end;
  SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$07);
  CloseHandle(f);{$i+}
  result:=ioresult+r;
end;

var 
  i,IcnNum,SW,e:integer;
  arg:array[1..8] of array[0..2*max_Path] of WideChar;

begin
  
  ZeroMemory(@arg,sizeOf(arg));
  // Get arguments
  for i:=1 to ParamCount do begin
    lstrcpyW(arg[i],PWideChar(ParamStrW(i)));    
    //MessageBoxW(0,arg[i],'GetCommandLineW',0);
  end;
  val(w2p(@arg[7][0]),IcnNum,e);if e<>0 then IcnNum:=0;
  val(w2p(@arg[8][0]),SW,e);if e<>0 then SW:=1;
  
  if paramCount<2 then begin
    
    if lstrcmpiW(arg[1],'x')=0 then begin      // Example
      SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$07);
      lstrcpyW(arg[1],'%USERPROFILE%\Desktop\Example.bat');ExpandEnvStrW(arg[1]);      
      if writeExample(arg[1])=0 then begin  
        write(' Written: "');Write(String(arg[1]));Writeln('"');
        Writeln;halt(0);  
      end else begin
        write(' Error writing "');Write(String(arg[1]));Writeln('"');  
        Writeln;halt(1);  
      end;
    end;    
   SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$0a);
    Writeln;
    Writeln(AppName);Writeln(Copyright);Writeln;
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$0b);
    Writeln('CREATELINK Link Target [Argument(s)] [WorkingDirectory] [Description] [IconPath] [IconNumber] [WndParams]'#13#10);    
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$07);  
    Write  ('Example (batch, using line-break ^):'#9#9#9);SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$0b);  
    Writeln('CREATELINK x   => EXAMPLE.BAT on user desktop'#13#10);    
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$07);  
    Writeln(' CreateLink "%USERPROFILE%\Desktop\Far Manager.lnk"^'#9'Link');
    Writeln('  "c:\Program Files\Far Manager\Far.exe"^'#9#9'Executable');    
    Writeln('  ""^'#9#9#9#9#9#9#9'Argument(s)');
    Writeln('  "c:\Program Files\Far Manager"^'#9#9#9'Working Directory');
    Writeln('  "Classical File Manager"^'#9#9#9#9'Description');
    Writeln('  "c:\Program Files\Far Manager\Far.exe"^'#9#9'Icon Path');
    Writeln('  2^'#9#9#9#9#9#9#9'Icon Number');
    Writeln('  3'#9#9#9#9#9#9#9'Show Command [1{Normal}|3{Maximized}|7{MinNoActive}]'#13#10);        
    Writeln('Overwrites existing link without prompt.'#13#10);
    halt(0);
  end;    
  
 if not CreateLinkExW(PWideChar(arg[2]),arg[3],arg[4],arg[1],arg[5],arg[6],IcnNum,SW) then begin    
    write(' Error creating: "');Write(String(ParamStr(1)));Writeln('"');    
    Writeln;
    halt(1);
  end else begin
    write(' Link created: "');Write(String(ParamStr(1)));Writeln('"');
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE),$07);
    Writeln;      
    halt(0);  
  end;  
end.
