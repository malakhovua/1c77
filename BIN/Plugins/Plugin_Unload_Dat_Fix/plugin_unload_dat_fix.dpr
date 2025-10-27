{$OPTIMIZATION OFF}

Library log_ert;
//������ ������� ���� ������� ����� ��������� � ����� ��������� ����� 1Cv77.dat.
//���� - ������� �������� � �������� ����������.
// ----------------------------------------------------------------------------
// 28.10.2011 - ���������� ��� Win7/ ������� ��������������� - MinWin (http://ru.wikipedia.org/wiki/MinWin)
// avgreen@molvest.org.ua

uses Windows, SysUtils, PatchMemory, HookMemory, Dialogs, uBalloon;

var g_HandleDat: Cardinal; //Handle ����� DAT
var g_NameDat: String; //��� ����� DAT
var g_Flag: Boolean;
var g_WriteMode: Boolean;
var g_IsOpenDat: Boolean;

var g_Checkpoint: Cardinal;
var g_Counter64: Int64;
var g_Size64: Int64;

var g_buf: array[0..65536] of char;
var g_pbuf: Integer;

var  g_Icon: HICON;
var g_IconWnd: HWND;

//var hm: THookMemory;



///////////////////////////////////////////////////////////////
procedure ShowBalloon(p_Message, p_Header: String);
begin
  if g_IconWnd=0 then begin
    g_IconWnd:=GetForegroundWindow();
    g_Icon := LoadIcon(0, IDI_APPLICATION);
    DZAddTrayIcon(g_IconWnd, 1, g_Icon, p_Header);
  end;
  DZBalloonTrayIcon(g_IconWnd, 1, 10, p_Message, p_Header, bitInfo);
//  Sleep(ms);
//  DZRemoveTrayIcon(g_IconWnd, 1);

end;


///////////////////////////////////////////////////////////////
Function GetDllPath: String;
var
  TheFileName : array[0..MAX_PATH] of char;
begin
  FillChar(TheFileName, sizeof(TheFileName), #0);
  GetModuleFileName(hInstance, TheFileName, sizeof(TheFileName));
  Result:= trim(''+TheFileName);
end;




///////////////////////////////////////////////////////////////
//��������� �� ������������ �������
type TCreateFile = function (
  lpFileName: pchar;
  dwDesiredAccess: DWORD;
  dwShareMode: DWORD;
  LPSECURITY_ATTRIBUTES: PSecurityAttributes;
  dwCreationDisposition: DWORD;
  dwFlagsAndAttributes: DWORD;
  hTemplateFile: DWORD
):Cardinal; StdCall;



///////////////////////////////////////////////////////////////
//��������� �� ������������ �������
type TCloseHandle = function (
  hObject: THandle
):BOOL; StdCall;


///////////////////////////////////////////////////////////////
//��������� �� ������������ �������
type TSetFilePointer = function (
    hFile: THandle; // handle of file
    lDistanceToMove: Longint;   // number of bytes to move file pointer
    lpDistanceToMoveHigh: Pointer;  // address of high-order word of distance to move
    dwMoveMethod: DWORD     // how to move
):DWORD; StdCall;


///////////////////////////////////////////////////////////////
//��������� �� ������������ �������
type TRead_File = function (
    pBuffer: Pointer;
    Bytes: DWORD
):DWORD; StdCall;


///////////////////////////////////////////////////////////////
var g_CreateFile_orig: TCreateFile;
var g_CloseHandle_orig: TCloseHandle;




///////////////////////////////////////////////////////////////
function OpenDatFile(fname: String): String;
var Dlg: TOpenDialog;
var dir: String;
begin
 dir:=SysUtils.GetCurrentDir;
 Dlg := TOpenDialog.Create(nil);
 Dlg.Filter := 'DAT ����� (*.dat)|*.dat|All files (*.*)|*.*';
 Dlg.InitialDir:=ExtractFilePath(fname);
 Dlg.DefaultExt:='DAT';
 Dlg.FileName:='romix.dat';
 Dlg.Title:='������� ����, ������ ����� �������� 1cv77.dat';
 if Dlg.Execute then begin
    Result:=Dlg.FileName;
 end else begin
    Result:=fname;
 end;
 Dlg.Free;
 SysUtils.SetCurrentDir(dir);
end;

///////////////////////////////////////////////////////////////
function GetFileSize1(hFile: THandle): Int64;
var hi, lo: DWORD;
begin
  lo:=Windows.GetFileSize(hFile, @hi);
  Result:=hi;
  Result:=(Result SHL 32) OR lo;
end;

///////////////////////////////////////////////////////////////
//�������-����������� CreateFile
function NewCreateFile(
  lpFileName: pchar;
  dwDesiredAccess: DWORD;
  dwShareMode: DWORD;
  LPSECURITY_ATTRIBUTES: PSecurityAttributes;
  dwCreationDisposition: DWORD;
  dwFlagsAndAttributes: DWORD;
  hTemplateFile: DWORD
):Cardinal; StdCall;

var s: String;
var p: Integer;
var newfn: String;
begin

  asm
    pusha
  end;

  s:=lpFileName;
  s:=LowerCase(s);
  s:=trim(s);
  p:=pos(LowerCase('1Cv77.dat'),s);
  g_Flag:=False;
  newfn:=lpFileName;

  if p>0 then begin
     if dwDesiredAccess=$40000000 then begin
       g_WriteMode:=True; //���� ������ �� ������
     end else begin
       g_WriteMode:=False;
     end;

    if not g_WriteMode then begin
      newfn:=OpenDatFile(newfn);
    end else begin
       ShowBalloon('��� �������� ����� ��������� ������������� ����� 1Cv77.dat (� ����� ZIP ������� ������ ���� DAT,'#13#10+
       ' � ��� 1Cv77.dat ����� ������ � �������� �������������� ���� ��� ������ romix.dat).', 'unload_dat_fix');
    end;
    g_Flag:=True;
  end;
  asm
    popa
  end;
  Result:=g_CreateFile_orig( //�������� ������������ CreateFile
      //lpFileName,
      pchar(newfn),
      dwDesiredAccess,
      dwShareMode,
      LPSECURITY_ATTRIBUTES,
      dwCreationDisposition,
      dwFlagsAndAttributes,
      hTemplateFile
  );

  if g_Flag then begin
    //MessageBox(0, pchar('Handle='+IntToHex(Result,8)), '***CreateFile',0);
    g_Size64:=GetFileSize1(Result);

    g_HandleDat:=Result;
    g_NameDat:=lpFileName;
    g_IsOpenDat:=True;
  end;

end;

///////////////////////////////////////////////////////////////
function FileNamesIsEqual(f1, f2: String): Boolean;
begin
  f1:=LowerCase(Trim(f1));
  f2:=LowerCase(Trim(f2));
  if f1=f2 then begin
    Result:=True;
  end else begin
    Result:=False;
  end;
end;


///////////////////////////////////////////////////////////////
//����� ���� � ������ g_NameDat � ���� � ���������; ������� ������ ���� ������ ���� ������ 1 ����.
procedure ReplaceDat();
var f: TextFile;
var newname: String;
begin

  newname:=ExtractFilePath(g_NameDat)+'romix.dat';

  if FileExists(newname) then begin
    DeleteFile(newname);
  end;
  RenameFile(g_NameDat, newname);
  AssignFile(f, g_NameDat);
  Rewrite(f);
  Write(f, '{}');
  CloseFile(f);
  ShowBalloon('�������� DAT ��������� � �����: '#13#10+newname+#13#10+
  '� ������ ZIP ��������� ������ ���� DAT', 'unload_dat_fix');

(*
  MessageBox(0, pchar('�������� DAT ��������� � �����: '#13#10+newname+#13#10+
  '� ������ ZIP ��������� ������ ���� DAT'), '(c)romix', MB_ICONINFORMATION);
*)

end;

//function CloseHandle(hObject: THandle): BOOL; stdcall;

///////////////////////////////////////////////////////////////
//�������-����������� CloseHandle
function NewCloseHandle(
hObject: THandle):BOOL; StdCall;
begin


  Result:=g_CloseHandle_orig( //�������� ������������ CreateFile
      hObject
  );

  asm
    pusha
  end;
  if g_NameDat<>'' then begin
    if g_HandleDat=hObject then begin
      //MessageBox(0, pchar('Handle='+IntToHex(hObject,8)), '***CloseHandle',0);
      //DZRemoveTrayIcon(Wnd, 1);

      if g_WriteMode then begin

        //if MessageBox(0, pchar('��������� ������������� ����� DAT? (� ����� ZIP ������� ������ ���� DAT).'), '(c)romix', MB_YESNO+MB_ICONQUESTION)=IDYES   then begin
          ReplaceDat();
        //end;
        g_NameDat:='';
        g_IsOpenDat:=False;
      end else begin
        //� ������ �������� ����� ��� ������
        //������� ����������� � ���� ��� ������������ ���������
        DZRemoveTrayIcon(g_IconWnd, 1);
        g_IconWnd:=0;
      end;
    end;
  end;
  asm
    popa
  end;
end;

//?Read@CBufdFile@@UAEIPAXI@Z



///////////////////////////////////////////////////////////////
//�������-����������� ?NextChar@CDB7Stream@@UAEXXZ
//public: virtual void __thiscall CDB7Stream::NextChar(void)



procedure New_NextChar(); stdcall;
var res: char;
var buf1: array[0..10] of char;
var handle: Cardinal;
var NumberOfBytesRead: Cardinal;
begin
  asm
    mov eax, [ecx+04]
    mov eax, [eax+04]
    mov handle, eax
    push ecx

  end;

  if handle=g_HandleDat then begin //���� ��� ���� ��������, �� ��������� ��� � ������������ (����� ��������)

    if g_pbuf=0 then begin
      //������ ���� � �����
      sleep(1);
      Windows.ReadFile(handle, g_buf, 65536, NumberOfBytesRead, nil);
      g_buf[NumberOfBytesRead+1]:=#0; //��������� ����� ����� �������� 0
    end;

    res:=g_buf[g_pbuf];

    asm
      mov eax, [g_pbuf]
      inc eax
      and eax, $FFFF //������� ���������� ��� ���������� 65536
      mov [g_pbuf], eax
    end;

    //if res=#0 then MessageBox(0, '��������� ����� �����', '*', 0);
  end else begin  //����� ��������� ������ ���� (���������, ��� �����������)
      Windows.ReadFile(handle, buf1, 1, NumberOfBytesRead, nil);
      if NumberOfBytesRead=0 then begin
        res:=#0;
      end else begin
        res:=buf1[0];
      end;
  end;

  //inc(g_pos32);
  inc(g_Checkpoint);
  inc(g_Counter64);

  asm
    pop ecx
    push eax
    mov al, byte ptr res
    mov [ecx+$38], al  //���������� ��������� ������ �� ����� ������
    pop eax
  end;

end;



///////////////////////////////////////////////////////////////
//?GetSinceCheckpoint@CDB7Stream@@QBEJXZ
//public: long __thiscall CDB7Stream::GetSinceCheckpoint(void)const

function New_GetSinceCheckpoint(): Cardinal; Stdcall;
begin
  result:=g_Checkpoint;
  //inc(g_Checkpoint);
end;

///////////////////////////////////////////////////////////////
//����������� ����� ��������
function IntToStr3(i64: Int64): String;
var s, s1: String;
i: Integer;
c: char;
begin
  s:=IntToStr(i64);
  s1:='';
  i:=Length(s);
  while i mod 3 <> 0 do begin
    s:=' '+s;
    i:=Length(s);
  end;
  for i:=1 to length(s) do begin
    c:=s[i];
    s1:=s1+c;
    if i mod 3=0 then s1:=s1+' ';
  end;
  Result:=trim(s1);
end;

///////////////////////////////////////////////////////////////
//����������, ������� ��������� n1 ���������� �� n2
function Percent(n1, n2: Int64): String;
var s: String;
var f: Extended;
begin
  f:=n1*10000/n2;
  f:=round(f)/100;
  s:=//FloatToStr(f);
  format('%5.2f', [f]);
  Result:=s+'%';
end;


var  g_Wnd, g_Wnd1: HWND;

///////////////////////////////////////////////////////////////
//?Checkpoint@CDB7Stream@@QAEXXZ
//public: void __thiscall CDB7Stream::Checkpoint(void)

procedure New_Checkpoint(); Stdcall;
begin
  //MessageBox(0, 'checkpoint' , '*', 0);



  g_Checkpoint:=0;

  if g_Wnd=0 then begin
    g_Wnd:=FindWindow(nil,'�������� ������');
    g_Wnd1:=GetWindow(g_wnd, GW_OWNER);
    //MessageBox(0,pchar(IntToHex(g_wnd,8)),pchar('����'), MB_ICONINFORMATION);
  end;


  ShowBalloon(
  ''+IntToStr3(g_Counter64 div (1024*1024))+' �� '+
   IntToStr3(g_Size64 div (1024*1024))+
   ' ��',
   '��������� '+Percent(g_Counter64, g_Size64));

  (*
  SetWindowText(g_wnd, pchar(
  '��������� '+IntToStr3(g_Counter64 div 1024)+'  ��  '+
   IntToStr3(g_Size64 div 1024)+
   ' ��������  ('+Percent(g_Counter64, g_Size64)+')'
  ));

  SetWindowText(g_Wnd1, pchar(
    Percent(g_Counter64, g_Size64)
  ));
  *)

end;

//  avgreen@molvest.org.ua
function WinMin : Boolean;
var
  OSVersionInfo : TOSVersionInfo;
begin
  Result := False;
  OSVersionInfo.dwOSVersionInfoSize := sizeof(TOSVersionInfo);
  if GetVersionEx(OSVersionInfo) then
      begin
         //MessageBox(0,pchar('Major = ' + IntToStr(OSVersionInfo.DwMajorVersion) + ' Minor = ' + IntToStr(OSVersionInfo.DwMinorVersion)),pchar('OS version'), MB_ICONERROR);
         if OSVersionInfo.DwMajorVersion >= 6 then Result := True;
      end;
end;

///////////////////////////////////////////////////////////////
procedure patch(DllName:String);
  var  pm: TPatchMemory;
  var DllNameToFind: String;
begin

     pm:=TPatchMemory.Create;

     //--------------------------------------
     //�������� ��������� ������� CreateFile ��� ��������� DLL

     pm.DllNameToPatch:=DllName;
//     pm.DllNameToFind:='KERNEL32.dll';   // avgreen@molvest.org.ua
     if WinMin then DllNameToFind := 'API-MS-Win-Core-File' else DllNameToFind := 'KERNEL32.dll'; // ������ ��� ���������� 'API-MS-Win-Core-File-L1-1-0.dll' �� ��������� ����� ���������� ������ � ��� � ����� ����� ������� - ���� �� ���������
     pm.DllNameToFind:=DllNameToFind;
     pm.FuncNameToFind:='CreateFileA';
     pm.NewFunctionAddr:=@NewCreateFile;
     try
       if pm.Patch then begin
         @g_CreateFile_orig:=pm.OldFunctionAddr;
       end;
     except
       on E:Exception do
       MessageBox(0,pchar(E.Message),pchar(DllName), MB_ICONERROR);
     end;

     //--------------------------------------
     //�������� ��������� ������� CloseHandle ��� ��������� DLL


     pm.DllNameToPatch:=DllName;
//     pm.DllNameToFind:='KERNEL32.dll';   // avgreen@molvest.org.ua
     if WinMin then DllNameToFind := 'API-MS-Win-Core-Handle' else DllNameToFind := 'KERNEL32.dll'; // ������ ��� ���������� 'API-MS-Win-Core-Handle-L1-1-0.dll' �� ��������� ����� ���������� ������ � ��� � ����� ����� ������� - ���� �� ���������
     pm.DllNameToFind:=DllNameToFind;
     pm.FuncNameToFind:='CloseHandle';
     pm.NewFunctionAddr:=@NewCloseHandle;
     try
       if pm.Patch then begin
         @g_CloseHandle_orig:=pm.OldFunctionAddr;
       end;
     except
       on E:Exception do
       MessageBox(0,pchar(E.Message),pchar(DllName), MB_ICONERROR);
     end;
     //MessageBox(0,pchar('*'),'', 0);

     //--------------------------------------
     pm.Free;
end;

//////////////////////////////////////////////////////////

  var  hm: THookMemory;
begin
  g_NameDat:='';
  g_IsOpenDat:=False;
  //MessageBox(0, 'plugin_unload_dat_fix v5', '(c) romix', MB_ICONINFORMATION);
  patch('mfc42.dll');


     hm:=THookMemory.Create;

     //--------------------------------------
     //�������� ������� ?NextChar@CDB7Stream@@UAEXXZ
     hm.DllName:='BkEnd.dll';
     hm.FuncName:='?NextChar@CDB7Stream@@UAEXXZ';
     hm.NewFunctionAddr:=@New_NextChar;

     try
       hm.Hook;
     except
       on E:Exception do
       MessageBox(0,pchar(E.Message),pchar('?NextChar@CDB7Stream@@UAEXXZ'), MB_ICONERROR);
     end;
     g_pbuf:=0;
     //g_pos32:=0;
     g_Checkpoint:=0;
     g_Counter64:=0;


     hm.DllName:='BkEnd.dll';
     hm.FuncName:='?GetSinceCheckpoint@CDB7Stream@@QBEJXZ';
     hm.NewFunctionAddr:=@New_GetSinceCheckpoint;

     try
       hm.Hook;
     except
       on E:Exception do
       MessageBox(0,pchar(E.Message),pchar('GetSinceCheckpoint@CDB7Stream@@QBEJXZ'), MB_ICONERROR);
     end;

     hm.DllName:='BkEnd.dll';
     hm.FuncName:='?Checkpoint@CDB7Stream@@QAEXXZ';
     hm.NewFunctionAddr:=@New_Checkpoint;

     try
       hm.Hook;
     except
       on E:Exception do
       MessageBox(0,pchar(E.Message),pchar('?Checkpoint@CDB7Stream@@QAEXXZ'), MB_ICONERROR);
     end;

     g_Wnd:=0;
     g_Wnd1:=0;
     g_IconWnd:=0;


end.
