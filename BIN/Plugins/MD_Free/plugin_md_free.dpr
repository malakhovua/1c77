//(c) romix, 2006

Library plugin_md_free;
uses
  Windows,
  SysUtils,
  PatchMemory,
  uBalloon;

var g_PathToMD: String; //���� � ����� MD
var g_Thread: THandle;
var g_ThreadID: DWORD;

var g_Delay: Integer;//����� � ��������
var g_Message: String;//��������� � ����


///////////////////////////////////////////////////////////////
procedure GetIniFile();
var f: TextFile;
var fname: String;
var s: String;
var p: Integer;
var name, value: String;
begin
  fname:=g_PathToMD+'MD_FREE\stop.ini';
  //MessageBox(0, pchar(fname), 'ini', 0);

  g_Delay:=0;
  g_Message:='';

  if FileExists(fname)=False then Exit;
  //MessageBox(0, pchar(fname), '���� ����������', 0);

  AssignFile(f, fname);
  Reset(f);
  Repeat
    Readln(f, s);
    p:=pos(';', s);
    if p>0 then s:=Copy(s,1,p-1); //�������� ����������� ����� ;
    p:=pos('=', s);
    if p=0 then Continue;
    name:=Copy(s,1, p-1); //��� � ���� ���=��������
    name:=trim(name);
    value:=Copy(s, p+1, Length(s)-p+1); //��������
    value:=trim(value);

    //MessageBox(0, pchar(name), pchar(value), 0);

    if name='��������������' then begin
      g_Delay:=StrToInt(Value);
    end;
    if name='���������' then begin
      g_Message:=Trim(Value);
    end;

  Until Eof(f);
  CloseFile(f);

end;


///////////////////////////////////////////////////////////////
procedure ShowBalloon(p_Message, p_Header: String; ms: Integer);
var
  Icon: HICON;
  Wnd: HWND;

begin
  Wnd:=GetForegroundWindow();
  Icon := LoadIcon(0, IDI_APPLICATION);
  DZAddTrayIcon(Wnd, 1, Icon, p_Header);
  DZBalloonTrayIcon(Wnd, 1, 10, p_Message, p_Header, bitInfo);
  Sleep(ms);
  DZRemoveTrayIcon(Wnd, 1);

end;


//////////////////////////////////////////////////////////
procedure Close1C();
var cnt: Integer;
begin
     GetIniFile();
     if g_Delay=0 then Exit;
     if g_Message='' then g_Message:= '����� ����������� �������� 1�:�����������. ';

     cnt:=g_Delay;

     repeat
       ShowBalloon(g_Message+#13#10+g_PathToMD, '�������� '+IntToStr(cnt)+' ������', 750);
       Sleep(250);
       Dec(cnt);
     until cnt=0;

     Windows.TerminateProcess(Windows.GetCurrentProcess(),1);
end;

//////////////////////////////////////////////////////////
procedure ThreadFunc(P: Pointer); stdcall;
begin

     repeat
       //���������, ���������� �� ���������� ����
       //MessageBox(0, pchar(g_PathToMD+'MD_Free\stop.ini'), '', 0);
       if SysUtils.FileExists(g_PathToMD+'MD_Free\stop.ini') then begin
         //��������� ������ 1�
         Close1C();
       end;
       //Windows.MessageBeep(MB_OK); //�������� ������
       Windows.sleep(5000); //����� � �������������, ����� �� ��������� ���������
     until False;
end;




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

var g_CreateFile_orig: TCreateFile;




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
begin
  asm
    pusha
  end;

  s:=lpFileName;
  s:=LowerCase(s);
  s:=trim(s);
  p:=pos('\syslog\1cv7.mlg',s);
  if p>0 then begin
    Delete(s, p+1, Length(s)-p+1);
    g_PathToMD:=s;
    //MessageBox(0, pchar(s), 'MD_FREE', 0);
  end;

  asm
    popa
  end;

//E:\MD_CHANGER\Hook_1C\Plugins\v77\SYSLOG\1cv7.mlg


    Result:=g_CreateFile_orig( //�������� ������������ CreateFile
        lpFileName,
        dwDesiredAccess,
        dwShareMode,
        LPSECURITY_ATTRIBUTES,
        dwCreationDisposition,
        dwFlagsAndAttributes,
        hTemplateFile
    );


end;

///////////////////////////////////////////////////////////////
function GetTempPathStr: String;
var
  Buffer: array[0..1023] of Char;
begin
  SetString(Result, Buffer, GetTempPath(Sizeof(Buffer)-1,Buffer));
end;



///////////////////////////////////////////////////////////////
procedure patch(DllName: String);
  var  pm: TPatchMemory;
begin

//MessageBox(0,pchar(''+DllName),'md_free',0);

     pm:=TPatchMemory.Create;

     //--------------------------------------
     //�������� ��������� ������� CreateFile ��� ��������� DLL

     pm.DllNameToPatch:=DllName;
     pm.DllNameToFind:='KERNEL32.dll';
     pm.FuncNameToFind:='CreateFileA';
     pm.NewFunctionAddr:=@NewCreateFile;
     try
       pm.Patch;
       @g_CreateFile_orig:=pm.OldFunctionAddr;
     except
     end;


     //--------------------------------------
     pm.Free;
end;





//////////////////////////////////////////////////////////
begin
  //MessageBox(0, 'Start', 'md_free', MB_ICONINFORMATION);
  //patch('frame.dll');
  patch('seven.dll');
  g_PathToMD:='';
  g_Thread := Windows.CreateThread(nil, 0, @ThreadFunc, nil, 0, g_ThreadID);
  if g_Thread = 0 then Windows.MessageBox(0, '������ ������������� ������', 'MD_Free', MB_ICONWARNING);


end.
