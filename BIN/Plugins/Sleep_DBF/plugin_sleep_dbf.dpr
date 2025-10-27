//(c) romix, 2006

Library sleep_dbf;
//������ ��������� ����� Sleep � �������� ����������


uses
  Windows,
  SysUtils,
//  ShellAPI,

  PatchMemory,
  uBalloon,
  uINI in 'uINI.pas';

var g_Sleep: Integer; //������������ �������� �����, ��������, 1024
var g_CurrentSleep: Integer; //������� �������� �����, ���������� ��� 1,2,4...,1024
//var g_1SJOURN: DWORD; //���������� ����� 1sjourn
//var g_Loop: Boolean; //����������� ���� ��� ��������� �����������
var g_Balloon: Boolean; //���������� ��������� � ����
var g_SignalFile: String; //���������� ����
var g_CriticalSection: _RTL_CRITICAL_SECTION;



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
procedure GetIniFile();
var fname: String;
var ini: T_INI;
var dir: String;
begin

  EnterCriticalSection(g_CriticalSection);

  g_Sleep:=0;
  g_CurrentSleep:=1;
  g_Balloon:=True;
  g_SignalFile:='';

  ini:=T_INI.Create;
  try
    fname:=ChangeFileExt(GetDllPath(), '.ini');

    ini.Open(fname);

   while ini.ReadLine() do begin
      if ini.Name='�����' then begin
        g_Sleep:=ini.GetInteger;
      end else if ini.Name='���������' then begin
        g_Balloon:=ini.GetBoolean;
      end else if ini.Name='��������������' then begin
        g_SignalFile:=ini.Value;
        dir:=ExtractFilePath(g_SignalFile);
        if DirectoryExists(dir)=False then begin
          Raise Exception.Create('�������, ��������� � ��������� �������������� �� ����������: '+dir);
        end;
      end else begin
        Raise Exception.Create('����������� �������� ini-�����: '+fname+' '+ini.Name+'.');
      end;

    end;
    ini.Close;

  except
    on e:Exception do begin
     ini.Close;
     MessageBox(0, pchar('���������� ������ � ini-�����: '+fname+#13#10+
     pchar(e.Message)), 'plugin_sleep_dbf', MB_ICONWARNING);
    end;
  end;

  ini.Destroy;
  LeaveCriticalSection(g_CriticalSection);

end;


///////////////////////////////////////////////////////////////
procedure CreateSignalFile();
//������� ���������� ����
var f: TextFile;
begin

  if g_SignalFile='' then exit;
  {$I-}
  if FileExists(g_SignalFile) then exit;
  AssignFile(f, g_SignalFile);
  Rewrite(f);
  CloseFile(f);
  {$I+}
// *)


//  ShellExecute(0, 'open', pchar(g_SignalFile), nil, nil, SW_HIDE);


end;

///////////////////////////////////////////////////////////////
//��������� �� ������������ �������
type TLockFile = function (
    hFile,	// handle of file to lock
    dwFileOffsetLow,	// low-order word of lock region offset
    dwFileOffsetHigh,	// high-order word of lock region offset
    nNumberOfBytesToLockLow,	// low-order word of length to lock
    nNumberOfBytesToLockHigh:DWORD 	// high-order word of length to lock
):Boolean; StdCall;

var g_LockFile_orig: TLockFile;

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


///////////////////////////////////////////////////////////////
//�������-����������� LockFile
function NewLockFile(
    hFile,	// handle of file to lock
    dwFileOffsetLow,	// low-order word of lock region offset
    dwFileOffsetHigh,	// high-order word of lock region offset
    nNumberOfBytesToLockLow,	// low-order word of length to lock
    nNumberOfBytesToLockHigh:DWORD 	// high-order word of length to lock
):Boolean; StdCall;
//If the function fails, the return value is zero.

begin
    EnterCriticalSection(g_CriticalSection);
    Result:=g_LockFile_orig( //�������� ������������ LockFile
      hFile,	// handle of file to lock
      dwFileOffsetLow,	// low-order word of lock region offset
      dwFileOffsetHigh,	// high-order word of lock region offset
      nNumberOfBytesToLockLow,	// low-order word of length to lock
      nNumberOfBytesToLockHigh	// high-order word of length to lock
    );

      if Result=False then begin //��������� ������� ������������� ����
        if g_Sleep>0 then begin

          //������ �����. �������� ��� �������� - � ������� ������, � ���
          if (g_Balloon=True) and (g_CurrentSleep>250) then begin
            ShowBalloon('�������� ����������', '1�:�����������', g_CurrentSleep);
          end else begin
            Windows.Sleep(g_CurrentSleep); //������ �����
          end;

          //���������� ����� �����
          if g_CurrentSleep<g_Sleep then begin
            g_CurrentSleep:=g_CurrentSleep*2;
          end;


          //���������� ���� ��� ��������� ����������
          if g_CurrentSleep>=g_Sleep then begin
              CreateSignalFile();
          end;


          Result:=g_LockFile_orig( //�������� ������������ LockFile
            hFile,	// handle of file to lock
            dwFileOffsetLow,	// low-order word of lock region offset
            dwFileOffsetHigh,	// high-order word of lock region offset
            nNumberOfBytesToLockLow,	// low-order word of length to lock
            nNumberOfBytesToLockHigh	// high-order word of length to lock
          );


        end;
      end else begin //��� ������� ������� ���������� ���������� ����� � 1
        g_CurrentSleep:=1;
      end;
      LeaveCriticalSection(g_CriticalSection);
end;




///////////////////////////////////////////////////////////////
procedure patch(DllName:String);
  var  pm: TPatchMemory;
begin

     pm:=TPatchMemory.Create;
     //--------------------------------------
     pm.DllNameToPatch:=DllName;
     pm.DllNameToFind:='KERNEL32.dll';
     pm.FuncNameToFind:='LockFile';


     pm.NewFunctionAddr:=@NewLockFile;

     try
       if pm.Patch then begin
       @g_LockFile_orig:=pm.OldFunctionAddr;
       end;
     except
     end;
     //--------------------------------------
     pm.Free;
end;







//////////////////////////////////////////////////////////
begin

  InitializeCriticalSection(g_CriticalSection);

  //MessageBox(0, 'Start', 'sleep_dbf', MB_ICONINFORMATION);

  GetIniFile();
  patch('dbeng32.dll');
  end.
