//(c) romix, 2006

Library TestThread;
uses Windows, SysUtils;

var g_Thread: THandle;
var g_ThreadID: DWORD;

procedure ThreadFunc(P: Pointer); stdcall;
begin
     repeat
       MessageBeep(MB_OK);
       sleep(1000); //����� � �������������, ����� �� ��������� ���������
     until False;
end;


//////////////////////////////////////////////////////////
begin

  g_Thread := CreateThread(nil, 0, @ThreadFunc,
    nil, 0, g_ThreadID);
  if g_Thread = 0 then MessageBox(0, '������ ������������� ������', 'TestThread', MB_ICONINFORMATION);

  MessageBox(0, '�������� TestThread.dll ������ �������', 'TestThread', MB_ICONINFORMATION);

end.
