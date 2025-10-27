unit HookMemory;


Interface
uses Windows, ImageHlp, SysUtils;

var  BasePointer: pointer;
type THookMemory = Class
  DllName: String; //��� DLL, � ������� ����� ����������� ���������
  FuncName: String; //�������, ������� �� ����� ����������� (��������, 'CreateProcessA')
  NewFunctionAddr: Pointer; //����� ������� - ����������
  procedure Hook;//��������� ������ ����������� ������� �� ����
  procedure UnHook;  //�������� �������
  procedure ReHook;

        Constructor Create;
        Destructor Destroy; Override;

private
  hProcess: THandle;
  FuncAddr: Pointer;
  OldString5: Array[1..5] of char;
  NewString5: Array[1..5] of char;
end;


//////////////////////////////////////////////////////////
implementation

//////////////////////////////////////////////////////////
procedure say(s:String);
//���������� ������
begin
  MessageBox(0,pchar(s),'',0);
end;


//////////////////////////////////////////////////////////////
Constructor THookMemory.Create;
begin
 //say('Create');
  inherited Create;
  hProcess := Windows.GetCurrentProcess();
end;

//////////////////////////////////////////////////////////////
destructor THookMemory.Destroy;
begin
     inherited Destroy;
end;



//////////////////////////////////////////////////////////
procedure  ReadMem(const Adr; Const Buf; count : Integer );
var
  S, D: PChar;
  I: Integer;
begin
  S := PChar(Adr);
  D := PChar(@Buf);
  for I := 0 to count-1 do
  D[I] := S[I];
end;


//////////////////////////////////////////////////////////
procedure  WriteMem(Const Adr; Const Buf; count : Integer );
var
  S, D: PChar;
  I: Integer;
begin
  S := PChar(@Buf);
  D := PChar(Adr);
  for I := 0 to count-1 do
  D[I] := S[I];
end;


//////////////////////////////////////////////////////////
procedure THookMemory.Hook;
var off: Integer;
var old: DWORD;
begin


//� Windows 2000 DLL � ���� ������, ��������, ��� �� ���������
   if DWORD(GetModuleHandle(pchar(DllName)))=0 then begin
      LoadLibrary(pchar(DllName));
   end;


//�������� ����� �������, ������� �� ����� �����
  FuncAddr:=
    GetProcAddress(//������� Windows API
      GetModuleHandle(//������� Windows API
        pchar(DllName)
      ),
      pchar(FuncName)
    );


  NewString5[1]:=chr($E9);
  off:=Integer(NewFunctionAddr) - Integer(FuncAddr) - 5;

  Move(off, NewString5[2], 4);

    if VirtualProtect(//������� WinAPI
      FuncAddr, //�����
      5, //����� ���� (����� ��������� 1 ��� 2 �������� ������ �������� 4�)
      PAGE_EXECUTE_READWRITE, //�������� �������
      @old) //���� ����� ���������� ������ ��������
      =False then
        Raise Exception.Create('������ ��� VirtualProtect');

   ReadMem(FuncAddr, OldString5, 5);

   WriteMem(FuncAddr, NewString5, 5);

end;


//////////////////////////////////////////////////////////
procedure THookMemory.UnHook;
begin
  WriteMem(FuncAddr, OldString5, 5);
end;

//////////////////////////////////////////////////////////
procedure THookMemory.ReHook;
begin
  WriteMem(FuncAddr, NewString5, 5);
end;

end.
