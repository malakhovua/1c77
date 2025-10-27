unit PatchMemory;


Interface
uses Windows, ImageHlp, SysUtils;

var  BasePointer: pointer;
type TPatchMemory = Class
  DllNameToPatch: String; //��� DLL, � ������ ������� ������� ����� ����������� ���������
  DllNameToFind: String; //��� DLL, ������� ������� �� ����� ����������� (��������, 'KERNEL32.dll')
  FuncNameToFind: String; //�������, ������� �� ����� ����������� (��������, 'CreateProcessA')
  NewFunctionAddr: Pointer; //����� ������� - ����������
  OldFunctionAddr: Pointer; //������ ����� ���������� �������
  procedure Patch;  //��������� ������ ����������� ������� �� ����
  procedure UnPatch;  //�������� �������

        Constructor Create;
        Destructor Destroy; Override;

private

  AddrWherePatching: Pointer; //��� UnPatch
  function GetDwordByRVA(rva: dword):dword;
  function GetStringByRVA(rva: dword):pchar;
  function GetPointerByRVA(rva: DWORD):Pointer;
  function GetRVAByPointer(P: Pointer):DWORD;
  procedure WriteDwordToMemory(Kuda: Pointer; Data: DWORD);

end;


//////////////////////////////////////////////////////////
implementation

//////////////////////////////////////////////////////////
procedure say(s:String);
//���������� ������
begin
//  MessageBox(0,pchar(s),'',0);
end;


//////////////////////////////////////////////////////////////
Constructor TPatchMemory.Create;
begin
     inherited Create;
end;

//////////////////////////////////////////////////////////////
destructor TPatchMemory.Destroy;
begin
     say('Destroy');
     inherited Destroy;

end;



//////////////////////////////////////////////////
function TPatchMemory.GetDwordByRVA(rva: dword):dword;
//�������� �������� (4 �������� ����� ��� ����� - DWORD) �� ������ ��
//�������� (RVA)
begin
   asm
     push ebx;

     mov ebx, [rva];
     add ebx, [BasePointer];

     mov eax, [ebx];
     mov Result, eax;

     pop ebx;
   end;
end;

//////////////////////////////////////////////////
function TPatchMemory.GetStringByRVA(rva: dword):pchar;
//�������� ������ �� ��������� RVA
begin
   asm
     mov eax, [rva];
     add eax, [BasePointer];
     mov Result, eax;
   end;
end;

//////////////////////////////////////////////////
function TPatchMemory.GetPointerByRVA(rva: DWORD):Pointer;
//�������� ��������� �� RVA (�.�. ���������� � rva �������� BasePointer)
begin
   asm
     mov eax, rva;
     add eax, [BasePointer];
     mov Result, eax;
   end;
end;

//////////////////////////////////////////////////
function TPatchMemory.GetRVAByPointer(P: Pointer):DWORD;
//�������� ��������� RVA �� ���������
//(�.�. �������� �� ��������� �������� BasePointer)
begin
  asm
    mov eax, [p];
    sub eax, BasePointer;
    mov Result, eax;
  end;
end;

//////////////////////////////////////////////////
Procedure TPatchMemory.WriteDwordToMemory( //����� 4 ����� � ������ �� ���������.
  Kuda: Pointer; //����� ���� �����
  Data: DWORD //�������� ������� �����
);

var  BytesWritten: DWORD;
var  hProcess: THandle;
var old: DWORD;
begin

  hProcess := GetCurrentProcess(//������� WinAPI
  );


    //��������� ������ �� ������ �� ���������� ������

    VirtualProtect(//������� WinAPI
      kuda, //�����
      4, //����� ���� (����� ��������� 1 ��� 2 �������� ������ �������� 4�)
      PAGE_EXECUTE_READWRITE, //�������� �������
      @old); //���� ����� ���������� ������ ��������

    BytesWritten:=0;

    //���������� 4 �����
    WriteProcessMemory(//������� WinAPI
      hProcess,
      kuda,
      @data,
      4,
      BytesWritten);

 //��������������� ������� �������� �������
    VirtualProtect(//������� WinAPI
      kuda,
      4,
      old,
      @old);

end;



//////////////////////////////////////////////////////////
procedure TPatchMemory.Patch;
var ulSize: ULONG;
var rva: DWORD;
var ImportTableOffset: pointer;
var OffsetDllName: DWORD;
var OffsetFuncAddrs: DWORD;
var FunctionAddr: DWORD;
var DllName: String;
var FuncAddrToFind: DWORD;
begin

//� Windows 2000 DLL � ���� ������, ��������, ��� �� ���������
   if DWORD(GetModuleHandle(pchar(DllNameToPatch)))=0 then begin
      LoadLibrary(pchar(DllNameToPatch));
   end;

//�������� ����� �������, ������� �� ����� ����� � ������� IAT
  FuncAddrToFind:= DWORD(
    GetProcAddress(//������� Windows API
      GetModuleHandle(//������� Windows API
        pchar(DllNameToFind)
      ),
      pchar(FuncNameToFind)
    ));

  say(DllNameToFind+' '+FuncNameToFind);
  say('FuncAddrToFind: '+IntToHex(FuncAddrToFind,8));
  OldFunctionAddr:=Pointer(FuncAddrToFind);

//�������� �����, �� �������� ����������� � ������ dll - "������".
  BasePointer:=pointer(GetModuleHandle(//������� Windows API
  pChar(DllNameToPatch)));

//�������� �������� ������� �������
  ImportTableOffset:= ImageDirectoryEntryToData( //������� Windows API
  BasePointer, TRUE, IMAGE_DIRECTORY_ENTRY_IMPORT, ulSize);

//��������� �������� ������� ������� � ������ RVA
//(�������� ������������ ������ DLL)
  rva:=GetRvaByPointer(ImportTableOffset);

  repeat    {�������� �� ������� �������.
     ������ ������ � ������� ������� ����� ����� 20 ����:
     +0 - ��������� �� ������� ���� ������� (��� Borland �� ���)
     +4 - ?
     +8 - ?
     +12 - ��������� (RVA) �� ��� DLL
     +16 - ��������� �� ������� ������� �������
     }

         OffsetDllName := GetDwordByRVA(rva+12);
         if OffsetDllName = 0 then break; //���� ������� ���������, �������

          DllName := GetStringByRVA(OffsetDllName);//��� DLL

               OffsetFuncAddrs:=GetDwordByRva(rva+16);
                 repeat //���� �� ������ ������� DLL
                   FunctionAddr:=Dword(GetDwordByRva(OffsetFuncAddrs));
                   if FunctionAddr=0 then break; //���� ������� �����������, �������

                   if FunctionAddr=FuncAddrToFind then begin
                    //����� - ��������� ����
                     AddrWherePatching:=GetPointerByRva(OffsetFuncAddrs);
                     WriteDwordToMemory(
                      AddrWherePatching,
                      DWORD(NewFunctionAddr));
                   //say('AddrWherePatching: '+IntToHex(Dword(AddrWherePatching),8));
                   end;



                 inc(OffsetFuncAddrs,4);
               until false;
          rva:=rva+20;
  until false;

end;


//////////////////////////////////////////////////////////
procedure TPatchMemory.UnPatch;
begin
   say('AddrWherePatching: '+IntToHex(Dword(AddrWherePatching),8));
   say('OldFunctionAddr: '+IntToHex(Dword(OldFunctionAddr),8));
   WriteDwordToMemory(
     AddrWherePatching,
     DWORD(OldFunctionAddr));
end;


end.
