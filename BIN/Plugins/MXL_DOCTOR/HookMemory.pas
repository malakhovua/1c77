unit HookMemory;


Interface
uses Windows, ImageHlp, SysUtils;

var  BasePointer: pointer;
type THookMemory = Class
  DllName: String; //Имя DLL, в которой будем производить изменения
  FuncName: String; //Функция, которую мы хотим перехватить (например, 'CreateProcessA')
  NewFunctionAddr: Pointer; //Адрес функции - заменителя
  procedure Hook;//Выполняет замену стандартной функции на нашу
  procedure UnHook;  //Отменяет патчинг
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
//Отладочная печать
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


//В Windows 2000 DLL в этот момент, возможно, еще не загружена
   if DWORD(GetModuleHandle(pchar(DllName)))=0 then begin
      LoadLibrary(pchar(DllName));
   end;


//Получаем адрес функции, который мы хотим найти
  FuncAddr:=
    GetProcAddress(//Функция Windows API
      GetModuleHandle(//Функция Windows API
        pchar(DllName)
      ),
      pchar(FuncName)
    );


  NewString5[1]:=chr($E9);
  off:=Integer(NewFunctionAddr) - Integer(FuncAddr) - 5;

  Move(off, NewString5[2], 4);

    if VirtualProtect(//Функция WinAPI
      FuncAddr, //адрес
      5, //число байт (будут затронуты 1 или 2 страницы памяти размером 4К)
      PAGE_EXECUTE_READWRITE, //атрибуты доступа
      @old) //сюда будут возвращены старые атрибуты
      =False then
        Raise Exception.Create('Ошибка при VirtualProtect');

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
