Library hook_1c; //(c) romix, 2007
uses Windows, SysUtils, ShellAPI;


type T_ShellExecuteA = function (hWnd: HWND; Operation, FileName, Parameters, Directory: PAnsiChar; ShowCmd: Integer): HINST; stdcall;
type T_ShellExecuteW = function (hWnd: HWND; Operation, FileName, Parameters, Directory: PWideChar; ShowCmd: Integer): HINST; stdcall;
type T_ShellExecuteExA = function (lpExecInfo:Pointer): BOOL; stdcall;
type T_ShellExecuteExW = function (lpExecInfo:Pointer): BOOL; stdcall;

var hDll: HMODULE;

var pShellExecuteA: T_ShellExecuteA;
var pShellExecuteW: T_ShellExecuteW;
var pShellExecuteExA: T_ShellExecuteExA;
var pShellExecuteExW: T_ShellExecuteExW;



type T_INI = Class
  Name: String;
  Value: String;


  procedure Open(fname: String);
  procedure Close();
  function ReadLine(): Boolean;
  function GetInteger: Integer;
  function GetBoolean: Boolean;

  Constructor Create;
  Destructor Destroy; Override;

private
  g_f: TextFile;
  g_s: String;
  isOpen: Boolean;
end;



//////////////////////////////////////////////////////////////
procedure T_INI.Open(fname: String);
begin
    {$I+}
    try
      Assign(g_f, fname);
      Reset(g_f);
      isOpen:=True;
    except
      on E:Exception do begin
        {$I-}CloseFile(g_f);{$I+}
        Raise Exception.Create('Ошибка открытия файла: '+E.Message);
      end;
    end;
end;

//////////////////////////////////////////////////////////////
procedure T_INI.Close();
begin
    {$I-}CloseFile(g_f);{$I+}
    isOpen:=False;
end;

//////////////////////////////////////////////////////////////
Function T_INI.ReadLine(): Boolean;
var p:Integer;
begin
   if IsOpen=False then begin
     Raise Exception.Create('Файл ini необходимо сначала открыть.');
   end;
{$I+}
   g_s:='';
   Name:='';
   Value:='';

   Repeat
    try
      if eof(g_f) then begin
        Result:=False;
        Exit;
      end;
      Readln(g_f, g_s);
    except
      on E:Exception do begin
        {$I-}CloseFile(g_f);{$I+}
        Raise Exception.Create('Ошибка чтения файла ini:'+ E.Message);
      end;
    end;


    Result:=True;


    p:=pos(';', g_s);
    if p>0 then g_s:=Copy(g_s,1,p-1); //отрезаем комментарий после ;

    g_s:=trim(g_s);

    p:=pos('=', g_s);
    if p=0 then g_s:='';

   Until g_s<>'';


    Name:=Copy(g_s,1, p-1); //имя в паре Имя=Значение
    Name:=trim(Name);

    Value:=Copy(g_s, p+1, Length(g_s)-p+1); //значение
    Value:=trim(Value);

end;

//////////////////////////////////////////////////////////////
function T_INI.GetInteger: Integer;
begin
  try
    Result:=StrToInt(Value);
  except
    Raise Exception.Create('Ошибка чтения значения '+Name+' - ожидается целое число.');
  end;
end;

//////////////////////////////////////////////////////////////
function T_INI.GetBoolean: Boolean;
begin
    if Value='1' then Result:=True
    else if Value='Да' then Result:=True
    else if Value='Д' then Result:=True
    else if Value='Yes' then Result:=True
    else if Value='Y' then Result:=True
    else if Value='True' then Result:=True
    else if Value='Истина' then Result:=True

    else if Value='0' then Result:=False
    else if Value='Нет' then Result:=False
    else if Value='Н' then Result:=False
    else if Value='No' then Result:=False
    else if Value='N' then Result:=False
    else if Value='Ложь' then Result:=False
    else if Value='False' then Result:=False

    else
    Raise Exception.Create('Ошибка чтения значения '+Name+' - ожидается значение Да,Д,Yes,Y,1,True,Истина  либо  Нет,Н,No,N,0,False,Ложь.');
end;


//////////////////////////////////////////////////////////////
Constructor T_INI.Create;
begin
     inherited Create;
     isOpen:=False;
end;

//////////////////////////////////////////////////////////////
destructor T_INI.Destroy;
begin
     {$I-}CloseFile(g_f);{$I+}
     inherited Destroy;
end;




//////////////////////////////////////////////////////////
function sGetModuleFileName(): String;
var p:pchar;
begin
  GetMem(p,MAX_PATH);
  GetModuleFileName(0, p, MAX_PATH-1);
  result:=Trim(Strpas(p));
  Freemem(p);
end;

///////////////////////////////////////////////////////////////
procedure LoadDlls;
var fname: String;
var dir: String;
var ini: T_INI;
begin
  dir:=ExtractFilePath(sGetModuleFileName())+'Plugins\';
  fname:=dir+'Hook_1C.ini';

  //MessageBox(0, pchar(fname), 'ini', 0);

  if not DirectoryExists(dir) then begin

    MessageBox(0, pchar('Не найден каталог: '+#13#10+
    dir), 'Hook_1C.dll', MB_ICONWARNING);
    exit;
  end;


  if not FileExists(fname) then begin
    MessageBox(0, pchar('Не найден файл: '+#13#10+
    fname), 'Hook_1C.dll', MB_ICONWARNING);
    exit;
  end;

  ini:=T_INI.Create();
  ini.Open(fname);
  While ini.ReadLine Do begin
    if ini.name='LoadDll' then begin
      if LoadLibrary(pchar(dir+ini.value))=0 then begin
        MessageBox(0,pchar('Ошибка загрузки DLL: '+pchar(dir+ini.value)), 'LoadDll', MB_ICONWARNING);
      end;

    end else if ini.name='Message' then begin
       MessageBox(0,pchar(ini.value), 'Hook_1C (c) romix', MB_ICONINFORMATION);
    end;
  end; //while

  ini.Close;

end;






//////////////////////////////////////////////////////////
function sGetSystemDirectory(): String;
var p:pchar;
begin
  GetMem(p,MAX_PATH);
  GetSystemDirectory(p,MAX_PATH-1);
  result:=IncludeTrailingPathDelimiter(Strpas(p));
  Freemem(p);
end;



//Функция перехватчик
//////////////////////////////////////////////////////////
  function ShellExecuteA(hWnd: HWND; Operation, FileName, Parameters,
  Directory: PAnsiChar; ShowCmd: Integer): HINST; export; stdcall;
  begin
    Result:=pShellExecuteA(hWnd, Operation, FileName, Parameters,  Directory, ShowCmd);
  end;
  exports ShellExecuteA;

//Функция перехватчик
//////////////////////////////////////////////////////////
  function ShellExecuteW(hWnd: HWND; Operation, FileName, Parameters,
  Directory: PWideChar; ShowCmd: Integer): HINST; stdcall;
  begin
    Result:=pShellExecuteW(hWnd, Operation, FileName, Parameters,  Directory, ShowCmd);
  end;
  exports ShellExecuteW;


//Функция перехватчик
//////////////////////////////////////////////////////////
  function ShellExecuteExA(lpExecInfo:Pointer): BOOL; export; stdcall;
  begin
    Result:=pShellExecuteExA(lpExecInfo);
  end;
  exports ShellExecuteExA;

//Функция перехватчик
//////////////////////////////////////////////////////////
  function ShellExecuteExW(lpExecInfo:Pointer): BOOL; export; stdcall;
  begin
    Result:=pShellExecuteExW(lpExecInfo);
  end;
  exports ShellExecuteExW;


//////////////////////////////////////////////////////////
var DllName: String;
var ExeName: String;
begin



  ExeName:=LowerCase(sGetModuleFileName());

  DllName:='shell32.dll';
  hDll:=LoadLibrary(pchar(DllName));
  if(hDll <= 0) then raise exception.create('Неудачный вызов LoadLibrary '+DllName);


  @pShellExecuteA:=GetProcAddress(hDll, 'ShellExecuteA');
  if not assigned(pShellExecuteA) then raise exception.Create(' Неудачный вызов GetProcAddress "ShellExecuteA"');

  @pShellExecuteExA:=GetProcAddress(hDll, 'ShellExecuteExA');
  if not assigned(pShellExecuteExA) then raise exception.Create(' Неудачный вызов GetProcAddress "ShellExecuteExA"');

  @pShellExecuteW:=GetProcAddress(hDll, 'ShellExecuteW');
  if not assigned(pShellExecuteW) then raise exception.Create(' Неудачный вызов GetProcAddress "ShellExecuteW"');

  @pShellExecuteExW:=GetProcAddress(hDll, 'ShellExecuteExW');
  if not assigned(pShellExecuteExW) then raise exception.Create(' Неудачный вызов GetProcAddress "ShellExecuteExW"');

  LoadDlls();

end.
