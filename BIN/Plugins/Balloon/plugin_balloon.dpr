//(c) romix, 2006

Library plugin_balloon;
uses Windows, SysUtils, PatchMemory, uBalloon, Classes;

//Заменяет указанные сообщения MessageBox на сообщения в трее (Balloon tooltip)

var g_sz: TStringList;


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
procedure GetIniFile;
var f: TextFile;
var fname: String;
var s: String;
var p: Integer;
var name, value: String;
begin
  fname:=ChangeFileExt(GetDllPath(), '.ini');

  //MessageBox(0, pchar(fname), 'ini', 0);


  if FileExists(fname)=False then Exit;
  //MessageBox(0, pchar(fname), 'Файл существует', 0);

  AssignFile(f, fname);
  Reset(f);
  Repeat
    Readln(f, s);
    p:=pos(';', s);
    if p>0 then s:=Copy(s,1,p-1); //отрезаем комментарий после ;
    p:=pos('=', s);
    if p=0 then Continue;
    name:=Copy(s,1, p-1); //имя в паре Имя=Значение
    name:=trim(name);
    value:=Copy(s, p+1, Length(s)-p+1); //значение
    value:=trim(value);

    //MessageBox(0, pchar(name), pchar(value), 0);

    g_sz.Add(name+'='+value)

  Until Eof(f);
  CloseFile(f);

end;


//Указатель на оригинальную функцию
type TMessageBox = function (

    hWnd: HWND;	// handle of owner window
    lpText: pchar;	// address of text in message box
    lpCaption: pchar;	// address of title of message box
    uType: DWORD 	// style of message box
):Cardinal; StdCall;

var g_MessageBox_orig: TMessageBox;



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
//Функция-перехватчик MessageBox
function NewMessageBox(
    hWnd: HWND;	// handle of owner window
    lpText: pchar;	// address of text in message box
    lpCaption: pchar;	// address of title of message box
    uType: DWORD 	// style of message box
):Cardinal; StdCall;


var s, name: String;
var i: Integer;

begin

   s:=lpText;

   try

   for i:=0 to g_sz.Count-1 do begin
     name:=g_sz.Names[i];
     if pos(name,s)>0 then begin
       ShowBalloon(lpText, lpCaption, StrToInt(g_sz.Values[name]));
       Result:=0;
       Exit;
     end;
   end;

   except
     on E:Exception do begin
       s:=E.Message;
     end;
   end;

    Result:=g_MessageBox_orig( //вызываем оригинальную MessageBoxA
        hWnd,
        pchar(s),
        pchar(''+lpCaption),
        uType
    );


end;





///////////////////////////////////////////////////////////////
procedure patch(DllName:String);
  var  pm: TPatchMemory;
  var s: String;
begin

//MessageBox(0,pchar(''+DllName),'md_changer',0);

     s:=LowerCase(''+DllName);
     //Плагины не патчим
     if pos('plugin_', s)>0 then Exit;

     pm:=TPatchMemory.Create;

     //--------------------------------------
     //Перехват системной функции CreateFile для указанной DLL

     pm.DllNameToPatch:=DllName;
     pm.DllNameToFind:='USER32.dll';
     pm.FuncNameToFind:='MessageBoxA';
     pm.NewFunctionAddr:=@NewMessageBox;
     try
       pm.Patch;
       @g_MessageBox_orig:=pm.OldFunctionAddr;
     except
     end;



     //--------------------------------------
     pm.Free;
end;


/////////////////////////////////////////////////
type
  TModuleArray = array[0..400] of HMODULE;

/////////////////////////////////////////////////
//Перечисляет все DLL текущего процесса
//Вызывает для каждой из них процедуру patch(ИмяDLL);
function GetLoadedDLLList(): Boolean;
type
EnumModType = function (hProcess: Longint; lphModule: TModuleArray;
  cb: DWord; var lpcbNeeded: Longint): Boolean; stdcall;
var
  psapilib: HModule;
  EnumProc: Pointer;
  ma: TModuleArray;
  I: Longint;
  FileName: array[0..MAX_PATH] of Char;
  S: string;
begin
  Result := False;

  (* Данная функция запускается только для Widnows NT *)
  if Win32Platform <> VER_PLATFORM_WIN32_NT then begin
    MessageBox(0, 'Данная функция запускается только для Widnows NT/2000/XP', 'md_changer', 0);
    Exit;
  end;

  psapilib := LoadLibrary('psapi.dll');
  if psapilib = 0 then
    Exit;
  try
    EnumProc := GetProcAddress(psapilib, 'EnumProcessModules');
    if not Assigned(EnumProc) then
      Exit;
    FillChar(ma, SizeOF(TModuleArray), 0);
    if EnumModType(EnumProc)(GetCurrentProcess, ma, 400, I) then
     begin
      for I := 0 to 400 do
        if ma[i] <> 0 then
        begin
          FillChar(FileName, MAX_PATH, 0);
          GetModuleFileName(ma[i], FileName, MAX_PATH);
          if CompareText(ExtractFileExt(FileName), '.dll') = 0 then
          begin
            S := FileName;
            S := ExtractFileName(S);
            //MessageBox(0, pchar(s), '', 0);
            patch(S);
          end;
        end;
    end;
    Result := True;
  finally
    FreeLibrary(psapilib);
  end;
end;

//////////////////////////////////////////////////////////
begin
  //MessageBox(0, 'Start', 'md_changer', MB_ICONINFORMATION);
  GetLoadedDLLList();

  g_sz:=TStringList.Create;

  GetIniFile;



end.
