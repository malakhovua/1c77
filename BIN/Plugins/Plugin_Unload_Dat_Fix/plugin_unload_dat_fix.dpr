{$OPTIMIZATION OFF}

Library log_ert;
//Плагин выводит окно запроса перед открытием и перед закрытием файла 1Cv77.dat.
//Цель - лечение проблемы с большими выгрузками.
// ----------------------------------------------------------------------------
// 28.10.2011 - портирован под Win7/ Причина несовместимости - MinWin (http://ru.wikipedia.org/wiki/MinWin)
// avgreen@molvest.org.ua

uses Windows, SysUtils, PatchMemory, HookMemory, Dialogs, uBalloon;

var g_HandleDat: Cardinal; //Handle файла DAT
var g_NameDat: String; //Имя файла DAT
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
//Указатель на оригинальную функцию
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
//Указатель на оригинальную функцию
type TCloseHandle = function (
  hObject: THandle
):BOOL; StdCall;


///////////////////////////////////////////////////////////////
//Указатель на оригинальную функцию
type TSetFilePointer = function (
    hFile: THandle; // handle of file
    lDistanceToMove: Longint;   // number of bytes to move file pointer
    lpDistanceToMoveHigh: Pointer;  // address of high-order word of distance to move
    dwMoveMethod: DWORD     // how to move
):DWORD; StdCall;


///////////////////////////////////////////////////////////////
//Указатель на оригинальную функцию
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
 Dlg.Filter := 'DAT файлы (*.dat)|*.dat|All files (*.*)|*.*';
 Dlg.InitialDir:=ExtractFilePath(fname);
 Dlg.DefaultExt:='DAT';
 Dlg.FileName:='romix.dat';
 Dlg.Title:='Укажите файл, откуда взять выгрузку 1cv77.dat';
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
//Функция-перехватчик CreateFile
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
       g_WriteMode:=True; //файл открыт на запись
     end else begin
       g_WriteMode:=False;
     end;

    if not g_WriteMode then begin
      newfn:=OpenDatFile(newfn);
    end else begin
       ShowBalloon('При выгрузке будет отключено архивирование файла 1Cv77.dat (в архив ZIP попадет пустой файл DAT,'#13#10+
       ' а сам 1Cv77.dat будет лежать в каталоге информационной базы под именем romix.dat).', 'unload_dat_fix');
    end;
    g_Flag:=True;
  end;
  asm
    popa
  end;
  Result:=g_CreateFile_orig( //вызываем оригинальную CreateFile
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
//Пишет файл с именем g_NameDat в файл с подчерком; создает вместо него пустой файл длиной 1 байт.
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
  ShowBalloon('Выгрузка DAT находится в файле: '#13#10+newname+#13#10+
  'В архиве ZIP находится пустой файл DAT', 'unload_dat_fix');

(*
  MessageBox(0, pchar('Выгрузка DAT находится в файле: '#13#10+newname+#13#10+
  'В архиве ZIP находится пустой файл DAT'), '(c)romix', MB_ICONINFORMATION);
*)

end;

//function CloseHandle(hObject: THandle): BOOL; stdcall;

///////////////////////////////////////////////////////////////
//Функция-перехватчик CloseHandle
function NewCloseHandle(
hObject: THandle):BOOL; StdCall;
begin


  Result:=g_CloseHandle_orig( //вызываем оригинальную CreateFile
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

        //if MessageBox(0, pchar('Отключить архивирование файла DAT? (в архив ZIP попадет пустой файл DAT).'), '(c)romix', MB_YESNO+MB_ICONQUESTION)=IDYES   then begin
          ReplaceDat();
        //end;
        g_NameDat:='';
        g_IsOpenDat:=False;
      end else begin
        //В случае закрытия файла при чтении
        //удаляем пиктограмму в трее где отображаются сообщения
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
//Функция-перехватчик ?NextChar@CDB7Stream@@UAEXXZ
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

  if handle=g_HandleDat then begin //Если это файл выгрузки, то считываем его с кешированием (иначе медленно)

    if g_pbuf=0 then begin
      //Читаем файл в буфер
      sleep(1);
      Windows.ReadFile(handle, g_buf, 65536, NumberOfBytesRead, nil);
      g_buf[NumberOfBytesRead+1]:=#0; //признаком конца файла является 0
    end;

    res:=g_buf[g_pbuf];

    asm
      mov eax, [g_pbuf]
      inc eax
      and eax, $FFFF //счетчик обнуляется при достижении 65536
      mov [g_pbuf], eax
    end;

    //if res=#0 then MessageBox(0, 'Достигнут конец файла', '*', 0);
  end else begin  //Иначе считываем другой файл (побайтово, без кеширования)
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
    mov [ecx+$38], al  //возвращаем считанный символ по этому адресу
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
//Форматирует число триадами
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
//Возвращает, сколько процентов n1 составляет от n2
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
    g_Wnd:=FindWindow(nil,'Загрузка данных');
    g_Wnd1:=GetWindow(g_wnd, GW_OWNER);
    //MessageBox(0,pchar(IntToHex(g_wnd,8)),pchar('окно'), MB_ICONINFORMATION);
  end;


  ShowBalloon(
  ''+IntToStr3(g_Counter64 div (1024*1024))+' из '+
   IntToStr3(g_Size64 div (1024*1024))+
   ' Мб',
   'Загружено '+Percent(g_Counter64, g_Size64));

  (*
  SetWindowText(g_wnd, pchar(
  'Загружено '+IntToStr3(g_Counter64 div 1024)+'  из  '+
   IntToStr3(g_Size64 div 1024)+
   ' килобайт  ('+Percent(g_Counter64, g_Size64)+')'
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
     //Перехват системной функции CreateFile для указанной DLL

     pm.DllNameToPatch:=DllName;
//     pm.DllNameToFind:='KERNEL32.dll';   // avgreen@molvest.org.ua
     if WinMin then DllNameToFind := 'API-MS-Win-Core-File' else DllNameToFind := 'KERNEL32.dll'; // Полное имя библиотеки 'API-MS-Win-Core-File-L1-1-0.dll' но поскольку цифры обозначают версию и как я понял могут менятся - ищем по подстроке
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
     //Перехват системной функции CloseHandle для указанной DLL


     pm.DllNameToPatch:=DllName;
//     pm.DllNameToFind:='KERNEL32.dll';   // avgreen@molvest.org.ua
     if WinMin then DllNameToFind := 'API-MS-Win-Core-Handle' else DllNameToFind := 'KERNEL32.dll'; // Полное имя библиотеки 'API-MS-Win-Core-Handle-L1-1-0.dll' но поскольку цифры обозначают версию и как я понял могут менятся - ищем по подстроке
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
     //Перехват функции ?NextChar@CDB7Stream@@UAEXXZ
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
