//(c) romix, 2007

Library plugin_terminal_sleep;
//Плагин вставляет паузу Sleep в ожидание блокировки


uses
  Windows,
  SysUtils,
//  ShellAPI,

  PatchMemory,
  uBalloon,
  uINI in 'uINI.pas';

var g_Sleep: Integer; //Максимальное значение паузы, например, 1024
var g_CurrentSleep: Integer; //Текущее значение паузы, изменяется как 1,2,4...,1024
//var g_1SJOURN: DWORD; //Дескриптор файла 1sjourn
//var g_Loop: Boolean; //Бесконечный цикл при неудачных блокировках
var g_Balloon: Boolean; //Показываем сообщения в трее
var g_SignalFile: String; //Сигнальный файл
var g_CriticalSection: _RTL_CRITICAL_SECTION;
var g_DebugOnLoad: Boolean; //выводить отладочные сообщения при загрузке



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
  g_DebugOnLoad:=False;

  ini:=T_INI.Create;
  try
    fname:=ChangeFileExt(GetDllPath(), '.ini');

    ini.Open(fname);

   while ini.ReadLine() do begin
      if ini.Name='Пауза' then begin
        g_Sleep:=ini.GetInteger;
      end else if ini.Name='Сообщения' then begin
        g_Balloon:=ini.GetBoolean;
      end else if ini.Name='СигнальныйФайл' then begin
        g_SignalFile:=ini.Value;
        dir:=ExtractFilePath(g_SignalFile);
        if DirectoryExists(dir)=False then begin
          Raise Exception.Create('Каталог, указанный в параметре СигнальныйФайл не существует: '+dir);
        end;
      end else if ini.Name='СообщенияПриЗагрузке' then begin
        g_DebugOnLoad:=ini.GetBoolean;
      end else begin
        Raise Exception.Create('Неизвестный параметр ini-файла: '+fname+' '+ini.Name+'.');
      end;

    end;
    ini.Close;

  except
    on e:Exception do begin
     ini.Close;
     MessageBox(0, pchar('Обнаружена ошибка в ini-файле: '+fname+#13#10+
     pchar(e.Message)), 'plugin_sleep_dbf', MB_ICONWARNING);
    end;
  end;

  ini.Destroy;
  LeaveCriticalSection(g_CriticalSection);

end;


///////////////////////////////////////////////////////////////
procedure CreateSignalFile();
//Создает сигнальный файл
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
//Указатель на оригинальную функцию
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
//Функция-перехватчик LockFile
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
    Result:=g_LockFile_orig( //вызываем оригинальную LockFile
      hFile,	// handle of file to lock
      dwFileOffsetLow,	// low-order word of lock region offset
      dwFileOffsetHigh,	// high-order word of lock region offset
      nNumberOfBytesToLockLow,	// low-order word of length to lock
      nNumberOfBytesToLockHigh	// high-order word of length to lock
    );

      if Result=False then begin //неудачная попытка заблокировать файл
        if g_Sleep>0 then begin

          //Делаем паузу. Возможны два варианта - с показом флажка, и без
          if (g_Balloon=True) and (g_CurrentSleep>250) then begin
            ShowBalloon('Ожидание блокировки', '1С:Предприятие', g_CurrentSleep);
          end else begin
            Windows.Sleep(g_CurrentSleep); //Делаем паузу
          end;

          //Наращиваем паузу вдвое
          if g_CurrentSleep<g_Sleep then begin
            g_CurrentSleep:=g_CurrentSleep*2;
          end;


          //Сигнальный файл для обработок проведения
          if g_CurrentSleep>=g_Sleep then begin
              CreateSignalFile();
          end;


          Result:=g_LockFile_orig( //вызываем оригинальную LockFile
            hFile,	// handle of file to lock
            dwFileOffsetLow,	// low-order word of lock region offset
            dwFileOffsetHigh,	// high-order word of lock region offset
            nNumberOfBytesToLockLow,	// low-order word of length to lock
            nNumberOfBytesToLockHigh	// high-order word of length to lock
          );


        end;
      end else begin //При удачной попытке блокировки сбрасываем паузу в 1
        g_CurrentSleep:=1;
      end;
      LeaveCriticalSection(g_CriticalSection);
end;




///////////////////////////////////////////////////////////////
procedure patch();
  var pm: TPatchMemory;
  var ok: Integer;
begin

     pm:=TPatchMemory.Create;
     //--------------------------------------
     pm.DllNameToFind:='KERNEL32.dll';
     pm.FuncNameToFind:='LockFile';


     pm.NewFunctionAddr:=@NewLockFile;

     ok:=0;


     pm.DllNameToPatch:='dbeng32.dll'; //для 1С 7.7
     if pm.Patch() then begin
      @g_LockFile_orig:=pm.OldFunctionAddr;
      ok:=1;
      if g_DebugOnLoad then
      ShowBalloon('Установлен plugin_terminal_sleep', '1С:Предприятие 7.7', 1000);
     end;

     if ok=0 then begin
       pm.DllNameToPatch:='core81.dll';  //для 1С 8.1
       if pm.Patch() then begin
        @g_LockFile_orig:=pm.OldFunctionAddr;
        ok:=1;
        if g_DebugOnLoad then
        ShowBalloon('Установлен plugin_terminal_sleep', '1С:Предприятие 8.1', 1000);
       end;
     end;

     if ok=0 then begin
       pm.DllNameToPatch:='dbeng8.dll';  //для 1С 8.0
       if pm.Patch() then begin
        @g_LockFile_orig:=pm.OldFunctionAddr;
        if g_DebugOnLoad then
        ShowBalloon('Установлен plugin_terminal_sleep', '1С:Предприятие 8.0', 1000);
        ok:=1;
       end;
     end;

     if ok=0 then begin
        ShowBalloon('Не установлен plugin_terminal_sleep', '1С:Предприятие', 1000);
     end;
     //--------------------------------------
     pm.Free;
end;







//////////////////////////////////////////////////////////
begin

  InitializeCriticalSection(g_CriticalSection);

  GetIniFile();

  patch();
  end.
