//(c) romix, 2006

Library log_ert;
//Плагин ведет лог обращений к файлам ert


uses Windows, SysUtils, PatchMemory;

var g_PathToMD: String; //Каталог MD


///////////////////////////////////////////////////////////////
Function GetDllPath: String;
var
  TheFileName : array[0..MAX_PATH] of char;
begin
  FillChar(TheFileName, sizeof(TheFileName), #0);
  GetModuleFileName(hInstance, TheFileName, sizeof(TheFileName));
  Result:= trim(''+TheFileName);
end;


//////////////////////////////////////////////////////////
function Lead0(n: Integer): String;
begin
  Result:=IntToStr(n);
  if n<10 then Result:='0'+Result;
end;

//////////////////////////////////////////////////////////
function GetDateStr(dt: TDateTime): String;
var Year, Month, Day: Word;
begin
   DecodeDate(dt, Year, Month, Day);
   Result:=Lead0(Day)+'.'+Lead0(Month)+'.'+IntToStr(Year);
end;

//////////////////////////////////////////////////////////
function GetTimeStr(dt: TDateTime): String;
var Hour, Min, Sec, MSec: Word;
begin
   DecodeTime(dt, Hour, Min, Sec, MSec);
   Result:=Lead0(Hour)+':'+Lead0(Min)+':'+Lead0(Sec);
end;


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

var g_CreateFile_orig: TCreateFile;

///////////////////////////////////////////////////////////////
Function ReadComputerName:string;
 var
 i:DWORD;
 p:PChar;
begin
 i:=255;
 GetMem(p, i);
 GetComputerName(p, i);
 Result:=String(p);
 FreeMem(p);
end;

///////////////////////////////////////////////////////////////
Function GetUserFromWindows: string;
Var
  UserName    : string;
  UserNameLen : Dword;
Begin
  UserNameLen := 255;
  SetLength(userName, UserNameLen);
  If GetUserName(PChar(UserName), UserNameLen) Then
    Result := Copy(UserName,1,UserNameLen - 1)
  Else
    Result := 'Unknown';
End;



///////////////////////////////////////////////////////////////
procedure LogErtFile(ert_name: String);
//Создает сигнальный файл
var f: TextFile;
var fname: String;
var i: Integer;
var s: String;
var LogDir: String;

begin
  //MessageBox(0, pchar(ert_name), '***111', 0 );
  if g_PathToMD='' then exit;

  LogDir:=g_PathToMD+'LOG_ERT\';
  //MessageBox(0, pchar(LogDir), 'LogDir', 0 );

  try
    if not DirectoryExists(LogDir) then begin
     ForceDirectories(LogDir);
    end;
  except
    exit;
  end;

  fname:=IncludeTrailingPathDelimiter(trim(LogDir))+
  trim(ExtractFileName(ert_name))+'.log';



  AssignFile(f, fname);
  for i:=1 to 100 do begin
    {$I+}
    try
      if FileExists(fname) then begin
        Append(f);
      end else begin
        Rewrite(f);
      end;
    except
      sleep(100);
      continue;
    end;
    break;
  end;

  s:=trim(DateTimeToStr(Now)+'; user='+GetUserFromWindows+'; comp='+ReadComputerName);

  //Дополняем строку пробелами справа, чтобы длина составила 100 символов
  while length(s)<98 do begin
    s:=s+' ';
  end;

  try
    Writeln(f, s);
  except
  end;


  {$I-}
  CloseFile(f);
  {$I+}

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
var fname: String;
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
    //MessageBox(0, pchar(s), 'g_PathToMD', 0);
  end;


//E:\MD_CHANGER\Hook_1C\Plugins\v77\SYSLOG\1cv7.mlg



  s:=lpFileName;
  fname:=UpperCase(ExtractFileName(s));
  if pos('.ERT', fname)>0 then begin
    LogErtFile(s);
  end;

  asm
    popa
  end;
  
//  MessageBox(0, pchar(fname), 'CreateFile - sleep_dbf',0);
  Result:=g_CreateFile_orig( //вызываем оригинальную CreateFile
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
procedure patch(DllName:String);
  var  pm: TPatchMemory;
  var s: String;
begin

//MessageBox(0,pchar(''+DllName),'sleep_dbf',0);

     s:=LowerCase(''+DllName);
     //Плагины не патчим
     if pos('plugin_', s)>0 then Exit;

     pm:=TPatchMemory.Create;

     //--------------------------------------
     //Перехват системной функции CreateFile для указанной DLL

     pm.DllNameToPatch:=DllName;
     pm.DllNameToFind:='KERNEL32.dll';
     pm.FuncNameToFind:='CreateFileA';
     pm.NewFunctionAddr:=@NewCreateFile;
     try
       if pm.Patch then begin
         @g_CreateFile_orig:=pm.OldFunctionAddr;
       end;
       //if LowerCase(DllName)='dbeng32.dll' then
       //MessageBox(0,pchar(IntToStr(Integer(pm.OldFunctionAddr))),pchar(DllName), 0);

     except
       //on E:Exception do
       //MessageBox(0,pchar(E.Message),pchar(DllName), 0);
     end;



     //--------------------------------------
     pm.Free;
end;

//////////////////////////////////////////////////////////
begin
  //MessageBox(0, 'Start', 'sleep_dbf', MB_ICONINFORMATION);
  //GetLoadedDLLList();
  //patch('seven.dll');
  //patch('Basic.dll');
  //patch('BkEnd.dll');
  //patch('BLang.dll');
  //patch('br32.dll');

  patch('frame.dll');
  patch('seven.dll');

end.
