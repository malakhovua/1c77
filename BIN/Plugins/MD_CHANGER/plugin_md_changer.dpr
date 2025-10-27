//(c) romix, 2006

Library md_changer;
uses Windows, SysUtils, PatchMemory, uBalloon;

var g_PathToMD: String; //Путь к файлу MD
var g_NewConfig: String; //Путь к новой конфигурации относительно папки md_changer
var g_RestrictLoad: String; //Запрет на загрузку (строка сообщения)
var g_TempPath: String; //Путь к временной директории


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
procedure GetIniFile(dir: String);
var f: TextFile;
var fname: String;
var s: String;
var p: Integer;
var name, value: String;
begin
  fname:=dir+'MD_CHANGER\md_changer.ini';

  //MessageBox(0, pchar(fname), 'ini', 0);

  g_NewConfig:='';
  g_RestrictLoad:='';

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

    if name='НоваяКонфигурация' then begin
      g_NewConfig:=trim(Value);
    end else if name='ЗапретЗагрузки' then begin
     g_RestrictLoad:=trim(Value);
    end;

  Until Eof(f);
  CloseFile(f);

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
var path, fname: String;
begin

  s:=lpFileName;
  path:=ExtractFilePath(s);

  fname:=LowerCase(ExtractFileName(s));

  if fname='1cv7.md' then begin
    GetIniFile(path);
    if g_RestrictLoad<>'' then begin
       //MessageBox(0, pchar(g_RestrictLoad), 'База заблокирована', MB_ICONHAND);
       ShowBalloon(pchar(g_RestrictLoad), 'База заблокирована', 3000);
       //GetIniFile(path);
       Windows.TerminateProcess(Windows.GetCurrentProcess(),1);
    end;


    //MessageBox(0, pchar(s), ' CreateFile md_changer', MB_ICONINFORMATION);

    if g_NewConfig<>'' then begin
      s:=path+'MD_CHANGER\'+g_NewConfig+'\'+fname;
      //MessageBox(0, pchar(s), 'Замена', MB_ICONINFORMATION);
    end;
  end; //if



    Result:=g_CreateFile_orig( //вызываем оригинальную CreateFile
        pchar(s),
        dwDesiredAccess,
        dwShareMode,
        LPSECURITY_ATTRIBUTES,
        dwCreationDisposition,
        dwFlagsAndAttributes,
        hTemplateFile
    );


end;

///////////////////////////////////////////////////////////////
function GetTempPathStr: String;
var
  Buffer: array[0..1023] of Char;
begin
  SetString(Result, Buffer, GetTempPath(Sizeof(Buffer)-1,Buffer));
end;




//Указатель на оригинальную функцию
type TCopyFile = function (
  lpExistingFileName: pchar;
  lpNewFileName: pchar;
  bFailIfExists: boolean
):Boolean; StdCall;

var g_CopyFile_orig: TCopyFile;


///////////////////////////////////////////////////////////////
//Функция-перехватчик CopyFile
function NewCopyFile(
  lpExistingFileName: pchar;
  lpNewFileName: pchar;
  bFailIfExists: boolean
):Boolean; StdCall;

var temp: String;
var path1: String;
var path2: String;
var mdname: String;


begin
  //      MessageBox(0,pchar('Copy From: '+lpExistingFileName+#13#10+' to '+lpNewFileName+#13#10+
  //  ' FailIfEx: '+BoolToStr(bFailIfExists)), pchar(GetTempPathStr), 0);


    mdname:=ExtractFileName(''+lpExistingFileName);


    if lowercase(mdname)='1cv7.md' then begin
      //Запуск конфигуратора отличается тем, что он сначала копирует 1cv7.md во временую папку

      path1:=ExtractFilePath(''+lpExistingFileName);

      temp:=g_TempPath;
      temp:=ExcludeTrailingPathDelimiter(temp);
      temp:=LowerCase(temp);


      path2:=ExtractFilePath(lpNewFileName);
      path2:=ExcludeTrailingPathDelimiter(path2);
      path2:=LowerCase(path2);





      if temp=path2 then begin
         //MessageBox(0,pchar('!!! INI'), pchar(g_TempPath), 0);
         
         GetIniFile(path1);
         if g_NewConfig<>'' then begin

//          While g_NewConfig<>'' do begin

        ShowBalloon(pchar('Обнаружена замена конфигурации: '+path1+'MD_CHANGER\'+g_NewConfig+'\1cv7.md'+#13#10+
           'Запуск конфигуратора заблокирован. Пожалуйста, отключите замену конфигурации в файле md_changer.ini.')
        , 'md_changer', 10000);
           Windows.TerminateProcess(Windows.GetCurrentProcess(),1);

//             GetIniFile(path1);
//          end;
         end;

      end;

      sleep(1);
    end;


    Result:=g_CopyFile_orig( //вызываем оригинальную CopyFileA
        lpExistingFileName,
        lpNewFileName,
        bFailIfExists
    );
end;







///////////////////////////////////////////////////////////////
procedure patch();
  var  pm: TPatchMemory;
begin

     pm:=TPatchMemory.Create;

     //--------------------------------------
     //Перехват системной функции CreateFile

     pm.DllNameToPatch:='frame.dll';
     pm.DllNameToFind:='KERNEL32.dll';
     pm.FuncNameToFind:='CreateFileA';
     pm.NewFunctionAddr:=@NewCreateFile;
     try
         if pm.Patch then begin
           @g_CreateFile_orig:=pm.OldFunctionAddr;
           //MessageBox(0, pchar(DllName+' CreateFileA '+IntToHex(Integer(pm.OldFunctionAddr),8)), 'md_changer',0);
         end;
     except
     end;

     //Перехват системной функции CopyFile
     //--------------------------------------
     pm.DllNameToPatch:='bkend.dll';
     pm.DllNameToFind:='KERNEL32.dll';
     pm.FuncNameToFind:='CopyFileA';
     pm.NewFunctionAddr:=@NewCopyFile;

         try
           if pm.Patch then
           @g_CopyFile_orig:=pm.OldFunctionAddr;
           //MessageBox(0, pchar(DllName +' CopyFileA '+IntToHex(Integer(pm.OldFunctionAddr),8)), 'md_changer',0);

         except
         end;

     //--------------------------------------
     pm.Free;
end;

//////////////////////////////////////////////////////////
begin
//  MessageBox(0, pchar(ParamStr(1)), 'md_changer', MB_ICONINFORMATION);
  patch();
  g_PathToMD:='';
  g_TempPath:=GetTempPathStr();


end.
