//(c) romix, 2007

{$APPTYPE CONSOLE}
program patch_Hook_1C;
uses SysUtils, Windows;


//const FNAME='seven.dll';
const STRING1='SHELL32.dll';
const STRING2='Hook_1C.dll';
const LEN1=Length(STRING1);
const LEN2=Length(STRING2);

Const BufSize = 30*1024*1024; //30 мегов
Type
  PBuffer = ^TBuffer;
  TBuffer = array [1..BufSize] of Char;
var
  Size             : integer;
  Buffer           : PBuffer;
  infile, outfile  : File;
  p: Integer;
  str: array[1..LEN1] of char;
  ok: Boolean;
  fname: String;

begin
{$I+}

  if LEN1<>LEN2 then Raise Exception.Create('Длина строк STRING1 и STRING2 различается. Сделайте их одинаковыми!');

  Buffer:=nil;
  ok:=True;

  try

  fname:='seven.dll';
  if not FileExists(fname) then begin
    fname:='frntend.dll';
  end;

  if not FileExists(fname) then Raise Exception.Create('Файл не найден: seven.dll или frntend.dll.'#13#10+
  'Скопируйте эту программу в папку BIN 1С:Предприятие и запустите ее оттуда');


  if MessageBox(0, pchar('Вы хотите пропатчить файл: '+FNAME+'?'), '(c) romix, 2007', MB_OKCANCEL+MB_ICONQUESTION)<>IDOK  then Exit;





     try
       AssignFile(infile, fname);
       System.Reset(infile, 1);
     except
       Raise Exception.Create('Ошибка открытия файла: '+fname);
     end;


     New(Buffer);

     BlockRead(infile, Buffer^, BufSize, Size);

     p:=pos(STRING2, Buffer^);
     if p>0 then begin
       Raise Exception.Create('Файл уже пропатчен: '+fname);
     end;


     p:=pos(STRING1, Buffer^);
     if p=0 then begin
       Raise Exception.Create('Подстрока не найдена: '+STRING1);
     end;

     try
       AssignFile(outfile, FNAME+'.new');
       System.Rewrite(outfile, 1);
     except
       Raise Exception.Create('Ошибка создания файла: '+fname+'.new');
     end;

     BlockWrite(outfile, Buffer^, Size);

     if p>0 then begin
        str:=STRING2;
        seek(outfile, p-1);
        BlockWrite(outfile, str, LEN1);
     end;
  Except
     on E:Exception do begin
       MessageBox(0, pchar(E.Message), 'patch_hook_1c', MB_ICONERROR);
       ok:=False;
     end;
  End;
    {$I-}
       System.close(infile);
       System.close(outfile);
       if Buffer<>nil then Dispose(Buffer);
    {$I+}

   if ok then begin
     try
       if FileExists(fname+'.bak') then DeleteFile(pchar(fname+'.bak'));
       RenameFile(FNAME, FNAME+'.bak');
       RenameFile(FNAME+'.new', FNAME);
     except
       on E:Exception do begin
         ok:=False;
         MessageBox(0, pchar('Ошибка переименования файла: '+FNAME), '', MB_ICONERROR);
       end;
     end;
   end;

   if ok then
      MessageBox(0, pchar('Файл успешно пропатчен: '+FNAME), '', MB_OK);



end.
