{$APPTYPE CONSOLE}
program patch_bkend;
uses SysUtils, Windows;

const FNAME='BkEnd.dll';
const STRING1='ODBC32.DLL';
const STRING2='ODBC33.DLL';
const LEN1=Length(STRING1);

Const BufSize = 10*1024*1024; //10 мег
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
  Attributes: Word;


begin
{$I+}
  Buffer:=nil;
  ok:=True;
  if MessageBox(0, 'Вы хотите пропатчить файл: '+FNAME+'?', '(c) romix, 2006', MB_OKCANCEL+MB_ICONQUESTION)<>IDOK  then Exit;


  try

     if not FileExists(FNAME) then Raise Exception.Create('Файл не найден: '+FNAME);

     Attributes := FileGetAttr(FNAME);

     if (Attributes and faReadOnly)=faReadOnly then
       Raise Exception.Create('Файл имеет атрибут "только чтение": '+FNAME);

     if (Attributes and faSysFile)=faSysFile then
       Raise Exception.Create('Файл имеет атрибут "системный": '+FNAME);

     if (Attributes and faHidden)=faHidden then
       Raise Exception.Create('Файл имеет атрибут "скрытый": '+FNAME);

     try
       AssignFile(infile, FNAME);
       System.Reset(infile, 1);
     except
       on e:Exception do
         Raise Exception.Create('Ошибка открытия файла: '+FNAME+': '+e.Message);
     end;

     try
       AssignFile(outfile, FNAME+'.new');
       System.Rewrite(outfile, 1);
     except
       on e:Exception do
         Raise Exception.Create('Ошибка создания файла: '+FNAME+': '+e.Message);
     end;

     New(Buffer);

     BlockRead(infile, Buffer^, BufSize, Size);

     p:=pos(STRING2, Buffer^);
     if p>0 then begin
       Raise Exception.Create('Файл уже пропатчен: '+FNAME);
     end;


     p:=pos(STRING1, Buffer^);
     if p=0 then begin
       Raise Exception.Create('Подстрока не найдена: '+STRING1);
     end;

     BlockWrite(outfile, Buffer^, Size);

     if p>0 then begin
        str:=STRING2;
        seek(outfile, p-1);
        BlockWrite(outfile, str, LEN1);
     end;
  Except
     on E:Exception do begin
       MessageBox(0, pchar(E.Message), '', MB_ICONERROR);
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
       if FileExists(FNAME+'.bak') then DeleteFile(FNAME+'.bak');
       RenameFile(FNAME, FNAME+'.bak');
       RenameFile(FNAME+'.new', FNAME);
     except
       on E:Exception do begin
           MessageBox(0, pchar('Ошибка переименования файла: '+FNAME+' '+e.Message), '', MB_ICONERROR);
           ok:=False;
       end;
     end;
   end;

   if ok then
      MessageBox(0, 'Файл успешно пропатчен: '+FNAME, '', MB_OK);



end.
