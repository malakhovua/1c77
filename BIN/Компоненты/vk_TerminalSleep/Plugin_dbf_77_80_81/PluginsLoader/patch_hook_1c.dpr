//(c) romix, 2007

{$APPTYPE CONSOLE}
program patch_Hook_1C;
uses SysUtils, Windows;


//const FNAME='seven.dll';
const STRING1='SHELL32.dll';
const STRING2='Hook_1C.dll';
const LEN1=Length(STRING1);
const LEN2=Length(STRING2);

Const BufSize = 30*1024*1024; //30 �����
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

  if LEN1<>LEN2 then Raise Exception.Create('����� ����� STRING1 � STRING2 �����������. �������� �� �����������!');

  Buffer:=nil;
  ok:=True;

  try

  fname:='seven.dll';
  if not FileExists(fname) then begin
    fname:='frntend.dll';
  end;

  if not FileExists(fname) then Raise Exception.Create('���� �� ������: seven.dll ��� frntend.dll.'#13#10+
  '���������� ��� ��������� � ����� BIN 1�:����������� � ��������� �� ������');


  if MessageBox(0, pchar('�� ������ ���������� ����: '+FNAME+'?'), '(c) romix, 2007', MB_OKCANCEL+MB_ICONQUESTION)<>IDOK  then Exit;





     try
       AssignFile(infile, fname);
       System.Reset(infile, 1);
     except
       Raise Exception.Create('������ �������� �����: '+fname);
     end;


     New(Buffer);

     BlockRead(infile, Buffer^, BufSize, Size);

     p:=pos(STRING2, Buffer^);
     if p>0 then begin
       Raise Exception.Create('���� ��� ���������: '+fname);
     end;


     p:=pos(STRING1, Buffer^);
     if p=0 then begin
       Raise Exception.Create('��������� �� �������: '+STRING1);
     end;

     try
       AssignFile(outfile, FNAME+'.new');
       System.Rewrite(outfile, 1);
     except
       Raise Exception.Create('������ �������� �����: '+fname+'.new');
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
         MessageBox(0, pchar('������ �������������� �����: '+FNAME), '', MB_ICONERROR);
       end;
     end;
   end;

   if ok then
      MessageBox(0, pchar('���� ������� ���������: '+FNAME), '', MB_OK);



end.
