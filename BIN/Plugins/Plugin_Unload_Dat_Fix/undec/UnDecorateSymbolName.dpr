{$APPTYPE CONSOLE}
program x;
uses ImageHlp, Windows, SysUtils;
var szUndName: array[1..50000] of char;
var f1,f2: TextFile;
var s: String;
begin

  AssignFile(f1, 'basic.txt');
  AssignFile(f2, 'basic.und');
  Reset(f1);
  Rewrite(f2);
  Repeat
    Readln(f1, s);
    s:=trim(s);
    if s='' then Continue;



    if UnDecorateSymbolName(
    pchar(s),
    @szUndName,
    sizeof(szUndName),
    ///UNDNAME_NO_MS_KEYWORDS or UNDNAME_NO_ACCESS_SPECIFIERS  or UNDNAME_NO_MEMBER_TYPE or UNDNAME_NO_MEMBER_TYPE
    UNDNAME_COMPLETE)=0 then
      Raise Exception.Create('Ошибка');

    s:=Trim(String(pchar(@szUndName)));

    Writeln(f2, s);
  Until eof(f1);

  CloseFile(f1);
  CloseFile(f2);



   //MessageBox(0, @szUndName, '*', 0);

end.

//public: int __thiscall CBLModule::Compile(void)

