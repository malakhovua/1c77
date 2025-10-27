{$APPTYPE CONSOLE}
program GenerateProxyDll;
uses SysUtils, Windows;

var f2, f3: TextFile;
var s, path: String;

const DllName = 'odbc32.dll'; 

begin

  path:=ExtractFilePath(ParamStr(0)); //ѕуть - там же, где и текущий exe

  Assign(f2, path+'odbc33.txt'); //¬ходной файл с именами функций DLL
  Assign(f3, path+'odbc33.dpr'); //»сходный код прокси-DLL, который будет получен на выходе
  Reset(f2);
  Rewrite(f3);

  Writeln(f3, 'Library odbc33;');
  Repeat;
    Readln(f2, s);
    s:=trim(s);
    Writeln(f3, '//////////////////////////////////////////////////////////');
    Writeln(f3, 'procedure '+s+'_orig; external ' + '''' + DllName + '''' + ' name ' + '''' + s + '''' + ';');
    Writeln(f3, '  procedure '+s+'; export; stdcall;');
    Writeln(f3, '  begin ');
    Writeln(f3, '   asm ');
    Writeln(f3, '     jmp '+s+'_orig; ');
    Writeln(f3, '   end; ');
    Writeln(f3, '  end;');
    Writeln(f3, '  exports '+s+';');
  Until eof(f2);

  CloseFile(f2);

  Writeln(f3, '');
  Writeln(f3, 'begin');
  Writeln(f3, 'end.');
  CloseFile(f3);

end.
