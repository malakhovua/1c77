unit uINI;


interface

uses SysUtils;

type T_INI = Class
  Name: String;
  Value: String;


  procedure Open(fname: String);
  procedure Close();
  function ReadLine(): Boolean;
  function GetInteger: Integer;
  function GetBoolean: Boolean;

  Constructor Create;
  Destructor Destroy; Override;

private
  g_f: TextFile;
  g_s: String;
  isOpen: Boolean;
end;



implementation
//////////////////////////////////////////////////////////////
procedure T_INI.Open(fname: String);
begin
    {$I+}
    try
      Assign(g_f, fname);
      Reset(g_f);
      isOpen:=True;
    except
      on E:Exception do begin
        {$I-}CloseFile(g_f);{$I+}
        Raise Exception.Create('Ошибка открытия файла: '+E.Message);
      end;
    end;
end;

//////////////////////////////////////////////////////////////
procedure T_INI.Close();
begin
    {$I-}CloseFile(g_f);{$I+}
    isOpen:=False;
end;

//////////////////////////////////////////////////////////////
Function T_INI.ReadLine(): Boolean;
var p:Integer;
begin
   if IsOpen=False then begin
     Raise Exception.Create('Файл ini необходимо сначала открыть.');
   end;
{$I+}
   g_s:='';
   Name:='';
   Value:='';

   Repeat
    try
      if eof(g_f) then begin
        Result:=False;
        Exit;
      end;
      Readln(g_f, g_s);
    except
      on E:Exception do begin
        {$I-}CloseFile(g_f);{$I+}
        Raise Exception.Create('Ошибка чтения файла ini:'+ E.Message);
      end;
    end;


    Result:=True;


    p:=pos(';', g_s);
    if p>0 then g_s:=Copy(g_s,1,p-1); //отрезаем комментарий после ;

    g_s:=trim(g_s);

    p:=pos('=', g_s);
    if p=0 then g_s:='';

   Until g_s<>'';


    Name:=Copy(g_s,1, p-1); //имя в паре Имя=Значение
    Name:=trim(Name);

    Value:=Copy(g_s, p+1, Length(g_s)-p+1); //значение
    Value:=trim(Value);

end;

//////////////////////////////////////////////////////////////
function T_INI.GetInteger: Integer;
begin
  try
    Result:=StrToInt(Value);
  except
    Raise Exception.Create('Ошибка чтения значения '+Name+' - ожидается целое число.');
  end;
end;

//////////////////////////////////////////////////////////////
function T_INI.GetBoolean: Boolean;
begin
    if Value='1' then Result:=True
    else if Value='Да' then Result:=True
    else if Value='Д' then Result:=True
    else if Value='Yes' then Result:=True
    else if Value='Y' then Result:=True
    else if Value='True' then Result:=True
    else if Value='Истина' then Result:=True

    else if Value='0' then Result:=False
    else if Value='Нет' then Result:=False
    else if Value='Н' then Result:=False
    else if Value='No' then Result:=False
    else if Value='N' then Result:=False
    else if Value='Ложь' then Result:=False
    else if Value='False' then Result:=False

    else
    Raise Exception.Create('Ошибка чтения значения '+Name+' - ожидается значение Да,Д,Yes,Y,1,True,Истина  либо  Нет,Н,No,N,0,False,Ложь.');
end;


//////////////////////////////////////////////////////////////
Constructor T_INI.Create;
begin
     inherited Create;
     isOpen:=False;
end;

//////////////////////////////////////////////////////////////
destructor T_INI.Destroy;
begin
     {$I-}CloseFile(g_f);{$I+}
     inherited Destroy;
end;

end.
