//(c) romix, 2006

Library plugin_mxl_doctor;
//������ ��������� ������� � ����������� � Excel (25 ����� 7.7)
//����� ������ ����� ������� HTML � ��������� CSS ��� ���������� MXL->HTML

uses Windows, SysUtils, HookMemory;

const c_AddrToPatch = $250298A1;


var g_UseCSS: Integer; //������������ CSS
var g_PatchError: Integer; //������� ������ ���������� � Excel 25 ������ (������� �� ������� ������)
hm: THookMemory;


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
var f: TextFile;
var fname: String;
var s: String;
var p: Integer;
var name, value: String;
begin
  fname:=ChangeFileExt(GetDllPath(), '.ini');
  //MessageBox(0, pchar(fname), 'ini', 0);

  g_UseCSS:=0;
  g_PatchError:=0;

  if FileExists(fname)=False then Exit;
  //MessageBox(0, pchar(fname), '���� ����������', 0);

  AssignFile(f, fname);
  Reset(f);
  Repeat
    Readln(f, s);
    p:=pos(';', s);
    if p>0 then s:=Copy(s,1,p-1); //�������� ����������� ����� ;
    p:=pos('=', s);
    if p=0 then Continue;
    name:=Copy(s,1, p-1); //��� � ���� ���=��������
    name:=trim(name);
    value:=Copy(s, p+1, Length(s)-p+1); //��������
    value:=trim(value);

    //MessageBox(0, pchar(name), pchar(value), 0);

    if name='������������CSS' then begin
      g_UseCSS:=StrToInt(Value);
    end;
    if name='���������������������������' then begin
      g_PatchError:=StrToInt(Value);
    end;

  Until Eof(f);
  CloseFile(f);

end;

///////////////////////////////////////////////////////////////////////
//������� ��������� �� ������
function DelString(var s: String; const cut: String): Boolean;
var p: Integer;
begin
  Result:=False;
  p:=pos(cut, s);
  while p>0 do begin
    Delete(s, p, length(cut));
    p:=pos(cut, s);
    Result:=True;
  end;
end;

///////////////////////////////////////////////////////////////////////
procedure InsertCss(fname: String);
var f1, f2, css: TextFile;
var s: String;
var css_name: String;
begin
  if g_UseCSS=0 then exit;
 {$I+}

  try

  css_name:=ChangeFileExt(GetDllPath(), '.css');

  AssignFile(f1, fname);
  AssignFile(f2, fname+'.tmp');
  AssignFile(css, css_name);
  Reset(f1);
  Reset(css);
  Rewrite(f2);
  Readln(f1, s);
  Writeln(f2, s);

  Readln(f1, s);
  Writeln(f2, s);

  Repeat
    Readln(css, s);
    Writeln(f2, s);
  Until eof(css);

  Repeat
    Readln(f1, s);
    DelString(s, ' BORDERCOLOR=#ffffff');
    DelString(s, ' ALIGN=LEFT');
    DelString(s, ' ALIGN=RIGHT');
    DelString(s, ' VALIGN=CENTER');
    DelString(s, ' VALIGN=TOP');
    DelString(s, ' VALIGN=BOTTOM');

    DelString(s, ' SIZE=2');
    if DelString(s, '<FONT>') then
     DelString(s, '</FONT>');
//

    Writeln(f2, s);
  Until eof(f1);
  CloseFile(f1);
  CloseFile(f2);
  CloseFile(css);

  SysUtils.DeleteFile(fname);
  SysUtils.RenameFile(fname+'.tmp', fname+'.xls');

  except
     {$I-}
        CloseFile(f1);
        CloseFile(f2);
        CloseFile(css);
     {$I+}
  end;

end;


///////////////////////////////////////////////////////////////////////
//�������-����������� CSheetDoc::SaveAs (���������� � MXL)
//////////////////////////////////////////////////////////
function CSheetDoc_SaveAs_orig(fname: pchar; CSheetSaveAsType: DWORD): DWORD; stdcall;
  external 'Moxel.dll' name '?SaveAs@CSheetDoc@@QAEHPBDW4CSheetSaveAsType@@@Z';


///////////////////////////////////////////////////////////////
function NewCSheetDoc_SaveAs(fname: pchar; CSheetSaveAsType: DWORD): DWORD; stdcall;
//����������� ������� "��������� ���"
  begin
  //��������� ���� ���������� JMP
      asm
        push eax
        push ebx
        push ecx
        push edx
      end;

      if g_PatchError=1 then begin
        asm
          push eax
          mov ax, $E2EB;
          mov [WORD PTR c_AddrToPatch], ax
          pop eax
        end;
      end;


//     MessageBoxA(0, fname, pchar(IntToStr(CSheetSaveAsType)), 0);
     hm.UnHook;

      asm
        pop edx
        pop ecx
        pop ebx
        pop eax
      end;

     Result:=CSheetDoc_SaveAs_orig(fname, CSheetSaveAsType);
      asm
        push eax
        push ebx
        push ecx
        push edx
      end;

     hm.ReHook;

   //��������������� ���������� � �� �������� ���������
      if g_PatchError=1 then begin
        asm
          mov ax, $D2EB;
          mov [WORD PTR c_AddrToPatch], ax
        end;
      end;

      if CSheetSaveAsType=2 then begin
         if g_UseCSS=1   then begin
            //MessageBox(0, pchar(g_FileCSS), '*HTML', 0);
            InsertCss(fname);
         end;
      end;


      asm
        pop edx
        pop ecx
        pop ebx
        pop eax
      end;


  end;



procedure patch;
var w: Word;
var old: DWORD;
begin

          if g_PatchError=1 then begin
          asm
            mov ax, [WORD PTR c_AddrToPatch]
            mov w, ax
          end;
          if w<>$D2EB then begin
            MessageBox(0, '���������� plugin_mxl_doctor ���������� ������ �� 25 ����� 1�:����������� 7.7', 'plugin_mxl_doctor', MB_ICONERROR);
            exit;
          end;

          //������������� ���������� �� ������ ��� ����������, ������� ����� �������
          Windows.VirtualProtect(//������� WinAPI
            Pointer(c_AddrToPatch), //�����
            4, //����� ���� (����� ��������� 1 ��� 2 �������� ������ �������� 4�)
            PAGE_EXECUTE_READWRITE, //�������� �������
            @old); //���� ����� ���������� ������ ��������


          end;

          //������������� ����������� ������� "��������� ���"
           hm.DllName:='Moxel.dll';
           hm.FuncName:='?SaveAs@CSheetDoc@@QAEHPBDW4CSheetSaveAsType@@@Z';
           hm.NewFunctionAddr:=@NewCSheetDoc_SaveAs;

           hm.Hook;




end;



var g_Thread: THandle;
var g_ThreadID: DWORD;

procedure ThreadFunc(P: Pointer); stdcall;
begin
     repeat
       //MessageBeep(MB_OK);
         if Windows.GetModuleHandle('Moxel.dll')<>0 then begin
           patch;
           ExitThread(0);
         end;

       sleep(1000); //����� � �������������, ����� �� ��������� ���������
     until False;
end;


//////////////////////////////////////////////////////////
begin
  GetIniFile();

  hm:=THookMemory.Create;

  g_Thread := CreateThread(nil, 0, @ThreadFunc,
    nil, 0, g_ThreadID);
  if g_Thread = 0 then MessageBox(0, '������ ������������� ������', 'TestThread', MB_ICONINFORMATION);



end.
