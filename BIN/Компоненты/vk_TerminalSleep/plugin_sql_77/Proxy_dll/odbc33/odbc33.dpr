Library odbc33; //(c) romix, 2006
uses Windows, SysUtils, uBalloon, uWaitWnd;

var g_ms: Integer;
var g_sleep: Integer;
var g_CriticalSection: _RTL_CRITICAL_SECTION;
var g_Ignore_Error_HYT00: Integer;
var g_Ignore_Error_40001: Integer;
var g_UseBalloon: Integer;
var g_UseWindow: Integer;
var g_SignalFile: String;
var g_MsgMinPauseLevel: Integer;
var g_SignalFileLevel: Integer;


///////////////////////////////////////////////////////////////
procedure GetIniFile;
var f: TextFile;
var fname: String;
var s: String;
var p: Integer;
var name, value: String;
var dir: String;
begin
  dir:=ExtractFilePath(ParamStr(0));
  fname:=dir+'odbc33.ini';

  g_ms:=1024;
  g_sleep:=1;
  g_Ignore_Error_HYT00:=0;
  g_Ignore_Error_40001:=0;
  g_UseBalloon:=0;
  g_UseWindow:=0;
  g_SignalFile:='';
  g_MsgMinPauseLevel:=128;
  g_SignalFileLevel:=1024;


  if FileExists(fname)=False then Exit;

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

    if name='Ignore_Error_HYT00' then begin
      if Value='1' then g_Ignore_Error_HYT00:=1;
    end else if name='Ignore_Error_40001' then begin
      if Value='1' then g_Ignore_Error_40001:=1;
    end else if name='UseBalloon' then begin
      if Value='1' then g_UseBalloon:=1;
    end else if name='UseWindow' then begin
      if Value='1' then g_UseWindow:=1;
    end else if name='SignalFile' then begin
      g_SignalFile:=Value;
    end else if name='Delay' then begin
      g_ms:=StrToInt(Value);
    end else if name='MsgMinPauseLevel' then begin
      g_MsgMinPauseLevel:=StrToInt(Value);
    end else if name='SignalFileLevel' then begin
      g_SignalFileLevel:=StrToInt(Value);
    end;

  Until Eof(f);
  CloseFile(f);

end;


//////////////////////////////////////////////////////////
procedure CreateSignalFile();
var f: TextFile;
begin
  if g_SignalFile='' then exit;
  {$I-}
    if FileExists(g_SignalFile) then exit;
    AssignFile(f, g_SignalFile);
    Rewrite(f);
    CloseFile(f);
  {$I+}

end;

//////////////////////////////////////////////////////////
procedure sleep_icon(ms: Integer; ic_type: pchar; msg: String);
var
  Icon: HICON;
  Wnd: HWND;
begin
 if ms<g_MsgMinPauseLevel then begin
   sleep(ms);
   exit;
 end;
 if g_SignalFile<>'' then begin
    if ms>=g_SignalFileLevel then begin
      CreateSignalFile();
    end;  
 end;

 if g_UseBalloon=1 then begin //Показываем баллон
    Wnd:=GetForegroundWindow();
    Icon := LoadIcon(0, ic_type);
    DZAddTrayIcon(Wnd, 1, Icon, 'Пауза');
    if g_UseBalloon=1 then begin
      DZBalloonTrayIcon(Wnd, 1, 10, msg, '1С:Предприятие', bitInfo);
    end;
    sleep(ms);
    DZRemoveTrayIcon(Wnd, 1);
 end else if g_UseWindow=1 then begin //Показываем окно
    CreateWaitWindow('Пауза..', PChar(msg), ms);
 end else begin
    sleep(ms); //просто пауза, без показа окон
 end;
 
end;



(*
//////////////////////////////////////////////////////////
procedure LogQuery(s: String);
var f: TextFile;
begin
   {$I+}
   AssignFile(f, g_LogSQL);

   try
   if FileExists(g_LogSQL) then begin
     Append(f);
   end else begin
     Rewrite(f);
   end;

   Write(f, TimeToStr(now),': ');
   Writeln(f, s);

   finally
     {$I-}
       CloseFile(f);
     {$I+}
   end;
end;
*)


type pint = ^Integer;
//////////////////////////////////////////////////////////
function SQLGetDiagRecA_orig(p1, p2, p3: DWORD; p4: pchar; p5: pint; p6: pchar; p7: Integer; p8: pint):DWORD; stdcall; external 'odbc32.dll' name 'SQLGetDiagRecA';
  function SQLGetDiagRecA(p1, p2, p3: DWORD; p4: pchar; p5: pint; p6: pchar; p7: Integer; p8: pint):DWORD; export; stdcall;
  begin
     Result:=SQLGetDiagRecA_orig(p1, p2, p3, p4, p5, p6, p7, p8);
  end;
  exports SQLGetDiagRecA;


//////////////////////////////////////////////////////////
//Получаем код ошибки SQL, например 40001
function GetSQLError(StatementHandle: DWORD): String;
  var ErrorMsg: Array[1..201] of char;
  var SqlState: Array[1..10] of char;
  var msglen, NativeError: Integer;

const SQL_HANDLE_ENV = 1;
const SQL_HANDLE_DBC = 2;
const SQL_HANDLE_STMT = 3;
const SQL_HANDLE_DESC = 4;

begin
  FillChar(SqlState,Sizeof(SqlState),0);
//  MessageBox(0, 'SQLGetDiagRecA_orig', '', 0);
  SQLGetDiagRecA_orig(SQL_HANDLE_STMT, StatementHandle, 1, @SqlState, @NativeError, @ErrorMsg, 200, @msglen);
  Result:=trim(String(SqlState));
//  MessageBox(0, pchar(''+ErrorMsg), pchar(String(SqlState)), 0);

  //fprintf(stderr, "Error: Message-%s Statment-%s (%d)\n", msg, statm, ret);
end;

//////////////////////////////////////////////////////////
//Эту функцию перехватываем, чтобы вставить sleep
function  SQLExecDirectA_orig(StatementHandle: DWORD; p2, p3: pchar):DWORD; stdcall; external 'odbc32.dll' name 'SQLExecDirectA';
  function SQLExecDirectA(StatementHandle: DWORD; p2, p3: pchar):DWORD; export; stdcall;
  var w: DWORD;
  var ok: Integer;
  var err: String;
  begin
    Repeat
      ok:=1;
      Result:=SQLExecDirectA_orig(StatementHandle,p2,p3);
      w:=Result and $FFFF; //выделяем младшее слово
      if w=2 then begin //SQL_STILL_EXECUTING - запрос все еще выполняется
        ////Реагируем только на запрос ожидания блокировки таблицы журналов
        //if pos('{call _1sp__1SJOURN_TLockX}', p2)>0 then begin
          if g_ms>0 then begin
             sleep_icon(g_sleep, IDI_APPLICATION, 'Ожидание блокировки SQL');

             EnterCriticalSection(g_CriticalSection);
             if g_sleep < g_ms then g_sleep:=g_sleep*2;
             LeaveCriticalSection(g_CriticalSection);

             //ok:=0; //запрос будет повторен
             //MessageBox(0, 'SQL_STILL_EXECUTING', pchar(IntToStr(g_sleep)), 0);
          end;
        //end;
      end else if w=$FFFF then begin //SQL_ERROR - ошибка при выполнении запроса

         err:=GetSQLError(StatementHandle);
         //MessageBox(0, 'Ошибка при выполнении запроса', pchar(err), 0);

         if err='40001' then begin
         //MessageBox(0, 'Ошибка 40001', '***', 0);
          //Ловим ошибку:
          //SQL State: 40001 Native: 1205 Message: [Microsoft][ODBC SQL Server Driver][SQL Server]
          //Transaction (Process ID 54) was deadlocked on lock resources with another process and has been chosen as the deadlock victim.
          //Rerun the transaction.

          if g_Ignore_Error_40001=1 then begin
            sleep_icon(2048, IDI_HAND, 'Ошибка Deadlock 40001');
            ok:=0; //запрос будет повторен
          end;
         end else if err='HYT00' then begin
           // {Обработка.МассовоеПроведение.Форма.Модуль(11)}: SQL State: HYT00 Native: 0 Message:
           //[Microsoft][ODBC SQL Server Driver]Время ожидания истекло

          if g_Ignore_Error_HYT00=1 then begin
           sleep_icon(2048, IDI_QUESTION, 'Время ожидания истекло');
           ok:=0; //запрос будет повторен
          end;
         end;
      end else begin
         EnterCriticalSection(g_CriticalSection);
         g_sleep:=1; //при удачном SQL-запросе
         LeaveCriticalSection(g_CriticalSection);

      end;
    Until ok=1;
  end;
  exports SQLExecDirectA;



//Оставшиеся функции - просто заглушки, которые делают jmp на оригинальную функцию.
//Кадр стека там не прописан, поэтому использовать их для чего-либо полезного, в
//общем случае, нельзя.


//////////////////////////////////////////////////////////
procedure SQLExecDirect_orig; external 'odbc32.dll' name 'SQLExecDirect';
  procedure SQLExecDirect; export; stdcall;
  begin
   asm
     jmp SQLExecDirect_orig;
   end;
  end;
  exports SQLExecDirect;

//////////////////////////////////////////////////////////
procedure SQLExecDirectW_orig; external 'odbc32.dll' name 'SQLExecDirectW';
  procedure SQLExecDirectW; export; stdcall;
  begin
   asm
     jmp SQLExecDirectW_orig;
   end;
  end;
  exports SQLExecDirectW;

//////////////////////////////////////////////////////////
procedure CloseODBCPerfData_orig; external 'odbc32.dll' name 'CloseODBCPerfData';
  procedure CloseODBCPerfData; export; stdcall;
  begin 
   asm 
     jmp CloseODBCPerfData_orig;
   end; 
  end;
  exports CloseODBCPerfData;
//////////////////////////////////////////////////////////
procedure CollectODBCPerfData_orig; external 'odbc32.dll' name 'CollectODBCPerfData';
  procedure CollectODBCPerfData; export; stdcall;
  begin
   asm
     jmp CollectODBCPerfData_orig; 
   end;
  end;
  exports CollectODBCPerfData;
//////////////////////////////////////////////////////////
procedure CursorLibLockDbc_orig; external 'odbc32.dll' name 'CursorLibLockDbc';
  procedure CursorLibLockDbc; export; stdcall;
  begin
   asm 
     jmp CursorLibLockDbc_orig; 
   end;
  end;
  exports CursorLibLockDbc;
//////////////////////////////////////////////////////////
procedure CursorLibLockDesc_orig; external 'odbc32.dll' name 'CursorLibLockDesc';
  procedure CursorLibLockDesc; export; stdcall;
  begin 
   asm 
     jmp CursorLibLockDesc_orig; 
   end; 
  end;
  exports CursorLibLockDesc;
//////////////////////////////////////////////////////////
procedure CursorLibLockStmt_orig; external 'odbc32.dll' name 'CursorLibLockStmt';
  procedure CursorLibLockStmt; export; stdcall;
  begin 
   asm 
     jmp CursorLibLockStmt_orig; 
   end; 
  end;
  exports CursorLibLockStmt;
//////////////////////////////////////////////////////////
procedure CursorLibTransact_orig; external 'odbc32.dll' name 'CursorLibTransact';
  procedure CursorLibTransact; export; stdcall;
  begin 
   asm 
     jmp CursorLibTransact_orig; 
   end;
  end;
  exports CursorLibTransact;
//////////////////////////////////////////////////////////
procedure GetODBCSharedData_orig; external 'odbc32.dll' name 'GetODBCSharedData';
  procedure GetODBCSharedData; export; stdcall;
  begin
   asm 
     jmp GetODBCSharedData_orig;
   end; 
  end;
  exports GetODBCSharedData;
//////////////////////////////////////////////////////////
procedure LockHandle_orig; external 'odbc32.dll' name 'LockHandle';
  procedure LockHandle; export; stdcall;
  begin
   asm 
     jmp LockHandle_orig; 
   end; 
  end;
  exports LockHandle;
//////////////////////////////////////////////////////////
procedure MpHeapAlloc_orig; external 'odbc32.dll' name 'MpHeapAlloc';
  procedure MpHeapAlloc; export; stdcall;
  begin 
   asm 
     jmp MpHeapAlloc_orig; 
   end; 
  end;
  exports MpHeapAlloc;
//////////////////////////////////////////////////////////
procedure MpHeapCompact_orig; external 'odbc32.dll' name 'MpHeapCompact';
  procedure MpHeapCompact; export; stdcall;
  begin 
   asm 
     jmp MpHeapCompact_orig; 
   end; 
  end;
  exports MpHeapCompact;
//////////////////////////////////////////////////////////
procedure MpHeapCreate_orig; external 'odbc32.dll' name 'MpHeapCreate';
  procedure MpHeapCreate; export; stdcall;
  begin 
   asm 
     jmp MpHeapCreate_orig; 
   end; 
  end;
  exports MpHeapCreate;
//////////////////////////////////////////////////////////
procedure MpHeapDestroy_orig; external 'odbc32.dll' name 'MpHeapDestroy';
  procedure MpHeapDestroy; export; stdcall;
  begin
   asm 
     jmp MpHeapDestroy_orig; 
   end;
  end;
  exports MpHeapDestroy;
//////////////////////////////////////////////////////////
procedure MpHeapFree_orig; external 'odbc32.dll' name 'MpHeapFree';
  procedure MpHeapFree; export; stdcall;
  begin
   asm 
     jmp MpHeapFree_orig; 
   end; 
  end;
  exports MpHeapFree;
//////////////////////////////////////////////////////////
procedure MpHeapReAlloc_orig; external 'odbc32.dll' name 'MpHeapReAlloc';
  procedure MpHeapReAlloc; export; stdcall;
  begin 
   asm 
     jmp MpHeapReAlloc_orig; 
   end; 
  end;
  exports MpHeapReAlloc;
//////////////////////////////////////////////////////////
procedure MpHeapSize_orig; external 'odbc32.dll' name 'MpHeapSize';
  procedure MpHeapSize; export; stdcall;
  begin 
   asm 
     jmp MpHeapSize_orig; 
   end; 
  end;
  exports MpHeapSize;
//////////////////////////////////////////////////////////
procedure MpHeapValidate_orig; external 'odbc32.dll' name 'MpHeapValidate';
  procedure MpHeapValidate; export; stdcall;
  begin 
   asm 
     jmp MpHeapValidate_orig; 
   end; 
  end;
  exports MpHeapValidate;
//////////////////////////////////////////////////////////
procedure ODBCGetTryWaitValue_orig; external 'odbc32.dll' name 'ODBCGetTryWaitValue';
  procedure ODBCGetTryWaitValue; export; stdcall;
  begin 
   asm 
     jmp ODBCGetTryWaitValue_orig; 
   end; 
  end;
  exports ODBCGetTryWaitValue;
//////////////////////////////////////////////////////////
procedure ODBCInternalConnectW_orig; external 'odbc32.dll' name 'ODBCInternalConnectW';
  procedure ODBCInternalConnectW; export; stdcall;
  begin
   asm 
     jmp ODBCInternalConnectW_orig; 
   end; 
  end;
  exports ODBCInternalConnectW;
//////////////////////////////////////////////////////////
procedure ODBCQualifyFileDSNW_orig; external 'odbc32.dll' name 'ODBCQualifyFileDSNW';
  procedure ODBCQualifyFileDSNW; export; stdcall;
  begin 
   asm 
     jmp ODBCQualifyFileDSNW_orig; 
   end; 
  end;
  exports ODBCQualifyFileDSNW;
//////////////////////////////////////////////////////////
procedure ODBCSetTryWaitValue_orig; external 'odbc32.dll' name 'ODBCSetTryWaitValue';
  procedure ODBCSetTryWaitValue; export; stdcall;
  begin 
   asm 
     jmp ODBCSetTryWaitValue_orig; 
   end; 
  end;
  exports ODBCSetTryWaitValue;
//////////////////////////////////////////////////////////
procedure OpenODBCPerfData_orig; external 'odbc32.dll' name 'OpenODBCPerfData';
  procedure OpenODBCPerfData; export; stdcall;
  begin 
   asm 
     jmp OpenODBCPerfData_orig; 
   end; 
  end;
  exports OpenODBCPerfData;
//////////////////////////////////////////////////////////
procedure PostComponentError_orig; external 'odbc32.dll' name 'PostComponentError';
  procedure PostComponentError; export; stdcall;
  begin 
   asm 
     jmp PostComponentError_orig; 
   end; 
  end;
  exports PostComponentError;
//////////////////////////////////////////////////////////
procedure PostODBCComponentError_orig; external 'odbc32.dll' name 'PostODBCComponentError';
  procedure PostODBCComponentError; export; stdcall;
  begin
   asm 
     jmp PostODBCComponentError_orig; 
   end; 
  end;
  exports PostODBCComponentError;
//////////////////////////////////////////////////////////
procedure PostODBCError_orig; external 'odbc32.dll' name 'PostODBCError';
  procedure PostODBCError; export; stdcall;
  begin 
   asm 
     jmp PostODBCError_orig; 
   end; 
  end;
  exports PostODBCError;
//////////////////////////////////////////////////////////
procedure SQLAllocConnect_orig; external 'odbc32.dll' name 'SQLAllocConnect';
  procedure SQLAllocConnect; export; stdcall;
  begin 
   asm 
     jmp SQLAllocConnect_orig; 
   end; 
  end;
  exports SQLAllocConnect;
//////////////////////////////////////////////////////////
procedure SQLAllocEnv_orig; external 'odbc32.dll' name 'SQLAllocEnv';
  procedure SQLAllocEnv; export; stdcall;
  begin 
   asm 
     jmp SQLAllocEnv_orig; 
   end; 
  end;
  exports SQLAllocEnv;
//////////////////////////////////////////////////////////
procedure SQLAllocHandle_orig; external 'odbc32.dll' name 'SQLAllocHandle';
  procedure SQLAllocHandle; export; stdcall;
  begin 
   asm 
     jmp SQLAllocHandle_orig; 
   end; 
  end;
  exports SQLAllocHandle;
//////////////////////////////////////////////////////////
procedure SQLAllocHandleStd_orig; external 'odbc32.dll' name 'SQLAllocHandleStd';
  procedure SQLAllocHandleStd; export; stdcall;
  begin
   asm 
     jmp SQLAllocHandleStd_orig; 
   end; 
  end;
  exports SQLAllocHandleStd;
//////////////////////////////////////////////////////////
procedure SQLAllocStmt_orig; external 'odbc32.dll' name 'SQLAllocStmt';
  procedure SQLAllocStmt; export; stdcall;
  begin 
   asm 
     jmp SQLAllocStmt_orig; 
   end; 
  end;
  exports SQLAllocStmt;
//////////////////////////////////////////////////////////
procedure SQLBindCol_orig; external 'odbc32.dll' name 'SQLBindCol';
  procedure SQLBindCol; export; stdcall;
  begin 
   asm 
     jmp SQLBindCol_orig; 
   end; 
  end;
  exports SQLBindCol;
//////////////////////////////////////////////////////////
procedure SQLBindParam_orig; external 'odbc32.dll' name 'SQLBindParam';
  procedure SQLBindParam; export; stdcall;
  begin 
   asm 
     jmp SQLBindParam_orig; 
   end; 
  end;
  exports SQLBindParam;
//////////////////////////////////////////////////////////
procedure SQLBindParameter_orig; external 'odbc32.dll' name 'SQLBindParameter';
  procedure SQLBindParameter; export; stdcall;
  begin 
   asm 
     jmp SQLBindParameter_orig; 
   end; 
  end;
  exports SQLBindParameter;
//////////////////////////////////////////////////////////
procedure SQLBrowseConnect_orig; external 'odbc32.dll' name 'SQLBrowseConnect';
  procedure SQLBrowseConnect; export; stdcall;
  begin
   asm 
     jmp SQLBrowseConnect_orig; 
   end; 
  end;
  exports SQLBrowseConnect;
//////////////////////////////////////////////////////////
procedure SQLBrowseConnectA_orig; external 'odbc32.dll' name 'SQLBrowseConnectA';
  procedure SQLBrowseConnectA; export; stdcall;
  begin
   asm 
     jmp SQLBrowseConnectA_orig; 
   end; 
  end;
  exports SQLBrowseConnectA;
//////////////////////////////////////////////////////////
procedure SQLBrowseConnectW_orig; external 'odbc32.dll' name 'SQLBrowseConnectW';
  procedure SQLBrowseConnectW; export; stdcall;
  begin 
   asm 
     jmp SQLBrowseConnectW_orig; 
   end; 
  end;
  exports SQLBrowseConnectW;
//////////////////////////////////////////////////////////
procedure SQLBulkOperations_orig; external 'odbc32.dll' name 'SQLBulkOperations';
  procedure SQLBulkOperations; export; stdcall;
  begin 
   asm 
     jmp SQLBulkOperations_orig; 
   end; 
  end;
  exports SQLBulkOperations;
//////////////////////////////////////////////////////////
procedure SQLCancel_orig; external 'odbc32.dll' name 'SQLCancel';
  procedure SQLCancel; export; stdcall;
  begin 
   asm 
     jmp SQLCancel_orig; 
   end; 
  end;
  exports SQLCancel;
//////////////////////////////////////////////////////////
procedure SQLCloseCursor_orig; external 'odbc32.dll' name 'SQLCloseCursor';
  procedure SQLCloseCursor; export; stdcall;
  begin
   asm 
     jmp SQLCloseCursor_orig; 
   end; 
  end;
  exports SQLCloseCursor;
//////////////////////////////////////////////////////////
procedure SQLColAttribute_orig; external 'odbc32.dll' name 'SQLColAttribute';
  procedure SQLColAttribute; export; stdcall;
  begin 
   asm
     jmp SQLColAttribute_orig; 
   end; 
  end;
  exports SQLColAttribute;
//////////////////////////////////////////////////////////
procedure SQLColAttributeA_orig; external 'odbc32.dll' name 'SQLColAttributeA';
  procedure SQLColAttributeA; export; stdcall;
  begin 
   asm 
     jmp SQLColAttributeA_orig; 
   end; 
  end;
  exports SQLColAttributeA;
//////////////////////////////////////////////////////////
procedure SQLColAttributeW_orig; external 'odbc32.dll' name 'SQLColAttributeW';
  procedure SQLColAttributeW; export; stdcall;
  begin 
   asm 
     jmp SQLColAttributeW_orig; 
   end; 
  end;
  exports SQLColAttributeW;
//////////////////////////////////////////////////////////
procedure SQLColAttributes_orig; external 'odbc32.dll' name 'SQLColAttributes';
  procedure SQLColAttributes; export; stdcall;
  begin 
   asm 
     jmp SQLColAttributes_orig; 
   end; 
  end;
  exports SQLColAttributes;
//////////////////////////////////////////////////////////
procedure SQLColAttributesA_orig; external 'odbc32.dll' name 'SQLColAttributesA';
  procedure SQLColAttributesA; export; stdcall;
  begin
   asm 
     jmp SQLColAttributesA_orig; 
   end; 
  end;
  exports SQLColAttributesA;
//////////////////////////////////////////////////////////
procedure SQLColAttributesW_orig; external 'odbc32.dll' name 'SQLColAttributesW';
  procedure SQLColAttributesW; export; stdcall;
  begin 
   asm 
     jmp SQLColAttributesW_orig;
   end; 
  end;
  exports SQLColAttributesW;
//////////////////////////////////////////////////////////
procedure SQLColumnPrivileges_orig; external 'odbc32.dll' name 'SQLColumnPrivileges';
  procedure SQLColumnPrivileges; export; stdcall;
  begin 
   asm 
     jmp SQLColumnPrivileges_orig; 
   end; 
  end;
  exports SQLColumnPrivileges;
//////////////////////////////////////////////////////////
procedure SQLColumnPrivilegesA_orig; external 'odbc32.dll' name 'SQLColumnPrivilegesA';
  procedure SQLColumnPrivilegesA; export; stdcall;
  begin 
   asm 
     jmp SQLColumnPrivilegesA_orig; 
   end; 
  end;
  exports SQLColumnPrivilegesA;
//////////////////////////////////////////////////////////
procedure SQLColumnPrivilegesW_orig; external 'odbc32.dll' name 'SQLColumnPrivilegesW';
  procedure SQLColumnPrivilegesW; export; stdcall;
  begin 
   asm 
     jmp SQLColumnPrivilegesW_orig; 
   end; 
  end;
  exports SQLColumnPrivilegesW;
//////////////////////////////////////////////////////////
procedure SQLColumns_orig; external 'odbc32.dll' name 'SQLColumns';
  procedure SQLColumns; export; stdcall;
  begin
   asm 
     jmp SQLColumns_orig; 
   end; 
  end;
  exports SQLColumns;
//////////////////////////////////////////////////////////
procedure SQLColumnsA_orig; external 'odbc32.dll' name 'SQLColumnsA';
  procedure SQLColumnsA; export; stdcall;
  begin 
   asm 
     jmp SQLColumnsA_orig; 
   end;
  end;
  exports SQLColumnsA;
//////////////////////////////////////////////////////////
procedure SQLColumnsW_orig; external 'odbc32.dll' name 'SQLColumnsW';
  procedure SQLColumnsW; export; stdcall;
  begin 
   asm 
     jmp SQLColumnsW_orig; 
   end; 
  end;
  exports SQLColumnsW;
//////////////////////////////////////////////////////////
procedure SQLConnect_orig; external 'odbc32.dll' name 'SQLConnect';
  procedure SQLConnect; export; stdcall;
  begin 
   asm 
     jmp SQLConnect_orig; 
   end; 
  end;
  exports SQLConnect;
//////////////////////////////////////////////////////////
procedure SQLConnectA_orig; external 'odbc32.dll' name 'SQLConnectA';
  procedure SQLConnectA; export; stdcall;
  begin 
   asm 
     jmp SQLConnectA_orig; 
   end; 
  end;
  exports SQLConnectA;
//////////////////////////////////////////////////////////
procedure SQLConnectW_orig; external 'odbc32.dll' name 'SQLConnectW';
  procedure SQLConnectW; export; stdcall;
  begin
   asm 
     jmp SQLConnectW_orig; 
   end; 
  end;
  exports SQLConnectW;
//////////////////////////////////////////////////////////
procedure SQLCopyDesc_orig; external 'odbc32.dll' name 'SQLCopyDesc';
  procedure SQLCopyDesc; export; stdcall;
  begin 
   asm 
     jmp SQLCopyDesc_orig; 
   end; 
  end;
  exports SQLCopyDesc;
//////////////////////////////////////////////////////////
procedure SQLDataSources_orig; external 'odbc32.dll' name 'SQLDataSources';
  procedure SQLDataSources; export; stdcall;
  begin 
   asm 
     jmp SQLDataSources_orig; 
   end; 
  end;
  exports SQLDataSources;
//////////////////////////////////////////////////////////
procedure SQLDataSourcesA_orig; external 'odbc32.dll' name 'SQLDataSourcesA';
  procedure SQLDataSourcesA; export; stdcall;
  begin 
   asm 
     jmp SQLDataSourcesA_orig; 
   end; 
  end;
  exports SQLDataSourcesA;
//////////////////////////////////////////////////////////
procedure SQLDataSourcesW_orig; external 'odbc32.dll' name 'SQLDataSourcesW';
  procedure SQLDataSourcesW; export; stdcall;
  begin 
   asm 
     jmp SQLDataSourcesW_orig; 
   end; 
  end;
  exports SQLDataSourcesW;
//////////////////////////////////////////////////////////
procedure SQLDescribeCol_orig; external 'odbc32.dll' name 'SQLDescribeCol';
  procedure SQLDescribeCol; export; stdcall;
  begin
   asm 
     jmp SQLDescribeCol_orig; 
   end; 
  end;
  exports SQLDescribeCol;
//////////////////////////////////////////////////////////
procedure SQLDescribeColA_orig; external 'odbc32.dll' name 'SQLDescribeColA';
  procedure SQLDescribeColA; export; stdcall;
  begin 
   asm 
     jmp SQLDescribeColA_orig; 
   end; 
  end;
  exports SQLDescribeColA;
//////////////////////////////////////////////////////////
procedure SQLDescribeColW_orig; external 'odbc32.dll' name 'SQLDescribeColW';
  procedure SQLDescribeColW; export; stdcall;
  begin 
   asm 
     jmp SQLDescribeColW_orig; 
   end; 
  end;
  exports SQLDescribeColW;
//////////////////////////////////////////////////////////
procedure SQLDescribeParam_orig; external 'odbc32.dll' name 'SQLDescribeParam';
  procedure SQLDescribeParam; export; stdcall;
  begin 
   asm 
     jmp SQLDescribeParam_orig; 
   end; 
  end;
  exports SQLDescribeParam;
//////////////////////////////////////////////////////////
procedure SQLDisconnect_orig; external 'odbc32.dll' name 'SQLDisconnect';
  procedure SQLDisconnect; export; stdcall;
  begin 
   asm 
     jmp SQLDisconnect_orig; 
   end; 
  end;
  exports SQLDisconnect;
//////////////////////////////////////////////////////////
procedure SQLDriverConnect_orig; external 'odbc32.dll' name 'SQLDriverConnect';
  procedure SQLDriverConnect; export; stdcall;
  begin
   asm 
     jmp SQLDriverConnect_orig; 
   end; 
  end;
  exports SQLDriverConnect;
//////////////////////////////////////////////////////////
procedure SQLDriverConnectA_orig; external 'odbc32.dll' name 'SQLDriverConnectA';
  procedure SQLDriverConnectA; export; stdcall;
  begin 
   asm 
     jmp SQLDriverConnectA_orig; 
   end; 
  end;
  exports SQLDriverConnectA;
//////////////////////////////////////////////////////////
procedure SQLDriverConnectW_orig; external 'odbc32.dll' name 'SQLDriverConnectW';
  procedure SQLDriverConnectW; export; stdcall;
  begin 
   asm 
     jmp SQLDriverConnectW_orig; 
   end; 
  end;
  exports SQLDriverConnectW;
//////////////////////////////////////////////////////////
procedure SQLDrivers_orig; external 'odbc32.dll' name 'SQLDrivers';
  procedure SQLDrivers; export; stdcall;
  begin 
   asm 
     jmp SQLDrivers_orig; 
   end; 
  end;
  exports SQLDrivers;
//////////////////////////////////////////////////////////
procedure SQLDriversA_orig; external 'odbc32.dll' name 'SQLDriversA';
  procedure SQLDriversA; export; stdcall;
  begin 
   asm 
     jmp SQLDriversA_orig; 
   end; 
  end;
  exports SQLDriversA;
//////////////////////////////////////////////////////////
procedure SQLDriversW_orig; external 'odbc32.dll' name 'SQLDriversW';
  procedure SQLDriversW; export; stdcall;
  begin
   asm 
     jmp SQLDriversW_orig; 
   end; 
  end;
  exports SQLDriversW;
//////////////////////////////////////////////////////////
procedure SQLEndTran_orig; external 'odbc32.dll' name 'SQLEndTran';
  procedure SQLEndTran; export; stdcall;
  begin 
   asm 
     jmp SQLEndTran_orig; 
   end; 
  end;
  exports SQLEndTran;
//////////////////////////////////////////////////////////
procedure SQLError_orig; external 'odbc32.dll' name 'SQLError';
  procedure SQLError; export; stdcall;
  begin 
   asm 
     jmp SQLError_orig; 
   end; 
  end;
  exports SQLError;
//////////////////////////////////////////////////////////
procedure SQLErrorA_orig; external 'odbc32.dll' name 'SQLErrorA';
  procedure SQLErrorA; export; stdcall;
  begin 
   asm 
     jmp SQLErrorA_orig; 
   end; 
  end;
  exports SQLErrorA;
//////////////////////////////////////////////////////////
procedure SQLErrorW_orig; external 'odbc32.dll' name 'SQLErrorW';
  procedure SQLErrorW; export; stdcall;
  begin 
   asm 
     jmp SQLErrorW_orig; 
   end; 
  end;
  exports SQLErrorW;
//////////////////////////////////////////////////////////
procedure SQLExecute_orig; external 'odbc32.dll' name 'SQLExecute';
  procedure SQLExecute; export; stdcall;
  begin
   asm 
     jmp SQLExecute_orig; 
   end; 
  end;
  exports SQLExecute;
//////////////////////////////////////////////////////////
procedure SQLExtendedFetch_orig; external 'odbc32.dll' name 'SQLExtendedFetch';
  procedure SQLExtendedFetch; export; stdcall;
  begin 
   asm 
     jmp SQLExtendedFetch_orig; 
   end; 
  end;
  exports SQLExtendedFetch;
//////////////////////////////////////////////////////////
procedure SQLFetch_orig; external 'odbc32.dll' name 'SQLFetch';
  procedure SQLFetch; export; stdcall;
  begin 
   asm 
     jmp SQLFetch_orig; 
   end; 
  end;
  exports SQLFetch;
//////////////////////////////////////////////////////////
procedure SQLFetchScroll_orig; external 'odbc32.dll' name 'SQLFetchScroll';
  procedure SQLFetchScroll; export; stdcall;
  begin 
   asm 
     jmp SQLFetchScroll_orig; 
   end; 
  end;
  exports SQLFetchScroll;
//////////////////////////////////////////////////////////
procedure SQLForeignKeys_orig; external 'odbc32.dll' name 'SQLForeignKeys';
  procedure SQLForeignKeys; export; stdcall;
  begin 
   asm 
     jmp SQLForeignKeys_orig; 
   end; 
  end;
  exports SQLForeignKeys;
//////////////////////////////////////////////////////////
procedure SQLForeignKeysA_orig; external 'odbc32.dll' name 'SQLForeignKeysA';
  procedure SQLForeignKeysA; export; stdcall;
  begin
   asm 
     jmp SQLForeignKeysA_orig; 
   end; 
  end;
  exports SQLForeignKeysA;
//////////////////////////////////////////////////////////
procedure SQLForeignKeysW_orig; external 'odbc32.dll' name 'SQLForeignKeysW';
  procedure SQLForeignKeysW; export; stdcall;
  begin 
   asm 
     jmp SQLForeignKeysW_orig; 
   end; 
  end;
  exports SQLForeignKeysW;
//////////////////////////////////////////////////////////
procedure SQLFreeConnect_orig; external 'odbc32.dll' name 'SQLFreeConnect';
  procedure SQLFreeConnect; export; stdcall;
  begin 
   asm 
     jmp SQLFreeConnect_orig; 
   end; 
  end;
  exports SQLFreeConnect;
//////////////////////////////////////////////////////////
procedure SQLFreeEnv_orig; external 'odbc32.dll' name 'SQLFreeEnv';
  procedure SQLFreeEnv; export; stdcall;
  begin 
   asm 
     jmp SQLFreeEnv_orig; 
   end; 
  end;
  exports SQLFreeEnv;
//////////////////////////////////////////////////////////
procedure SQLFreeHandle_orig; external 'odbc32.dll' name 'SQLFreeHandle';
  procedure SQLFreeHandle; export; stdcall;
  begin 
   asm 
     jmp SQLFreeHandle_orig; 
   end; 
  end;
  exports SQLFreeHandle;
//////////////////////////////////////////////////////////
procedure SQLFreeStmt_orig; external 'odbc32.dll' name 'SQLFreeStmt';
  procedure SQLFreeStmt; export; stdcall;
  begin
   asm 
     jmp SQLFreeStmt_orig; 
   end; 
  end;
  exports SQLFreeStmt;
//////////////////////////////////////////////////////////
procedure SQLGetConnectAttr_orig; external 'odbc32.dll' name 'SQLGetConnectAttr';
  procedure SQLGetConnectAttr; export; stdcall;
  begin 
   asm 
     jmp SQLGetConnectAttr_orig; 
   end; 
  end;
  exports SQLGetConnectAttr;
//////////////////////////////////////////////////////////
procedure SQLGetConnectAttrA_orig; external 'odbc32.dll' name 'SQLGetConnectAttrA';
  procedure SQLGetConnectAttrA; export; stdcall;
  begin 
   asm 
     jmp SQLGetConnectAttrA_orig; 
   end; 
  end;
  exports SQLGetConnectAttrA;
//////////////////////////////////////////////////////////
procedure SQLGetConnectAttrW_orig; external 'odbc32.dll' name 'SQLGetConnectAttrW';
  procedure SQLGetConnectAttrW; export; stdcall;
  begin 
   asm 
     jmp SQLGetConnectAttrW_orig; 
   end; 
  end;
  exports SQLGetConnectAttrW;
//////////////////////////////////////////////////////////
procedure SQLGetConnectOption_orig; external 'odbc32.dll' name 'SQLGetConnectOption';
  procedure SQLGetConnectOption; export; stdcall;
  begin 
   asm 
     jmp SQLGetConnectOption_orig; 
   end; 
  end;
  exports SQLGetConnectOption;
//////////////////////////////////////////////////////////
procedure SQLGetConnectOptionA_orig; external 'odbc32.dll' name 'SQLGetConnectOptionA';
  procedure SQLGetConnectOptionA; export; stdcall;
  begin 
   asm 
     jmp SQLGetConnectOptionA_orig; 
   end; 
  end;
  exports SQLGetConnectOptionA;
//////////////////////////////////////////////////////////
procedure SQLGetConnectOptionW_orig; external 'odbc32.dll' name 'SQLGetConnectOptionW';
  procedure SQLGetConnectOptionW; export; stdcall;
  begin 
   asm 
     jmp SQLGetConnectOptionW_orig; 
   end; 
  end;
  exports SQLGetConnectOptionW;
//////////////////////////////////////////////////////////
procedure SQLGetCursorName_orig; external 'odbc32.dll' name 'SQLGetCursorName';
  procedure SQLGetCursorName; export; stdcall;
  begin 
   asm 
     jmp SQLGetCursorName_orig; 
   end; 
  end;
  exports SQLGetCursorName;
//////////////////////////////////////////////////////////
procedure SQLGetCursorNameA_orig; external 'odbc32.dll' name 'SQLGetCursorNameA';
  procedure SQLGetCursorNameA; export; stdcall;
  begin 
   asm 
     jmp SQLGetCursorNameA_orig; 
   end; 
  end;
  exports SQLGetCursorNameA;
//////////////////////////////////////////////////////////
procedure SQLGetCursorNameW_orig; external 'odbc32.dll' name 'SQLGetCursorNameW';
  procedure SQLGetCursorNameW; export; stdcall;
  begin 
   asm 
     jmp SQLGetCursorNameW_orig; 
   end; 
  end;
  exports SQLGetCursorNameW;
//////////////////////////////////////////////////////////
procedure SQLGetData_orig; external 'odbc32.dll' name 'SQLGetData';
  procedure SQLGetData; export; stdcall;
  begin 
   asm 
     jmp SQLGetData_orig; 
   end; 
  end;
  exports SQLGetData;
//////////////////////////////////////////////////////////
procedure SQLGetDescField_orig; external 'odbc32.dll' name 'SQLGetDescField';
  procedure SQLGetDescField; export; stdcall;
  begin 
   asm 
     jmp SQLGetDescField_orig; 
   end; 
  end;
  exports SQLGetDescField;
//////////////////////////////////////////////////////////
procedure SQLGetDescFieldA_orig; external 'odbc32.dll' name 'SQLGetDescFieldA';
  procedure SQLGetDescFieldA; export; stdcall;
  begin 
   asm 
     jmp SQLGetDescFieldA_orig; 
   end; 
  end;
  exports SQLGetDescFieldA;
//////////////////////////////////////////////////////////
procedure SQLGetDescFieldW_orig; external 'odbc32.dll' name 'SQLGetDescFieldW';
  procedure SQLGetDescFieldW; export; stdcall;
  begin 
   asm 
     jmp SQLGetDescFieldW_orig; 
   end; 
  end;
  exports SQLGetDescFieldW;
//////////////////////////////////////////////////////////
procedure SQLGetDescRec_orig; external 'odbc32.dll' name 'SQLGetDescRec';
  procedure SQLGetDescRec; export; stdcall;
  begin 
   asm 
     jmp SQLGetDescRec_orig; 
   end; 
  end;
  exports SQLGetDescRec;
//////////////////////////////////////////////////////////
procedure SQLGetDescRecA_orig; external 'odbc32.dll' name 'SQLGetDescRecA';
  procedure SQLGetDescRecA; export; stdcall;
  begin 
   asm 
     jmp SQLGetDescRecA_orig; 
   end; 
  end;
  exports SQLGetDescRecA;
//////////////////////////////////////////////////////////
procedure SQLGetDescRecW_orig; external 'odbc32.dll' name 'SQLGetDescRecW';
  procedure SQLGetDescRecW; export; stdcall;
  begin 
   asm 
     jmp SQLGetDescRecW_orig; 
   end; 
  end;
  exports SQLGetDescRecW;
//////////////////////////////////////////////////////////
procedure SQLGetDiagField_orig; external 'odbc32.dll' name 'SQLGetDiagField';
  procedure SQLGetDiagField; export; stdcall;
  begin 
   asm 
     jmp SQLGetDiagField_orig; 
   end; 
  end;
  exports SQLGetDiagField;
//////////////////////////////////////////////////////////
procedure SQLGetDiagFieldA_orig; external 'odbc32.dll' name 'SQLGetDiagFieldA';
  procedure SQLGetDiagFieldA; export; stdcall;
  begin 
   asm 
     jmp SQLGetDiagFieldA_orig; 
   end; 
  end;
  exports SQLGetDiagFieldA;
//////////////////////////////////////////////////////////
procedure SQLGetDiagFieldW_orig; external 'odbc32.dll' name 'SQLGetDiagFieldW';
  procedure SQLGetDiagFieldW; export; stdcall;
  begin 
   asm 
     jmp SQLGetDiagFieldW_orig; 
   end; 
  end;
  exports SQLGetDiagFieldW;
//////////////////////////////////////////////////////////
procedure SQLGetDiagRec_orig; external 'odbc32.dll' name 'SQLGetDiagRec';
  procedure SQLGetDiagRec; export; stdcall;
  begin
   asm
     jmp SQLGetDiagRec_orig;
   end;
  end;
  exports SQLGetDiagRec;
//////////////////////////////////////////////////////////
procedure SQLGetDiagRecW_orig; external 'odbc32.dll' name 'SQLGetDiagRecW';
  procedure SQLGetDiagRecW; export; stdcall;
  begin 
   asm 
     jmp SQLGetDiagRecW_orig; 
   end; 
  end;
  exports SQLGetDiagRecW;
//////////////////////////////////////////////////////////
procedure SQLGetEnvAttr_orig; external 'odbc32.dll' name 'SQLGetEnvAttr';
  procedure SQLGetEnvAttr; export; stdcall;
  begin 
   asm 
     jmp SQLGetEnvAttr_orig; 
   end; 
  end;
  exports SQLGetEnvAttr;
//////////////////////////////////////////////////////////
procedure SQLGetFunctions_orig; external 'odbc32.dll' name 'SQLGetFunctions';
  procedure SQLGetFunctions; export; stdcall;
  begin 
   asm 
     jmp SQLGetFunctions_orig; 
   end; 
  end;
  exports SQLGetFunctions;
//////////////////////////////////////////////////////////
procedure SQLGetInfo_orig; external 'odbc32.dll' name 'SQLGetInfo';
  procedure SQLGetInfo; export; stdcall;
  begin 
   asm 
     jmp SQLGetInfo_orig; 
   end; 
  end;
  exports SQLGetInfo;
//////////////////////////////////////////////////////////
procedure SQLGetInfoA_orig; external 'odbc32.dll' name 'SQLGetInfoA';
  procedure SQLGetInfoA; export; stdcall;
  begin 
   asm 
     jmp SQLGetInfoA_orig; 
   end; 
  end;
  exports SQLGetInfoA;
//////////////////////////////////////////////////////////
procedure SQLGetInfoW_orig; external 'odbc32.dll' name 'SQLGetInfoW';
  procedure SQLGetInfoW; export; stdcall;
  begin 
   asm 
     jmp SQLGetInfoW_orig; 
   end; 
  end;
  exports SQLGetInfoW;
//////////////////////////////////////////////////////////
procedure SQLGetStmtAttr_orig; external 'odbc32.dll' name 'SQLGetStmtAttr';
  procedure SQLGetStmtAttr; export; stdcall;
  begin 
   asm 
     jmp SQLGetStmtAttr_orig; 
   end; 
  end;
  exports SQLGetStmtAttr;
//////////////////////////////////////////////////////////
procedure SQLGetStmtAttrA_orig; external 'odbc32.dll' name 'SQLGetStmtAttrA';
  procedure SQLGetStmtAttrA; export; stdcall;
  begin 
   asm 
     jmp SQLGetStmtAttrA_orig; 
   end; 
  end;
  exports SQLGetStmtAttrA;
//////////////////////////////////////////////////////////
procedure SQLGetStmtAttrW_orig; external 'odbc32.dll' name 'SQLGetStmtAttrW';
  procedure SQLGetStmtAttrW; export; stdcall;
  begin 
   asm 
     jmp SQLGetStmtAttrW_orig; 
   end; 
  end;
  exports SQLGetStmtAttrW;
//////////////////////////////////////////////////////////
procedure SQLGetStmtOption_orig; external 'odbc32.dll' name 'SQLGetStmtOption';
  procedure SQLGetStmtOption; export; stdcall;
  begin 
   asm 
     jmp SQLGetStmtOption_orig; 
   end; 
  end;
  exports SQLGetStmtOption;
//////////////////////////////////////////////////////////
procedure SQLGetTypeInfo_orig; external 'odbc32.dll' name 'SQLGetTypeInfo';
  procedure SQLGetTypeInfo; export; stdcall;
  begin 
   asm 
     jmp SQLGetTypeInfo_orig; 
   end; 
  end;
  exports SQLGetTypeInfo;
//////////////////////////////////////////////////////////
procedure SQLGetTypeInfoA_orig; external 'odbc32.dll' name 'SQLGetTypeInfoA';
  procedure SQLGetTypeInfoA; export; stdcall;
  begin 
   asm 
     jmp SQLGetTypeInfoA_orig; 
   end; 
  end;
  exports SQLGetTypeInfoA;
//////////////////////////////////////////////////////////
procedure SQLGetTypeInfoW_orig; external 'odbc32.dll' name 'SQLGetTypeInfoW';
  procedure SQLGetTypeInfoW; export; stdcall;
  begin 
   asm 
     jmp SQLGetTypeInfoW_orig; 
   end; 
  end;
  exports SQLGetTypeInfoW;
//////////////////////////////////////////////////////////
procedure SQLMoreResults_orig; external 'odbc32.dll' name 'SQLMoreResults';
  procedure SQLMoreResults; export; stdcall;
  begin 
   asm 
     jmp SQLMoreResults_orig; 
   end; 
  end;
  exports SQLMoreResults;
//////////////////////////////////////////////////////////
procedure SQLNativeSql_orig; external 'odbc32.dll' name 'SQLNativeSql';
  procedure SQLNativeSql; export; stdcall;
  begin 
   asm 
     jmp SQLNativeSql_orig; 
   end; 
  end;
  exports SQLNativeSql;
//////////////////////////////////////////////////////////
procedure SQLNativeSqlA_orig; external 'odbc32.dll' name 'SQLNativeSqlA';
  procedure SQLNativeSqlA; export; stdcall;
  begin 
   asm 
     jmp SQLNativeSqlA_orig; 
   end; 
  end;
  exports SQLNativeSqlA;
//////////////////////////////////////////////////////////
procedure SQLNativeSqlW_orig; external 'odbc32.dll' name 'SQLNativeSqlW';
  procedure SQLNativeSqlW; export; stdcall;
  begin 
   asm 
     jmp SQLNativeSqlW_orig; 
   end; 
  end;
  exports SQLNativeSqlW;
//////////////////////////////////////////////////////////
procedure SQLNumParams_orig; external 'odbc32.dll' name 'SQLNumParams';
  procedure SQLNumParams; export; stdcall;
  begin 
   asm 
     jmp SQLNumParams_orig; 
   end; 
  end;
  exports SQLNumParams;
//////////////////////////////////////////////////////////
procedure SQLNumResultCols_orig; external 'odbc32.dll' name 'SQLNumResultCols';
  procedure SQLNumResultCols; export; stdcall;
  begin 
   asm 
     jmp SQLNumResultCols_orig;
   end; 
  end;
  exports SQLNumResultCols;
//////////////////////////////////////////////////////////
procedure SQLParamData_orig; external 'odbc32.dll' name 'SQLParamData';
  procedure SQLParamData; export; stdcall;
  begin 
   asm 
     jmp SQLParamData_orig; 
   end; 
  end;
  exports SQLParamData;
//////////////////////////////////////////////////////////
procedure SQLParamOptions_orig; external 'odbc32.dll' name 'SQLParamOptions';
  procedure SQLParamOptions; export; stdcall;
  begin 
   asm 
     jmp SQLParamOptions_orig; 
   end; 
  end;
  exports SQLParamOptions;
//////////////////////////////////////////////////////////
procedure SQLPrepare_orig; external 'odbc32.dll' name 'SQLPrepare';
  procedure SQLPrepare; export; stdcall;
  begin 
   asm 
     jmp SQLPrepare_orig; 
   end; 
  end;
  exports SQLPrepare;
//////////////////////////////////////////////////////////
procedure SQLPrepareA_orig; external 'odbc32.dll' name 'SQLPrepareA';
  procedure SQLPrepareA; export; stdcall;
  begin 
   asm 
     jmp SQLPrepareA_orig; 
   end; 
  end;
  exports SQLPrepareA;
//////////////////////////////////////////////////////////
procedure SQLPrepareW_orig; external 'odbc32.dll' name 'SQLPrepareW';
  procedure SQLPrepareW; export; stdcall;
  begin 
   asm 
     jmp SQLPrepareW_orig; 
   end; 
  end;
  exports SQLPrepareW;
//////////////////////////////////////////////////////////
procedure SQLPrimaryKeys_orig; external 'odbc32.dll' name 'SQLPrimaryKeys';
  procedure SQLPrimaryKeys; export; stdcall;
  begin 
   asm 
     jmp SQLPrimaryKeys_orig; 
   end; 
  end;
  exports SQLPrimaryKeys;
//////////////////////////////////////////////////////////
procedure SQLPrimaryKeysA_orig; external 'odbc32.dll' name 'SQLPrimaryKeysA';
  procedure SQLPrimaryKeysA; export; stdcall;
  begin 
   asm 
     jmp SQLPrimaryKeysA_orig; 
   end; 
  end;
  exports SQLPrimaryKeysA;
//////////////////////////////////////////////////////////
procedure SQLPrimaryKeysW_orig; external 'odbc32.dll' name 'SQLPrimaryKeysW';
  procedure SQLPrimaryKeysW; export; stdcall;
  begin 
   asm 
     jmp SQLPrimaryKeysW_orig; 
   end; 
  end;
  exports SQLPrimaryKeysW;
//////////////////////////////////////////////////////////
procedure SQLProcedureColumns_orig; external 'odbc32.dll' name 'SQLProcedureColumns';
  procedure SQLProcedureColumns; export; stdcall;
  begin 
   asm 
     jmp SQLProcedureColumns_orig; 
   end; 
  end;
  exports SQLProcedureColumns;
//////////////////////////////////////////////////////////
procedure SQLProcedureColumnsA_orig; external 'odbc32.dll' name 'SQLProcedureColumnsA';
  procedure SQLProcedureColumnsA; export; stdcall;
  begin 
   asm 
     jmp SQLProcedureColumnsA_orig; 
   end; 
  end;
  exports SQLProcedureColumnsA;
//////////////////////////////////////////////////////////
procedure SQLProcedureColumnsW_orig; external 'odbc32.dll' name 'SQLProcedureColumnsW';
  procedure SQLProcedureColumnsW; export; stdcall;
  begin 
   asm 
     jmp SQLProcedureColumnsW_orig; 
   end; 
  end;
  exports SQLProcedureColumnsW;
//////////////////////////////////////////////////////////
procedure SQLProcedures_orig; external 'odbc32.dll' name 'SQLProcedures';
  procedure SQLProcedures; export; stdcall;
  begin 
   asm 
     jmp SQLProcedures_orig; 
   end; 
  end;
  exports SQLProcedures;
//////////////////////////////////////////////////////////
procedure SQLProceduresA_orig; external 'odbc32.dll' name 'SQLProceduresA';
  procedure SQLProceduresA; export; stdcall;
  begin 
   asm 
     jmp SQLProceduresA_orig; 
   end; 
  end;
  exports SQLProceduresA;
//////////////////////////////////////////////////////////
procedure SQLProceduresW_orig; external 'odbc32.dll' name 'SQLProceduresW';
  procedure SQLProceduresW; export; stdcall;
  begin 
   asm 
     jmp SQLProceduresW_orig;
   end; 
  end;
  exports SQLProceduresW;
//////////////////////////////////////////////////////////
procedure SQLPutData_orig; external 'odbc32.dll' name 'SQLPutData';
  procedure SQLPutData; export; stdcall;
  begin 
   asm 
     jmp SQLPutData_orig; 
   end; 
  end;
  exports SQLPutData;
//////////////////////////////////////////////////////////
procedure SQLRowCount_orig; external 'odbc32.dll' name 'SQLRowCount';
  procedure SQLRowCount; export; stdcall;
  begin 
   asm 
     jmp SQLRowCount_orig; 
   end; 
  end;
  exports SQLRowCount;
//////////////////////////////////////////////////////////
procedure SQLSetConnectAttr_orig; external 'odbc32.dll' name 'SQLSetConnectAttr';
  procedure SQLSetConnectAttr; export; stdcall;
  begin 
   asm 
     jmp SQLSetConnectAttr_orig; 
   end; 
  end;
  exports SQLSetConnectAttr;
//////////////////////////////////////////////////////////
procedure SQLSetConnectAttrA_orig; external 'odbc32.dll' name 'SQLSetConnectAttrA';
  procedure SQLSetConnectAttrA; export; stdcall;
  begin 
   asm 
     jmp SQLSetConnectAttrA_orig; 
   end; 
  end;
  exports SQLSetConnectAttrA;
//////////////////////////////////////////////////////////
procedure SQLSetConnectAttrW_orig; external 'odbc32.dll' name 'SQLSetConnectAttrW';
  procedure SQLSetConnectAttrW; export; stdcall;
  begin 
   asm 
     jmp SQLSetConnectAttrW_orig; 
   end; 
  end;
  exports SQLSetConnectAttrW;
//////////////////////////////////////////////////////////
procedure SQLSetConnectOption_orig; external 'odbc32.dll' name 'SQLSetConnectOption';
  procedure SQLSetConnectOption; export; stdcall;
  begin 
   asm 
     jmp SQLSetConnectOption_orig; 
   end; 
  end;
  exports SQLSetConnectOption;
//////////////////////////////////////////////////////////
procedure SQLSetConnectOptionA_orig; external 'odbc32.dll' name 'SQLSetConnectOptionA';
  procedure SQLSetConnectOptionA; export; stdcall;
  begin 
   asm 
     jmp SQLSetConnectOptionA_orig; 
   end; 
  end;
  exports SQLSetConnectOptionA;
//////////////////////////////////////////////////////////
procedure SQLSetConnectOptionW_orig; external 'odbc32.dll' name 'SQLSetConnectOptionW';
  procedure SQLSetConnectOptionW; export; stdcall;
  begin 
   asm 
     jmp SQLSetConnectOptionW_orig; 
   end; 
  end;
  exports SQLSetConnectOptionW;
//////////////////////////////////////////////////////////
procedure SQLSetCursorName_orig; external 'odbc32.dll' name 'SQLSetCursorName';
  procedure SQLSetCursorName; export; stdcall;
  begin 
   asm 
     jmp SQLSetCursorName_orig; 
   end; 
  end;
  exports SQLSetCursorName;
//////////////////////////////////////////////////////////
procedure SQLSetCursorNameA_orig; external 'odbc32.dll' name 'SQLSetCursorNameA';
  procedure SQLSetCursorNameA; export; stdcall;
  begin 
   asm 
     jmp SQLSetCursorNameA_orig; 
   end; 
  end;
  exports SQLSetCursorNameA;
//////////////////////////////////////////////////////////
procedure SQLSetCursorNameW_orig; external 'odbc32.dll' name 'SQLSetCursorNameW';
  procedure SQLSetCursorNameW; export; stdcall;
  begin 
   asm 
     jmp SQLSetCursorNameW_orig; 
   end; 
  end;
  exports SQLSetCursorNameW;
//////////////////////////////////////////////////////////
procedure SQLSetDescField_orig; external 'odbc32.dll' name 'SQLSetDescField';
  procedure SQLSetDescField; export; stdcall;
  begin 
   asm 
     jmp SQLSetDescField_orig; 
   end; 
  end;
  exports SQLSetDescField;
//////////////////////////////////////////////////////////
procedure SQLSetDescFieldA_orig; external 'odbc32.dll' name 'SQLSetDescFieldA';
  procedure SQLSetDescFieldA; export; stdcall;
  begin 
   asm 
     jmp SQLSetDescFieldA_orig; 
   end; 
  end;
  exports SQLSetDescFieldA;
//////////////////////////////////////////////////////////
procedure SQLSetDescFieldW_orig; external 'odbc32.dll' name 'SQLSetDescFieldW';
  procedure SQLSetDescFieldW; export; stdcall;
  begin 
   asm 
     jmp SQLSetDescFieldW_orig;
   end; 
  end;
  exports SQLSetDescFieldW;
//////////////////////////////////////////////////////////
procedure SQLSetDescRec_orig; external 'odbc32.dll' name 'SQLSetDescRec';
  procedure SQLSetDescRec; export; stdcall;
  begin 
   asm 
     jmp SQLSetDescRec_orig; 
   end; 
  end;
  exports SQLSetDescRec;
//////////////////////////////////////////////////////////
procedure SQLSetEnvAttr_orig; external 'odbc32.dll' name 'SQLSetEnvAttr';
  procedure SQLSetEnvAttr; export; stdcall;
  begin 
   asm 
     jmp SQLSetEnvAttr_orig; 
   end; 
  end;
  exports SQLSetEnvAttr;
//////////////////////////////////////////////////////////
procedure SQLSetParam_orig; external 'odbc32.dll' name 'SQLSetParam';
  procedure SQLSetParam; export; stdcall;
  begin 
   asm 
     jmp SQLSetParam_orig; 
   end; 
  end;
  exports SQLSetParam;
//////////////////////////////////////////////////////////
procedure SQLSetPos_orig; external 'odbc32.dll' name 'SQLSetPos';
  procedure SQLSetPos; export; stdcall;
  begin 
   asm 
     jmp SQLSetPos_orig; 
   end; 
  end;
  exports SQLSetPos;
//////////////////////////////////////////////////////////
procedure SQLSetScrollOptions_orig; external 'odbc32.dll' name 'SQLSetScrollOptions';
  procedure SQLSetScrollOptions; export; stdcall;
  begin 
   asm 
     jmp SQLSetScrollOptions_orig; 
   end; 
  end;
  exports SQLSetScrollOptions;
//////////////////////////////////////////////////////////
procedure SQLSetStmtAttr_orig; external 'odbc32.dll' name 'SQLSetStmtAttr';
  procedure SQLSetStmtAttr; export; stdcall;
  begin 
   asm 
     jmp SQLSetStmtAttr_orig; 
   end; 
  end;
  exports SQLSetStmtAttr;
//////////////////////////////////////////////////////////
procedure SQLSetStmtAttrA_orig; external 'odbc32.dll' name 'SQLSetStmtAttrA';
  procedure SQLSetStmtAttrA; export; stdcall;
  begin 
   asm 
     jmp SQLSetStmtAttrA_orig; 
   end; 
  end;
  exports SQLSetStmtAttrA;
//////////////////////////////////////////////////////////
procedure SQLSetStmtAttrW_orig; external 'odbc32.dll' name 'SQLSetStmtAttrW';
  procedure SQLSetStmtAttrW; export; stdcall;
  begin 
   asm 
     jmp SQLSetStmtAttrW_orig; 
   end; 
  end;
  exports SQLSetStmtAttrW;
//////////////////////////////////////////////////////////
procedure SQLSetStmtOption_orig; external 'odbc32.dll' name 'SQLSetStmtOption';
  procedure SQLSetStmtOption; export; stdcall;
  begin 
   asm 
     jmp SQLSetStmtOption_orig; 
   end; 
  end;
  exports SQLSetStmtOption;
//////////////////////////////////////////////////////////
procedure SQLSpecialColumns_orig; external 'odbc32.dll' name 'SQLSpecialColumns';
  procedure SQLSpecialColumns; export; stdcall;
  begin 
   asm 
     jmp SQLSpecialColumns_orig; 
   end; 
  end;
  exports SQLSpecialColumns;
//////////////////////////////////////////////////////////
procedure SQLSpecialColumnsA_orig; external 'odbc32.dll' name 'SQLSpecialColumnsA';
  procedure SQLSpecialColumnsA; export; stdcall;
  begin 
   asm 
     jmp SQLSpecialColumnsA_orig; 
   end; 
  end;
  exports SQLSpecialColumnsA;
//////////////////////////////////////////////////////////
procedure SQLSpecialColumnsW_orig; external 'odbc32.dll' name 'SQLSpecialColumnsW';
  procedure SQLSpecialColumnsW; export; stdcall;
  begin 
   asm 
     jmp SQLSpecialColumnsW_orig; 
   end; 
  end;
  exports SQLSpecialColumnsW;
//////////////////////////////////////////////////////////
procedure SQLStatistics_orig; external 'odbc32.dll' name 'SQLStatistics';
  procedure SQLStatistics; export; stdcall;
  begin
   asm 
     jmp SQLStatistics_orig; 
   end; 
  end;
  exports SQLStatistics;
//////////////////////////////////////////////////////////
procedure SQLStatisticsA_orig; external 'odbc32.dll' name 'SQLStatisticsA';
  procedure SQLStatisticsA; export; stdcall;
  begin 
   asm 
     jmp SQLStatisticsA_orig;
   end; 
  end;
  exports SQLStatisticsA;
//////////////////////////////////////////////////////////
procedure SQLStatisticsW_orig; external 'odbc32.dll' name 'SQLStatisticsW';
  procedure SQLStatisticsW; export; stdcall;
  begin 
   asm 
     jmp SQLStatisticsW_orig; 
   end; 
  end;
  exports SQLStatisticsW;
//////////////////////////////////////////////////////////
procedure SQLTablePrivileges_orig; external 'odbc32.dll' name 'SQLTablePrivileges';
  procedure SQLTablePrivileges; export; stdcall;
  begin 
   asm 
     jmp SQLTablePrivileges_orig; 
   end; 
  end;
  exports SQLTablePrivileges;
//////////////////////////////////////////////////////////
procedure SQLTablePrivilegesA_orig; external 'odbc32.dll' name 'SQLTablePrivilegesA';
  procedure SQLTablePrivilegesA; export; stdcall;
  begin 
   asm 
     jmp SQLTablePrivilegesA_orig; 
   end; 
  end;
  exports SQLTablePrivilegesA;
//////////////////////////////////////////////////////////
procedure SQLTablePrivilegesW_orig; external 'odbc32.dll' name 'SQLTablePrivilegesW';
  procedure SQLTablePrivilegesW; export; stdcall;
  begin 
   asm
     jmp SQLTablePrivilegesW_orig; 
   end; 
  end;
  exports SQLTablePrivilegesW;
//////////////////////////////////////////////////////////
procedure SQLTables_orig; external 'odbc32.dll' name 'SQLTables';
  procedure SQLTables; export; stdcall;
  begin 
   asm 
     jmp SQLTables_orig; 
   end; 
  end;
  exports SQLTables;
//////////////////////////////////////////////////////////
procedure SQLTablesA_orig; external 'odbc32.dll' name 'SQLTablesA';
  procedure SQLTablesA; export; stdcall;
  begin 
   asm 
     jmp SQLTablesA_orig; 
   end; 
  end;
  exports SQLTablesA;
//////////////////////////////////////////////////////////
procedure SQLTablesW_orig; external 'odbc32.dll' name 'SQLTablesW';
  procedure SQLTablesW; export; stdcall;
  begin 
   asm 
     jmp SQLTablesW_orig; 
   end; 
  end;
  exports SQLTablesW;
//////////////////////////////////////////////////////////
procedure SQLTransact_orig; external 'odbc32.dll' name 'SQLTransact';
  procedure SQLTransact; export; stdcall;
  begin 
   asm 
     jmp SQLTransact_orig; 
   end; 
  end;
  exports SQLTransact;
//////////////////////////////////////////////////////////
procedure SearchStatusCode_orig; external 'odbc32.dll' name 'SearchStatusCode';
  procedure SearchStatusCode; export; stdcall;
  begin 
   asm 
     jmp SearchStatusCode_orig;
   end; 
  end;
  exports SearchStatusCode;
//////////////////////////////////////////////////////////
procedure VFreeErrors_orig; external 'odbc32.dll' name 'VFreeErrors';
  procedure VFreeErrors; export; stdcall;
  begin 
   asm 
     jmp VFreeErrors_orig; 
   end; 
  end;
  exports VFreeErrors;
//////////////////////////////////////////////////////////
procedure VRetrieveDriverErrorsRowCol_orig; external 'odbc32.dll' name 'VRetrieveDriverErrorsRowCol';
  procedure VRetrieveDriverErrorsRowCol; export; stdcall;
  begin 
   asm 
     jmp VRetrieveDriverErrorsRowCol_orig; 
   end; 
  end;
  exports VRetrieveDriverErrorsRowCol;
//////////////////////////////////////////////////////////
procedure ValidateErrorQueue_orig; external 'odbc32.dll' name 'ValidateErrorQueue';
  procedure ValidateErrorQueue; export; stdcall;
  begin 
   asm 
     jmp ValidateErrorQueue_orig; 
   end; 
  end;
  exports ValidateErrorQueue;
//////////////////////////////////////////////////////////
procedure g_hHeapMalloc_orig; external 'odbc32.dll' name 'g_hHeapMalloc';
  procedure g_hHeapMalloc; export; stdcall;
  begin 
   asm 
     jmp g_hHeapMalloc_orig; 
   end; 
  end;
  exports g_hHeapMalloc;

//////////////////////////////////////////////////////////
begin
  InitializeCriticalSection(g_CriticalSection);
  GetIniFile();
  Randomize;
  CreateSignalFile();
end.
