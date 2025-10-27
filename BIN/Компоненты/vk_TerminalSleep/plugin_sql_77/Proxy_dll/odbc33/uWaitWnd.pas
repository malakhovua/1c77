unit uWaitWnd;
//Процедура для окна с показом сообщения (настаивал Drunkovsky)

interface
uses
  Windows,
  Messages;

const
  mycName = 'TWaitForm';
  MainWindowStyle         = WS_CAPTION;
var
  MainWindowClass    : TWndClass{Ex};
  hMainWindow        : HWND;
  hFontNormal,
  Label1        : HWND;
  canClose      : boolean;

procedure CreateWaitWindow(wndCap, WndText: PChar; tim: integer);

implementation

// Главная оконная процедура
// =============================================================================
function WindowProc(Wnd: HWND; Msg: Integer; WParam: WPARAM; LParam: LPARAM): LRESULT; stdcall;
begin
  Result := 0;
  case Msg of
    WM_DESTROY:
    begin
      Result:=0;
      PostQuitMessage(0);
    end;
    wm_close: begin
      if not canClose then Result:=0
      else Result := DefWindowProc(Wnd, Msg, WParam, LParam);
    end;
    WM_TIMER:
      begin
        KillTimer( wnd, 1 );
        canClose := true;
        SendMessage(wnd, WM_CLOSE, 0, 0);
        Result:=0
      end;
    WM_CTLCOLORSTATIC: // Изменения цвета STATIC
    begin
      if LParam = Label1 then
      begin
        Result := DefWindowProc(Wnd, Msg, WParam, LParam);
        SetTextColor(WParam, $FF0000);
      end;
    end;
  else
    Result := DefWindowProc(Wnd, Msg, WParam, LParam);
  end;
end;

function IntToStr (l: LONGINT): STRING;
begin
 Str (l, result);
end;

procedure CreateWaitWindow(wndCap, WndText: PChar; tim: integer);
// Здесь программа стартует
// =============================================================================
var
  msg:    TMSG;
  rc:     TRECT;
  br:     HBRUSH;
begin
  // Инициализируем оконный класс
  ZeroMemory (@MainWindowClass, sizeof (MainWindowClass));
  canClose:= false;

  br := CreateSolidBrush ($ffffff);

  MainWindowClass.lpfnWndProc := @WindowProc;
  MainWindowClass.hInstance := hInstance;
  MainWindowClass.hCursor := LoadCursor (0, IDC_ARROW);
  MainWindowClass.hbrBackground := br;
  MainWindowClass.lpszClassName := mycName;

  if RegisterClass (MainWindowClass) = 0 then begin
    MessageBox (0, PChar ('Невозможно зарегистрировать оконный класс. Код ошибки: ' + IntToStr (GetLastError)),
                'Ошибка', MB_OK or MB_ICONERROR);
    DeleteObject (br);
    EXIT;
  end;
  rc.Right := 200;
  rc.Bottom := 100;
  rc.Left := GetSystemMetrics (SM_CXSCREEN) div 2 - rc.Right div 2;
  rc.Top := GetSystemMetrics (SM_CYSCREEN) div 2 - rc.Bottom div 2;

  AdjustWindowRectEx (rc, MainWindowStyle, FALSE, WS_EX_TOOLWINDOW);

  hMainWindow := CreateWindowEx (WS_EX_TOOLWINDOW, MainWindowClass.lpszClassName, wndCap,
                                MainWindowStyle, rc.Left, rc.Top, rc.Right, rc.Bottom, 0, 0,
                                hInstance, nil);

  if hMainWindow = 0 then begin
    MessageBox (0, PChar ('Невозможно создать окно! Код ошибки : ' + IntToStr (GetLastError)),
                'Ошибка', MB_OK or MB_ICONERROR);
    UnregisterClass (MainWindowClass.lpszClassName, hInstance);
    DeleteObject (br);
    exit;
  end;

  // Создаем Label
  Label1 := CreateWindow('STATIC', WndText,
    WS_VISIBLE or WS_CHILD, 10, 13, 230, 14, hMainWindow, 0, hInstance, nil);

  // Создаем нужный шрифт
  hFontNormal := CreateFont(-14, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET,
                      OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY,
                      DEFAULT_PITCH or FF_DONTCARE, 'MS Sans Serif');

  // назначаем этот шрифт всем оконным элементам
  if hFontNormal <> 0 then
    SendMessage(Label1, WM_SETFONT, hFontNormal, 0);

  UpdateWindow(hMainWindow);
  // Показываем окно
  ShowWindow(hMainWindow, SW_SHOW);
  // Ставим его поверх всех
  SetWindowPos(hMainWindow, HWND_TOPMOST, 0, 0, 0, 0,
    SWP_NOSIZE or SWP_NOMOVE);
  SetTimer( hMainWindow, 1, tim, nil );
  while GetMessage(msg, 0, 0, 0) do begin
    TranslateMessage(msg);
    DispatchMessage(msg);
  end;
  DestroyWindow(hMainWindow);
  UnregisterClass (mycName, hInstance);

end;

end.
