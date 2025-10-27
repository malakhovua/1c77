uses Windows, SysUtils, PatchMemory, HookMemory, Dialogs, uBalloon;

function ToString(Value: Variant): String;
begin
  case TVarData(Value).VType of
    varSmallInt,
    varInteger   : Result := IntToStr(Value);
    varSingle,
    varDouble,
    varCurrency  : Result := FloatToStr(Value);
    varDate      : Result := FormatDateTime('dd/mm/yyyy', Value);
    varBoolean   : if Value then Result := 'T' else Result := 'F';
    varString    : Result := Value;
    else            Result := IntToStr(Value);
  end;
end;

type
  TWinVersion = (wvUnknown,wv95,wv98,wvME,wvNT3,wvNT4,wvW2K,wvXP,wv2003, wvVista, wvSeven);

function DetectWinVersion : TWinVersion;
var
  OSVersionInfo : TOSVersionInfo;
begin
  Result := wvUnknown;                      // ??????????? ?????? ??
  OSVersionInfo.dwOSVersionInfoSize := sizeof(TOSVersionInfo);
  if GetVersionEx(OSVersionInfo)
    then
      begin
       MessageBox(0,pchar('Major = ' + IntToStr(OSVersionInfo.DwMajorVersion) + ' Minor = ' + IntToStr(OSVersionInfo.DwMinorVersion)),pchar('OS version'), MB_ICONERROR);
        case OSVersionInfo.DwMajorVersion of
          3:  Result := wvNT3;              // Windows NT 3
          4:  case OSVersionInfo.DwMinorVersion of
                0: if OSVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT
                   then Result := wvNT4     // Windows NT 4
                   else Result := wv95;     // Windows 95
                10: Result := wv98;         // Windows 98
                90: Result := wvME;         // Windows ME
              end;
          5:  case OSVersionInfo.DwMinorVersion of
                0: Result := wvW2K;         // Windows 2000
                1: Result := wvXP;          // Windows XP
                2: Result := wv2003;        // Windows 2003
              end;
          6:  case OSVersionInfo.DwMinorVersion of
                0: Result := wvVista;         // Windows Vista
                1: Result := wvSeven;          // Windows 7
              end;
        end;
      end;
end;

function DetectWinVersionStr : string;
const
  VersStr : array[TWinVersion] of string = (
    'Unknown',
    'Windows 95',
    'Windows 98',
    'Windows ME',
    'Windows NT 3',
    'Windows NT 4',
    'Windows 2000',
    'Windows XP',
    'Windows 2003',
    'Windows Vista',
    'Windows 7'
    );
begin
  Result := VersStr[DetectWinVersion];
end;

  var s: String;
begin
  s := DetectWinVersionStr();
  MessageBox(0, pchar(s), 'OS version info', MB_ICONINFORMATION);
end.