unit EineInstanz_Unit;

interface

implementation

uses Windows, Dialogs, Sysutils;

var
  mHandle: THandle;    // Mutexhandle

Initialization
  mHandle := CreateMutex(NIL, True, 'StartBox'); // Anwendungsname
  if GetLastError = ERROR_ALREADY_EXISTS then
  begin
    Messagedlg('StartBox l‰uft bereits!', mtInformation, [mbok], 0);
    Halt;
  end;

finalization   // ... und Schluﬂ
  if mHandle <> 0 then
    CloseHandle(mHandle)
end.
