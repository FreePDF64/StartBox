program StartBox;

uses
  Winapi.Windows,
  System.SysUtils,
  Vcl.Forms,
  OfficeLauncher in 'OfficeLauncher.pas' {Form1};

{$R *.res}

var
  hMutex: THandle;
  MutexName: string;
begin
  MutexName := 'Mutex_' + ExtractFileName(ParamStr(0));
  hMutex := CreateMutex(nil, True, PChar(MutexName));

  if (hMutex = 0) or (GetLastError = ERROR_ALREADY_EXISTS) then
  begin
    MessageBox(0,
      'StartBox läuft bereits!',
      'Hinweis',
      MB_ICONINFORMATION or MB_OK);
    Halt;
  end;

  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;

  CloseHandle(hMutex);
end.
