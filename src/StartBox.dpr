program StartBox;

uses
  Vcl.Forms,
  OfficeLauncher in 'OfficeLauncher.pas' {Form1},
  EineInstanz_Unit in 'EineInstanz_Unit.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
