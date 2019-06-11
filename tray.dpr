program tray;

uses
  Forms,
  trayUnit1 in 'trayUnit1.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.ShowMainForm:=False;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
