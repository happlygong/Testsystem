program Mytitle;

uses
  Forms,
  MyTitles in 'MyTitles.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
