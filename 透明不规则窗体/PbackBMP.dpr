program PbackBMP;

uses
  Forms,
  backBMP in 'backBMP.pas' {Form1};


begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
