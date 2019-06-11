program thread_p;

uses
  Forms,
  thread_n in 'thread_n.pas' {Form1};

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
