program PrintDoc;

uses
  Forms,
  PrintDocUnit1 in 'PrintDocUnit1.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '北京金飞腾标签打印';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
