program PrintDoc;

uses
  Forms,
  PrintDocUnit1 in 'PrintDocUnit1.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '��������ڱ�ǩ��ӡ';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
