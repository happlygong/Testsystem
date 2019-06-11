program Project1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1};

{$R *.res}
{$R 'myPic.res'}
begin
  Application.Initialize;
  Application.Title := '未来教育发货单合并邮件';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
