program DB2e;

uses
  Forms,sysutils,windows,
  DB2Ex in 'DB2Ex.pas' {Form1},
  DBe in 'DBe.pas' {DM: TDataModule};

{$R *.res}

begin
  if not FileExists('data\PenMDB.mdb') then
  begin
   messagebox(0,'当前目录下没有需要的Data文件夹，'+#13+#13+'或该文件夹下没有所需数据库文件，'+#13+#13+'请确定本程序的当前位置'+#13+#13+'检查确认后重新运行。',
   '错误',MB_Ok or MB_ICONHAND);
   halt;
  end;
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TDM, DM);
  Application.Run;
end.
