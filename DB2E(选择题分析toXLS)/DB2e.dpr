program DB2e;

uses
  Forms,sysutils,windows,
  DB2Ex in 'DB2Ex.pas' {Form1},
  DBe in 'DBe.pas' {DM: TDataModule};

{$R *.res}

begin
  if not FileExists('data\PenMDB.mdb') then
  begin
   messagebox(0,'��ǰĿ¼��û����Ҫ��Data�ļ��У�'+#13+#13+'����ļ�����û���������ݿ��ļ���'+#13+#13+'��ȷ��������ĵ�ǰλ��'+#13+#13+'���ȷ�Ϻ��������С�',
   '����',MB_Ok or MB_ICONHAND);
   halt;
  end;
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TDM, DM);
  Application.Run;
end.
