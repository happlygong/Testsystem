program CreateTaoTi;

uses
    Forms, Windows,
    SysUtils,
    MainformU in 'MainformU.pas' {MainForm},
    DataModul in 'DataModul.pas' {DataMd: TDataModule};

{$R *.res}
var
    str1: widestring = 'Provider=Microsoft.Jet.OLEDB.4.0;';
    str2: widestring = 'Data Source=';
    str3: WideString = ';Jet OLEDB:Database Password="WEILAIJIAOYU";';
    str: widestring;
    flag: Integer = 0;
begin
    if FileExists('Choicemdb.mdb') then
    begin
        {   messagebox(0,'��ǰĿ¼��û����Ҫ�����ݿ��ļ�<destiny.mdb>��'+#13+#13+
                        '�뽫����������ݿ��ļ�<destiny.mdb>����ͬһ��Ŀ¼��'+#13+#13+
                        '���ȷ�Ϻ��������С�',
           '����',MB_Ok or MB_ICONHAND);
           halt; }
           //DM.AC1.ConnectionString:=str1+str2+'destiny.mdb'+str3;
        str0 := str1 + str2 + 'Choicemdb.mdb' + str3;
        flag := 1;
    end
    else if FileExists('data\Choicemdb.mdb') then
    begin
        //DM.AC1.ConnectionString:=str1+str2+'data\destiny.mdb'+str3;
        str0 := str1 + str2 + 'data\Choicemdb.mdb' + str3;
        flag := 2;
    end;
    //else
    if flag = 0 then
    begin
        MessageBox(0, '�ڵ�ǰĿ¼�»�ǰĿ¼   ' + #13 + #13 +
            '��DataĿ¼��û�з�����   ' + #13 + #13 +
            '�����ļ�"Choicemdb.mdb"' + #13 + #13 +
            '���鱾����ǰλ��' + #13 + #13 + 'Ȼ����������', '����', MB_OK or
                MB_ICONQUESTION);
        halt;
    end;
    Application.Initialize;
    Application.CreateForm(TMainForm, MainForm);
    Application.CreateForm(TDataMd, DataMd);
    Application.Run;
end.

