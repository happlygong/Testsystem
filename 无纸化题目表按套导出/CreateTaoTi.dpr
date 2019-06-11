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
        {   messagebox(0,'当前目录下没有需要的数据库文件<destiny.mdb>，'+#13+#13+
                        '请将本程序和数据库文件<destiny.mdb>放在同一个目录下'+#13+#13+
                        '检查确认后重新运行。',
           '错误',MB_Ok or MB_ICONHAND);
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
        MessageBox(0, '在当前目录下或当前目录   ' + #13 + #13 +
            '的Data目录下没有发现期   ' + #13 + #13 +
            '望的文件"Choicemdb.mdb"' + #13 + #13 +
            '请检查本程序当前位置' + #13 + #13 + '然后重新运行', '错误', MB_OK or
                MB_ICONQUESTION);
        halt;
    end;
    Application.Initialize;
    Application.CreateForm(TMainForm, MainForm);
    Application.CreateForm(TDataMd, DataMd);
    Application.Run;
end.

