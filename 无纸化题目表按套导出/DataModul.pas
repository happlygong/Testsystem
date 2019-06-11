unit DataModul;

interface

uses
  SysUtils, Classes, DB, ADODB;

type
  TDataMd = class(TDataModule)
    ADOConn: TADOConnection;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DataMd: TDataMd;
    str0:WideString;
implementation

{$R *.dfm}

procedure TDataMd.DataModuleCreate(Sender: TObject);
begin
DataMd.ADOConn.ConnectionString:=str0;{创建连接字符串}
end;

end.
