unit dm2;

interface

uses
  ADODB, Classes, DB;

type
  TDM = class(TDataModule)
    AC1: TADOConnection;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DM: TDM;
  str0:WideString;
implementation

{$R *.dfm}

procedure TDM.DataModuleCreate(Sender: TObject);
begin
DM.AC1.ConnectionString:=str0;
end;

end.
