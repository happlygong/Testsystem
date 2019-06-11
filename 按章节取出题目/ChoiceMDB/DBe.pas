unit DBe;

interface

uses
  Classes, ADODB, DB;

type
  TDM = class(TDataModule)
    con1: TADOConnection;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DM: TDM;

implementation

{$R *.dfm}

end.