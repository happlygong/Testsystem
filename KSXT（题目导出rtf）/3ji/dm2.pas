unit dm2;

interface

uses
  ADODB, Classes, DB;

type
  TDM = class(TDataModule)
    AC1: TADOConnection;
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
