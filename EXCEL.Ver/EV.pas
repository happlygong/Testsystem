unit EV;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ComObj, StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  v:Variant;
  ver:Integer;
begin
  try
    v:=CreateOleObject('Excel.Application');
    ver:=v.version;
  except
    MessageBox(0,'没有正确安装Excle','',MB_OK);
    v.quit;
    Exit;
  end;
  case ver of
  1..10:ShowMessage('OLD Office');
  11:ShowMessage('office 11');
  12:ShowMessage('Office 12(2007)');
  14:ShowMessage('Office 2010');
  15:ShowMessage('Office 2013');
  else
    ShowMessage('Unkonw');
  end;
end;

end.
