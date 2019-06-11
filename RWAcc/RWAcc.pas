unit RWAcc;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ComObj, StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Label1: TLabel;
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
   accapp:Variant;
begin
//Halt;
  try
    accapp:=CreateOLEObject('Access.Application');
    accapp.visible:=true;
  except
    showmessage('³öÊ¦»¯Ê§°Ü');
    accapp.quit;
    exit;
   end;
  accapp.openCurrentDataBase('D:\My Documents\db1.mdb');
  
end;

end.
