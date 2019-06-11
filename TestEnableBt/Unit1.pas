unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    procedure Edit1Change(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  vs:boolean=False;
implementation

{$R *.dfm}

procedure TForm1.Edit1Change(Sender: TObject);

begin
  if Edit1.Text='111111' then
     begin
       if not vs then
       vs:=True; Button1.Enabled:=True;
     end else
     begin
       if vs then
       vs:=False;Button1.Enabled:=False;
     end;
end;

end.
