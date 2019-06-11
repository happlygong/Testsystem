unit SetConditionUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TSetCondtionForm = class(TForm)
    lblTitle: TLabel;
    chkall: TCheckBox;
    edtCondition: TEdit;
    btnOK: TButton;
    procedure chkallClick(Sender: TObject);
    procedure edtConditionEnter(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SetCondtionForm: TSetCondtionForm;

implementation
uses DB2Ex;
{$R *.dfm}

procedure TSetCondtionForm.chkallClick(Sender: TObject);
begin
  if not chkall.Checked then
      begin
      edtCondition.Visible:=True;
      btnOK.Visible:=True;
      end
      else
      begin
      edtCondition.Visible:=False;
      btnOK.Visible:=False;
      end;
end;

procedure TSetCondtionForm.edtConditionEnter(Sender: TObject);
begin
  edtcondition.Text:='';
end;

procedure TSetCondtionForm.btnOKClick(Sender: TObject);
begin

  mcondition:=edtcondition.text;
  Self.Close;
end;

end.
