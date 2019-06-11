unit CallDlls;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    btn1: TButton;
    procedure btn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  procedure SynAPP(App:THandle);stdcall;external 'DLLForms.dll';
  procedure ShowForm;stdcall;external 'DLLForms.dll';

implementation

{$R *.dfm}

procedure TForm1.btn1Click(Sender: TObject);
begin
   SynAPP(Application.Handle);
   ShowForm ;
end;

end.
