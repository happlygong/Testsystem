unit DllForm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs;

type
  TDllFrm = class(TForm)
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DllFrm: TDllFrm;
  procedure SynAPP(App:THandle);stdcall;
  procedure ShowForm;stdcall;

implementation

{$R *.dfm}
procedure SynAPP(App:THandle );stdcall;
begin
   Application.Handle := App;
end;

procedure ShowForm;stdcall;
begin
   try
     Dllfrm := TDllFrm.Create (Application);
        DllFrm.showmodal;
     finally
       FreeAndNil(DllFrm);
    end;
end;

end.
