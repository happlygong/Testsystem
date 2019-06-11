unit list;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    ListBox1: TListBox;
    Button1: TButton;
    Button2: TButton;
    Edit1: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure ListBox1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
function My_RefreshForm(MyStringList: TStringList): Boolean;
procedure My_KillForm(S: String);

implementation

{$R *.dfm}

function My_RefreshForm(MyStringList: TStringList): Boolean;
var
  hCurrentWindow: HWnd;
  szText: array[0..254] of Char;
begin
  MyStringList.Clear;
  hCurrentWindow := GetWindow(application.Handle, GW_HWNDFIRST);
  while hCurrentWindow <> 0 do
  begin
    if GetWindowText(hCurrentWindow, @szText, 255) > 0 then
      MyStringList.Add(StrPas(@szText));
    hCurrentWindow := GetWindow(hCurrentWindow, GW_HWNDNEXT);
  end;
  Result := True;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  FormStrings: TStringList;
begin
  FormStrings := TStringList.Create;
  My_RefreshForm(FormStrings);
  ListBox1.Items := FormStrings;
  FormStrings.Free;
  label1.Caption := IntToStr(listbox1.Count);
end;

procedure TForm1.ListBox1Click(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to (listbox1.Count - 1) do
  begin
    try
      if listbox1.Selected[i] then
        edit1.Text := listbox1.Items.Strings[i];
    finally
    end;
  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  sender1: TObject;      //不知是否这样写
begin
  try
    my_killform(edit1.Text);
  finally
  end;
  button1click(sender1);    //去按按钮1
end;

procedure My_KillForm(S: String);
var
  Exehandle: Thandle;
begin
  ExeHandle := FindWindow(nil, PChar(S));
  if ExeHandle <> 0 then
    PostMessage(ExeHandle, WM_Quit, 0, 0);
end;

end.

