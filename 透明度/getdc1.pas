unit getdc1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, ComCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    BitBtn1: TBitBtn;
    Edit1: TEdit;
    BitBtn2: TBitBtn;
    TrackBar1: TTrackBar;
    ComboBox1: TComboBox;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure TrackBar1Change(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  h1,c1,h2,c2:dword;
implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  hDC:HWND;
  hWin:HWND;
  color,i:integer;
begin
  //hDC:=GetDC(findwindow(nil,'touming'));
  hDC:=GetDC(self.Handle);

  color:=GetPixel(hDC,0,0);
  for i:=0 to 255 do
  SetLayeredWindowAttributes(self.Handle,0,i,2);//最后参数为1，指定颜色透明，为2，整个窗口透明
  label2.Caption:=inttohex(h1,8);
  label3.Caption:=inttohex(c1,8);
  label4.Caption:=inttohex(self.Handle,8);
  label5.Caption:=inttohex(color,8);
  SetLayeredWindowAttributes(self.Handle,color,0,1);
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  WindowExs,St:dword;
  hDC:HWND;
  hWin:HWND;
  color:cardinal;
begin
    WindowExs  :=   GetWindowLong(form1.handle, GWL_EXSTYLE) ;
    WindowExs  :=   WindowExs   Or   WS_EX_LAYERED;//  or WS_EX_TOOLWINDOW ;
    WindowExs  :=   WindowExs   and  (not WS_EX_APPWINDOW);
    St         :=   GetWindowLong(form1.Handle,GWL_STYLE);
    SetWindowLong (  form1.Handle,   GWL_EXSTYLE,   WindowExs);
    //SetWindowLong ( form1.handle,GWL_STYLE,$94CA00C4);
    //SetWindowLong ( form1.handle,GWL_EXSTYLE,$00010109 or WS_EX_TOPMOST);
    {SetWindowLong (application.Handle,GWL_EXSTYLE,
                   GetWindowLong(application.Handle,GWL_EXSTYLE)or WS_EX_TOOLWINDOW
                   and not WS_EX_APPWINDOW);
    // }
    hDC:=GetDC(form1.Handle);
    h1:=form1.Handle;
    color:=GetPixel(hDC,50,50);//可能因为此时还没有生成窗口所以得到的值不对
    c1:=color;
    SetLayeredWindowAttributes(form1.Handle,color,0,0);
    //SetLayeredWindowAttributes(
    //Application.Title:='';
end;
  //指定一个窗口透明-根据窗口名称
procedure TForm1.BitBtn1Click(Sender: TObject);
var
  WindowExs:dword;
  hDC:HWND;
  hWin:HWND;
  color:cardinal;
begin
    if edit1.text='' then edit1.text:='新仙剑-S1-1服 - 37abc';//'新仙剑-360卫士248区-360游戏中心';
    hWin:=FindWindow(nil,pchar(edit1.text));
    if hWin=0 then exit;
    WindowExs  :=   GetWindowLong(hWin, GWL_EXSTYLE) ;
    WindowExs  :=   WindowExs   Or   WS_EX_LAYERED   ;
    SetWindowLong ( hWin,   GWL_EXSTYLE,   WindowExs);
{    hDC:=GetDC(hWin);
    h1:=hWin;
    color:=GetPixel(hDC,50,50);//可能因为此时还没有生成窗口所以得到的值不对
    c1:=color;}
    SetLayeredWindowAttributes(hWin,0,trackbar1.Position,2);
    //SetLayeredWindowAttributes(
end;
  //指定一个窗口透明-根据窗口类名称
procedure TForm1.BitBtn2Click(Sender: TObject);
var
  WindowExs:dword;
  hDC:HWND;
  hWin:HWND;
  color:cardinal;
begin
    if edit1.text='' then edit1.text:='UnityWndClass';
    hWin:=FindWindow(pchar(edit1.text),nil);
    if hWin=0 then exit;
    WindowExs  :=   GetWindowLong(hWin, GWL_EXSTYLE) ;
    WindowExs  :=   WindowExs   Or   WS_EX_LAYERED   ;
    SetWindowLong ( hWin,   GWL_EXSTYLE,   WindowExs);
{    hDC:=GetDC(hWin);
    h1:=hWin;
    color:=GetPixel(hDC,50,50);//可能因为此时还没有生成窗口所以得到的值不对
    c1:=color;}
    SetLayeredWindowAttributes(hWin,$975518,trackbar1.Position,3);
    //SetLayeredWindowAttributes(
end;

procedure TForm1.TrackBar1Change(Sender: TObject);
begin
  BitBtn1Click(Sender);
  BitBtn2Click(Sender);
  label2.caption:=inttostr(trackbar1.position);
end;

procedure TForm1.ComboBox1Change(Sender: TObject);
begin
Edit1.Text:=ComboBox1.Text;
end;

end.
