unit getdc1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, ComCtrls, jpeg;

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
    Image1: TImage;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure TrackBar1Change(Sender: TObject);
    procedure FormMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Image1Click(Sender: TObject);
    procedure Image1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Image1StartDrag(Sender: TObject;
      var DragObject: TDragObject);
    procedure Image1DragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure FormShow(Sender: TObject);
    procedure FormActivate(Sender: TObject);
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
  RoundRect(hDC,10,10,100,100,20,40);
  color:=GetPixel(hDC,0,0);
  for i:=0 to 255 do
  SetLayeredWindowAttributes(self.Handle,0,i,2);//最后参数为1，指定颜色透明，为2，整个窗口透明
  label2.Caption:='form1:'+inttohex(h1,8);
  label3.Caption:='颜色 :'+inttohex(c1,8);
  label4.Caption:='Selfh:'+inttohex(self.Handle,8);
  label5.Caption:='颜色2:'+inttohex(color,8);
  Label1.Caption:='app.h:'+inttohex(Application.Handle,8);
  SetLayeredWindowAttributes(self.Handle,color,0,1);
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  WindowExs,St:dword;
  hDC:HWND;
  hWin:HWND;
  color:cardinal;
begin
    ShowMessage('oncreate');
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
    SetLayeredWindowAttributes(form1.Handle,color,0,1);
    //SetLayeredWindowAttributes(
    //Application.Title:='';
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
var
  WindowExs:dword;
  hDC:HWND;
  hWin:HWND;
  color:cardinal;
begin
    if edit1.text='' then edit1.text:='touming';
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

procedure TForm1.BitBtn2Click(Sender: TObject);
var
  WindowExs:dword;
  hDC:HWND;
  hWin:HWND;
  color:cardinal;
begin
    if edit1.text='' then edit1.text:='touming';
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

procedure TForm1.FormMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
   UpdateWindow(Self.Handle);
   ReleaseCapture;
   SendMessage(Self.Handle,WM_NCLBUTTONDOWN,HTCAPTION,0);

end;

procedure TForm1.Image1Click(Sender: TObject);
begin
Halt;
end;

procedure TForm1.Image1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
   UpdateWindow(Self.Handle);
   ReleaseCapture;
   SendMessage(Self.Handle,WM_NCLBUTTONDOWN,HTCAPTION,0);
end;

procedure TForm1.Image1StartDrag(Sender: TObject;
  var DragObject: TDragObject);
begin
   UpdateWindow(Self.Handle);
   ReleaseCapture;
   SendMessage(Self.Handle,WM_NCLBUTTONDOWN,HTCAPTION,0);

end;

procedure TForm1.Image1DragDrop(Sender, Source: TObject; X, Y: Integer);
begin
   UpdateWindow(Self.Handle);
   ReleaseCapture;
   SendMessage(Self.Handle,WM_NCLBUTTONDOWN,HTCAPTION,0);

end;

procedure TForm1.FormShow(Sender: TObject);
var
  hDC:HWND;
  hWin:HWND;
  color,i:integer;
begin
  //hDC:=GetDC(findwindow(nil,'touming'));
  hDC:=GetDC(self.Handle);
  RoundRect(hDC,10,10,100,100,20,40);
  color:=GetPixel(hDC,0,0);
  for i:=0 to 255 do
  SetLayeredWindowAttributes(self.Handle,0,i,2);//最后参数为1，指定颜色透明，为2，整个窗口透明
  label2.Caption:=inttohex(h1,8);
  label3.Caption:=inttohex(c1,8);
  label4.Caption:=inttohex(self.Handle,8);
  label5.Caption:=inttohex(color,8);
  SetLayeredWindowAttributes(self.Handle,color,0,1);
end;

procedure TForm1.FormActivate(Sender: TObject);
begin
ShowMessage('onActiove');
end;

end.
