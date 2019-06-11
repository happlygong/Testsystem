unit backBMP;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, jpeg, ExtCtrls, StdCtrls, Menus;

type
  TForm1 = class(TForm)
    Button1: TButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    procedure Image1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure FormPaint(Sender: TObject);
  private
    { Private declarations }
  //按下左键时可以移动窗体
  procedure WMLButtonDown(var m: TMessage); message WM_LBUTTONDOWN;

  public
    { Public declarations }
  end;
const
  IDB_CIRCLE1=101;
  IDB_BACK1=100;
  IDB_MASK1=102;
var
  Form1: TForm1;

implementation

{$R *.dfm}
{$R mybmp.res}
procedure tform1.WMLButtonDown(var m: TMessage);  //这里一定要加TFORM1
begin
  ReleaseCapture;
  SendMessage(handle,wm_nclbuttondown,HTCAPTION,0);

end;
procedure TForm1.Image1Click(Sender: TObject);
var
  wdc,circldc,maskdc,bmpdc,mdc:HDC;
  hWin,hold:HWND;
  color,i:integer;
  windowsExs:DWORD;
  bCir,bMask:TBitmap;
begin
  windowsexs:=GetWindowLong(Self.Handle,GWL_EXSTYLE);
  windowsexs:=windowsExs or WS_EX_LAYERED;
  SetWindowLong(Self.Handle,GWL_EXSTYLE,windowsExs);
  wdc:=GetWindowDC(Self.handle);
  color:=GetPixel(wdc,0,0);
  SetLayeredWindowAttributes(Self.Handle,color,0,1);
  bCir:=TBitmap.Create;bMask:=TBitmap.Create;
  bCir.Handle:=loadbitmap(HInstance,MAKEINTRESOURCE(101));//'IDB_MASK1'); 用字符串的形式不行
  bMask.Handle:=LoadBitmap(HInstance,MAKEINTRESOURCE(102));//'IDB_CIRCLE1');
  //下面两句是从外部图片加载
  //bCir.LoadFromFile('C:\masm32\Source\Test\Chapter07\test\2503.bmp');
  //bMask.LoadFromFile('C:\masm32\Source\Test\Chapter07\test\2501.bmp');
  mdc:=CreateCompatibleDC(wdc);
  hold:=SelectObject(mdc,bMask.Handle);
  BitBlt(wdc,0,0,400,700,mdc,0,0,SRCAND);
  bmpdc:=CreateCompatibleDC(bmpdc);
  SelectObject(mdc,bCir.Handle);
  BitBlt(wdc,0,0,400,700,mdc,0,0,SRCPAINT);
  SelectObject(wdc,hold);
  DeleteDC(mdc);
  Beep;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  WindowsExs:DWORD;
  dc:HDC;
  color:Integer;
begin
  windowsexs:=GetWindowLong(Self.Handle,GWL_EXSTYLE);
  windowsexs:=windowsExs or WS_EX_LAYERED;
  SetWindowLong(Self.Handle,GWL_EXSTYLE,windowsExs);
  dc:=GetDC(Self.Handle);
  color:=GetPixel(dc,0,0);
  SetLayeredWindowAttributes(Self.Handle,color,255,2);
end;

procedure TForm1.FormActivate(Sender: TObject);
begin
ShowMessage('dldAc');
end;

procedure TForm1.FormShow(Sender: TObject);
begin
ShowMessage('show');
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  bmpdc:HDC;
  hWin:HWND;
  color,i:integer;
  windowsExs:DWORD;
begin
  windowsexs:=GetWindowLong(Self.Handle,GWL_EXSTYLE);
  windowsexs:=windowsExs or WS_EX_LAYERED;
  SetWindowLong(Self.Handle,GWL_EXSTYLE,windowsExs);
  bmpdc:=GetDC(Self.handle);
  color:=GetPixel(bmpdc,50,50);
  SetLayeredWindowAttributes(Self.Handle,color,0,1);
end;
  //按下左键时可以移动窗体

procedure TForm1.N1Click(Sender: TObject);
begin
halt;
end;

procedure TForm1.FormPaint(Sender: TObject);
var
  wdc,circldc,maskdc,bmpdc,mdc:HDC;
  hWin,hold:HWND;
  color,i:integer;
  windowsExs:DWORD;
  bCir,bMask:TBitmap;
begin
  windowsexs:=GetWindowLong(Self.Handle,GWL_EXSTYLE);
  windowsexs:=windowsExs or WS_EX_LAYERED;
  SetWindowLong(Self.Handle,GWL_EXSTYLE,windowsExs);
  wdc:=GetWindowDC(Self.handle);
  color:=GetPixel(wdc,0,0);
  SetLayeredWindowAttributes(Self.Handle,color,0,1);
  bCir:=TBitmap.Create;bMask:=TBitmap.Create;
  bCir.Handle:=loadbitmap(HInstance,MAKEINTRESOURCE(101));//'IDB_MASK1'); 用字符串的形式不行
  bMask.Handle:=LoadBitmap(HInstance,MAKEINTRESOURCE(102));//'IDB_CIRCLE1');
  //下面两句是从外部图片加载
  //bCir.LoadFromFile('C:\masm32\Source\Test\Chapter07\test\2503.bmp');
  //bMask.LoadFromFile('C:\masm32\Source\Test\Chapter07\test\2501.bmp');
  mdc:=CreateCompatibleDC(wdc);
  hold:=SelectObject(mdc,bMask.Handle);
  BitBlt(wdc,0,0,500,700,mdc,0,0,SRCAND);
  bmpdc:=CreateCompatibleDC(bmpdc);
  SelectObject(mdc,bCir.Handle);
  BitBlt(wdc,0,0,500,700,mdc,0,0,SRCPAINT);
  SelectObject(wdc,hold);
  DeleteDC(mdc);
  Beep;
end;

end.
