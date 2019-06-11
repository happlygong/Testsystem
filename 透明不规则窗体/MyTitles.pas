unit MyTitles;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,jpeg;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    Button2: TButton;

    procedure Button1Click(Sender: TObject);
    procedure FormCanResize(Sender: TObject; var NewWidth,
      NewHeight: Integer; var Resize: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Button2Click(Sender: TObject);
  private
    Procedure MoveForm(var M:TWMNCHITTEST);Message WM_NCHITTEST;
    //����һ�Զ����¼������ء�WM_NCHITTEST����Ϣ    { Private declarations }
  // ���ڱ������ϰ��������ťʱ����ù���
  procedure WMNcLButtonDown(var m: TMessage); message WM_NCLBUTTONDOWN;
  // ���ڱ������Ϸſ������ťʱ����ù���
  procedure WMNcLButtonUp(var m: TMessage); message WM_NCLBUTTONUP;
  // ���ڱ��������ƶ����ʱ����ù���
  procedure WMNcMouseMove(var m: TMessage); message WM_NCMOUSEMOVE;
  // ���ڱ�������˫��������ťʱ����ù���
  procedure WMNcLButtonDBLClk(var m: TMessage); message WM_NCLBUTTONDBLCLK;
  // ���ڱ������ϰ�������Ұ�ťʱ����ù���
  procedure WMNcRButtonDown(var m: TMessage); message WM_NCRBUTTONDOWN;
  // ����������ʱ����ù���
  procedure WMNcPaint(var m: TMessage); message WM_NCPAINT;
  // ���������ڼ�����Ǽ���֮���л�ʱ����ù���
  procedure WMNcActivate(var m: TMessage); message WM_NCACTIVATE;

  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  win:TRect;
implementation

{$R *.dfm}


Procedure TForm1.MoveForm (var M:TWMNCHITTEST);
begin
inHerited;                                  //�̳У�������Լ��������Ժ���¼�
if (M.Result=HTCLIENT)                    //��������ڿͻ���
//and ((GetKeyState(vk_CONTROL) < 0))            //��⡰Ctrl�����Ƿ���
then M.Result:=HTCAPTION;                    //���ġ�.Result�����ֵ
end;

procedure tform1.WMNcLButtonDown(var m: TMessage);
begin
  inherited;
  Label1.Caption:='MouseDown';
end;
procedure tform1.WMNcLButtonUp(var m: TMessage);
begin
  inherited;
  Label1.Caption:=Label1.caption+'  MouseUp';
end;
procedure Tform1.WMNcMouseMove(var m: TMessage);
var
  bmp:TBitmap;
  ttrec,tcrec:TRect;
  pt:TPoint;
  jpg:TJPEGImage;
begin
  inherited;
  GetCursorPos(pt);Canvas.Handle:=GetWindowDC(Self.Handle);
  caption:=Format('mouse: X,Y:%d,%d',[pt.X,pt.Y]);
  GetWindowRect(Self.Handle,ttrec);
  tcrec:=GetClientRect;
  ttrec.Left:=ttrec.Right-50;
  ttrec.Right:=ttrec.Right-25;
  bmp:=TBitmap.Create;jpg:=TJPEGImage.Create;
  if PtInRect(ttrec,pt) then
  begin
//    jpg.LoadFromFile('f:\ͼƬ\8860092.jpg');
    bmp.LoadFromFile('40.bmp');
    Canvas.Draw(ttrec.Left,0,jpg);
  end else
  begin
    bmp.LoadFromFile('40.bmp');
    Canvas.Draw(ttrec.Left,0,bmp);
  end;
end;
procedure TForm1.WMNcLButtonDBLClk(var m: TMessage);
begin
  inherited;
  Label1.Caption:=label1.caption+'˫����';
end;
procedure TFORM1.WMNcRButtonDown(var m: TMessage);
begin
  inherited;
  Label1.Caption:=Label1.caption+#13+'RBdown';
end;
procedure TFORM1.WMNcPaint(var m: TMessage);
var
  tDC:HDC;
  titlebmp:TBitmap;
begin
  inherited;
//  SetWindowRgn(Self.handle, CreateRoundRectRgn(0,0,self.width,self.height,5,5),True);
  Canvas.Handle:=GetWindowDC(Self.Handle);
  win:=GetClientRect;
  titlebmp:=TBitmap.Create;
//  titlebmp.LoadFromFile('f:\ͼƬ\4411708.bmp');
//  Canvas.Handle:=tDC;
//  Canvas.Draw(700,0,titlebmp);
  //Canvas.Draw(0,0,titlebmp);
  titlebmp.LoadFromFile('4.bmp');
  Canvas.Draw(win.right-28,0,titlebmp);
  titlebmp.loadfromfile('4.bmp');
  canvas.draw(win.right-50,0,titlebmp);
  ReleaseDC(Self.Handle,Canvas.Handle);
  titlebmp.Free;
  Canvas.Handle:=0;
end;
procedure TFORM1.WMNcActivate(var m: TMessage);
begin
  inherited;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  tDC:HDC;
  titlebmp:TBitmap;
  sidew,sideh,bw,bh:Integer;
begin
  Canvas.Handle:=GetWindowDC(Self.Handle);
  titlebmp:=TBitmap.Create;
  titlebmp.LoadFromFile('068.bmp');
//  Canvas.Handle:=tDC;
  Canvas.Draw(0,0,titlebmp);
  ReleaseDC(Self.Handle,Canvas.Handle);
  titlebmp.Free;
  Canvas.Handle:=0;
  sidew:=GetSystemMetrics(SM_CXFRAME);//�߿���
  sideh:=GetSystemMetrics(SM_CYFRAME);
  bw   :=GetSystemMetrics(SM_CXSIZE); //��������ť�ĳߴ�
  bh   :=GetSystemMetrics(SM_CYSIZE);
  Button1.caption:=IntToStr(sidew)+':'+inttostr(sideh)+':'+inttostr(bw)+':'+inttostr(bh);
end;

procedure TForm1.FormCanResize(Sender: TObject; var NewWidth,
  NewHeight: Integer; var Resize: Boolean);
var
  tDC:HDC;
  titlebmp:TBitmap;
begin
  Canvas.Handle:=GetWindowDC(Self.Handle);
  titlebmp:=TBitmap.Create;
  titlebmp.LoadFromFile('40.bmp');
//  Canvas.Handle:=tDC;
  Canvas.Draw(700,0,titlebmp);
  Canvas.Draw(0,0,titlebmp);
  ReleaseDC(Self.Handle,Canvas.Handle);
  titlebmp.Free;
  Canvas.Handle:=0;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  ExSty:DWORD;
begin
{  ExSty:=GetWindowLong(Handle,GWL_EXSTYLE);
  ExSty:=ExSty or WS_EX_TRANSPARENT or WS_EX_LAYERED;
  SetWindowLong(Handle,GWL_EXSTYLE,ExSty);
  SetLayeredWindowAttributes(Handle,cardinal(clBtnFace),125,LWA_ALPHA);
  MoveWindow(Handle,Screen.Width-Self.Width,0,Self.Width,Self.Height,false);
}
SetWindowRgn(Self.handle, CreateRoundRectRgn(0,0,self.width,self.height,5,5),True);
end;

procedure TForm1.FormActivate(Sender: TObject);
var
  tDC:HDC;
  titlebmp:TBitmap;
begin
  getwindowrect(Self.Handle,win);
  win:=GetClientRect;
end;

procedure TForm1.FormResize(Sender: TObject);
var
  tDC:HDC;
  titlebmp:TBitmap;
begin
  SetWindowRgn(Self.handle, CreateRoundRectRgn(0,0,self.width,self.height,5,5),True);
  Canvas.Handle:=GetWindowDC(Self.Handle);
  win:=GetClientRect;
  titlebmp:=TBitmap.Create;
  titlebmp.LoadFromFile('40.bmp');
//  Canvas.Handle:=tDC;
  Canvas.Draw(700,0,titlebmp);
  //Canvas.Draw(0,0,titlebmp);
  titlebmp.LoadFromFile('4.bmp');
  Canvas.Draw(win.right-28,0,titlebmp);
  titlebmp.loadfromfile('4.bmp');
  canvas.draw(win.right-50,0,titlebmp);
  ReleaseDC(Self.Handle,Canvas.Handle);
  titlebmp.Free;
  Canvas.Handle:=0;
end;

procedure TForm1.FormMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
  var
    pt:TPoint; rc:trect;
begin
  GetCursorPos(pt);
  GetWindowRect(Self.Handle,rc);
  rc.Left:=rc.Left+(rc.Right-rc.left) div 2 -20;
  rc.Top:=rc.Top+(rc.Bottom-rc.Top)div 2 -20;
  rc.Right:=rc.left+40;
  rc.Bottom:=rc.Top+40;
  if PtInRect(rc,pt) then
  caption:=Format('TTTOOOPP:XY:%d,%d',[pt.x,pt.y])
  else
  caption:=Format('mouse: X,Y:%d,%d',[pt.X,pt.Y]);
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  Application.Terminate;
end;

end.
