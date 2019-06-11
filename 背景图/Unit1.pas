unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Menus;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    Memo1: TMemo;
    Panel1: TPanel;
    procedure FormCreate(Sender: TObject);
    procedure Panel1Resize(Sender: TObject);
    procedure Panel1Click(Sender: TObject);
    procedure Memo1Enter(Sender: TObject);
    procedure Edit1Enter(Sender: TObject);
    //1.定义方法
    procedure USER_SYSMENU(var Msg: TWMMenuSelect); message WM_SYSCommand;
    procedure WMNCPaint(var Msg: TWMNCPaint); message WM_NCPAINT;
    procedure FormPaint(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

//2.实现方法

procedure TForm1.WMNCPaint(var Msg: TWMNCPaint);
var
  dc: hDc;
  Pen: hPen;
  OldPen: hPen;
  OldBrush: hBrush;
begin
  inherited;
  //获取本窗口设备上下文
  dc := GetWindowDC(Handle);
  msg.Result := 1;
  //创建画笔，实线、宽度为l、红色
  Pen := CreatePen(PS_SOLID, 10, RGB(255, 0, 0));
  //将新创建的画笔选入窗体的设备上下文
  OldPen := SelectObject(dc, Pen);
  //将系统库存的空画刷入窗体的设备上下文
  OldBrush := SelectObject(dc, GetStockObject(NULL_BRUSH));
  //给窗体“镶边”
  Rectangle(dc, 0, 0, Form1.Width, Form1.Height);
  //恢复旧画笔和旧画刷
  SelectObject(dc, OldBrush);
  SelectObject(dc, oldPen);
  //删除新创建的画笔，释放系统资源
  DeleteObject(Pen);
  //释放设备上下文
  ReleaseDC(Handle, Canvas.Handle);
end;

//3.DBGrid控件描边

procedure TForm1.FormPaint(Sender: TObject);
var
  Rct: TRect;
begin
  Rct := Rect(Memo1.Left, Memo1.Top, Memo1.Left + Memo1.Width, Memo1.top +
    Memo1.Height);
  with Form1.Canvas do
  begin
    Pen.Color := clRed;
    Pen.Width := 5;
    Brush.Style := bsClear;
    Rectangle(Rct);
  end;
  Rct := Rect(Button1.Left, Button1.Top, Button1.Left + Button1.Width,
    Button1.top + Button1.Height);
  with Form1.Canvas do
  begin
    Pen.Color := clBlue;
    Pen.Width := 5;
    Brush.Style := bsClear;
    Rectangle(Rct);
  end;
  Rct := Rect(0, 0, Form1.Width - 18, Form1.Height - 40);
  with Form1.Canvas do
  begin
    Pen.Color := clYellow;
    Pen.Width := 5;
    Brush.Style := bsClear;
    Rectangle(Rct);
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  bmp: TBitmap;
  hMenu: DWORD;
begin
  hMenu := getsystemmenu(handle, false);
  {获取系统菜单句柄}
  appendmenu(hMenu, MF_SEPARATOR, 0, nil);
  appendmenu(hMenu, MF_STRING, 100, '关于(&A)');
  AppendMenu(hMenu, MF_STRING, 101, '帮忙(&H)');
  //   AppendMenu(hMenu,MF_POPUP or MF_STRING,PopupMenu1.Handle,'Sub Menu');
     {加入用户菜单}
  {   DeleteMenu(hMenu,0,MF_BYPOSITION); //删掉最上面一个菜单项
     DeleteMenu(hmenu,1,MF_BYPOSITION);
     DeleteMenu(hmenu,2,MF_BYPOSITION);
        }//DeleteMenu(hmenu,0,MF_BYPOSITION);
  Bmp := TBitmap.Create;
  Bmp.LoadFromFile('E:\范二朋\rev\Dbox\邮件合并生成Doc\bmp\布\1.bmp');
  //Bmp.LoadFromResourceID(0,seed);
  Brush.Bitmap := Bmp; //Image1.Picture.Bitmap;
  Panel1.repaint;
  Memo1.Repaint;
  Edit1.Repaint;
  Button1.Repaint;

end;

procedure TForm1.Panel1Resize(Sender: TObject);
var
  bmp: TBitmap;
begin
  Bmp := TBitmap.Create;
  Bmp.LoadFromFile('E:\范二朋\rev\Dbox\邮件合并生成Doc\bmp\布\2.bmp');
  //Bmp.LoadFromResourceID(0,seed);
  Panel1.Brush.Bitmap := Bmp; //Image1.Picture.Bitmap;
end;

procedure TForm1.Panel1Click(Sender: TObject);
var
  bmp: TBitmap;
begin
  Bmp := TBitmap.Create;
  Bmp.LoadFromFile('E:\范二朋\rev\Dbox\邮件合并生成Doc\bmp\布\2.bmp');
  //Bmp.LoadFromResourceID(0,seed);
  Brush.Bitmap := Bmp; //Image1.Picture.Bitmap;
  Self.Update;
end;

procedure TForm1.Memo1Enter(Sender: TObject);
var
  bmp: TBitmap;
begin
  Bmp := TBitmap.Create;
  Bmp.LoadFromFile('E:\范二朋\rev\Dbox\邮件合并生成Doc\bmp\布\3.bmp');
  //Bmp.LoadFromResourceID(0,seed);
  Memo1.Brush.Bitmap := Bmp; //Image1.Picture.Bitmap;
  memo1.Repaint;
end;

procedure TForm1.Edit1Enter(Sender: TObject);
var
  bmp: TBitmap;
begin
  Bmp := TBitmap.Create;
  Bmp.LoadFromFile('E:\范二朋\rev\Dbox\邮件合并生成Doc\bmp\布\3.bmp');
  //Bmp.LoadFromResourceID(0,seed);
  Memo1.Brush.Bitmap := Bmp; //Image1.Picture.Bitmap;
  Edit1.Repaint;
end;

procedure TForm1.user_sysmenu(var msg: TWMMENUSELECT);
begin
  self.Caption := IntToStr(msg.iditem) + ' ';
  inherited; { 作缺省处理,必须调用这一过程}
end;

end.

