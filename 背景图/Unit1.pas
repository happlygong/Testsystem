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
    //1.���巽��
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

//2.ʵ�ַ���

procedure TForm1.WMNCPaint(var Msg: TWMNCPaint);
var
  dc: hDc;
  Pen: hPen;
  OldPen: hPen;
  OldBrush: hBrush;
begin
  inherited;
  //��ȡ�������豸������
  dc := GetWindowDC(Handle);
  msg.Result := 1;
  //�������ʣ�ʵ�ߡ����Ϊl����ɫ
  Pen := CreatePen(PS_SOLID, 10, RGB(255, 0, 0));
  //���´����Ļ���ѡ�봰����豸������
  OldPen := SelectObject(dc, Pen);
  //��ϵͳ���Ŀջ�ˢ�봰����豸������
  OldBrush := SelectObject(dc, GetStockObject(NULL_BRUSH));
  //�����塰��ߡ�
  Rectangle(dc, 0, 0, Form1.Width, Form1.Height);
  //�ָ��ɻ��ʺ;ɻ�ˢ
  SelectObject(dc, OldBrush);
  SelectObject(dc, oldPen);
  //ɾ���´����Ļ��ʣ��ͷ�ϵͳ��Դ
  DeleteObject(Pen);
  //�ͷ��豸������
  ReleaseDC(Handle, Canvas.Handle);
end;

//3.DBGrid�ؼ����

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
  {��ȡϵͳ�˵����}
  appendmenu(hMenu, MF_SEPARATOR, 0, nil);
  appendmenu(hMenu, MF_STRING, 100, '����(&A)');
  AppendMenu(hMenu, MF_STRING, 101, '��æ(&H)');
  //   AppendMenu(hMenu,MF_POPUP or MF_STRING,PopupMenu1.Handle,'Sub Menu');
     {�����û��˵�}
  {   DeleteMenu(hMenu,0,MF_BYPOSITION); //ɾ��������һ���˵���
     DeleteMenu(hmenu,1,MF_BYPOSITION);
     DeleteMenu(hmenu,2,MF_BYPOSITION);
        }//DeleteMenu(hmenu,0,MF_BYPOSITION);
  Bmp := TBitmap.Create;
  Bmp.LoadFromFile('E:\������\rev\Dbox\�ʼ��ϲ�����Doc\bmp\��\1.bmp');
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
  Bmp.LoadFromFile('E:\������\rev\Dbox\�ʼ��ϲ�����Doc\bmp\��\2.bmp');
  //Bmp.LoadFromResourceID(0,seed);
  Panel1.Brush.Bitmap := Bmp; //Image1.Picture.Bitmap;
end;

procedure TForm1.Panel1Click(Sender: TObject);
var
  bmp: TBitmap;
begin
  Bmp := TBitmap.Create;
  Bmp.LoadFromFile('E:\������\rev\Dbox\�ʼ��ϲ�����Doc\bmp\��\2.bmp');
  //Bmp.LoadFromResourceID(0,seed);
  Brush.Bitmap := Bmp; //Image1.Picture.Bitmap;
  Self.Update;
end;

procedure TForm1.Memo1Enter(Sender: TObject);
var
  bmp: TBitmap;
begin
  Bmp := TBitmap.Create;
  Bmp.LoadFromFile('E:\������\rev\Dbox\�ʼ��ϲ�����Doc\bmp\��\3.bmp');
  //Bmp.LoadFromResourceID(0,seed);
  Memo1.Brush.Bitmap := Bmp; //Image1.Picture.Bitmap;
  memo1.Repaint;
end;

procedure TForm1.Edit1Enter(Sender: TObject);
var
  bmp: TBitmap;
begin
  Bmp := TBitmap.Create;
  Bmp.LoadFromFile('E:\������\rev\Dbox\�ʼ��ϲ�����Doc\bmp\��\3.bmp');
  //Bmp.LoadFromResourceID(0,seed);
  Memo1.Brush.Bitmap := Bmp; //Image1.Picture.Bitmap;
  Edit1.Repaint;
end;

procedure TForm1.user_sysmenu(var msg: TWMMENUSELECT);
begin
  self.Caption := IntToStr(msg.iditem) + ' ';
  inherited; { ��ȱʡ����,���������һ����}
end;

end.

