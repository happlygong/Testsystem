unit trayUnit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ShellAPI, Menus;

const
  WM_TRAYNOTIFY = 3000;
type
  TForm1 = class(TForm)
    Button1: TButton;
    PopupMenu1: TPopupMenu;
    Button2: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure N11Click(Sender: TObject);
    procedure n21Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);

  private
    { Private declarations }
  public
    MINIFTPFF: TNotifyIconData;
    MsgTaskbarRestart: Cardinal;
    procedure WndProc(var Msg: TMessage); override;
    procedure MiniToTaskbar; //��С��������
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

{ TForm1 }

procedure TForm1.MiniToTaskbar;
begin
  MsgTaskbarRestart := RegisterWindowMessage('TaskbarCreated');
  //ShowWindow(Application.Handle, SW_SHOWMINIMIZED);
  //ShowWindow(Application.Handle, SW_HIDE);
  Form1.Hide;
  MINIFTPFF.cbSize := SizeOf(TNotifyIconData);
  MINIFTPFF.Wnd := Handle;
  MINIFTPFF.uID := 1985629;
  MINIFTPFF.uFlags := NIF_MESSAGE or NIF_ICON or NIF_TIP;
  MINIFTPFF.uCallbackMessage := WM_TRAYNOTIFY;
  MINIFTPFF.hIcon := Application.Icon.Handle;
  MINIFTPFF.szTip := '���Ǹ����Գ���';
  Shell_NotifyIcon(NIM_ADD, @MINIFTPFF);
end;

procedure TForm1.WndProc(var Msg: TMessage);
var
  mouse:TPoint;
begin
  inherited;
  if Msg.LParam = WM_LBUTTONDBLCLK then
  begin //˫��ͼƬͼ�껹ԭ����
    //ShowWindow(Application.Handle, SW_SHOW);
    //ShowWindow(Application.Handle, SW_SHOWNORMAL);
    Form1.Show;
    Application.BringToFront;
    Shell_NotifyIcon(NIM_Delete, @MINIFTPFF);
  end
  else if Msg.LParam = WM_RBUTTONUP then
  begin
    SetForegroundWindow(Handle);
    GetCursorPos(mouse);//����������
    PopupMenu1.Popup(mouse.X, mouse.Y);//������괦��ʾ�����˵�
  {end
  else if msg.lParam = WM_LBUTTONUP then
  begin
    SetForegroundWindow(Handle);
    GetCursorPos(mouse);//����������
    PopupMenu1.Popup(mouse.X, mouse.Y);//������괦��ʾ�����˵�
}
  end;
// if (Msg.Msg = WM_NCACTIVATE) then
//  PopupForm.Destroy;
// ��Ӧ�����������ؽ���Ϣ
  if Msg.Msg = MsgTaskbarRestart then
    MiniToTaskbar; // ��С��������
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  MiniToTaskbar;
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  MiniToTaskbar;
  canclose:=False;
end;

procedure TForm1.N11Click(Sender: TObject);
begin
    //ShowWindow(Application.Handle, SW_SHOW);
    //ShowWindow(Application.Handle, SW_SHOWNORMAL);
    Form1.Show;
    Application.BringToFront;
    Shell_NotifyIcon(NIM_Delete, @MINIFTPFF);

end;

procedure TForm1.n21Click(Sender: TObject);
begin
   Shell_NotifyIcon(NIM_DELETE,@MINIFTPFF);
   PostQuitMessage(0);
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
application.Terminate;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
MiniToTaskbar;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  MsgTaskbarRestart := RegisterWindowMessage('TaskbarCreated');
  //ShowWindow(Application.Handle, SW_SHOWMINIMIZED);
  //ShowWindow(Application.Handle, SW_HIDE);
  MINIFTPFF.cbSize := SizeOf(TNotifyIconData);
  MINIFTPFF.Wnd := Handle;
  MINIFTPFF.uID := 1985629;
  MINIFTPFF.uFlags := NIF_MESSAGE or NIF_ICON or NIF_TIP;
  MINIFTPFF.uCallbackMessage := WM_TRAYNOTIFY;
  MINIFTPFF.hIcon := Application.Icon.Handle;
  MINIFTPFF.szTip := '��������ʱ��С�����Գ���';
  Shell_NotifyIcon(NIM_ADD, @MINIFTPFF);
end;

end.

