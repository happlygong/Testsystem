unit MenuUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus;

type
  TForm1 = class(TForm)
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
  private
    { Private declarations }
    procedure  user_sysmenu(var msg:twmmenuselect);
    message wm_syscommand;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
var
 hMenu:DWORD;
 bCir:hBitmap;
begin
   hMenu:=getsystemmenu(handle,false);
   {��ȡϵͳ�˵����}
   appendmenu(hMenu,MF_SEPARATOR,0,nil);
   appendmenu(hMenu,MF_STRING,100,'����(&A)');
   AppendMenu(hMenu,MF_STRING,101,'��æ(&H)');
   AppendMenu(hMenu,MF_POPUP or MF_STRING,PopupMenu1.Handle,'Sub Menu');
   {�����û��˵�}
   //DeleteMenu(hMenu,0,MF_BYPOSITION); //ɾ��������һ���˵���
{   DeleteMenu(hmenu,1,MF_BYPOSITION);
   DeleteMenu(hmenu,2,MF_BYPOSITION);
}   //DeleteMenu(hmenu,0,MF_BYPOSITION);
    bCir:=LoadImage(HInstance,'back2.bmp',IMAGE_BITMAP,0,0,LR_LOADFROMFILE or LR_CREATEDIBSECTION);
    SetMenuItemBitmaps(hMenu,7,MF_BYPOSITION,bCir,bCir);//���˵����
    SetMenuItemBitmaps(hMenu,101,MF_BYCOMMAND,bCir,bCir);//���˵�ID
    SetMenuItemBitmaps(hMenu,PopupMenu1.Handle,MF_BYCOMMAND,bCir,bCir);//���˵�ID
end;

procedure  TForm1.user_sysmenu(var msg:TWMMENUSELECT);
begin
   self.Caption:=IntToStr(msg.iditem)+' ';
   if msg.iditem=100 then
      begin MessageBox(0,''+
      #13+#13'2011��11��11��',
      '���ڱ�����',MB_OK or MB_ICONASTERISK);

      { Ҳ����setwindowpos()��ʵ�ִ�����ǰ�˹���}
      end
   else if msg.IDItem=GetMenuitemid(PopupMenu1.Handle,0) then
      N1Click(self)
   else  if msg.IDItem=GetMenuitemid(PopupMenu1.Handle,1) then
      N2Click(self)
   else if msg.IDItem=GetMenuitemid(PopupMenu1.Handle,2) then
      N3Click(self)
   else
      inherited;     { ��ȱʡ����,���������һ����}
end;
procedure TForm1.N1Click(Sender: TObject);
begin
Self.Caption:='�ϲ�һ';
//Self.Caption:=n1.Items[0].Name;
end;

procedure TForm1.N2Click(Sender: TObject);
begin
  Self.Caption:='�ϲ˶�';
  showmessage('�����Ӳ˵�����������');
end;

procedure TForm1.N3Click(Sender: TObject);
begin
Self.Caption:='�ϲ���';
end;

end.
