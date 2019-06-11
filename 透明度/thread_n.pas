unit thread_n;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    RichEdit1: TRichEdit;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  hth:Thandle;
implementation

{$R *.dfm}
procedure th1;
var
 i:integer;
begin
  for i:=0 to 100 do
  with form1 do begin
  begin
   canvas.textout(0,0,inttostr(i));
   sleep(100);
  end;
  end;
end;
procedure th2;
begin
  with form1 do begin
  if (WaitForSingleObject(hth,INFINITE)=WAIT_OBJECT_0) then
  button1.Enabled:=true;
  canvas.textout(80,0,'�߳̽���.');
  richedit1.Visible:=true;
  end;
end;
procedure TForm1.Button1Click(Sender: TObject);
var
  tmp:longword;
  //hDC:HDC
  bmap:TBitmap;
begin
    richedit1.Visible:=false;
    canvas.textout(80,0,'                 ');
    button1.Enabled:=false;
    hth:=createthread(nil,0,@th1,nil,0,tmp);
    createthread(nil,0,@th2,nil,0,tmp);
    canvas.TextOut(20,0,'�߳̿�ʼ...');

Canvas.Handle := GetWindowDC(Self.Handle); {Get Device context for drawing}
//DrawButtonFace(Canvas, TitleButton, 1, bsAutoDetect, False, False, False);
bmap := TBitmap.Create;
bmap.LoadFromFile('d:\dll\400920.bmp');
Canvas.Draw(0, 0, bmap);

ReleaseDC(Self.Handle, Canvas.Handle);
bmap.Free;
Canvas.Handle := 0;

end;

end.
