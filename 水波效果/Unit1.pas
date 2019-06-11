unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, jpeg, StdCtrls;

type
  TForm1 = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Button1: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  WaterHandle:Integer;
implementation

{$R *.dfm}

//��ʼ������ˮ��ͼ������ڸ����
//HWD		����ˮ��Ч�����þ����
//DrawTextSize	��ʼ����ˮ��Ч���ϵ��ı�����, �粻ʹ����Ϊ0
//DrawBmpSize	��ʼ����ˮ��Ч���ϵ�λͼ����, �粻ʹ����Ϊ0
function WaterInit(HWD: HWND; DrawTextSize, DrawBmpSize: Integer):Integer;stdcall;
external 'WaterLib.dll'; 

//����ˮ��ͼ�������С��λ��
Procedure WaterSetBounds(Handle, Left, Top, Width, Height: Integer); stdcall;
external 'WaterLib.dll';

//ClickBlob	�������ʱ������ˮ��Ч�����ȣ�0 ��ʾ����
//Damping	ˮ������ϵ��
//RandomBlob	���������ˮ�������ȣ�0 ��ʾ����
//RandomDelay	�������ˮ�ε���ʱ
//TrackBlob	����ƶ��켣��ˮ�εķ��ȣ�0 ��ʾ����
Procedure WaterSetType(Handle, ClickBlob, Damping, RandomBlob, RandomDelay, TrackBlob: Integer);stdcall;
external 'WaterLib.dll';

//����ˮ��ͼ�����ͼƬ�ļ�
Procedure WaterSetFile(Handle: Integer; FileName: AnsiString); stdcall;
external 'WaterLib.dll';

//����ˮ��ͼ������Ƿ�����ˮ��Ч��
Procedure WaterSetActive(Handle: Integer; Active: Boolean); stdcall;
external 'WaterLib.dll';

//����ˮ��ͼ����������
Procedure WaterSetParentWindow(Handle: Integer; HWND: HWND); stdcall;
external 'WaterLib.dll';

//��ջ����ϵ�ˮ��Ч��
Procedure WaterClear(Handle: Integer);stdcall;
external 'WaterLib.dll';

//�ͷ�ָ�������ˮ��ͼ�����
Procedure WaterFree(Handle: Integer);stdcall;
external 'WaterLib.dll';

//�ͷ�ȫ��ˮ��ͼ�����
Procedure WaterAllFree;stdcall;
external 'WaterLib.dll';

//�����Ƿ�֧������, True֧������ʾSkyGz.Com��ʶ, False��֧����رձ�ʶ��ʾ
Procedure WaterSupportAuthor(SupportAuthor: Boolean);stdcall;
external 'WaterLib.dll';

//����API��Ҫ����֧�����ߵı�ʶ��ʾʱ�ſ�ʹ��
//��Ҫ������������ˮ��Ч��ͼ�ϻ��ı���λͼ
procedure WaterDrawTextBrush(Handle, Index: Integer; BrushColor: TColor; BrushStyle: TBrushStyle);stdcall;
external 'WaterDrawTextBrush@files:WaterLib.dll'; 

Procedure WaterDrawTextFont(Handle, Index: Integer; FontName: AnsiString; FontSize: Integer; FontColor: TColor; FontCharset: Byte);stdcall;
external 'WaterDrawTextFont@files:WaterLib.dll'; 

procedure WaterDrawBmpBrush(Handle, Index: Integer; BrushColor: TColor; BrushStyle: TBrushStyle);stdcall;
external 'WaterDrawBmpBrush@files:WaterLib.dll'; 

Procedure WaterDrawBmpFont(Handle, Index: Integer; FontName: AnsiString; FontSize: Integer; FontColor: TColor; FontCharset: Byte);stdcall;
external 'WaterDrawBmpFont@files:WaterLib.dll'; 

Procedure WaterDrawBitmap(Handle, Index: Integer; X, Y: Integer; HBitmap: LongWord; Transparent: Boolean; TransparentColor: TColor);stdcall;
external 'WaterDrawBitmap@files:WaterLib.dll'; 

Procedure WaterDrawText(Handle, Index: Integer; X, Y: Integer; S: AnsiString); stdcall;
external 'WaterDrawText@files:WaterLib.dll'; 
procedure TForm1.FormCreate(Sender: TObject);
var
  F: AnsiString; 
begin
  WaterSupportAuthor(False);
  //֧�ֱ�����Ʒ, ���鿪����ʶ��ʾ.
  //ֻ�п����ù��ܺ�, �ſ���ʹ��WaterDraw****���API�Ĺ���

  F:= 'c:\temp\WizardImage.bmp';
  Form1.Image1.Picture.SaveToFile(F);
  WaterHandle := WaterInit(Form1.Handle, 2, 2);
  WaterSetBounds(WaterHandle, Form1.Image1.Left, Form1.Image1.Top, Form1.Image1.Width, Form1.Image1.Height);
  WaterSetFile(WaterHandle, AnsiString(F));
  WaterSetType(WaterHandle, 100, 1, 100, 100, 200 );
//ClickBlob	�������ʱ������ˮ��Ч�����ȣ�0 ��ʾ����
//Damping	ˮ������ϵ��
//RandomBlob	���������ˮ�������ȣ�0 ��ʾ����
//RandomDelay	�������ˮ�ε���ʱ
//TrackBlob	����ƶ��켣��ˮ�εķ��ȣ�0 ��ʾ����
  WaterSetActive(WaterHandle, True);
  DeleteFile(F);
end;

procedure CurPageChanged(CurPageID: Integer);
begin
{  Case CurPageID of
    wpWelcome : WaterSetParentWindow(WaterHandle, WizardForm.WelcomePage.Handle);  //��ˮ���ƶ�����һ�������
    wpFinished: WaterSetParentWindow(WaterHandle, WizardForm.FinishedPage.Handle); //��ˮ���ƶ�����һ�������
  end;
  }
end;

//�ͷ�����ˮ������
procedure DeinitializeSetup();
begin
  WaterAllFree;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  WaterAllFree;
end;

end.
