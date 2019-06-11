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

//初始化创建水波图像组件于父句柄
//HWD		创建水波效果到该句柄上
//DrawTextSize	初始化在水波效果上的文本数量, 如不使用设为0
//DrawBmpSize	初始化在水波效果上的位图数量, 如不使用设为0
function WaterInit(HWD: HWND; DrawTextSize, DrawBmpSize: Integer):Integer;stdcall;
external 'WaterLib.dll'; 

//设置水波图像组件大小与位置
Procedure WaterSetBounds(Handle, Left, Top, Width, Height: Integer); stdcall;
external 'WaterLib.dll';

//ClickBlob	点击画面时产生的水滴效果幅度，0 表示禁用
//Damping	水滴阻尼系数
//RandomBlob	随机产生的水滴最大幅度，0 表示禁用
//RandomDelay	随机产生水滴的延时
//TrackBlob	鼠标移动轨迹下水滴的幅度，0 表示禁用
Procedure WaterSetType(Handle, ClickBlob, Damping, RandomBlob, RandomDelay, TrackBlob: Integer);stdcall;
external 'WaterLib.dll';

//设置水波图像组件图片文件
Procedure WaterSetFile(Handle: Integer; FileName: AnsiString); stdcall;
external 'WaterLib.dll';

//设置水波图像组件是否启动水波效果
Procedure WaterSetActive(Handle: Integer; Active: Boolean); stdcall;
external 'WaterLib.dll';

//设置水波图像组件父句柄
Procedure WaterSetParentWindow(Handle: Integer; HWND: HWND); stdcall;
external 'WaterLib.dll';

//清空画面上的水滴效果
Procedure WaterClear(Handle: Integer);stdcall;
external 'WaterLib.dll';

//释放指定句柄的水波图像组件
Procedure WaterFree(Handle: Integer);stdcall;
external 'WaterLib.dll';

//释放全部水波图像组件
Procedure WaterAllFree;stdcall;
external 'WaterLib.dll';

//设置是否支持作者, True支持则显示SkyGz.Com标识, False不支持则关闭标识显示
Procedure WaterSupportAuthor(SupportAuthor: Boolean);stdcall;
external 'WaterLib.dll';

//以下API需要启动支持作者的标识显示时才可使用
//主要功能作用是在水波效果图上画文本和位图
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
  //支持本人作品, 建议开启标识显示.
  //只有开启该功能后, 才可以使用WaterDraw****相关API的功能

  F:= 'c:\temp\WizardImage.bmp';
  Form1.Image1.Picture.SaveToFile(F);
  WaterHandle := WaterInit(Form1.Handle, 2, 2);
  WaterSetBounds(WaterHandle, Form1.Image1.Left, Form1.Image1.Top, Form1.Image1.Width, Form1.Image1.Height);
  WaterSetFile(WaterHandle, AnsiString(F));
  WaterSetType(WaterHandle, 100, 1, 100, 100, 200 );
//ClickBlob	点击画面时产生的水滴效果幅度，0 表示禁用
//Damping	水滴阻尼系数
//RandomBlob	随机产生的水滴最大幅度，0 表示禁用
//RandomDelay	随机产生水滴的延时
//TrackBlob	鼠标移动轨迹下水滴的幅度，0 表示禁用
  WaterSetActive(WaterHandle, True);
  DeleteFile(F);
end;

procedure CurPageChanged(CurPageID: Integer);
begin
{  Case CurPageID of
    wpWelcome : WaterSetParentWindow(WaterHandle, WizardForm.WelcomePage.Handle);  //将水波移动到另一个句柄上
    wpFinished: WaterSetParentWindow(WaterHandle, WizardForm.FinishedPage.Handle); //将水波移动到另一个句柄上
  end;
  }
end;

//释放所有水波对象
procedure DeinitializeSetup();
begin
  WaterAllFree;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  WaterAllFree;
end;

end.
