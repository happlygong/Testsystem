unit Covernttime;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, strutils;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    ODg1: TOpenDialog;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    CheckBox1: TCheckBox;
    Edit1: TEdit;
    Memo1: TMemo;
    Button2: TButton;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure FormPaint(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

  {$R *.dfm}
  //ת��ʱ���ʽ
function CovFileDate(Fd: _FileTime): TDateTime;
var
  Tct: _SystemTime;
  Temp: _FileTime;
begin
  FileTimeToLocalFileTime(Fd, Temp);
  FileTimeToSystemTime(Temp, Tct);
  CovFileDate := SystemTimeToDateTime(Tct);
end;

procedure GetFileTime(const Tf: string);
  { ��ȡ�ļ�ʱ�䣬Tf��ʾĿ���ļ�·�������� }
const
  //Model = 'yyyy/mm/dd/hh/mmss'; { �趨ʱ���ʽ Win7����ϵͳ/��ͬ-}
  Model = 'yyyy_mm_dd_hh_mmss'; { �趨ʱ���ʽ Win7������/��-��ͬ}
var
  Tp: TSearchRec; { ����TpΪһ�����Ҽ�¼ }
  T1, T2, T3: String;
begin
  FindFirst(Tf, faAnyFile, Tp); //����Ŀ���ļ�
  T1 := FormatDateTime(Model, CovFileDate(Tp.FindData.ftCreationTime));
  { �����ļ��Ĵ���ʱ�� }
  T2 := FormatDateTime(Model, CovFileDate(Tp.FindData.ftLastWriteTime));
  { �����ļ����޸�ʱ�� }
  T3 := FormatDateTime(Model, Now);
  { �����ļ��ĵ�ǰ����ʱ�� }
  FindClose(Tp);
  form1.label2.Caption := '�ļ�����ʱ��:' + T1;
  form1.label3.Caption := '�ļ��޸�ʱ��:' + T2;
  form1.Label4.Caption := '�ļ�����ʱ��:' + T3;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  b: Byte;
  tmpstr: String;
begin
  b := 0;
  if odg1.Execute then
    label1.Caption := odg1.FileName;
  label5.Caption := expandfilename(odg1.FileName);
  label6.Caption := ExtractFileName(odg1.FileName);
  label7.Caption := extractfilepath(odg1.filename);
  tmpstr := odg1.FileName;
  tmpstr := AnsiReverseString(tmpstr);
  tmpstr := Copy(tmpstr, Pos('.', tmpstr) + 1, Length(tmpstr));
  tmpstr := AnsiReverseString(tmpstr);
  label8.Caption := tmpstr;
  GetFileTime(odg1.FileName);
end;

procedure TForm1.FormPaint(Sender: TObject);
begin
  MessageBox(0, '������', 'OnPaint�¼�����', mb_ok);
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  T: TSystemTime;
  s: string;
begin
  GetLocalTime(T);
  s := Format('%d-', [T.wYear]);
  s := s + Format('%.2d-', [T.wMonth]);
  s := s + Format('%.2d-', [T.wDay]);
  s := s + Format('%.2d-', [T.wHour]);
  s := s + Format('%.2d-', [T.wMinute]);
  s := s + Format('%.2d', [T.wSecond]);
  Edit1.text := s;
end;

procedure TForm1.Edit1KeyPress(Sender: TObject; var Key: Char);
begin
  if not (Key in ['0'..'9']) then
    Key := #0;
end;

end.

