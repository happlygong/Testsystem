unit ExcelF;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ComObj, StdCtrls,DateUtils;

type
  TForm1 = class(TForm)
    btn1: TButton;
    edt1: TEdit;
    lbl1: TLabel;
    procedure btn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  v,range:   Variant;
  Sheet :   Variant;

implementation

{$R *.dfm}

procedure TForm1.btn1Click(Sender: TObject);
var
  usrange:Variant;
  findrow,m:Integer;
begin
   try
     v:=CreateoleObject('Excel.Application');
     v.visible:=False;
   except
    ShowMessage('��ʼ��ʧ�ܣ��޷��򿪹���������û�а�װXecel������������');
    v.displayalerts:=False;
    v.quit;
    Exit;
   end;
     v.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'������.xls');
     Sheet:=v.WorkBooks[1].WorkSheets[1];
   try
     range:=Sheet.cells.find('��ͬ����');
     //ShowMessage('�ҵ����У�'+inttostr(range.row)+' �У�'+IntToStr(range.column));
     usrange:=Sheet.usedrange;
     if not VarIsEmpty(range)then
     begin
       findrow:=range.row;
       repeat
         m:=monthsbetween(Now,Sheet.cells[range.row,range.column+1].value);
         lbl1.Caption:=lbl1.Caption+'������'+sheet.cells[range.row,range.column-3].value+#8+':'+
              DateToStr(Sheet.cells[range.row,range.column+1].value)+#13+#10+
             'ʱ���ȥ��'+IntToStr(m)+'��'+#13;
         range:=usrange.findnext(range);
         //ShowMessage('��һ���ҵ����У�'+inttostr(range.row)+' �У�'+IntToStr(range.column));
       until VarIsEmpty(range) or (findrow=range.row);
     end;
   except
      on e: Exception do
     begin ShowMessage(e.Message);
     //ShowMessage('��������ʧ��');
     v.quit;
     Exit;
   end;
   end;
   v.quit;
end;

end.
