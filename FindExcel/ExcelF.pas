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
    ShowMessage('初始化失败，无法打开工作表，可能没有安装Xecel，或其他问题');
    v.displayalerts:=False;
    v.quit;
    Exit;
   end;
     v.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'工作表.xls');
     Sheet:=v.WorkBooks[1].WorkSheets[1];
   try
     range:=Sheet.cells.find('合同日期');
     //ShowMessage('找到的行：'+inttostr(range.row)+' 列：'+IntToStr(range.column));
     usrange:=Sheet.usedrange;
     if not VarIsEmpty(range)then
     begin
       findrow:=range.row;
       repeat
         m:=monthsbetween(Now,Sheet.cells[range.row,range.column+1].value);
         lbl1.Caption:=lbl1.Caption+'姓名：'+sheet.cells[range.row,range.column-3].value+#8+':'+
              DateToStr(Sheet.cells[range.row,range.column+1].value)+#13+#10+
             '时间过去：'+IntToStr(m)+'月'+#13;
         range:=usrange.findnext(range);
         //ShowMessage('下一处找到的行：'+inttostr(range.row)+' 列：'+IntToStr(range.column));
       until VarIsEmpty(range) or (findrow=range.row);
     end;
   except
      on e: Exception do
     begin ShowMessage(e.Message);
     //ShowMessage('查找命令失败');
     v.quit;
     Exit;
   end;
   end;
   v.quit;
end;

end.
