unit DB2Ex;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,ComObj,ADODB,StrUtils, ExtCtrls,
  MMSystem;

type
  TForm1 = class(TForm)
    btn1: TButton;
    bevel1:tBevel;
    label1:Tlabel;
    Timer1: TTimer;
    Button1: TButton;
    grp1: TGroupBox;
    chk1: TCheckBox;
    chk2: TCheckBox;
    chk3: TCheckBox;
    chk4: TCheckBox;
    chk5: TCheckBox;
    chk6: TCheckBox;
    chk7: TCheckBox;
    chk8: TCheckBox;
    chk9: TCheckBox;
    btn3: TButton;
    btn4: TButton;
    btn5: TButton;
    btn6: TButton;
    btn7: TButton;
    btn8: TButton;
    rg1: TRadioGroup;
    btn2: TButton;
    btn9: TButton;
    chk10: TCheckBox;
    procedure btn1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Button1Click(Sender: TObject);   
    procedure btn3Click(Sender: TObject);
    procedure btn4Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btn5Click(Sender: TObject);
    procedure btn6Click(Sender: TObject);
    procedure btn7Click(Sender: TObject);
    procedure btn8Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure btn9Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;
const
  MAX=10;
var
  Form1 :   TForm1;
  v,range:   Variant;
  Sheet :   Variant;
  sum   :   array[1..4,1..MAX] of Double;
  str:string;
  inplay:Boolean=true;
  is3ji:Boolean=False;
  is2c:boolean=False;
  TableName :array[0..3] of string=('二级C语言一级考点表','二级C语言二级考点表','二级C语言三级考点表','二级C语言四级考点表');
//  TableName :array[0..3] of string=('二级Access一级考点表','二级Access二级考点表','二级Access三级考点表','二级Access四级考点表');
//  TableName :array[0..3] of string=('二级VB一级考点表','二级VB二级考点表','二级VB三级考点表','二级VB四级考点表');
//  TableName :array[0..3] of string=('二级VF一级考点表','二级VF二级考点表','二级VF三级考点表','二级VF四级考点表');
//  TableName :array[0..3] of string=('三级网络技术一级考点表','三级网络技术二级考点表','三级网络技术三级考点表','三级网络技术四级考点表');
  table:string;
implementation
uses DBe;
{$R *.dfm}
//数据表是否存在
Function TableExist(pConn:TADOConnection; pcTable:string ):boolean;overload;
var
  tmpFldList : TStrings ;
  nLoop : integer ;
begin
    Result := False ;
    tmpFldList := TStringList.Create ;
    pConn.GetTableNames( tmpFldList, True ); // 包含系统表
    for nLoop := 0 to tmpFldList.Count - 1 do
    begin
        if uppercase( tmpFldList[nLoop] ) = uppercase( pcTable ) then
        begin
            Result := True ;
            break ;
        end;
    end;
    tmpFldList.Free ;
end;

////比较考点////
procedure getsum(level:Integer;str:string;table:string);
var
  i,th,fz,cou:Integer;   bili:Double;
  ti:tadoquery;
  kaodian:string;
  begin
    ti:=Tadoquery.create(nil);
    TI.connection:=DBe.dm.con1;
    ti.SQL.Clear;
    ti.SQL.Add('select TH,FZ,考点1,考点1比例,考点2,考点2比例,考点3,考点3比例,考点4,考点4比例,考点5,考点5比例,考点6,考点6比例 from '+table);ti.open;
    cou:=ti.RecordCount;
    for i:=1 to cou do
    begin
      kaodian:=ti.fieldbyname('考点1').AsString;
      if kaodian=str then
        begin
         th:=ti.fieldbyname('TH').AsInteger;
         fz:=ti.fieldbyname('FZ').AsInteger;
         bili:=ti.fieldbyname('考点1比例').AsFloat;
         sum[level,th]:=sum[level,th]+fz*bili;
        end;
      kaodian:=ti.fieldbyname('考点2').AsString;
      if kaodian='' then begin ti.Next;Continue;end;
      if kaodian=str then
        begin
         th:=ti.fieldbyname('TH').AsInteger;
         fz:=ti.fieldbyname('FZ').AsInteger;
         bili:=ti.fieldbyname('考点2比例').AsFloat;
         sum[level,th]:=sum[level,th]+fz*bili;
        end;
      kaodian:=ti.fieldbyname('考点3').AsString;
      if kaodian='' then begin ti.Next;Continue;end;
      if kaodian=str then
        begin
         th:=ti.fieldbyname('TH').AsInteger;
         fz:=ti.fieldbyname('FZ').AsInteger;
         bili:=ti.fieldbyname('考点3比例').AsFloat;
         sum[level,th]:=sum[level,th]+fz*bili;
        end;
      kaodian:=ti.fieldbyname('考点4').AsString;
      if kaodian='' then begin ti.Next;Continue;end;
      if kaodian=str then
        begin
         th:=ti.fieldbyname('TH').AsInteger;
         fz:=ti.fieldbyname('FZ').AsInteger;
         bili:=ti.fieldbyname('考点4比例').AsFloat;
         sum[level,th]:=sum[level,th]+fz*bili;
        end;
      kaodian:=ti.fieldbyname('考点5').AsString;
      if kaodian='' then begin ti.Next;Continue;end;
      if kaodian=str then
        begin
         th:=ti.fieldbyname('TH').AsInteger;
         fz:=ti.fieldbyname('FZ').AsInteger;
         bili:=ti.fieldbyname('考点5比例').AsFloat;
         sum[level,th]:=sum[level,th]+fz*bili;
        end;
      kaodian:=ti.fieldbyname('考点6').AsString;
      if kaodian='' then begin ti.Next;Continue;end;
      if kaodian=str then
        begin
         th:=ti.fieldbyname('TH').AsInteger;
         fz:=ti.fieldbyname('FZ').AsInteger;
         bili:=ti.fieldbyname('考点6比例').AsFloat;
         sum[level,th]:=sum[level,th]+fz*bili;
        end;
      ti.Next;
    end;
    ti.Free;
  end;
//生成单个表不除100
procedure TForm1.btn1Click(Sender: TObject);
var
row1,row2,row3,crow,i,j,k,l,m:Integer;
big,litle:Double;
aq1,aq2,aq3,aq4:TADOQuery;//4个表的查询
cou1,cou2,cou3,cou4:integer;//4个表的记录个数
sumer:array[1..4,1..MAX]of Double;
totle:array[1..MAX] of double;
field,field1,field2,field3,field4:string;
begin
    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('数据表:'''+TableName[1]+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    if rg1.itemindex=0 then table:='题目表'  else table:='模拟题表';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('数据表:'''+Table+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    try
      v:=CreateOleObject('Excel.Application');//如果这个字符串中有空格就会出'无效的类别字符串'的错
      v.Visible:=True;
                v.WorkBooks.Add;   //新建Excel   文件
                //v.WorkBooks[1].WorkSheets[1].name:= '电脑报';//第一页标题
                //v.WorkBooks[1].WorkSheets[2].name:= '编程乐园';
                //v.WorkBooks[1].WorkSheets[3].name:= '都来看呀 ';
                v.activesheet.columns[1].columnwidth:=20;
                //v.Activesheet.usedrange.font.size:=8;
                Sheet:=v.WorkBooks[1].WorkSheets[1];
                v.workBooks[1].Styles['Normal'].Font.size:=10;
{                range:=Sheet.cells.select;
                v.selection.font.size:=10;
                v.Selection.FormatConditions.Delete;
                v.Selection.FormatConditions.Add(1, 3,'0');
                v.Selection.FormatConditions[1].Font.ColorIndex := 15;
}
                v.range['a2'].HorizontalAlignment := 3;
                sheet.cells[2,1]:='合计';
                v.range['b:k'].FormatConditions.Delete;
                v.range['b:k'].FormatConditions.Add(1, 3,'0');
                v.range['b:k'].FormatConditions[1].Font.ColorIndex := 15;
                if is3ji then begin
                  v.range['b2:k2'].FormatConditions.Add(1,1,'0','99.995');
                  v.range['b2:k2'].FormatConditions[2].Font.ColorIndex := 3;
                  v.range['b2:k2'].FormatConditions.Add(1,7,'100.005');end
                else begin
                 v.range['b2:k2'].FormatConditions.Add(1,1,'0','69.995');
                 v.range['b2:k2'].FormatConditions[2].Font.ColorIndex := 3;
                 v.range['b2:k2'].FormatConditions.Add(1,7,'70.005');
                end;
                v.range['b2:k2'].FormatConditions[3].Font.ColorIndex := 3;
                v.range['b:k'].columnwidth:=7.5;
                sheet.range['b3'].select;
                v.ActiveWindow.FreezePanes:= True;
                sheet.range['a1'].select;
                //Sheet.Cells[1,1]:= '好看 ';//单元格内容
                //Sheet.Cells[1,2]:= '确实 ';
                //Sheet.Cells[2,1]:= '我喜欢 ';
{                v.activesheet.rows[2].select;
                v.Selection.Interior.ColorIndex:= 32;      
                v.range['b2:j2'].interior.colorindex:=39;//粉色
 }
    except
        showmessage( '初始化Excel失败，可能没装Excel，或者其他错误；请重新再试。 ');
        v.DisplayAlerts:=false;//是否提示存盘
        v.Quit;//退出Excel
        exit;
    end;
    {Application.Restore;
    Application.BringToFront;}
    field1:=' 一级考点名称 ';
    field2:=' 二级考点名称 ';
    field3:=' 三级考点名称 ';
    field4:=' 四级考点名称 ';
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+' from '+TableName[0]+'  order by 一级考点ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
 btn1.Enabled:=false;
 with Form1 do
 begin
   Sheet.Cells[1,2]:='1';
   Sheet.Cells[1,3]:='2';
   Sheet.Cells[1,4]:='3';
   Sheet.Cells[1,5]:='4';
   Sheet.Cells[1,6]:='5';
   Sheet.Cells[1,7]:='6';
   Sheet.Cells[1,8]:='7';
   Sheet.Cells[1,9]:='8';   Sheet.Cells[1,10]:='9';
   Sheet.Cells[1,11]:='10';
  crow:=2;for i:=1 to MAX do totle[i]:=0;
  for i:=1 to cou1 do
  begin
    caption:='第'+inttostr(i)+'章 处理中...';
    for m:=1 to MAX do sum[1,m]:=0;
    crow:=crow+1;//inc(row);
    row1:=crow;
    v.range['a'+inttostr(crow)+':k'+inttostr(crow)].interior.colorindex:=24;
    str:= aq1.fieldbyname('一级考点名称').asstring;
    Sheet.cells[crow,1]:=' '+aq1.fieldbyname('一级考点名称').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);
    for j:=1 to MAX do
      Sheet.cells[crow,j+1]:=sum[1,j];
    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+TableName[1]+' where [二级考点所属一级考点ID]='+inttostr(i)+' order by 二级考点ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     crow:=crow+1;    row2:=crow;
     Sheet.cells[crow,1]:='     '+aq2.fieldbyname('二级考点名称').asstring;
     str:=aq2.fieldbyname('二级考点名称').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);
     for k:=1 to MAX do Sheet.cells[crow,k+1]:=sum[2,k];
     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+TableName[2]+' where ([三级考点所属一级考点ID]='+inttostr(i)+
                ') and ([三级考点所属二级考点ID]='+inttostr(j)+') order by 三级考点ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
       crow:=crow+1;    row3:=crow;
       Sheet.cells[crow,1]:='         '+aq3.fieldbyname('三级考点名称').asstring;
       str:=aq3.fieldbyname('三级考点名称').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+TableName[3]+' where ([四级考点所属一级考点ID]='+inttostr(i)+
                ') and ([四级考点所属二级考点ID]='+inttostr(j)+
                ') and ([四级考点所属三级考点ID]='+inttostr(k)+
                ') order by 四级考点ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
        crow:=crow+1;
        Sheet.cells[crow,1]:='             '+aq4.fieldbyname('四级考点名称').asstring;
        str:=aq4.fieldbyname('四级考点名称').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);
        for m:=1 to MAX do begin
          Sheet.cells[crow,m+1]:=sum[4,m];
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          end;
        aq4.Next;
       end;//第4层结束，第3层的值是第3层的值加所有第4层的和
       for l:=1 to MAX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          Sheet.cells[row3,l+1]:=sum[3,l];
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          end;
       aq3.Next;
     end; //第3层结束
     for l:=1 to MAX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          Sheet.cells[row2,l+1]:=sum[2,l];
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          end;
     aq2.Next;
    end;
    for l:=1 to MAX do begin
          sum[1,l]:=sum[1,l]+sumer[2,l];
          Sheet.cells[row1,l+1]:=sum[1,l];
          totle[l]:=totle[l]+sum[1,l];
          Sheet.cells[2,l+1]:=totle[l];
          end;
   aq1.Next;
   end;
    aq4.Free;
     aq3.free;
     aq2.Free;
     aq1.Free;
 end;
   caption:='完成';
   if is3ji then begin big:=100.005;litle:=99.995;end
     else begin big:=70.005;litle:=69.995;end;
   for i:=1 to MAX do
   begin
     if (totle[i]<litle) or (totle[i]>big) then
      begin
        case i of
          1:  chk1.checked:=True;
          2:  chk2.checked:=True;
          3:  chk3.checked:=True;
          4:  chk4.checked:=True;
          5:  chk5.checked:=True;
          6:  chk6.checked:=True;
          7:  chk7.checked:=True;
          8:  chk8.checked:=True;
          9:  chk9.checked:=True;
          10: chk10.Checked:=True;
        end;
      end;
   end;
    if chk1.Checked or chk2.Checked or chk3.Checked or chk4.Checked or
       chk5.Checked or chk6.Checked or chk7.Checked or chk8.Checked or
       chk9.Checked or chk10.Checked then
       begin
         Application.Restore;
         Application.BringToFront;
         if MessageBox(0,'发现不和谐的数据，要把它拿出来晒一晒吗','注意',MB_ICONINFORMATION+MB_YESNO+MB_DEFBUTTON1)=mryes then
           begin
             grp1.Visible:=true;
             btn3.Enabled:=True;
             btn4.Enabled:=False;
             btn3Click(Sender);
           end;
       end else begin
    messagebox(0,'全部数据计算分析导出完毕','完成数据导出',MB_OK);
    end;
    btn1.Enabled:=true;
    
end;


procedure TForm1.Timer1Timer(Sender: TObject);
var
  table:array [0..4] of string;
  i:Integer;
  s1,s2:string;
begin
  Timer1.Enabled:=False;
  s1:='您将要分析的是：  '+#13+#13+'    ';
  table[0]:='二级C语言一级考点表';
  table[1]:='二级VB一级考点表';
  table[2]:='二级VF一级考点表';
  table[3]:='二级Access一级考点表';
  table[4]:='三级网络技术一级考点表';
  for i:=0 to 4 do
   begin
    if TableExist(DBe.DM.con1,table[i]) then Break;
   end;
  case i of
     0:begin
       TableName[0]:='二级C语言一级考点表';
       TableName[1]:='二级C语言二级考点表';
       TableName[2]:='二级C语言三级考点表';
       TableName[3]:='二级C语言四级考点表';s2:='二级C语言';
       is2c:=True;
       end;
     1:begin
       TableName[0]:='二级VB一级考点表';
       TableName[1]:='二级VB二级考点表';
       TableName[2]:='二级VB三级考点表';
       TableName[3]:='二级VB四级考点表';s2:='二级VB';
       end;
     2:begin
       TableName[0]:='二级VF一级考点表';
       TableName[1]:='二级VF二级考点表';
       TableName[2]:='二级VF三级考点表';
       TableName[3]:='二级VF四级考点表';s2:='二级VF';
       end;
     3:begin
       TableName[0]:='二级Access一级考点表';
       TableName[1]:='二级Access二级考点表';
       TableName[2]:='二级Access三级考点表';
       TableName[3]:='二级Access四级考点表';s2:='二级Access';
       end;
     4:begin
       TableName[0]:='三级网络技术一级考点表';
       TableName[1]:='三级网络技术二级考点表';
       TableName[2]:='三级网络技术三级考点表';
       TableName[3]:='三级网络技术四级考点表';s2:='三级网络技术';
       is3ji:=True;
       end;
     else
       MessageBox(0,'数据库错误,没有找到可用的表,本程序将退出.','检查数据库',MB_OK);
       Halt;
   end;

   label1.Caption:=s1+s2;
   end;
//按四级标题分4个表导出
procedure TForm1.Button1Click(Sender: TObject);
var
row1,row2,row3,row4,i,j,k,l,m:Integer;
aq1,aq2,aq3,aq4:TADOQuery;//4个表的查询
cou1,cou2,cou3,cou4:integer;//4个表的记录个数
sumer:array[1..4,1..MAX]of Double;
totle:array[1..MAX] of double;
sht:array[1..4]of variant;
field1,field2,field3,field4:string;
begin
    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('数据表:'''+TableName[1]+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    if rg1.itemindex=0 then table:='题目表'  else table:='模拟题表';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('数据表:'''+Table+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    try
      v:=CreateOleObject('Excel.Application');//如果这个字符串中有空格就会出'无效的类别字符串'的错
      v.Visible:=True;
                v.WorkBooks.Add;   //新建Excel   文件
                v.sheets.add(after:=v.sheets[3]);
                v.workBooks[1].Styles['Normal'].Font.size:=10;
                for i:=1 to 4 do
                begin
                  v.workBooks[1].WorkSheets[i].columns[1].columnwidth:=30;
                  v.workBooks[1].WorkSheets[i].select;
                  sht[i]:=v.WorkBooks[1].WorkSheets[i];
                  sht[i].range['b2'].select;
                  v.ActiveWindow.FreezePanes:= True;
                  v.range['b:j'].columnwidth:=7.5;
                  Sht[i].Cells[1,2]:='1';
                  Sht[i].Cells[1,3]:='2';
                  Sht[i].Cells[1,4]:='3';
                  Sht[i].Cells[1,5]:='4';
                  Sht[i].Cells[1,6]:='5';
                  Sht[i].Cells[1,7]:='6';
                  Sht[i].Cells[1,8]:='7';
                  Sht[i].Cells[1,9]:='8';
                  Sht[i].Cells[1,10]:='9';
                  Sht[i].Cells[1,11]:='10';
                end;
                v.workBooks[1].WorkSheets[1].select;
                sht[1].range['a1'].select;
    except
        showmessage( '初始化Excel失败，可能没装Excel，或者其他错误；请重新再试。 ');
        v.DisplayAlerts:=false;//是否提示存盘
        v.Quit;//退出Excel
        exit;
    end;
    {Application.Restore;
    Application.BringToFront;}
    field1:=' 一级考点名称 ';
    field2:=' 二级考点名称 ';
    field3:=' 三级考点名称 ';
    field4:=' 四级考点名称 ';

    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+'  from '+TableName[0]+'  order by 一级考点ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
 button1.Enabled:=false;
 for i:=1 to cou1 do
   begin
     sht[1].cells[i+1,1]:=aq1.fieldbyname('一级考点名称').asstring;
     aq1.Next;
   end;
  aq1.First;
  v.range['a'+inttostr(cou1+2)].HorizontalAlignment := 3;
  sht[1].cells[cou1+2,1]:='合计';
 with Form1 do
 begin
  row1:=1;row2:=1;row3:=1;row4:=1;
  //crow:=1;  //for i:=1 to MAX do totle[i]:=0;//各章合计
  for i:=1 to cou1 do
  begin
    caption:='第'+inttostr(i)+'章 处理中...';
    for m:=1 to MAX do sum[1,m]:=0;
    row1:=row1+1;//inc(row);

    v.range['a'+inttostr(row1)+':k'+inttostr(row1)].interior.colorindex:=24;
    str:= aq1.fieldbyname('一级考点名称').asstring;
    //sht[1].cells[row1,1]:=aq1.fieldbyname('一级考点名称').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);
    for j:=1 to MAX do    sht[1].cells[row1,j+1]:=sum[1,j]/100;    //本章各套合计
    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+TableName[1]+' where [二级考点所属一级考点ID]='+inttostr(i)+' order by 二级考点ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     row2:=row2+1;    //row2:=crow;
     sht[2].cells[row2,1]:=aq2.fieldbyname('二级考点名称').asstring;
     str:=aq2.fieldbyname('二级考点名称').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);
     for k:=1 to MAX do Sht[2].cells[row2,k+1]:=sum[2,k]/100; //本二级小节各套合计
     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+TableName[2]+' where ([三级考点所属一级考点ID]='+inttostr(i)+
                ') and ([三级考点所属二级考点ID]='+inttostr(j)+') order by 三级考点ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
       row3:=row3+1;    //row3:=crow;
       sht[3].cells[row3,1]:=aq3.fieldbyname('三级考点名称').asstring;
       str:=aq3.fieldbyname('三级考点名称').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+TableName[3]+' where ([四级考点所属一级考点ID]='+inttostr(i)+
                ') and ([四级考点所属二级考点ID]='+inttostr(j)+
                ') and ([四级考点所属三级考点ID]='+inttostr(k)+
                ') order by 四级考点ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
        row4:=row4+1;
        sht[4].cells[row4,1]:=aq4.fieldbyname('四级考点名称').asstring;
        str:=aq4.fieldbyname('四级考点名称').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);
        for m:=1 to MAX do begin
          sht[4].cells[row4,m+1]:=sum[4,m]/100;
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          end;
        aq4.Next;
       end;//第4层结束，第3层的值是第3层的值加所有第4层的和
       for l:=1 to MAX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          sht[3].cells[row3,l+1]:=sum[3,l]/100;
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          end;
       aq3.Next;
     end; //第3层结束
     for l:=1 to MAX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          sht[2].cells[row2,l+1]:=sum[2,l]/100;
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          end;
     aq2.Next;
    end;
    for l:=1 to MAX do begin
          sum[1,l]:=sum[1,l]+sumer[2,l];
          sht[1].cells[row1,l+1]:=sum[1,l]/100;
          totle[l]:=totle[l]+sum[1,l];
          Sht[1].cells[cou1+2,l+1]:=totle[l];
          end;
   aq1.Next;
   end;
 aq4.Free;
 aq3.free;
 aq2.Free;
 aq1.Free;
 end;
 caption:='完成';
 messagebox(0,'全部数据计算分析导出完毕','完成数据导出',MB_OK);
 Button1.Enabled:=true;
end;

function getstr(s:string):string;
begin
  getstr:=copy(s,Pos(' ',s)+1,Length(s)-pos(' ',s));
end;
//找出与考点目录不一致的考点
procedure TForm1.btn3Click(Sender: TObject);
var
crow,i,j,k:Integer;
aq1,aq2,aq3,aq4,ti:TADOQuery;//4个表的查询
cou1,cou2,cou3,cou4,ticou:integer;//4个表的记录个数
sqls:string;s:array[1..MAX]of string;
isbreak:Boolean;
field,field1,field2,field3,field4:string;
begin
    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('数据表:'''+TableName[1]+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    if rg1.itemindex=0 then table:='题目表'  else table:='模拟题表';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('数据表:'''+Table+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    field1:=' 一级考点名称 ';
    field2:=' 二级考点名称 ';
    field3:=' 三级考点名称 ';
    field4:=' 四级考点名称 ';
    field:=' TH,XH,考点1,考点2,考点3,考点4,考点5,考点6 ';
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;    ti:=TADOQuery.Create(nil);ti.Connection:=DBe.DM.con1;
    AQ1.sql.Clear;aq2.SQL.Clear;aq3.SQL.Clear;aq4.SQL.Clear;
    AQ1.SQL.Add('select '+field1+'  from '+TableName[0]+'  order by 一级考点ID ASC');
    AQ2.SQL.Add('select '+field2+' from '+TableName[1]+'  order by 二级考点ID ASC');
    AQ3.SQL.Add('select '+field3+' from '+TableName[2]+'  order by 三级考点ID ASC');
    AQ4.SQL.Add('select '+field4+' from '+TableName[3]+'  order by 四级考点ID ASC');
    AQ1.open;  aq2.Open; aq3.Open;  aq4.Open;
    cou1:=AQ1.recordcount ;  cou2:=aq2.RecordCount;
    cou3:=aq3.RecordCount;   cou4:=aq4.RecordCount;
 //btn2.Enabled:=false;
    if not(chk1.Checked or chk2.Checked or chk3.Checked or
           chk4.Checked or chk5.Checked or chk6.Checked or
           chk7.Checked or chk8.Checked or chk9.Checked or
           chk10.Checked ) then
           begin
             MessageBox(0,'请选择要核对的套号','提示',MB_OK);
             exit;
           end;
    ti.SQL.Clear;
    sqls:='select '+field+' from '+Table+' where ';
    i:=1;//字符串数组开始值
    if chk1.Checked then  begin s[i]:=' TH=1 ';i:=i+1; end;
    if chk2.Checked then  begin s[i]:=' TH=2 ';i:=i+1; end;
    if chk3.Checked then  begin s[i]:=' TH=3 ';i:=i+1; end;
    if chk4.Checked then  begin s[i]:=' TH=4 ';i:=i+1; end;
    if chk5.Checked then  begin s[i]:=' TH=5 ';i:=i+1; end;
    if chk6.Checked then  begin s[i]:=' TH=6 ';i:=i+1; end;
    if chk7.Checked then  begin s[i]:=' TH=7 ';i:=i+1; end;
    if chk8.Checked then  begin s[i]:=' TH=8 ';i:=i+1; end;
    if chk9.Checked then  begin s[i]:=' TH=9 ';i:=i+1; end;
    if chk10.Checked then  begin s[i]:=' TH=10 ';i:=i+1; end;
    sqls:=sqls+s[1];
    if (i-1)>1 then
    for j:=2 to i-1 do
     begin
       sqls:=sqls+'OR'+s[j];
     end;
    ti.SQL.Add(sqls); ti.Open;
    ticou:=ti.RecordCount;
    crow:=1;
    for i:=1 to ticou do
    begin
      for j:=1 to 6 do
      begin
        isbreak:=False;//假设每个考点最开始都没有找到
        aq1.First;aq2.First;aq3.First;aq4.First;//全部指向第一个记录
        str:=ti.fieldByname('考点'+inttostr(j)).AsString;
        if str='' then Break;
        for k:=1 to cou1 do
         begin
           if str=getstr(aq1.fieldbyname('一级考点名称').AsString) then
              begin isbreak:=True;Break;end;   //找到了就不再继续
           aq1.Next;
         end;
         if isbreak then Continue;   //找到了就直接找下一个考点
        for k:=1 to cou2 do
         begin
           if str=getstr(aq2.fieldbyname('二级考点名称').AsString) then
              begin isbreak:=True;Break;end;   //找到了就不再继续
           aq2.Next;
         end;
         if isbreak then Continue;   //找到了就直接找下一个考点
        for k:=1 to cou3 do
         begin
           if str=getstr(aq3.fieldbyname('三级考点名称').AsString) then
              begin isbreak:=True;Break;end;   //找到了就不再继续
           aq3.Next;
         end;
         if isbreak then Continue;   //找到了就直接找下一个考点
        for k:=1 to cou4 do
         begin
           if str=getstr(aq4.fieldbyname('四级考点名称').AsString) then
              begin isbreak:=True;Break;end;   //找到了就不再继续
           aq4.Next;
         end;
         if isbreak then Continue;   //找到了就直接找下一个考点
         if crow=1 then
         begin
           try
              v:=CreateOleObject('Excel.Application');//如果这个字符串中有空格就会出'无效的类别字符串'的错
              v.Visible:=True;
              v.WorkBooks.Add;   //新建Excel   文件
                Sheet:=v.WorkBooks[1].WorkSheets[1];
                v.workBooks[1].Styles['Normal'].Font.size:=10;
                sheet.range['a1'].select;
                Sheet.Cells[1,1]:= '发现不一致的考点如下： ';//单元格内容
                Sheet.Cells[2,1]:= 'TH';
                Sheet.Cells[2,2]:= 'XH';
                Sheet.Cells[2,3]:= '考点名';
                crow:=crow+1;
           except
              showmessage( '初始化Excel失败，可能没装Excel，或者其他错误；请重新再试。 ');
              v.DisplayAlerts:=false;//是否提示存盘
              v.Quit;//退出Excel
              exit;
           end;
         end;
         crow:=crow+1;
         sheet.cells[crow,1]:=ti.FieldByName('TH').AsInteger;
         sheet.cells[crow,2]:=ti.FieldByName('XH').AsInteger;
         sheet.cells[crow,3]:=str;
      end;
      ti.Next;
    end;
    ti.Free;aq1.Free;aq2.Free;aq3.Free;aq4.Free;
    if crow=1 then
    messagebox(0,'没有找到不一致的考点'+#13+#13+
    '请检查比例或分值是否填写正确,'+#13+#13+
    '或者是否有漏填的考点或比例','抱歉',mb_ok);

    {Application.Restore;
    Application.BringToFront;}



end;

procedure TForm1.btn4Click(Sender: TObject);
begin
 btn3.Enabled:=True;
 grp1.Visible:=True;
 btn4.Enabled:=False;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  if FileExists('data\1.mp3') then
  begin
  MCISendString('OPEN data\1.mp3 TYPE mpegvideo alias bgm',nil,0,0);
  MCISendString('PLAY bgm repeat', nil,0,0);
  end else btn6.Visible:=False;
  Application.HintColor := $00C08000;
  screen.HintFont.Color:=$00FFFFFF;
  Screen.HintFont.Name:= '楷体_GB2312';
  Screen.HintFont.Size:=12;
  application.HintHidePause:=10000;
end;
function numtohz(num:Integer):string;
begin
  case num of
    1: numtohz:='一';
    2: numtohz:='二';
    3: numtohz:='三';
    4: numtohz:='四';
    5: numtohz:='五';
    6: numtohz:='六';
    7: numtohz:='七';
    8: numtohz:='八';
    9: numtohz:='九';
    end;
end;
//新组题比例直接写入数据库
procedure TForm1.btn5Click(Sender: TObject);
var
i,j,k,l,m:Integer;
aq1,aq2,aq3,aq4:TADOQuery;//4个表的查询
cou1,cou2,cou3,cou4:integer;//4个表的记录个数
sumer:array[1..4,1..MAX]of Double;
//zt:array[1..9]of string;
field1,field2,field3,field4:string;
begin
    if rg1.ItemIndex=0 then Table:='题目表' else table:='模拟题表';
    if MessageBox(0,PCHar('即将改写数据库中模拟题比例分析结果，'+#13+#13+'数据来源是“'+Table+'”，确定吗？'),'警告',MB_YESNO or MB_DEFBUTTON2	or MB_ICONWARNING)=mrno then Exit;
    if MessageBox(0,'确定一定以及肯定要改写数据库吗？','警告将覆盖原有数据',MB_YESNO or MB_DEFBUTTON2	or MB_ICONWARNING)=mrno then Exit;
    if MessageBox(0,'操作将无法撤销！！！','警告！！',MB_YESNO or MB_DEFBUTTON2	or MB_ICONWARNING)=mrno then Exit;

    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('数据表:'''+TableName[1]+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;

    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('数据表:'''+Table+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;

    {Application.Restore;
    Application.BringToFront;}
//      field1:='*';
//      field2:='*';
//      field3:='*';
//      field4:='*';
{    if is2c then begin
      zt[1]:='2008年4月分值比例';zt[2]:='2008年9月分值比例';
      zt[3]:='2009年3月分值比例';zt[4]:='2009年9月分值比例';
      zt[5]:='2010年3月分值比例';zt[6]:='2010年9月分值比例';
      zt[7]:='2011年3月分值比例';MuX:=7;
      field1:=' 编号,一级考点名称,[2008年4月分值比例],[2008年9月分值比例],[2009年3月分值比例],[2009年9月分值比例],[2010年3月分值比例],[2010年9月分值比例],[2011年3月分值比例] ';
      field2:=' 编号,二级考点名称,[2008年4月分值比例],[2008年9月分值比例],[2009年3月分值比例],[2009年9月分值比例],[2010年3月分值比例],[2010年9月分值比例],[2011年3月分值比例] ';
      field3:=' 编号,三级考点名称,[2008年4月分值比例],[2008年9月分值比例],[2009年3月分值比例],[2009年9月分值比例],[2010年3月分值比例],[2010年9月分值比例],[2011年3月分值比例] ';
      field4:=' 编号,四级考点名称,[2008年4月分值比例],[2008年9月分值比例],[2009年3月分值比例],[2009年9月分值比例],[2010年3月分值比例],[2010年9月分值比例],[2011年3月分值比例] ';
      end
      else begin
      zt[1]:='2007年4月分值比例';zt[2]:='2007年9月分值比例';
      zt[3]:='2008年4月分值比例';zt[4]:='2008年9月分值比例';
      zt[5]:='2009年3月分值比例';zt[6]:='2009年9月分值比例';
      zt[7]:='2010年3月分值比例';zt[8]:='2010年9月分值比例';
      zt[9]:='2011年3月分值比例';//MuX:=9;
}      field1:=' 编号,一级考点名称,第一套模拟题分值比例,第二套模拟题分值比例,第三套模拟题分值比例,第四套模拟题分值比例,第五套模拟题分值比例,第六套模拟题分值比例,第七套模拟题分值比例,第八套模拟题分值比例,第九套模拟题分值比例 ';
      field2:=' 编号,二级考点名称,第一套模拟题分值比例,第二套模拟题分值比例,第三套模拟题分值比例,第四套模拟题分值比例,第五套模拟题分值比例,第六套模拟题分值比例,第七套模拟题分值比例,第八套模拟题分值比例,第九套模拟题分值比例 ';
      field3:=' 编号,三级考点名称,第一套模拟题分值比例,第二套模拟题分值比例,第三套模拟题分值比例,第四套模拟题分值比例,第五套模拟题分值比例,第六套模拟题分值比例,第七套模拟题分值比例,第八套模拟题分值比例,第九套模拟题分值比例 ';
      field4:=' 编号,四级考点名称,第一套模拟题分值比例,第二套模拟题分值比例,第三套模拟题分值比例,第四套模拟题分值比例,第五套模拟题分值比例,第六套模拟题分值比例,第七套模拟题分值比例,第八套模拟题分值比例,第九套模拟题分值比例 ';
      //end;//如果不select‘编号’将会发生键列不足之类的错误
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select *  from '+TableName[0]+'  order by 一级考点ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
    btn5.Enabled:=false;//本按钮暂时失效以防连续点击
 with Form1 do
 begin
  //row1:=1;row2:=1;row3:=1;row4:=1;
  //crow:=1;  //for i:=1 to MAX do totle[i]:=0;//各章合计
  for i:=1 to cou1 do
  begin
    caption:='第'+inttostr(i)+'章 处理中...';
    for m:=1 to MAX do sum[1,m]:=0;
    //row1:=row1+1;//inc(row);

    str:= aq1.fieldbyname('一级考点名称').asstring;
    //sht[1].cells[row1,1]:=aq1.fieldbyname('一级考点名称').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);
{    for j:=1 to MAX do
       aq1['模拟题第'+inttostr(j)+'套分值比例']:=sum[1,j];    //本章各套合计
}    aq2.sql.clear;
    aq2.sql.add('select * from '+TableName[1]+' where [二级考点所属一级考点ID]='+inttostr(i)+' order by 二级考点ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     //row2:=row2+1;    //row2:=crow;
     //sht[2].cells[row2,1]:=aq2.fieldbyname('二级考点名称').asstring;
     str:=aq2.fieldbyname('二级考点名称').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);
{     for k:=1 to MAX do
             aq2['模拟题第'+inttostr(j)+'套分值比例']:=sum[1,j]; //本二级小节各套合计
}     aq3.sql.clear;
     aq3.sql.add('select * from '+TableName[2]+' where ([三级考点所属一级考点ID]='+inttostr(i)+
                ') and ([三级考点所属二级考点ID]='+inttostr(j)+') order by 三级考点ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
     //  row3:=row3+1;    //row3:=crow;
       //sht[3].cells[row3,1]:=aq3.fieldbyname('三级考点名称').asstring;
       str:=aq3.fieldbyname('三级考点名称').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);

       aq4.sql.clear;
       aq4.sql.add('select * from '+TableName[3]+' where ([四级考点所属一级考点ID]='+inttostr(i)+
                ') and ([四级考点所属二级考点ID]='+inttostr(j)+
                ') and ([四级考点所属三级考点ID]='+inttostr(k)+
                ') order by 四级考点ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
       //row4:=row4+1;
        //sht[4].cells[row4,1]:=aq4.fieldbyname('四级考点名称').asstring;
        str:=aq4.fieldbyname('四级考点名称').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);
        aq4.Edit;
        for m:=1 to MAX do begin
          //aq4.FieldByName('模拟题第'+inttostr(m)+'套分值比例').AsFloat:=sum[4,m];//亦可用
          aq4['第'+numtohz(m)+'套模拟题分值比例']:=sum[4,m] /100 ;//可用
          //sht[4].cells[row4,m+1]:=sum[4,m];
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          end;
        aq4.Post;
        aq4.Next;
       end;//第4层结束，第3层的值是第3层的值加所有第4层的和
       aq3.Edit;
       for l:=1 to MAX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          aq3['第'+numtohz(l)+'套模拟题分值比例']:=sum[3,l] / 100;
          //sht[3].cells[row3,l+1]:=sum[3,l];
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          end;
       aq3.Post;
       aq3.Next;
     end; //第3层结束
     aq2.Edit;
     for l:=1 to MAX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          aq2['第'+numtohz(l)+'套模拟题分值比例']:=sum[2,l] / 100;
          //sht[2].cells[row2,l+1]:=sum[2,l];
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          end;
     aq2.Post;
     aq2.Next;
    end;
    aq1.Edit;
    for l:=1 to MAX do begin
          sum[1,l]:=sum[1,l]+sumer[2,l];
          aq1['第'+numtohz(l)+'套模拟题分值比例']:=sum[1,l] / 100;
          //sht[1].cells[row1,l+1]:=sum[1,l];
          //totle[l]:=totle[l]+sum[1,l];
          //Sht[1].cells[cou1+2,l+1]:=totle[l];//总计，在最后一章的下面
          end;
    aq1.Post;
   aq1.Next;
   end;
 aq4.Free;
 aq3.free;
 aq2.Free;
 aq1.Free;
 end;
 caption:='完成';
 messagebox(0,'全部数据已正确地写入数据库中','成功写入全部数据',MB_OK);
 btn5.Enabled:=true;
end;

procedure TForm1.btn6Click(Sender: TObject);
begin
  inplay:=not inplay;
  if inplay then
   begin
      MCISendString('PLAY bgm repeat', nil, 0,0);
      btn6.Caption:='V';
   end
   else begin
      MCISendString('STOP bgm', '', 0, handle);
      MCISendString('CLOSE ANIMATION', '', 0,0);
      btn6.Caption:='X';
   end;
end;
//生成除100的形式
procedure TForm1.btn7Click(Sender: TObject);
var
row1,row2,row3,crow,i,j,k,l,m:Integer;
big,litle:Double;
aq1,aq2,aq3,aq4:TADOQuery;//4个表的查询
cou1,cou2,cou3,cou4:integer;//4个表的记录个数
sumer:array[1..4,1..MAX]of Double;
totle:array[1..MAX] of double;
field,field1,field2,field3,field4:string;
begin
    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('数据表:'''+TableName[1]+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    if rg1.itemindex=0 then table:='题目表'  else table:='模拟题表';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('数据表:'''+Table+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    try
      v:=CreateOleObject('Excel.Application');//如果这个字符串中有空格就会出'无效的类别字符串'的错
      v.Visible:=True;
                v.WorkBooks.Add;   //新建Excel   文件
                //v.WorkBooks[1].WorkSheets[1].name:= '电脑报';//第一页标题
                //v.WorkBooks[1].WorkSheets[2].name:= '编程乐园';
                //v.WorkBooks[1].WorkSheets[3].name:= '都来看呀 ';
                v.activesheet.columns[1].columnwidth:=20;
                //v.Activesheet.usedrange.font.size:=8;
                Sheet:=v.WorkBooks[1].WorkSheets[1];
                v.workBooks[1].Styles['Normal'].Font.size:=10;
{                range:=Sheet.cells.select;
                v.selection.font.size:=10;
                v.Selection.FormatConditions.Delete;
                v.Selection.FormatConditions.Add(1, 3,'0');
                v.Selection.FormatConditions[1].Font.ColorIndex := 15;
}
                v.range['a2'].HorizontalAlignment := 3;
                sheet.cells[2,1]:='合计';
                v.range['b:j'].FormatConditions.Delete;
                v.range['b:j'].FormatConditions.Add(1, 3,'0');
                v.range['b:j'].FormatConditions[1].Font.ColorIndex := 15;
                if is3ji then begin
                  v.range['b2:j2'].FormatConditions.Add(1,1,'0','99.995');
                  v.range['b2:j2'].FormatConditions[2].Font.ColorIndex := 3;
                  v.range['b2:j2'].FormatConditions.Add(1,7,'100.005');end
                else begin
                 v.range['b2:j2'].FormatConditions.Add(1,1,'0','69.995');
                 v.range['b2:j2'].FormatConditions[2].Font.ColorIndex := 3;
                 v.range['b2:j2'].FormatConditions.Add(1,7,'70.005');
                end;
                v.range['b2:j2'].FormatConditions[3].Font.ColorIndex := 3;
                v.range['b:j'].columnwidth:=7.5;
                sheet.range['b3'].select;
                v.ActiveWindow.FreezePanes:= True;
                sheet.range['a1'].select;
                //Sheet.Cells[1,1]:= '好看 ';//单元格内容
                //Sheet.Cells[1,2]:= '确实 ';
                //Sheet.Cells[2,1]:= '我喜欢 ';
{                v.activesheet.rows[2].select;
                v.Selection.Interior.ColorIndex:= 32;      
                v.range['b2:j2'].interior.colorindex:=39;//粉色
 }
    except
        showmessage( '初始化Excel失败，可能没装Excel，或者其他错误；请重新再试。 ');
        v.DisplayAlerts:=false;//是否提示存盘
        v.Quit;//退出Excel
        exit;
    end;
    {Application.Restore;
    Application.BringToFront;}
    field1:=' 一级考点名称 ';
    field2:=' 二级考点名称 ';
    field3:=' 三级考点名称 ';
    field4:=' 四级考点名称 ';
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+' from '+TableName[0]+'  order by 一级考点ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
 btn7.Enabled:=false;
 with Form1 do
 begin
   Sheet.Cells[1,2]:='1';
   Sheet.Cells[1,3]:='2';
   Sheet.Cells[1,4]:='3';
   Sheet.Cells[1,5]:='4';
   Sheet.Cells[1,6]:='5';
   Sheet.Cells[1,7]:='6';
   Sheet.Cells[1,8]:='7';
   Sheet.Cells[1,9]:='8';   Sheet.Cells[1,10]:='9';
  crow:=2;for i:=1 to MAX do totle[i]:=0;
  for i:=1 to cou1 do
  begin
    caption:='第'+inttostr(i)+'章 处理中...';
    for m:=1 to MAX do sum[1,m]:=0;
    crow:=crow+1;//inc(row);
    row1:=crow;
    v.range['a'+inttostr(crow)+':j'+inttostr(crow)].interior.colorindex:=24;
    str:= aq1.fieldbyname('一级考点名称').asstring;
    Sheet.cells[crow,1]:=' '+aq1.fieldbyname('一级考点名称').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);
    for j:=1 to MAX do
      Sheet.cells[crow,j+1]:=sum[1,j]/100;
    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+TableName[1]+' where [二级考点所属一级考点ID]='+inttostr(i)+' order by 二级考点ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     crow:=crow+1;    row2:=crow;
     Sheet.cells[crow,1]:='     '+aq2.fieldbyname('二级考点名称').asstring;
     str:=aq2.fieldbyname('二级考点名称').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);
     for k:=1 to MAX do Sheet.cells[crow,k+1]:=sum[2,k]/100;
     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+TableName[2]+' where ([三级考点所属一级考点ID]='+inttostr(i)+
                ') and ([三级考点所属二级考点ID]='+inttostr(j)+') order by 三级考点ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
       crow:=crow+1;    row3:=crow;
       Sheet.cells[crow,1]:='         '+aq3.fieldbyname('三级考点名称').asstring;
       str:=aq3.fieldbyname('三级考点名称').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+TableName[3]+' where ([四级考点所属一级考点ID]='+inttostr(i)+
                ') and ([四级考点所属二级考点ID]='+inttostr(j)+
                ') and ([四级考点所属三级考点ID]='+inttostr(k)+
                ') order by 四级考点ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
        crow:=crow+1;
        Sheet.cells[crow,1]:='             '+aq4.fieldbyname('四级考点名称').asstring;
        str:=aq4.fieldbyname('四级考点名称').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);
        for m:=1 to MAX do begin
          Sheet.cells[crow,m+1]:=sum[4,m]/100;
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          end;
        aq4.Next;
       end;//第4层结束，第3层的值是第3层的值加所有第4层的和
       for l:=1 to MAX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          Sheet.cells[row3,l+1]:=sum[3,l]/100;
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          end;
       aq3.Next;
     end; //第3层结束
     for l:=1 to MAX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          Sheet.cells[row2,l+1]:=sum[2,l]/100;
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          end;
     aq2.Next;
    end;
    for l:=1 to MAX do begin
          sum[1,l]:=sum[1,l]+sumer[2,l];
          Sheet.cells[row1,l+1]:=sum[1,l]/100;
          totle[l]:=totle[l]+sum[1,l];
          Sheet.cells[2,l+1]:=totle[l];
          end;
   aq1.Next;
   end;
    aq4.Free;
     aq3.free;
     aq2.Free;
     aq1.Free;
 end;
   caption:='完成';

   if is3ji then begin big:=100.005;litle:=99.995;end
     else begin big:=70.005;litle:=69.995;end;
   for i:=1 to MAX do
   begin
     if (totle[i]<litle) or (totle[i]>big) then
      begin
        case i of
          1:  chk1.checked:=True;
          2:  chk2.checked:=True;
          3:  chk3.checked:=True;
          4:  chk4.checked:=True;
          5:  chk5.checked:=True;
          6:  chk6.checked:=True;
          7:  chk7.checked:=True;
          8:  chk8.checked:=True;
          9:  chk9.checked:=True;
          10: chk10.Checked:=True;
        end;
      end;
   end;
    if chk1.Checked or chk2.Checked or chk3.Checked or chk4.Checked or
       chk5.Checked or chk6.Checked or chk7.Checked or chk8.Checked or
       chk9.Checked or chk10.Checked  then
       begin
         Application.Restore;
         Application.BringToFront;
         if MessageBox(0,'发现不和谐的数据，要把它拿出来晒一晒吗','注意',MB_ICONINFORMATION+MB_YESNO+MB_DEFBUTTON1)=mryes then
           begin
             grp1.Visible:=true;
             btn3.Enabled:=True;
             btn4.Enabled:=False;
             btn3Click(Sender);
           end;
       end else begin
    messagebox(0,'全部数据计算分析导出完毕','完成数据导出',MB_OK);
    end;
    btn7.Enabled:=true;


end;
//写入真题比例
procedure TForm1.btn8Click(Sender: TObject);
var
i,j,k,l,m,MuX:Integer;
aq1,aq2,aq3,aq4:TADOQuery;//4个表的查询
cou1,cou2,cou3,cou4:integer;//4个表的记录个数
sumer:array[1..4,1..MAX]of Double;
zt:array[1..10]of string;
field1,field2,field3,field4:string;
begin
    if MessageBox(0,'即将改写数据库中真题比例分析结果，确定吗？','警告',MB_YESNO or MB_DEFBUTTON2	or MB_ICONWARNING)=mrno then Exit;
    if MessageBox(0,'确定一定以及肯定要改写数据库吗？','警告将覆盖原有数据',MB_YESNO or MB_DEFBUTTON2	or MB_ICONWARNING)=mrno then Exit;
    if MessageBox(0,'操作将无法撤销！！！','警告！！',MB_YESNO or MB_DEFBUTTON2	or MB_ICONWARNING)=mrno then Exit;

    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('数据表:'''+TableName[1]+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    table:='题目表';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('数据表:'''+Table+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;

    {Application.Restore;
    Application.BringToFront;}
//      field1:='*';
//      field2:='*';
//      field3:='*';
//      field4:='*';
    if is2c then begin
      zt[1]:='2008年4月分值比例';zt[2]:='2008年9月分值比例';
      zt[3]:='2009年3月分值比例';zt[4]:='2009年9月分值比例';
      zt[5]:='2010年3月分值比例';zt[6]:='2010年9月分值比例';
      zt[7]:='2011年3月分值比例';zt[8]:='2011年9月分值比例';MuX:=8;
      field1:=' 编号,一级考点名称,[2008年4月分值比例],[2008年9月分值比例],[2009年3月分值比例],[2009年9月分值比例],[2010年3月分值比例],[2010年9月分值比例],[2011年3月分值比例],[2011年9月分值比例] ';
      field2:=' 编号,二级考点名称,[2008年4月分值比例],[2008年9月分值比例],[2009年3月分值比例],[2009年9月分值比例],[2010年3月分值比例],[2010年9月分值比例],[2011年3月分值比例],[2011年9月分值比例] ';
      field3:=' 编号,三级考点名称,[2008年4月分值比例],[2008年9月分值比例],[2009年3月分值比例],[2009年9月分值比例],[2010年3月分值比例],[2010年9月分值比例],[2011年3月分值比例],[2011年9月分值比例] ';
      field4:=' 编号,四级考点名称,[2008年4月分值比例],[2008年9月分值比例],[2009年3月分值比例],[2009年9月分值比例],[2010年3月分值比例],[2010年9月分值比例],[2011年3月分值比例],[2011年9月分值比例] ';
      end
      else begin
      zt[1]:='2007年4月分值比例';zt[2]:='2007年9月分值比例';
      zt[3]:='2008年4月分值比例';zt[4]:='2008年9月分值比例';
      zt[5]:='2009年3月分值比例';zt[6]:='2009年9月分值比例';
      zt[7]:='2010年3月分值比例';zt[8]:='2010年9月分值比例';
      zt[9]:='2011年3月分值比例';zt[10]:='2011年9月分值比例';MuX:=10;
      field1:=' 编号,一级考点名称,[2007年4月分值比例],[2007年9月分值比例],[2008年4月分值比例],[2008年9月分值比例],[2009年3月分值比例],[2009年9月分值比例],[2010年3月分值比例],[2010年9月分值比例],[2011年3月分值比例],[2011年9月分值比例] ';
      field2:=' 编号,二级考点名称,[2007年4月分值比例],[2007年9月分值比例],[2008年4月分值比例],[2008年9月分值比例],[2009年3月分值比例],[2009年9月分值比例],[2010年3月分值比例],[2010年9月分值比例],[2011年3月分值比例],[2011年9月分值比例] ';
      field3:=' 编号,三级考点名称,[2007年4月分值比例],[2007年9月分值比例],[2008年4月分值比例],[2008年9月分值比例],[2009年3月分值比例],[2009年9月分值比例],[2010年3月分值比例],[2010年9月分值比例],[2011年3月分值比例],[2011年9月分值比例] ';
      field4:=' 编号,四级考点名称,[2007年4月分值比例],[2007年9月分值比例],[2008年4月分值比例],[2008年9月分值比例],[2009年3月分值比例],[2009年9月分值比例],[2010年3月分值比例],[2010年9月分值比例],[2011年3月分值比例],[2011年9月分值比例] ';
      end;
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+'  from '+TableName[0]+'  order by 一级考点ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
    btn8.Enabled:=false;//本按钮暂时失效以防连续点击
 with Form1 do
 begin
  //row1:=1;row2:=1;row3:=1;row4:=1;
  //crow:=1;  //for i:=1 to MAX do totle[i]:=0;//各章合计
  for i:=1 to cou1 do
  begin
    caption:='第'+inttostr(i)+'章 处理中...';
    for m:=1 to MAX do sum[1,m]:=0;
    //row1:=row1+1;//inc(row);

    str:= aq1.fieldbyname('一级考点名称').asstring;
    //sht[1].cells[row1,1]:=aq1.fieldbyname('一级考点名称').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);
{    for j:=1 to MAX do
       aq1['模拟题第'+inttostr(j)+'套分值比例']:=sum[1,j];    //本章各套合计
}    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+TableName[1]+' where [二级考点所属一级考点ID]='+inttostr(i)+' order by 二级考点ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     //row2:=row2+1;    //row2:=crow;
     //sht[2].cells[row2,1]:=aq2.fieldbyname('二级考点名称').asstring;
     str:=aq2.fieldbyname('二级考点名称').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);
{     for k:=1 to MAX do
             aq2['模拟题第'+inttostr(j)+'套分值比例']:=sum[1,j]; //本二级小节各套合计
}     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+TableName[2]+' where ([三级考点所属一级考点ID]='+inttostr(i)+
                ') and ([三级考点所属二级考点ID]='+inttostr(j)+') order by 三级考点ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
     //  row3:=row3+1;    //row3:=crow;
       //sht[3].cells[row3,1]:=aq3.fieldbyname('三级考点名称').asstring;
       str:=aq3.fieldbyname('三级考点名称').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+TableName[3]+' where ([四级考点所属一级考点ID]='+inttostr(i)+
                ') and ([四级考点所属二级考点ID]='+inttostr(j)+
                ') and ([四级考点所属三级考点ID]='+inttostr(k)+
                ') order by 四级考点ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
       //row4:=row4+1;
        //sht[4].cells[row4,1]:=aq4.fieldbyname('四级考点名称').asstring;
        str:=aq4.fieldbyname('四级考点名称').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);
        aq4.Edit;
        for m:=1 to MuX do begin
          //aq4.FieldByName('模拟题第'+inttostr(m)+'套分值比例').AsFloat:=sum[4,m];//亦可用
          aq4[zt[m]]:=sum[4,m] /100 ;//可用
          //sht[4].cells[row4,m+1]:=sum[4,m];
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          end;
        aq4.Post;
        aq4.Next;
       end;//第4层结束，第3层的值是第3层的值加所有第4层的和
       aq3.Edit;
       for l:=1 to MuX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          aq3[zt[l]]:=sum[3,l] / 100;
          //sht[3].cells[row3,l+1]:=sum[3,l];
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          end;
       aq3.Post;
       aq3.Next;
     end; //第3层结束
     aq2.Edit;
     for l:=1 to MuX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          aq2[zt[l]]:=sum[2,l] / 100;
          //sht[2].cells[row2,l+1]:=sum[2,l];
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          end;
     aq2.Post;
     aq2.Next;
    end;
    aq1.Edit;
    for l:=1 to MuX do begin
          sum[1,l]:=sum[1,l]+sumer[2,l];
          aq1[zt[l]]:=sum[1,l] / 100;
          //sht[1].cells[row1,l+1]:=sum[1,l];
          //totle[l]:=totle[l]+sum[1,l];
          //Sht[1].cells[cou1+2,l+1]:=totle[l];//总计，在最后一章的下面
          end;
    aq1.Post;
   aq1.Next;
   end;
 aq4.Free;
 aq3.free;
 aq2.Free;
 aq1.Free;
 end;
 caption:='完成';
 messagebox(0,'全部数据已正确地写入数据库中','成功写入全部数据',MB_OK);
 btn8.Enabled:=true;

end;

procedure TForm1.btn2Click(Sender: TObject);
var
row1,row2,row3,crow,i,j,k,l,m:Integer;
big,litle:Double;
aq1,aq2,aq3,aq4:TADOQuery;//4个表的查询
cou1,cou2,cou3,cou4:integer;//4个表的记录个数
PubTable:array[0..3]of string;
sumer:array[1..4,1..MAX]of Double;
totle:array[1..MAX] of double;
field,field1,field2,field3,field4:string;
begin
   PubTable[0]:='公共基础知识一级考点表';
   PubTable[1]:='公共基础知识二级考点表';
   PubTable[2]:='公共基础知识三级考点表';
   PubTable[3]:='公共基础知识四级考点表';
    if not(TableExist(DBe.DM.con1,PubTable[1])) then
    begin
       MessageBox(0,PCHar('数据表:'''+PubTable[1]+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    if rg1.itemindex=0 then table:='题目表'  else table:='模拟题表';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('数据表:'''+Table+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    try
      v:=CreateOleObject('Excel.Application');//如果这个字符串中有空格就会出'无效的类别字符串'的错
      v.Visible:=True;
                v.WorkBooks.Add;   //新建Excel   文件
                v.WorkBooks[1].WorkSheets[1].name:= Table+'－公共';//第一页标题
                //v.WorkBooks[1].WorkSheets[2].name:= '编程乐园';
                //v.WorkBooks[1].WorkSheets[3].name:= '都来看呀 ';
                v.activesheet.columns[1].columnwidth:=20;
                //v.Activesheet.usedrange.font.size:=8;
                Sheet:=v.WorkBooks[1].WorkSheets[1];
                v.workBooks[1].Styles['Normal'].Font.size:=10;
{                range:=Sheet.cells.select;
                v.selection.font.size:=10;
                v.Selection.FormatConditions.Delete;
                v.Selection.FormatConditions.Add(1, 3,'0');
                v.Selection.FormatConditions[1].Font.ColorIndex := 15;
}
                v.range['a2'].HorizontalAlignment := 3;
                sheet.cells[2,1]:='合计';
                v.range['b:k'].FormatConditions.Delete;
                v.range['b:k'].FormatConditions.Add(1, 3,'0');
                v.range['b:k'].FormatConditions[1].Font.ColorIndex := 15;
                if is3ji then begin
                  v.range['b2:k2'].FormatConditions.Add(1,1,'0','99.995');
                  v.range['b2:k2'].FormatConditions[2].Font.ColorIndex := 3;
                  v.range['b2:k2'].FormatConditions.Add(1,7,'100.005');end
                else begin
                 v.range['b2:k2'].FormatConditions.Add(1,1,'0','69.995');
                 v.range['b2:k2'].FormatConditions[2].Font.ColorIndex := 3;
                 v.range['b2:k2'].FormatConditions.Add(1,7,'70.005');
                end;
                v.range['b2:k2'].FormatConditions[3].Font.ColorIndex := 3;
                v.range['b:k'].columnwidth:=7.5;
                sheet.range['b3'].select;
                v.ActiveWindow.FreezePanes:= True;
                sheet.range['a1'].select;
                //Sheet.Cells[1,1]:= '好看 ';//单元格内容
                //Sheet.Cells[1,2]:= '确实 ';
                //Sheet.Cells[2,1]:= '我喜欢 ';
{                v.activesheet.rows[2].select;
                v.Selection.Interior.ColorIndex:= 32;      
                v.range['b2:j2'].interior.colorindex:=39;//粉色
 }
    except
        showmessage( '初始化Excel失败，可能没装Excel，或者其他错误；请重新再试。 ');
        v.DisplayAlerts:=false;//是否提示存盘
        v.Quit;//退出Excel
        exit;
    end;
    {Application.Restore;
    Application.BringToFront;}
    field1:=' 一级考点名称 ';
    field2:=' 二级考点名称 ';
    field3:=' 三级考点名称 ';
    field4:=' 四级考点名称 ';
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+' from '+PubTable[0]+'  order by 一级考点ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
 btn2.Enabled:=false;
 with Form1 do
 begin
   Sheet.Cells[1,2]:='1';
   Sheet.Cells[1,3]:='2';
   Sheet.Cells[1,4]:='3';
   Sheet.Cells[1,5]:='4';
   Sheet.Cells[1,6]:='5';
   Sheet.Cells[1,7]:='6';
   Sheet.Cells[1,8]:='7';
   Sheet.Cells[1,9]:='8';   Sheet.Cells[1,10]:='9';
   Sheet.Cells[1,11]:='10';
  crow:=2;for i:=1 to MAX do totle[i]:=0;
  for i:=1 to cou1 do
  begin
    caption:='第'+inttostr(i)+'章 处理中...';
    for m:=1 to MAX do sum[1,m]:=0;
    crow:=crow+1;//inc(row);
    row1:=crow;
    v.range['a'+inttostr(crow)+':k'+inttostr(crow)].interior.colorindex:=24;//颜色范围到K列
    str:= aq1.fieldbyname('一级考点名称').asstring;
    Sheet.cells[crow,1]:=' '+aq1.fieldbyname('一级考点名称').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);
    for j:=1 to MAX do
      Sheet.cells[crow,j+1]:=sum[1,j]/100;
    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+PubTable[1]+' where [二级考点所属一级考点ID]='+inttostr(i)+' order by 二级考点ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     crow:=crow+1;    row2:=crow;
     Sheet.cells[crow,1]:='     '+aq2.fieldbyname('二级考点名称').asstring;
     str:=aq2.fieldbyname('二级考点名称').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);
     for k:=1 to MAX do Sheet.cells[crow,k+1]:=sum[2,k]/100;
     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+PubTable[2]+' where ([三级考点所属一级考点ID]='+inttostr(i)+
                ') and ([三级考点所属二级考点ID]='+inttostr(j)+') order by 三级考点ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
       crow:=crow+1;    row3:=crow;
       Sheet.cells[crow,1]:='         '+aq3.fieldbyname('三级考点名称').asstring;
       str:=aq3.fieldbyname('三级考点名称').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+PubTable[3]+' where ([四级考点所属一级考点ID]='+inttostr(i)+
                ') and ([四级考点所属二级考点ID]='+inttostr(j)+
                ') and ([四级考点所属三级考点ID]='+inttostr(k)+
                ') order by 四级考点ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
        crow:=crow+1;
        Sheet.cells[crow,1]:='             '+aq4.fieldbyname('四级考点名称').asstring;
        str:=aq4.fieldbyname('四级考点名称').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);
        for m:=1 to MAX do begin
          Sheet.cells[crow,m+1]:=sum[4,m]/100;
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          end;
        aq4.Next;
       end;//第4层结束，第3层的值是第3层的值加所有第4层的和
       for l:=1 to MAX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          Sheet.cells[row3,l+1]:=sum[3,l]/100;
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          end;
       aq3.Next;
     end; //第3层结束
     for l:=1 to MAX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          Sheet.cells[row2,l+1]:=sum[2,l]/100;
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          end;
     aq2.Next;
    end;
    for l:=1 to MAX do begin
          sum[1,l]:=sum[1,l]+sumer[2,l];
          Sheet.cells[row1,l+1]:=sum[1,l]/100;
          totle[l]:=totle[l]+sum[1,l];
          Sheet.cells[2,l+1]:=totle[l];
          end;
   aq1.Next;
   end;
    aq4.Free;
     aq3.free;
     aq2.Free;
     aq1.Free;
 end;
   caption:='完成';
  {
   if is3ji then begin big:=100.005;litle:=99.995;end
     else begin big:=70.005;litle:=69.995;end;
   for i:=1 to MAX do
   begin
     if (totle[i]<litle) or (totle[i]>big) then
      begin
        case i of
          1:  chk1.checked:=True;
          2:  chk2.checked:=True;
          3:  chk3.checked:=True;
          4:  chk4.checked:=True;
          5:  chk5.checked:=True;
          6:  chk6.checked:=True;
          7:  chk7.checked:=True;
          8:  chk8.checked:=True;
          9:  chk9.checked:=True;
        end;
      end;
   end;}
 {   if chk1.Checked or chk2.Checked or chk3.Checked or chk4.Checked or
       chk5.Checked or chk6.Checked or chk7.Checked or chk8.Checked or
       chk9.Checked then
       begin
         Application.Restore;
         Application.BringToFront;
         if MessageBox(0,'发现不和谐的数据，要把它拿出来晒一晒吗','注意',MB_ICONINFORMATION+MB_YESNO+MB_DEFBUTTON1)=mryes then
           begin
             grp1.Visible:=true;
             btn3.Enabled:=True;
             btn4.Enabled:=False;
             btn3Click(Sender);
           end;
       end else begin
 }   messagebox(0,'全部数据计算分析导出完毕','完成数据导出',MB_OK);
 //   end;
    btn2.Enabled:=true;
end;
//公共基础输出四个表
procedure TForm1.btn9Click(Sender: TObject);
var
row1,row2,row3,row4,i,j,k,l,m:Integer;
aq1,aq2,aq3,aq4:TADOQuery;//4个表的查询
cou1,cou2,cou3,cou4:integer;//4个表的记录个数
sumer:array[1..4,1..MAX]of Double;
totle:array[1..MAX] of double;
sht:array[1..4]of variant;
field1,field2,field3,field4:string;
PubTable:array[0..3]of string;

begin
   PubTable[0]:='公共基础知识一级考点表';
   PubTable[1]:='公共基础知识二级考点表';
   PubTable[2]:='公共基础知识三级考点表';
   PubTable[3]:='公共基础知识四级考点表';

    if not(TableExist(DBe.DM.con1,PubTable[1])) then
    begin
       MessageBox(0,PCHar('数据表:'''+PubTable[1]+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    if rg1.itemindex=0 then table:='题目表'  else table:='模拟题表';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('数据表:'''+Table+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '出错了',MB_OK or MB_ICONWARNING);
       halt;
    end;
    try
      v:=CreateOleObject('Excel.Application');//如果这个字符串中有空格就会出'无效的类别字符串'的错
      v.Visible:=True;
                v.WorkBooks.Add;   //新建Excel   文件
                v.WorkBooks[1].WorkSheets[1].name:= Table+'－公共';
                v.sheets.add(after:=v.sheets[3]);
                v.workBooks[1].Styles['Normal'].Font.size:=10;
                for i:=1 to 4 do
                begin
                  v.workBooks[1].WorkSheets[i].columns[1].columnwidth:=30;
                  v.workBooks[1].WorkSheets[i].select;
                  sht[i]:=v.WorkBooks[1].WorkSheets[i];
                  sht[i].range['b2'].select;
                  v.ActiveWindow.FreezePanes:= True;
                  v.range['b:j'].columnwidth:=7.5;
                  Sht[i].Cells[1,2]:='1';
                  Sht[i].Cells[1,3]:='2';
                  Sht[i].Cells[1,4]:='3';
                  Sht[i].Cells[1,5]:='4';
                  Sht[i].Cells[1,6]:='5';
                  Sht[i].Cells[1,7]:='6';
                  Sht[i].Cells[1,8]:='7';
                  Sht[i].Cells[1,9]:='8';
                  Sht[i].Cells[1,10]:='9';
                  Sht[i].Cells[1,11]:='10';//第10套表头
                end;
                v.workBooks[1].WorkSheets[1].select;
                sht[1].range['a1'].select;
    except
        showmessage( '初始化Excel失败，可能没装Excel，或者其他错误；请重新再试。 ');
        v.DisplayAlerts:=false;//是否提示存盘
        v.Quit;//退出Excel
        exit;
    end;
    {Application.Restore;
    Application.BringToFront;}
    field1:=' 一级考点名称 ';
    field2:=' 二级考点名称 ';
    field3:=' 三级考点名称 ';
    field4:=' 四级考点名称 ';

    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+'  from '+PubTable[0]+'  order by 一级考点ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
 btn9.Enabled:=false;
 for i:=1 to cou1 do
   begin
     sht[1].cells[i+1,1]:=aq1.fieldbyname('一级考点名称').asstring;
     aq1.Next;
   end;
  aq1.First;
  v.range['a'+inttostr(cou1+2)].HorizontalAlignment := 3;
  sht[1].cells[cou1+2,1]:='合计';
 with Form1 do
 begin
  row1:=1;row2:=1;row3:=1;row4:=1;
  //crow:=1;  //for i:=1 to MAX do totle[i]:=0;//各章合计
  for i:=1 to cou1 do
  begin
    caption:='第'+inttostr(i)+'章 处理中...';
    for m:=1 to MAX do sum[1,m]:=0;
    row1:=row1+1;//inc(row);
//下句中':k'表示k列（第11列）
    v.range['a'+inttostr(row1)+':k'+inttostr(row1)].interior.colorindex:=24;
    str:= aq1.fieldbyname('一级考点名称').asstring;
    //sht[1].cells[row1,1]:=aq1.fieldbyname('一级考点名称').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);
    for j:=1 to MAX do    sht[1].cells[row1,j+1]:=sum[1,j]/100;    //本章各套合计
    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+PubTable[1]+' where [二级考点所属一级考点ID]='+inttostr(i)+' order by 二级考点ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     row2:=row2+1;    //row2:=crow;
     sht[2].cells[row2,1]:=aq2.fieldbyname('二级考点名称').asstring;
     str:=aq2.fieldbyname('二级考点名称').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);
     for k:=1 to MAX do Sht[2].cells[row2,k+1]:=sum[2,k]/100; //本二级小节各套合计
     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+PubTable[2]+' where ([三级考点所属一级考点ID]='+inttostr(i)+
                ') and ([三级考点所属二级考点ID]='+inttostr(j)+') order by 三级考点ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
       row3:=row3+1;    //row3:=crow;
       sht[3].cells[row3,1]:=aq3.fieldbyname('三级考点名称').asstring;
       str:=aq3.fieldbyname('三级考点名称').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+PubTable[3]+' where ([四级考点所属一级考点ID]='+inttostr(i)+
                ') and ([四级考点所属二级考点ID]='+inttostr(j)+
                ') and ([四级考点所属三级考点ID]='+inttostr(k)+
                ') order by 四级考点ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
        row4:=row4+1;
        sht[4].cells[row4,1]:=aq4.fieldbyname('四级考点名称').asstring;
        str:=aq4.fieldbyname('四级考点名称').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);
        for m:=1 to MAX do begin
          sht[4].cells[row4,m+1]:=sum[4,m]/100;
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          end;
        aq4.Next;
       end;//第4层结束，第3层的值是第3层的值加所有第4层的和
       for l:=1 to MAX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          sht[3].cells[row3,l+1]:=sum[3,l]/100;
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          end;
       aq3.Next;
     end; //第3层结束
     for l:=1 to MAX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          sht[2].cells[row2,l+1]:=sum[2,l]/100;
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          end;
     aq2.Next;
    end;
    for l:=1 to MAX do begin
          sum[1,l]:=sum[1,l]+sumer[2,l];
          sht[1].cells[row1,l+1]:=sum[1,l]/100;
          totle[l]:=totle[l]+sum[1,l];
          Sht[1].cells[cou1+2,l+1]:=totle[l];
          end;
   aq1.Next;
   end;
 aq4.Free;
 aq3.free;
 aq2.Free;
 aq1.Free;
 end;
 caption:='完成';
 messagebox(0,'全部数据计算分析导出完毕','完成数据导出',MB_OK);
 Btn9.Enabled:=true;
end;

end.
