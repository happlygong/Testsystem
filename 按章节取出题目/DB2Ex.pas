unit DB2Ex;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,ComObj,ADODB,StrUtils, ExtCtrls,
  MMSystem, RxRichEd;

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
    btn10: TButton;
    edt1: TRxRichEdit;
    btn11: TButton;
    btn12: TButton;
    chk11: TCheckBox;
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
    procedure btn10Click(Sender: TObject);
    procedure btn11Click(Sender: TObject);
    procedure btn12Click(Sender: TObject);
    procedure rg1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;
//const
  //MAX=10;
var
  Form1 :   TForm1;
  v,range:   Variant;
  Sheet :   Variant; MAX:integer=10;
  sum   :   array of array of Double;
  str:string;
  inplay:Boolean=true;
  is3ji:Boolean=False;
  is2c:boolean=False;
  TableName :array[0..3] of string=('����C����һ�������','����C���Զ��������','����C�������������','����C�����ļ������');
//  TableName :array[0..3] of string=('����Accessһ�������','����Access���������','����Access���������','����Access�ļ������');
//  TableName :array[0..3] of string=('����VBһ�������','����VB���������','����VB���������','����VB�ļ������');
//  TableName :array[0..3] of string=('����VFһ�������','����VF���������','����VF���������','����VF�ļ������');
//  TableName :array[0..3] of string=('�������缼��һ�������','�������缼�����������','�������缼�����������','�������缼���ļ������');
  table:string;
implementation
uses DBe;
{$R *.dfm}
//���ݱ��Ƿ����
Function TableExist(pConn:TADOConnection; pcTable:string ):boolean;overload;
var
  tmpFldList : TStrings ;
  nLoop : integer ;
begin
    Result := False ;
    tmpFldList := TStringList.Create ;
    pConn.GetTableNames( tmpFldList, True ); // ����ϵͳ��
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

////�ȽϿ�������ֵ////
procedure getsum(level:Integer;str:string;table:string);
var
  i,th,fz,cou:Integer;   bili:Double;
  ti:tadoquery;
  kaodian:string;
  begin
    ti:=Tadoquery.create(nil);
    TI.connection:=DBe.dm.con1;
    ti.SQL.Clear;
    ti.SQL.Add('select TH,FZ,����1,����1����,����2,����2����,����3,����3����,����4,����4����,����5,����5����,����6,����6���� from '+table);ti.open;
    cou:=ti.RecordCount;
    for i:=1 to cou do
    begin
      kaodian:=ti.fieldbyname('����1').AsString;
      if kaodian=str then
        begin
         th:=ti.fieldbyname('TH').AsInteger;
         fz:=ti.fieldbyname('FZ').AsInteger;
         bili:=ti.fieldbyname('����1����').AsFloat;
         sum[level,th]:=sum[level,th]+fz*bili;
        end;
      kaodian:=ti.fieldbyname('����2').AsString;
      if kaodian='' then begin ti.Next;Continue;end;
      if kaodian=str then
        begin
         th:=ti.fieldbyname('TH').AsInteger;
         fz:=ti.fieldbyname('FZ').AsInteger;
         bili:=ti.fieldbyname('����2����').AsFloat;
         sum[level,th]:=sum[level,th]+fz*bili;
        end;
      kaodian:=ti.fieldbyname('����3').AsString;
      if kaodian='' then begin ti.Next;Continue;end;
      if kaodian=str then
        begin
         th:=ti.fieldbyname('TH').AsInteger;
         fz:=ti.fieldbyname('FZ').AsInteger;
         bili:=ti.fieldbyname('����3����').AsFloat;
         sum[level,th]:=sum[level,th]+fz*bili;
        end;
      kaodian:=ti.fieldbyname('����4').AsString;
      if kaodian='' then begin ti.Next;Continue;end;
      if kaodian=str then
        begin
         th:=ti.fieldbyname('TH').AsInteger;
         fz:=ti.fieldbyname('FZ').AsInteger;
         bili:=ti.fieldbyname('����4����').AsFloat;
         sum[level,th]:=sum[level,th]+fz*bili;
        end;
      kaodian:=ti.fieldbyname('����5').AsString;
      if kaodian='' then begin ti.Next;Continue;end;
      if kaodian=str then
        begin
         th:=ti.fieldbyname('TH').AsInteger;
         fz:=ti.fieldbyname('FZ').AsInteger;
         bili:=ti.fieldbyname('����5����').AsFloat;
         sum[level,th]:=sum[level,th]+fz*bili;
        end;
      kaodian:=ti.fieldbyname('����6').AsString;
      if kaodian='' then begin ti.Next;Continue;end;
      if kaodian=str then
        begin
         th:=ti.fieldbyname('TH').AsInteger;
         fz:=ti.fieldbyname('FZ').AsInteger;
         bili:=ti.fieldbyname('����6����').AsFloat;
         sum[level,th]:=sum[level,th]+fz*bili;
        end;
      ti.Next;
    end;
    ti.Free;
  end;
function getstr(s:string):string;
begin
  getstr:=copy(s,Pos(' ',s)+1,Length(s)-pos(' ',s));
end;
//�ȽϿ��㣬���¼���Ƿ��г���
procedure CheckTestPoint(THstr:string);
var
crow,i,j,k:Integer;
aq1,aq2,aq3,aq4,ti:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4,ticou:integer;//4����ļ�¼����
sqls:string;s:array of string;
isbreak,biaotoufirst:Boolean;
field,field1,field2,field3,field4:string;

begin
    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+TableName[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       exit;
    end;
    Setlength(s,MAX+1);
    if Form1.rg1.itemindex=0 then table:='��Ŀ��'  else table:='ģ�����';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+Table+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       exit;
    end;
    field1:=' һ���������� ';
    field2:=' ������������ ';
    field3:=' ������������ ';
    field4:=' �ļ��������� ';
    field:=' TH,XH,����1,����2,����3,����4,����5,����6,����1����,FZ ';
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;    ti:=TADOQuery.Create(nil);ti.Connection:=DBe.DM.con1;
    AQ1.sql.Clear;aq2.SQL.Clear;aq3.SQL.Clear;aq4.SQL.Clear;
    AQ1.SQL.Add('select '+field1+' from '+TableName[0]+'  order by һ������ID ASC');
    AQ2.SQL.Add('select '+field2+' from '+TableName[1]+'  order by ��������ID ASC');
    AQ3.SQL.Add('select '+field3+' from '+TableName[2]+'  order by ��������ID ASC');
    AQ4.SQL.Add('select '+field4+' from '+TableName[3]+'  order by �ļ�����ID ASC');
    AQ1.open;  aq2.Open; aq3.Open;  aq4.Open;
    cou1:=AQ1.recordcount ;  cou2:=aq2.RecordCount;
    cou3:=aq3.RecordCount;   cou4:=aq4.RecordCount;
 //btn2.Enabled:=false;
    
    ti.SQL.Clear;
    sqls:='select '+field+' from '+Table;
    
    sqls:=sqls+THstr;//2012.4.18��Ӳ�ѯ��������Ŀ���������
{    if is2c then   //�����2��C ѡ������40����������46��ʼΪC��
    sqls:=sqls+' and ((XH>40 And XH<46) OR (XH<11))'
    else           //�����ǰ����C��Ҳ����3�����Ǿ��Ƕ���������Ŀ
    sqls:=sqls+' and ((XH>35 And XH<41) OR (XH<11))';
}    ti.SQL.Add(sqls); ti.Open;
    ticou:=ti.RecordCount;
    crow:=1;
    for i:=1 to ticou do
    begin
      for j:=1 to 6 do
      begin
        isbreak:=False;//����ÿ�������ʼ��û���ҵ�
        aq1.First;aq2.First;aq3.First;aq4.First;//ȫ��ָ���һ����¼
        str:=ti.fieldByname('����'+inttostr(j)).AsString;
        if str='' then Break;
        for k:=1 to cou1 do
         begin
           if str=getstr(aq1.fieldbyname('һ����������').AsString) then
              begin isbreak:=True;Break;end;   //�ҵ��˾Ͳ��ټ���
           aq1.Next;
         end;
         if isbreak then Continue;   //�ҵ��˾�ֱ������һ������
        for k:=1 to cou2 do
         begin
           if str=getstr(aq2.fieldbyname('������������').AsString) then
              begin isbreak:=True;Break;end;   //�ҵ��˾Ͳ��ټ���
           aq2.Next;
         end;
         if isbreak then Continue;   //�ҵ��˾�ֱ������һ������
        for k:=1 to cou3 do
         begin
           if str=getstr(aq3.fieldbyname('������������').AsString) then
              begin isbreak:=True;Break;end;   //�ҵ��˾Ͳ��ټ���
           aq3.Next;
         end;
         if isbreak then Continue;   //�ҵ��˾�ֱ������һ������
        for k:=1 to cou4 do
         begin
           if str=getstr(aq4.fieldbyname('�ļ���������').AsString) then
              begin isbreak:=True;Break;end;   //�ҵ��˾Ͳ��ټ���
           aq4.Next;
         end;
         if isbreak then Continue;   //�ҵ��˾�ֱ������һ������
         if crow=1 then
         begin
           try
              v:=CreateOleObject('Excel.Application');//�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
              v.visible:=false;
              v.WorkBooks.Add;   //�½�Excel   �ļ�
                Sheet:=v.WorkBooks[1].WorkSheets[1];
                v.workBooks[1].Styles['Normal'].Font.size:=10;
                sheet.range['a1'].select;
                Sheet.Cells[1,1]:= '���ֲ�һ�µĿ������£� ';//��Ԫ������
                Sheet.Cells[2,1]:= 'TH';
                Sheet.Cells[2,2]:= 'XH';
                Sheet.Cells[2,3]:= '������';
                crow:=crow+1;
           except
              showmessage( '��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
              v.DisplayAlerts:=false;//�Ƿ���ʾ����
              v.Quit;//�˳�Excel
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
    aq1.Free;aq2.Free;aq3.Free;aq4.Free;
    ti.First;
    biaotoufirst:=true;
    for i:=1 to ticou do
    begin
      if (ti.FieldByName('FZ').asInteger<=0) or (ti.FieldByName('FZ').asInteger>10) or (ti.FieldByName('����1����').AsFloat<=0) or (ti.FieldByName('����1����').AsFloat>1) then
      begin
         if crow=1 then
         begin
           try
              v:=CreateOleObject('Excel.Application');//�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
              v.Visible:=True;
              v.WorkBooks.Add;   //�½�Excel   �ļ�
                Sheet:=v.WorkBooks[1].WorkSheets[1];
                v.workBooks[1].Styles['Normal'].Font.size:=10;               
           except
              showmessage( '��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
              v.DisplayAlerts:=false;//�Ƿ���ʾ����
              v.Quit;//�˳�Excel
              exit;
           end;
         end; 
         if biaotoufirst then
         begin
            crow:=crow+1;
            sheet.cells[crow,1]:='���·�ֵ�򿼵����������:';
            crow:=crow+1;
            sheet.cells[crow,1]:='TH';
            sheet.cells[crow,2]:='XH';
            sheet.cells[crow,3]:='��ʾ��Ϣ';
            biaotoufirst:=false;
         end;
         crow:=crow+1;
         sheet.cells[crow,1]:=ti.FieldByName('TH').AsInteger;
         sheet.cells[crow,2]:=ti.FieldByName('XH').AsInteger;
         sheet.cells[crow,3]:='��������ֶλ�FZ�ֶ������⣬�������ݿ⣡';         
      end;
      ti.Next;
    end;
    ti.Free;
    if crow=1 then
    messagebox(0,'û���ҵ���һ�µĿ���'+#13+#13+
    '����������ֵ�Ƿ���д��ȷ,'+#13+#13+
    '�����Ƿ���©��Ŀ�������','��Ǹ',mb_ok);

    {Application.Restore;
    Application.BringToFront;}
end;

//�ȽϹ����������㣬���¼���Ƿ��г���
procedure CheckPublicTestPoint(THstr:string);
var
crow,i,j,k:Integer;
aq1,aq2,aq3,aq4,ti:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4,ticou:integer;//4����ļ�¼����
sqls:string;s:array of string;
isbreak,biaotoufirst:Boolean;
field,field1,field2,field3,field4:string;
PublicTable:array [0..3] of string;
begin
    PublicTable[0]:='��������֪ʶһ�������';
    PublicTable[1]:='��������֪ʶ���������';
    PublicTable[2]:='��������֪ʶ���������';
    PublicTable[3]:='��������֪ʶ�ļ������';
    if not(TableExist(DBe.DM.con1,PublicTable[1])) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+PublicTable[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       exit;
    end;
    Setlength(s,MAX+1);
    if Form1.rg1.itemindex=0 then table:='��Ŀ��'  else table:='ģ�����';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+Table+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       exit;
    end;
    field1:=' һ���������� ';
    field2:=' ������������ ';
    field3:=' ������������ ';
    field4:=' �ļ��������� ';
    field:=' TH,XH,����1,����2,����3,����4,����5,����6,����1����,FZ ';
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;    ti:=TADOQuery.Create(nil);ti.Connection:=DBe.DM.con1;
    AQ1.sql.Clear;aq2.SQL.Clear;aq3.SQL.Clear;aq4.SQL.Clear;
    AQ1.SQL.Add('select '+field1+' from '+PublicTable[0]+'  order by һ������ID ASC');
    AQ2.SQL.Add('select '+field2+' from '+PublicTable[1]+'  order by ��������ID ASC');
    AQ3.SQL.Add('select '+field3+' from '+PublicTable[2]+'  order by ��������ID ASC');
    AQ4.SQL.Add('select '+field4+' from '+PublicTable[3]+'  order by �ļ�����ID ASC');
    AQ1.open;  aq2.Open; aq3.Open;  aq4.Open;
    cou1:=AQ1.recordcount ;  cou2:=aq2.RecordCount;
    cou3:=aq3.RecordCount;   cou4:=aq4.RecordCount;
 //btn2.Enabled:=false;
    
    ti.SQL.Clear;
    sqls:='select '+field+' from '+Table;
    
    sqls:=sqls+THstr;//2012.4.18��Ӳ�ѯ��������Ŀ���������
    if is2c then   //�����2��C ѡ������40����������46��ʼΪC��
    sqls:=sqls+' and ((XH>40 And XH<46) OR (XH<11))'//2012.6.6����
    else           //�����ǰ����C��Ҳ����3�����Ǿ��Ƕ���������Ŀ
    sqls:=sqls+' and ((XH>35 And XH<41) OR (XH<11))';
    sqls:=sqls+' and (XH<11)';//20140701 ��ֽ�����Թ�������Ϊǰ10��
    ti.SQL.Add(sqls); ti.Open;
    ticou:=ti.RecordCount;
    crow:=1;
    for i:=1 to ticou do
    begin
      for j:=1 to 6 do
      begin
        isbreak:=False;//����ÿ�������ʼ��û���ҵ�
        aq1.First;aq2.First;aq3.First;aq4.First;//ȫ��ָ���һ����¼
        str:=ti.fieldByname('����'+inttostr(j)).AsString;
        if str='' then Break;
        for k:=1 to cou1 do
         begin
           if str=getstr(aq1.fieldbyname('һ����������').AsString) then
              begin isbreak:=True;Break;end;   //�ҵ��˾Ͳ��ټ���
           aq1.Next;
         end;
         if isbreak then Continue;   //�ҵ��˾�ֱ������һ������
        for k:=1 to cou2 do
         begin
           if str=getstr(aq2.fieldbyname('������������').AsString) then
              begin isbreak:=True;Break;end;   //�ҵ��˾Ͳ��ټ���
           aq2.Next;
         end;
         if isbreak then Continue;   //�ҵ��˾�ֱ������һ������
        for k:=1 to cou3 do
         begin
           if str=getstr(aq3.fieldbyname('������������').AsString) then
              begin isbreak:=True;Break;end;   //�ҵ��˾Ͳ��ټ���
           aq3.Next;
         end;
         if isbreak then Continue;   //�ҵ��˾�ֱ������һ������
        for k:=1 to cou4 do
         begin
           if str=getstr(aq4.fieldbyname('�ļ���������').AsString) then
              begin isbreak:=True;Break;end;   //�ҵ��˾Ͳ��ټ���
           aq4.Next;
         end;
         if isbreak then Continue;   //�ҵ��˾�ֱ������һ������
         if crow=1 then
         begin
           try
              v:=CreateOleObject('Excel.Application');//�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
              v.Visible:=True;
              v.WorkBooks.Add;   //�½�Excel   �ļ�
                Sheet:=v.WorkBooks[1].WorkSheets[1];
                v.workBooks[1].Styles['Normal'].Font.size:=10;
                sheet.range['a1'].select;
                Sheet.Cells[1,1]:= '���ֲ�һ�µĿ������£� ';//��Ԫ������
                Sheet.Cells[2,1]:= 'TH';
                Sheet.Cells[2,2]:= 'XH';
                Sheet.Cells[2,3]:= '������';
                crow:=crow+1;
           except
              showmessage( '��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
              v.DisplayAlerts:=false;//�Ƿ���ʾ����
              v.Quit;//�˳�Excel
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
    aq1.Free;aq2.Free;aq3.Free;aq4.Free;
    ti.First;
    biaotoufirst:=true;
    for i:=1 to ticou do
    begin
      if (ti.FieldByName('FZ').asInteger<=0) or (ti.FieldByName('FZ').asInteger>10) or (ti.FieldByName('����1����').AsFloat<=0) or (ti.FieldByName('����1����').AsFloat>1) then
      begin
         if crow=1 then
         begin
           try
              v:=CreateOleObject('Excel.Application');//�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
              v.Visible:=True;
              v.WorkBooks.Add;   //�½�Excel   �ļ�
                Sheet:=v.WorkBooks[1].WorkSheets[1];
                v.workBooks[1].Styles['Normal'].Font.size:=10;               
           except
              showmessage( '��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
              v.DisplayAlerts:=false;//�Ƿ���ʾ����
              v.Quit;//�˳�Excel
              exit;
           end;
         end; 
         if biaotoufirst then
         begin
            crow:=crow+1;
            sheet.cells[crow,1]:='���·�ֵ�򿼵����������:';
            crow:=crow+1;
            sheet.cells[crow,1]:='TH';
            sheet.cells[crow,2]:='XH';
            sheet.cells[crow,3]:='��ʾ��Ϣ';
            biaotoufirst:=false;
         end;
         crow:=crow+1;
         sheet.cells[crow,1]:=ti.FieldByName('TH').AsInteger;
         sheet.cells[crow,2]:=ti.FieldByName('XH').AsInteger;
         sheet.cells[crow,3]:='��������ֶλ�FZ�ֶ������⣬�������ݿ⣡';         
      end;
      ti.Next;
    end;
    ti.Free;
    if crow=1 then
    messagebox(0,'û���ҵ���һ�µĿ���'+#13+#13+
    '����������ֵ�Ƿ���д��ȷ,'+#13+#13+
    '�����Ƿ���©��Ŀ�������','��Ǹ',mb_ok);

    {Application.Restore;
    Application.BringToFront;}
end;

//���ɵ�������100
procedure TForm1.btn1Click(Sender: TObject);
var
row1,row2,row3,crow,i,j,k,l,m,counter:Integer;
big,litle,gailv:Double;
avg:array[1..4] of Double;
aq1,aq2,aq3,aq4:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4:integer;//4����ļ�¼����
sumer:array of array of Double;
totle:array of double;
field,field1,field2,field3,field4:string;
TableXiaBiao:string;//20140701 �кŴ���26ʱ�����һ����ĸ��AA��AB��AC...
 first:boolean; sqlstr:string;
begin
    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+TableName[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    SetLength(sumer,5,MAX+1);SetLength(totle,MAX+1);
    if rg1.itemindex=0 then table:='��Ŀ��'  else table:='ģ�����';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+Table+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    TableXiaBiao:='';
    i:=ord('a')+MAX+2;  //ȷ���еķ�Χ
    if i>122 then    //20140701��ĸz������122
    begin
      Tablexiabiao:='a'+chr(97+i-123);
    end
    else
    TableXiaBiao:=chr(i);
    try
      v:=CreateOleObject('Excel.Application');//�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
      v.Visible:=false;
                v.WorkBooks.Add;   //�½�Excel   �ļ�
                //v.WorkBooks[1].WorkSheets[1].name:= '���Ա�';//��һҳ����
                //v.WorkBooks[1].WorkSheets[2].name:= '�����԰';
                //v.WorkBooks[1].WorkSheets[3].name:= '������ѽ ';
                v.activesheet.columns[1].columnwidth:=20;  //��1�п�
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
                sheet.cells[2,1]:='�ϼ�';
                v.range['b:'+TableXiaBiao].FormatConditions.Delete;
                v.range['b:'+TableXiaBiao].FormatConditions.Add(1, 3,'0');
                v.range['b:'+TableXiaBiao].FormatConditions[1].Font.ColorIndex := 15;
                v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Delete;
                if is3ji then begin
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Add(1,1,'0','99.995');
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions[1].Font.ColorIndex := 3;
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Add(1,7,'100.005');end
                else begin
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Add(1,1,'0','29.995');
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions[1].Font.ColorIndex := 3;
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Add(1,7,'30.005');
                end;
                v.range['b2:'+TableXiaBiao+'2'].FormatConditions[2].Font.ColorIndex := 3;
                v.range['b:'+TableXiaBiao].columnwidth:=5;
                sheet.range['b3'].select;
                v.ActiveWindow.FreezePanes:= True;
                sheet.range['a1'].select;
                //Sheet.Cells[1,1]:= '�ÿ� ';//��Ԫ������
                //Sheet.Cells[1,2]:= 'ȷʵ ';
                //Sheet.Cells[2,1]:= '��ϲ�� ';
{                v.activesheet.rows[2].select;
                v.Selection.Interior.ColorIndex:= 32;      
                v.range['b2:'+TableXiaBiao+'2'].interior.colorindex:=39;//��ɫ
 }
    except
        on e:exception do
        begin
        ShowMessage(e.Message);
        showmessage( '��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
        v.DisplayAlerts:=false;//�Ƿ���ʾ����
        v.Quit;//�˳�Excel
        exit;
        end;
    end;
    {Application.Restore;
    Application.BringToFront;}
    field1:=' һ���������� ';
    field2:=' ������������ ';
    field3:=' ������������ ';
    field4:=' �ļ��������� ';
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+' from '+TableName[0]+'  order by һ������ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
 btn1.Enabled:=false;
 counter:=0;//������ʼ��Ϊ0
 avg[1]:=0;avg[2]:=0;avg[3]:=0;avg[4]:=0;//ƽ��ֵ��ʼ��Ϊ0
 with Form1 do
 begin
   for i:=1 to MAX do
   Sheet.Cells[1,i+1]:=IntToStr(i);
   Sheet.cells[1,MAX+2]:='ƽ��ֵ';
   Sheet.cells[1,MAX+3]:='���˸���';

  crow:=2;     //�ӵ�2�п�ʼ�����2���Ǻϼƣ�
  for i:=1 to MAX do totle[i]:=0; //��ϼ�ֵ�ı�����ʼ��Ϊ0
  for i:=1 to cou1 do
  begin
    caption:='��'+inttostr(i)+'�� ������...';
    for m:=1 to MAX do sum[1,m]:=0;
    crow:=crow+1;//inc(row);//��һ��
    row1:=crow;
    v.range['a'+inttostr(crow)+':'+TableXiaBiao+inttostr(crow)].interior.colorindex:=24;
    str:= aq1.fieldbyname('һ����������').asstring;
    Sheet.cells[crow,1]:=' '+aq1.fieldbyname('һ����������').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);  ////�ȽϿ�������ֵ////
    for j:=1 to MAX do
      Sheet.cells[crow,j+1]:=sum[1,j];
    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+TableName[1]+' where [������������һ������ID]='+inttostr(i)+' order by ��������ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     crow:=crow+1;    row2:=crow;
     Sheet.cells[crow,1]:='     '+aq2.fieldbyname('������������').asstring;
     str:=aq2.fieldbyname('������������').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);
     for k:=1 to MAX do Sheet.cells[crow,k+1]:=sum[2,k];
     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+TableName[2]+' where ([������������һ������ID]='+inttostr(i)+
                ') and ([��������������������ID]='+inttostr(j)+') order by ��������ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
       crow:=crow+1;    row3:=crow;
       Sheet.cells[crow,1]:='         '+aq3.fieldbyname('������������').asstring;
       str:=aq3.fieldbyname('������������').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);   ////�ȽϿ�������ֵ////

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+TableName[3]+' where ([�ļ���������һ������ID]='+inttostr(i)+
                ') and ([�ļ�����������������ID]='+inttostr(j)+
                ') and ([�ļ�����������������ID]='+inttostr(k)+
                ') order by �ļ�����ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
        crow:=crow+1;
        Sheet.cells[crow,1]:='             '+aq4.fieldbyname('�ļ���������').asstring;
        str:=aq4.fieldbyname('�ļ���������').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);      ////�ȽϿ�������ֵ////
        for m:=1 to MAX do begin
          Sheet.cells[crow,m+1]:=sum[4,m];
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          avg[4]:=avg[4]+sum[4,m];        //�����ܺϣ�׼������ƽ��ֵ
          if sum[4,m]>0 then counter:=counter+1; //����0�ͼ����������������
          end;
         Sheet.cells[crow,MAX+2]:=avg[4]/MAX;//�ļ������ƽ��ֵ
         Sheet.cells[crow,MAX+3]:=counter/MAX;//�ļ�����ĸ���
         avg[4]:=0;counter:=0;//���³�ʼΪ0
        aq4.Next;
       end;//��4���������3���ֵ�ǵ�3���ֵ�����е�4��ĺ�
       for l:=1 to MAX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          Sheet.cells[row3,l+1]:=sum[3,l];
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          avg[3]:=avg[3]+sum[3,l];        //�����ܺϣ�׼������ƽ��ֵ
          if sum[3,l]>0 then counter:=counter+1; //����0�ͼ����������������
          end;
         Sheet.cells[row3,MAX+2]:=avg[3]/MAX;//���������ƽ��ֵ
         Sheet.cells[row3,MAX+3]:=counter/MAX;//��������ĸ���
         avg[3]:=0;counter:=0;//���³�ʼΪ0
       aq3.Next;
     end; //��3�����
     for l:=1 to MAX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          Sheet.cells[row2,l+1]:=sum[2,l];
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          avg[2]:=avg[2]+sum[2,l];        //�����ܺϣ�׼������ƽ��ֵ
          if sum[2,l]>0 then counter:=counter+1; //����0�ͼ����������������
          end;
         Sheet.cells[row2,MAX+2]:=avg[2]/MAX;//���������ƽ��ֵ
         Sheet.cells[row2,MAX+3]:=counter/MAX;//��������ĸ���
         avg[2]:=0;counter:=0;//���³�ʼΪ0
     aq2.Next;
    end;
    for l:=1 to MAX do begin              //����ϼ�
          sum[1,l]:=sum[1,l]+sumer[2,l];
          Sheet.cells[row1,l+1]:=sum[1,l];
          totle[l]:=totle[l]+sum[1,l];
          Sheet.cells[2,l+1]:=totle[l];
          avg[1]:=avg[1]+sum[1,l];  //�����ֵ��ӣ�׼������ƽ��
          if sum[1,l]>0 then counter:=counter+1; //����0��ֵ�ͼ�����׼���������
          end;
       Sheet.cells[row1,MAX+2]:=avg[1] / MAX;//һ�������ƽ��ֵ
       Sheet.cells[row1,MAX+3]:=counter/MAX; //һ������Ŀ��˸���
       avg[1]:=0;counter:=0;//���³�ʼ��Ϊ0
   aq1.Next;
   end;
     aq4.Free;
     aq3.free;
     aq2.Free;
     aq1.Free;
 end;
   caption:='���';
   v.visible:=true;//��ʾ��
   if is3ji then begin big:=100.005;litle:=99.995;end
     else begin big:=30.005;litle:=29.995;end;
   first:=true;   //��־
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
          11: chk11.checked:=True;
        end;
        if first then
         begin
          first:=false;
          sqlstr:=' where (th='+inttostr(i);
         end else
          sqlstr:=sqlstr+' or th='+inttostr(i);
      end;
   end;
   if not first then    //�в���Ԥ�ڵ�����
     begin
      sqlstr:=sqlstr+')';
      if not is3ji then
{         if is2c then
           sqlstr:=sqlstr+' and ((XH>10 And XH<41) OR (XH>45))'
         else           //�����ǰ����C��Ҳ����3�����Ǿ��Ƕ���������Ŀ
           sqlstr:=sqlstr+' and ((XH>10 And XH<) OR (XH>40))';
}        sqlstr:=sqlstr+' and (XH>10)';///20140701��ֽ������ֻ��ѡ���⡣40����ǰ10���ǹ�������
         Application.Restore;
         Application.BringToFront;
      if MessageBox(0,'���ֲ���г�����ݣ�Ҫ�����ó���ɹһɹ��','ע��',MB_ICONINFORMATION+MB_YESNO+MB_DEFBUTTON1)=mryes then
      CheckTestPoint(sqlstr);
     end else
     messagebox(0,'ȫ�����ݼ�������������','������ݵ���',MB_OK);
{    if chk1.Checked or chk2.Checked or chk3.Checked or chk4.Checked or
       chk5.Checked or chk6.Checked or chk7.Checked or chk8.Checked or
       chk9.Checked or chk10.Checked or chk11.Checked then
       begin
         if MessageBox(0,'���ֲ���г�����ݣ�Ҫ�����ó���ɹһɹ��','ע��',MB_ICONINFORMATION+MB_YESNO+MB_DEFBUTTON1)=mryes then
           begin
             grp1.Visible:=true;
             btn3.Enabled:=True;
             btn4.Enabled:=False;
             btn3Click(Sender);
           end;
       end else begin
    messagebox(0,'ȫ�����ݼ�������������','������ݵ���',MB_OK);
    end;
}    btn1.Enabled:=true;
     v.Visible:=True;
   
end;


procedure TForm1.Timer1Timer(Sender: TObject);
var
  table:array [0..4] of string;
  i:Integer;
  s1,s2:string;
  aq:TADOQuery;
begin
  Timer1.Enabled:=False;
  s1:='����Ҫ�������ǣ�  '+#13+#13+'    ';
  table[0]:='����C����һ�������';
  table[1]:='����VBһ�������';
  table[2]:='����VFһ�������';
  table[3]:='����Accessһ�������';
  table[4]:='�������缼��һ�������';
  if TableExist(DBE.DM.con1,'��Ŀ��') then
     begin 
       aq:=TADOQuery.Create(nil);
       aq.Connection:=DBe.DM.con1;
       AQ.sql.Clear;
       AQ.SQL.Add('select TH  from ��Ŀ��  Group by TH');
       AQ.open;
       MAX:=AQ.recordcount ;
       aq.free;
       Setlength(sum,5,MAX+1);
     end
     else MessageBox(0,'û���ҵ���Ŀ���������ݿ�','û��Ŀ��',MB_OK);
  for i:=0 to 4 do
   begin
    if TableExist(DBe.DM.con1,table[i]) then Break;
   end;
  case i of
     0:begin
       TableName[0]:='����C����һ�������';
       TableName[1]:='����C���Զ��������';
       TableName[2]:='����C�������������';
       TableName[3]:='����C�����ļ������';s2:='����C����';
       is2c:=True;
       end;
     1:begin
       TableName[0]:='����VBһ�������';
       TableName[1]:='����VB���������';
       TableName[2]:='����VB���������';
       TableName[3]:='����VB�ļ������';s2:='����VB';
       end;
     2:begin
       TableName[0]:='����VFһ�������';
       TableName[1]:='����VF���������';
       TableName[2]:='����VF���������';
       TableName[3]:='����VF�ļ������';s2:='����VF';
       end;
     3:begin
       TableName[0]:='����Accessһ�������';
       TableName[1]:='����Access���������';
       TableName[2]:='����Access���������';
       TableName[3]:='����Access�ļ������';s2:='����Access';
       end;
     4:begin
       TableName[0]:='�������缼��һ�������';
       TableName[1]:='�������缼�����������';
       TableName[2]:='�������缼�����������';
       TableName[3]:='�������缼���ļ������';s2:='�������缼��';
       is3ji:=True;
       end;
     else
       MessageBox(0,'���ݿ����,û���ҵ����õı�,�������˳�.','������ݿ�',MB_OK);
       Halt;
   end;

   label1.Caption:=s1+s2;
   end;
//���ļ������4������
procedure TForm1.Button1Click(Sender: TObject);
var
row1,row2,row3,row4,i,j,k,l,m,counter:Integer;
aq1,aq2,aq3,aq4:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4:integer;//4����ļ�¼����
sumer:array of array of Double;
totle:array of double;
sht:array[1..4]of variant;
field1,field2,field3,field4:string;
big,litle:double;
avg:array[1..4]of Double;
TableXiaBiao:string;//20140701 �кŴ���26ʱ�����һ����ĸ��AA��AB��AC...
first:boolean; sqlstr:string;
begin
    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+TableName[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    SetLength(sumer,5,MAX+1);SetLength(totle,MAX+1);
    if rg1.itemindex=0 then table:='��Ŀ��'  else table:='ģ�����';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+Table+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    TableXiaBiao:='a';
    i:=ord('a')+MAX+2;  //ȷ���еķ�Χ
    if i>122 then    //20140701��ĸz������122
    begin
      Tablexiabiao:='a'+chr(97+i-123);
    end
    else
    TableXiaBiao:=chr(i);
    try
      v:=CreateOleObject('Excel.Application');//�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
      v.visible:=false;
                v.WorkBooks.Add;   //�½�Excel   �ļ�
                v.sheets.add(after:=v.sheets[3]);
                v.workBooks[1].Styles['Normal'].Font.size:=10;
                for i:=1 to 4 do
                begin
                  v.workBooks[1].WorkSheets[i].columns[1].columnwidth:=30;
                  v.workBooks[1].WorkSheets[i].select;
                  sht[i]:=v.WorkBooks[1].WorkSheets[i];
                  sht[i].range['b2'].select;
                  v.ActiveWindow.FreezePanes:= True;
                  v.range['b:'+TableXiaBiao].columnwidth:=5;
                  for j:=1 to MAX do
                  Sht[i].Cells[1,j+1]:=IntToStr(j);
                  sht[i].cells[1,MAX+2]:='ƽ��ֵ';
                  Sht[i].Cells[1,MAX+3]:='���˸���';
                  Sht[i].range['b:'+TableXiaBiao].FormatConditions.Delete;
                  Sht[i].range['b:'+TableXiaBiao].FormatConditions.Add(1, 3,'0');
                  sht[i].range['b:'+TableXiaBiao].FormatConditions[1].Font.ColorIndex := 15;
                end;
                v.workBooks[1].WorkSheets[1].select;
                sht[1].range['a1'].select;
    except
        showmessage( '��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
        v.DisplayAlerts:=false;//�Ƿ���ʾ����
        v.Quit;//�˳�Excel
        exit;
    end;
    {Application.Restore;
    Application.BringToFront;}
    field1:=' һ���������� ';
    field2:=' ������������ ';
    field3:=' ������������ ';
    field4:=' �ļ��������� ';

    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+'  from '+TableName[0]+'  order by һ������ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
 button1.Enabled:=false;
 counter:=0;avg[1]:=0;avg[2]:=0;avg[3]:=0;avg[4]:=0;//���ʼ�����ƽ��ֵ��ʼΪ0
 for i:=1 to cou1 do
   begin
     sht[1].cells[i+1,1]:=aq1.fieldbyname('һ����������').asstring;
     aq1.Next;
   end;
  aq1.First;
  v.range['a'+inttostr(cou1+2)].HorizontalAlignment := 3;
  sht[1].cells[cou1+2,1]:='�ϼ�';
 with Form1 do
 begin
  row1:=1;row2:=1;row3:=1;row4:=1;
  //crow:=1;  //for i:=1 to MAX do totle[i]:=0;//���ºϼ�
  for i:=1 to cou1 do
  begin
    caption:='��'+inttostr(i)+'�� ������...';
    for m:=1 to MAX do sum[1,m]:=0;
    row1:=row1+1;//inc(row);

    v.range['a'+inttostr(row1)+':'+TableXiaBiao+inttostr(row1)].interior.colorindex:=24;
    str:= aq1.fieldbyname('һ����������').asstring;
    //sht[1].cells[row1,1]:=aq1.fieldbyname('һ����������').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);   ////�ȽϿ�������ֵ////
    for j:=1 to MAX do    sht[1].cells[row1,j+1]:=sum[1,j]/100;    //���¸��׺ϼ�
    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+TableName[1]+' where [������������һ������ID]='+inttostr(i)+' order by ��������ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     row2:=row2+1;    //row2:=crow;
     sht[2].cells[row2,1]:=aq2.fieldbyname('������������').asstring;
     str:=aq2.fieldbyname('������������').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);  ////�ȽϿ�������ֵ////
     for k:=1 to MAX do Sht[2].cells[row2,k+1]:=sum[2,k]/100; //������С�ڸ��׺ϼ�
     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+TableName[2]+' where ([������������һ������ID]='+inttostr(i)+
                ') and ([��������������������ID]='+inttostr(j)+') order by ��������ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
       row3:=row3+1;    //row3:=crow;
       sht[3].cells[row3,1]:=aq3.fieldbyname('������������').asstring;
       str:=aq3.fieldbyname('������������').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);  ////�ȽϿ�������ֵ////

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+TableName[3]+' where ([�ļ���������һ������ID]='+inttostr(i)+
                ') and ([�ļ�����������������ID]='+inttostr(j)+
                ') and ([�ļ�����������������ID]='+inttostr(k)+
                ') order by �ļ�����ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
        row4:=row4+1;
        sht[4].cells[row4,1]:=aq4.fieldbyname('�ļ���������').asstring;
        str:=aq4.fieldbyname('�ļ���������').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);   ////�ȽϿ�������ֵ////
        for m:=1 to MAX do begin
            sht[4].cells[row4,m+1]:=sum[4,m]/100;
            sumer[4,m]:=sumer[4,m]+sum[4,m];
            avg[4]:=avg[4]+sum[4,m];//ÿ�еĸ���ֵ�ĺ�
            if sum[4,m]>0 then counter:=counter+1;//�������0�ͼ�����׼���������
           end;
           sht[4].cells[row4,MAX+2]:=avg[4]/MAX; //ƽ��ֵ
           sht[4].cells[row4,MAX+3]:=counter/MAX;//����
           avg[4]:=0;counter:=0;//���³�ʼ��Ϊ0��׼����һ������ļ���
        aq4.Next;
       end;//��4���������3���ֵ�ǵ�3���ֵ�����е�4��ĺ�
       for l:=1 to MAX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          sht[3].cells[row3,l+1]:=sum[3,l]/100;
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          avg[3]:=avg[3]+sum[3,l];//ÿ�еĸ���ֵ�ĺ�
          if sum[3,l]>0 then counter:=counter+1;//�������0�ͼ�����׼���������
          end;
         sht[3].cells[row3,MAX+2]:=avg[3]/MAX; //ƽ��ֵ
         sht[3].cells[row3,MAX+3]:=counter/MAX;//����
         avg[3]:=0;counter:=0;//���³�ʼ��Ϊ0��׼����һ������ļ���
       aq3.Next;
     end; //��3�����
        for l:=1 to MAX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          sht[2].cells[row2,l+1]:=sum[2,l]/100;
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          avg[2]:=avg[2]+sum[2,l];//ÿ�еĸ���ֵ�ĺ�
          if sum[2,l]>0 then counter:=counter+1;//�������0�ͼ�����׼���������
        end;
        sht[2].cells[row2,MAX+2]:=avg[2]/MAX; //ƽ��ֵ
        sht[2].cells[row2,MAX+3]:=counter/MAX;//����
        avg[2]:=0;counter:=0;//���³�ʼ��Ϊ0��׼����һ������ļ���
     aq2.Next;
    end;
        for l:=1 to MAX do begin
          sum[1,l]:=sum[1,l]+sumer[2,l];
          sht[1].cells[row1,l+1]:=sum[1,l]/100;
          totle[l]:=totle[l]+sum[1,l];
          Sht[1].cells[cou1+2,l+1]:=totle[l];
          avg[1]:=avg[1]+sum[1,l];//ÿ�еĸ���ֵ�ĺ�
          if sum[1,l]>0 then counter:=counter+1;//�������0�ͼ�����׼���������
        end;
        sht[1].cells[row1,MAX+2]:=avg[1]/MAX; //ƽ��ֵ
        sht[1].cells[row1,MAX+3]:=counter/MAX;//����
        avg[1]:=0;counter:=0;//���³�ʼ��Ϊ0��׼����һ������ļ���
   aq1.Next;
   end;
   aq4.Free;
   aq3.free;
   aq2.Free;
   aq1.Free;
 end;
    caption:='���';
   v.visible:=true;//��ʾ��

   if is3ji then begin big:=100.005;litle:=99.995;end
     else begin big:=30.005;litle:=29.995;end;
   first:=true;   //��־
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
          11: chk11.checked:=True;
        end;
        if first then
         begin
          first:=false;
          sqlstr:=' where (th='+inttostr(i);
         end else
          sqlstr:=sqlstr+' or th='+inttostr(i);
      end;
   end;
   if not first then    //�в���Ԥ�ڵ�����
     begin
      sqlstr:=sqlstr+')';
      if not is3ji then
         {if is2c then
           sqlstr:=sqlstr+' and ((XH>10 And XH<41) OR (XH>45))'
         else           //�����ǰ����C��Ҳ����3�����Ǿ��Ƕ���������Ŀ
           sqlstr:=sqlstr+' and ((XH>10 And XH<36) OR (XH>40))';
         }
         sqlstr:=sqlstr+' and (XH>10)';   //20140701 ��ֽ��������������ǰ10��
         Application.Restore;
         Application.BringToFront;
      if MessageBox(0,'���ֲ���г�����ݣ�Ҫ�����ó���ɹһɹ��','ע��',MB_ICONINFORMATION+MB_YESNO+MB_DEFBUTTON1)=mryes then
      CheckTestPoint(sqlstr);
     end else
     messagebox(0,'ȫ�����ݼ�������������','������ݵ���',MB_OK);
 Button1.Enabled:=true;
end;


//�ҳ��뿼��Ŀ¼��һ�µĿ���
procedure TForm1.btn3Click(Sender: TObject);
var
crow,i,j,k:Integer;
aq1,aq2,aq3,aq4,ti:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4,ticou:integer;//4����ļ�¼����
sqls:string;s:array of string;
isbreak:Boolean;
field,field1,field2,field3,field4:string;
begin
    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+TableName[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    Setlength(s,MAX+1);
    if rg1.itemindex=0 then table:='��Ŀ��'  else table:='ģ�����';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+Table+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    field1:=' һ���������� ';
    field2:=' ������������ ';
    field3:=' ������������ ';
    field4:=' �ļ��������� ';
    field:=' TH,XH,����1,����2,����3,����4,����5,����6 ';
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;    ti:=TADOQuery.Create(nil);ti.Connection:=DBe.DM.con1;
    AQ1.sql.Clear;aq2.SQL.Clear;aq3.SQL.Clear;aq4.SQL.Clear;
    AQ1.SQL.Add('select '+field1+' from '+TableName[0]+'  order by һ������ID ASC');
    AQ2.SQL.Add('select '+field2+' from '+TableName[1]+'  order by ��������ID ASC');
    AQ3.SQL.Add('select '+field3+' from '+TableName[2]+'  order by ��������ID ASC');
    AQ4.SQL.Add('select '+field4+' from '+TableName[3]+'  order by �ļ�����ID ASC');
    AQ1.open;  aq2.Open; aq3.Open;  aq4.Open;
    cou1:=AQ1.recordcount ;  cou2:=aq2.RecordCount;
    cou3:=aq3.RecordCount;   cou4:=aq4.RecordCount;
 //btn2.Enabled:=false;
    if not(chk1.Checked or chk2.Checked or chk3.Checked or
           chk4.Checked or chk5.Checked or chk6.Checked or
           chk7.Checked or chk8.Checked or chk9.Checked or
           chk10.Checked or chk11.checked) then
           begin
             MessageBox(0,'��ѡ��Ҫ�˶Ե��׺�','��ʾ',MB_OK);
             exit;
           end;
    ti.SQL.Clear;
    sqls:='select '+field+' from '+Table+' where (';
    i:=1;//�ַ������鿪ʼֵ
    if chk1.Checked then  begin s[i]:= ' TH=1 ';i:=i+1; end;
    if chk2.Checked then  begin s[i]:= ' TH=2 ';i:=i+1; end;
    if chk3.Checked then  begin s[i]:= ' TH=3 ';i:=i+1; end;
    if chk4.Checked then  begin s[i]:= ' TH=4 ';i:=i+1; end;
    if chk5.Checked then  begin s[i]:= ' TH=5 ';i:=i+1; end;
    if chk6.Checked then  begin s[i]:= ' TH=6 ';i:=i+1; end;
    if chk7.Checked then  begin s[i]:= ' TH=7 ';i:=i+1; end;
    if chk8.Checked then  begin s[i]:= ' TH=8 ';i:=i+1; end;
    if chk9.Checked then  begin s[i]:= ' TH=9 ';i:=i+1; end;
    if chk10.Checked then  begin s[i]:=' TH=10 ';i:=i+1; end;
    if chk11.Checked then  begin s[i]:=' TH=11 ';i:=i+1; end;
    sqls:=sqls+s[1];
    if (i-1)>1 then
    for j:=2 to i-1 do
     begin
       sqls:=sqls+'OR'+s[j];
     end;
    sqls:=sqls+')';//2012.4.18��Ӳ�ѯ��������Ŀ���������
    if is2c then   //�����2��C ѡ������40����������46��ʼΪC��
    sqls:=sqls+' and ((XH>10 And XH<41) OR (XH>45))'
    else           //�����ǰ����C��Ҳ����3�����Ǿ��Ƕ���������Ŀ
    if not is3ji then sqls:=sqls+' and ((XH>10 And XH<36) OR (XH>40))';
    ti.SQL.Add(sqls); ti.Open;
    ticou:=ti.RecordCount;
    crow:=1;
    for i:=1 to ticou do
    begin
      for j:=1 to 6 do
      begin
        isbreak:=False;//����ÿ�������ʼ��û���ҵ�
        aq1.First;aq2.First;aq3.First;aq4.First;//ȫ��ָ���һ����¼
        str:=ti.fieldByname('����'+inttostr(j)).AsString;
        if str='' then Break;
        for k:=1 to cou1 do
         begin
           if str=getstr(aq1.fieldbyname('һ����������').AsString) then
              begin isbreak:=True;Break;end;   //�ҵ��˾Ͳ��ټ���
           aq1.Next;
         end;
         if isbreak then Continue;   //�ҵ��˾�ֱ������һ������
        for k:=1 to cou2 do
         begin
           if str=getstr(aq2.fieldbyname('������������').AsString) then
              begin isbreak:=True;Break;end;   //�ҵ��˾Ͳ��ټ���
           aq2.Next;
         end;
         if isbreak then Continue;   //�ҵ��˾�ֱ������һ������
        for k:=1 to cou3 do
         begin
           if str=getstr(aq3.fieldbyname('������������').AsString) then
              begin isbreak:=True;Break;end;   //�ҵ��˾Ͳ��ټ���
           aq3.Next;
         end;
         if isbreak then Continue;   //�ҵ��˾�ֱ������һ������
        for k:=1 to cou4 do
         begin
           if str=getstr(aq4.fieldbyname('�ļ���������').AsString) then
              begin isbreak:=True;Break;end;   //�ҵ��˾Ͳ��ټ���
           aq4.Next;
         end;
         if isbreak then Continue;   //�ҵ��˾�ֱ������һ������
         if crow=1 then
         begin
           try
              v:=CreateOleObject('Excel.Application');//�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
              v.Visible:=True;
              v.WorkBooks.Add;   //�½�Excel   �ļ�
                Sheet:=v.WorkBooks[1].WorkSheets[1];
                v.workBooks[1].Styles['Normal'].Font.size:=10;
                sheet.range['a1'].select;
                Sheet.Cells[1,1]:= '���ֲ�һ�µĿ������£� ';//��Ԫ������
                Sheet.Cells[2,1]:= 'TH';
                Sheet.Cells[2,2]:= 'XH';
                Sheet.Cells[2,3]:= '������';
                crow:=crow+1;
           except
              showmessage( '��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
              v.DisplayAlerts:=false;//�Ƿ���ʾ����
              v.Quit;//�˳�Excel
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
    messagebox(0,'û���ҵ���һ�µĿ���'+#13+#13+
    '����������ֵ�Ƿ���д��ȷ,'+#13+#13+
    '�����Ƿ���©��Ŀ�������','��Ǹ',mb_ok);

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
  Screen.HintFont.Name:= '����_GB2312';
  Screen.HintFont.Size:=12;
  application.HintHidePause:=10000;
end;
function numtohz(num:Integer):string;
begin
  case num of
    1: numtohz:='һ';
    2: numtohz:='��';
    3: numtohz:='��';
    4: numtohz:='��';
    5: numtohz:='��';
    6: numtohz:='��';
    7: numtohz:='��';
    8: numtohz:='��';
    9: numtohz:='��';
    end;
end;
//���������ֱ��д�����ݿ�
procedure TForm1.btn5Click(Sender: TObject);
var
i,j,k,l,m:Integer;
aq1,aq2,aq3,aq4:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4:integer;//4����ļ�¼����
sumer:array of array of Double;
//zt:array[1..9]of string;
field1,field2,field3,field4:string;
begin
    if rg1.ItemIndex=0 then Table:='��Ŀ��' else table:='ģ�����';
    if MessageBox(0,PCHar('������д���ݿ���ģ����������������'+#13+#13+'������Դ�ǡ�'+Table+'����ȷ����'),'����',MB_YESNO or MB_DEFBUTTON2	or MB_ICONWARNING)=mrno then Exit;
    if MessageBox(0,'ȷ��һ���Լ��϶�Ҫ��д���ݿ���','���潫����ԭ������',MB_YESNO or MB_DEFBUTTON2	or MB_ICONWARNING)=mrno then Exit;
    if MessageBox(0,'�������޷�����������','���棡��',MB_YESNO or MB_DEFBUTTON2	or MB_ICONWARNING)=mrno then Exit;

    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+TableName[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;

    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+Table+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    SetLength(sumer,5,MAX+1);
    {Application.Restore;
    Application.BringToFront;}
//      field1:='*';
//      field2:='*';
//      field3:='*';
//      field4:='*';
{    if is2c then begin
      zt[1]:='2008��4�·�ֵ����';zt[2]:='2008��9�·�ֵ����';
      zt[3]:='2009��3�·�ֵ����';zt[4]:='2009��9�·�ֵ����';
      zt[5]:='2010��3�·�ֵ����';zt[6]:='2010��9�·�ֵ����';
      zt[7]:='2011��3�·�ֵ����';MuX:=7;
      field1:=' ���,һ����������,[2008��4�·�ֵ����],[2008��9�·�ֵ����],[2009��3�·�ֵ����],[2009��9�·�ֵ����],[2010��3�·�ֵ����],[2010��9�·�ֵ����],[2011��3�·�ֵ����] ';
      field2:=' ���,������������,[2008��4�·�ֵ����],[2008��9�·�ֵ����],[2009��3�·�ֵ����],[2009��9�·�ֵ����],[2010��3�·�ֵ����],[2010��9�·�ֵ����],[2011��3�·�ֵ����] ';
      field3:=' ���,������������,[2008��4�·�ֵ����],[2008��9�·�ֵ����],[2009��3�·�ֵ����],[2009��9�·�ֵ����],[2010��3�·�ֵ����],[2010��9�·�ֵ����],[2011��3�·�ֵ����] ';
      field4:=' ���,�ļ���������,[2008��4�·�ֵ����],[2008��9�·�ֵ����],[2009��3�·�ֵ����],[2009��9�·�ֵ����],[2010��3�·�ֵ����],[2010��9�·�ֵ����],[2011��3�·�ֵ����] ';
      end
      else begin
      zt[1]:='2007��4�·�ֵ����';zt[2]:='2007��9�·�ֵ����';
      zt[3]:='2008��4�·�ֵ����';zt[4]:='2008��9�·�ֵ����';
      zt[5]:='2009��3�·�ֵ����';zt[6]:='2009��9�·�ֵ����';
      zt[7]:='2010��3�·�ֵ����';zt[8]:='2010��9�·�ֵ����';
      zt[9]:='2011��3�·�ֵ����';//MuX:=9;
}      field1:=' ���,һ����������,��һ��ģ�����ֵ����,�ڶ���ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,�ڰ���ģ�����ֵ����,�ھ���ģ�����ֵ���� ';
      field2:=' ���,������������,��һ��ģ�����ֵ����,�ڶ���ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,�ڰ���ģ�����ֵ����,�ھ���ģ�����ֵ���� ';
      field3:=' ���,������������,��һ��ģ�����ֵ����,�ڶ���ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,�ڰ���ģ�����ֵ����,�ھ���ģ�����ֵ���� ';
      field4:=' ���,�ļ���������,��һ��ģ�����ֵ����,�ڶ���ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,������ģ�����ֵ����,�ڰ���ģ�����ֵ����,�ھ���ģ�����ֵ���� ';
      //end;//�����select����š����ᷢ�����в���֮��Ĵ���
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select *  from '+TableName[0]+'  order by һ������ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
    btn5.Enabled:=false;//����ť��ʱʧЧ�Է��������
 with Form1 do
 begin
  //row1:=1;row2:=1;row3:=1;row4:=1;
  //crow:=1;  //for i:=1 to MAX do totle[i]:=0;//���ºϼ�
  for i:=1 to cou1 do
  begin
    caption:='��'+inttostr(i)+'�� ������...';
    for m:=1 to MAX do sum[1,m]:=0;
    //row1:=row1+1;//inc(row);

    str:= aq1.fieldbyname('һ����������').asstring;
    //sht[1].cells[row1,1]:=aq1.fieldbyname('һ����������').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);    ////�ȽϿ�������ֵ////
{    for j:=1 to MAX do
       aq1['ģ�����'+inttostr(j)+'�׷�ֵ����']:=sum[1,j];    //���¸��׺ϼ�
}    aq2.sql.clear;
    aq2.sql.add('select * from '+TableName[1]+' where [������������һ������ID]='+inttostr(i)+' order by ��������ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     //row2:=row2+1;    //row2:=crow;
     //sht[2].cells[row2,1]:=aq2.fieldbyname('������������').asstring;
     str:=aq2.fieldbyname('������������').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);  ////�ȽϿ�������ֵ////
{     for k:=1 to MAX do
             aq2['ģ�����'+inttostr(j)+'�׷�ֵ����']:=sum[1,j]; //������С�ڸ��׺ϼ�
}     aq3.sql.clear;
     aq3.sql.add('select * from '+TableName[2]+' where ([������������һ������ID]='+inttostr(i)+
                ') and ([��������������������ID]='+inttostr(j)+') order by ��������ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
     //  row3:=row3+1;    //row3:=crow;
       //sht[3].cells[row3,1]:=aq3.fieldbyname('������������').asstring;
       str:=aq3.fieldbyname('������������').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);  ////�ȽϿ�������ֵ////

       aq4.sql.clear;
       aq4.sql.add('select * from '+TableName[3]+' where ([�ļ���������һ������ID]='+inttostr(i)+
                ') and ([�ļ�����������������ID]='+inttostr(j)+
                ') and ([�ļ�����������������ID]='+inttostr(k)+
                ') order by �ļ�����ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
       //row4:=row4+1;
        //sht[4].cells[row4,1]:=aq4.fieldbyname('�ļ���������').asstring;
        str:=aq4.fieldbyname('�ļ���������').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);  ////�ȽϿ�������ֵ////
        aq4.Edit;
        for m:=1 to MAX do begin
          //aq4.FieldByName('ģ�����'+inttostr(m)+'�׷�ֵ����').AsFloat:=sum[4,m];//�����
          aq4['��'+numtohz(m)+'��ģ�����ֵ����']:=sum[4,m] /100 ;//����
          //sht[4].cells[row4,m+1]:=sum[4,m];
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          end;
        aq4.Post;
        aq4.Next;
       end;//��4���������3���ֵ�ǵ�3���ֵ�����е�4��ĺ�
       aq3.Edit;
       for l:=1 to MAX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          aq3['��'+numtohz(l)+'��ģ�����ֵ����']:=sum[3,l] / 100;
          //sht[3].cells[row3,l+1]:=sum[3,l];
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          end;
       aq3.Post;
       aq3.Next;
     end; //��3�����
     aq2.Edit;
     for l:=1 to MAX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          aq2['��'+numtohz(l)+'��ģ�����ֵ����']:=sum[2,l] / 100;
          //sht[2].cells[row2,l+1]:=sum[2,l];
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          end;
     aq2.Post;
     aq2.Next;
    end;
    aq1.Edit;
    for l:=1 to MAX do begin
          sum[1,l]:=sum[1,l]+sumer[2,l];
          aq1['��'+numtohz(l)+'��ģ�����ֵ����']:=sum[1,l] / 100;
          //sht[1].cells[row1,l+1]:=sum[1,l];
          //totle[l]:=totle[l]+sum[1,l];
          //Sht[1].cells[cou1+2,l+1]:=totle[l];//�ܼƣ������һ�µ�����
          end;
    aq1.Post;
   aq1.Next;
   end;
 aq4.Free;
 aq3.free;
 aq2.Free;
 aq1.Free;
 end;
 caption:='���';
 messagebox(0,'ȫ����������ȷ��д�����ݿ���','�ɹ�д��ȫ������',MB_OK);
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
//���ɳ�100����ʽ
procedure TForm1.btn7Click(Sender: TObject);
var
row1,row2,row3,crow,i,j,k,l,m,counter:Integer;
avg:array[1..4]of Double;
big,litle:Double;
aq1,aq2,aq3,aq4:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4:integer;//4����ļ�¼����
sumer:array of array of Double;
totle:array of double;
field,field1,field2,field3,field4:string;
TableXiaBiao:char; first:boolean; sqlstr:string;
begin
    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+TableName[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    if rg1.itemindex=0 then table:='��Ŀ��'  else table:='ģ�����';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+Table+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    SetLength(sumer,5,MAX+1);SetLength(totle,MAX+1);
    TableXiaBiao:='a';i:=ord(TableXiaBiao)+MAX+2;TableXiaBiao:=chr(i);
    try
      v:=CreateOleObject('Excel.Application');//�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
      v.visible:=false;
                v.WorkBooks.Add;   //�½�Excel   �ļ�
                //v.WorkBooks[1].WorkSheets[1].name:= '���Ա�';//��һҳ����
                //v.WorkBooks[1].WorkSheets[2].name:= '�����԰';
                //v.WorkBooks[1].WorkSheets[3].name:= '������ѽ ';
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
                sheet.cells[2,1]:='�ϼ�';
                v.range['b:'+TableXiaBiao].FormatConditions.Delete;
                v.range['b:'+TableXiaBiao].FormatConditions.Add(1, 3,'0');
                v.range['b:'+TableXiaBiao].FormatConditions[1].Font.ColorIndex := 15;
                if is3ji then begin
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Add(1,1,'0','99.995');
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions[2].Font.ColorIndex := 3;
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Add(1,7,'100.005');end
                else begin
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Add(1,1,'0','29.995');
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions[2].Font.ColorIndex := 3;
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Add(1,7,'30.005');
                end;
                v.range['b2:'+TableXiaBiao+'2'].FormatConditions[3].Font.ColorIndex := 3;
                v.range['b:'+TableXiaBiao].columnwidth:=5;
                sheet.range['b3'].select;
                v.ActiveWindow.FreezePanes:= True;
                sheet.range['a1'].select;
                //Sheet.Cells[1,1]:= '�ÿ� ';//��Ԫ������
                //Sheet.Cells[1,2]:= 'ȷʵ ';
                //Sheet.Cells[2,1]:= '��ϲ�� ';
{                v.activesheet.rows[2].select;
                v.Selection.Interior.ColorIndex:= 32;      
                v.range['b2:'+TableXiaBiao+'2'].interior.colorindex:=39;//��ɫ
 }
    except
        showmessage( '��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
        v.DisplayAlerts:=false;//�Ƿ���ʾ����
        v.Quit;//�˳�Excel
        exit;
    end;
    {Application.Restore;
    Application.BringToFront;}
    field1:=' һ���������� ';
    field2:=' ������������ ';
    field3:=' ������������ ';
    field4:=' �ļ��������� ';
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+' from '+TableName[0]+'  order by һ������ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
 btn7.Enabled:=false;
 avg[1]:=0;avg[2]:=0;avg[3]:=0;avg[4]:=0;counter:=0;
 with Form1 do
 begin
   for i:=1 to MAX do
   Sheet.Cells[1,i+1]:=inttostr(i);
   Sheet.Cells[1,MAX+2]:='ƽ��ֵ';
   Sheet.Cells[1,MAX+3]:='���˸���';
  crow:=2;for i:=1 to MAX do totle[i]:=0;
  for i:=1 to cou1 do
  begin
    caption:='��'+inttostr(i)+'�� ������...';
    for m:=1 to MAX do sum[1,m]:=0;
    crow:=crow+1;//inc(row);
    row1:=crow;
    v.range['a'+inttostr(crow)+':'+TableXiaBiao+inttostr(crow)].interior.colorindex:=24;
    str:= aq1.fieldbyname('һ����������').asstring;
    Sheet.cells[crow,1]:=' '+aq1.fieldbyname('һ����������').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);    ////�ȽϿ�������ֵ////
    for j:=1 to MAX do
      Sheet.cells[crow,j+1]:=sum[1,j]/100;
    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+TableName[1]+' where [������������һ������ID]='+inttostr(i)+' order by ��������ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     crow:=crow+1;    row2:=crow;
     Sheet.cells[crow,1]:='     '+aq2.fieldbyname('������������').asstring;
     str:=aq2.fieldbyname('������������').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);  ////�ȽϿ�������ֵ////
     for k:=1 to MAX do Sheet.cells[crow,k+1]:=sum[2,k]/100;
     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+TableName[2]+' where ([������������һ������ID]='+inttostr(i)+
                ') and ([��������������������ID]='+inttostr(j)+') order by ��������ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
       crow:=crow+1;    row3:=crow;
       Sheet.cells[crow,1]:='         '+aq3.fieldbyname('������������').asstring;
       str:=aq3.fieldbyname('������������').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);   ////�ȽϿ�������ֵ////

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+TableName[3]+' where ([�ļ���������һ������ID]='+inttostr(i)+
                ') and ([�ļ�����������������ID]='+inttostr(j)+
                ') and ([�ļ�����������������ID]='+inttostr(k)+
                ') order by �ļ�����ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
        crow:=crow+1;
        Sheet.cells[crow,1]:='             '+aq4.fieldbyname('�ļ���������').asstring;
        str:=aq4.fieldbyname('�ļ���������').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);  ////�ȽϿ�������ֵ////
        for m:=1 to MAX do begin
          Sheet.cells[crow,m+1]:=sum[4,m]/100;
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          avg[4]:=avg[4]+sum[4,m];
          if sum[4,m]>0 then counter:=counter+1;
          end;
         Sheet.Cells[crow,MAX+2]:=avg[4]/MAX;
         Sheet.Cells[crow,MAX+3]:=counter/MAX;
         avg[4]:=0;counter:=0;
        aq4.Next;
       end;//��4���������3���ֵ�ǵ�3���ֵ�����е�4��ĺ�
       for l:=1 to MAX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          Sheet.cells[row3,l+1]:=sum[3,l]/100;
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          avg[3]:=avg[3]+sum[3,l];
          if sum[3,l]>0 then counter:=counter+1;
          end;
          Sheet.Cells[row3,MAX+2]:=avg[3]/MAX;
          Sheet.Cells[row3,MAX+3]:=counter/MAX;
          avg[3]:=0;counter:=0;
       aq3.Next;
     end; //��3�����
     for l:=1 to MAX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          Sheet.cells[row2,l+1]:=sum[2,l]/100;
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          avg[2]:=avg[2]+sum[2,l];
          if sum[2,l]>0 then counter:=counter+1;
          end;
          Sheet.Cells[row2,MAX+2]:=avg[2]/MAX;
          Sheet.Cells[row2,MAX+3]:=counter/MAX;
          avg[2]:=0;counter:=0;
     aq2.Next;
    end;
    for l:=1 to MAX do begin
          sum[1,l]:=sum[1,l]+sumer[2,l];
          Sheet.cells[row1,l+1]:=sum[1,l]/100;
          totle[l]:=totle[l]+sum[1,l];
          Sheet.cells[2,l+1]:=totle[l];
          avg[1]:=avg[1]+sum[1,l];
          if sum[1,l]>0 then counter:=counter+1;
          end;
          Sheet.Cells[row1,MAX+2]:=avg[1]/MAX;
          Sheet.Cells[row1,MAX+3]:=counter/MAX;
          avg[1]:=0;counter:=0;
   aq1.Next;
   end;
    aq4.Free;
     aq3.free;
     aq2.Free;
     aq1.Free;
 end;
   caption:='���';
   v.visible:=true;//��ʾ��


   if is3ji then begin big:=100.005;litle:=99.995;end
     else begin big:=30.005;litle:=29.995;end;
   first:=true;   //��־
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
          11: chk11.checked:=True;
        end;
        if first then
         begin
          first:=false;
          sqlstr:=' where (th='+inttostr(i);
         end else
          sqlstr:=sqlstr+' or th='+inttostr(i);
      end;
   end;
   if not first then    //�в���Ԥ�ڵ�����
     begin
      sqlstr:=sqlstr+')';
      if not is3ji then
       {  if is2c then
           sqlstr:=sqlstr+' and ((XH>10 And XH<41) OR (XH>45))'
         else           //�����ǰ����C��Ҳ����3�����Ǿ��Ƕ���������Ŀ
           sqlstr:=sqlstr+' and ((XH>10 And XH<36) OR (XH>40))';
        }
         sqlstr:=sqlstr+' and (XH>10)';///20140701
         Application.Restore;
         Application.BringToFront;
      if MessageBox(0,'���ֲ���г�����ݣ�Ҫ�����ó���ɹһɹ��','ע��',MB_ICONINFORMATION+MB_YESNO+MB_DEFBUTTON1)=mryes then
      CheckTestPoint(sqlstr);
     end else
     messagebox(0,'ȫ�����ݼ�������������','������ݵ���',MB_OK);
    btn7.Enabled:=true;


end;
//д���������
procedure TForm1.btn8Click(Sender: TObject);
var
i,j,k,l,m,MuX:Integer;
aq1,aq2,aq3,aq4:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4:integer;//4����ļ�¼����
sumer:array of array of Double;
zt:array[1..10]of string;
field1,field2,field3,field4:string;
begin
    if MessageBox(0,'������д���ݿ�������������������ȷ����','����',MB_YESNO or MB_DEFBUTTON2	or MB_ICONWARNING)=mrno then Exit;
    if MessageBox(0,'ȷ��һ���Լ��϶�Ҫ��д���ݿ���','���潫����ԭ������',MB_YESNO or MB_DEFBUTTON2	or MB_ICONWARNING)=mrno then Exit;
    if MessageBox(0,'�������޷�����������','���棡��',MB_YESNO or MB_DEFBUTTON2	or MB_ICONWARNING)=mrno then Exit;

    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+TableName[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    table:='��Ŀ��';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+Table+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    SetLength(sumer,5,MAX+1);
    {Application.Restore;
    Application.BringToFront;}
//      field1:='*';
//      field2:='*';
//      field3:='*';
//      field4:='*';
    if is2c then begin
      zt[1]:='2008��4�·�ֵ����';zt[2]:='2008��9�·�ֵ����';
      zt[3]:='2009��3�·�ֵ����';zt[4]:='2009��9�·�ֵ����';
      zt[5]:='2010��3�·�ֵ����';zt[6]:='2010��9�·�ֵ����';
      zt[7]:='2011��3�·�ֵ����';zt[8]:='2011��9�·�ֵ����';MuX:=8;
      field1:=' ���,һ����������,[2008��4�·�ֵ����],[2008��9�·�ֵ����],[2009��3�·�ֵ����],[2009��9�·�ֵ����],[2010��3�·�ֵ����],[2010��9�·�ֵ����],[2011��3�·�ֵ����],[2011��9�·�ֵ����] ';
      field2:=' ���,������������,[2008��4�·�ֵ����],[2008��9�·�ֵ����],[2009��3�·�ֵ����],[2009��9�·�ֵ����],[2010��3�·�ֵ����],[2010��9�·�ֵ����],[2011��3�·�ֵ����],[2011��9�·�ֵ����] ';
      field3:=' ���,������������,[2008��4�·�ֵ����],[2008��9�·�ֵ����],[2009��3�·�ֵ����],[2009��9�·�ֵ����],[2010��3�·�ֵ����],[2010��9�·�ֵ����],[2011��3�·�ֵ����],[2011��9�·�ֵ����] ';
      field4:=' ���,�ļ���������,[2008��4�·�ֵ����],[2008��9�·�ֵ����],[2009��3�·�ֵ����],[2009��9�·�ֵ����],[2010��3�·�ֵ����],[2010��9�·�ֵ����],[2011��3�·�ֵ����],[2011��9�·�ֵ����] ';
      end
      else begin
      zt[1]:='2007��4�·�ֵ����';zt[2]:='2007��9�·�ֵ����';
      zt[3]:='2008��4�·�ֵ����';zt[4]:='2008��9�·�ֵ����';
      zt[5]:='2009��3�·�ֵ����';zt[6]:='2009��9�·�ֵ����';
      zt[7]:='2010��3�·�ֵ����';zt[8]:='2010��9�·�ֵ����';
      zt[9]:='2011��3�·�ֵ����';zt[10]:='2011��9�·�ֵ����';MuX:=10;
      field1:=' ���,һ����������,[2007��4�·�ֵ����],[2007��9�·�ֵ����],[2008��4�·�ֵ����],[2008��9�·�ֵ����],[2009��3�·�ֵ����],[2009��9�·�ֵ����],[2010��3�·�ֵ����],[2010��9�·�ֵ����],[2011��3�·�ֵ����],[2011��9�·�ֵ����] ';
      field2:=' ���,������������,[2007��4�·�ֵ����],[2007��9�·�ֵ����],[2008��4�·�ֵ����],[2008��9�·�ֵ����],[2009��3�·�ֵ����],[2009��9�·�ֵ����],[2010��3�·�ֵ����],[2010��9�·�ֵ����],[2011��3�·�ֵ����],[2011��9�·�ֵ����] ';
      field3:=' ���,������������,[2007��4�·�ֵ����],[2007��9�·�ֵ����],[2008��4�·�ֵ����],[2008��9�·�ֵ����],[2009��3�·�ֵ����],[2009��9�·�ֵ����],[2010��3�·�ֵ����],[2010��9�·�ֵ����],[2011��3�·�ֵ����],[2011��9�·�ֵ����] ';
      field4:=' ���,�ļ���������,[2007��4�·�ֵ����],[2007��9�·�ֵ����],[2008��4�·�ֵ����],[2008��9�·�ֵ����],[2009��3�·�ֵ����],[2009��9�·�ֵ����],[2010��3�·�ֵ����],[2010��9�·�ֵ����],[2011��3�·�ֵ����],[2011��9�·�ֵ����] ';
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
    AQ1.SQL.Add('select '+field1+'  from '+TableName[0]+'  order by һ������ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
    btn8.Enabled:=false;//����ť��ʱʧЧ�Է��������
 with Form1 do
 begin
  //row1:=1;row2:=1;row3:=1;row4:=1;
  //crow:=1;  //for i:=1 to MAX do totle[i]:=0;//���ºϼ�
  for i:=1 to cou1 do
  begin
    caption:='��'+inttostr(i)+'�� ������...';
    for m:=1 to MAX do sum[1,m]:=0;
    //row1:=row1+1;//inc(row);

    str:= aq1.fieldbyname('һ����������').asstring;
    //sht[1].cells[row1,1]:=aq1.fieldbyname('һ����������').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);   ////�ȽϿ�������ֵ////
{    for j:=1 to MAX do
       aq1['ģ�����'+inttostr(j)+'�׷�ֵ����']:=sum[1,j];    //���¸��׺ϼ�
}    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+TableName[1]+' where [������������һ������ID]='+inttostr(i)+' order by ��������ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     //row2:=row2+1;    //row2:=crow;
     //sht[2].cells[row2,1]:=aq2.fieldbyname('������������').asstring;
     str:=aq2.fieldbyname('������������').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);  ////�ȽϿ�������ֵ////
{     for k:=1 to MAX do
             aq2['ģ�����'+inttostr(j)+'�׷�ֵ����']:=sum[1,j]; //������С�ڸ��׺ϼ�
}     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+TableName[2]+' where ([������������һ������ID]='+inttostr(i)+
                ') and ([��������������������ID]='+inttostr(j)+') order by ��������ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
     //  row3:=row3+1;    //row3:=crow;
       //sht[3].cells[row3,1]:=aq3.fieldbyname('������������').asstring;
       str:=aq3.fieldbyname('������������').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);  ////�ȽϿ�������ֵ////

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+TableName[3]+' where ([�ļ���������һ������ID]='+inttostr(i)+
                ') and ([�ļ�����������������ID]='+inttostr(j)+
                ') and ([�ļ�����������������ID]='+inttostr(k)+
                ') order by �ļ�����ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
       //row4:=row4+1;
        //sht[4].cells[row4,1]:=aq4.fieldbyname('�ļ���������').asstring;
        str:=aq4.fieldbyname('�ļ���������').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);  ////�ȽϿ�������ֵ////
        aq4.Edit;
        for m:=1 to MuX do begin
          //aq4.FieldByName('ģ�����'+inttostr(m)+'�׷�ֵ����').AsFloat:=sum[4,m];//�����
          aq4[zt[m]]:=sum[4,m] /100 ;//����
          //sht[4].cells[row4,m+1]:=sum[4,m];
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          end;
        aq4.Post;
        aq4.Next;
       end;//��4���������3���ֵ�ǵ�3���ֵ�����е�4��ĺ�
       aq3.Edit;
       for l:=1 to MuX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          aq3[zt[l]]:=sum[3,l] / 100;
          //sht[3].cells[row3,l+1]:=sum[3,l];
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          end;
       aq3.Post;
       aq3.Next;
     end; //��3�����
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
          //Sht[1].cells[cou1+2,l+1]:=totle[l];//�ܼƣ������һ�µ�����
          end;
    aq1.Post;
   aq1.Next;
   end;
 aq4.Free;
 aq3.free;
 aq2.Free;
 aq1.Free;
 end;
 caption:='���';
 messagebox(0,'ȫ����������ȷ��д�����ݿ���','�ɹ�д��ȫ������',MB_OK);
 btn8.Enabled:=true;

end;


//��������һ����
procedure TForm1.btn2Click(Sender: TObject);
var
row1,row2,row3,crow,i,j,k,l,m,counter:Integer;
avg:array[1..4] of Double;
big,little:Double;
aq1,aq2,aq3,aq4:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4:integer;//4����ļ�¼����
PubTable:array[0..3]of string;
sumer:array of array of Double;
totle:array  of double;
field,field1,field2,field3,field4:string;
TableXiaBiao:string;//20140701
sqlstr:string;first:Boolean;
begin
   if is3ji then begin MessageBox(0,'����û�й���������лл!','��ʾ',MB_OK or MB_ICONWARNING);exit;end;
   PubTable[0]:='��������֪ʶһ�������';
   PubTable[1]:='��������֪ʶ���������';
   PubTable[2]:='��������֪ʶ���������';
   PubTable[3]:='��������֪ʶ�ļ������';
    if not(TableExist(DBe.DM.con1,PubTable[1])) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+PubTable[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    if rg1.itemindex=0 then table:='��Ŀ��'  else table:='ģ�����';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+Table+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    SetLength(sumer,5,MAX+1);SetLength(totle,MAX+1);
    TableXiaBiao:='a';
    i:=ord('a')+MAX+2;  //ȷ���еķ�Χ
    if i>122 then    //20140701��ĸz������122
    begin
      Tablexiabiao:='a'+chr(97+i-123);
    end
    else
    TableXiaBiao:=chr(i);

    try
      v:=CreateOleObject('Excel.Application');//�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
      v.Visible:=false;
                v.WorkBooks.Add;   //�½�Excel   �ļ�
                v.WorkBooks[1].WorkSheets[1].name:= Table+'������';//��һҳ����
                //v.WorkBooks[1].WorkSheets[2].name:= '�����԰';
                //v.WorkBooks[1].WorkSheets[3].name:= '������ѽ ';
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
                sheet.cells[2,1]:='�ϼ�';
                v.range['b:'+TableXiaBiao].FormatConditions.Delete;
                v.range['b:'+TableXiaBiao].FormatConditions.Add(1, 3,'0');
                v.range['b:'+TableXiaBiao].FormatConditions[1].Font.ColorIndex := 15;
{                if is3ji then begin
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Add(1,1,'0','99.995');
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions[2].Font.ColorIndex := 3;
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Add(1,7,'100.005');end
                else begin
}                 v.range['b2:'+TableXiaBiao+'2'].FormatConditions.delete;
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Add(1,1,'0','9.995');
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions[1].Font.ColorIndex := 3;
                  v.range['b2:'+TableXiaBiao+'2'].FormatConditions.Add(1,7,'10.005');
{                end;}
                v.range['b2:'+TableXiaBiao+'2'].FormatConditions[2].Font.ColorIndex := 3;
                v.range['b:'+TableXiaBiao].columnwidth:=5;
                sheet.range['b3'].select;
                v.ActiveWindow.FreezePanes:= True;
                sheet.range['a1'].select;
                //Sheet.Cells[1,1]:= '�ÿ� ';//��Ԫ������
                //Sheet.Cells[1,2]:= 'ȷʵ ';
                //Sheet.Cells[2,1]:= '��ϲ�� ';
{                v.activesheet.rows[2].select;
                v.Selection.Interior.ColorIndex:= 32;
                v.range['b2:'+TableXiaBiao+'2'].interior.colorindex:=39;//��ɫ
 }
    except
        showmessage( '��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
        v.DisplayAlerts:=false;//�Ƿ���ʾ����
        v.Quit;//�˳�Excel
        exit;
    end;
    {Application.Restore;
    Application.BringToFront;}
    field1:=' һ���������� ';
    field2:=' ������������ ';
    field3:=' ������������ ';
    field4:=' �ļ��������� ';
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+' from '+PubTable[0]+'  order by һ������ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
 avg[1]:=0;avg[2]:=0;avg[3]:=0;avg[4]:=0;counter:=0;//ƽ��ֵ��ʼ��Ϊ0
 btn2.Enabled:=false;
 with Form1 do
 begin
   for i:=1 to MAX do
   Sheet.Cells[1,i+1]:=IntToStr(i);
   Sheet.cells[1,MAX+2]:='ƽ��ֵ';
   Sheet.cells[1,MAX+3]:='���˸���';

  crow:=2;for i:=1 to MAX do totle[i]:=0;
  for i:=1 to cou1 do
  begin
    caption:='��'+inttostr(i)+'�� ������...';
    for m:=1 to MAX do sum[1,m]:=0;
    crow:=crow+1;//inc(row);
    row1:=crow;
    v.range['a'+inttostr(crow)+':'+TableXiaBiao+inttostr(crow)].interior.colorindex:=24;//��ɫ��Χ��K��
    str:= aq1.fieldbyname('һ����������').asstring;
    Sheet.cells[crow,1]:=' '+aq1.fieldbyname('һ����������').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);   ////�ȽϿ�������ֵ////
    for j:=1 to MAX do
      Sheet.cells[crow,j+1]:=sum[1,j]/100;
    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+PubTable[1]+' where [������������һ������ID]='+inttostr(i)+' order by ��������ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     crow:=crow+1;    row2:=crow;
     Sheet.cells[crow,1]:='     '+aq2.fieldbyname('������������').asstring;
     str:=aq2.fieldbyname('������������').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);  ////�ȽϿ�������ֵ////
     for k:=1 to MAX do Sheet.cells[crow,k+1]:=sum[2,k]/100;
     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+PubTable[2]+' where ([������������һ������ID]='+inttostr(i)+
                ') and ([��������������������ID]='+inttostr(j)+') order by ��������ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
       crow:=crow+1;    row3:=crow;
       Sheet.cells[crow,1]:='         '+aq3.fieldbyname('������������').asstring;
       str:=aq3.fieldbyname('������������').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table);   ////�ȽϿ�������ֵ////

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+PubTable[3]+' where ([�ļ���������һ������ID]='+inttostr(i)+
                ') and ([�ļ�����������������ID]='+inttostr(j)+
                ') and ([�ļ�����������������ID]='+inttostr(k)+
                ') order by �ļ�����ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
        crow:=crow+1;
        Sheet.cells[crow,1]:='             '+aq4.fieldbyname('�ļ���������').asstring;
        str:=aq4.fieldbyname('�ļ���������').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);  ////�ȽϿ�������ֵ////
        for m:=1 to MAX do begin
          Sheet.cells[crow,m+1]:=sum[4,m]/100;
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          avg[4]:=avg[4]+sum[4,m];        //�����ܺϣ�׼������ƽ��ֵ
          if sum[4,m]>0 then counter:=counter+1; //����0�ͼ����������������
        end;
         Sheet.cells[crow,MAX+2]:=avg[4]/MAX;//�ļ������ƽ��ֵ
         Sheet.cells[crow,MAX+3]:=counter/MAX;//�ļ�����ĸ���
         avg[4]:=0;counter:=0;//���³�ʼΪ0
        aq4.Next;
       end;//��4���������3���ֵ�ǵ�3���ֵ�����е�4��ĺ�
       for l:=1 to MAX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          Sheet.cells[row3,l+1]:=sum[3,l]/100;
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          avg[3]:=avg[3]+sum[3,l];        //�����ܺϣ�׼������ƽ��ֵ
          if sum[3,l]>0 then counter:=counter+1; //����0�ͼ����������������
       end;
         Sheet.cells[row3,MAX+2]:=avg[3]/MAX;//���������ƽ��ֵ
         Sheet.cells[row3,MAX+3]:=counter/MAX;//��������ĸ���
         avg[3]:=0;counter:=0;//���³�ʼΪ0
       aq3.Next;
     end; //��3�����
     for l:=1 to MAX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          Sheet.cells[row2,l+1]:=sum[2,l]/100;
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          avg[2]:=avg[2]+sum[2,l];        //�����ܺϣ�׼������ƽ��ֵ
          if sum[2,l]>0 then counter:=counter+1; //����0�ͼ����������������
     end;
         Sheet.cells[row2,MAX+2]:=avg[2]/MAX;//���������ƽ��ֵ
         Sheet.cells[row2,MAX+3]:=counter/MAX;//��������ĸ���
         avg[2]:=0;counter:=0;//���³�ʼΪ0
     aq2.Next;
    end;
    for l:=1 to MAX do begin
          sum[1,l]:=sum[1,l]+sumer[2,l];
          Sheet.cells[row1,l+1]:=sum[1,l]/100;
          totle[l]:=totle[l]+sum[1,l];
          Sheet.cells[2,l+1]:=totle[l];
          avg[1]:=avg[1]+sum[1,l];  //�����ֵ��ӣ�׼������ƽ��
          if sum[1,l]>0 then counter:=counter+1; //����0��ֵ�ͼ�����׼���������
    end;
       Sheet.cells[row1,MAX+2]:=avg[1] / MAX;//һ�������ƽ��ֵ
       Sheet.cells[row1,MAX+3]:=counter/MAX; //һ������Ŀ��˸���
       avg[1]:=0;counter:=0;//���³�ʼ��Ϊ0
   aq1.Next;
   end;
    aq4.Free;
     aq3.free;
     aq2.Free;
     aq1.Free;
 end;
   caption:='���';
   v.Visible:=True;//��ʾ��

     big:=10.005;little:=9.995;
     first:=true;   //��־
     for i:=1 to MAX do
     begin
       if (totle[i]<little) or (totle[i]>big) then
         if first then
         begin
          first:=false;
          sqlstr:=' where (th='+inttostr(i);
         end else
          sqlstr:=sqlstr+' or th='+inttostr(i);
     end;
     if not first then    //�в���30������
     begin
      sqlstr:=sqlstr+')';
         Application.Restore;
         Application.BringToFront;
      if MessageBox(0,'���ֲ���г�����ݣ�Ҫ�����ó���ɹһɹ��','ע��',MB_ICONINFORMATION+MB_YESNO+MB_DEFBUTTON1)=mryes then
      CheckPublicTestPoint(sqlstr);
     end else

 {   if chk1.Checked or chk2.Checked or chk3.Checked or chk4.Checked or
       chk5.Checked or chk6.Checked or chk7.Checked or chk8.Checked or
       chk9.Checked then
       begin
         if MessageBox(0,'���ֲ���г�����ݣ�Ҫ�����ó���ɹһɹ��','ע��',MB_ICONINFORMATION+MB_YESNO+MB_DEFBUTTON1)=mryes then
           begin
             grp1.Visible:=true;
             btn3.Enabled:=True;
             btn4.Enabled:=False;
             btn3Click(Sender);
           end;
       end else begin
 }   messagebox(0,'ȫ�����ݼ�������������','������ݵ���',MB_OK);
 //   end;
    btn2.Enabled:=true;
end;
//������������ĸ���
procedure TForm1.btn9Click(Sender: TObject);
var
row1,row2,row3,row4,i,j,k,l,m,counter:Integer;
avg:array[1..4]of Double;
aq1,aq2,aq3,aq4:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4:integer;//4����ļ�¼����
sumer:array of array of Double;
totle:array of double;
sht:array[1..4]of variant;
field1,field2,field3,field4:string;
PubTable:array[0..3]of string;
TableXiaBiao:string;
sqlstr:string;first:Boolean;
big,little:Double;
begin
   if is3ji then begin MessageBox(0,'����û�й���������лл!','��ʾ',MB_OK or MB_ICONWARNING);exit;end;
   PubTable[0]:='��������֪ʶһ�������';
   PubTable[1]:='��������֪ʶ���������';
   PubTable[2]:='��������֪ʶ���������';
   PubTable[3]:='��������֪ʶ�ļ������';

    if not(TableExist(DBe.DM.con1,PubTable[1])) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+PubTable[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    if rg1.itemindex=0 then table:='��Ŀ��'  else table:='ģ�����';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+Table+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    SetLength(sumer,5,MAX+1);SetLength(totle,MAX+1);
    TableXiaBiao:='a';
    i:=ord('a')+MAX+2;  //ȷ���еķ�Χ
    if i>122 then    //20140701��ĸz������122
    begin
      Tablexiabiao:='a'+chr(97+i-123);
    end
    else
    TableXiaBiao:=chr(i);

    try
      v:=CreateOleObject('Excel.Application');//�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
      v.visible:=false;
                v.WorkBooks.Add;   //�½�Excel   �ļ�
                v.WorkBooks[1].WorkSheets[1].name:= Table+'������';
                v.sheets.add(after:=v.sheets[3]);
                v.workBooks[1].Styles['Normal'].Font.size:=10;
                for i:=1 to 4 do
                begin
                  v.workBooks[1].WorkSheets[i].columns[1].columnwidth:=30;
                  v.workBooks[1].WorkSheets[i].select;
                  sht[i]:=v.WorkBooks[1].WorkSheets[i];
                  sht[i].range['b2'].select;
                  v.ActiveWindow.FreezePanes:= True;
                  v.range['b:'+TableXiaBiao].columnwidth:=5;
                  for j:=1 to MAX do
                  Sht[i].Cells[1,j+1]:=IntToStr(j);
                  sht[i].cells[1,MAX+2]:='ƽ��ֵ';
                  Sht[i].Cells[1,MAX+3]:='���˸���';
                  Sht[i].range['b:'+TableXiaBiao].FormatConditions.Delete;
                  Sht[i].range['b:'+TableXiaBiao].FormatConditions.Add(1, 3,'0');
                  sht[i].range['b:'+TableXiaBiao].FormatConditions[1].Font.ColorIndex := 15;
                end;
                v.workBooks[1].WorkSheets[1].select;
                sht[1].range['a1'].select;
    except
        showmessage( '��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
        v.DisplayAlerts:=false;//�Ƿ���ʾ����
        v.Quit;//�˳�Excel
        exit;
    end;
    {Application.Restore;
    Application.BringToFront;}
    field1:=' һ���������� ';
    field2:=' ������������ ';
    field3:=' ������������ ';
    field4:=' �ļ��������� ';

    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+'  from '+PubTable[0]+'  order by һ������ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
 counter:=0;avg[1]:=0;avg[2]:=0;avg[3]:=0;avg[4]:=0;//���ʼ�����ƽ��ֵ��ʼΪ0
 btn9.Enabled:=false;
 for i:=1 to cou1 do
   begin
     sht[1].cells[i+1,1]:=aq1.fieldbyname('һ����������').asstring;
     aq1.Next;
   end;
  aq1.First;
  v.range['a'+inttostr(cou1+2)].HorizontalAlignment := 3;
  sht[1].cells[cou1+2,1]:='�ϼ�';
 with Form1 do
 begin
  row1:=1;row2:=1;row3:=1;row4:=1;
  //crow:=1;  //for i:=1 to MAX do totle[i]:=0;//���ºϼ�
  for i:=1 to cou1 do
  begin
    caption:='��'+inttostr(i)+'�� ������...';
    for m:=1 to MAX do sum[1,m]:=0;
    row1:=row1+1;//inc(row);
//�¾���':k'��ʾk�У���11�У�
    v.range['a'+inttostr(row1)+':'+TableXiaBiao+inttostr(row1)].interior.colorindex:=24;
    str:= aq1.fieldbyname('һ����������').asstring;
    //sht[1].cells[row1,1]:=aq1.fieldbyname('һ����������').asstring;
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getsum(1,str,Table);  ////�ȽϿ�������ֵ////
    for j:=1 to MAX do    sht[1].cells[row1,j+1]:=sum[1,j]/100;    //���¸��׺ϼ�
    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+PubTable[1]+' where [������������һ������ID]='+inttostr(i)+' order by ��������ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    for m:=1 to MAX do sumer[2,m]:=0;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     for m:=1 to MAX do sum[2,m]:=0;
     row2:=row2+1;    //row2:=crow;
     sht[2].cells[row2,1]:=aq2.fieldbyname('������������').asstring;
     str:=aq2.fieldbyname('������������').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getsum(2,str,Table);  ////�ȽϿ�������ֵ////
     for k:=1 to MAX do Sht[2].cells[row2,k+1]:=sum[2,k]/100; //������С�ڸ��׺ϼ�
     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+PubTable[2]+' where ([������������һ������ID]='+inttostr(i)+
                ') and ([��������������������ID]='+inttostr(j)+') order by ��������ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     for m:=1 to MAX do sumer[3,m]:=0;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       for m:=1 to MAX do sum[3,m]:=0;
       row3:=row3+1;    //row3:=crow;
       sht[3].cells[row3,1]:=aq3.fieldbyname('������������').asstring;
       str:=aq3.fieldbyname('������������').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getsum(3,str,Table); ////�ȽϿ�������ֵ////

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+PubTable[3]+' where ([�ļ���������һ������ID]='+inttostr(i)+
                ') and ([�ļ�����������������ID]='+inttostr(j)+
                ') and ([�ļ�����������������ID]='+inttostr(k)+
                ') order by �ļ�����ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       for m:=1 to MAX do sumer[4,m]:=0;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        for m:=1 to MAX do sum[4,m]:=0;
        row4:=row4+1;
        sht[4].cells[row4,1]:=aq4.fieldbyname('�ļ���������').asstring;
        str:=aq4.fieldbyname('�ļ���������').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getsum(4,str,Table);   ////�ȽϿ�������ֵ////
        for m:=1 to MAX do begin
          sht[4].cells[row4,m+1]:=sum[4,m]/100;
          sumer[4,m]:=sumer[4,m]+sum[4,m];
          avg[4]:=avg[4]+sum[4,m];//ÿ�еĸ���ֵ�ĺ�
          if sum[4,m]>0 then counter:=counter+1;//�������0�ͼ�����׼���������
        end;
           sht[4].cells[row4,MAX+2]:=avg[4]/MAX; //ƽ��ֵ
           sht[4].cells[row4,MAX+3]:=counter/MAX;//����
           avg[4]:=0;counter:=0;//���³�ʼ��Ϊ0��׼����һ������ļ���
        aq4.Next;
       end;//��4���������3���ֵ�ǵ�3���ֵ�����е�4��ĺ�
       for l:=1 to MAX do begin
          sum[3,l]:=sum[3,l]+sumer[4,l];
          sht[3].cells[row3,l+1]:=sum[3,l]/100;
          sumer[3,l]:=sumer[3,l]+sum[3,l];
          avg[3]:=avg[3]+sum[3,l];//ÿ�еĸ���ֵ�ĺ�
          if sum[3,l]>0 then counter:=counter+1;//�������0�ͼ�����׼���������
          end;
         sht[3].cells[row3,MAX+2]:=avg[3]/MAX; //ƽ��ֵ
         sht[3].cells[row3,MAX+3]:=counter/MAX;//����
         avg[3]:=0;counter:=0;//���³�ʼ��Ϊ0��׼����һ������ļ���
       aq3.Next;
     end; //��3�����
     for l:=1 to MAX do begin
          sum[2,l]:=sum[2,l]+sumer[3,l];
          sht[2].cells[row2,l+1]:=sum[2,l]/100;
          sumer[2,l]:=sumer[2,l]+sum[2,l];
          avg[2]:=avg[2]+sum[2,l];//ÿ�еĸ���ֵ�ĺ�
          if sum[2,l]>0 then counter:=counter+1;//�������0�ͼ�����׼���������
        end;
        sht[2].cells[row2,MAX+2]:=avg[2]/MAX; //ƽ��ֵ
        sht[2].cells[row2,MAX+3]:=counter/MAX;//����
        avg[2]:=0;counter:=0;//���³�ʼ��Ϊ0��׼����һ������ļ���
     aq2.Next;
    end;
    for l:=1 to MAX do begin
          sum[1,l]:=sum[1,l]+sumer[2,l];
          sht[1].cells[row1,l+1]:=sum[1,l]/100;
          totle[l]:=totle[l]+sum[1,l];
          Sht[1].cells[cou1+2,l+1]:=totle[l];
          avg[1]:=avg[1]+sum[1,l];//ÿ�еĸ���ֵ�ĺ�
          if sum[1,l]>0 then counter:=counter+1;//�������0�ͼ�����׼���������
        end;
        sht[1].cells[row1,MAX+2]:=avg[1]/MAX; //ƽ��ֵ
        sht[1].cells[row1,MAX+3]:=counter/MAX;//����
        avg[1]:=0;counter:=0;//���³�ʼ��Ϊ0��׼����һ������ļ���
   aq1.Next;
   end;
 aq4.Free;
 aq3.free;
 aq2.Free;
 aq1.Free;
 end;
   caption:='���';
   v.visible:=true;//��ʾ��

      big:=10.005;little:=9.995;//20140701
     first:=true;
     for i:=1 to MAX do
     begin
       if (totle[i]<little) or (totle[i]>big) then
         if first then
         begin
          first:=false;
          sqlstr:=' where (th='+inttostr(i);
         end else
          sqlstr:=sqlstr+' or th='+inttostr(i);
     end;
     if not first then
     begin
      sqlstr:=sqlstr+')';
      CheckPublicTestPoint(sqlstr);
     end else
 messagebox(0,'ȫ�����ݼ�������������','������ݵ���',MB_OK);
 Btn9.Enabled:=true;
end;
//get��Ŀ��������1��
procedure getTI(kaod:string;tabl:string);
var
  i,{th,fz,}cou:Integer;
  ti:tadoquery;
  //kaodian:string;
begin
    ti:=Tadoquery.create(nil);
    TI.connection:=dm.con1;
    ti.SQL.Clear;
    ti.SQL.Add('select TH,XH,TM,DA,JX,����1 from '
                +tabl+' where ����1="'+kaod+'" order by TH,XH ASC');
    ti.open;
    cou:=ti.RecordCount;
with Form1 do
begin
    for i:=1 to cou do
    begin
      edt1.SelAttributes.Color:=clRed;
      if is2c then
      case ti.FieldByName('TH').AsInteger   of
          1:edt1.Lines.Append('��������2008��4�µ�'+ti.fieldbyname('XH').AsString+'���������');
          2:edt1.Lines.Append('��������2008��9�µ�'+ti.fieldbyname('XH').AsString+'���������');
          3:edt1.Lines.Append('��������2009��3�µ�'+ti.fieldbyname('XH').AsString+'���������');
          4:edt1.Lines.Append('��������2009��9�µ�'+ti.fieldbyname('XH').AsString+'���������');
          5:edt1.Lines.Append('��������2010��3�µ�'+ti.fieldbyname('XH').AsString+'���������');
          6:edt1.Lines.Append('��������2010��9�µ�'+ti.fieldbyname('XH').AsString+'���������');
          7:edt1.Lines.Append('��������2011��3�µ�'+ti.fieldbyname('XH').AsString+'���������');
          8:edt1.Lines.Append('��������2011��9�µ�'+ti.fieldbyname('XH').AsString+'���������');
          else edt1.Lines.Append('��������2012���Ժ����������������');
       end else
       case ti.FieldByName('TH').AsInteger   of
          1:edt1.Lines.Append('��������2007��4�µ�'+ti.fieldbyname('XH').AsString+'���������');
          2:edt1.Lines.Append('��������2007��9�µ�'+ti.fieldbyname('XH').AsString+'���������');
          3:edt1.Lines.Append('��������2008��4�µ�'+ti.fieldbyname('XH').AsString+'���������');
          4:edt1.Lines.Append('��������2008��9�µ�'+ti.fieldbyname('XH').AsString+'���������');
          5:edt1.Lines.Append('��������2009��3�µ�'+ti.fieldbyname('XH').AsString+'���������');
          6:edt1.Lines.Append('��������2009��9�µ�'+ti.fieldbyname('XH').AsString+'���������');
          7:edt1.Lines.Append('��������2010��3�µ�'+ti.fieldbyname('XH').AsString+'���������');
          8:edt1.Lines.Append('��������2010��9�µ�'+ti.fieldbyname('XH').AsString+'���������');
          9:edt1.Lines.Append('��������2011��3�µ�'+ti.fieldbyname('XH').AsString+'���������');
          10:edt1.Lines.Append('��������2011��9�µ�'+ti.fieldbyname('XH').AsString+'���������');
          else edt1.Lines.Append('��������2012���Ժ����������������');
       end;

      edt1.Lines.Append(ti.fieldbyname('TM').AsString);
      edt1.SelAttributes.Color:=clGreen;
      edt1.Lines.Append('�𰸣�'+ti.fieldbyname('DA').AsString);
      edt1.Lines.Append(ti.fieldbyname('JX').AsString);
      ti.Next;
   end;
end;
ti.Free;
end;
//�������������������Ŀ
procedure TForm1.btn10Click(Sender: TObject);
var
{row1,row2,row3,crow,m,}i,j,k,l:Integer;
//big,litle:Double;
aq1,aq2,aq3,aq4:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4:integer;//4����ļ�¼����
sumer:array of array of Double;
totle:array of double;
{field,}field1,field2,field3,field4:string;
begin
    if not(TableExist(DBe.DM.con1,TableName[1])) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+TableName[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    if rg1.itemindex=0 then table:='��Ŀ��'  else table:='ģ�����';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+Table+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    SetLength(sumer,5,MAX+1);SetLength(totle,MAX+1);
    {Application.Restore;
    Application.BringToFront;}
    field1:=' һ���������� ';
    field2:=' ������������ ';
    field3:=' ������������ ';
    field4:=' �ļ��������� ';
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+' from '+TableName[0]+'  order by һ������ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
 btn10.Enabled:=false;
 with Form1 do
 begin

  for i:=1 to cou1 do
  begin
    caption:='��'+inttostr(i)+'�� ������...';
    edt1.SelAttributes.Size:=22;
    str:= aq1.fieldbyname('һ����������').asstring;
    edt1.Lines.Append(aq1.fieldbyname('һ����������').asstring);
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getTI(str,Table);
    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+TableName[1]+' where [������������һ������ID]='+inttostr(i)+' order by ��������ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     edt1.SelAttributes.Size:=18;
     edt1.Lines.Append(aq2.fieldbyname('������������').asstring);
     str:=aq2.fieldbyname('������������').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getTI(str,Table);
     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+TableName[2]+' where ([������������һ������ID]='+inttostr(i)+
                ') and ([��������������������ID]='+inttostr(j)+') order by ��������ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       edt1.SelAttributes.Size:=14;
       edt1.Lines.Append(aq3.fieldbyname('������������').asstring);
       str:=aq3.fieldbyname('������������').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getTI(str,Table);

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+TableName[3]+' where ([�ļ���������һ������ID]='+inttostr(i)+
                ') and ([�ļ�����������������ID]='+inttostr(j)+
                ') and ([�ļ�����������������ID]='+inttostr(k)+
                ') order by �ļ�����ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        edt1.Lines.Append(aq4.fieldbyname('�ļ���������').asstring);
        str:=aq4.fieldbyname('�ļ���������').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getTI(str,Table);
        aq4.Next;
       end;//��4���������3���ֵ�ǵ�3���ֵ�����е�4��ĺ�
       aq3.Next;
     end; //��3�����
     aq2.Next;
    end;
   aq1.Next;
   end;
     aq4.Free;
     aq3.free;
     aq2.Free;
     aq1.Free;
     edt1.Lines.SaveToFile(copy(TableName[0],1,Length(TableName[0])-10)+'���������.rtf');
 end;
   caption:='���';

    btn10.Enabled:=true;
end;
//���㰴����ֵ��С����
procedure TForm1.btn11Click(Sender: TObject);
var
  stmp,tablen:string;
  c,i,j,count:Integer;
  kaodian:array[1..6] of string;
  bili:array[1..6]of Double;
  btmp:Double;
  aq0:TADOQuery;
begin
  btn11.Enabled:=false;
  aq0:=TADOQuery.Create(nil);
  aq0.Connection:=DM.con1;
  if rg1.itemindex=0 then tablen:='��Ŀ��' else
     tablen:='ģ�����';
  if not TableExist(dbe.DM.con1,tablen) then
    begin
     ShowMessage('û��'+tablen+'��');btn11.Enabled:=True;Exit;
    end;

  aq0.SQL.Clear;
  aq0.SQL.Add('select * from '+tablen+' order by TH,XH ASC');
  aq0.Open;
  count:=aq0.RecordCount;
  for c:=1 to count do
  begin
    kaodian[1]:=aq0.fieldbyname('����1').AsString;
    kaodian[2]:=aq0.fieldbyname('����2').AsString;
    kaodian[3]:=aq0.fieldbyname('����3').AsString;
    kaodian[4]:=aq0.fieldbyname('����4').AsString;
    kaodian[5]:=aq0.fieldbyname('����5').AsString;
    kaodian[6]:=aq0.fieldbyname('����6').AsString;
    bili[1]:=aq0.fieldbyname('����1����').AsFloat;
    bili[2]:=aq0.fieldbyname('����2����').AsFloat;
    bili[3]:=aq0.fieldbyname('����3����').AsFloat;
    bili[4]:=aq0.fieldbyname('����4����').AsFloat;
    bili[5]:=aq0.fieldbyname('����5����').AsFloat;
    bili[6]:=aq0.fieldbyname('����6����').AsFloat;
    for j:=1 to 6 do
    begin
      for i:=1 to 5 do
      begin
        if bili[i]=0 then Break;
        if bili[i]<bili[i+1] then
        begin
           btmp:=bili[i];bili[i]:=bili[i+1];bili[i+1]:=btmp;
           stmp:=kaodian[i];kaodian[i]:=kaodian[i+1];kaodian[i+1]:=stmp;
        end;
      end;
    end;
    aq0.Edit;
    aq0['����1']:=kaodian[1];aq0['����2']:=kaodian[2];aq0['����3']:=kaodian[3];
    aq0['����4']:=kaodian[4];aq0['����5']:=kaodian[5];aq0['����6']:=kaodian[6];
    aq0['����1����']:=bili[1];aq0['����2����']:=bili[2];
    aq0['����3����']:=bili[3];aq0['����4����']:=bili[4];
    aq0['����5����']:=bili[5];aq0['����6����']:=bili[6];
    aq0.Post;
    aq0.Next;
  end;
  aq0.Free;
  btn11.Caption:='��������������';
  btn11.Enabled:=True;
end;
//��������������ؿ�������
procedure TForm1.btn12Click(Sender: TObject);
var
i,j,k,l:Integer;
//big,litle:Double;
aq1,aq2,aq3,aq4:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4:integer;//4����ļ�¼����
PubTable:array[0..3]of string;
field1,field2,field3,field4:string;
begin
   if is3ji then begin MessageBox(0,'����û�й�������','��ʾ',MB_OK or MB_ICONINFORMATION);exit;end;
   PubTable[0]:='��������֪ʶһ�������';
   PubTable[1]:='��������֪ʶ���������';
   PubTable[2]:='��������֪ʶ���������';
   PubTable[3]:='��������֪ʶ�ļ������';
    if not(TableExist(DBe.DM.con1,PubTable[1])) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+PubTable[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;
    if rg1.itemindex=0 then table:='��Ŀ��'  else table:='ģ�����';
    if not(TableExist(DBe.DM.con1,Table)) then
    begin
       MessageBox(0,PCHar('���ݱ�:'''+Table+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
       halt;
    end;

    {Application.Restore;
    Application.BringToFront;}
    field1:=' һ���������� ';
    field2:=' ������������ ';
    field3:=' ������������ ';
    field4:=' �ļ��������� ';
    aq1:=TADOQuery.Create(nil);
    aq1.Connection:=DBe.DM.con1;
    aq2:=TADOQuery.Create(nil);
    aq2.connection:=DBE.dm.con1;
    aq3:=TADOQuery.Create(nil);
    aq3.connection:=DBE.dm.con1;
    aq4:=TADOQuery.Create(nil);
    aq4.connection:=DBE.dm.con1;
    AQ1.sql.Clear;
    AQ1.SQL.Add('select '+field1+' from '+PubTable[0]+'  order by һ������ID ASC');
    AQ1.open;
    cou1:=AQ1.recordcount ;
    btn12.Enabled:=false;
 with Form1 do
 begin

  for i:=1 to cou1 do
  begin
    caption:='��'+inttostr(i)+'�� ������...';
    edt1.SelAttributes.Size:=22;
    str:= aq1.fieldbyname('һ����������').asstring;
    edt1.Lines.Append(aq1.fieldbyname('һ����������').asstring);
    str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
    getTI(str,Table);
    aq2.sql.clear;
    aq2.sql.add('select '+field2+' from '+PubTable[1]+' where [������������һ������ID]='+inttostr(i)+' order by ��������ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     edt1.SelAttributes.Size:=18;
     edt1.Lines.Append(aq2.fieldbyname('������������').asstring);
     str:=aq2.fieldbyname('������������').asstring;
     str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
     getTI(str,Table);
     aq3.sql.clear;
     aq3.sql.add('select '+field3+' from '+PubTable[2]+' where ([������������һ������ID]='+inttostr(i)+
                ') and ([��������������������ID]='+inttostr(j)+') order by ��������ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       edt1.SelAttributes.Size:=14;
       edt1.Lines.Append(aq3.fieldbyname('������������').asstring);
       str:=aq3.fieldbyname('������������').asstring;
       str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
       getTI(str,Table);

       aq4.sql.clear;
       aq4.sql.add('select '+field4+' from '+PubTable[3]+' where ([�ļ���������һ������ID]='+inttostr(i)+
                ') and ([�ļ�����������������ID]='+inttostr(j)+
                ') and ([�ļ�����������������ID]='+inttostr(k)+
                ') order by �ļ�����ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        edt1.Lines.Append(aq4.fieldbyname('�ļ���������').asstring);
        str:=aq4.fieldbyname('�ļ���������').asstring;
        str:=copy(str,Pos(' ',str)+1,Length(str)-pos(' ',str));
        getTI(str,Table);
        aq4.Next;
       end;//��4���������3���ֵ�ǵ�3���ֵ�����е�4��ĺ�
       aq3.Next;
     end; //��3�����
     aq2.Next;
    end;
   aq1.Next;
   end;
    aq4.Free;
     aq3.free;
     aq2.Free;
     aq1.Free;
     edt1.Lines.SaveToFile(copy(PubTable[0],1,Length(PubTable[0])-10)+'���������.rtf');
 end;
   caption:='���';

    btn10.Enabled:=true;
end;
//ѡ����Ŀ����ģ�����ȷ������������
procedure TForm1.rg1Click(Sender: TObject);
var
  aq:TADOQuery;
begin
       aq:=TADOQuery.Create(nil);
       aq.Connection:=DBe.DM.con1;
       AQ.sql.Clear;
       if rg1.itemindex=0 then 
          if TableExist(DBe.DM.con1,'��Ŀ��') then
            AQ.SQL.Add('select TH  from ��Ŀ��  Group by TH')
          else
          begin showmessage('��Ŀ�����ڣ�');exit;end
       else
          if TableExist(DBe.DM.con1,'ģ�����') then
            AQ.SQL.Add('select TH  from ģ�����  Group by TH')
          else
          begin showmessage('��Ŀ�����ڣ�');exit;end;
       AQ.open;
       MAX:=AQ.recordcount ;
       aq.free;
       Setlength(sum,5,MAX+1);
end;

end.
