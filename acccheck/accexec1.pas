unit accexec1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids,ADODB,StdCtrls,ShellAPI,
  activex, ComCtrls, //��Ϊ�̲߳����activex
  ComObj, ExtCtrls;

type
  TForm1 = class(TForm)
    StringGrid1: TStringGrid;
    btn1: TButton;
    tmr1: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure btn1Click(Sender: TObject);
    procedure tmr1Timer(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  ///////������ʼ��ֻ������ȫ�ֱ���///////////////////////
//  TableName :array[0..3] of string=('����C����һ�������','����C���Զ��������','����C�������������','����C�����ļ������');
  TableName :array[0..3] of string=('����Accessһ�������','����Access���������','����Access���������','����Access�ļ������');
//  TableName :array[0..3] of string=('����VBһ�������','����VB���������','����VB���������','����VB�ļ������');
//  TableName :array[0..3] of string=('����VFһ�������','����VF���������','����VF���������','����VF�ļ������');
//  TableName :array[0..3] of string=('�������缼��һ�������','�������缼�����������','�������缼�����������','�������缼���ļ������');
  str:string;
  uni:array of Integer; //��Ԥ�������������ж�̬����
  arrhi:Integer;        //һ����¼����Ϊ�����������
  thid:THandle;
implementation
uses ExDm;
{$R *.dfm}
procedure getdata;
var
//v,sheet:Variant;
ccol,crow,i,j,k,l:Integer;
val1,val2,val3,val4,val5,val6,val7,val8,val9,avg:Double;
aq1,aq2,aq3,aq4:TADOQuery;//4����Ĳ�ѯ
cou1,cou2,cou3,cou4:integer;//4����ļ�¼����

begin
 CoInitialize(nil);  //��Ϊ  �̲߳����
 aq1:=TADOQuery.Create(nil);
 aq1.Connection:=ExDm.DM.con1;

 aq2:=TADOQuery.Create(nil);
 aq2.connection:=Exdm.dm.con1;
 aq3:=TADOQuery.Create(nil);
 aq3.connection:=Exdm.dm.con1;
 aq4:=TADOQuery.Create(nil);
 aq4.connection:=Exdm.dm.con1;
 AQ1.sql.Clear;
  AQ1.SQL.Add('select *  from '+TableName[0]+'  order by һ������ID ASC');
  AQ1.open;
  cou1:=AQ1.recordcount ;arrhi:=cou1;SetLength(uni,cou1);
 AQ2.sql.Clear;
  AQ2.SQL.Add('select *  from '+TableName[1]+'  order by ������������һ������ID,��������ID ASC');
  AQ2.open;
  cou2:=AQ2.recordcount ;
 AQ3.sql.Clear;
  AQ3.SQL.Add('select *  from '+TableName[2]+'  order by ������������һ������ID,��������������������ID,��������ID ASC');
  AQ3.open;
  cou3:=AQ3.recordcount ;
 AQ4.sql.Clear;
  AQ4.SQL.Add('select *  from '+TableName[3]+'  order by �ļ���������һ������ID,�ļ�����������������ID,�ļ�����������������ID,�ļ�����ID ASC');
  AQ4.open;
  cou4:=AQ4.recordcount ;
 with Form1 do
 begin
 StringGrid1.Cells[1,0]:='2007��4������';
 StringGrid1.Cells[2,0]:='2007��9������';
 StringGrid1.Cells[3,0]:='2008��4������';
 StringGrid1.Cells[4,0]:='2008��9������';
 StringGrid1.Cells[5,0]:='2009��3������';
 StringGrid1.Cells[6,0]:='2009��9������';
 StringGrid1.Cells[7,0]:='2010��3������';
 StringGrid1.Cells[8,0]:='2010��9������';
 StringGrid1.Cells[9,0]:='2011��3������';
 StringGrid1.Cells[10,0]:='����ƽ��';

  StringGrid1.RowCount:=cou1+cou2+cou3+cou4+3;
  StringGrid1.ColCount:=11;
  for ccol:=1 to 10 do begin
    StringGrid1.Cells[ccol,1]:='��ֵ����';
    StringGrid1.ColWidths[ccol]:=100;
    end;
  crow:=1;
  for i:=1 to cou1 do
  begin
   with StringGrid1 do
   begin
    crow:=crow+1;//inc(row);
    uni[i]:=crow;//���Ʊ�ɫ��ÿ������һ����������ɫ��ͬ��
    cells[0,crow]:=' '+aq1.fieldbyname('һ����������').asstring;
    val1:=aq1.fieldbyname('2007���ϰ����ֵ����').asfloat; val1:=val1*100;
    cells[1,crow]:=FloattoStrF(val1,ffNumber,10,2)+'%';//������2λ�Ǿ��ȣ��˴���Ϊ10
    val2:=aq1.fieldbyname('2007���°����ֵ����').asfloat; val2:=val2*100;
    cells[2,crow]:=FloattoStrF(val2,ffNumber,10,2)+'%';
    val3:=aq1.fieldbyname('2008���ϰ����ֵ����').asfloat; val3:=val3*100;
    cells[3,crow]:=FloattoStrF(val3,ffNumber,10,2)+'%';
    val4:=aq1.fieldbyname('2008���°����ֵ����').asfloat; val4:=val4*100;
    cells[4,crow]:=FloattoStrF(val4,ffNumber,10,2)+'%';
    val5:=aq1.fieldbyname('2009���ϰ����ֵ����').asfloat; val5:=val5*100;
    cells[5,crow]:=FloattoStrF(val5,ffNumber,10,2)+'%';
    val6:=aq1.fieldbyname('2009���°����ֵ����').asfloat; val6:=val6*100;
    cells[6,crow]:=FloattoStrF(val6,ffNumber,10,2)+'%';
    val7:=aq1.fieldbyname('2010���ϰ����ֵ����').asfloat; val7:=val7*100;
    cells[7,crow]:=FloattoStrF(val7,ffNumber,10,2)+'%';
    val8:=aq1.fieldbyname('2010���°����ֵ����').asfloat; val8:=val8*100;
    cells[8,crow]:=FloattoStrF(val8,ffNumber,10,2)+'%';
    val9:=aq1.fieldbyname('2011���ϰ����ֵ����').asfloat; val9:=val9*100;
    cells[9,crow]:=FloattoStrF(val9,ffNumber,10,2)+'%';
    avg:=(val1+val2+val3+val4+val5+val6+val7+val8+val9)/9;
    Cells[10,crow]:=FloatToStrF(avg,ffNumber,10,2)+'%';
    aq2.sql.clear;
    aq2.sql.add('select * from '+TableName[1]+' where [������������һ������ID]='+inttostr(i)+' order by ��������ID ASC');
    aq2.open;
    cou2:=aq2.RecordCount;
    if cou2>0 then
    for j:=1 to cou2 do
    begin
     crow:=crow+1;
     cells[0,crow]:='     '+aq2.fieldbyname('������������').asstring;
     val1:=aq2.fieldbyname('2007���ϰ����ֵ����').asfloat; val1:=val1*100;
     cells[1,crow]:=FloattoStrF(val1,ffNumber,10,2)+'%';//������2λ�Ǿ��ȣ��˴���Ϊ10
     val2:=aq2.fieldbyname('2007���°����ֵ����').asfloat; val2:=val2*100;
     cells[2,crow]:=FloattoStrF(val2,ffNumber,10,2)+'%';
     val3:=aq2.fieldbyname('2008���ϰ����ֵ����').asfloat; val3:=val3*100;
     cells[3,crow]:=FloattoStrF(val3,ffNumber,10,2)+'%';
     val4:=aq2.fieldbyname('2008���°����ֵ����').asfloat; val4:=val4*100;
     cells[4,crow]:=FloattoStrF(val4,ffNumber,10,2)+'%';
     val5:=aq2.fieldbyname('2009���ϰ����ֵ����').asfloat; val5:=val5*100;
     cells[5,crow]:=FloattoStrF(val5,ffNumber,10,2)+'%';
     val6:=aq2.fieldbyname('2009���°����ֵ����').asfloat; val6:=val6*100;
     cells[6,crow]:=FloattoStrF(val6,ffNumber,10,2)+'%';
     val7:=aq2.fieldbyname('2010���ϰ����ֵ����').asfloat; val7:=val7*100;
     cells[7,crow]:=FloattoStrF(val7,ffNumber,10,2)+'%';
     val8:=aq2.fieldbyname('2010���°����ֵ����').asfloat; val8:=val8*100;
     cells[8,crow]:=FloattoStrF(val8,ffNumber,10,2)+'%';
     val9:=aq2.fieldbyname('2011���ϰ����ֵ����').asfloat; val9:=val9*100;
     cells[9,crow]:=FloattoStrF(val9,ffNumber,10,2)+'%';
     avg:=(val1+val2+val3+val4+val5+val6+val7+val8+val9)/9;
     Cells[10,crow]:=FloatToStrF(avg,ffNumber,10,2)+'%';
     aq3.sql.clear;
     aq3.sql.add('select * from '+TableName[2]+' where ([������������һ������ID]='+inttostr(i)+
                ') and ([��������������������ID]='+inttostr(j)+') order by ��������ID ASC');
     aq3.open;
     cou3:=aq3.RecordCount;
     if cou3>0 then
     for k:=1 to cou3 do
     begin
       crow:=crow+1;
       cells[0,crow]:='         '+aq3.fieldbyname('������������').asstring;
       val1:=aq3.fieldbyname('2007���ϰ����ֵ����').asfloat; val1:=val1*100;
       cells[1,crow]:=FloattoStrF(val1,ffNumber,10,2)+'%';//������2λ�Ǿ��ȣ��˴���Ϊ10
       val2:=aq3.fieldbyname('2007���°����ֵ����').asfloat; val2:=val2*100;
       cells[2,crow]:=FloattoStrF(val2,ffNumber,10,2)+'%';
       val3:=aq3.fieldbyname('2008���ϰ����ֵ����').asfloat; val3:=val3*100;
       cells[3,crow]:=FloattoStrF(val3,ffNumber,10,2)+'%';
       val4:=aq3.fieldbyname('2008���°����ֵ����').asfloat; val4:=val4*100;
       cells[4,crow]:=FloattoStrF(val4,ffNumber,10,2)+'%';
       val5:=aq3.fieldbyname('2009���ϰ����ֵ����').asfloat; val5:=val5*100;
       cells[5,crow]:=FloattoStrF(val5,ffNumber,10,2)+'%';
       val6:=aq3.fieldbyname('2009���°����ֵ����').asfloat; val6:=val6*100;
       cells[6,crow]:=FloattoStrF(val6,ffNumber,10,2)+'%';
       val7:=aq3.fieldbyname('2010���ϰ����ֵ����').asfloat; val7:=val7*100;
       cells[7,crow]:=FloattoStrF(val7,ffNumber,10,2)+'%';
       val8:=aq3.fieldbyname('2010���°����ֵ����').asfloat; val8:=val8*100;
       cells[8,crow]:=FloattoStrF(val8,ffNumber,10,2)+'%';
       val9:=aq3.fieldbyname('2011���ϰ����ֵ����').asfloat; val9:=val9*100;
       cells[9,crow]:=FloattoStrF(val9,ffNumber,10,2)+'%';
       avg:=(val1+val2+val3+val4+val5+val6+val7+val8+val9)/9;
       Cells[10,crow]:=FloatToStrF(avg,ffNumber,10,2)+'%';
       aq4.sql.clear;
       aq4.sql.add('select * from '+TableName[3]+' where ([�ļ���������һ������ID]='+inttostr(i)+
                ') and ([�ļ�����������������ID]='+inttostr(j)+
                ') and ([�ļ�����������������ID]='+inttostr(k)+
                ') order by �ļ�����ID ASC');
       aq4.open;
       cou4:=aq4.RecordCount;
       if cou4>0 then
       for l:=1 to cou4 do
       begin
        crow:=crow+1;
        cells[0,crow]:='             '+aq4.fieldbyname('�ļ���������').asstring;
        val1:=aq4.fieldbyname('2007���ϰ����ֵ����').asfloat; val1:=val1*100;
        cells[1,crow]:=FloattoStrF(val1,ffNumber,10,2)+'%';//������2λ�Ǿ��ȣ��˴���Ϊ10
        val2:=aq4.fieldbyname('2007���°����ֵ����').asfloat; val2:=val2*100;
        cells[2,crow]:=FloattoStrF(val2,ffNumber,10,2)+'%';
        val3:=aq4.fieldbyname('2008���ϰ����ֵ����').asfloat; val3:=val3*100;
        cells[3,crow]:=FloattoStrF(val3,ffNumber,10,2)+'%';
        val4:=aq4.fieldbyname('2008���°����ֵ����').asfloat; val4:=val4*100;
        cells[4,crow]:=FloattoStrF(val4,ffNumber,10,2)+'%';
        val5:=aq4.fieldbyname('2009���ϰ����ֵ����').asfloat; val5:=val5*100;
        cells[5,crow]:=FloattoStrF(val5,ffNumber,10,2)+'%';
        val6:=aq4.fieldbyname('2009���°����ֵ����').asfloat; val6:=val6*100;
        cells[6,crow]:=FloattoStrF(val6,ffNumber,10,2)+'%';
        val7:=aq4.fieldbyname('2010���ϰ����ֵ����').asfloat; val7:=val7*100;
        cells[7,crow]:=FloattoStrF(val7,ffNumber,10,2)+'%';
        val8:=aq4.fieldbyname('2010���°����ֵ����').asfloat; val8:=val8*100;
        cells[8,crow]:=FloattoStrF(val8,ffNumber,10,2)+'%';
        val9:=aq4.fieldbyname('2011���ϰ����ֵ����').asfloat; val9:=val9*100;
        cells[9,crow]:=FloattoStrF(val9,ffNumber,10,2)+'%';
        avg:=(val1+val2+val3+val4+val5+val6+val7+val8+val9)/9;
        Cells[10,crow]:=FloatToStrF(avg,ffNumber,10,2)+'%';
        aq4.Next;
       end;
       aq3.Next;
     end;
     aq2.Next;
    end;
   aq1.Next;
   end;
  end;
  aq1.free;aq2.free;aq3.free;aq4.free;

 end;
 couninitialize;//��Ϊʹ�����߳��������
end;

procedure TForm1.FormCreate(Sender: TObject);

begin
   getdata;
end;

procedure TForm1.StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  w,h,i:Integer;
begin
if (ACol=0) then
begin
  with TStringGrid(Sender).Canvas do
  begin
    Pen.Color:=clBlack;
    MoveTo(Rect.Right,Rect.top-1);
    LineTo(Rect.Right,Rect.Bottom+1);
    MoveTo(Rect.Left,Rect.Bottom);
    LineTo(Rect.Right,Rect.Bottom);
    Brush.Color := clBtnFace ;
    FillRect(Rect);
    TextOut(Rect.Left+0, Rect.Top +2, TStringGrid(Sender).Cells[ACol, ARow]);//�Լ�����λ�ð�
  end;
end;
if (ARow=0)or(ARow=1) then
begin
  with TStringGrid(Sender).Canvas do
  begin
    Brush.Color := clBtnFace ;
    FillRect(Rect);
    w:=TextWidth(TStringGrid(Sender).cells[acol,arow]);
    h:=TextHeight(TStringGrid(Sender).Cells[acol,arow]);
    TextOut(Rect.Left+(Rect.Right-rect.Left-w) div 2,
            Rect.Top+(Rect.Bottom-rect.Top-h)div 2,
            TStringGrid(Sender).Cells[acol,arow]);
  end;
end;
if(ACol>0)and(arow>1)then
  begin
  for i:=1 to arrhi do
   begin
     if ARow=uni[i] then
     begin
       TStringGrid(Sender).Canvas.Brush.Color:=$c0eac0;
       TStringGrid(Sender).Canvas.FillRect(Rect);
     end;
   end;
   with tstringgrid(Sender).Canvas do
   begin
   //Pen.Color:=clBlack;
   //Brush.Color:=$cccccc;
   FillRect(Rect);
   str:=TStringGrid(Sender).Cells[acol,arow];
   if str='0.00%' then
    Font.Color:=$cccccc;
   {DrawText(TStringGrid(Sender).Canvas.Handle,
                PChar(TStringGrid(Sender).Cells[ACol, ARow]),
                Length(TStringGrid(Sender).Cells[ACol, ARow]),
                Rect,
                DT_CENTER   or   DT_VCENTER   or   DT_SINGLELINE);
   }
    w:=TextWidth(TStringGrid(Sender).cells[acol,arow]);
    h:=TextHeight(TStringGrid(Sender).Cells[acol,arow]);
    TextOut(Rect.Left+(Rect.Right-rect.Left-w) div 2,
            Rect.Top+(Rect.Bottom-rect.Top-h)div 2,
            TStringGrid(Sender).Cells[acol,arow]);
   end;
  end;
  
end;

procedure TForm1.btn1Click(Sender: TObject);
begin
getdata;
end;
//�ȴ������߳̽������һЩ����
procedure waitthread;
begin
  with form1 do begin
   if(WaitForSingleObject(thid,INFINITE)=WAIT_OBJECT_0)then
     begin
      messagebox(0,'ok','ok',MB_OK);
     end;
  end;
end;

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

procedure TForm1.tmr1Timer(Sender: TObject);
 var
   tmp:LongWord;
begin
 if not(TableExist(Exdm.DM.con1,TableName[1]))then
 begin
  tmr1.Enabled:=False;
  MessageBox(0,PCHar('���ݱ�:'''+TableName[1]+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '������',MB_OK or MB_ICONWARNING);
  halt;
 end;
   thid:=CreateThread(nil,0,@getdata,nil,0,TMP);
   if thid=0 then
   begin
     MessageBox(0,'���ִ���,�߳��޷�����','����',MB_OK);
     Exit;
   end;
   CreateThread(nil,0,@waitthread,nil,0,tmp);
   tmr1.Enabled:=False;
end;

end.
