unit Unit1;

interface

uses
  Windows, Messages, SysUtils,  Forms,Graphics,
  adodb,  StdCtrls, RxRichEd, Classes, Controls,
  shellapi, Dialogs, AppEvnts, Variants,About,
  activex, ComCtrls, ExtCtrls;//��Ϊ�̲߳����activex

type
  TForm1 = class(TForm)
    Button1: TButton;
    RxRichEdit1: TRxRichEdit;
    Label1: TLabel;
    RG1: TRadioGroup;
    RG2: TRadioGroup;
    Label2: TLabel;
    Label3: TLabel;
    rg3: TRadioGroup;
    chk1: TCheckBox;
    GRPorderby1: TGroupBox;
    grpWhere1: TGroupBox;
    rbTHXH: TRadioButton;
    rbBianhao: TRadioButton;
    chkNothing: TCheckBox;
    chkBianHao: TCheckBox;
    chkTH: TCheckBox;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Label3MouseEnter(Sender: TObject);
    procedure Label3MouseLeave(Sender: TObject);
    procedure Label3Click(Sender: TObject);
    procedure chk1Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure RG1Click(Sender: TObject);
  private
    { Private declarations }
    procedure  user_sysmenu(var msg:twmmenuselect);
    message wm_syscommand;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  hthread:Thandle;//�߳�
  TableName:string;
  TaoHao:array[0..11] of string=(
  '2007��4��','2007��9��','2008��4��','2008��9��','2009��3��','2009��9��',
  '2010��3��','2010��9��','2011��3��','2011��9��','2012��3��','2012��9��');
  Is2c:Boolean=False;
  Hmax:Integer=15;
  Lmin:Integer;
  SQLwhere,SQLorder:string;{��ѯ��������������ʽ20130710}
implementation
uses dm2;
{$R *.dfm}
procedure  TForm1.user_sysmenu(var msg:TWMMENUSELECT);
begin
   if msg.iditem=100 then
      begin {MessageBox(0,'δ���������ݵ�������'+#13+'    �ڲ�ʹ��   '+
      #13+#13'2011��11��11��',
      '���ڱ�����',MB_OK or MB_ICONASTERISK);}
      aboutbox.show;
      { Ҳ����setwindowpos()��ʵ�ִ�����ǰ�˹���}
      end
   else
      inherited;     { ��ȱʡ����,���������һ����}
end;
procedure shangji1; //һ���ϻ�����
var
  i,num:integer;
  aq:TADOQuery;
begin
  CoInitialize(nil);  //��Ϊ  �̲߳����
  aq:=TADOQuery.Create(nil);
  aq.connection:=dm2.dm.ac1;
  AQ.sql.Clear;
  AQ.SQL.Add('select *  from '+TableName+' order by TH ASC');
  AQ.open;
  num:=AQ.recordcount ;
  //showmessage(inttostr(num));
  with form1 do begin
   label2.caption:='����'+inttostr(num)+'����¼��';
  for i:=1 to num do
    begin
      try
      //label3.caption:='���ڴ����'+inttostr(i)+'��...';
      Canvas.textout(60,290,'���ڴ����'+inttostr(i)+'��...');
      //pbar1.Position:=(pbar1.Max div num)*i;
      {rxrichedit1.Font.Color:=clBlue;
      rxrichedit1.Font.Size:=18;}
      rxrichedit1.Paragraph.Alignment:=pacenter;
      rxrichedit1.SelAttributes.Color:=clblue;
      rxrichedit1.SelAttributes.Size:=16;
      RXrichedit1.Lines.Append('��'+aq.FieldByname('TH').asstring+'����');
      rxrichedit1.Paragraph.Alignment:=paleftjustify;
      rxrichedit1.SelAttributes.Color:=clblue;
      //rxrichedit1.SelAttributes.Size:=11;
      rxrichedit1.Lines.Append('1.����������');
      rxrichedit1.Lines.append(AQ.FieldByname('CZT').asstring);
      if chk1.Checked and (rg3.itemindex=0) then
        begin
          rxrichedit1.SelAttributes.Color:=clGreen;
          RxRichEdit1.Lines.Append('���������������');
          RxRichEdit1.Lines.Append(aq.FieldByname('CZTPX').AsString);
        end;
      rxrichedit1.SelAttributes.Color:=clblue;
      //rxrichedit1.SelAttributes.Size:=11;
      rxrichedit1.Lines.Append('2.����¼��');
      rxrichedit1.Lines.append(AQ.FieldByname('LRWZ').AsString);
      if chk1.Checked and (rg3.itemindex=0) then
        begin
          rxrichedit1.SelAttributes.Color:=clGreen;
          RxRichEdit1.Lines.Append('*����¼����û�н���*');
          //RxRichEdit1.Lines.Append(aq.FieldByname('CZTPX').AsString);
        end;
      rxrichedit1.SelAttributes.Color:=clblue;
      //rxrichedit1.SelAttributes.Size:=11;
      rxrichedit1.Lines.Append('3.Word������');
      rxrichedit1.Lines.append(AQ.FieldByname('WORD').AsString);
      if chk1.Checked and (rg3.ItemIndex=0) then
        begin
          rxrichedit1.SelAttributes.Color:=clgreen;
          //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('=======����======');
          rxrichedit1.Lines.append(AQ.FieldByname('WORDPX').AsString);
        end;  
      rxrichedit1.SelAttributes.Color:=clblue;
      //rxrichedit1.SelAttributes.Size:=11;
      rxrichedit1.Lines.Append('4.Excel������');
      rxrichedit1.Lines.append(AQ.FieldByname('EXCEL').AsString);
      if chk1.Checked and (rg3.ItemIndex=0) then
        begin
          rxrichedit1.SelAttributes.Color:=clgreen;
          //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('=======����======');
          rxrichedit1.Lines.append(AQ.FieldByname('EXCELPX').AsString);
        end;  
      rxrichedit1.SelAttributes.Color:=clblue;
      rxrichedit1.SelAttributes.Size:=11;
      rxrichedit1.Lines.Append('5.PowerPoint������');
      rxrichedit1.Lines.append(AQ.FieldByname('POWERPOINT').AsString);
      if chk1.Checked and (rg3.ItemIndex=0) then
        begin
          rxrichedit1.SelAttributes.Color:=clgreen;
          //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('=======����======');
          rxrichedit1.Lines.append(AQ.FieldByname('POWERPOINTPX').AsString);
        end;  
      rxrichedit1.SelAttributes.Color:=clblue;
      //rxrichedit1.SelAttributes.Size:=11;
      rxrichedit1.Lines.Append('6.������');
      rxrichedit1.Lines.append(AQ.FieldByname('INTERNET').AsString);
      if chk1.Checked and (rg3.ItemIndex=0) then
        begin
          rxrichedit1.SelAttributes.Color:=clgreen;
          //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('=======����======');
          rxrichedit1.Lines.append(AQ.FieldByname('INTERNETPX').AsString);
        end;  

      except
      end;
      aq.Next;
    end;
  if chk1.Checked and(rg3.ItemIndex=1) then
   begin
     aq.first;
     for i:=1 to num do
        begin
          try
          Canvas.textout(60,290,'���ڴ����'+inttostr(i)+'������...');
          rxrichedit1.Paragraph.Alignment:=pacenter;
          rxrichedit1.SelAttributes.Color:=clGreen;
          rxrichedit1.SelAttributes.Size:=16;
          RXrichedit1.Lines.Append('��'+aq.FieldByname('TH').asstring+'�������');
          rxrichedit1.Paragraph.Alignment:=paleftjustify;
          rxrichedit1.SelAttributes.Color:=clGreen;
          rxrichedit1.Lines.Append('1.����������');
          rxrichedit1.Lines.append(AQ.FieldByname('CZTPX').asstring);
          rxrichedit1.SelAttributes.Color:=clGreen;
          RxRichEdit1.Lines.Append('2.*����¼����û�н���*');
          //RxRichEdit1.Lines.Append(aq.FieldByname('CZTPX').AsString);
          rxrichedit1.SelAttributes.Color:=clgreen;
          //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('3.Word���������');
          rxrichedit1.Lines.append(AQ.FieldByname('WORDPX').AsString);
          rxrichedit1.SelAttributes.Color:=clgreen;
          //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('4.Excel���������');
          rxrichedit1.Lines.append(AQ.FieldByname('EXCELPX').AsString);
          rxrichedit1.SelAttributes.Color:=clgreen;
          //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('5.PowerPoint�����======');
          rxrichedit1.Lines.append(AQ.FieldByname('POWERPOINTPX').AsString);
          rxrichedit1.SelAttributes.Color:=clgreen;
          //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('6.���������');
          rxrichedit1.Lines.append(AQ.FieldByname('INTERNETPX').AsString);
          except
          end;
      aq.Next;
      end;
    end;
  FreeandNil(aq);
  rxrichedit1.Lines.SaveToFile(TableName+'.rtf');
  end;
  couninitialize;//��Ϊʹ�����߳��������
end;
procedure shangji2;  //�����ϻ�����
var
  i,num:integer;
  aq:TADOQuery;
begin
  CoInitialize(nil);  //��Ϊ  �̲߳����
  {//richedit1.lines.Text:=dbrichedit1.Text;
  richedit1.Lines.Append(dbrichedit1.Text);
  richedit1.Lines.SaveToFile('output.rtf');
  }
  aq:=TADOQuery.Create(nil);
  aq.connection:=dm2.dm.ac1;
  AQ.sql.Clear;
  AQ.SQL.Add('select *  from '+TableName+'  order by TH ASC');
  AQ.open;
  num:=AQ.recordcount ;
  //showmessage(inttostr(num));
  with form1 do begin
   label2.caption:='����'+inttostr(num)+'����¼��';
  for i:=1 to num do
    begin
      try
      //label3.caption:='���ڴ����'+inttostr(i)+'��...';
      Canvas.textout(60,290,'���ڴ����'+inttostr(i)+'��...');
      //pbar1.Position:=(pbar1.Max div num)*i;
      {rxrichedit1.Font.Color:=clBlue;
      rxrichedit1.Font.Size:=18;}
      rxrichedit1.Paragraph.Alignment:=pacenter;
      rxrichedit1.SelAttributes.Color:=clblue;
      rxrichedit1.SelAttributes.Size:=16;
      RXrichedit1.Lines.Append('��'+aq.FieldByname('TH').asstring+'����');
      rxrichedit1.Paragraph.Alignment:=paleftjustify;
      rxrichedit1.SelAttributes.Color:=clblue;
      //rxrichedit1.SelAttributes.Size:=11;
      rxrichedit1.Lines.Append('��Ŀһ������������');
      rxrichedit1.Lines.append(AQ.FieldByname('TM1').asstring);
      if chk1.Checked and (rg3.itemindex=0) then
        begin
          rxrichedit1.SelAttributes.Color:=clGreen;
          RxRichEdit1.Lines.Append('���������������');
          RxRichEdit1.Lines.Append(aq.FieldByname('PX1').AsString);
        end;
      rxrichedit1.SelAttributes.Color:=clBlue;
      //rxrichedit1.SelAttributes.Size:=11;
      rxrichedit1.Lines.Append('��Ŀ������Ӧ����');
      rxrichedit1.Lines.append(AQ.FieldByname('TM2').AsString);
      if (chk1.Checked and (rg3.itemindex=0)) then
        begin
          rxrichedit1.SelAttributes.Color:=clGreen;
          RxRichEdit1.Lines.Append('��Ӧ���������');
          RxRichEdit1.Lines.Append(aq.FieldByname('PX2').AsString);
        end;
      rxrichedit1.SelAttributes.Color:=clblue;
      //rxrichedit1.SelAttributes.Size:=11;
      rxrichedit1.Lines.Append('��Ŀ�����ۺ�Ӧ����');
      rxrichedit1.Lines.append(AQ.FieldByname('TM3').AsString);
      if chk1.Checked and (rg3.itemindex=0) then
        begin
          rxrichedit1.SelAttributes.Color:=clGreen;
          RxRichEdit1.Lines.Append('�ۺ�Ӧ���������');
          RxRichEdit1.Lines.Append(aq.FieldByname('PX3').AsString);
        end;
      except
      end;
      aq.Next;
    end;
  if (chk1.Checked and(rg3.ItemIndex=1)) then begin
  aq.First;
  for i:=1 to num do
    begin
      try
      Canvas.textout(60,290,'���ڴ����'+inttostr(i)+'������...');
      rxrichedit1.Paragraph.Alignment:=pacenter;
      rxrichedit1.SelAttributes.Color:=clGreen;
      rxrichedit1.SelAttributes.Size:=16;
      RXrichedit1.Lines.Append('��'+aq.FieldByname('TH').asstring+'�������');
      rxrichedit1.Paragraph.Alignment:=paleftjustify;
      rxrichedit1.SelAttributes.Color:=clGreen;
      //rxrichedit1.SelAttributes.Size:=11;
      rxrichedit1.Lines.Append('��Ŀһ������������');
      rxrichedit1.Lines.append(AQ.FieldByname('PX1').asstring);
      rxrichedit1.SelAttributes.Color:=clGreen;
      //rxrichedit1.SelAttributes.Size:=11;
      rxrichedit1.Lines.Append('��Ŀ������Ӧ����');
      rxrichedit1.Lines.append(AQ.FieldByname('PX2').AsString);
      rxrichedit1.SelAttributes.Color:=clGreen;
      //rxrichedit1.SelAttributes.Size:=11;
      rxrichedit1.Lines.Append('��Ŀ�����ۺ�Ӧ����');
      rxrichedit1.Lines.append(AQ.FieldByname('PX3').AsString);
      except

      end;
      aq.Next;
    end;
  end;
  FreeandNil(aq);
  rxrichedit1.Lines.SaveToFile(TableName+'.rtf');
  end;
  couninitialize;//��Ϊʹ�����߳��������
end;
procedure bishi2;//�������Դ���
var
  i,num:integer;
  aq:TADOQuery;  thao,xhao:Integer;
begin
  if    Form1.rbbianhao.Checked then
      sqlorder:=' order by ���'
      else if Form1.rbTHXH.Checked then
              SQLorder:=' order by TH,XH'
            else SQLorder:=' ';
   if Form1.chkBianHao.Checked then
      SQLwhere:=' WHERE (Left([���],3))="N24" ';
   if Form1.chkTH.Checked then SQLwhere:=' ';

  CoInitialize(nil);  //��Ϊ  �̲߳����
  aq:=TADOQuery.Create(nil);
  aq.connection:=dm2.dm.ac1;
  AQ.sql.Clear;
  AQ.SQL.Add('select *  from '+TableName+sqlwhere+sqlorder);
  AQ.open;
  num:=AQ.recordcount ;
  with form1 do begin
   label2.caption:='����'+inttostr(num)+'����¼��';
  for i:=1 to num do
    begin
      try
      //label3.caption:='���ڴ����'+inttostr(i)+'��...';
      Canvas.textout(60,290,'���ڴ����'+inttostr(i)+'��...');
      rxrichedit1.Paragraph.Alignment:=pacenter;
      rxrichedit1.SelAttributes.Color:=clblue;
      rxrichedit1.SelAttributes.Size:=18;
      thao:=aq.FieldByName('TH').asinteger;
//{      if (RG1.itemindex=1) then    begin
//        if aq.FieldByname('XH').AsInteger=1 then
//        RXrichedit1.Lines.Append('��'+aq.FieldByname('TH').asstring+'��ȫ��ģ��');
//      end
//      else
//       begin
//        if aq.FieldByName('XH').AsInteger=1 then
//        if is2c then
//           RxRichEdit1.Lines.Append(TaoHao[thao+1]+'����')
//        else  RxRichEdit1.Lines.Append(TaoHao[thao-1]+'����');
//       end;
//}      rxrichedit1.Paragraph.Alignment:=paleftjustify;
//      rxrichedit1.SelAttributes.Color:=clGreen;
//      //rxrichedit1.SelAttributes.Size:=11;
//      xhao:=aq.FieldByName('XH').AsInteger;
//      if xhao=1 then begin
        rxrichedit1.Paragraph.Alignment:=paLeftJustify;
        rxrichedit1.SelAttributes.Color:=clblue;
        rxrichedit1.SelAttributes.Size:=14;
//        RXrichedit1.Lines.Append('һ��ѡ����');
//        end;
//
//      if (Is2c and (xhao<41)) then
     rxrichedit1.Lines.Append('('+aq.fieldByname('���').AsString+')');
//      if  (Is2c and (xhao=41)) then
//      rxrichedit1.Lines.Append('���������');
//      if  (Is2c and (xhao>40)) then
//      rxrichedit1.Lines.Append('('+inttostr(xhao-40)+')');
//      if (not Is2c) and (xhao<36) then
//      rxrichedit1.Lines.Append('('+aq.fieldByname('XH').AsString+')');
//      if(not Is2c)and(xhao=36)then
//      rxrichedit1.Lines.Append('���������');
//      if(not Is2c)and(xhao>35)then
//      rxrichedit1.Lines.Append('('+inttostr(xhao-35)+')');
      rxrichedit1.Lines.append(AQ.FieldByname('TM').asstring);
//      if chk1.Checked and(rg3.itemindex=0)then
//        begin
          rxrichedit1.SelAttributes.Color:=clgreen;
          RXrichedit1.Lines.Append('�𰸣�'+aq.FieldByname('DA').asstring);
          rxrichedit1.SelAttributes.Color:=clblue;
      //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('������');
          rxrichedit1.Lines.append(AQ.FieldByname('JX').AsString);
//        end;
//      except
//      end;
//      aq.Next;
//    end;
//  if chk1.Checked and(rg3.itemindex=1) then
//    begin
//      aq.First;
//      for i:=1 to num do
//        begin
//         try
//          Canvas.textout(60,290,'���ڴ����'+inttostr(i)+'������...');
//          rxrichedit1.Paragraph.Alignment:=pacenter;
//          rxrichedit1.SelAttributes.Color:=clblue;
//          rxrichedit1.SelAttributes.Size:=18;
//          thao:=aq.FieldByName('TH').asinteger;
//          if (RG1.itemindex=1) then  begin
//            if aq.FieldByName('XH').AsInteger=1 then begin
//           RXrichedit1.Lines.Append('��'+aq.FieldByname('TH').asstring+'��ȫ��ģ�������');
//          rxrichedit1.Paragraph.Alignment:=paleftjustify;
//          rxrichedit1.SelAttributes.Color:=clred;
//          rxrichedit1.SelAttributes.Size:=14;
//
//          RxRichEdit1.Lines.Append('һ��ѡ����');end;
//           end
//          else
//       begin
//        if aq.FieldByName('XH').AsInteger=1 then
//        if is2c then
//           RxRichEdit1.Lines.Append(TaoHao[thao+1]+'������')
//        else  RxRichEdit1.Lines.Append(TaoHao[thao-1]+'������');
//       end;
//          rxrichedit1.Paragraph.Alignment:=paleftjustify;
//          rxrichedit1.SelAttributes.Color:=clblue;
//          //rxrichedit1.SelAttributes.Size:=11;
//      xhao:=aq.FieldByName('XH').AsInteger;
//      if (Is2c and (xhao<41)) then
//      rxrichedit1.Lines.Append('('+aq.fieldByname('XH').AsString+')');
//      if  (Is2c and (xhao=40)) then
//      rxrichedit1.Lines.Append('���������');
//      if  (Is2c and (xhao>40)) then
//      rxrichedit1.Lines.Append('('+inttostr(xhao-40)+')');
////      if (not Is2c) and (xhao=1) then
////      rxrichedit1.Lines.Append('һ��ѡ����');
//      if (not Is2c) and (xhao<36) then
//      rxrichedit1.Lines.Append('('+aq.fieldByname('XH').AsString+')');
//      if(not Is2c)and(xhao=36)then
//      rxrichedit1.Lines.Append('���������');
//      if(not Is2c)and(xhao>35)then
//      rxrichedit1.Lines.Append('('+inttostr(xhao-35)+')');
//          rxrichedit1.SelAttributes.Color:=clgreen;
//          RXrichedit1.Lines.Append('�𰸣�'+aq.FieldByname('DA').asstring);
//          rxrichedit1.SelAttributes.Color:=clblue;
//          //rxrichedit1.SelAttributes.Size:=11;
//          rxrichedit1.Lines.Append('������');
//          rxrichedit1.Lines.append(AQ.FieldByname('JX').AsString);

         except
        end;
      aq.Next;
      end;

  
  FreeandNil(aq);
  rxrichedit1.Lines.SaveToFile(TableName+'.rtf');
  end;
  couninitialize;//��Ϊʹ�����߳��������
end;

//�ȴ������߳̽������һЩ����
procedure waitthread;
begin
  with form1 do begin
   if(WaitForSingleObject(hthread,INFINITE)=WAIT_OBJECT_0)then
     begin
     button1.Enabled:=true;
     rxrichedit1.Visible:=true;
     label1.Visible:=false;
     label3.Visible:=true;
     label3.Caption:=ExtractFilePath(Application.Exename)+TableName+'.rtf';
     label3.Left:=width div 2-label3.Width div 2;
     hthread:=0;
     end;
    if MessageBox(0,PChar('���ݴ�����ϡ��Ƿ��'+#13+'<'+TableName+'.rtf>'+#13+'�����Ϊ.doc�ļ�������С�ļ������'),
                  '�������',MB_ICONINFORMATION+MB_YESNO+MB_DEFBUTTON1)=mrYes then
               shellexecute(0,nil,PChar(label3.Caption),nil,nil,SW_NORMAL);
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
procedure TForm1.Button1Click(Sender: TObject);
var
  Tmp:longword;
begin
 if (rg1.itemindex=-1) then
 begin
   messagebox(0,'��ѡ��Ҫ�����������⻹��ģ���⻹���ϻ���','��ʾ',MB_OK or MB_ICONEXCLAMATION);exit;
 end
 else
 if (rg1.itemindex=2) and (RG2.itemindex=-1) then
 begin
    messagebox(0,'��ѡ��Ҫ�����Ŀ�Ŀ','��ʾ',MB_OK or MB_ICONEXCLAMATION);
    exit;
 end;
 if(rg1.itemindex=2)and(rg2.ItemIndex=0)then
   TableName:='TWO_ACCESS_SJT';
 if(rg1.itemindex=2)and(rg2.ItemIndex=1)then
   TableName:='TWO_VB_SJT';
 if(rg1.itemindex=2)and(rg2.ItemIndex=2)then
   TableName:='TWO_C_SJT';
 if(rg1.itemindex=2)and(rg2.ItemIndex=3)then
   TableName:='TWO_VF_SJT';
 if(rg1.itemindex=2)and(rg2.ItemIndex=4)then
   TableName:='TWO_CAA_SJT';
 if(rg1.itemindex=2)and(rg2.ItemIndex=5)then
   TableName:='ONE_SJT';
{ if(rg1.itemindex=1)and(rg2.ItemIndex=0)then
   begin TableName:='TWO_ACCESS_XZT'; Lmin:=13;end;
 if(rg1.itemindex=1)and(rg2.ItemIndex=1)then
   begin TableName:='TWO_VB_XZT'; Lmin:=15;end;
 if(rg1.itemindex=1)and(rg2.ItemIndex=2)then
   begin TableName:='TWO_C_XZT'; Lmin:=9;Is2c:=True;end;
 if(rg1.itemindex=1)and(rg2.ItemIndex=3)then
   begin TableName:='TWO_VF_XZT'; Lmin:=15;end;
 if(rg1.itemindex=1)and(rg2.ItemIndex=4)then
   TableName:='TWO_CAA_XZT';
 if(rg1.itemindex=1)and((rg2.ItemIndex=5)or(rg2.ItemIndex=6))then
   TableName:='ONE_XZT';
 if(rg1.itemindex=0)and(rg2.ItemIndex=6)then
   TableName:='ONE_B_SJT';
 }
 if TableExist(dm2.DM.AC1,'����C����һ�������') then is2c:=True;
 if RG1.itemindex=0 then Tablename:='��Ŀ��';
 if RG1.ItemIndex=1 then TableName:='ģ�����';
 if not(TableExist(dm2.DM.AC1,TableName))then
 begin
  MessageBox(0,PCHar('���ݱ�:'''+TableName+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
  '�ֳ�����',MB_OK or MB_ICONWARNING);
  exit;
 end;
 if FileExists(TableName+'.rtf') then
  if MessageBox(0,PChar('��ǰĿ¼�Ѿ�����<'+TableName+'.rtf>�ļ�����������'
  +#13+#13+'����<��>��ť���ǣ�����<��>ȡ�����β�����'+#13+#13+'��ѡ��'),
  '�ļ��Ѵ���', MB_ICONINFORMATION + mb_YesNo + MB_DEFBUTTON2) = mrNo then
   Exit;
  button1.Enabled:=false;
  label1.Visible:=true;
  RxrichEdit1.Lines.Clear;
  label3.visible:=false;
  label2.Caption:='';
  if rg1.ItemIndex<2 then
     hthread:=createthread(nil,0,@bishi2,nil,0,tmp) else
{  if (rg2.itemindex>4) and (rg2.ItemIndex<7) then
     hthread:=createthread(nil,0,@shangji1,nil,0,tmp)else
  if rg2.itemindex<5 then                           }
     hthread:=createthread(nil,0,@shangji2,nil,0,tmp);
  if hthread=0 then
  begin
    messagebox(0,'�����߳�ʧ�ܣ���ϵ���򿪷���Ա��','����',MB_OK or MB_ICONEXCLAMATION);
    exit;
  end;
  CreateThread(nil,0,@waitthread,nil,0,tmp);
end;

procedure TForm1.FormCreate(Sender: TObject);
var
 hmenu:integer;
begin
   hmenu:=getsystemmenu(handle,false);
   {��ȡϵͳ�˵����}
   appendmenu(hmenu,MF_SEPARATOR,0,nil);
   appendmenu(hmenu,MF_STRING,100,'����(&A)');
   {�����û��˵�}
   DeleteMenu(hmenu,0,MF_BYPOSITION); //ɾ��������һ���˵���
{   DeleteMenu(hmenu,1,MF_BYPOSITION);
   DeleteMenu(hmenu,2,MF_BYPOSITION);
}   //DeleteMenu(hmenu,0,MF_BYPOSITION);
end;


procedure TForm1.Label3MouseEnter(Sender: TObject);
begin
 label3.font.Color:=clblue;
 label3.Font.Style:=[fsbold,fsunderline];
 label3.Left:=Width div 2-label3.Width div 2;
end;

procedure TForm1.Label3MouseLeave(Sender: TObject);
begin
 label3.Font.Color:=clblue;
 label3.Font.Style:=[];
 label3.Left:=Width div 2-label3.Width div 2;

end;

procedure TForm1.Label3Click(Sender: TObject);
begin
   shellexecute(0,nil,PChar(label3.Caption),nil,nil,SW_NORMAL);
end;

procedure TForm1.chk1Click(Sender: TObject);
begin
  if chk1.Checked then
     rg3.enabled:=True
  else rg3.Enabled:=False;   
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
 if hthread<>0 then
  if MessageBox(0,'�߳���δȫ��������ǿ�ƹرճ���'+#13+
                  '���ܻᵼ�����ݶ�ʧ����ȷ��.','��ʾ',MB_OKCANCEL+MB_ICONWARNING)=IDOK then
     begin    TerminateThread(hthread,0);         CanClose:=True;end
                  else CanClose:=False;
end;

procedure TForm1.RG1Click(Sender: TObject);
begin
  if RG1.ItemIndex=2 then
      RG2.Enabled:=True
      else
      RG2.Enabled:=False;
end;

end.
