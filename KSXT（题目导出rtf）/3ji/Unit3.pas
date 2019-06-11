unit Unit3;

interface

uses
  Windows, Messages, SysUtils,  Forms,Graphics,
  adodb,  StdCtrls, RxRichEd, Classes, Controls,
  shellapi,
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
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Label3MouseEnter(Sender: TObject);
    procedure Label3MouseLeave(Sender: TObject);
    procedure Label3Click(Sender: TObject);
    procedure chk1Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
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
  TaoHao:array[0..14] of string=('2005��4��','2005��9��','2006��4��','2006��9��',
  '2007��4��','2007��9��','2008��4��','2008��9��','2009��3��','2009��9��',
  '2010��3��','2010��9��','2011��3��','2011��9��','2012��3��');

implementation
uses dm2;
{$R *.dfm}
procedure  TForm1.user_sysmenu(var msg:TWMMENUSELECT);
begin
   if msg.iditem=100 then
      MessageBox(0,'δ���������ݵ�������'+#13+'    �ڲ�ʹ��   '+
      #13+#13+'    2011��7��',
      '���ڱ�����',MB_OK or MB_ICONASTERISK)
      { Ҳ����setwindowpos()��ʵ�ִ�����ǰ�˹���}
   else
      inherited;     { ��ȱʡ����,���������һ����}
end;

procedure shangji3;  //�����ϻ�����
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

      rxrichedit1.Lines.append(AQ.FieldByname('TM2').AsString);
      if (chk1.Checked and (rg3.itemindex=0)) then
        begin
          rxrichedit1.SelAttributes.Color:=clGreen;
          RxRichEdit1.Lines.Append('������');
          RxRichEdit1.Lines.Append(aq.FieldByname('PX2').AsString);
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

      rxrichedit1.Lines.append(AQ.FieldByname('PX2').AsString);

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
procedure bishi3;//�������Դ���
var
  i,num:integer;
  aq:TADOQuery;   thao,xhao:Integer;
begin
  CoInitialize(nil);  //��Ϊ  �̲߳����
  aq:=TADOQuery.Create(nil);
  aq.connection:=dm2.dm.ac1;
  AQ.sql.Clear;
  AQ.SQL.Add('select *  from '+TableName+' order by TH,XH ASC');
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
      rxrichedit1.SelAttributes.Size:=14;
      thao:=aq.FieldByName('TH').asinteger;
      if (thao<13) then
      RXrichedit1.Lines.Append('��'+aq.FieldByname('TH').asstring+'��ģ����')
      else
      RxRichEdit1.Lines.Append(TaoHao[thao-13]+'����');
      rxrichedit1.Paragraph.Alignment:=paleftjustify;
      rxrichedit1.SelAttributes.Color:=clblue;
      //rxrichedit1.SelAttributes.Size:=11;
      xhao:=aq.FieldByName('XH').AsInteger;
      if xhao<61 then
      rxrichedit1.Lines.Append('ѡ���⣨'+aq.fieldByname('XH').AsString+'����')
      else
      RxRichEdit1.Lines.Append('����⣨'+inttostr(xhao-60)+'����');
      rxrichedit1.Lines.append(AQ.FieldByname('TM').asstring);
      if chk1.Checked and(rg3.itemindex=0)then
        begin
          rxrichedit1.SelAttributes.Color:=clgreen;
          RXrichedit1.Lines.Append('�𰸣�'+aq.FieldByname('DA').asstring);
          rxrichedit1.SelAttributes.Color:=clblue;
      //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('������');
          rxrichedit1.Lines.append(AQ.FieldByname('PX').AsString);
        end;
      except
      end;
      aq.Next;
    end;
  if chk1.Checked and(rg3.itemindex=1) then
    begin
      aq.First;
      for i:=1 to num do
        begin
         try
          Canvas.textout(60,290,'���ڴ����'+inttostr(i)+'������...');
          rxrichedit1.Paragraph.Alignment:=pacenter;
          rxrichedit1.SelAttributes.Color:=clblue;
          rxrichedit1.SelAttributes.Size:=14;
      thao:=aq.FieldByName('TH').asinteger;
      if (thao<13) then
      RXrichedit1.Lines.Append('��'+aq.FieldByname('TH').asstring+'��ģ�������')
      else
      RxRichEdit1.Lines.Append(TaoHao[thao-13]+'�������');
          rxrichedit1.Paragraph.Alignment:=paleftjustify;
          rxrichedit1.SelAttributes.Color:=clblue;
          //rxrichedit1.SelAttributes.Size:=11;
      xhao:=aq.FieldByName('XH').AsInteger;
      if xhao<61 then
      rxrichedit1.Lines.Append('ѡ���⣨'+aq.fieldByname('XH').AsString+'�������')
      else
      RxRichEdit1.Lines.Append('����⣨'+inttostr(xhao-60)+'�������');
          rxrichedit1.SelAttributes.Color:=clgreen;
          RXrichedit1.Lines.Append('�𰸣�'+aq.FieldByname('DA').asstring);
          rxrichedit1.SelAttributes.Color:=clblue;
          //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('������');
          rxrichedit1.Lines.append(AQ.FieldByname('PX').AsString);

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
 if (rg1.itemindex=-1)or(rg2.ItemIndex=-1) then
 begin
   messagebox(0,'��ѡ�����ͺͿ�Ŀ','��ʾ',MB_OK or MB_ICONEXCLAMATION);exit;
 end;
 if(rg1.itemindex=0)then
   TableName:='THREE_NET_SJT';
 if(rg1.itemindex=1)then
   TableName:='THREE_NET_XZT';

 if not(TableExist(dm2.DM.AC1,TableName))then
 begin
  MessageBox(0,PCHar('�ף����ݱ�:'''+TableName+'''�ƺ�������!'+#13+#13+'�������ݿ�!'),
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
  if rg1.ItemIndex=1 then
  hthread:=createthread(nil,0,@bishi3,nil,0,tmp) else
  if rg1.itemindex=0 then
  hthread:=createthread(nil,0,@shangji3,nil,0,tmp);
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
   appendmenu(hmenu,MF_STRING,100,'����');
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
   begin TerminateThread(hthread,0);   CanClose:=True;end
                  else CanClose:=False;
end;

end.
