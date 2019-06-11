unit Unit3;

interface

uses
  Windows, Messages, SysUtils,  Forms,Graphics,
  adodb,  StdCtrls, RxRichEd, Classes, Controls,
  shellapi,
  activex, ComCtrls, ExtCtrls;//因为线程才添加activex

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
  hthread:Thandle;//线程
  TableName:string;
  TaoHao:array[0..14] of string=('2005年4月','2005年9月','2006年4月','2006年9月',
  '2007年4月','2007年9月','2008年4月','2008年9月','2009年3月','2009年9月',
  '2010年3月','2010年9月','2011年3月','2011年9月','2012年3月');

implementation
uses dm2;
{$R *.dfm}
procedure  TForm1.user_sysmenu(var msg:TWMMENUSELECT);
begin
   if msg.iditem=100 then
      MessageBox(0,'未来教育数据导出程序'+#13+'    内部使用   '+
      #13+#13+'    2011年7月',
      '关于本程序',MB_OK or MB_ICONASTERISK)
      { 也可以setwindowpos()来实现处于最前端功能}
   else
      inherited;     { 作缺省处理,必须调用这一过程}
end;

procedure shangji3;  //三级上机处理
var
  i,num:integer;
  aq:TADOQuery;
begin
  CoInitialize(nil);  //因为  线程才添加
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
   label2.caption:='共有'+inttostr(num)+'条记录。';
  for i:=1 to num do
    begin
      try
      //label3.caption:='正在处理第'+inttostr(i)+'条...';
      Canvas.textout(60,290,'正在处理第'+inttostr(i)+'条...');
      //pbar1.Position:=(pbar1.Max div num)*i;
      {rxrichedit1.Font.Color:=clBlue;
      rxrichedit1.Font.Size:=18;}
      rxrichedit1.Paragraph.Alignment:=pacenter;
      rxrichedit1.SelAttributes.Color:=clblue;
      rxrichedit1.SelAttributes.Size:=16;
      RXrichedit1.Lines.Append('第'+aq.FieldByname('TH').asstring+'套题');
      rxrichedit1.Paragraph.Alignment:=paleftjustify;
      rxrichedit1.SelAttributes.Color:=clblue;

      rxrichedit1.Lines.append(AQ.FieldByname('TM2').AsString);
      if (chk1.Checked and (rg3.itemindex=0)) then
        begin
          rxrichedit1.SelAttributes.Color:=clGreen;
          RxRichEdit1.Lines.Append('解析：');
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
      Canvas.textout(60,290,'正在处理第'+inttostr(i)+'条解析...');
      rxrichedit1.Paragraph.Alignment:=pacenter;
      rxrichedit1.SelAttributes.Color:=clGreen;
      rxrichedit1.SelAttributes.Size:=16;
      RXrichedit1.Lines.Append('第'+aq.FieldByname('TH').asstring+'套题解析');
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
  couninitialize;//因为使用了线程所以添加
end;
procedure bishi3;//三级笔试处理
var
  i,num:integer;
  aq:TADOQuery;   thao,xhao:Integer;
begin
  CoInitialize(nil);  //因为  线程才添加
  aq:=TADOQuery.Create(nil);
  aq.connection:=dm2.dm.ac1;
  AQ.sql.Clear;
  AQ.SQL.Add('select *  from '+TableName+' order by TH,XH ASC');
  AQ.open;
  num:=AQ.recordcount ;
  with form1 do begin
   label2.caption:='共有'+inttostr(num)+'条记录。';
  for i:=1 to num do
    begin
      try
      //label3.caption:='正在处理第'+inttostr(i)+'条...';
      Canvas.textout(60,290,'正在处理第'+inttostr(i)+'条...');
      rxrichedit1.Paragraph.Alignment:=pacenter;
      rxrichedit1.SelAttributes.Color:=clblue;
      rxrichedit1.SelAttributes.Size:=14;
      thao:=aq.FieldByName('TH').asinteger;
      if (thao<13) then
      RXrichedit1.Lines.Append('第'+aq.FieldByname('TH').asstring+'套模拟题')
      else
      RxRichEdit1.Lines.Append(TaoHao[thao-13]+'真题');
      rxrichedit1.Paragraph.Alignment:=paleftjustify;
      rxrichedit1.SelAttributes.Color:=clblue;
      //rxrichedit1.SelAttributes.Size:=11;
      xhao:=aq.FieldByName('XH').AsInteger;
      if xhao<61 then
      rxrichedit1.Lines.Append('选择题（'+aq.fieldByname('XH').AsString+'）题')
      else
      RxRichEdit1.Lines.Append('填空题（'+inttostr(xhao-60)+'）题');
      rxrichedit1.Lines.append(AQ.FieldByname('TM').asstring);
      if chk1.Checked and(rg3.itemindex=0)then
        begin
          rxrichedit1.SelAttributes.Color:=clgreen;
          RXrichedit1.Lines.Append('答案：'+aq.FieldByname('DA').asstring);
          rxrichedit1.SelAttributes.Color:=clblue;
      //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('评析：');
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
          Canvas.textout(60,290,'正在处理第'+inttostr(i)+'条解析...');
          rxrichedit1.Paragraph.Alignment:=pacenter;
          rxrichedit1.SelAttributes.Color:=clblue;
          rxrichedit1.SelAttributes.Size:=14;
      thao:=aq.FieldByName('TH').asinteger;
      if (thao<13) then
      RXrichedit1.Lines.Append('第'+aq.FieldByname('TH').asstring+'套模拟题解析')
      else
      RxRichEdit1.Lines.Append(TaoHao[thao-13]+'真题解析');
          rxrichedit1.Paragraph.Alignment:=paleftjustify;
          rxrichedit1.SelAttributes.Color:=clblue;
          //rxrichedit1.SelAttributes.Size:=11;
      xhao:=aq.FieldByName('XH').AsInteger;
      if xhao<61 then
      rxrichedit1.Lines.Append('选择题（'+aq.fieldByname('XH').AsString+'）题解析')
      else
      RxRichEdit1.Lines.Append('填空题（'+inttostr(xhao-60)+'）题解析');
          rxrichedit1.SelAttributes.Color:=clgreen;
          RXrichedit1.Lines.Append('答案：'+aq.FieldByname('DA').asstring);
          rxrichedit1.SelAttributes.Color:=clblue;
          //rxrichedit1.SelAttributes.Size:=11;
          rxrichedit1.Lines.Append('评析：');
          rxrichedit1.Lines.append(AQ.FieldByname('PX').AsString);

         except
        end;
      aq.Next;
      end;

    end;
  FreeandNil(aq);
  rxrichedit1.Lines.SaveToFile(TableName+'.rtf');
  end;
  couninitialize;//因为使用了线程所以添加
end;

//等待导出线程结束后的一些操作
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
   if MessageBox(0,PChar('数据处理完毕。是否打开'+#13+'<'+TableName+'.rtf>'+#13+'（另存为.doc文件或许会减小文件体积）'),
                  '处理完毕',MB_ICONINFORMATION+MB_YESNO+MB_DEFBUTTON1)=mrYes then
               shellexecute(0,nil,PChar(label3.Caption),nil,nil,SW_NORMAL);

  end;
end;
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
procedure TForm1.Button1Click(Sender: TObject);
var
  Tmp:longword;
begin
 if (rg1.itemindex=-1)or(rg2.ItemIndex=-1) then
 begin
   messagebox(0,'请选择题型和科目','提示',MB_OK or MB_ICONEXCLAMATION);exit;
 end;
 if(rg1.itemindex=0)then
   TableName:='THREE_NET_SJT';
 if(rg1.itemindex=1)then
   TableName:='THREE_NET_XZT';

 if not(TableExist(dm2.DM.AC1,TableName))then
 begin
  MessageBox(0,PCHar('咦？数据表:'''+TableName+'''似乎不存在!'+#13+#13+'请检查数据库!'),
  '又出错了',MB_OK or MB_ICONWARNING);
  exit;
 end;
 if FileExists(TableName+'.rtf') then
  if MessageBox(0,PChar('当前目录已经存在<'+TableName+'.rtf>文件。覆盖它吗？'
  +#13+#13+'单击<是>按钮覆盖，单击<否>取消本次操作。'+#13+#13+'请选择。'),
  '文件已存在', MB_ICONINFORMATION + mb_YesNo + MB_DEFBUTTON2) = mrNo then
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
    messagebox(0,'创建线程失败，联系程序开发人员。','错误',MB_OK or MB_ICONEXCLAMATION);
    exit;
  end;
  CreateThread(nil,0,@waitthread,nil,0,tmp);
end;

procedure TForm1.FormCreate(Sender: TObject);
var
 hmenu:integer;
begin
   hmenu:=getsystemmenu(handle,false);
   {获取系统菜单句柄}
   appendmenu(hmenu,MF_SEPARATOR,0,nil);
   appendmenu(hmenu,MF_STRING,100,'关于');
   {加入用户菜单}
   DeleteMenu(hmenu,0,MF_BYPOSITION); //删掉最上面一个菜单项
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
  if MessageBox(0,'线程尚未全部结束，强制关闭程序'+#13+
                  '可能会导致数据丢失，请确认.','提示',MB_OKCANCEL+MB_ICONWARNING)=IDOK then
   begin TerminateThread(hthread,0);   CanClose:=True;end
                  else CanClose:=False;
end;

end.
