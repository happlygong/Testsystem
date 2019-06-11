Unit PrintDocUnit1;

Interface

Uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComObj, WordXP, ExtCtrls, ShellAPI, Buttons, ImgList,
  IniFiles, ComCtrls, StrUtils;

Type
  TForm1 = Class(TForm)
    ListBox1      : TListBox;
    Panel1        : TPanel;
    lblWholeSale  : TLabel;
    lblCity       : TLabel;
    lblCustom     : TLabel;
    lblPackage    : TLabel;
    lblPackageNum : TLabel;
    edtWholeSale  : TEdit;
    edtCity       : TEdit;
    edtCustom     : TEdit;
    edtPackage    : TEdit;
    edtPackageNum : TEdit;
    lblPrintCopies: TLabel;
    lblStartNum   : TLabel;
    lblEndNum     : TLabel;
    edtPrintCopies: TEdit;
    edtStartNum   : TEdit;
    edtEndNum     : TEdit;
    GroupBox1     : TGroupBox;
    GroupBox2     : TGroupBox;
    btn8040       : TButton;
    btn9060       : TButton;
    btnNetSale    : TButton;
    btnPersonSale : TButton;
    btnPrint      : TButton;
    Label1        : TLabel;
    Label2        : TLabel;
    btnDataMaintain: TButton;
    btnManualInput: TBitBtn;
    ImageList1    : TImageList;
    lbledtPre     : TLabeledEdit;
    lbledtSep     : TLabeledEdit;
    lbledtExt     : TLabeledEdit;
    chkPre        : TCheckBox;
    chkSep        : TCheckBox;
    chkExt        : TCheckBox;
    udPrtCopy     : TUpDown;
    udStart       : TUpDown;
    udEnd         : TUpDown;
    Procedure btnNetSaleClick(Sender: TObject);
    Procedure btnPersonSaleClick(Sender: TObject);
    Procedure btn8040Click(Sender: TObject);
    Procedure btn9060Click(Sender: TObject);
    Procedure edtPrintCopiesKeyPress(Sender: TObject; Var Key: Char);
    Procedure edtEndNumKeyPress(Sender: TObject; Var Key: Char);
    Procedure edtStartNumKeyPress(Sender: TObject; Var Key: Char);
    Procedure edtPrintCopiesChange(Sender: TObject);
    Procedure ListBox1Click(Sender: TObject);
    Procedure btnPrintClick(Sender: TObject);
    Procedure FormCreate(Sender: TObject);
    Procedure btnDataMaintainClick(Sender: TObject);
    Procedure FormDestroy(Sender: TObject);
    Procedure edtEndNumChange(Sender: TObject);
    Procedure btnManualInputClick(Sender: TObject);
    Procedure edtCustomKeyPress(Sender: TObject; Var Key: Char);
    Procedure edtPrintCopiesClick(Sender: TObject);
    Procedure edtStartNumClick(Sender: TObject);
    Procedure edtEndNumClick(Sender: TObject);
    Procedure edtEndNumExit(Sender: TObject);
    Procedure edtStartNumExit(Sender: TObject);
    Procedure edtPrintCopiesExit(Sender: TObject);
    Procedure edtStartNumChange(Sender: TObject);
    Procedure lbledtPreChange(Sender: TObject);
    Procedure lbledtSepChange(Sender: TObject);
    Procedure lbledtExtChange(Sender: TObject);
    Procedure chkPreClick(Sender: TObject);
    Procedure ListBox1DrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
  Private
    { Private declarations }

  Public
    { Public declarations }

  End;

Var
  Form1: TForm1;
  WholeSaleNum, PackageNum: String; //批销单号、包件编号
  CityName, CustomName: String;    //城市、客户
  CustomFileName: String; //客户列表文件名
  TempletFileName: String; //模板文件名
  fNetSale, fPersonSale, f8040, f9060: Boolean;
  w, Doc: Variant; //准备打开Word
  fCustomMody: Boolean; //是否新增了客户，如果是就会写入Excel表
  fFormatMody: Boolean; //是否修改过包件前后缀。如果修改过就保存到ini文件中
  OldCity: String; //为了设置不同城市不同颜色
  Color1, Color2: TColor;

Implementation

{$R *.dfm}

Procedure GetCustomList(); ////获取客户列表
Var
  i: Integer;
  v, Sheet: Variant;
  CellStr: String;
Begin
  If Not FileExists(CustomFileName) Then
  Begin  //MessageBox第1个参数用Form1.handle时，点别处这个窗口会闪
    MessageBox(Form1.Handle, PChar(CustomFilename + '文件未找到或文件不存在，请检查是否改名或者移动位置。'), PChar('错误'), 0);
    Exit; //退出本过程
  End;
  Try
    v := CreateOleObject('Excel.Application');
      //如果这个字符串中有空格就会出'无效的类别字符串'的错
    v.Visible := False;
    //v.WorkBooks.Add;   //新建Excel   文件
    v.workbooks.open(CustomFileName);
    Sheet := v.WorkBooks[1].WorkSheets[1];
  Except
    Showmessage('初始化Excel失败，可能没装Excel，或者其他错误；请重新再试。 ');
    v.DisplayAlerts := false; //是否提示存盘
    v.Quit; //退出Excel
    exit;
  End;
  i := 2;
  CellStr := Sheet.Cells[i, 1].Text + ' - ' + Sheet.Cells[i, 2].Text;
  CityName := Sheet.Cells[i, 1].Text; //为了设置不同城市不同颜色
  OldCity := CityName;               //为了设置不同城市不同颜色
  Form1.ListBox1.Items.Clear; //清空列表
  color1 := clBlue;                   //为了设置不同城市不同颜色
  Color2 := clMaroon;                 //为了设置不同城市不同颜色
  Repeat
    Form1.listbox1.Items.Add(CellStr);
    i := i + 1;
    CellStr := Sheet.cells[i, 1].Text + ' - ' + Sheet.cells[i, 2].Text;
    CityName := Sheet.Cells[i, 1].Text; //为了设置不同城市不同颜色
    //Cellstr:=Concat(Sheet.cells[i,1],  Sheet.cells[i,2]);
  Until CellStr = ' - ';
  v.quit;
  Form1.Label1.Caption := '列表中共有' + inttostr(i - 2) + '条客户记录';
End;

Procedure SetTempletFile(); ////设定模板文件名
Var
  i: Integer;
  tmpfile: String;
Begin
  If (Not f8040) And (Not f9060) Then
    Form1.Label2.caption := '请选择标签';
  If (Not fNetSale) And (Not fPersonSale) Then
    Form1.Label2.caption := '请选择客户';
  If f8040 And fNetSale Then
  Begin
    Form1.Label2.caption := '网销80X40';
    templetfilename := ExtractFilePath(ParamStr(0)) + '数据模板\网销8040T.doc';
  End;
  If f8040 And fPersonSale Then
  Begin
    Form1.Label2.caption := '客户80X40';
    TempletFileName := ExtractFilePath(ParamStr(0)) + '数据模板\客户8040T.doc';
  End;
  If f9060 And fNetSale Then
  Begin
    Form1.Label2.caption := '网销90X60';
    TempletFileName := ExtractFilePath(ParamStr(0)) + '数据模板\网销9060T.doc';
  End;
  If f9060 And fPersonSale Then
  Begin
    Form1.Label2.caption := '客户90X60';
    TempletFileName := ExtractFilePath(ParamStr(0)) + '数据模板\客户9060T.doc';
  End;
  Randomize;
  tmpfile := Format('%.5d', [Random(10000)]);
  tmpfile := ExtractFilePath(ParamStr(0)) + 'tmp\tmp' + tmpfile + '.doc'; //生成模板副本文件名
  ForceDirectories(ExtractFilePath(ParamStr(0)) + 'tmp');
  If CopyFile(PChar(Templetfilename), PChar(tmpfile), true) Then//复制模板
    TempletFileName := tmpfile;

  If fNetSale Then         ////////////////如果是网销
  Begin
    Form1.edtPackage.Text := Trim(PackageNum);
    Form1.lblPackage.Visible := True;     ////网销有包件编号
    Form1.edtPackage.Visible := True;
    If Form1.lblpackage.top > Form1.lblPackageNum.Top Then
    Begin
      i := Form1.lblPackage.Top;
      Form1.lblPackage.Top := Form1.lblPackageNum.Top;
      Form1.lblPackageNum.Top := i;
      Form1.edtPackage.Top := Form1.lblPackage.Top;
      Form1.edtpackageNum.top := Form1.lblPackageNum.Top;
    End;
  End
  Else
  Begin
    Form1.lblPackage.Visible := False; //客户没有包件编号
    Form1.edtPackage.Visible := False;
    If Form1.lblpackage.top < Form1.lblPackageNum.Top Then
    Begin
      i := Form1.lblPackage.Top;
      Form1.lblPackage.Top := Form1.lblPackageNum.Top;
      Form1.lblPackageNum.Top := i;
      Form1.edtPackage.Top := Form1.lblPackage.Top;
      Form1.edtpackageNum.top := Form1.lblPackageNum.Top;
    End;
  End;
End;

Procedure OpenTempletFile();  ////打开模板文件
Var
  temLabWholeSale, WholeSaleNum, PackageNum: String;
Begin
  If Trim(ExtractFileName(TempletFileName)) = '' Then
    exit;
  If Not FileExists(TempletFileName) Then
  Begin   //MessageBox第1个参数如果是0，对话框就不能有最高控制权
    MessageBox(Application.handle, PChar(TempletFilename + '文件未找到或文件不存在，请检查是否改名或者移动位置。'), PChar('错误'), 0);
    fnetsale := False; //20170822在某种情况下模板文件被删除就需要重新点一下重新打开
    fpersonsale := False;
    Form1.ListBox1.Clear;
    Exit; //退出本过程
  End;
  Try
    w := CreateOleObject('Word.Application');
      //如果这个字符串中有空格就会出'无效的类别字符串'的错
    w.Visible := False;
    Doc := w.Documents.Open(FileName := TempletFileName, readonly := false);
//    Doc.saveas(tmpfile);
  Except
    showmessage('初始化Word失败，可能没装Word，或者其他错误；请重新再试。 ');
    f8040 := False;
    f9060 := False;
    Form1.label2.caption := '请选择标签';
    w.DisplayAlerts := false; //是否提示存盘
    w.Quit; //退出Excel
    exit;
  End;
//  doc.printout(); //也可以
{  Doc.PrintOut(     //这个方法需要Uses WordXP
    Background:=True,
    Append:=False,
    Range:=wdPrintCurrentPage,
    Item:=wdPrintDocumentContent,
    Copies:='1',
    Pages:='1',
    PageType:=wdPrintAllPages,
    PrintToFile:=False,
    Collate:=True,
    ManualDuplexPrint:=False);
    }
{    for i:=1 to 13 do
    begin
    //W.ActiveDocument.Range(Start:=0, End:=0);//开始改变的启止位置
    W.ActiveDocument.Range.InsertAfter(Text:=IntToStr(i));//在Word中增加字符'myvc'
    W.ActiveDocument.Range.InsertParagraphAfter;
    //Doc.PrintOut;//不带括号也可以
    end;
    //w.selection.endkey;
    //w.selection.typetext('不保存');
}
  If Doc.tables.count < 1 Then
  Begin
    ShowMessage('模板文件中发现没有表格。模板选择错误？');
    w.Quit;
    Exit;
  End;
    //ShowMessage(Doc.tables.item(1).cell(5,1).range.Text);
    //ShowMessage('ToBe continue');
{  For i := 1 To 15 Do
  Begin
    Doc.Tables.Item(1).Cell(5, 2).range.TEXT := '15' + IntToStr(i);
    Doc.PrintOut;
  End;}
  temLabWholeSale := Doc.tables.item(1).cell(1, 1).range.Text;
  WholeSaleNum := Doc.Tables.Item(1).cell(1, 2).range.Text;
  PackageNum := Doc.Tables.Item(1).Cell(4, 2).range.Text;
  Form1.lblWholeSale.Caption := Trim(temLabWholeSale);
  Form1.edtWholeSale.Text := Trim(WholeSaleNum);
  //Form1.edtCity.Text := trim(CityName);
  //Form1.edtCustom.Text := Trim(CustomName);
  If fNetSale Then         ////////////////如果是网销
  Begin
    Form1.edtPackage.Text := Trim(PackageNum);
  End;
  //w.quit;
End;

Procedure TForm1.btnNetSaleClick(Sender: TObject);
Begin
  If fNetSale Then
    Exit; //如果当前按钮上次已点
  If Not VarIsEmpty(w) Then//如果当前打开过文件
  Try
    W.Quit;
  Except  //如果W已经被意外关闭
  End;
  CustomFileName := ExtractFilePath(ParamStr(0)) + '数据模板\网销.xls';
  GetCustomList();
  fNetSale := True;
  fPersonSale := False;
  SetTempletFile();
  OpenTempletFile();
End;

Procedure TForm1.btnPersonSaleClick(Sender: TObject);
Begin
  If fPersonSale Then
    Exit; //如果当前按钮上次已点
  If Not VarIsEmpty(w) Then//如果当前打开过文件
  Try
    W.Quit;
  Except
    w := varempty; //如果W已经被外部关闭，尝试给它赋空值
  End;
  CustomFileName := ExtractFilePath(ParamStr(0)) + '数据模板\客户.xls';
  GetCustomList();
  fNetSale := False;
  fPersonSale := True;
  SetTempletFile();
  OpenTempletFile();
End;

Procedure TForm1.btn8040Click(Sender: TObject);
Begin
  If f8040 Then
    Exit; //如果当前按钮上次已点
  If Not VarIsEmpty(w) Then//如果当前打开过文件
    W.Quit;

  f8040 := True;
  f9060 := False;
  SetTempletFile();
  OpenTempletFile();
End;

Procedure TForm1.btn9060Click(Sender: TObject);
Begin
  If f9060 Then
    Exit; //如果当前按钮上次已点
  If Not VarIsEmpty(w) Then//如果当前打开过文件
    W.Quit;
  f8040 := False;
  f9060 := True;
  SetTempletFile();
  OpenTempletFile();
End;

Procedure TForm1.edtPrintCopiesKeyPress(Sender: TObject; Var Key: Char);
Begin
  If Not (Key In ['0'..'9', #8]) Then
    Key := #0;  //只允许输入数字
End;

Procedure TForm1.edtEndNumKeyPress(Sender: TObject; Var Key: Char);
Begin
  If Not (Key In ['0'..'9', #8]) Then
    Key := #0;  //只允许输入数字
End;

Procedure TForm1.edtStartNumKeyPress(Sender: TObject; Var Key: Char);
Begin
  If Not (Key In ['0'..'9', #8]) Then
    Key := #0;  //只允许输入数字
End;

Procedure ChangeFormat(); //更改包件前后缀、间隔符
Begin
  With Form1 Do
  Begin
    If chkPre.Checked And chkSep.Checked And chkExt.Checked Then
      edtPackageNum.Text := Trim(lbledtPre.text) + edtprintcopies.Text + Trim(lbledtSep.Text) + edtstartNum.Text + trim(lbledtExt.Text);
    If (chkPre.Checked = False) And chkSep.Checked And chkExt.Checked Then
      edtPackageNum.Text := edtprintcopies.Text + Trim(lbledtSep.Text) + edtstartNum.Text + trim(lbledtExt.Text);
    If chkPre.Checked And (chkSep.Checked = False) And chkExt.Checked Then
      edtPackageNum.Text := Trim(lbledtPre.text) + edtprintcopies.Text + trim(lbledtExt.Text);
    If chkPre.Checked And chkSep.Checked And (chkExt.Checked = False) Then
      edtPackageNum.Text := Trim(lbledtPre.text) + edtprintcopies.Text + Trim(lbledtSep.Text) + edtstartNum.Text;
    If (chkPre.Checked = False) And (chkSep.Checked = False) And chkExt.Checked Then
      edtPackageNum.Text := edtprintcopies.Text + trim(lbledtExt.Text);
    If (chkPre.Checked = False) And chkSep.Checked And (chkExt.Checked = False) Then
      edtPackageNum.Text := edtprintcopies.Text + Trim(lbledtSep.Text) + edtstartNum.Text;
    If (chkPre.Checked = False) And (chkSep.Checked = False) And (chkExt.Checked = False) Then
      edtPackageNum.Text := edtprintcopies.Text;
    If chkPre.Checked And (chkSep.Checked = False) And (chkExt.Checked = False) Then
      edtPackageNum.Text := Trim(lbledtPre.text) + edtprintcopies.Text;
  End;
End;

Procedure TForm1.edtPrintCopiesChange(Sender: TObject);
Begin     //如果总包件数为空，则置为0
  If trim(edtPrintCopies.Text) <> '' Then
  Begin
    edtEndNum.Text := edtPrintCopies.Text;
    ChangeFormat();
  End
  Else
    edtPrintCopies.Text := '0';
End;

Procedure TForm1.ListBox1Click(Sender: TObject);
Var
  strs: TStrings;
  i: Integer;
  constr: String;
Begin
  fcustommody := false; //每次点击客户列表都取消手动输入标记，以免在打印之前使用者改变主意
  constr := ListBox1.Items.Strings[ListBox1.ItemIndex];
  strs := TStringList.Create;
  strs.Delimiter := '-';  //以什么为分割符
//strs.QuoteChar := '|';  //去掉字符串两端的这个符号
//    strs.DelimitedText := constr;//但如果子字符中有空格时不能正解拆分
//    SplitColumns(constr, strs, '-');//该函数没有上述的BUG 需要uses ldStrings;
  ExtractStrings(['-'], [], PChar(constr), strs); //2017.9.1可以不拆分空格
//  For i := 0 To strs.Count - 1 Do
//    ShowMessage(strs[i]);
  If strs.count = 1 Then
  Begin
    edtCity.Text := Trim(strs[0]);
    edtCustom.Text := '';
  End
  Else If strs.count = 2 Then
  Begin
    edtCity.Text := Trim(strs[0]);
    edtCustom.Text := Trim(strs[1]);
  End;
  //edtWholeSale.SetFocus;
End;

//添加客户到Excel文件中，下次可直接读取
Procedure AddCustomList();
Var
  i: Integer;
  v, Sheet: Variant;
  CellStr: String;
Begin
  If Not FileExists(CustomFileName) Then
  Begin
    MessageBox(0, PChar(CustomFilename + '文件未找到或文件不存在，请检查是否改名或者移动位置。'), PChar('错误'), 0);
    Exit; //退出本过程
  End;
  Try
    v := CreateOleObject('Excel.Application');
      //如果这个字符串中有空格就会出'无效的类别字符串'的错
    v.Visible := False;
    //v.WorkBooks.Add;   //新建Excel   文件
    v.workbooks.open(CustomFileName);
    Sheet := v.WorkBooks[1].WorkSheets[1];
  Except
    Showmessage('初始化Excel失败，可能没装Excel，或者其他错误；请重新再试。 ');
    v.DisplayAlerts := false; //是否提示存盘
    v.Quit; //退出Excel
    exit;
  End;
  i := 2;
  //CellStr := Sheet.Cells[i, 1].Text + ' - ' + Sheet.Cells[i, 2].Text;
  Repeat
    i := i + 1;
    CellStr := Sheet.cells[i, 1].Text + ' - ' + Sheet.cells[i, 2].Text;
    //Cellstr:=Concat(Sheet.cells[i,1],  Sheet.cells[i,2]);
  Until CellStr = ' - ';
  Sheet.Cells[i, 1].Value := CityName;  //填入第i行第1列，城市
  Sheet.Cells[i, 2].Value := CustomName; //填入第i行第2列，姓名
  v.ActiveWorkBook.Save;                //保存
  v.quit;                               //关闭
End;


////打印 并保存本次打印到'生成目录'
Procedure TForm1.btnPrintClick(Sender: TObject);
Var
  lngCount, lngStart, lngEnd: Integer;
  i: Integer;
  Row: Integer; //包件数在表格中的行
  AppPath: String;
  strPre, strExt: String;
Begin
  WholeSaleNum := Trim(edtWholeSale.Text);  //批销单号
  PackageNum := Trim(edtPackage.Text);      //包件编号 网销的单子中有
  CityName := Trim(edtCity.Text);           //城市名
  CustomName := Trim(edtCustom.Text);       //客户名
  Try
    lngCount := StrToInt(edtPrintCopies.text); //货物总包件数
    lngStart := StrToInt(edtStartNum.text);    //打印开始编码
    lngEnd := StrToInt(edtEndNum.Text);        //打印结束编码
  Except
    ShowMessage('包件、编号不能为空值！');
    Exit;
  End;
  If VarIsEmpty(W) Then //如果W值为空，说明还没有打开模板文件
  Begin
    MessageBox(Handle, '需要先点一下“网销店铺”或“客户单位”按钮。' + #13#13 + '点完后出现列表才可点其他地方。', '提示', MB_ICONWARNING);
    Exit;
  End;
  If (CityName = '') And (CustomName = '') Then
  Begin
    MessageBox(Handle, '1.请先在左侧列表中选择客户。' + #13#13 + '2.如果在左侧列表中没有你要的客户，请手动输入！', '友情提示', MB_ICONWARNING);
    Exit;
  End;
  //////////////////////////////////////////////////////////////
  If chkPre.Checked And chkSep.Checked And chkExt.Checked Then
  Begin
    strPre := Trim(lbledtPre.text) + edtprintcopies.Text + Trim(lbledtSep.Text);
    strExt := trim(lbledtExt.Text);
  End;
  If (chkPre.Checked = False) And chkSep.Checked And chkExt.Checked Then
  Begin
    strPre := edtprintcopies.Text + Trim(lbledtSep.Text);
    strExt := trim(lbledtExt.Text);
  End;
  If chkPre.Checked And (chkSep.Checked = False) And chkExt.Checked Then
  Begin
    strPre := Trim(lbledtPre.text) + edtprintcopies.Text;
    strExt := trim(lbledtExt.Text);
  End;
  If chkPre.Checked And chkSep.Checked And (chkExt.Checked = False) Then
  Begin
    strPre := Trim(lbledtPre.text) + edtprintcopies.Text + Trim(lbledtSep.Text);
    strExt := '';
  End;
  If (chkPre.Checked = False) And (chkSep.Checked = False) And chkExt.Checked Then
  Begin
    strPre := edtprintcopies.Text;
    strExt := trim(lbledtExt.Text);
  End;
  If (chkPre.Checked = False) And chkSep.Checked And (chkExt.Checked = False) Then
  Begin
    strPre := edtprintcopies.Text + Trim(lbledtSep.Text);
    strExt := '';
  End;
  If (chkPre.Checked = False) And (chkSep.Checked = False) And (chkExt.Checked = False) Then
  Begin
    strPre := edtprintcopies.Text;
    strExt := '';
  End;
  If chkPre.Checked And (chkSep.Checked = False) And (chkExt.Checked = False) Then
  Begin
    strPre := Trim(lbledtPre.text) + edtprintcopies.Text;
    strExt := '';
  End;
  //////////////////////////////////////////////////////////////
  If chkSep.Checked Then
  Begin
    If MessageBox(Handle, PChar('总包件数：' + inttostr(lngCount) + #10#13 + '开始编号：' + InttoStr(lngStart) + #10#13 + '结束编号：' + IntToStr(lngEnd)), '请确认单号、客户、包件信息', MB_OKCancel + MB_DefButton2 + MB_ICONINFORMATION) = IDCancel Then
      Exit;
  End
  Else If MessageBox(Handle, PChar('总包件数：' + inttostr(lngCount) + #13#13 + '只打印总包件数！'), '请确认单号、客户、包件信息', MB_OKCancel + MB_DefButton2 + MB_ICONINFORMATION) = IDCancel Then
    Exit;

  Try
    Doc.tables.Item(1).Cell(1, 2).range.Text := WholeSaleNum; //批销单号放入表格第1行第2列
    Doc.Tables.Item(1).Cell(2, 2).range.Text := CityName; //城市放入表格第2行第2列
    Doc.Tables.Item(1).Cell(3, 2).range.Text := CustomName; //客户名放入表格第3行第2列
    Row := 4; //客户单在第4行是包件数
    If fNetSale Then
    Begin
      Doc.Tables.Item(1).Cell(4, 2).range.Text := PackageNum; //网销包件编号放入表格第4行第2列
      Row := 5; //网销单在第5行是包件数
    End;
    For i := lngStart To lngEnd Do
    Begin
      Doc.Tables.Item(1).Cell(Row, 2).range.Text := strPre + inttostr(i) + strExt;
      Doc.Printout;
    End;
    AppPath := ExtractFilePath(Application.ExeName);
    ForceDirectories(AppPath + '生成目录');
    Doc.saveas(AppPath + '生成目录/' + CityName + '-' + CustomName);
  Except
    On E: Exception Do
      Showmessage(E.ClassName + ':' + #13#10 + E.message + #13#10 + '如果提示“拒绝呼叫”请检查模板打印范围是否超出页边距');
    //W.Quit;
  End;
  //w.quit;
  If fCustomMody Then  //如果在客户输入框中输入过就保存
    AddCustomList();
  fCustomMody := False; //输入标记取消
End;

Procedure SetiniConfig();
Var
  filename: String;
  Myinifile: TIniFile;
Begin
  With Form1 Do
  Begin
    filename := ExtractFilePath(ParamStr(0)) + 'PrintDoc.ini';
    Myinifile := TIniFile.Create(filename);
    Myinifile.WriteString('CONFIG', 'Pre', lbledtPre.Text);
    Myinifile.WriteString('CONFIG', 'Sep', lbledtSep.Text);
    Myinifile.WriteString('CONFIG', 'Ext', lbledtExt.Text);
  End;
  Myinifile.Free;
End;

Procedure Getiniconfig();
Var
  filename: String;
  Myinifile: TIniFile;
Begin
  filename := ExtractFilePath(ParamStr(0)) + 'PrintDoc.ini';
  If Not FileExists(filename) Then
    exit;
  With Form1 Do
  Begin
    Myinifile := TIniFile.Create(filename);
    lbledtPre.Text := Myinifile.ReadString('CONFIG', 'Pre', '');
    lbledtSep.Text := Myinifile.readString('CONFIG', 'Sep', '');
    lbledtExt.Text := Myinifile.ReadString('CONFIG', 'Ext', '');
  End;
  Myinifile.Free;
End;

Procedure TForm1.FormCreate(Sender: TObject);
Var
  f1, f2: String;
Begin
  f8040 := True; //默认为80X40的标签尺寸
  If FileExists(ExtractFilePath(ParamStr(0)) + 'PrintDoc.ini') Then
    GetiniConfig();
  fformatmody := False;
  f1 := 'del ' + ExtractFilePath(ParamStr(0)) + 'tmp\*.doc';
  f2 := 'del ' + ExtractFilePath(ParamStr(0)) + 'tmp\*.docx';
  //WinExec(PChar('cmd dir '+f),SW_SHOW);
  ShellExecute(0, nil, PChar('cmd.exe '), PChar('/c ' + f1), nil, sw_normal);
  ShellExecute(0, nil, PChar('cmd.exe '), PChar('/c ' + f2), nil, sw_normal);
End;

Procedure TForm1.btnDataMaintainClick(Sender: TObject);
Begin
//  if fcustommody then Label2.caption:='变了';
//  if VarIsEmpty(W) then
//  Label2.Caption:='W is Null'
//  else
//  Label2.Caption:='not Null';
  If CustomFileName = '' Then
  Begin
    customfilename := ExtractFilePath(ParamStr(0)) + '数据模板';
    ShellExecute(0, nil, PChar('explorer.exe'), PChar('/e, ' + CustomFileName), nil, SW_NORMAL);
  End
  Else
    ShellExecute(0, nil, PChar('explorer.exe'), PChar('/e, ' + '/select,' + CustomFileName), nil, SW_NORMAL);

End;

Procedure TForm1.FormDestroy(Sender: TObject);
Begin
  If Not VarIsEmpty(w) Then//如果当前打开过文件
  Try
    W.Quit;
  Except //如果W已经被意外关闭
  End;
  If fFormatMody Then
    SetiniConfig();
End;

Procedure TForm1.edtEndNumChange(Sender: TObject);
Begin
  If edtEndNum.text = '' Then
    edtEndNum.text := '0';
  If StrToInt(edtEndNum.text) > StrToInt(edtPrintCopies.text) Then
    edtEndNum.Text := edtPrintCopies.Text;
End;

Procedure TForm1.btnManualInputClick(Sender: TObject);
Begin
  btnManualInput.Glyph := nil;
  If edtCustom.Enabled Then
  Begin
    edtCustom.Enabled := False;
    edtCity.Enabled := False;
    edtPackageNum.Enabled := False;
    ImageList1.GetBitmap(1, BtnManualinput.Glyph);
  End
  Else
  Begin
    edtCustom.Enabled := True;
    edtCity.Enabled := True;
    edtPackageNum.Enabled := True;
    ImageList1.GetBitmap(0, BtnManualinput.Glyph);
  End;
End;

Procedure TForm1.edtCustomKeyPress(Sender: TObject; Var Key: Char);
Begin
  fcustommody := True;
End;

//鼠标点到Edit控件上时，选中Edit控件中的内容
Procedure TForm1.edtPrintCopiesClick(Sender: TObject);
Begin
  edtPrintCopies.SelectAll;
  edtPrintCopies.onclick := nil;
End;

Procedure TForm1.edtStartNumClick(Sender: TObject);
Begin
  edtStartNum.SelectAll;
  edtStartNum.onclick := nil;
End;

Procedure TForm1.edtEndNumClick(Sender: TObject);
Begin
  edtEndNum.SelectAll;
  edtEndNum.onclick := nil;
End;
//为了让它下次进来还是有选中全部内容的效果，需要这句。啊好复杂

Procedure TForm1.edtEndNumExit(Sender: TObject);
Begin
  edtEndNum.OnClick := edtEndNumClick;
End;

Procedure TForm1.edtStartNumExit(Sender: TObject);
Begin
  edtStartNum.OnClick := edtStartNumClick;
End;

Procedure TForm1.edtPrintCopiesExit(Sender: TObject);
Begin
  edtPrintCopies.OnClick := edtPrintCopiesClick;
End;

Procedure TForm1.edtStartNumChange(Sender: TObject);
Begin
  If trim(edtStartNum.Text) <> '' Then
  Begin
    ChangeFormat();
  End
  Else
    edtStartNum.Text := '0';

End;

Procedure TForm1.lbledtPreChange(Sender: TObject);
Begin
  fformatmody := True;
  chkPre.Checked := True;
  ChangeFormat();
End;

Procedure TForm1.lbledtSepChange(Sender: TObject);
Begin
  fformatmody := True;
  chkSep.Checked := true;
  ChangeFormat();
End;

Procedure TForm1.lbledtExtChange(Sender: TObject);
Begin
  fformatmody := True;
  chkExt.Checked := True;
  ChangeFormat();
End;

Procedure TForm1.chkPreClick(Sender: TObject);
Begin
  ChangeFormat();
End;

Procedure TForm1.ListBox1DrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
Var
  position: Integer;
  str: String;
Begin
  str := ListBox1.Items.Strings[Index];
  str := LeftStr(str, 2);
  If Oldcity <> str Then      //如果前一行和当前行不同，则换颜色
  Begin
    ListBox1.Canvas.FillRect(Rect);
    ListBox1.Canvas.Font.Color := Color2;
    Color2 := Color1;
    Color1 := ListBox1.Canvas.Font.Color;
  End
  Else
  Begin                      //否则，还用原来颜色
    ListBox1.Canvas.FillRect(Rect);
    ListBox1.Canvas.Font.Color := Color1;
  End;

{  position := AnsiPos('东', str);
  if position > 0 then
  begin
    ListBox1.Canvas.FillRect(Rect);
    ListBox1.Canvas.Font.Color := clGreen;
  end;
 }
  ListBox1.Canvas.TextOut(Rect.Left + 1, Rect.Top + 0, ListBox1.Items.Strings[Index]);
  OldCity := str;
  If odSelected In State Then        //当选定时的颜色
  Begin
    listbox1.Canvas.Brush.Color := clhighlight;
    ListBox1.Canvas.TextRect(Rect, Rect.Left, Rect.Top, ListBox1.Items[Index]);
  End;
End;

End.

