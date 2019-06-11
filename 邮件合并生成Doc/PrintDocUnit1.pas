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
  WholeSaleNum, PackageNum: String; //�������š��������
  CityName, CustomName: String;    //���С��ͻ�
  CustomFileName: String; //�ͻ��б��ļ���
  TempletFileName: String; //ģ���ļ���
  fNetSale, fPersonSale, f8040, f9060: Boolean;
  w, Doc: Variant; //׼����Word
  fCustomMody: Boolean; //�Ƿ������˿ͻ�������Ǿͻ�д��Excel��
  fFormatMody: Boolean; //�Ƿ��޸Ĺ�����ǰ��׺������޸Ĺ��ͱ��浽ini�ļ���
  OldCity: String; //Ϊ�����ò�ͬ���в�ͬ��ɫ
  Color1, Color2: TColor;

Implementation

{$R *.dfm}

Procedure GetCustomList(); ////��ȡ�ͻ��б�
Var
  i: Integer;
  v, Sheet: Variant;
  CellStr: String;
Begin
  If Not FileExists(CustomFileName) Then
  Begin  //MessageBox��1��������Form1.handleʱ�����������ڻ���
    MessageBox(Form1.Handle, PChar(CustomFilename + '�ļ�δ�ҵ����ļ������ڣ������Ƿ���������ƶ�λ�á�'), PChar('����'), 0);
    Exit; //�˳�������
  End;
  Try
    v := CreateOleObject('Excel.Application');
      //�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
    v.Visible := False;
    //v.WorkBooks.Add;   //�½�Excel   �ļ�
    v.workbooks.open(CustomFileName);
    Sheet := v.WorkBooks[1].WorkSheets[1];
  Except
    Showmessage('��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
    v.DisplayAlerts := false; //�Ƿ���ʾ����
    v.Quit; //�˳�Excel
    exit;
  End;
  i := 2;
  CellStr := Sheet.Cells[i, 1].Text + ' - ' + Sheet.Cells[i, 2].Text;
  CityName := Sheet.Cells[i, 1].Text; //Ϊ�����ò�ͬ���в�ͬ��ɫ
  OldCity := CityName;               //Ϊ�����ò�ͬ���в�ͬ��ɫ
  Form1.ListBox1.Items.Clear; //����б�
  color1 := clBlue;                   //Ϊ�����ò�ͬ���в�ͬ��ɫ
  Color2 := clMaroon;                 //Ϊ�����ò�ͬ���в�ͬ��ɫ
  Repeat
    Form1.listbox1.Items.Add(CellStr);
    i := i + 1;
    CellStr := Sheet.cells[i, 1].Text + ' - ' + Sheet.cells[i, 2].Text;
    CityName := Sheet.Cells[i, 1].Text; //Ϊ�����ò�ͬ���в�ͬ��ɫ
    //Cellstr:=Concat(Sheet.cells[i,1],  Sheet.cells[i,2]);
  Until CellStr = ' - ';
  v.quit;
  Form1.Label1.Caption := '�б��й���' + inttostr(i - 2) + '���ͻ���¼';
End;

Procedure SetTempletFile(); ////�趨ģ���ļ���
Var
  i: Integer;
  tmpfile: String;
Begin
  If (Not f8040) And (Not f9060) Then
    Form1.Label2.caption := '��ѡ���ǩ';
  If (Not fNetSale) And (Not fPersonSale) Then
    Form1.Label2.caption := '��ѡ��ͻ�';
  If f8040 And fNetSale Then
  Begin
    Form1.Label2.caption := '����80X40';
    templetfilename := ExtractFilePath(ParamStr(0)) + '����ģ��\����8040T.doc';
  End;
  If f8040 And fPersonSale Then
  Begin
    Form1.Label2.caption := '�ͻ�80X40';
    TempletFileName := ExtractFilePath(ParamStr(0)) + '����ģ��\�ͻ�8040T.doc';
  End;
  If f9060 And fNetSale Then
  Begin
    Form1.Label2.caption := '����90X60';
    TempletFileName := ExtractFilePath(ParamStr(0)) + '����ģ��\����9060T.doc';
  End;
  If f9060 And fPersonSale Then
  Begin
    Form1.Label2.caption := '�ͻ�90X60';
    TempletFileName := ExtractFilePath(ParamStr(0)) + '����ģ��\�ͻ�9060T.doc';
  End;
  Randomize;
  tmpfile := Format('%.5d', [Random(10000)]);
  tmpfile := ExtractFilePath(ParamStr(0)) + 'tmp\tmp' + tmpfile + '.doc'; //����ģ�帱���ļ���
  ForceDirectories(ExtractFilePath(ParamStr(0)) + 'tmp');
  If CopyFile(PChar(Templetfilename), PChar(tmpfile), true) Then//����ģ��
    TempletFileName := tmpfile;

  If fNetSale Then         ////////////////���������
  Begin
    Form1.edtPackage.Text := Trim(PackageNum);
    Form1.lblPackage.Visible := True;     ////�����а������
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
    Form1.lblPackage.Visible := False; //�ͻ�û�а������
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

Procedure OpenTempletFile();  ////��ģ���ļ�
Var
  temLabWholeSale, WholeSaleNum, PackageNum: String;
Begin
  If Trim(ExtractFileName(TempletFileName)) = '' Then
    exit;
  If Not FileExists(TempletFileName) Then
  Begin   //MessageBox��1�����������0���Ի���Ͳ�������߿���Ȩ
    MessageBox(Application.handle, PChar(TempletFilename + '�ļ�δ�ҵ����ļ������ڣ������Ƿ���������ƶ�λ�á�'), PChar('����'), 0);
    fnetsale := False; //20170822��ĳ�������ģ���ļ���ɾ������Ҫ���µ�һ�����´�
    fpersonsale := False;
    Form1.ListBox1.Clear;
    Exit; //�˳�������
  End;
  Try
    w := CreateOleObject('Word.Application');
      //�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
    w.Visible := False;
    Doc := w.Documents.Open(FileName := TempletFileName, readonly := false);
//    Doc.saveas(tmpfile);
  Except
    showmessage('��ʼ��Wordʧ�ܣ�����ûװWord�����������������������ԡ� ');
    f8040 := False;
    f9060 := False;
    Form1.label2.caption := '��ѡ���ǩ';
    w.DisplayAlerts := false; //�Ƿ���ʾ����
    w.Quit; //�˳�Excel
    exit;
  End;
//  doc.printout(); //Ҳ����
{  Doc.PrintOut(     //���������ҪUses WordXP
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
    //W.ActiveDocument.Range(Start:=0, End:=0);//��ʼ�ı����ֹλ��
    W.ActiveDocument.Range.InsertAfter(Text:=IntToStr(i));//��Word�������ַ�'myvc'
    W.ActiveDocument.Range.InsertParagraphAfter;
    //Doc.PrintOut;//��������Ҳ����
    end;
    //w.selection.endkey;
    //w.selection.typetext('������');
}
  If Doc.tables.count < 1 Then
  Begin
    ShowMessage('ģ���ļ��з���û�б��ģ��ѡ�����');
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
  If fNetSale Then         ////////////////���������
  Begin
    Form1.edtPackage.Text := Trim(PackageNum);
  End;
  //w.quit;
End;

Procedure TForm1.btnNetSaleClick(Sender: TObject);
Begin
  If fNetSale Then
    Exit; //�����ǰ��ť�ϴ��ѵ�
  If Not VarIsEmpty(w) Then//�����ǰ�򿪹��ļ�
  Try
    W.Quit;
  Except  //���W�Ѿ�������ر�
  End;
  CustomFileName := ExtractFilePath(ParamStr(0)) + '����ģ��\����.xls';
  GetCustomList();
  fNetSale := True;
  fPersonSale := False;
  SetTempletFile();
  OpenTempletFile();
End;

Procedure TForm1.btnPersonSaleClick(Sender: TObject);
Begin
  If fPersonSale Then
    Exit; //�����ǰ��ť�ϴ��ѵ�
  If Not VarIsEmpty(w) Then//�����ǰ�򿪹��ļ�
  Try
    W.Quit;
  Except
    w := varempty; //���W�Ѿ����ⲿ�رգ����Ը�������ֵ
  End;
  CustomFileName := ExtractFilePath(ParamStr(0)) + '����ģ��\�ͻ�.xls';
  GetCustomList();
  fNetSale := False;
  fPersonSale := True;
  SetTempletFile();
  OpenTempletFile();
End;

Procedure TForm1.btn8040Click(Sender: TObject);
Begin
  If f8040 Then
    Exit; //�����ǰ��ť�ϴ��ѵ�
  If Not VarIsEmpty(w) Then//�����ǰ�򿪹��ļ�
    W.Quit;

  f8040 := True;
  f9060 := False;
  SetTempletFile();
  OpenTempletFile();
End;

Procedure TForm1.btn9060Click(Sender: TObject);
Begin
  If f9060 Then
    Exit; //�����ǰ��ť�ϴ��ѵ�
  If Not VarIsEmpty(w) Then//�����ǰ�򿪹��ļ�
    W.Quit;
  f8040 := False;
  f9060 := True;
  SetTempletFile();
  OpenTempletFile();
End;

Procedure TForm1.edtPrintCopiesKeyPress(Sender: TObject; Var Key: Char);
Begin
  If Not (Key In ['0'..'9', #8]) Then
    Key := #0;  //ֻ������������
End;

Procedure TForm1.edtEndNumKeyPress(Sender: TObject; Var Key: Char);
Begin
  If Not (Key In ['0'..'9', #8]) Then
    Key := #0;  //ֻ������������
End;

Procedure TForm1.edtStartNumKeyPress(Sender: TObject; Var Key: Char);
Begin
  If Not (Key In ['0'..'9', #8]) Then
    Key := #0;  //ֻ������������
End;

Procedure ChangeFormat(); //���İ���ǰ��׺�������
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
Begin     //����ܰ�����Ϊ�գ�����Ϊ0
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
  fcustommody := false; //ÿ�ε���ͻ��б�ȡ���ֶ������ǣ������ڴ�ӡ֮ǰʹ���߸ı�����
  constr := ListBox1.Items.Strings[ListBox1.ItemIndex];
  strs := TStringList.Create;
  strs.Delimiter := '-';  //��ʲôΪ�ָ��
//strs.QuoteChar := '|';  //ȥ���ַ������˵��������
//    strs.DelimitedText := constr;//��������ַ����пո�ʱ����������
//    SplitColumns(constr, strs, '-');//�ú���û��������BUG ��Ҫuses ldStrings;
  ExtractStrings(['-'], [], PChar(constr), strs); //2017.9.1���Բ���ֿո�
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

//��ӿͻ���Excel�ļ��У��´ο�ֱ�Ӷ�ȡ
Procedure AddCustomList();
Var
  i: Integer;
  v, Sheet: Variant;
  CellStr: String;
Begin
  If Not FileExists(CustomFileName) Then
  Begin
    MessageBox(0, PChar(CustomFilename + '�ļ�δ�ҵ����ļ������ڣ������Ƿ���������ƶ�λ�á�'), PChar('����'), 0);
    Exit; //�˳�������
  End;
  Try
    v := CreateOleObject('Excel.Application');
      //�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
    v.Visible := False;
    //v.WorkBooks.Add;   //�½�Excel   �ļ�
    v.workbooks.open(CustomFileName);
    Sheet := v.WorkBooks[1].WorkSheets[1];
  Except
    Showmessage('��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
    v.DisplayAlerts := false; //�Ƿ���ʾ����
    v.Quit; //�˳�Excel
    exit;
  End;
  i := 2;
  //CellStr := Sheet.Cells[i, 1].Text + ' - ' + Sheet.Cells[i, 2].Text;
  Repeat
    i := i + 1;
    CellStr := Sheet.cells[i, 1].Text + ' - ' + Sheet.cells[i, 2].Text;
    //Cellstr:=Concat(Sheet.cells[i,1],  Sheet.cells[i,2]);
  Until CellStr = ' - ';
  Sheet.Cells[i, 1].Value := CityName;  //�����i�е�1�У�����
  Sheet.Cells[i, 2].Value := CustomName; //�����i�е�2�У�����
  v.ActiveWorkBook.Save;                //����
  v.quit;                               //�ر�
End;


////��ӡ �����汾�δ�ӡ��'����Ŀ¼'
Procedure TForm1.btnPrintClick(Sender: TObject);
Var
  lngCount, lngStart, lngEnd: Integer;
  i: Integer;
  Row: Integer; //�������ڱ���е���
  AppPath: String;
  strPre, strExt: String;
Begin
  WholeSaleNum := Trim(edtWholeSale.Text);  //��������
  PackageNum := Trim(edtPackage.Text);      //������� �����ĵ�������
  CityName := Trim(edtCity.Text);           //������
  CustomName := Trim(edtCustom.Text);       //�ͻ���
  Try
    lngCount := StrToInt(edtPrintCopies.text); //�����ܰ�����
    lngStart := StrToInt(edtStartNum.text);    //��ӡ��ʼ����
    lngEnd := StrToInt(edtEndNum.Text);        //��ӡ��������
  Except
    ShowMessage('��������Ų���Ϊ��ֵ��');
    Exit;
  End;
  If VarIsEmpty(W) Then //���WֵΪ�գ�˵����û�д�ģ���ļ�
  Begin
    MessageBox(Handle, '��Ҫ�ȵ�һ�¡��������̡��򡰿ͻ���λ����ť��' + #13#13 + '���������б�ſɵ������ط���', '��ʾ', MB_ICONWARNING);
    Exit;
  End;
  If (CityName = '') And (CustomName = '') Then
  Begin
    MessageBox(Handle, '1.����������б���ѡ��ͻ���' + #13#13 + '2.���������б���û����Ҫ�Ŀͻ������ֶ����룡', '������ʾ', MB_ICONWARNING);
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
    If MessageBox(Handle, PChar('�ܰ�������' + inttostr(lngCount) + #10#13 + '��ʼ��ţ�' + InttoStr(lngStart) + #10#13 + '������ţ�' + IntToStr(lngEnd)), '��ȷ�ϵ��š��ͻ���������Ϣ', MB_OKCancel + MB_DefButton2 + MB_ICONINFORMATION) = IDCancel Then
      Exit;
  End
  Else If MessageBox(Handle, PChar('�ܰ�������' + inttostr(lngCount) + #13#13 + 'ֻ��ӡ�ܰ�������'), '��ȷ�ϵ��š��ͻ���������Ϣ', MB_OKCancel + MB_DefButton2 + MB_ICONINFORMATION) = IDCancel Then
    Exit;

  Try
    Doc.tables.Item(1).Cell(1, 2).range.Text := WholeSaleNum; //�������ŷ������1�е�2��
    Doc.Tables.Item(1).Cell(2, 2).range.Text := CityName; //���з������2�е�2��
    Doc.Tables.Item(1).Cell(3, 2).range.Text := CustomName; //�ͻ����������3�е�2��
    Row := 4; //�ͻ����ڵ�4���ǰ�����
    If fNetSale Then
    Begin
      Doc.Tables.Item(1).Cell(4, 2).range.Text := PackageNum; //����������ŷ������4�е�2��
      Row := 5; //�������ڵ�5���ǰ�����
    End;
    For i := lngStart To lngEnd Do
    Begin
      Doc.Tables.Item(1).Cell(Row, 2).range.Text := strPre + inttostr(i) + strExt;
      Doc.Printout;
    End;
    AppPath := ExtractFilePath(Application.ExeName);
    ForceDirectories(AppPath + '����Ŀ¼');
    Doc.saveas(AppPath + '����Ŀ¼/' + CityName + '-' + CustomName);
  Except
    On E: Exception Do
      Showmessage(E.ClassName + ':' + #13#10 + E.message + #13#10 + '�����ʾ���ܾ����С�����ģ���ӡ��Χ�Ƿ񳬳�ҳ�߾�');
    //W.Quit;
  End;
  //w.quit;
  If fCustomMody Then  //����ڿͻ��������������ͱ���
    AddCustomList();
  fCustomMody := False; //������ȡ��
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
  f8040 := True; //Ĭ��Ϊ80X40�ı�ǩ�ߴ�
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
//  if fcustommody then Label2.caption:='����';
//  if VarIsEmpty(W) then
//  Label2.Caption:='W is Null'
//  else
//  Label2.Caption:='not Null';
  If CustomFileName = '' Then
  Begin
    customfilename := ExtractFilePath(ParamStr(0)) + '����ģ��';
    ShellExecute(0, nil, PChar('explorer.exe'), PChar('/e, ' + CustomFileName), nil, SW_NORMAL);
  End
  Else
    ShellExecute(0, nil, PChar('explorer.exe'), PChar('/e, ' + '/select,' + CustomFileName), nil, SW_NORMAL);

End;

Procedure TForm1.FormDestroy(Sender: TObject);
Begin
  If Not VarIsEmpty(w) Then//�����ǰ�򿪹��ļ�
  Try
    W.Quit;
  Except //���W�Ѿ�������ر�
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

//���㵽Edit�ؼ���ʱ��ѡ��Edit�ؼ��е�����
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
//Ϊ�������´ν���������ѡ��ȫ�����ݵ�Ч������Ҫ��䡣���ø���

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
  If Oldcity <> str Then      //���ǰһ�к͵�ǰ�в�ͬ������ɫ
  Begin
    ListBox1.Canvas.FillRect(Rect);
    ListBox1.Canvas.Font.Color := Color2;
    Color2 := Color1;
    Color1 := ListBox1.Canvas.Font.Color;
  End
  Else
  Begin                      //���򣬻���ԭ����ɫ
    ListBox1.Canvas.FillRect(Rect);
    ListBox1.Canvas.Font.Color := Color1;
  End;

{  position := AnsiPos('��', str);
  if position > 0 then
  begin
    ListBox1.Canvas.FillRect(Rect);
    ListBox1.Canvas.Font.Color := clGreen;
  end;
 }
  ListBox1.Canvas.TextOut(Rect.Left + 1, Rect.Top + 0, ListBox1.Items.Strings[Index]);
  OldCity := str;
  If odSelected In State Then        //��ѡ��ʱ����ɫ
  Begin
    listbox1.Canvas.Brush.Color := clhighlight;
    ListBox1.Canvas.TextRect(Rect, Rect.Left, Rect.Top, ListBox1.Items[Index]);
  End;
End;

End.

