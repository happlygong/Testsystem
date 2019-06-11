unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Controls, Forms,
  Dialogs, StdCtrls, ComObj, ShellAPI, ExtCtrls,Graphics,
  StrUtils, rxAnimate, rxGIFCtrl ;

type
  TForm1 = class(TForm)
    ListBox1: TListBox;
    Button1: TButton;
    Button2: TButton;
    OpenDialog1: TOpenDialog;
    PrintDialog1: TPrintDialog;
    Edit1: TEdit;
    Edit2: TEdit;
    Button3: TButton;
    Button4: TButton;
    Label1: TLabel;
    ListBox2: TListBox;
    RxGIFAnimator1: TRxGIFAnimator;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ListBox2DblClick(Sender: TObject);
    
  private
    { Private declarations }
  public
    { Public declarations }
    procedure AppMessage(var Msg: TMsg; var Handled: Boolean);
  end;

var
  Form1: TForm1;
  v, range: Variant;
  Sheet: Variant;
  CustomFileName: string; //�ͻ��б��ļ�
  TempletFileName: string; //ģ���ļ�
  Bmp: TBitmap;
implementation

{$R *.dfm}

//{$R 'myPic.res'}
procedure TForm1.Button1Click(Sender: TObject);
var
  CellStr: string;
  i: Integer;
begin
  if not FileExists(CustomFileName) then
  begin
    MessageBox(0, PChar(CustomFilename +
      '�ļ�δ�ҵ����ļ������ڣ������Ƿ���������ƶ�λ�á�'), PChar('����'), 0);
    Exit; //�˳�������
  end;
  try
    v := CreateOleObject('Excel.Application');
      //�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
    v.Visible := False;
    //v.WorkBooks.Add;   //�½�Excel   �ļ�
    v.workbooks.open(CustomFileName);
    //v.WorkBooks[1].WorkSheets[1].name:= '���Ա�';//��һҳ����
    //v.WorkBooks[1].WorkSheets[2].name:= '�����԰';
    //v.WorkBooks[1].WorkSheets[3].name:= '������ѽ ';
//                v.activesheet.columns[1].columnwidth:=20;//���õ�1�п�
    //v.Activesheet.usedrange.font.size:=8;
    Sheet := v.WorkBooks[1].WorkSheets[1];
    //                v.workBooks[1].Styles['Normal'].Font.size:=10;//���õ�1����������ֺ�
    {                range:=Sheet.cells.select;
                    v.selection.font.size:=10;
                    v.Selection.FormatConditions.Delete;
                    v.Selection.FormatConditions.Add(1, 3,'0');
                    v.Selection.FormatConditions[1].Font.ColorIndex := 15;
    }
    {                v.range['a2'].HorizontalAlignment := 3;
                    sheet.cells[2,1]:='�ϼ�';
                    v.range['b:k'].FormatConditions.Delete;
                    v.range['b:k'].FormatConditions.Add(1, 3,'0');
                    v.range['b:k'].FormatConditions[1].Font.ColorIndex := 15;

                    v.range['b2:k2'].FormatConditions[3].Font.ColorIndex := 3;
                    v.range['b:k'].columnwidth:=7.5;
                    sheet.range['b3'].select;
                    v.ActiveWindow.FreezePanes:= True;
                    sheet.range['a1'].select;
                    //Sheet.Cells[1,1]:= '�ÿ� ';//��Ԫ������
                    //Sheet.Cells[1,2]:= 'ȷʵ ';
      }//Sheet.Cells[2,1]:= '��ϲ�� ';
    {                v.activesheet.rows[2].select;
                    v.Selection.Interior.ColorIndex:= 32;
                    v.range['b2:j2'].interior.colorindex:=39;//��ɫ
     }
  except
    showmessage('��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
    v.DisplayAlerts := false; //�Ƿ���ʾ����
    v.Quit; //�˳�Excel
    exit;
  end;
  i := 2;
  CellStr := Sheet.Cells[i, 1].Text + ' - ' + Sheet.Cells[i, 2].Text;
  ListBox1.Items.Clear;//����б�

  repeat
    listbox1.Items.Add(CellStr);
    i := i + 1;
    CellStr := Sheet.cells[i, 1].Text + ' - ' + Sheet.cells[i, 2].Text;
    //Cellstr:=Concat(Sheet.cells[i,1],  Sheet.cells[i,2]);
  until CellStr = ' - ';
  v.quit;
  Label1.Caption:='����'+inttostr(i-2)+'����¼';
end;
//��ȡ�ļ����µ��ļ�
procedure  SearchFile(path:string);//ע��,path����Ҫ��'\';
   var   
       SearchRec:TSearchRec;   
       found:integer;   
   begin   
       Form1.ListBox2.Clear;
       found:=FindFirst(path+'*.*',faAnyFile,SearchRec);
       while    found=0    do
         begin
             if (SearchRec.Name<>'.')and(SearchRec.Name<>'..')
             and(SearchRec.Attr<>faDirectory) and (leftstr(SearchRec.Name,2)<>'~$') then
                 Form1.ListBox2.items.add(path+SearchRec.Name);
             found:=FindNext(SearchRec);
         end;
       FindClose(SearchRec);   
   end;
//�ϲ��ļ�
procedure TForm1.Button2Click(Sender: TObject);
var
  City, Custom: string;
  w: Variant;
  Doc: OleVariant;
  i: Integer;
  CurrentPath:string;//��ǰ·��
begin
  currentpath:=ExtractFilePath(Application.ExeName); //����/��
//  currentpath:=ExtractFileDir(Application.ExeName);//����/
  ForceDirectories(CurrentPath+'����Ŀ¼');
  if not FileExists(TempletFileName) then
  begin
    showMessage(TempletFilename +
      '�ļ�δ�ҵ����ļ������ڣ������Ƿ���������ƶ�λ�á�');
    Exit; //�˳�������
  end;
  if not FileExists(CustomFileName) then
  begin
    showMessage(CustomFilename +
      '�ļ�δ�ҵ����ļ������ڣ������Ƿ���������ƶ�λ�á�');
    Exit; //�˳�������
  end;

  try
    w := CreateOleObject('Word.Application');
      //�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
    w.Visible := True;
    Doc := w.Documents.Open(FileName := templetfilename, readonly := true);

    {var MSWord: Variant;
     begin
    MSWord := CreateOLEObject('Word.Application');//����Word
    MSWord.Documents.Open(FileName:='D:\Temp\temp.doc', ReadOnly:=True);
    //���ⲿWord�ĵ�
    MSWord.Visible := 1;//�Ƿ���ʾ�ļ��༭
    MSWord.ActiveDocument.Range(Start:=0, End:=0);//��ʼ�ı����ֹλ��
    MSWord.ActiveDocument.Range.InsertAfter(Text:='myvc');//��Word�������ַ�'myvc'
    MSWord.ActiveDocument.Range.InsertParagraphAfter;
    MSWord.ActiveDocument.Range.Font.Size := 72;//�����С
    MSWord.ActiveDocument.Range.Font.Name := 'Arial';//��������}
  except
    showmessage('��ʼ��Wordʧ�ܣ�����ûװWord�����������������������ԡ� ');
    w.DisplayAlerts := false; //�Ƿ���ʾ����
    w.Quit; //�˳�Excel
    exit;
  end;

  try
    v := CreateOleObject('Excel.Application');
      //�������ַ������пո�ͻ��'��Ч������ַ���'�Ĵ�
    v.Visible := False;
    v.workbooks.open(CustomFileName);
    Sheet := v.WorkBooks[1].WorkSheets[1];
  except
    showmessage('��ʼ��Excelʧ�ܣ�����ûװExcel�����������������������ԡ� ');
    v.DisplayAlerts := false; //�Ƿ���ʾ����
    v.Quit; //�˳�Excel
    exit;
  end;
  i := 2;
  City := Sheet.Cells[i, 1]; //  ��1��Ϊ������
  Custom := Sheet.Cells[i, 2]; //��2��Ϊ�ͻ���
  ListBox2.Items.Clear;//�����ʾ�б�
  while City <> '' do
  begin
    try
    Doc.tables.item(1).Cell(1,2).range.Text:='123456789';
    Doc.tables.item(1).Cell(2, 2).range.TEXT := City; //Cell(�У���)
    Doc.tables.item(1).Cell(3, 2).range.TEXT := Custom; //Cell(�У���)
    Doc.saveas(Currentpath+'����Ŀ¼/'+City + '-' + Custom);
    i := i + 1;
    Label1.Caption:='�Ѵ���'+inttostr(i-2)+'����¼';
    City := Sheet.Cells[i, 1];
    Custom := Sheet.Cells[i, 2];
    searchfile(Currentpath+'����Ŀ¼\');
    except
      on e:exception do
      ShowMessage(e.message);

    end;
  end;
  w.quit;
  v.DisplayAlerts := false; //�Ƿ���ʾ����
  v.quit;

end;


//ѡ��ͻ��б�Դ�ļ�

procedure TForm1.Button3Click(Sender: TObject);
begin
  OpenDialog1.Filter := 'Excel�ļ���*.xls;*.xlsx��|*.xls;*.xlsx';
  if OpenDialog1.Execute then
  begin
    Edit1.Text := OpenDialog1.FileName;
    customfilename := OpenDialog1.FileName;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  seed:Integer;
begin
  //�ӱ���  ////////////////////////////////////
  try
  Bmp := TBitmap.Create;
  //Bmp.LoadFromFile('C:\test56.bmp');
  Randomize;
  seed:=1010+random(5);
  Bmp.LoadFromResourceID(0,seed);
  Brush.Bitmap := Bmp;//Image1.Picture.Bitmap;
  except
     on E: Exception do
      showmessage(E.ClassName+':' + #13#10 +E.message + #13#10 +'����û��ϵ��');
  end;
  //////////////////////////////////////////////
  customfilename := 'D:\Backup\�ҵ��ĵ�\������1.xlsx';
  templetfilename := 'D:\Backup\�ҵ��ĵ�\����ģ��2����.docx';
  //������Ҫ�����ļ�WM_DROPFILES�Ϸ���Ϣ
  DragAcceptFiles(Edit1.Handle, TRUE); //��Edit1�����ļ��Ϸ�
  DragAcceptFiles(Edit2.Handle, TRUE); //��Edit2�����ļ��Ϸ�
  //����AppMessage����������������Ϣ
  Application.OnMessage := AppMessage;
end;

//ѡ�񷢻���ģ���ļ�

procedure TForm1.Button4Click(Sender: TObject);
begin
  OpenDialog1.Filter := 'Word�ļ���*.doc;*.docx��|*.doc;*.docx';
  if OpenDialog1.Execute then
  begin
    Edit2.Text := OpenDialog1.FileName;
    TempletFileName := OpenDialog1.FileName;
  end;
end;

procedure TForm1.AppMessage(var Msg: TMsg; var Handled: Boolean);
var
  nFiles, I: Integer;
  Filename: string;
begin
  //
  // ע�⣡������Ϣ����ͨ�����
  // ��Ҫ�ڴ˹����б�д����Ļ�����Ҫ��ʱ������Ĵ��룬����Ӱ����������
  //
  // �ж��Ƿ��Ƿ��͵�ָ���ؼ���WM_DROPFILES��Ϣ
  if (Msg.message = WM_DROPFILES) and (Msg.hwnd = Edit1.Handle) then
  begin
    // ȡdropped files������
    nFiles := DragQueryFile(Msg.wParam, $FFFFFFFF, nil, 0);
    // ѭ��ȡÿ�������ļ���ȫ�ļ���
    try
      for I := 0 to nFiles - 1 do
      begin
        // Ϊ�ļ������仺�� allocate memory
        SetLength(Filename, 80);
        // ȡ�ļ��� read the file name
        DragQueryFile(Msg.wParam, I, PChar(Filename), 80);
        Filename := PChar(Filename);
        //��ȫ�ļ����ֽ���ļ�����·��
        Edit1.text := Filename;
        if (UpperCase(ExtractFileExt(Filename)) = '.XLS') or
          (UpperCase(ExtractFileExt(Filename)) = '.XLSX') then
          CustomFileName := Filename //�ͻ��б�
        else
          MessageBox(0, '�ļ����Ͳ��ԣ�', 'Ҫ����Excel�ļ�', 0);
      end;
    finally
      //��������ϷŲ���
      DragFinish(Msg.wParam);
    end;
    //��ʶ�Ѵ�����������Ϣ
    Handled := True;
  end;

  if (Msg.message = WM_DROPFILES) and (Msg.hwnd = Edit2.Handle) then
  begin
    // ȡdropped files������
    nFiles := DragQueryFile(Msg.wParam, $FFFFFFFF, nil, 0);
    // ѭ��ȡÿ�������ļ���ȫ�ļ���
    try
      for I := 0 to nFiles - 1 do
      begin
        // Ϊ�ļ������仺�� allocate memory
        SetLength(Filename, 80);
        // ȡ�ļ��� read the file name
        DragQueryFile(Msg.wParam, I, PChar(Filename), 80);
        Filename := PChar(Filename);
        //��ȫ�ļ����ֽ���ļ�����·��
        Edit2.text := Filename;
        if (UpperCase(ExtractFileExt(Filename)) = '.DOC') or
          (UpperCase(ExtractFileExt(Filename)) = '.DOCX') then
          TempletFileName := Filename //ģ���ļ�
        else
          MessageBox(0, '�ļ����Ͳ��ԣ�', 'Ҫ����Word�ļ�', 0);
      end;
    finally
      //��������ϷŲ���
      DragFinish(Msg.wParam);
    end;
    //��ʶ�Ѵ�����������Ϣ
    Handled := True;
  end;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  Bmp.free;
end;

procedure TForm1.ListBox2DblClick(Sender: TObject);
var
  strFileName:string;
begin
  strfilename:=ListBox2.Items.Strings[ListBox2.ItemIndex];
  ShellExecute(0, nil, PChar('explorer.exe'),PChar('/e, ' + '/select,' + strFileName), nil, SW_NORMAL);
end;

end.

