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
  CustomFileName: string; //客户列表文件
  TempletFileName: string; //模板文件
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
      '文件未找到或文件不存在，请检查是否改名或者移动位置。'), PChar('错误'), 0);
    Exit; //退出本过程
  end;
  try
    v := CreateOleObject('Excel.Application');
      //如果这个字符串中有空格就会出'无效的类别字符串'的错
    v.Visible := False;
    //v.WorkBooks.Add;   //新建Excel   文件
    v.workbooks.open(CustomFileName);
    //v.WorkBooks[1].WorkSheets[1].name:= '电脑报';//第一页标题
    //v.WorkBooks[1].WorkSheets[2].name:= '编程乐园';
    //v.WorkBooks[1].WorkSheets[3].name:= '都来看呀 ';
//                v.activesheet.columns[1].columnwidth:=20;//设置第1列宽
    //v.Activesheet.usedrange.font.size:=8;
    Sheet := v.WorkBooks[1].WorkSheets[1];
    //                v.workBooks[1].Styles['Normal'].Font.size:=10;//设置第1个工作表的字号
    {                range:=Sheet.cells.select;
                    v.selection.font.size:=10;
                    v.Selection.FormatConditions.Delete;
                    v.Selection.FormatConditions.Add(1, 3,'0');
                    v.Selection.FormatConditions[1].Font.ColorIndex := 15;
    }
    {                v.range['a2'].HorizontalAlignment := 3;
                    sheet.cells[2,1]:='合计';
                    v.range['b:k'].FormatConditions.Delete;
                    v.range['b:k'].FormatConditions.Add(1, 3,'0');
                    v.range['b:k'].FormatConditions[1].Font.ColorIndex := 15;

                    v.range['b2:k2'].FormatConditions[3].Font.ColorIndex := 3;
                    v.range['b:k'].columnwidth:=7.5;
                    sheet.range['b3'].select;
                    v.ActiveWindow.FreezePanes:= True;
                    sheet.range['a1'].select;
                    //Sheet.Cells[1,1]:= '好看 ';//单元格内容
                    //Sheet.Cells[1,2]:= '确实 ';
      }//Sheet.Cells[2,1]:= '我喜欢 ';
    {                v.activesheet.rows[2].select;
                    v.Selection.Interior.ColorIndex:= 32;
                    v.range['b2:j2'].interior.colorindex:=39;//粉色
     }
  except
    showmessage('初始化Excel失败，可能没装Excel，或者其他错误；请重新再试。 ');
    v.DisplayAlerts := false; //是否提示存盘
    v.Quit; //退出Excel
    exit;
  end;
  i := 2;
  CellStr := Sheet.Cells[i, 1].Text + ' - ' + Sheet.Cells[i, 2].Text;
  ListBox1.Items.Clear;//清空列表

  repeat
    listbox1.Items.Add(CellStr);
    i := i + 1;
    CellStr := Sheet.cells[i, 1].Text + ' - ' + Sheet.cells[i, 2].Text;
    //Cellstr:=Concat(Sheet.cells[i,1],  Sheet.cells[i,2]);
  until CellStr = ' - ';
  v.quit;
  Label1.Caption:='共有'+inttostr(i-2)+'条记录';
end;
//获取文件夹下的文件
procedure  SearchFile(path:string);//注意,path后面要有'\';
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
//合并文件
procedure TForm1.Button2Click(Sender: TObject);
var
  City, Custom: string;
  w: Variant;
  Doc: OleVariant;
  i: Integer;
  CurrentPath:string;//当前路径
begin
  currentpath:=ExtractFilePath(Application.ExeName); //带“/”
//  currentpath:=ExtractFileDir(Application.ExeName);//不带/
  ForceDirectories(CurrentPath+'生成目录');
  if not FileExists(TempletFileName) then
  begin
    showMessage(TempletFilename +
      '文件未找到或文件不存在，请检查是否改名或者移动位置。');
    Exit; //退出本过程
  end;
  if not FileExists(CustomFileName) then
  begin
    showMessage(CustomFilename +
      '文件未找到或文件不存在，请检查是否改名或者移动位置。');
    Exit; //退出本过程
  end;

  try
    w := CreateOleObject('Word.Application');
      //如果这个字符串中有空格就会出'无效的类别字符串'的错
    w.Visible := True;
    Doc := w.Documents.Open(FileName := templetfilename, readonly := true);

    {var MSWord: Variant;
     begin
    MSWord := CreateOLEObject('Word.Application');//连接Word
    MSWord.Documents.Open(FileName:='D:\Temp\temp.doc', ReadOnly:=True);
    //打开外部Word文档
    MSWord.Visible := 1;//是否显示文件编辑
    MSWord.ActiveDocument.Range(Start:=0, End:=0);//开始改变的启止位置
    MSWord.ActiveDocument.Range.InsertAfter(Text:='myvc');//在Word中增加字符'myvc'
    MSWord.ActiveDocument.Range.InsertParagraphAfter;
    MSWord.ActiveDocument.Range.Font.Size := 72;//字体大小
    MSWord.ActiveDocument.Range.Font.Name := 'Arial';//字体名称}
  except
    showmessage('初始化Word失败，可能没装Word，或者其他错误；请重新再试。 ');
    w.DisplayAlerts := false; //是否提示存盘
    w.Quit; //退出Excel
    exit;
  end;

  try
    v := CreateOleObject('Excel.Application');
      //如果这个字符串中有空格就会出'无效的类别字符串'的错
    v.Visible := False;
    v.workbooks.open(CustomFileName);
    Sheet := v.WorkBooks[1].WorkSheets[1];
  except
    showmessage('初始化Excel失败，可能没装Excel，或者其他错误；请重新再试。 ');
    v.DisplayAlerts := false; //是否提示存盘
    v.Quit; //退出Excel
    exit;
  end;
  i := 2;
  City := Sheet.Cells[i, 1]; //  第1列为城市名
  Custom := Sheet.Cells[i, 2]; //第2列为客户名
  ListBox2.Items.Clear;//清空显示列表
  while City <> '' do
  begin
    try
    Doc.tables.item(1).Cell(1,2).range.Text:='123456789';
    Doc.tables.item(1).Cell(2, 2).range.TEXT := City; //Cell(行，列)
    Doc.tables.item(1).Cell(3, 2).range.TEXT := Custom; //Cell(行，列)
    Doc.saveas(Currentpath+'生成目录/'+City + '-' + Custom);
    i := i + 1;
    Label1.Caption:='已处理'+inttostr(i-2)+'条记录';
    City := Sheet.Cells[i, 1];
    Custom := Sheet.Cells[i, 2];
    searchfile(Currentpath+'生成目录\');
    except
      on e:exception do
      ShowMessage(e.message);

    end;
  end;
  w.quit;
  v.DisplayAlerts := false; //是否提示存盘
  v.quit;

end;


//选择客户列表源文件

procedure TForm1.Button3Click(Sender: TObject);
begin
  OpenDialog1.Filter := 'Excel文件（*.xls;*.xlsx）|*.xls;*.xlsx';
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
  //加背景  ////////////////////////////////////
  try
  Bmp := TBitmap.Create;
  //Bmp.LoadFromFile('C:\test56.bmp');
  Randomize;
  seed:=1010+random(5);
  Bmp.LoadFromResourceID(0,seed);
  Brush.Bitmap := Bmp;//Image1.Picture.Bitmap;
  except
     on E: Exception do
      showmessage(E.ClassName+':' + #13#10 +E.message + #13#10 +'不过没关系。');
  end;
  //////////////////////////////////////////////
  customfilename := 'D:\Backup\我的文档\工作簿1.xlsx';
  templetfilename := 'D:\Backup\我的文档\店铺模板2表格版.docx';
  //设置需要处理文件WM_DROPFILES拖放消息
  DragAcceptFiles(Edit1.Handle, TRUE); //让Edit1接受文件拖放
  DragAcceptFiles(Edit2.Handle, TRUE); //让Edit2接受文件拖放
  //设置AppMessage过程来捕获所有消息
  Application.OnMessage := AppMessage;
end;

//选择发货单模板文件

procedure TForm1.Button4Click(Sender: TObject);
begin
  OpenDialog1.Filter := 'Word文件（*.doc;*.docx）|*.doc;*.docx';
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
  // 注意！所有消息都将通过这里！
  // 不要在此过程中编写过多的或者需要长时间操作的代码，否则将影响程序的性能
  //
  // 判断是否是发送到指定控件的WM_DROPFILES消息
  if (Msg.message = WM_DROPFILES) and (Msg.hwnd = Edit1.Handle) then
  begin
    // 取dropped files的数量
    nFiles := DragQueryFile(Msg.wParam, $FFFFFFFF, nil, 0);
    // 循环取每个拖下文件的全文件名
    try
      for I := 0 to nFiles - 1 do
      begin
        // 为文件名分配缓冲 allocate memory
        SetLength(Filename, 80);
        // 取文件名 read the file name
        DragQueryFile(Msg.wParam, I, PChar(Filename), 80);
        Filename := PChar(Filename);
        //将全文件名分解程文件名和路径
        Edit1.text := Filename;
        if (UpperCase(ExtractFileExt(Filename)) = '.XLS') or
          (UpperCase(ExtractFileExt(Filename)) = '.XLSX') then
          CustomFileName := Filename //客户列表
        else
          MessageBox(0, '文件类型不对！', '要求是Excel文件', 0);
      end;
    finally
      //结束这次拖放操作
      DragFinish(Msg.wParam);
    end;
    //标识已处理了这条消息
    Handled := True;
  end;

  if (Msg.message = WM_DROPFILES) and (Msg.hwnd = Edit2.Handle) then
  begin
    // 取dropped files的数量
    nFiles := DragQueryFile(Msg.wParam, $FFFFFFFF, nil, 0);
    // 循环取每个拖下文件的全文件名
    try
      for I := 0 to nFiles - 1 do
      begin
        // 为文件名分配缓冲 allocate memory
        SetLength(Filename, 80);
        // 取文件名 read the file name
        DragQueryFile(Msg.wParam, I, PChar(Filename), 80);
        Filename := PChar(Filename);
        //将全文件名分解程文件名和路径
        Edit2.text := Filename;
        if (UpperCase(ExtractFileExt(Filename)) = '.DOC') or
          (UpperCase(ExtractFileExt(Filename)) = '.DOCX') then
          TempletFileName := Filename //模板文件
        else
          MessageBox(0, '文件类型不对！', '要求是Word文件', 0);
      end;
    finally
      //结束这次拖放操作
      DragFinish(Msg.wParam);
    end;
    //标识已处理了这条消息
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

