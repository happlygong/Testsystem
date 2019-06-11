unit list_Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, tlhelp32, StdCtrls, ComCtrls, PSAPI, Buttons, shellapi;

type
  TForm1 = class(TForm)
    ListBox1: TListBox;
    Button1: TButton;
    Label1: TLabel;
    Edit1: TEdit;
    ListView1: TListView;
    Button2: TButton;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    BitBtn1: TBitBtn;
    procedure Button1Click(Sender: TObject);
    procedure ListBox1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure ListView1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
function My_RefreshForm(MyStringList: TStringList): Boolean;

implementation

{$R *.dfm}
function My_RefreshForm(MyStringList: TStringList): Boolean;
var
  hCurrentWindow: HWnd;
  szText: array[0..254] of Char;
  lppe: tprocessentry32;
  sshandle: thandle;
  hh: hwnd;
  found: Boolean;
begin
  MyStringList.Clear;
  sshandle := createtoolhelp32snapshot(TH32CS_SNAPALL, 0);
  found := process32first(sshandle, lppe);
  while found do
  begin
    MyStringList.Add(StrPas(@lppe.szExeFile));
  {if (uppercase(extractfilename(lppe.szExeFile))=s) or (uppercase  (lppe.szExeFile)=s) then
    begin
      hh:=OpenProcess(PROCESS_ALL_ACCESS,true,lppe.th32ProcessID);
      TerminateProcess(hh,0);
    end;
    }
    found := process32next(sshandle, lppe);
  end;
  CloseHandle(sshandle);

  Result := True;
end;


procedure TForm1.Button1Click(Sender: TObject);
var
  FormStrings: TStringList;
begin
  FormStrings := TStringList.Create;
  My_RefreshForm(FormStrings);
  ListBox1.Items := FormStrings;
  FormStrings.Free;
  label1.Caption := IntToStr(listbox1.Count);
  {
  }
end;

procedure TForm1.ListBox1Click(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to (listbox1.Count - 1) do
  begin
    try
      if listbox1.Selected[i] then
        edit1.Text := listbox1.Items.Strings[i];
    finally
    end;
  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  processIDs: array[0..$3FFF - 1] of DWORD;  //进程ID
  Count: DWORD;
  i: Integer;
  ProcHand: THandle; //进程的句柄
  ModHand: HMODULE;  //模块的句柄
  ModName: array[0..MAX_PATH] of Char;//模块文件名
begin
  ListView1.Items.Clear;
  //列举所有正在运行的进程
  if not EnumProcesses(@processIDs, SizeOf(processIDs), Count) then
    raise Exception.Create('列举进程出错，确认是否安装了PSAPI.DLL!');
  for i := 0 to (Count div Sizeof(DWORD)) - 1 do
  begin
    ProcHand := OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_VM_READ,
      False, processIDs[i]);//查询方式打开进程
    if ProcHand > 0 then
      try
        //列举进程的首模块
        if EnumProcessModules(Prochand, @ModHand, sizeof(ModHand),
          Count)        //获取模块的文件名
          then
          if GetModuleFileNameEx(Prochand, ModHand, ModName,
            SizeOf(ModName)) > 0 then
          begin
            with ListView1.Items.Add, SubItems do
            begin
              Caption := ModName;
              Data := Pointer(processIDs[I]);//保存进程的ID
              Add(IntToStr(processIDs[I]));
            end;
          end;
      finally
        CloseHandle(ProcHand);
      end;
  end;
end;

procedure TForm1.ListView1Click(Sender: TObject);
begin
  edit1.Text := extractfilepath(listview1.Selected.Caption);
  label3.Caption := extractfilepath(listview1.Selected.Caption);
  label5.Caption := extractfilename(listview1.Selected.Caption);
  bitbtn1.Enabled := True;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
  shellexecutea(0, 'explore', PChar(label3.Caption), nil, nil, sw_shownormal);
end;

end.

