unit searchu;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ToolWin, ExtCtrls, Menus, Buttons, CheckLst,
  StrUtils;

type
  TForm1 = class(TForm)
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    Basict: TGroupBox;
    goods: TGroupBox;
    atk: TGroupBox;
    ck: TCheckListBox;
    Button4: TButton;
    Button3: TButton;
    Button2: TButton;
    Button1: TButton;
    LED_Level: TLabeledEdit;
    LED_Exper: TLabeledEdit;
    LED_HP: TLabeledEdit;
    LED_MAXHP: TLabeledEdit;
    LED_RAGE: TLabeledEdit;
    LED_MRAGE: TLabeledEdit;
    LED_MP: TLabeledEdit;
    LED_MAXMP: TLabeledEdit;
    LED_Wuli1: TLabeledEdit;
    LED_Fang1: TLabeledEdit;
    LED_Lucky1: TLabeledEdit;
    LED_Will1: TLabeledEdit;
    LED_Speed1: TLabeledEdit;
    LED_Wuli: TLabeledEdit;
    LED_Fang: TLabeledEdit;
    LED_Lucky: TLabeledEdit;
    LED_Speed: TLabeledEdit;
    LED_Will: TLabeledEdit;
    atkall: TCheckBox;
    inteam: TCheckBox;
    Hostcheck: TCheckBox;
    Button5: TButton;
    Button6: TButton;
    BitBtn3: TBitBtn;
    LMoney: TLabeledEdit;
    BitBtn1: TBitBtn;
    player: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    BitBtn2: TBitBtn;
    wupin: TCheckBox;
    tupu: TCheckBox;
    procedure Button1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure TabSheet2Show(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure ckClickCheck(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormPaint(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure LED_LevelKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  filepathname: string;
  Levelad, level: DWord;
  Experad, exper: DWord;
  HPad, HP: DWord;
  MaxHPad, MaxHP: DWord;
  Ragead, Rage: DWord;
  MPad, MP: DWord;
  MAXMPad, MaxMP: DWord;
  Wuliad, Wuli: Dword;
  Fangad, Fang: Dword;
  Speedad, Speed: Dword;
  Luckyad, Lucky: Dword;
  Willad, Will: Dword;
  Wuli1ad, Wuli1: Dword;
  Fang1ad, Fang1: Dword;
  Speed1ad, Speed1: Dword;
  Lucky1ad, Lucky1: Dword;
  Will1ad, Will1: Dword;
  //tianxiangad,huishengad,zhiningwanad,hanlingguoad,tougudingad:Dword;
  //zhentianliead:dword;
  maxhp1ad, maxmp1ad: Dword;
  moneyad: dword;
  AddressMemStream: TMemoryStream;
  tempFileMemStream: TMemoryStream;
  firstad, midad, endad: Dword;
  inteamad, atkallad, atkall2ad: DWord;
  conatk: Integer;
  Tp: TSearchRec;

implementation

uses
  aboutU;

  {$R *.dfm}
  {$R data.res}

//转换文件时间
function CovFileDate(Fd: _FileTime): TDateTime;
var
  Tct: _SystemTime;
  Temp: _FileTime;
begin
  FileTimeToLocalFileTime(Fd, Temp);
  FileTimeToSystemTime(Temp, Tct);
  CovFileDate := SystemTimeToDateTime(Tct);
end;

procedure getvalue(tempaddress: DWord); //实际值
begin
  if tempaddress = 999999999 then
    exit;
  //    experad :=tempaddress+821;        // 经验值
  hpad := tempaddress + 861;        // 精
  levelad := tempaddress + 911;        // 等级
  luckyad := tempaddress + 990;        // 运气
  maxhp1ad := tempaddress + 1070;       // 最大精
  maxmp1ad := tempaddress + 1091;       // 最大神
  mpad := tempaddress + 1108;       // 神
  wuliad := tempaddress + 1213;       // 武力
  ragead := tempaddress + 1268;       // 气
  speedad := tempaddress + 1288;       // 速度
  fangad := tempaddress + 1345;       // 防御
  willad := tempaddress + 1364;       //灵力？

  with form1 do
  begin
    {    tempFileMemstream.Seek(experad,sofrombeginning);
        tempFileMemStream.ReadBuffer(exper,4);
        LED_Exper.Text:=inttostr(exper); }
    tempFileMemstream.Seek(hpad, sofrombeginning);
    tempFileMemStream.ReadBuffer(hp, 4);
    LED_hp.Text := IntToStr(hp);
    tempFileMemstream.Seek(levelad, sofrombeginning);
    tempFileMemStream.ReadBuffer(level, 4);
    LED_level.Text := IntToStr(level);
    tempFileMemstream.Seek(luckyad, sofrombeginning);
    tempFileMemStream.ReadBuffer(lucky, 4);
    LED_lucky.Text := IntToStr(lucky);
    tempFileMemstream.Seek(mpad, sofrombeginning);
    tempFileMemStream.ReadBuffer(mp, 4);
    LED_MP.Text := IntToStr(mp);
    tempFileMemstream.Seek(wuliad, sofrombeginning);
    tempFileMemStream.ReadBuffer(wuli, 4);
    LED_wuli.Text := IntToStr(wuli);
    tempFileMemstream.Seek(ragead, sofrombeginning);
    tempFileMemStream.ReadBuffer(rage, 4);
    LED_rage.Text := IntToStr(rage);
    LED_Mrage.Text := '100';
    tempFileMemstream.Seek(Fangad, sofrombeginning);
    tempFileMemStream.ReadBuffer(Fang, 4);
    LED_Fang.Text := IntToStr(Fang);
    tempFileMemstream.Seek(Speedad, sofrombeginning);
    tempFileMemStream.ReadBuffer(Speed, 4);
    LED_Speed.Text := IntToStr(Speed);
    tempFileMemstream.Seek(Willad, sofrombeginning);
    tempFileMemStream.ReadBuffer(Will, 4);
    LED_Will.Text := IntToStr(Will);
  end;
end;

procedure getvalue2(tempaddress: DWord); //基础值
begin
  if tempaddress = 999999999 then
    exit;
  experad := tempaddress + 820;         // 经验值
  maxhpad := tempaddress + 1069;        // 最大精
  maxmpad := tempaddress + 1090;        // 最大神
  lucky1ad := tempaddress + 989;        // 运气
  wuli1ad := tempaddress + 1212;        // 武力
  speed1ad := tempaddress + 1287;       // 速度
  fang1ad := tempaddress + 1344;        // 防御
  will1ad := tempaddress + 1363;        //灵力？

  with form1 do
  begin
    tempFileMemstream.Seek(experad, sofrombeginning);
    tempFileMemStream.ReadBuffer(exper, 4);
    LED_Exper.Text := IntToStr(exper);
    tempFileMemstream.Seek(lucky1ad, sofrombeginning);
    tempFileMemStream.ReadBuffer(lucky1, 4);
    LED_lucky1.Text := IntToStr(lucky1);
    tempFileMemstream.Seek(wuli1ad, sofrombeginning);
    tempFileMemStream.ReadBuffer(wuli1, 4);
    LED_wuli1.Text := IntToStr(wuli1);
    tempFileMemstream.Seek(Fang1ad, sofrombeginning);
    tempFileMemStream.ReadBuffer(Fang1, 4);
    LED_Fang1.Text := IntToStr(Fang1);
    tempFileMemstream.Seek(Speed1ad, sofrombeginning);
    tempFileMemStream.ReadBuffer(Speed1, 4);
    LED_Speed1.Text := IntToStr(Speed1);
    tempFileMemstream.Seek(Will1ad, sofrombeginning);
    tempFileMemStream.ReadBuffer(Will1, 4);
    LED_Will1.Text := IntToStr(Will1);
    tempFileMemstream.Seek(maxhpad, sofrombeginning);
    tempFileMemStream.ReadBuffer(maxhp, 4);
    LED_maxhp.Text := IntToStr(maxhp);
    tempFileMemstream.Seek(maxmpad, sofrombeginning);
    tempFileMemStream.ReadBuffer(maxmp, 4);
    LED_maxmp.Text := IntToStr(maxmp);
  end;
end;

function Getaddress(s1: string; Count: Integer): DWord;
var
  tempAddress: Dword;
  chartoFind: Integer;
  i, j: Integer;
  helpStringLength: Integer;
  TempLength: DWORD;               ////字节大小
  TempSize: DWORD;
  ReadCharByte: Byte;
begin
  j := 0;  //计数 和count 决定搜索几次停下来
           //s1 := 'Playerr'; //要搜索的字符串
  chartoFind := 1;
  //filepathname := '33.dat';
  // tempFileMemStream := TmemoryStream.Create;
  try
    begin
      if True then
      begin
        helpStringLength := length(s1);
        TempLength := tempFileMemStream.Size;
        TempSize := TempLength;
        chartoFind := 1;
        for i := 0 to TempLength - 1 do
        begin
          tempFileMemStream.Seek(i, sofrombeginning);
          tempFileMemStream.ReadBuffer(ReadCharByte, 1);
          if chr(ReadCharByte) = s1[chartoFind] then
          begin
            inc(chartoFind);
            if chartoFind >= Dword(helpStringLength) + 1 then
            //found the string
            begin
              j := j + 1;
              tempAddress := i - helpStringLength + 1;
              chartoFind := 1;
              if j = Count then
                break;
            end;
          end
          else
            chartoFind := 1;
        end;
        if j = 0 then
        begin
          ShowMessage('找不到关键字，请检查文件是否正确！');
          Result := 999999999;
          exit;
        end;
             //form1.edit1.Text := s1;
             //form1.edit2.Text := IntToStr(tempaddress);
      end;   //text String scan
    end;     ///try end;
  finally
    //FreeAndNil(tempFileMemStream);
  end;
  Result := tempAddress;
end;

function getaddress2(val1: Integer): Dword;
var
  i: Integer;
  getval: Word;
  templength, tempaddress: Dword;
begin
  templength := tempFileMemStream.Size;
  for i := 0 to templength - 1 do
  begin
    if i = templength - 2 then
    begin
      Result := 0;
      break;
    end;
    tempFileMemStream.Seek(i, sofrombeginning);
    tempFileMemStream.ReadBuffer(getval, 2);
    if getval = val1 then
    begin
      tempaddress := i - 4;
      Result := tempaddress;
      break;
    end;
    Result := 0;
  end;
end;

//仙术是否学全
procedure checkhost(nub: Integer);
var
  len, raiselen: Integer;
  i, j: Integer;
  endad, buf, xad: Dword;
  buf1: array of Dword;
begin
  buf := 0;
  firstad := getaddress('MGC_HOST_SER_FLAG', nub); //找到标记
  xad := firstad + 17;                            //定位仙术数量处
  case nub of
    1, 2, 3:
      midad := getaddress('BASICT', nub + 1) - 37;
    4:
      midad := getaddress('CameraAutoSeek', 0) - 33;
  end;
  endad := midad;
{  tempFileMemStream.Seek(xad, sofrombeginning);
  tempFileMemStream.ReadBuffer(buf, 4);2008.3.14日发现算法有误，这几句根本就多余
  //   midad:=0;
  if buf > 0 then  //假如有仙术了
  begin}
  for i := firstad to endad do
  begin
    tempFileMemStream.Seek(i, sofrombeginning);
    tempFileMemStream.ReadBuffer(buf, 4);
    if (buf >= 5500) and (buf <= 5631) then
    begin
      tempFileMemStream.Seek(i - 4, sofrombeginning);
      tempFileMemStream.ReadBuffer(buf, 4);
      if (buf > 0) and (buf < 26) then
      begin
        midad := i - 4; //定位在武技数量处。
        break;
      end;
    end;
  end;
{  end
  else
    midad := firstad + 57;//如果没有学仙术,算出武技地址，技能不详
 }
  //    if midad=0 then
  form1.Hostcheck.Enabled := True;
  Form1.hostcheck.Checked := False;
  //len:=tempFileMemStream.Size-midad;//后面的长度要复制到别处
  //如果startad加上417字节等于midad那就说明已添加了不再添加
  //让复选框失效
  if midad >= firstad + 417 then
  begin   //使选框失效
    form1.hostcheck.Enabled := False;
    exit;
  end;
  //f1.SaveToFile('33.dat.bak');
   { raiselen:=417-(midad-firstad);//文件要增加的长度
    f1.Seek(midad,sofrombeginning);
    f2.CopyFrom(f1,len);//复制后面要保留的部分
    firstad:=firstad+17;//准备定位到写入位置
    f1.Size:=f1.Size+raiselen;
    f1.Seek(firstad,sofrombeginning);
    for i:=0 to 99 do
      f1.WriteBuffer(buf1[i],4);
    f1.CopyFrom(f2,0);
    f1.SaveToFile('33.dat');
    showmessage('好了，现在学会了！');
  }  //    f2.SaveToFile('midad.dat');
end;

procedure edithost;
var
  i, len, raiselen: Integer;
  f2: TmemoryStream;
  buf1: array of Dword;
begin
  SetLength(buf1, 100);
  buf1[0] := 6;
  buf1[1] := 5000; //$1388
  buf1[2] := 5001; //$1389
  buf1[3] := 5002; //$138A
  buf1[4] := 5003; //$138B
  buf1[5] := 5004; //$138C
  buf1[6] := 5005; //$138D
  buf1[7] := 6;
  buf1[8] := 5000;
  buf1[9] := 5;
  buf1[10] := 5001;
  buf1[11] := 5;
  buf1[12] := 5002;
  buf1[13] := 6;
  buf1[14] := 5003;
  buf1[15] := 6;
  buf1[16] := 5004;
  buf1[17] := 7;
  buf1[18] := 5005;
  buf1[19] := 7;

  buf1[20] := 6;
  buf1[21] := 5006;
  buf1[22] := 5007;
  buf1[23] := 5008;
  buf1[24] := 5009;
  buf1[25] := 5010;
  buf1[26] := 5011;
  buf1[27] := 6;
  buf1[28] := 5006;
  buf1[29] := 5;
  buf1[30] := 5007;
  buf1[31] := 5;
  buf1[32] := 5008;
  buf1[33] := 6;
  buf1[34] := 5009;
  buf1[35] := 6;
  buf1[36] := 5010;
  buf1[37] := 7;
  buf1[38] := 5011;
  buf1[39] := 7;

  buf1[40] := 6;
  buf1[41] := 5012;
  buf1[42] := 5013;
  buf1[43] := 5014;
  buf1[44] := 5015;
  buf1[45] := 5016;
  buf1[46] := 5017;
  buf1[47] := 6;
  buf1[48] := 5012;
  buf1[49] := 5;
  buf1[50] := 5013;
  buf1[51] := 5;
  buf1[52] := 5014;
  buf1[53] := 6;
  buf1[54] := 5015;
  buf1[55] := 6;
  buf1[56] := 5016;
  buf1[57] := 7;
  buf1[58] := 5017;
  buf1[59] := 7;

  buf1[60] := 6;
  buf1[61] := 5018;
  buf1[62] := 5019;
  buf1[63] := 5020;
  buf1[64] := 5021;
  buf1[65] := 5022;
  buf1[66] := 5023;
  buf1[67] := 6;
  buf1[68] := 5018;
  buf1[69] := 5;
  buf1[70] := 5019;
  buf1[71] := 5;
  buf1[72] := 5020;
  buf1[73] := 6;
  buf1[74] := 5021;
  buf1[75] := 6;
  buf1[76] := 5022;
  buf1[77] := 7;
  buf1[78] := 5023;
  buf1[79] := 7;

  buf1[80] := 6;
  buf1[81] := 5024;
  buf1[82] := 5025;
  buf1[83] := 5026;
  buf1[84] := 5027;
  buf1[85] := 5028;
  buf1[86] := 5029;
  buf1[87] := 6;
  buf1[88] := 5024;
  buf1[89] := 5;
  buf1[90] := 5025;
  buf1[91] := 5;
  buf1[92] := 5026;
  buf1[93] := 6;
  buf1[94] := 5027;
  buf1[95] := 6;
  buf1[96] := 5028;
  buf1[97] := 7;
  buf1[98] := 5029;
  buf1[99] := 7;

  len := tempFileMemStream.Size - midad; //后面的长度要复制到别处
  raiselen := 417 - (midad - firstad);  //文件要增加的长度
  if raiselen = 0 then
    exit;
  f2 := TMemoryStream.Create;
  tempFileMemStream.Seek(midad, sofrombeginning);
  f2.CopyFrom(tempFileMemStream, len); //复制后面要保留的部分
  //   firstad:=firstad+17;//准备定位到写入位置
  tempFileMemStream.Size := tempFileMemStream.Size + raiselen;
  tempFileMemStream.Seek(firstad + 17, sofrombeginning);
  for i := 0 to 99 do
    tempFileMemStream.WriteBuffer(buf1[i], 4);
  tempFileMemStream.CopyFrom(f2, 0);
  FreeAndNil(f2);
  midad := firstad + 417; //以保证后面对特技的修改地址按新地址进行
end;

procedure getatk; //读取特技
var
  i: Integer;
  buf: Dword;
  showmsg: boolean;
begin
  showmsg := true;
  with form1 do
  begin
    for i := 0 to 24 do
    begin
      ck.Checked[i] := False;
      ck.ItemEnabled[i] := True;
    end;
    tempFileMemStream.Seek(midad, sofrombeginning);
    tempFileMemStream.ReadBuffer(buf, 4);
    conatk := buf;
    if conatk > 25 then
      conatk := 25;
    for i := 1 to conatk do
    begin
      tempFileMemStream.ReadBuffer(buf, 4);
      case buf of
        5500:
          ck.Checked[0] := True;
        5501:
          ck.Checked[1] := True;
        5502:
          ck.Checked[2] := True;
        5503:
          ck.Checked[3] := True;
        5504:
          ck.Checked[4] := True;
        5505:
          ck.Checked[5] := True;
        5506:
          ck.Checked[6] := True;
        5507:
          ck.Checked[7] := True;
        5508:
          ck.Checked[8] := True;
        5509:
          ck.Checked[9] := True;
        5510:
          ck.Checked[10] := True;
        5511:
          ck.Checked[11] := True;
        5512:
          ck.Checked[12] := True;
        5513:
          ck.Checked[13] := True;
        5514:
          ck.Checked[14] := True;
        5515:
          ck.Checked[15] := True;
        5516:
          ck.Checked[16] := True;
        5517:
          ck.Checked[17] := True;
        5518:
          ck.Checked[18] := True;
        5519:
          ck.Checked[19] := True;
        5520:
          ck.Checked[20] := True;
        5521:
          ck.Checked[21] := True;
        5522:
          ck.Checked[22] := True;
        5523:
          ck.Checked[23] := True;
        5524:
          ck.Checked[24] := True;

      else
        begin
          if showmsg then
          begin
            ShowMessage('武技识别数据超出范围!' + #13#10 + '' + #13#10 + '如果修改后发现错误请把备份文件*.dat.old改回*.dat恢复!');
            showmsg := false;
          end;
        end;
      end;
    end;
    for i := 0 to 24 do
      if not ck.Checked[i] then
        ck.ItemEnabled[i] := False;
  end; //end with do
end;

procedure setatk(numb: Integer);
var
  con, i: Integer;
begin
  con := 0;
  with form1 do
  begin
    for i := 0 to 24 do
      if ck.Checked[i] then
      begin
        //ck.ItemEnabled[i]:=true;
        con := con + 1;
      end;
    if con = numb then
      for i := 0 to 24 do
        if not ck.Checked[i] then
          ck.ItemEnabled[i] := False;
    if con < numb then
      for i := 0 to 24 do
        if not ck.Checked[i] then
          ck.ItemEnabled[i] := True;
  end; //end with do
end;

procedure getinteam(numb: Integer);
var
  buf, a1, a2: Dword;
begin
  inteamad := getaddress('InTeam', numb) + 6;
  tempFileMemStream.Seek(inteamad, sofrombeginning);
  tempFileMemStream.readbuffer(buf, 4);
  with form1 do
  begin
    case buf of
      0:
        inteam.Checked := False;
      1:
        inteam.Checked := True;
    else
      ShowMessage('Error!');
    end;
    atkallad := inteamad + 30;
    atkall2ad := atkallad + 31;
    tempFileMemStream.Seek(atkallad, sofrombeginning);
    tempFileMemStream.readbuffer(a1, 4);
    tempFileMemStream.Seek(atkall2ad, sofrombeginning);
    tempFileMemStream.readbuffer(a2, 4);
    if (a1 = 1) and (a2 = 1) then
      atkall.Checked := True
    else
      atkall.Checked := False;
  end;
end;

procedure editinteam; //修改在队伍中和全攻击
var
  inteam1, a1, a2: Integer;
begin
  with form1 do
  begin
    if inteam.Checked then
      inteam1 := 1
    else
      inteam1 := 0;
    if atkall.Checked then
    begin
      a1 := 1;
      a2 := 1;
    end
    else
    begin
      a1 := 0;
      a2 := 0;
    end;
  end; //withend
  tempFileMemStream.Seek(inteamad, sofrombeginning);
  tempFileMemStream.writebuffer(inteam1, 1);
  tempFileMemStream.Seek(atkallad, sofrombeginning);
  tempFileMemStream.writebuffer(a1, 1);
  tempFileMemStream.Seek(atkall2ad, sofrombeginning);
  tempFileMemStream.writebuffer(a2, 1);
end;

procedure editatk;  //修改特技
var
  i, j, con: Integer;
  buf: Dword;
begin
  j := 0; //计数
  tempFilememStream.Seek(midad, sofrombeginning);
  tempFileMemStream.ReadBuffer(buf, 4);
  con := buf;
  with form1 do
  begin
    for i := 0 to 24 do
      if ck.Checked[i] then
      begin
        buf := 5500 + i;
        tempFileMemStream.WriteBuffer(buf, 4);
        j := j + 1;
      end;
  end;
  if con <> j then
    ShowMessage('editatk Error!');
end;

//图谱全
procedure fulltupu;
var
  res: tresourcestream;
  i, j, len: Integer;
  buf, startad, stopad: dword;
  f2: TmemoryStream;
begin
  f2 := tmemorystream.Create;
  startad := getaddress('money', 1) + 13;
  stopad := getaddress('EC_SERIALIZE_FLAG', 1) - 8;
  res := tresourceStream.Create(hinstance, 'book', PChar('blueprint'));

  j := res.Size - (stopad - startad);
  if j < 0 then
  begin
    ShowMessage('本存档数据多于已知数据，图谱将不修改！');
    exit;
  end;
  len := tempFileMemStream.size - stopad;
  tempFileMemStream.Seek(stopad, sofrombeginning);
  f2.CopyFrom(tempFileMemStream, len);
  tempFileMemStream.Size := tempFileMemStream.Size + j;
  tempFileMemStream.Seek(startad, sofrombeginning);
  tempFileMemStream.CopyFrom(res, 0);
  tempFileMemStream.CopyFrom(f2, 0);
  //tempFileMemStream.SaveToFile('33.dat');
  res.Free;
  //tempFileMemStream.Free;
  f2.Free;
  ShowMessage('图谱全部拥有^_^');
end;

// 物品全部
procedure fullgoods;
var
  res: tresourcestream;
  i, j, len: Integer;
  buf, startad, stopad, tempad: dword;
  f2: TmemoryStream;
begin
  f2 := tmemorystream.Create;

  startad := getaddress('EC_SERIALIZE_FLAG', 0) + 17;
  tempFileMemStream.Seek(startad, sofrombeginning);
  tempFileMemStream.ReadBuffer(buf, 4);
  if buf = 0 then
    startad := startad + 4; //如果没装备
  tempad := startad;
  if buf > 0 then //如果有装备
  begin
    stopad := getaddress('P_MGR_SERIALIZE_FLAG', 1) - 8;
    //label5.caption:=inttostr(endad);
    for i := startad to stopad do
    begin
      tempFileMemStream.Seek(i, sofrombeginning);
      tempFileMemStream.ReadBuffer(buf, 4);
      if buf = 1244032 then
        startad := i + 152;
      //label4.Caption:=inttostr(firstad);
    end;
    if startad = tempad then//如果是繁体则会找不到$80 FB 12，则查找13
      for i := startad to stopad do
      begin
        tempFileMemStream.Seek(i, sofrombeginning);
        tempFileMemStream.ReadBuffer(buf, 4);
        if buf = 1309568 then
          startad := i + 152;
      end;
    //如果还是找不到呢?
  end;

  tempFileMemStream.Seek(startad, sofrombeginning);
  tempFileMemStream.ReadBuffer(buf, 4);
  if buf > 6 then
//  begin
    ShowMessage('修改物品。' + #13#10 + '如果修改后出现错误，请把备份文件*.dat.old改回*.dat恢复！');
//    exit;
//  end    //
//  else   //网友反映会出现上面的提示2008.3.14
  startad := startad + buf * 4 + 8;
  //label4.caption:=inttostr(firstad);
  stopad := startad + 4;  //结束地址赋初值，后面此地址会变
  for i := 1 to 5 do
  begin
    tempFileMemStream.Seek(stopad, sofrombeginning);
    tempFileMemStream.ReadBuffer(buf, 4);
    stopad := stopad + buf * 12 + 8;  //结束地址
    //label6.Caption:=inttostr(midad);
  end;

  res := tresourceStream.Create(hinstance, 'good', PChar('goods'));
  j := res.Size - (stopad - startad);
  if j < 0 then
  begin
    ShowMessage('本存档数据多于已知数据，物品不能修改！');
    exit;
  end;
  len := tempFileMemStream.size - stopad;
  tempFileMemStream.Seek(stopad, sofrombeginning);
  f2.CopyFrom(tempFileMemStream, len);
  tempFileMemStream.Size := tempFileMemStream.Size + j;
  tempFileMemStream.Seek(startad, sofrombeginning);
  tempFileMemStream.CopyFrom(res, 0);
  tempFileMemStream.CopyFrom(f2, 0);
  //tempFileMemStream.SaveToFile('33.dat');
  res.Free;
  //tempFileMemStream.Free;
  f2.Free;
  ShowMessage('物品全部拥有');
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  if filepathname = '' then
    n2.Click;
  if FilepathName = '' then
    exit;
  getvalue(getaddress('Playerr', 1));
  getvalue2(getaddress('BASICT', 1));
  checkhost(1);
  getatk;
  getinteam(1);
  player.Caption := '云天河';
end;

procedure TForm1.N2Click(Sender: TObject);
const
  Model = 'yyyy_mm_dd_hh_mmss'; { 设定时间格式 win7以前的系统/、-等同}
var
  T1, T3: string; //取当前时间做为文件名
begin
  if opendialog1.Execute then
  begin
    filepathname := opendialog1.FileName;
    label1.Caption := filepathname;
    if filepathname = '' then
      exit;
    tempFileMemStream := TmemoryStream.Create;
    tempFileMemStream.LoadFromFile(FilePathName);
    FindFirst(filepathname, faAnyFile, TP); //查找目标文件
    T3 := FormatDateTime(Model, Now);
    T1 := filepathname;
    T1 := AnsiReverseString(T1);
    T1 := Copy(T1, Pos('.', T1) + 1, Length(T1));
    T1 := AnsiReverseString(T1);
    T1 := T1 + '-' + T3 + '.dat';
    //tempFileMemStream.SaveToFile(filepathname+'.old');
    //用copyfile可以保存原文件的修改时间
    copyfile(PChar(filepathname), PChar(T1), False);
  end;
  if filepathname <> '' then
  begin
    TabSheet2Show(Sender);

  end;
//  button1.Click;
  led_exper.Text := '';
  led_level.Text := '';
  led_hp.Text := '';
  led_maxhp.Text := '';
  led_mp.Text := '';
  led_maxmp.Text := '';
  led_rage.text := '';
  led_wuli1.Text := '';
  led_fang1.Text := '';
  led_speed1.Text := '';
  led_lucky1.Text := '';
  led_will1.Text := '';
  button6.Caption := '自动修改';
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  if filepathname = '' then
    n2.Click;
  if FilepathName = '' then
    exit;
  getvalue(getaddress('Playerr', 2));
  getvalue2(getaddress('BASICT', 2));
  checkhost(2);
  getatk;
  getinteam(2);
  player.Caption := '韩菱纱';
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  if filepathname = '' then
    n2.Click;
  if FilepathName = '' then
    exit;
  getvalue(getaddress('Playerr', 3));
  getvalue2(getaddress('BASICT', 3));
  checkhost(3);
  getatk;
  getinteam(3);
  player.Caption := '柳梦璃';
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
  if filepathname = '' then
    n2.Click;
  if FilepathName = '' then
    exit;
  getvalue(getaddress('Playerr', 4));
  getvalue2(getaddress('BASICT', 4));
  checkhost(4);
  getatk;
  getinteam(4);
  player.Caption := '慕容紫英';
end;
//人物属性保存到内存流

procedure TForm1.Button5Click(Sender: TObject);
var
  level: DWord;
  exper: DWord;
  HP: DWord;
  MaxHP: DWord;
  Rage: DWord;
  MP: DWord;
  MaxMP: DWord;
  Wuli1: Dword;
  Fang1: Dword;
  Speed1: Dword;
  Lucky1: Dword;
  Will1: Dword;
begin
  if filepathname = '' then
    exit;
  if led_exper.Modified then
  begin//showmessage('modify');
    tempfilememstream.Seek(experad, sofrombeginning);
    exper := StrToInt(led_exper.Text);
    tempfileMemstream.WriteBuffer(exper, 4);
  end;
  if led_fang1.Modified then
  begin
    tempfilememstream.seek(fang1ad, sofrombeginning);
    Fang1 := StrToInt(led_fang1.Text);
    tempfilememstream.WriteBuffer(Fang1, 4);
  end;
  if led_hp.Modified then
  begin
    tempfilememstream.seek(hpad, sofrombeginning);
    HP := StrToInt(led_hp.Text);
    tempfilememstream.WriteBuffer(HP, 4);
  end;
  if led_level.Modified then
  begin
    tempfilememstream.seek(Levelad, sofrombeginning);
    level := StrToInt(led_level.Text);
    tempfilememstream.WriteBuffer(level, 4);
  end;
  if led_Lucky1.Modified then
  begin
    tempfilememstream.seek(Lucky1ad, sofrombeginning);
    Lucky1 := StrToInt(led_Lucky1.Text);
    tempfilememstream.WriteBuffer(Lucky1, 4);
  end;
  if led_maxhp.Modified then
  begin
    tempfilememstream.seek(maxhpad, sofrombeginning);
    MaxHP := StrToInt(led_maxhp.Text);
    tempfilememstream.WriteBuffer(MaxHP, 4);
    tempfilememstream.seek(maxhp1ad, sofrombeginning);
    tempfilememstream.WriteBuffer(MaxHP, 4);
  end;
  if led_maxmp.Modified then
  begin
    tempfilememstream.seek(maxmpad, sofrombeginning);
    MaxMP := StrToInt(led_maxmp.Text);
    tempfilememstream.WriteBuffer(MaxMP, 4);
    tempfilememstream.seek(maxmp1ad, sofrombeginning);
    tempfilememstream.WriteBuffer(MaxMP, 4);
  end;
  if led_mp.Modified then
  begin
    tempfilememstream.seek(mpad, sofrombeginning);
    MP := StrToInt(led_mp.Text);
    tempfilememstream.WriteBuffer(MP, 4);
  end;
  if led_rage.Modified then
  begin
    tempfilememstream.seek(ragead, sofrombeginning);
    Rage := StrToInt(led_rage.Text);
    tempfilememstream.WriteBuffer(Rage, 4);
  end;
  if led_speed1.Modified then
  begin
    tempfilememstream.seek(speed1ad, sofrombeginning);
    Speed1 := StrToInt(led_speed1.Text);
    tempfilememstream.WriteBuffer(Speed1, 4);
  end;
  if led_will1.Modified then
  begin
    tempfilememstream.seek(will1ad, sofrombeginning);
    Will1 := StrToInt(led_will1.Text);
    tempfilememstream.WriteBuffer(Will1, 4);
  end;
  if led_wuli1.Modified then
  begin
    tempfilememstream.seek(wuli1ad, sofrombeginning);
    Wuli1 := StrToInt(led_wuli1.Text);
    tempfilememstream.WriteBuffer(Wuli1, 4);
  end;
  if hostcheck.Enabled and hostcheck.Checked then
    edithost; //仙术修改
  //button1.Click;//为
  editinteam;
  editatk; //修改个人特技
end;

procedure TForm1.N3Click(Sender: TObject);
begin
  if filepathname = '' then
  begin
    ShowMessage('没有打开文档');
    exit;
  end;
  tempfileMemstream.SaveToFile(filepathname);
  showmessage('文件已保存！');
  FindFirst(filepathname, faAnyFile, TP);
  //FreeAndNil(tempFileMemStream);
end;

procedure TForm1.TabSheet2Show(Sender: TObject);
var
//  tianxiang, huisheng, zhiningwan, hanlingguo, touguding: Integer;
//  zhentianlie: Integer;
  money: Integer;
  i: Integer;
  getval: Word;
  templength, tempaddress: Dword;
begin
   {tianxiang:=0;   //0x0BD7
   huisheng:=0;    //0x0BD6
   zhiningwan:=0;  //0X0bC3
   hanlingguo:=0;  //0X0BBD
   touguding:=0;   //0X0C0F
   zhentianlie:=0;   //0X0C10
   if filepathname='' then exit;
   LED1.Enabled:=False;
   LED2.Enabled:=False;
   LED3.Enabled:=False;
   LED4.Enabled:=False;
   LED5.Enabled:=False;Led6.Enabled:=false; }
  moneyad := getaddress('money', 1) + 5;
  if moneyad < 1000000004 then
  begin
    tempFileMemstream.Seek(moneyad, sofrombeginning);
    tempFileMemStream.ReadBuffer(money, 4);
    Lmoney.Text := IntToStr(money);
  end;
  bitbtn3.Enabled := True;
  { tianxiangad:=getaddress2(3031);
   huishengad:=getaddress2(3030);
   zhiningwanad:=getaddress2(3011);
   hanlingguoad:=getaddress2(3005);
   tougudingad:=getaddress2(3087);
   zhentianliead:=getaddress2(3088);
   if tianxiangad<>0 then begin
   tempFileMemstream.Seek(tianxiangad,sofrombeginning);
   tempFileMemStream.ReadBuffer(tianxiang,2);
   led1.Enabled:=true;
   LED1.Text:=inttostr(tianxiang); end;
   if huishengad<>0 then begin
   tempFileMemstream.Seek(huishengad,sofrombeginning);
   tempFileMemStream.ReadBuffer(huisheng,2);
   led2.Enabled:=true;
   LED2.Text:=inttostr(huisheng); end;
   if zhiningwanad<>0 then begin
   tempFileMemstream.Seek(zhiningwanad,sofrombeginning);
   tempFileMemStream.ReadBuffer(zhiningwan,2);
   Led3.Enabled:=true;
   LED3.Text:=inttostr(zhiningwan);end;
   if hanlingguoad<>0 then begin
   tempFileMemstream.Seek(hanlingguoad,sofrombeginning);
   tempFileMemStream.ReadBuffer(hanlingguo,2);
   Led4.Enabled:=true;
   LED4.Text:=inttostr(hanlingguo);end;
   if tougudingad<>0 then begin
   tempFileMemstream.Seek(tougudingad,sofrombeginning);
   tempFileMemStream.ReadBuffer(touguding,2);
   Led5.Enabled:=true;
   LED5.Text:=inttostr(touguding);end;
   if zhentianliead<>0 then begin
   tempFileMemstream.Seek(zhentianliead,sofrombeginning);
   tempFileMemStream.ReadBuffer(zhentianlie,2);
   Led6.Enabled:=true;
   LED6.Text:=inttostr(zhentianlie);end;
   }
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
var
  tianxiang, huisheng, zhiningwan, hanlingguo, touguding: Integer;
  zhentianlie: Integer;
  money: dword;
begin
  If Filepathname = '' Then
    Exit;
  { if led1.Modified then
   begin
   tempFileMemstream.Seek(tianxiangad,sofrombeginning);
   tianxiang:=strtoint(led1.Text);
   tempFileMemStream.WriteBuffer(tianxiang,2);
   end;
   if led2.Modified then
   begin
   tempFileMemstream.Seek(huishengad,sofrombeginning);
   huisheng:=strtoint(led2.Text);
   tempFileMemStream.WriteBuffer(huisheng,2);
   end;
   if led3.Modified then
   begin
   tempFileMemstream.Seek(zhiningwanad,sofrombeginning);
   zhiningwan:=strtoint(led3.Text);
   tempFileMemStream.WriteBuffer(zhiningwan,2);
   end;
   if led4.Modified then
   begin
   tempFileMemstream.Seek(hanlingguoad,sofrombeginning);
   hanlingguo:=strtoint(led4.Text);
   tempFileMemStream.WriteBuffer(hanlingguo,2);
   end;
   if led5.Modified then
   begin
   tempFileMemstream.Seek(tougudingad,sofrombeginning);
   touguding:=strtoint(led5.Text);
   tempFileMemStream.WriteBuffer(touguding,2);
   end;
   }
  { if led6.Modified then
   begin
   tempFileMemstream.Seek(zhentianliead,sofrombeginning);
   zhentianlie:=strtoint(led6.Text);
   tempFileMemStream.WriteBuffer(zhentianlie,2);
   end; }
  if wupin.Checked then
    fullgoods;
  if tupu.Checked then
    fulltupu;
  button1.Click;
  if (moneyad = 0) or (moneyad > 999999999) then
    exit; //一定要用括号
  if LMoney.Modified then
  begin
    tempFileMemstream.Seek(moneyad, sofrombeginning);
    money := StrToInt(Lmoney.Text);
    tempFileMemStream.WriteBuffer(money, 4);
  end;
end;

procedure automody; //修改物品数量
var
  i: Integer;
  getval, buf: Dword;
  templength, tempaddress: Dword;
  startad, stopad: dword;
begin
  templength := getaddress('BASICT', 1);
  startad := getaddress('EC_SERIALIZE_FLAG', 0) + 17;
  tempFileMemStream.Seek(startad, sofrombeginning);
  tempFileMemStream.ReadBuffer(buf, 4);
  if buf = 0 then
    startad := startad + 4; //如果没有装备物品将从这里开始
  stopad := getaddress('P_MGR_SERIALIZE_FLAG', 1) - 8; //物品结束地址
  if buf > 0 then//如果有装备
  begin
    //label5.caption:=inttostr(endad);
    for i := startad to stopad do//查找装备标志
    begin
      tempFileMemStream.Seek(i, sofrombeginning);
      tempFileMemStream.ReadBuffer(buf, 4);
      if buf = 1244032 then
        startad := i + 152; //找到标志，计算物品开始地址
      //label4.Caption:=inttostr(firstad);
    end;
  end;

  tempFileMemStream.Seek(startad, sofrombeginning);
  tempFileMemStream.ReadBuffer(buf, 4);
{  if buf > 6 then
  begin
    ShowMessage('貌似非正常文件！');

  end;
}
  for i := startad to stopad do
  begin
    if i = templength - 2 then
    begin
      break;
    end;
    tempFileMemStream.Seek(i, sofrombeginning);
    tempFileMemStream.ReadBuffer(getval, 4);
    if (getval > 2999) and (getval < 3328) then //找到物品代号
    begin
      //不修改剧情物品
      if (getval = 3103) or (getval = 3245) or (getval = 3246) then
        break;
      //定位物品数量地址
      tempaddress := i - 4;
      tempfilememstream.Seek(tempaddress, sofrombeginning);
      tempfilememstream.ReadBuffer(getval, 4); //取物品数量
      if getval < 255 then
      begin
        getval := getval or 240;
        //             getval:=255;
        tempfilememstream.Seek(tempaddress, sofrombeginning);
        tempfilememstream.WriteBuffer(getval, 4);
      end;
    end;
  end;
end;

procedure TForm1.BitBtn3Click(Sender: TObject);
begin
  if Filepathname = '' then
    exit;
  automody;
  bitbtn3.Enabled := False;
  //tabsheet2Show(sender);
end;

procedure TForm1.ckClickCheck(Sender: TObject);
begin
  setatk(conatk);
end;

procedure TForm1.N4Click(Sender: TObject);
begin
  FreeAndNil(tempFileMemStream);
  label1.Caption := Filepathname + '文件已关闭！';
  Filepathname := '';
//  filepathname := '';
  //label1.Caption:='文件已关闭';
end;

procedure TForm1.N6Click(Sender: TObject);
begin
  AboutBox.Show;
end;

procedure TForm1.Button6Click(Sender: TObject);
var
  money: Dword; //钱
  hpmp: dword; //精神
  gfyl: dword; //攻防运灵
  rage: dword; //气
  exp: dword;  //经验
  level: dword; //等级
  i, it: Integer;
begin
  if Filepathname = '' then
    exit;
  money := 99999999;
  hpmp := 999999;
  gfyl := 9999;
  rage := 100;
  exp := 2000000;
  level := 90;
  it := 1;
  tempfilememstream.Seek(moneyad, sofrombeginning);
  tempfilememstream.WriteBuffer(money, 4);
  automody;
  for i := 1 to 4 do
  begin
    getvalue(getaddress('Playerr', i));
    getvalue2(getaddress('BASICT', i));
    getinteam(i);
    tempfilememstream.Seek(inteamad, sofrombeginning);
    tempfilememstream.WriteBuffer(it, 1);
    tempfilememstream.Seek(hpad, sofrombeginning);
    tempfilememstream.WriteBuffer(hpmp, 4);
    tempfilememstream.Seek(mpad, sofrombeginning);
    tempfilememstream.WriteBuffer(hpmp, 4);
    tempfilememstream.Seek(maxhpad, sofrombeginning);
    tempfilememstream.WriteBuffer(hpmp, 4);
    tempfilememstream.Seek(maxmpad, sofrombeginning);
    tempfilememstream.WriteBuffer(hpmp, 4);
    tempfilememstream.Seek(maxhp1ad, sofrombeginning);
    tempfilememstream.WriteBuffer(hpmp, 4);
    tempfilememstream.Seek(maxmp1ad, sofrombeginning);
    tempfilememstream.WriteBuffer(hpmp, 4);
    tempfilememstream.Seek(wuli1ad, sofrombeginning);
    tempfilememstream.WriteBuffer(gfyl, 4);
    tempfilememstream.Seek(fang1ad, sofrombeginning);
    tempfilememstream.WriteBuffer(gfyl, 4);
    tempfilememstream.Seek(lucky1ad, sofrombeginning);
    tempfilememstream.WriteBuffer(gfyl, 4);
    tempfilememstream.Seek(will1ad, sofrombeginning);
    tempfilememstream.WriteBuffer(gfyl, 4);
    tempfilememstream.Seek(ragead, sofrombeginning);
    tempfilememstream.WriteBuffer(rage, 4);
    tempfilememstream.Seek(experad, sofrombeginning);
    tempfilememstream.WriteBuffer(exp, 4);
    tempfilememstream.Seek(levelad, sofrombeginning);
    tempfilememstream.WriteBuffer(level, 4);
  end;

  lmoney.Text := IntToStr(money);
  led_level.Text := IntToStr(level);
  led_exper.Text := IntToStr(exp);
  led_hp.Text := '999999';
  led_maxhp.Text := '999999';
  led_rage.Text := '100';
  led_mp.Text := '999999';
  led_maxmp.Text := '999999';
  led_wuli1.Text := '9999';
  led_fang1.Text := '9999';
  led_lucky1.Text := '9999';
  led_will1.Text := '9999';
  led_speed1.Text := '';
  inteam.Checked := True;
  button6.Caption := 'OK!修改好了。^_^';
  button1.Click;
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
  n3.Click;
  if Filepathname = '' then
    exit;
{
  FreeAndNil(tempFileMemStream);
  label1.Caption := filepathname + '文件已关闭！';
  filepathname := '';
  filepathname := '';
}
end;

procedure TForm1.FormPaint(Sender: TObject);
var
  ts: Tsearchrec;
begin
  if Filepathname = '' then
    exit;
//  showmessage('画窗体事件');
  if fileexists(filepathname) then
  begin
    FindFirst(Filepathname, faAnyFile, ts); //查找目标文件
    if covfiledate(tp.FindData.ftlastwritetime) <> covfiledate(ts.FindData.ftLastWriteTime) then
    begin
      if messagedlg('该存档文件已经被其他程序修改' + #13#10 + '是否要重新载入此文件？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        tempfilememstream.LoadFromFile(Filepathname);
        TabSheet2Show(Sender);
        button1.Click;
      end;
      FindFirst(Filepathname, faAnyFile, TP); //不管选是、否都把时间改为新的
    end;
  end
  else
  begin
     //showmessage('File is lost');//文件不存在也不提示
  end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FreeAndNil(tempFileMemStream);
  label1.Caption := Filepathname + '文件已关闭！';
  Filepathname := '';
end;

procedure TForm1.LED_LevelKeyPress(Sender: TObject; var Key: Char);
begin
  if not (Key in ['0'..'9', #8, #9]) then
    Key := #0;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  Application.HintColor := $00C08000;
  Screen.HintFont.Color := $00cccccc;
  Screen.HintFont.Name := '楷体_GB2312';
  Screen.HintFont.Size := 12;
  Application.HintHidePause := 10000;
end;

end.

