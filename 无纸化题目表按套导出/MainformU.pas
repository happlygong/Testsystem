unit MainformU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, RxRichEd, ADODB;

type
  TMainForm = class(TForm)
    RxRichEdit1: TRxRichEdit;
    StartCreate: TButton;
    StartJieXi: TButton;
    chk1: TCheckBox;
    CreateTiJiexi: TButton;
    procedure StartCreateClick(Sender: TObject);
    procedure StartJieXiClick(Sender: TObject);
    procedure CreateTiJiexiClick(Sender: TObject);
    //  procedure CreateTaoTi();
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation
uses DataModul; {引用模块}
{$R *.dfm}

procedure CreateTaoTi();
var
  i, j: Integer;
  num, Count: Integer;
  TH:array of Integer;
  aq: TADOQuery;
begin
  aq := TADOQuery.Create(nil);
  aq.connection := datamodul.datamd.adoconn;
  AQ.sql.Clear;
  AQ.SQL.Add('select th  from 题目表 group by TH order by TH');
  AQ.open;
  num := AQ.recordcount;
  Setlength(TH, num + 1);//设定数据元素个数
  for i:=1 to num do
  begin
    th[i]:=aq.fieldbyname('TH').AsInteger;
    aq.Next;
  end;
  with MainForm do
  begin
    RxRichEdit1.Lines.Clear;
    for i := 1 to num do
    begin
      aq.sql.clear;
      aq.sql.add('select * from 题目表 where th=' + IntToStr(TH[i]) +
        ' order by xh ');
      aq.open;
      count := aq.recordcount;
      rxrichedit1.Paragraph.Alignment := pacenter;
      rxrichedit1.SelAttributes.Color := clblue;
      rxrichedit1.SelAttributes.Size := 16;
      RXrichedit1.Lines.Append('第' + IntToStr(TH[i]) + '套题');
      rxrichedit1.Paragraph.Alignment := paleftjustify;
      rxrichedit1.SelAttributes.Color := clblue;
      for j := 1 to Count do
      begin
        //  RXrichedit1.Lines.Append('(第'+inttostr(i)+'套题第'+aq.fieldbyname('xh').asstring+'题）'+aq.fieldbyname('编号').asstring);
        RxRichEdit1.SelAttributes.Color:=clRed;
        RXrichedit1.Lines.Append('(' + aq.fieldbyname('XH').AsString + ')');
        if chk1.Checked then
        begin
        rxrichedit1.SelAttributes.Color := clblue;
        RXrichedit1.Lines.Append('考点：'+aq.fieldbyname('考点1').asstring);
        end;
        RXrichedit1.Lines.Append(aq.fieldbyname('TM').asstring);
        RxRichEdit1.SelAttributes.Color:=clGreen;
        RXrichedit1.Lines.Append('(A)');
        RXrichedit1.Lines.Append(aq.fieldbyname('OptionA').asstring);
        RxRichEdit1.SelAttributes.Color:=clGreen;
        RXrichedit1.Lines.Append('(B)');
        RXrichedit1.Lines.Append(aq.fieldbyname('OptionB').asstring);
        RxRichEdit1.SelAttributes.Color:=clGreen;
        RXrichedit1.Lines.Append('(C)');
        RXrichedit1.Lines.Append(aq.fieldbyname('OptionC').asstring);
        RxRichEdit1.SelAttributes.Color:=clGreen;
        RXrichedit1.Lines.Append('(D)');
        RXrichedit1.Lines.Append(aq.fieldbyname('OptionD').asstring);
        aq.Next;
      end;

    end;
    FreeandNil(aq);
    rxrichedit1.Lines.SaveToFile('题目表套题.rtf');
  end;
end;
Procedure CreateJieXi();
 var
  i, j: Integer;
  num, Count: Integer;
  TH:array of Integer;
  aq: TADOQuery;
begin
  aq := TADOQuery.Create(nil);
  aq.connection := datamodul.datamd.adoconn;
  AQ.sql.Clear;
  AQ.SQL.Add('select th  from 题目表 group by TH order by TH');
  AQ.open;
  num := AQ.recordcount;
  Setlength(TH, num + 1);//设定数据元素个数
  for i:=1 to num do
  begin
    th[i]:=aq.fieldbyname('TH').AsInteger;
    aq.Next;
  end;
  with MainForm do
  begin
    RxRichEdit1.Lines.Clear;
    for i := 1 to num do
    begin
      aq.sql.clear;
      aq.sql.add('select * from 题目表 where th=' + IntToStr(TH[i]) +
        ' order by xh ');
      aq.open;
      count := aq.recordcount;
      rxrichedit1.Paragraph.Alignment := pacenter;
      rxrichedit1.SelAttributes.Color := clblue;
      rxrichedit1.SelAttributes.Size := 16;
      RXrichedit1.Lines.Append('第' + IntToStr(TH[i]) + '套题');
      rxrichedit1.Paragraph.Alignment := paleftjustify;
      rxrichedit1.SelAttributes.Color := clblue;
      for j := 1 to Count do
      begin
        //  RXrichedit1.Lines.Append('(第'+inttostr(i)+'套题第'+aq.fieldbyname('xh').asstring+'题）'+aq.fieldbyname('编号').asstring);
        RxRichEdit1.SelAttributes.Color:=clRed;
        RXrichedit1.Lines.Append('(' + aq.fieldbyname('XH').AsString + ')');
        RXrichedit1.Lines.Append('答案：'+aq.fieldbyname('DA').asstring);

        RXrichedit1.Lines.Append(aq.fieldbyname('JX').asstring);

        aq.Next;
      end;

    end;
    FreeandNil(aq);
    rxrichedit1.Lines.SaveToFile('题目表套题答案和解析.rtf');
  end;
end;
procedure TMainForm.StartCreateClick(Sender: TObject);
begin
  StartCreate.Enabled:=False;
  CreateTaoTi;
  StartCreate.Caption:='导出题目完成！！';
end;

procedure TMainForm.StartJieXiClick(Sender: TObject);
begin
  StartJieXi.Enabled:=False;
  CreateJieXi;
  StartJieXi.Caption:='导出答案解析完成！！';
end;

//导出题目和解析

procedure CreateTaoTiJieXi();
var
  i, j: Integer;
  num, Count: Integer;
  TH:array of Integer;
  aq: TADOQuery;
begin
  aq := TADOQuery.Create(nil);
  aq.connection := datamodul.datamd.adoconn;
  AQ.sql.Clear;
  AQ.SQL.Add('select th  from 题目表 group by TH order by TH');
  AQ.open;
  num := AQ.recordcount;
  Setlength(TH, num + 1);//设定数据元素个数
  for i:=1 to num do
  begin
    th[i]:=aq.fieldbyname('TH').AsInteger;
    aq.Next;
  end;
  with MainForm do
  begin
    RxRichEdit1.Lines.Clear;
    for i := 1 to num do
    begin
      aq.sql.clear;
      aq.sql.add('select * from 题目表 where th=' + IntToStr(TH[i]) +
        ' order by xh ');
      aq.open;
      count := aq.recordcount;
      rxrichedit1.Paragraph.Alignment := pacenter;
      rxrichedit1.SelAttributes.Color := clblue;
      rxrichedit1.SelAttributes.Size := 16;
      RXrichedit1.Lines.Append('第' + IntToStr(TH[i]) + '套题');
      rxrichedit1.Paragraph.Alignment := paleftjustify;
      rxrichedit1.SelAttributes.Color := clblue;
      for j := 1 to Count do
      begin
        //  RXrichedit1.Lines.Append('(第'+inttostr(i)+'套题第'+aq.fieldbyname('xh').asstring+'题）'+aq.fieldbyname('编号').asstring);
        RxRichEdit1.SelAttributes.Color:=clRed;
        RXrichedit1.Lines.Append('(' + aq.fieldbyname('XH').AsString + ')');
        if chk1.Checked then
        begin
        rxrichedit1.SelAttributes.Color := clblue;
        RXrichedit1.Lines.Append('考点：'+aq.fieldbyname('考点1').asstring);
        end;
        RXrichedit1.Lines.Append(aq.fieldbyname('TM').asstring);
        RxRichEdit1.SelAttributes.Color:=clGreen;
        RXrichedit1.Lines.Append('(A)');
        RXrichedit1.Lines.Append(aq.fieldbyname('OptionA').asstring);
        RxRichEdit1.SelAttributes.Color:=clGreen;
        RXrichedit1.Lines.Append('(B)');
        RXrichedit1.Lines.Append(aq.fieldbyname('OptionB').asstring);
        RxRichEdit1.SelAttributes.Color:=clGreen;
        RXrichedit1.Lines.Append('(C)');
        RXrichedit1.Lines.Append(aq.fieldbyname('OptionC').asstring);
        RxRichEdit1.SelAttributes.Color:=clGreen;
        RXrichedit1.Lines.Append('(D)');
        RXrichedit1.Lines.Append(aq.fieldbyname('OptionD').asstring);
        RxRichEdit1.SelAttributes.Color:=clRed;

        RXrichedit1.Lines.Append('答案：'+aq.fieldbyname('DA').asstring);

        RXrichedit1.Lines.Append(aq.fieldbyname('JX').asstring);

        aq.Next;
      end;

    end;
    FreeandNil(aq);
    rxrichedit1.Lines.SaveToFile('题目表套题和解析.rtf');
  end;
end;



procedure TMainForm.CreateTiJiexiClick(Sender: TObject);
begin
  CreateTiJiexi.Enabled:=False;
  createtaotiJieXi;
  CreateTiJiexi.Caption:='导出 题目和解析完成！';
end;

end.

