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
uses DataModul; {����ģ��}
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
  AQ.SQL.Add('select th  from ��Ŀ�� group by TH order by TH');
  AQ.open;
  num := AQ.recordcount;
  Setlength(TH, num + 1);//�趨����Ԫ�ظ���
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
      aq.sql.add('select * from ��Ŀ�� where th=' + IntToStr(TH[i]) +
        ' order by xh ');
      aq.open;
      count := aq.recordcount;
      rxrichedit1.Paragraph.Alignment := pacenter;
      rxrichedit1.SelAttributes.Color := clblue;
      rxrichedit1.SelAttributes.Size := 16;
      RXrichedit1.Lines.Append('��' + IntToStr(TH[i]) + '����');
      rxrichedit1.Paragraph.Alignment := paleftjustify;
      rxrichedit1.SelAttributes.Color := clblue;
      for j := 1 to Count do
      begin
        //  RXrichedit1.Lines.Append('(��'+inttostr(i)+'�����'+aq.fieldbyname('xh').asstring+'�⣩'+aq.fieldbyname('���').asstring);
        RxRichEdit1.SelAttributes.Color:=clRed;
        RXrichedit1.Lines.Append('(' + aq.fieldbyname('XH').AsString + ')');
        if chk1.Checked then
        begin
        rxrichedit1.SelAttributes.Color := clblue;
        RXrichedit1.Lines.Append('���㣺'+aq.fieldbyname('����1').asstring);
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
    rxrichedit1.Lines.SaveToFile('��Ŀ������.rtf');
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
  AQ.SQL.Add('select th  from ��Ŀ�� group by TH order by TH');
  AQ.open;
  num := AQ.recordcount;
  Setlength(TH, num + 1);//�趨����Ԫ�ظ���
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
      aq.sql.add('select * from ��Ŀ�� where th=' + IntToStr(TH[i]) +
        ' order by xh ');
      aq.open;
      count := aq.recordcount;
      rxrichedit1.Paragraph.Alignment := pacenter;
      rxrichedit1.SelAttributes.Color := clblue;
      rxrichedit1.SelAttributes.Size := 16;
      RXrichedit1.Lines.Append('��' + IntToStr(TH[i]) + '����');
      rxrichedit1.Paragraph.Alignment := paleftjustify;
      rxrichedit1.SelAttributes.Color := clblue;
      for j := 1 to Count do
      begin
        //  RXrichedit1.Lines.Append('(��'+inttostr(i)+'�����'+aq.fieldbyname('xh').asstring+'�⣩'+aq.fieldbyname('���').asstring);
        RxRichEdit1.SelAttributes.Color:=clRed;
        RXrichedit1.Lines.Append('(' + aq.fieldbyname('XH').AsString + ')');
        RXrichedit1.Lines.Append('�𰸣�'+aq.fieldbyname('DA').asstring);

        RXrichedit1.Lines.Append(aq.fieldbyname('JX').asstring);

        aq.Next;
      end;

    end;
    FreeandNil(aq);
    rxrichedit1.Lines.SaveToFile('��Ŀ������𰸺ͽ���.rtf');
  end;
end;
procedure TMainForm.StartCreateClick(Sender: TObject);
begin
  StartCreate.Enabled:=False;
  CreateTaoTi;
  StartCreate.Caption:='������Ŀ��ɣ���';
end;

procedure TMainForm.StartJieXiClick(Sender: TObject);
begin
  StartJieXi.Enabled:=False;
  CreateJieXi;
  StartJieXi.Caption:='�����𰸽�����ɣ���';
end;

//������Ŀ�ͽ���

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
  AQ.SQL.Add('select th  from ��Ŀ�� group by TH order by TH');
  AQ.open;
  num := AQ.recordcount;
  Setlength(TH, num + 1);//�趨����Ԫ�ظ���
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
      aq.sql.add('select * from ��Ŀ�� where th=' + IntToStr(TH[i]) +
        ' order by xh ');
      aq.open;
      count := aq.recordcount;
      rxrichedit1.Paragraph.Alignment := pacenter;
      rxrichedit1.SelAttributes.Color := clblue;
      rxrichedit1.SelAttributes.Size := 16;
      RXrichedit1.Lines.Append('��' + IntToStr(TH[i]) + '����');
      rxrichedit1.Paragraph.Alignment := paleftjustify;
      rxrichedit1.SelAttributes.Color := clblue;
      for j := 1 to Count do
      begin
        //  RXrichedit1.Lines.Append('(��'+inttostr(i)+'�����'+aq.fieldbyname('xh').asstring+'�⣩'+aq.fieldbyname('���').asstring);
        RxRichEdit1.SelAttributes.Color:=clRed;
        RXrichedit1.Lines.Append('(' + aq.fieldbyname('XH').AsString + ')');
        if chk1.Checked then
        begin
        rxrichedit1.SelAttributes.Color := clblue;
        RXrichedit1.Lines.Append('���㣺'+aq.fieldbyname('����1').asstring);
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

        RXrichedit1.Lines.Append('�𰸣�'+aq.fieldbyname('DA').asstring);

        RXrichedit1.Lines.Append(aq.fieldbyname('JX').asstring);

        aq.Next;
      end;

    end;
    FreeandNil(aq);
    rxrichedit1.Lines.SaveToFile('��Ŀ������ͽ���.rtf');
  end;
end;



procedure TMainForm.CreateTiJiexiClick(Sender: TObject);
begin
  CreateTiJiexi.Enabled:=False;
  createtaotiJieXi;
  CreateTiJiexi.Caption:='���� ��Ŀ�ͽ�����ɣ�';
end;

end.

