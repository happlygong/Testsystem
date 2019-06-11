object Form1: TForm1
  Left = 206
  Top = 96
  Width = 701
  Height = 490
  Caption = #21382#24180#32771#28857#27604#20363
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  DesignSize = (
    693
    463)
  PixelsPerInch = 96
  TextHeight = 13
  object StringGrid1: TStringGrid
    Left = 0
    Top = 0
    Width = 710
    Height = 490
    Anchors = [akLeft, akTop, akRight, akBottom]
    ColCount = 2
    Ctl3D = False
    DefaultColWidth = 164
    DefaultRowHeight = 16
    FixedCols = 0
    RowCount = 3
    FixedRows = 2
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goRowSelect, goThumbTracking]
    ParentCtl3D = False
    ParentFont = False
    TabOrder = 0
    OnDrawCell = StringGrid1DrawCell
  end
  object btn1: TButton
    Left = 504
    Top = 216
    Width = 75
    Height = 25
    Caption = 'btn1'
    TabOrder = 1
    Visible = False
    OnClick = btn1Click
  end
  object tmr1: TTimer
    OnTimer = tmr1Timer
    Left = 200
    Top = 48
  end
end
