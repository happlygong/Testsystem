object Form1: TForm1
  Left = 192
  Top = 107
  Width = 378
  Height = 461
  Anchors = []
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 136
    Top = 56
    Width = 75
    Height = 25
    Caption = #26368#23567#21270#21040#25176#30424
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 136
    Top = 216
    Width = 75
    Height = 25
    Caption = #36864#20986
    TabOrder = 1
    OnClick = Button2Click
  end
  object PopupMenu1: TPopupMenu
    Left = 160
    Top = 128
    object N11: TMenuItem
      Caption = #24674#22797#31383#21475
      OnClick = N11Click
    end
    object n21: TMenuItem
      Caption = #36864#20986#31243#24207
      OnClick = n21Click
    end
  end
end
