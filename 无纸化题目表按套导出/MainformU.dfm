object MainForm: TMainForm
  Left = 207
  Top = 126
  Width = 448
  Height = 563
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = #39064#30446#34920#25353#22871#23548#20986
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object RxRichEdit1: TRxRichEdit
    Left = 312
    Top = 32
    Width = 49
    Height = 9
    DrawEndPage = False
    Font.Charset = GB2312_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    Visible = False
  end
  object StartCreate: TButton
    Left = 128
    Top = 104
    Width = 121
    Height = 41
    Caption = #24320#22987#29983#25104#39064#30446
    TabOrder = 1
    OnClick = StartCreateClick
  end
  object StartJieXi: TButton
    Left = 128
    Top = 180
    Width = 121
    Height = 41
    Caption = #29983#25104#31572#26696#19982#35299#26512
    TabOrder = 2
    OnClick = StartJieXiClick
  end
  object chk1: TCheckBox
    Left = 296
    Top = 96
    Width = 97
    Height = 17
    Caption = #23548#20986#32771#28857
    Checked = True
    State = cbChecked
    TabOrder = 3
  end
  object CreateTiJiexi: TButton
    Left = 128
    Top = 256
    Width = 121
    Height = 41
    Caption = #29983#25104' '#39064#30446#21644#35299#26512
    TabOrder = 4
    OnClick = CreateTiJiexiClick
  end
end
