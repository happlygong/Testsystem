object Form1: TForm1
  Left = 178
  Top = 121
  Width = 247
  Height = 365
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = #33719#21462#32771#28857#20869#23481#20026#25991#20214
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  DesignSize = (
    239
    338)
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 8
    Top = 0
    Width = 217
    Height = 89
  end
  object Label1: TLabel
    Left = 39
    Top = 18
    Width = 182
    Height = 168
    Anchors = [akLeft, akTop, akRight, akBottom]
    AutoSize = False
    Font.Charset = GB2312_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = #26999#20307'_GB2312'
    Font.Style = []
    ParentFont = False
    WordWrap = True
  end
  object grp1: TGroupBox
    Left = 8
    Top = 1
    Width = 217
    Height = 89
    Caption = #22871#21495
    TabOrder = 2
    Visible = False
    object chk1: TCheckBox
      Left = 31
      Top = 19
      Width = 50
      Height = 17
      Caption = '1'
      TabOrder = 0
    end
    object chk2: TCheckBox
      Left = 31
      Top = 42
      Width = 50
      Height = 17
      Caption = '2'
      TabOrder = 1
    end
    object chk3: TCheckBox
      Left = 31
      Top = 65
      Width = 50
      Height = 17
      Caption = '3'
      TabOrder = 2
    end
    object chk4: TCheckBox
      Left = 100
      Top = 18
      Width = 50
      Height = 17
      Caption = '4'
      TabOrder = 3
    end
    object chk5: TCheckBox
      Left = 100
      Top = 41
      Width = 50
      Height = 17
      Caption = '5'
      TabOrder = 4
    end
    object chk6: TCheckBox
      Left = 100
      Top = 66
      Width = 50
      Height = 17
      Caption = '6'
      TabOrder = 5
    end
    object chk7: TCheckBox
      Left = 163
      Top = 16
      Width = 50
      Height = 17
      Caption = '7'
      TabOrder = 6
    end
    object chk8: TCheckBox
      Left = 163
      Top = 40
      Width = 50
      Height = 17
      Caption = '8'
      TabOrder = 7
    end
    object chk9: TCheckBox
      Left = 164
      Top = 66
      Width = 50
      Height = 17
      Caption = '9'
      TabOrder = 8
    end
  end
  object btn1: TButton
    Left = 4
    Top = 140
    Width = 116
    Height = 25
    Hint = #25226#25152#26377#25968#25454#25353#31456#33410#39034#24207#25918#21040#19968#20010'Excel'#34920#20013
    Caption = #24320#22987'('#29983#25104#21333#20010#34920')'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 0
    Visible = False
    OnClick = btn1Click
  end
  object Button1: TButton
    Left = 4
    Top = 172
    Width = 116
    Height = 25
    Hint = #25353#19968#33267#22235#32423#26631#39064#25226#32467#26524#25968#25454#20998#21035#23384#25918#21040#22235#20010'Excel'#24037#20316#34920#20013
    Caption = #25353#26631#39064#32423#21035#20998#20026'4'#20010#34920
    ParentShowHint = False
    ShowHint = True
    TabOrder = 1
    Visible = False
    OnClick = Button1Click
  end
  object btn3: TButton
    Left = 122
    Top = 140
    Width = 110
    Height = 25
    Hint = 
      #25972#22871#39064#24635#20998#19981#36275'70'#25110'100'#20998#26102#65292#13#10#24456#21487#33021#26159#26377#19982#30446#24405#19981#19968#26679#30340#32771#13#10#28857#24405#20837#20102#25968#25454#24211#20013#65292#21333#20987#27492#25353#13#10#38062#21487#20197#25226#23427#25214#20986#26469#12290'|'#25972#22871#39064#24635#20998#19981 +
      #36275'70'#25110'100'#20998#26102#65292#24456#21487#33021'#13'#26159#26377#19982#30446#24405#19981#19968#26679#30340#32771#28857#24405#20837#20102#25968#25454#24211#20013#65292#21333#20987#20123#25353#38062#21487#20197#25226#23427#25214#20986#26469#12290
    Caption = #25214#20986#19981#19968#33268#30340#32771#28857
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 3
    Visible = False
    OnClick = btn3Click
  end
  object btn4: TButton
    Left = 122
    Top = 172
    Width = 110
    Height = 25
    Hint = #20351#19978#38754#37027#20010#25353#38062#21487#29992
    Caption = #28608#27963#25214#19981#19968#33268#32771#28857
    ParentShowHint = False
    ShowHint = True
    TabOrder = 4
    Visible = False
    OnClick = btn4Click
  end
  object btn5: TButton
    Left = 145
    Top = 204
    Width = 65
    Height = 25
    Hint = #23558#27169#25311#39064#27604#20363#30452#25509#20889#20837#21508#32423#32771#28857#25968#25454#24211#20013
    Caption = #32467#26524#30452#25509#20889#20837#25968#25454#24211
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
    Visible = False
    OnClick = btn5Click
  end
  object btn6: TButton
    Left = 212
    Top = 207
    Width = 20
    Height = 20
    Hint = #32972#26223#38899#20048#24320#20851
    Caption = 'V'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 6
    OnClick = btn6Click
  end
  object btn7: TButton
    Left = 4
    Top = 204
    Width = 68
    Height = 25
    Hint = #29983#25104#19968#20010#21333#34920#65292#25968#25454#38500'100'
    Caption = #38500'100'#30340#21333#34920
    ParentShowHint = False
    ShowHint = True
    TabOrder = 7
    Visible = False
    OnClick = btn7Click
  end
  object btn8: TButton
    Left = 76
    Top = 204
    Width = 65
    Height = 25
    Hint = #30495#39064#21508#32423#27604#20363#30452#25509#20889#20837#25968#25454#24211
    Caption = #20889#20837#30495#39064#27604#20363
    ModalResult = 1
    ParentShowHint = False
    ShowHint = True
    TabOrder = 8
    Visible = False
    OnClick = btn8Click
  end
  object rg1: TRadioGroup
    Left = 8
    Top = 94
    Width = 217
    Height = 33
    Hint = #36873#25321#25968#25454#20998#26512#30340#26469#28304#26159'"'#39064#30446#34920'"'#65292#36824#26159'"'#27169#25311#39064#34920'".'
    Columns = 2
    ItemIndex = 1
    Items.Strings = (
      #39064#30446#34920'('#30495#39064')'
      #27169#25311#39064#34920)
    ParentShowHint = False
    ShowHint = True
    TabOrder = 9
    Visible = False
  end
  object btn2: TButton
    Left = 8
    Top = 240
    Width = 75
    Height = 25
    Caption = #20844#20849#22522#30784'1'
    TabOrder = 10
    Visible = False
    OnClick = btn2Click
  end
  object btn9: TButton
    Left = 120
    Top = 240
    Width = 75
    Height = 25
    Caption = #20844#20849#22522#30784'4 '
    TabOrder = 11
    Visible = False
    OnClick = btn9Click
  end
  object btn10: TButton
    Left = 24
    Top = 288
    Width = 177
    Height = 25
    Caption = #33719#21462#32771#28857#20869#23481#21040#25991#20214
    TabOrder = 12
    OnClick = btn10Click
  end
  object edt1: TRxRichEdit
    Left = 88
    Top = 256
    Width = 0
    Height = 0
    DrawEndPage = False
    Font.Charset = GB2312_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = #21326#25991#26032#39759
    Font.Style = []
    ParentFont = False
    TabOrder = 13
  end
  object Timer1: TTimer
    Interval = 100
    OnTimer = Timer1Timer
    Left = 14
    Top = 132
  end
end
