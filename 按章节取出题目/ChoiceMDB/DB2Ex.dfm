object Form1: TForm1
  Left = 226
  Top = 135
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = #20998#26512#21508#22871#32771#28857#30334#20998#27604
  ClientHeight = 320
  ClientWidth = 237
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  ShowHint = True
  OnShow = FormShow
  DesignSize = (
    237
    320)
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
    Width = 180
    Height = 50
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
  object img1: TImage
    Left = 216
    Top = 110
    Width = 19
    Height = 12
    Picture.Data = {
      055449636F6E0000010001002020100000000000E80200001600000028000000
      2000000040000000010004000000000080020000000000000000000000000000
      0000000000000000000080000080000000808000800000008000800080800000
      C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000
      FFFFFF0000000000000000000000000000000000000000000000000000000000
      0000000000000000000000009999000000000000000000000000000999901000
      000000000003333000000099990110000000000000B033330000099990110000
      0000000000BB0333300000000110000000000000000BB0333300009999000000
      000000000000BB00000000000000000000000000000003333000000000000000
      0008000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000EEEE00000000000
      0000000000000000EEEE060000000000000000000000000EEEE0660000000000
      00000000000000EEEE0660000000000000000000000000000066000000000000
      000000000000000EEEE000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000080000000000000000000000000
      00000818000000000000000000000008000000800000000000000BB000000081
      80000000000000000000BB00000000080000000000000000000BB00000000000
      0000800000000000008000000000000000081800000000000000000000080000
      0000800000000000000000000081800000000000000000000000000000080000
      0080000000000000000000000000000008180000000000000000000000000000
      00800000FFFFFFFFFFFF0FFFFFFE07FFE1FC03FFC0F803FF807007FF80300FFF
      C0181FFFE01C3FFFF03FFFCFF87FFF8FFFFFFF0FFF87FE1FFF03FC3FFE03F87F
      FC03F0FFF803E1FFF807C3FFFC0F87FFFE1F0FFFFFFE1FFDFFFC3FF8FFF87EFD
      FFF0FC7FFFE1FEFFFFC3FFF7FFC7FFE3FFFFEFF7FFFFC7FFFFFFEFDFFFFFFF8F
      FFFFFFDF}
    OnClick = img1Click
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
    OnClick = Button1Click
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
      Left = 17
      Top = 19
      Width = 50
      Height = 17
      Caption = '1'
      TabOrder = 0
    end
    object chk2: TCheckBox
      Left = 17
      Top = 42
      Width = 50
      Height = 17
      Caption = '2'
      TabOrder = 1
    end
    object chk3: TCheckBox
      Left = 17
      Top = 65
      Width = 50
      Height = 17
      Caption = '3'
      TabOrder = 2
    end
    object chk4: TCheckBox
      Left = 55
      Top = 18
      Width = 50
      Height = 17
      Caption = '4'
      TabOrder = 3
    end
    object chk5: TCheckBox
      Left = 55
      Top = 41
      Width = 50
      Height = 17
      Caption = '5'
      TabOrder = 4
    end
    object chk6: TCheckBox
      Left = 55
      Top = 66
      Width = 50
      Height = 17
      Caption = '6'
      TabOrder = 5
    end
    object chk7: TCheckBox
      Left = 95
      Top = 16
      Width = 50
      Height = 17
      Caption = '7'
      TabOrder = 6
      Visible = False
    end
    object chk8: TCheckBox
      Left = 95
      Top = 40
      Width = 50
      Height = 17
      Caption = '8'
      TabOrder = 7
    end
    object chk9: TCheckBox
      Left = 96
      Top = 66
      Width = 50
      Height = 17
      Caption = '9'
      TabOrder = 8
    end
    object chk10: TCheckBox
      Left = 131
      Top = 17
      Width = 30
      Height = 17
      Caption = '10'
      TabOrder = 9
    end
    object edt1: TRxRichEdit
      Left = 8
      Top = 48
      Width = 1
      Height = 0
      DrawEndPage = False
      Font.Charset = GB2312_CHARSET
      Font.Color = clBlue
      Font.Height = -15
      Font.Name = #23435#20307
      Font.Style = []
      ParentFont = False
      TabOrder = 11
    end
    object chk11: TCheckBox
      Left = 132
      Top = 41
      Width = 30
      Height = 17
      Caption = '11'
      TabOrder = 10
    end
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
    OnClick = btn4Click
  end
  object btn5: TButton
    Left = 145
    Top = 203
    Width = 65
    Height = 25
    Hint = #23558#27169#25311#39064#27604#20363#30452#25509#20889#20837#21508#32423#32771#28857#25968#25454#24211#20013
    Caption = #32467#26524#30452#25509#20889#20837#25968#25454#24211
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
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
    Top = 203
    Width = 68
    Height = 25
    Hint = #29983#25104#19968#20010#21333#34920#65292#25968#25454#38500'100'
    Caption = #38500'100'#30340#21333#34920
    ParentShowHint = False
    ShowHint = True
    TabOrder = 7
    OnClick = btn7Click
  end
  object btn8: TButton
    Left = 76
    Top = 203
    Width = 65
    Height = 25
    Hint = #30495#39064#21508#32423#27604#20363#30452#25509#20889#20837#25968#25454#24211
    Caption = #20889#20837#30495#39064#27604#20363
    ModalResult = 1
    ParentShowHint = False
    ShowHint = True
    TabOrder = 8
    OnClick = btn8Click
  end
  object rg1: TRadioGroup
    Left = 8
    Top = 94
    Width = 217
    Height = 33
    Hint = #36873#25321#25968#25454#20998#26512#30340#26469#28304#26159'"'#39064#30446#34920'"'#65292#36824#26159'"'#27169#25311#39064#34920'".'
    Columns = 2
    ItemIndex = 0
    Items.Strings = (
      #39064#30446#34920'('#30495#39064')'
      #27169#25311#39064#34920)
    ParentShowHint = False
    ShowHint = True
    TabOrder = 9
    OnClick = rg1Click
  end
  object btn2: TButton
    Left = 8
    Top = 235
    Width = 89
    Height = 25
    Hint = #35745#31639#20844#20849#22522#30784#37096#20998#30340#20998#20540#27604#20363#36755#20986#21040#19968#20010#34920
    Caption = #20844#20849#22522#30784'1'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 10
    OnClick = btn2Click
  end
  object btn9: TButton
    Left = 104
    Top = 235
    Width = 121
    Height = 25
    Hint = #35745#31639#20844#20849#22522#30784#37096#20998#20998#20540#27604#20363#25353#26631#39064#36755#20986#21040'4'#20010#34920#20013
    Caption = #20844#20849#22522#30784'4 '
    ParentShowHint = False
    ShowHint = True
    TabOrder = 11
    OnClick = btn9Click
  end
  object btn10: TButton
    Left = 104
    Top = 264
    Width = 121
    Height = 25
    Hint = #21508#31456#33410#19979#30340#30495#39064#23548#20986#20026#25991#20214
    Caption = #23548#20986#21508#31456#33410#19979#30340#39064#30446
    ParentShowHint = False
    ShowHint = True
    TabOrder = 12
    OnClick = btn10Click
  end
  object btn11: TButton
    Left = 8
    Top = 264
    Width = 89
    Height = 25
    Hint = #20351#32771#28857#27604#20363#25353#39640#20302#25490#24207#65292#21516#26102#35843#25972#23545#24212#32771#28857
    Caption = #25490#24207#32771#28857#27604#20363
    ParentShowHint = False
    ShowHint = True
    TabOrder = 13
    OnClick = btn11Click
  end
  object btn12: TButton
    Left = 16
    Top = 291
    Width = 193
    Height = 25
    Hint = #23548#20986#20844#20849#22522#30784#21508#31456#33410#19979#30340#30495#39064#30446#20026#25991#20214
    Caption = #20844#20849#22522#30784#31456#33410#19979#30340#39064#30446
    ParentShowHint = False
    ShowHint = True
    TabOrder = 14
    OnClick = btn12Click
  end
  object chk_disp_anay: TCheckBox
    Left = 117
    Top = 127
    Width = 115
    Height = 135
    Caption = #21253#21547#35299#26512
    TabOrder = 16
  end
  object chk_disp_answer: TCheckBox
    Left = 4
    Top = 127
    Width = 113
    Height = 114
    Caption = #21253#21547#31572#26696
    TabOrder = 15
  end
  object chk_sgn_answer: TCheckBox
    Left = 6
    Top = 232
    Width = 217
    Height = 33
    Caption = #31572#26696#21644#35299#26512#21333#29420#23548#20986#19968#20010#25991#20214
    TabOrder = 17
  end
  object edtAnswer1: TRxRichEdit
    Left = 72
    Top = 144
    Width = 153
    Height = 17
    DrawEndPage = False
    Font.Charset = GB2312_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #23435#20307
    Font.Style = []
    ParentFont = False
    TabOrder = 18
    Visible = False
  end
  object Timer1: TTimer
    Interval = 100
    OnTimer = Timer1Timer
    Left = 14
    Top = 132
  end
end
