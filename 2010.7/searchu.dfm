object Form1: TForm1
  Left = 220
  Top = 130
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = #20185#21073#22235'1.1'#23384#26723#32534#36753#22120'-happlygong'
  ClientHeight = 475
  ClientWidth = 498
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnClose = FormClose
  OnPaint = FormPaint
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 0
    Top = 16
    Width = 123
    Height = 13
    Alignment = taRightJustify
    AutoSize = False
    BiDiMode = bdLeftToRight
    Caption = #23384#30424#25991#20214#26410#25171#24320'  '
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlue
    Font.Height = -12
    Font.Name = #23435#20307
    Font.Style = []
    ParentBiDiMode = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
  end
  object player: TGroupBox
    Left = 136
    Top = 0
    Width = 361
    Height = 473
    Caption = #20154#29289#23646#24615
    TabOrder = 5
    object Label2: TLabel
      Left = 320
      Top = 40
      Width = 24
      Height = 48
      Font.Charset = GB2312_CHARSET
      Font.Color = clGray
      Font.Height = -48
      Font.Name = #26999#20307'_GB2312'
      Font.Style = []
      ParentFont = False
      Transparent = True
      WordWrap = True
    end
  end
  object Basict: TGroupBox
    Left = 147
    Top = 20
    Width = 209
    Height = 417
    Caption = #22522#26412#23646#24615
    TabOrder = 6
    object LED_Level: TLabeledEdit
      Left = 19
      Top = 44
      Width = 42
      Height = 21
      Hint = #31561#32423#26368#22823'99'
      Ctl3D = True
      EditLabel.Width = 27
      EditLabel.Height = 13
      EditLabel.Hint = '1~99'
      EditLabel.Caption = #31561#32423' '
      EditLabel.ParentShowHint = False
      EditLabel.ShowHint = True
      ParentCtl3D = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnKeyPress = LED_LevelKeyPress
    end
    object LED_Exper: TLabeledEdit
      Left = 109
      Top = 44
      Width = 80
      Height = 21
      Hint = #32463#39564#20540#20915#23450#31561#32423
      EditLabel.Width = 45
      EditLabel.Height = 13
      EditLabel.Hint = #32463#39564#20540#20915#23450#20154#29289#30340#31561#32423#65292#13#30452#25509#20462#25913#31561#32423#27809#26377#29992#12290
      EditLabel.Caption = #32463#39564#20540'   '
      EditLabel.ParentShowHint = False
      EditLabel.ShowHint = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnKeyPress = LED_LevelKeyPress
    end
    object LED_HP: TLabeledEdit
      Left = 19
      Top = 85
      Width = 80
      Height = 21
      EditLabel.Width = 12
      EditLabel.Height = 13
      EditLabel.Caption = #31934
      TabOrder = 2
      OnKeyPress = LED_LevelKeyPress
    end
    object LED_MAXHP: TLabeledEdit
      Left = 109
      Top = 85
      Width = 80
      Height = 21
      EditLabel.Width = 60
      EditLabel.Height = 13
      EditLabel.Caption = #31934#26368#22823#20540'    '
      TabOrder = 3
      OnKeyPress = LED_LevelKeyPress
    end
    object LED_RAGE: TLabeledEdit
      Left = 19
      Top = 127
      Width = 80
      Height = 21
      EditLabel.Width = 12
      EditLabel.Height = 13
      EditLabel.Caption = #27668
      TabOrder = 4
      OnKeyPress = LED_LevelKeyPress
    end
    object LED_MRAGE: TLabeledEdit
      Left = 109
      Top = 127
      Width = 80
      Height = 21
      Hint = #27668#26368#22823#22266#23450#20026'100'#19981#33021#20462#25913
      EditLabel.Width = 57
      EditLabel.Height = 13
      EditLabel.Caption = #27668#26368#22823#20540'   '
      ParentShowHint = False
      ReadOnly = True
      ShowHint = True
      TabOrder = 5
    end
    object LED_MP: TLabeledEdit
      Left = 19
      Top = 168
      Width = 80
      Height = 21
      EditLabel.Width = 12
      EditLabel.Height = 13
      EditLabel.Caption = #31070
      TabOrder = 6
      OnKeyPress = LED_LevelKeyPress
    end
    object LED_MAXMP: TLabeledEdit
      Left = 109
      Top = 168
      Width = 80
      Height = 21
      EditLabel.Width = 60
      EditLabel.Height = 13
      EditLabel.Caption = #31070#26368#22823#20540'    '
      TabOrder = 7
      OnKeyPress = LED_LevelKeyPress
    end
    object LED_Wuli1: TLabeledEdit
      Left = 19
      Top = 210
      Width = 80
      Height = 21
      Hint = #22522#30784#20540
      EditLabel.Width = 12
      EditLabel.Height = 13
      EditLabel.Caption = #25915
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      OnKeyPress = LED_LevelKeyPress
    end
    object LED_Fang1: TLabeledEdit
      Left = 19
      Top = 252
      Width = 80
      Height = 21
      EditLabel.Width = 12
      EditLabel.Height = 13
      EditLabel.Caption = #38450
      TabOrder = 9
      OnKeyPress = LED_LevelKeyPress
    end
    object LED_Lucky1: TLabeledEdit
      Left = 19
      Top = 335
      Width = 80
      Height = 21
      EditLabel.Width = 18
      EditLabel.Height = 13
      EditLabel.Caption = #36816'  '
      TabOrder = 11
      OnKeyPress = LED_LevelKeyPress
    end
    object LED_Will1: TLabeledEdit
      Left = 19
      Top = 377
      Width = 80
      Height = 21
      EditLabel.Width = 18
      EditLabel.Height = 13
      EditLabel.Caption = #28789'  '
      TabOrder = 12
      OnKeyPress = LED_LevelKeyPress
    end
    object LED_Speed1: TLabeledEdit
      Left = 19
      Top = 293
      Width = 80
      Height = 21
      EditLabel.Width = 12
      EditLabel.Height = 13
      EditLabel.Caption = #36895
      TabOrder = 10
      OnKeyPress = LED_LevelKeyPress
    end
    object LED_Wuli: TLabeledEdit
      Left = 109
      Top = 210
      Width = 80
      Height = 21
      Hint = #23454#38469#20540
      EditLabel.Width = 15
      EditLabel.Height = 13
      EditLabel.Caption = #25915' '
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 13
    end
    object LED_Fang: TLabeledEdit
      Left = 109
      Top = 251
      Width = 80
      Height = 21
      EditLabel.Width = 15
      EditLabel.Height = 13
      EditLabel.Caption = #38450' '
      Enabled = False
      TabOrder = 14
    end
    object LED_Lucky: TLabeledEdit
      Left = 109
      Top = 334
      Width = 80
      Height = 21
      EditLabel.Width = 18
      EditLabel.Height = 13
      EditLabel.Caption = #36816'  '
      Enabled = False
      TabOrder = 16
    end
    object LED_Speed: TLabeledEdit
      Left = 109
      Top = 293
      Width = 80
      Height = 21
      EditLabel.Width = 18
      EditLabel.Height = 13
      EditLabel.Caption = #36895'  '
      Enabled = False
      TabOrder = 15
    end
    object LED_Will: TLabeledEdit
      Left = 109
      Top = 376
      Width = 80
      Height = 21
      EditLabel.Width = 18
      EditLabel.Height = 13
      EditLabel.Caption = #28789'  '
      Enabled = False
      TabOrder = 17
    end
  end
  object goods: TGroupBox
    Left = 0
    Top = 180
    Width = 129
    Height = 147
    Caption = #29289#21697
    TabOrder = 4
    object BitBtn3: TBitBtn
      Left = 12
      Top = 82
      Width = 107
      Height = 25
      Hint = #20351#24050#24471#21040#30340#29289#21697#25968#37327#36798#20108#30334#22810#20010#13#19981#20462#25913#39#21095#24773#39#29289#21697
      Caption = #21482#25913#29616#26377#29289#21697#25968#37327
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BitBtn3Click
    end
    object LMoney: TLabeledEdit
      Left = 10
      Top = 36
      Width = 109
      Height = 21
      Hint = #24314#35758#20540'9999999'
      EditLabel.Width = 30
      EditLabel.Height = 13
      EditLabel.Caption = #37329#38065'  '
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnKeyPress = LED_LevelKeyPress
    end
    object BitBtn1: TBitBtn
      Left = 12
      Top = 113
      Width = 106
      Height = 25
      Hint = #20462#25913#20869#23384#25968#25454
      Caption = #30830#35748#20462#25913
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BitBtn1Click
    end
    object wupin: TCheckBox
      Left = 9
      Top = 61
      Width = 53
      Height = 17
      Hint = #33719#24471'246'#31181#29289#21697#27599#31181#29289#21697'255'#20010
      Caption = #29289#21697#20840
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #23435#20307
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
    end
    object tupu: TCheckBox
      Left = 67
      Top = 61
      Width = 55
      Height = 17
      Hint = #33719#24471'150'#31181#22270#35889
      Caption = #22270#35889#20840
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #23435#20307
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
    end
  end
  object atk: TGroupBox
    Left = 375
    Top = 20
    Width = 105
    Height = 353
    Caption = #29305#25216
    TabOrder = 7
    object ck: TCheckListBox
      Left = 8
      Top = 17
      Width = 89
      Height = 328
      Hint = #35201#20808#21246#36873#25481#24050#26377#30340#13#25165#33021#20877#36873#20854#20182#30340#13#29305#25216#24635#25968#19981#21464
      OnClickCheck = ckClickCheck
      Ctl3D = False
      ItemHeight = 13
      Items.Strings = (
        #29454#20861
        #39134#32701#31661
        #33853#26143#24335
        #33181#35010
        #36880#26376#24335
        #29378#29022
        #24696#22825#36143#26085#24335
        #25628#22218#25506#23453
        #20940#31354#25688#26143
        #28895#38632#22842#39746
        #20094#22372#19968#25527
        #20116#27602#30722
        #26080#24433#36830#21073#35776
        #27785#27700#28070#24515
        #22825#29572#20116#38899
        #29071#27264#20928#34915
        #33487#21512#36890#31373
        #37257#29983#26790#27515
        #39746#26790#39749#26354
        #19977#25165#26397#20803
        #22235#26041#32899#25947
        #21270#30456#30495#22914#21073
        #20116#28789#24402#23447
        #21315#26041#27544#20809#21073
        #19978#28165#30772#20113#21073)
      ParentCtl3D = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
    end
  end
  object Button4: TButton
    Left = 4
    Top = 137
    Width = 125
    Height = 25
    Caption = #24917#23481#32043#33521
    TabOrder = 3
    OnClick = Button4Click
  end
  object Button3: TButton
    Left = 4
    Top = 105
    Width = 125
    Height = 25
    Caption = #26611#26790#29827
    TabOrder = 2
    OnClick = Button3Click
  end
  object Button2: TButton
    Left = 4
    Top = 73
    Width = 125
    Height = 25
    Caption = #38889#33777#32433
    TabOrder = 1
    OnClick = Button2Click
  end
  object Button1: TButton
    Left = 4
    Top = 41
    Width = 125
    Height = 25
    Caption = #20113#22825#27827
    TabOrder = 0
    OnClick = Button1Click
  end
  object atkall: TCheckBox
    Left = 376
    Top = 421
    Width = 97
    Height = 17
    Hint = #19968#27425#25915#20987#20840#37096#30340#25932#20154
    Caption = #20840#20307#25915#20987
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = #23435#20307
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 10
  end
  object inteam: TCheckBox
    Left = 376
    Top = 399
    Width = 97
    Height = 17
    Caption = #22312#38431#20237#20013
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = #23435#20307
    Font.Style = []
    ParentFont = False
    TabOrder = 9
  end
  object Hostcheck: TCheckBox
    Left = 376
    Top = 376
    Width = 97
    Height = 17
    Hint = #22914#26524#24819#33258#24049#23398#20185#26415#65292#13#20294#20185#26415#28857#25968#21448#19981#22815#65292#13#21487#20197#25226#31561#32423#25913#39640#20123#65292#13#20185#26415#28857#25968#23601#20250#22686#21152#12290#13#23398#23436#20877#25226#31561#32423#25913#22238#26469
    Caption = #20185#26415#20840#23398#24471
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = #23435#20307
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 8
  end
  object Button5: TButton
    Left = 147
    Top = 440
    Width = 337
    Height = 25
    Hint = #20462#25913#20869#23384#25968#25454
    Caption = #30830#35748#20462#25913
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = #23435#20307
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 11
    OnClick = Button5Click
  end
  object BitBtn2: TBitBtn
    Left = 8
    Top = 360
    Width = 117
    Height = 107
    Hint = #23558#20869#23384#25968#25454#20445#23384#21040#25991#26723
    Caption = #20445#23384#25991#20214
    ParentShowHint = False
    ShowHint = True
    TabOrder = 13
    OnClick = BitBtn2Click
  end
  object Button6: TButton
    Left = 8
    Top = 331
    Width = 117
    Height = 25
    Hint = #33258#21160#20462#25913#38065#12289#29289#13#31934#27668#31070#25915#38450#36816#28789#13#22235#20154#37117#22312#38431#20237#20013
    Caption = #33258#21160#20462#25913
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = #23435#20307
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 12
    OnClick = Button6Click
  end
  object OpenDialog1: TOpenDialog
    DefaultExt = '*.dat'
    Filter = '*.dat|*.dat|*.*|*.*'
    InitialDir = 'C:\Program Files\SoftStar\PAL4\save'
    Left = 52
    Top = 8
  end
  object SaveDialog1: TSaveDialog
    Left = 84
    Top = 8
  end
  object MainMenu1: TMainMenu
    Left = 16
    Top = 8
    object N1: TMenuItem
      Caption = #25991#20214'(&F)'
      object N2: TMenuItem
        Caption = #25171#24320#23384#26723'(&O)'
        OnClick = N2Click
      end
      object N3: TMenuItem
        Caption = #20445#23384#25991#26723'(&S)'
        OnClick = N3Click
      end
      object N4: TMenuItem
        Caption = #20851#38381'(&C)'
        OnClick = N4Click
      end
    end
    object N5: TMenuItem
      Caption = #20851#20110'(&H)'
      object N6: TMenuItem
        Caption = #20851#20110'(&A)'
        OnClick = N6Click
      end
    end
  end
end
