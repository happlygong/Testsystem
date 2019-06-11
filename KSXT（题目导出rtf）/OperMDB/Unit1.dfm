object Form1: TForm1
  Left = 200
  Top = 125
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  BorderWidth = 1
  Caption = #25968#25454#23548#20986':->OperMDB.mdb'
  ClientHeight = 335
  ClientWidth = 403
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  Visible = True
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 66
    Top = 260
    Width = 117
    Height = 13
    Caption = #27491#22312#22788#29702#65292#35831#31245#20505'...    '
    Visible = False
  end
  object Label2: TLabel
    Left = 80
    Top = 275
    Width = 3
    Height = 13
  end
  object Label3: TLabel
    Left = 60
    Top = 290
    Width = 3
    Height = 13
    Cursor = crHandPoint
    Hint = #21478#23384#20026'.doc'#25991#20214#21487#20197#20943#23567#25991#20214#20307#31215
    Alignment = taCenter
    ParentShowHint = False
    ShowHint = True
    OnClick = Label3Click
    OnMouseEnter = Label3MouseEnter
    OnMouseLeave = Label3MouseLeave
  end
  object RxRichEdit1: TRxRichEdit
    Left = 24
    Top = 72
    Width = 201
    Height = 89
    DrawEndPage = False
    Alignment = taCenter
    Font.Charset = GB2312_CHARSET
    Font.Color = clRed
    Font.Height = -14
    Font.Name = #23435#20307
    Font.Style = []
    ImeName = #20013#25991' - QQ'#20116#31508#36755#20837#27861
    ParentFont = False
    TabOrder = 1
    Visible = False
  end
  object Button1: TButton
    Left = 76
    Top = 306
    Width = 81
    Height = 25
    Caption = #23548#20986#25968#25454
    TabOrder = 0
    OnClick = Button1Click
  end
  object RG1: TRadioGroup
    Left = 16
    Top = 8
    Width = 369
    Height = 49
    Caption = #35831#36873#25321#31867#22411
    Columns = 3
    Items.Strings = (
      #21382#24180#30495#39064
      #27169#25311#35797#39064
      #19978#26426#25805#20316#39064)
    TabOrder = 2
    OnClick = RG1Click
  end
  object RG2: TRadioGroup
    Left = 16
    Top = 64
    Width = 209
    Height = 105
    Caption = #35831#36873#25321#31185#30446
    Columns = 2
    Enabled = False
    Items.Strings = (
      #20108#32423'Access'
      #20108#32423'VB'
      #20108#32423'C'
      #20108#32423'VF'
      #20108#32423'C++'
      #19968#32423'MS '
      #20108#32423'MS')
    TabOrder = 3
  end
  object rg3: TRadioGroup
    Left = 16
    Top = 177
    Width = 209
    Height = 81
    Enabled = False
    ItemIndex = 1
    Items.Strings = (
      #35299#26512#22312#27599#19968#39064#21518#38754
      #35299#26512#22312#25152#26377#39064#21518#38754)
    TabOrder = 4
  end
  object chk1: TCheckBox
    Left = 32
    Top = 177
    Width = 97
    Height = 17
    Caption = #21253#25324#35299#26512#12288#12288
    Constraints.MaxHeight = 17
    Constraints.MaxWidth = 97
    TabOrder = 5
    OnClick = chk1Click
  end
  object GRPorderby1: TGroupBox
    Left = 248
    Top = 72
    Width = 137
    Height = 72
    Caption = #25490#24207#26041#24335
    TabOrder = 6
    object rbTHXH: TRadioButton
      Left = 32
      Top = 24
      Width = 113
      Height = 17
      Caption = #25353'TH'#12289'XH'
      TabOrder = 0
    end
    object rbBianhao: TRadioButton
      Left = 32
      Top = 49
      Width = 113
      Height = 17
      Caption = #25353#32534#21495
      TabOrder = 1
    end
  end
  object grpWhere1: TGroupBox
    Left = 248
    Top = 168
    Width = 137
    Height = 89
    Caption = #23548#20986#26465#20214#65288#33539#22260#65289
    TabOrder = 7
    object chkNothing: TCheckBox
      Left = 32
      Top = 19
      Width = 97
      Height = 17
      Caption = #26080#26465#20214
      TabOrder = 0
    end
    object chkBianHao: TCheckBox
      Left = 32
      Top = 43
      Width = 97
      Height = 17
      Caption = #32534#21495#33539#22260
      Checked = True
      State = cbChecked
      TabOrder = 1
    end
    object chkTH: TCheckBox
      Left = 32
      Top = 67
      Width = 97
      Height = 17
      Caption = 'TH'#33539#22260
      TabOrder = 2
    end
  end
end
