object Form1: TForm1
  Left = 388
  Top = 104
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  BorderWidth = 1
  Caption = #25968#25454#23548#20986
  ClientHeight = 335
  ClientWidth = 239
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
    Width = 209
    Height = 49
    Caption = #35831#36873#25321#31867#22411
    Columns = 2
    Items.Strings = (
      #19978#26426#39064
      #31508#35797#39064)
    TabOrder = 2
  end
  object RG2: TRadioGroup
    Left = 16
    Top = 64
    Width = 209
    Height = 105
    Caption = #35831#36873#25321#31185#30446
    Columns = 2
    ItemIndex = 0
    Items.Strings = (
      #19977#32423#32593#32476)
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
end
