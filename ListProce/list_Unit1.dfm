object Form1: TForm1
  Left = 192
  Top = 107
  Width = 633
  Height = 486
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = #21015#20986#31995#32479#24403#21069#36827#31243
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 100
    Top = 61
    Width = 32
    Height = 13
    Caption = 'Label1'
  end
  object Label2: TLabel
    Left = 248
    Top = 62
    Width = 66
    Height = 13
    Caption = #25152#22312#36335#24452#65306'  '
  end
  object Label3: TLabel
    Left = 320
    Top = 62
    Width = 32
    Height = 13
    Caption = 'Label3'
  end
  object Label4: TLabel
    Left = 432
    Top = 16
    Width = 54
    Height = 13
    Caption = #25991#20214#21517#65306'  '
  end
  object Label5: TLabel
    Left = 488
    Top = 16
    Width = 32
    Height = 13
    Caption = 'Label5'
  end
  object ListBox1: TListBox
    Left = 0
    Top = 90
    Width = 241
    Height = 369
    ItemHeight = 13
    TabOrder = 0
    OnClick = ListBox1Click
  end
  object Button1: TButton
    Left = 0
    Top = 56
    Width = 75
    Height = 25
    Caption = #21015#20986#36827#31243
    TabOrder = 1
    OnClick = Button1Click
  end
  object Edit1: TEdit
    Left = 8
    Top = 8
    Width = 329
    Height = 21
    TabOrder = 2
    Text = 'Edit1'
  end
  object ListView1: TListView
    Left = 272
    Top = 89
    Width = 353
    Height = 369
    Columns = <
      item
        Caption = #25991#20214#21517
        Width = 285
      end
      item
        Caption = 'ID'
        Width = 60
      end>
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    IconOptions.Arrangement = iaLeft
    IconOptions.AutoArrange = True
    IconOptions.WrapText = False
    ParentFont = False
    TabOrder = 3
    ViewStyle = vsReport
    OnClick = ListView1Click
  end
  object Button2: TButton
    Left = 144
    Top = 56
    Width = 75
    Height = 25
    Caption = 'Button2'
    TabOrder = 4
    OnClick = Button2Click
  end
  object BitBtn1: TBitBtn
    Left = 352
    Top = 8
    Width = 25
    Height = 25
    Hint = #25171#24320#27492#30446#24405
    Caption = 'BitBtn1'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
    OnClick = BitBtn1Click
  end
end
