object Form1: TForm1
  Left = 192
  Top = 107
  Width = 436
  Height = 319
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 48
    Top = 40
    Width = 75
    Height = 25
    Caption = 'Button1'
    TabOrder = 0
  end
  object Button2: TButton
    Left = 48
    Top = 82
    Width = 75
    Height = 25
    Caption = 'Button2'
    TabOrder = 1
  end
  object Button3: TButton
    Left = 48
    Top = 125
    Width = 75
    Height = 25
    Caption = 'Button3'
    TabOrder = 2
  end
  object Button4: TButton
    Left = 48
    Top = 168
    Width = 75
    Height = 25
    Caption = 'Button4'
    TabOrder = 3
  end
  object ListBox1: TListBox
    Left = 152
    Top = 40
    Width = 121
    Height = 113
    ItemHeight = 13
    Items.Strings = (
      #19968#32423'B'
      #19968#32423'MS'
      #20108#32423'C'
      #20108#32423'VB'
      #20108#32423'VF'
      #20108#32423'Access'
      #19977#32423#32593#32476
      #19977#32423#25968#25454#24211)
    TabOrder = 4
    OnClick = ListBox1Click
  end
end
