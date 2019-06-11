object Form1: TForm1
  Left = 192
  Top = 107
  Width = 415
  Height = 311
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 160
    Top = 168
    Width = 75
    Height = 25
    Caption = 'Button1'
    Enabled = False
    TabOrder = 0
  end
  object Edit1: TEdit
    Left = 128
    Top = 56
    Width = 121
    Height = 21
    TabOrder = 1
    Text = 'Edit1'
    OnChange = Edit1Change
  end
end
