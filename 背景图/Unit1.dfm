object Form1: TForm1
  Left = 177
  Top = 245
  Width = 979
  Height = 563
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnPaint = FormPaint
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 24
    Top = 40
    Width = 123
    Height = 73
    Caption = 'Button1'
    TabOrder = 0
  end
  object Edit1: TEdit
    Left = 176
    Top = 8
    Width = 409
    Height = 121
    AutoSize = False
    TabOrder = 1
    Text = 'Edit1'
    OnEnter = Edit1Enter
  end
  object Memo1: TMemo
    Left = 120
    Top = 144
    Width = 473
    Height = 289
    TabOrder = 2
    OnEnter = Memo1Enter
  end
  object Panel1: TPanel
    Left = 648
    Top = 56
    Width = 201
    Height = 369
    TabOrder = 3
    OnClick = Panel1Click
    OnResize = Panel1Resize
  end
end
