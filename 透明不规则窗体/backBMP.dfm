object Form1: TForm1
  Left = 192
  Top = 107
  BorderStyle = bsNone
  Caption = 'Form1'
  ClientHeight = 97
  ClientWidth = 116
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PopupMenu = PopupMenu1
  OnCreate = FormCreate
  OnPaint = FormPaint
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 8
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Button1'
    TabOrder = 0
    OnClick = Button1Click
  end
  object PopupMenu1: TPopupMenu
    Left = 48
    Top = 72
    object N1: TMenuItem
      Caption = #36864#20986
      OnClick = N1Click
    end
  end
end
