object SetCondtionForm: TSetCondtionForm
  Left = 492
  Top = 147
  Width = 298
  Height = 243
  BorderIcons = [biSystemMenu]
  Caption = #35774#32622#26597#35810#26465#20214
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object lblTitle: TLabel
    Left = 24
    Top = 16
    Width = 96
    Height = 13
    Caption = #25968#25454#24211#20013#26377#39064#30446#65306
  end
  object chkall: TCheckBox
    Left = 32
    Top = 128
    Width = 97
    Height = 17
    Caption = #20840#36873
    Checked = True
    State = cbChecked
    TabOrder = 0
    OnClick = chkallClick
  end
  object edtCondition: TEdit
    Left = 32
    Top = 152
    Width = 209
    Height = 21
    TabOrder = 1
    Text = #20363#22914#65306'TH>8 and TH<12  Or TH=14'
    Visible = False
    OnEnter = edtConditionEnter
  end
  object btnOK: TButton
    Left = 96
    Top = 176
    Width = 75
    Height = 25
    Caption = #30830'    '#23450
    TabOrder = 2
    Visible = False
    OnClick = btnOKClick
  end
end
