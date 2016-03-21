object frmMain: TfrmMain
  Left = 192
  Top = 133
  Width = 385
  Height = 540
  Caption = 'frmMain'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object edNumbertoGenerate: TEdit
    Left = 24
    Top = 32
    Width = 121
    Height = 21
    TabOrder = 0
    Text = '0'
  end
  object btnGenerate: TButton
    Left = 184
    Top = 32
    Width = 75
    Height = 25
    Caption = 'Generate'
    TabOrder = 1
    OnClick = btnGenerateClick
  end
  object MemoAddresses: TMemo
    Left = 24
    Top = 64
    Width = 329
    Height = 361
    TabOrder = 2
  end
end
