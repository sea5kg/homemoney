object FormSelectClass: TFormSelectClass
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = #1042#1099#1073#1088#1072#1090#1100' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1102
  ClientHeight = 100
  ClientWidth = 378
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 13
    Width = 3
    Height = 13
  end
  object ComboBox1: TComboBox
    Left = 8
    Top = 32
    Width = 345
    Height = 21
    TabOrder = 0
  end
  object Button1: TButton
    Left = 205
    Top = 59
    Width = 75
    Height = 25
    Caption = 'OK'
    ModalResult = 1
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 286
    Top = 59
    Width = 75
    Height = 25
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 2
    OnClick = Button2Click
  end
end
