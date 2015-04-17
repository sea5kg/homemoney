object Form1: TForm1
  Left = 0
  Top = 0
  Caption = #1044#1086#1084#1072#1096#1085#1103#1103' '#1073#1091#1093#1075#1072#1083#1090#1077#1088#1080#1103' ('#1074#1089#1087#1086#1084#1086#1075#1072#1090#1077#1083#1100#1085#1072#1103' '#1087#1088#1086#1075#1088#1072#1084#1084#1072')'
  ClientHeight = 365
  ClientWidth = 922
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  DesignSize = (
    922
    365)
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 48
    Width = 113
    Height = 13
    Caption = #1044#1083#1103' '#1088#1072#1073#1086#1090#1099' '#1089' '#1084#1077#1089#1103#1094#1077#1084
  end
  object Label2: TLabel
    Left = 16
    Top = 144
    Width = 162
    Height = 13
    Caption = #1044#1083#1103' '#1088#1072#1073#1086#1090#1099' '#1089' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1103#1084#1080
  end
  object Edit1: TEdit
    Left = 8
    Top = 8
    Width = 789
    Height = 21
    Anchors = [akLeft, akTop, akRight]
    Color = clBtnFace
    ReadOnly = True
    TabOrder = 0
  end
  object Button1: TButton
    Left = 803
    Top = 6
    Width = 111
    Height = 25
    Anchors = [akTop, akRight]
    Caption = #1042#1099#1073#1088#1072#1090#1100' '#1092#1072#1081#1083
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 24
    Top = 72
    Width = 321
    Height = 25
    Caption = #1056#1072#1089#1089#1095#1080#1090#1072#1090#1100' '#1089#1091#1084#1084#1099' '#1087#1086' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1103#1084
    Enabled = False
    TabOrder = 2
  end
  object Button3: TButton
    Left = 24
    Top = 103
    Width = 321
    Height = 25
    Caption = #1042#1099#1074#1077#1089#1090#1080' '#1082#1083#1072#1089#1089#1080#1092#1080#1094#1080#1088#1086#1074#1072#1085#1085#1099#1081' '#1089#1087#1080#1089#1086#1082' '#1080#1079' '#1090#1077#1082#1091#1097#1077#1075#1086' '#1083#1080#1089#1090#1072
    Enabled = False
    TabOrder = 3
  end
  object Button5: TButton
    Left = 24
    Top = 163
    Width = 321
    Height = 25
    Caption = #1055#1077#1088#1077#1089#1090#1088#1086#1080#1090#1100' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1080
    Enabled = False
    TabOrder = 4
    OnClick = Button5Click
  end
  object Button6: TButton
    Left = 24
    Top = 194
    Width = 321
    Height = 25
    Caption = #1054#1090#1089#1086#1088#1090#1080#1088#1086#1074#1072#1090#1100' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1080
    Enabled = False
    TabOrder = 5
    OnClick = Button6Click
  end
  object ComboBox1: TComboBox
    Left = 160
    Top = 45
    Width = 185
    Height = 21
    Style = csDropDownList
    TabOrder = 6
  end
  object Log: TMemo
    Left = 8
    Top = 225
    Width = 906
    Height = 110
    Anchors = [akLeft, akTop, akRight, akBottom]
    Color = clBtnFace
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 7
  end
  object ProgressBar1: TProgressBar
    Left = 8
    Top = 341
    Width = 906
    Height = 16
    Anchors = [akLeft, akRight, akBottom]
    TabOrder = 8
  end
  object Button7: TButton
    Left = 360
    Top = 163
    Width = 217
    Height = 25
    Caption = #1055#1088#1086#1089#1084#1086#1090#1088' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1081
    Enabled = False
    TabOrder = 9
    OnClick = Button7Click
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Excel-file|*.xls*'
    Left = 576
    Top = 56
  end
end
