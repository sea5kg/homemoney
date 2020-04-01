object Form1: TForm1
  Left = 0
  Top = 0
  Caption = #1044#1086#1084#1072#1096#1085#1103#1103' '#1073#1091#1093#1075#1072#1083#1090#1077#1088#1080#1103' ('#1074#1089#1087#1086#1084#1086#1075#1072#1090#1077#1083#1100#1085#1072#1103' '#1087#1088#1086#1075#1088#1072#1084#1084#1072')'
  ClientHeight = 377
  ClientWidth = 922
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  DesignSize = (
    922
    377)
  PixelsPerInch = 96
  TextHeight = 13
  object lblStatus: TLabel
    Left = 8
    Top = 346
    Width = 48
    Height = 23
    Anchors = [akLeft, akBottom]
    Caption = '1111'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -19
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    ExplicitTop = 334
  end
  object edtFile: TEdit
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
  object Log: TMemo
    Left = 8
    Top = 139
    Width = 906
    Height = 178
    Anchors = [akLeft, akTop, akRight, akBottom]
    Color = clBtnFace
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 2
  end
  object ProgressBar1: TProgressBar
    Left = 8
    Top = 323
    Width = 906
    Height = 16
    Anchors = [akLeft, akRight, akBottom]
    TabOrder = 3
  end
  object btnViewClassifications: TButton
    Left = 535
    Top = 60
    Width = 226
    Height = 25
    Action = actViewClassifications
    TabOrder = 4
    Visible = False
  end
  object GroupBox1: TGroupBox
    Left = 9
    Top = 43
    Width = 257
    Height = 90
    Caption = #1044#1083#1103' '#1088#1072#1073#1086#1090#1099' '#1089' '#1084#1077#1089#1103#1094#1077#1084
    TabOrder = 5
    object cmbMonth: TComboBox
      Left = 15
      Top = 25
      Width = 226
      Height = 21
      Style = csDropDownList
      TabOrder = 0
    end
    object btnCalcClassifications: TButton
      Left = 15
      Top = 56
      Width = 226
      Height = 25
      Action = actCalcClassification
      TabOrder = 1
    end
  end
  object GroupBox2: TGroupBox
    Left = 272
    Top = 43
    Width = 257
    Height = 90
    Caption = #1044#1083#1103' '#1088#1072#1073#1086#1090#1099' '#1089' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1103#1084#1080
    TabOrder = 6
    object btnRedesignClassifications: TButton
      Left = 15
      Top = 25
      Width = 226
      Height = 25
      Action = actRedesignClassifications
      TabOrder = 0
    end
    object btnSortClassifications: TButton
      Left = 15
      Top = 56
      Width = 226
      Height = 25
      Action = actSortClassifications
      TabOrder = 1
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Excel-file|*.xls*'
    Left = 816
    Top = 56
  end
  object ActionList1: TActionList
    Left = 704
    Top = 48
    object actCalcClassification: TAction
      Caption = #1056#1072#1089#1089#1095#1080#1090#1072#1090#1100' '#1087#1086' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1103#1084
      OnExecute = actCalcClassificationExecute
      OnUpdate = actCalcClassificationUpdate
    end
    object actViewClassifications: TAction
      Caption = #1055#1088#1086#1089#1084#1086#1090#1088' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1081
      OnExecute = actViewClassificationsExecute
      OnUpdate = actViewClassificationsUpdate
    end
    object actRedesignClassifications: TAction
      Caption = #1055#1077#1088#1077#1089#1090#1088#1086#1080#1090#1100' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1080
      OnExecute = actRedesignClassificationsExecute
      OnUpdate = actRedesignClassificationsUpdate
    end
    object actSortClassifications: TAction
      Caption = #1054#1090#1089#1086#1088#1090#1080#1088#1086#1074#1072#1090#1100' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1080
      OnExecute = actSortClassificationsExecute
      OnUpdate = actSortClassificationsUpdate
    end
  end
end
