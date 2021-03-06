object FormMain: TFormMain
  Left = 435
  Top = 282
  Width = 976
  Height = 256
  Caption = #1047#1072#1075#1088#1091#1079#1082#1072' '#1090#1077#1089#1090#1086#1074' '#1080#1079' Excel'
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
  object LabelCurrentQuestion: TLabel
    Left = 8
    Top = 8
    Width = 133
    Height = 13
    Caption = #1053#1086#1084#1077#1088' '#1090#1077#1082#1091#1097#1077#1075#1086' '#1074#1086#1087#1088#1086#1089#1072':'
  end
  object LabelFileName: TLabel
    Left = 8
    Top = 32
    Width = 111
    Height = 13
    Caption = #1048#1084#1103' '#1090#1077#1082#1091#1097#1077#1075#1086' '#1092#1072#1081#1083#1072':'
  end
  object LabelFilePath: TLabel
    Left = 8
    Top = 56
    Width = 70
    Height = 13
    Caption = #1055#1091#1090#1100' '#1082' '#1092#1072#1081#1083#1091':'
  end
  object LabelStatus: TLabel
    Left = 192
    Top = 8
    Width = 93
    Height = 13
    Caption = #1057#1090#1072#1090#1091#1089': '#1086#1078#1080#1076#1072#1085#1080#1077'.'
  end
  object ButtonOpen: TButton
    Left = 8
    Top = 88
    Width = 105
    Height = 25
    Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1090#1077#1084#1091
    TabOrder = 0
    OnClick = ButtonOpenClick
  end
  object Memo1: TMemo
    Left = 0
    Top = 128
    Width = 960
    Height = 89
    Align = alBottom
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssVertical
    TabOrder = 1
  end
  object ListBoxFileList: TListBox
    Left = 496
    Top = 0
    Width = 464
    Height = 128
    Align = alRight
    ItemHeight = 13
    TabOrder = 2
    Visible = False
  end
  object ButtonPutStudents: TButton
    Left = 120
    Top = 88
    Width = 169
    Height = 25
    Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1089#1087#1080#1089#1086#1082' '#1089#1090#1091#1076#1077#1085#1090#1086#1074
    TabOrder = 3
    OnClick = ButtonPutStudentsClick
  end
  object IBDatabaseImport: TIBDatabase
    Connected = True
    DatabaseName = 'D:\'#1055#1088#1086#1075#1088#1072#1084#1084#1099'\'#1058#1077#1089#1090#1099'\'#1041#1072#1079#1099'\TEST_VSMA.GDB'
    Params.Strings = (
      'user_name=sysdba'
      'password=masterkey'
      'lc_ctype=WIN1251')
    LoginPrompt = False
    DefaultTransaction = IBTransactionTheme
    IdleTimer = 0
    SQLDialect = 1
    TraceFlags = []
    Left = 88
    Top = 48
  end
  object IBUpdateSQLQuestion: TIBUpdateSQL
    RefreshSQL.Strings = (
      'Select '
      '  N_VOPR,'
      '  N_TEMA,'
      '  NAME_VOPR_L,'
      '  NAME_VOPR_P,'
      '  TIP_VOPR,'
      '  KOL_OTV,'
      '  VOPR_PIC'
      'from vopros '
      'where'
      '  N_VOPR = :N_VOPR')
    ModifySQL.Strings = (
      'update vopros'
      'set'
      '  N_VOPR = :N_VOPR,'
      '  N_TEMA = :N_TEMA,'
      '  NAME_VOPR_L = :NAME_VOPR_L,'
      '  NAME_VOPR_P = :NAME_VOPR_P,'
      '  TIP_VOPR = :TIP_VOPR,'
      '  KOL_OTV = :KOL_OTV,'
      '  VOPR_PIC = :VOPR_PIC'
      'where'
      '  N_VOPR = :OLD_N_VOPR')
    InsertSQL.Strings = (
      'insert into vopros'
      '  (N_VOPR, N_TEMA, NAME_VOPR_L, NAME_VOPR_P, TIP_VOPR, KOL_OTV, '
      'VOPR_PIC)'
      'values'
      
        '  (:N_VOPR, :N_TEMA, :NAME_VOPR_L, :NAME_VOPR_P, :TIP_VOPR, :KOL' +
        '_OTV, '
      ':VOPR_PIC)')
    DeleteSQL.Strings = (
      'delete from vopros'
      'where'
      '  N_VOPR = :OLD_N_VOPR')
    Left = 168
    Top = 48
  end
  object IBQueryQuestion: TIBQuery
    Database = IBDatabaseImport
    Transaction = IBTransactionTheme
    BufferChunks = 1000
    CachedUpdates = True
    SQL.Strings = (
      'select * from vopros')
    UpdateObject = IBUpdateSQLQuestion
    GeneratorField.Field = 'N_VOPR'
    GeneratorField.Generator = 'NEW_VOPR'
    Left = 128
    Top = 48
    object IBQueryQuestionN_VOPR: TIntegerField
      FieldName = 'N_VOPR'
      Origin = 'VOPROS.N_VOPR'
      Required = True
    end
    object IBQueryQuestionN_TEMA: TIntegerField
      FieldName = 'N_TEMA'
      Origin = 'VOPROS.N_TEMA'
      Required = True
    end
    object IBQueryQuestionNAME_VOPR_L: TIBStringField
      FieldName = 'NAME_VOPR_L'
      Origin = 'VOPROS.NAME_VOPR_L'
      Size = 3000
    end
    object IBQueryQuestionNAME_VOPR_P: TIBStringField
      FieldName = 'NAME_VOPR_P'
      Origin = 'VOPROS.NAME_VOPR_P'
      Size = 3000
    end
    object IBQueryQuestionTIP_VOPR: TSmallintField
      FieldName = 'TIP_VOPR'
      Origin = 'VOPROS.TIP_VOPR'
      Required = True
    end
    object IBQueryQuestionKOL_OTV: TIntegerField
      FieldName = 'KOL_OTV'
      Origin = 'VOPROS.KOL_OTV'
      Required = True
    end
    object IBQueryQuestionVOPR_PIC: TBlobField
      FieldName = 'VOPR_PIC'
      Origin = 'VOPROS.VOPR_PIC'
      Size = 8
    end
  end
  object IBQueryAnswer: TIBQuery
    Database = IBDatabaseImport
    Transaction = IBTransactionTheme
    BufferChunks = 1000
    CachedUpdates = True
    SQL.Strings = (
      'select * from otvet')
    UpdateObject = IBUpdateSQLAnswer
    GeneratorField.Field = 'N_OTV'
    GeneratorField.Generator = 'NEW_OTVET'
    Left = 208
    Top = 48
    object IBQueryAnswerN_OTV: TIntegerField
      FieldName = 'N_OTV'
      Origin = 'OTVET.N_OTV'
      Required = True
    end
    object IBQueryAnswerN_VOPR: TIntegerField
      FieldName = 'N_VOPR'
      Origin = 'OTVET.N_VOPR'
      Required = True
    end
    object IBQueryAnswerNAME_OTV: TIBStringField
      FieldName = 'NAME_OTV'
      Origin = 'OTVET.NAME_OTV'
      Size = 255
    end
    object IBQueryAnswerPR_PR: TSmallintField
      FieldName = 'PR_PR'
      Origin = 'OTVET.PR_PR'
      Required = True
    end
    object IBQueryAnswerN_OTV_S: TIntegerField
      FieldName = 'N_OTV_S'
      Origin = 'OTVET.N_OTV_S'
    end
    object IBQueryAnswerOTV_PIC: TBlobField
      FieldName = 'OTV_PIC'
      Origin = 'OTVET.OTV_PIC'
      Size = 8
    end
  end
  object IBUpdateSQLAnswer: TIBUpdateSQL
    RefreshSQL.Strings = (
      'Select '
      '  N_OTV,'
      '  N_VOPR,'
      '  NAME_OTV,'
      '  PR_PR,'
      '  N_OTV_S,'
      '  OTV_PIC'
      'from otvet '
      'where'
      '  N_OTV = :N_OTV')
    ModifySQL.Strings = (
      'update otvet'
      'set'
      '  N_OTV = :N_OTV,'
      '  N_VOPR = :N_VOPR,'
      '  NAME_OTV = :NAME_OTV,'
      '  PR_PR = :PR_PR,'
      '  N_OTV_S = :N_OTV_S,'
      '  OTV_PIC = :OTV_PIC'
      'where'
      '  N_OTV = :OLD_N_OTV')
    InsertSQL.Strings = (
      'insert into otvet'
      '  (N_OTV, N_VOPR, NAME_OTV, PR_PR, N_OTV_S, OTV_PIC)'
      'values'
      '  (:N_OTV, :N_VOPR, :NAME_OTV, :PR_PR, :N_OTV_S, :OTV_PIC)')
    DeleteSQL.Strings = (
      'delete from otvet'
      'where'
      '  N_OTV = :OLD_N_OTV')
    Left = 248
    Top = 48
  end
  object IBQueryTheme: TIBQuery
    Database = IBDatabaseImport
    Transaction = IBTransactionTheme
    BufferChunks = 1000
    CachedUpdates = True
    SQL.Strings = (
      'select * from tema')
    UpdateObject = IBUpdateSQLTheme
    GeneratorField.Field = 'N_TEMA'
    GeneratorField.Generator = 'NEW_TEMA'
    Left = 288
    Top = 48
    object IBQueryThemeN_TEMA: TIntegerField
      FieldName = 'N_TEMA'
      Origin = 'TEMA.N_TEMA'
      Required = True
    end
    object IBQueryThemeNAME_TEMA: TIBStringField
      FieldName = 'NAME_TEMA'
      Origin = 'TEMA.NAME_TEMA'
      Size = 255
    end
    object IBQueryThemeKOL_VOPR: TIntegerField
      FieldName = 'KOL_VOPR'
      Origin = 'TEMA.KOL_VOPR'
      Required = True
    end
  end
  object IBUpdateSQLTheme: TIBUpdateSQL
    RefreshSQL.Strings = (
      'Select '
      '  N_TEMA,'
      '  NAME_TEMA,'
      '  KOL_VOPR'
      'from tema '
      'where'
      '  N_TEMA = :N_TEMA')
    ModifySQL.Strings = (
      'update tema'
      'set'
      '  N_TEMA = :N_TEMA,'
      '  NAME_TEMA = :NAME_TEMA,'
      '  KOL_VOPR = :KOL_VOPR'
      'where'
      '  N_TEMA = :OLD_N_TEMA')
    InsertSQL.Strings = (
      'insert into tema'
      '  (N_TEMA, NAME_TEMA, KOL_VOPR)'
      'values'
      '  (:N_TEMA, :NAME_TEMA, :KOL_VOPR)')
    DeleteSQL.Strings = (
      'delete from tema'
      'where'
      '  N_TEMA = :OLD_N_TEMA')
    Left = 328
    Top = 48
  end
  object IBTransactionTheme: TIBTransaction
    Active = True
    DefaultDatabase = IBDatabaseImport
    AutoStopAction = saNone
    Left = 368
    Top = 48
  end
  object IBQueryStudents: TIBQuery
    BufferChunks = 1000
    CachedUpdates = False
    Left = 408
    Top = 48
  end
end
