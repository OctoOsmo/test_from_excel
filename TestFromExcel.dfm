object FormMain: TFormMain
  Left = 849
  Top = 273
  Width = 461
  Height = 259
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
    Top = 131
    Width = 445
    Height = 89
    Align = alBottom
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssVertical
    TabOrder = 1
  end
  object Button1: TButton
    Left = 128
    Top = 88
    Width = 75
    Height = 25
    Caption = 'Button1'
    TabOrder = 2
    OnClick = Button1Click
  end
  object IBQueryImport: TIBQuery
    Database = IBDatabaseImport
    Transaction = IBTransactionImport
    BufferChunks = 1000
    CachedUpdates = False
    Left = 328
    Top = 8
  end
  object IBDatabaseImport: TIBDatabase
    Connected = True
    DatabaseName = 'D:\'#1055#1088#1086#1075#1088#1072#1084#1084#1099'\'#1058#1077#1089#1090#1099'\'#1041#1072#1079#1099'\TEST_VSMA.GDB'
    Params.Strings = (
      'user_name=sysdba'
      'password=masterkey'
      'lc_ctype=WIN1251')
    LoginPrompt = False
    DefaultTransaction = IBTransactionImport
    IdleTimer = 0
    SQLDialect = 1
    TraceFlags = []
    Left = 368
    Top = 8
  end
  object IBTransactionImport: TIBTransaction
    Active = True
    DefaultDatabase = IBDatabaseImport
    AutoStopAction = saNone
    Left = 408
    Top = 8
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
      
        '  (N_VOPR, N_TEMA, NAME_VOPR_L, NAME_VOPR_P, TIP_VOPR, KOL_OTV, ' +
        'VOPR_PIC)'
      'values'
      
        '  (:N_VOPR, :N_TEMA, :NAME_VOPR_L, :NAME_VOPR_P, :TIP_VOPR, :KOL' +
        '_OTV, :VOPR_PIC)')
    DeleteSQL.Strings = (
      'delete from vopros'
      'where'
      '  N_VOPR = :OLD_N_VOPR')
    Left = 368
    Top = 56
  end
  object IBQueryQuestion: TIBQuery
    Database = IBDatabaseImport
    Transaction = IBTransactionQuestion
    BufferChunks = 1000
    CachedUpdates = True
    SQL.Strings = (
      'select * from vopros')
    UpdateObject = IBUpdateSQLQuestion
    GeneratorField.Field = 'N_VOPR'
    GeneratorField.Generator = 'NEW_VOPR'
    Left = 328
    Top = 56
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
  object DataSourceQuestion: TDataSource
    DataSet = IBQueryQuestion
    Left = 288
    Top = 56
  end
  object IBTransactionQuestion: TIBTransaction
    Active = True
    DefaultDatabase = IBDatabaseImport
    AutoStopAction = saNone
    Left = 408
    Top = 56
  end
end
