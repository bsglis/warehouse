object Form1: TForm1
  Left = 191
  Top = 146
  Width = 870
  Height = 640
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object DBGrid1: TDBGrid
    Left = 216
    Top = 16
    Width = 513
    Height = 481
    DataSource = DataSource1
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
  end
  object Button1: TButton
    Left = 680
    Top = 544
    Width = 75
    Height = 25
    Caption = 'Request'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 336
    Top = 552
    Width = 75
    Height = 25
    Caption = 'Button2'
    TabOrder = 2
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 48
    Top = 528
    Width = 75
    Height = 25
    Caption = 'Button3'
    TabOrder = 3
    OnClick = Button3Click
  end
  object DataSource1: TDataSource
    DataSet = ADOQuery1
    Left = 32
    Top = 8
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=14790322;Persist Security Info=True' +
      ';User ID=sa;Initial Catalog=reg;Data Source=FS\SQLEXPRESS'
    LoginPrompt = False
    Provider = 'SQLOLEDB.1'
    Left = 32
    Top = 56
  end
  object ADOQuery1: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      ' SELECT [a].[ID'#1050#1086#1076#1052#1094']'
      '      ,[a].['#1040#1083#1100#1090#1077#1088#1053#1072#1080#1084']'
      '      ,[a].['#1058#1077#1082#1055#1088#1072#1081#1089#1062#1077#1085#1072']'
      '      ,[a].[KO]'
      '      ,[a].[CODE_POST]'
      '      ,[a].[CENA]'
      '      ,[a].[OPTCENA]'
      '      ,[a].[PRICE]'
      '      ,[a].[PRICE2]'
      '      ,[a].[uuser]'
      '      ,[a].[date_up]'
      '      ,[b].[id]'
      '      ,[b].['#1055#1088#1086#1080#1079#1074#1086#1076#1080#1090#1077#1083#1100']'
      '      ,[b].[Rem]'
      '      ,[c].[ID_NN]'
      '      ,[c].['#1050#1086#1076#1052#1062']'
      '      ,[d].['#1050#1086#1076#1057#1082#1083#1072#1076#1072']'
      '      ,[d].[ID_NN]'
      '      ,[d].[KOL0]'
      '      ,[d].[KOLTEK]'
      '      ,[d].[NamUpd]'
      '      ,[d].[DateUpd]'
      '      ,[e].[id'#1082#1086#1076#1084#1094']'
      '      ,[e].[id'#1082#1086#1076#1085#1072#1080#1084#1074#1085#1086#1088#1084#1077']'
      '      ,[e].[naim]'
      '      ,[e].['#1050#1088#1053#1072#1080#1084#1058#1086#1074']'
      '      '
      
        '  FROM [reg].[dbo].[26'#1061#1072#1088#1072#1082#1055#1088#1086#1076#1091#1082#1094#1080#1080'] a, [reg].[dbo].[lab_'#1089#1087#1088#1055#1088#1086 +
        #1080#1079#1074#1086#1076#1080#1090#1077#1083#1080'] b, [reg].[dbo].[00 MC_NN] c, [reg].[dbo].[00_'#1052#1057'] d, ' +
        '[reg].[dbo].[00'#1057#1087#1088#1048#1084#1077#1085#1052#1062'] e where [CODE_POST] = [b].[id] and [a]' +
        '.[ID'#1050#1086#1076#1052#1094'] = [c].['#1050#1086#1076#1052#1062'] and ['#1050#1086#1076#1057#1082#1083#1072#1076#1072'] = 13 and [c].[ID_NN] = ' +
        '[d].[ID_NN] and [a].[ID'#1050#1086#1076#1052#1094'] = [e].[ID'#1050#1086#1076#1052#1094']  order by  [e].[na' +
        'im]')
    Left = 32
    Top = 104
  end
end
