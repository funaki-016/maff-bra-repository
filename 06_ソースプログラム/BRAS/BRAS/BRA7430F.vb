Imports Microsoft.VisualBasic.FileIO
Imports System.Text
Imports System.Text.RegularExpressions

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.12.16 |Daiko               | 要件No1 バージョン区分追加
'//  REV_002   | 2023.08.11 |Daiko               | 変更要件No.5 設定可能項番の追加、営農の項番存在チェックのバージョンを無視
'//  REV_003   | 2023.12.15 |Daiko               | 既存不具合修正（設定読込時に画面項目のクリア処理を追加）
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 帳票出力画面
''' </summary>
''' <remarks></remarks>
Public Class BRA7430F

    '===============================
    Private _censusNo As String

    '任意項目番号定義
    Private Enum 営農個人_項目番号
        お問合せ先 = 1
        集計結果表
        営個還2_レーダーチャート
        営個還3_分析指標の経年変化
        営個還4_グラフ項目
        営個還7_生産概況内訳
        営個還8_分析指標
        営個還9_統計表
    End Enum

    Private Enum 営農法人_項目番号
        お問合せ先 = 1
        集計結果表
        営法還2_レーダーチャート
        営法還3_分析指標の経年変化
        営法還4_グラフ項目
        営法還7_生産概況内訳
        営法還8_分析指標
        営法還9_統計表
    End Enum

    Private Enum 生産費_項目番号
        お問合せ先 = 1
        集計結果表
        還元4_統計表
        還元6_レーダーチャートの表示費目
        還元8_主要費目
        還元9_主要作業
        還元10_物財費内訳
    End Enum

    '任意項目明細番号定義(共通)
    Private Enum お問合せ先
        局_事務所_事務局 = 1
        県拠点等
        担当者
        問合せ番号先
        FAX
        住所
    End Enum

    Private Enum 集計結果表
        比較データ1 = 1
        比較データ2
        比較データ3
        比較データ4
    End Enum

    'CSV区切文字
    Private Const CSV_DELIMITER As String = ","
    'コードページ_Shift_JIS
    Private Const CODEPAGE_SHIFT_JIS As String = "Shift_JIS"

    'CSV項目定義
    Private Enum ヘッダ
        ファイルタイトル
        調査区分
    End Enum

    Private Enum データ部
        項目名称
        項目番号
        明細番号
        設定値
    End Enum

    Private Const ファイルタイトル As String = "還元資料パラメータ"

#Region "個別結果表項番チェック定義"

#Region "営農個人"
    Private Shared ALLOWLIST_EIKO2 As New List(Of String) From {
        "K011118", "K011119", "K011120", "K011121", "K011122", "K011123", "K011124", "K011125", "K011126", "K011127",
        "K011128", "K011129", "K011130", "K011131", "K011132", "K011133", "K011139"
    }
    'REV_002↓
    'Private Shared ALLOWLIST_EIKO3 As New List(Of String) From {
    '    "K010509", "K010510", "K010511", "K010512", "K010513", "K010514", "K010515", "K010516", "K010517", "K010518",
    '    "K010519", "K010521", "K010522", "K010523", "K010524", "K010525", "K010526", "K010527", "K010528", "K010529",
    '    "K010530", "K010531", "K010532", "K010533", "K010534", "K010535", "K010536", "K010538", "K010539", "K010540",
    '    "K010541", "K010542", "K010543", "K010544", "K010545", "K010547", "K010548", "K010549"
    '}
    Private Shared ALLOWLIST_EIKO3 As New List(Of String) From {
        "K010509", "K010510", "K010511", "K010512", "K010513", "K010514", "K010515", "K010516", "K010517", "K010518",
        "K010519", "K010521", "K010522", "K010523", "K010524", "K010525", "K010526", "K010527", "K010528", "K010529",
        "K010530", "K010531", "K010532", "K010533", "K010534", "K010535", "K010536", "K010538", "K010539", "K010540",
        "K010541", "K010542", "K010543", "K010544", "K010545", "K010547", "K010548", "K010549", "K010550", "K010551"
    }
    'REV_002↑

    'REV_002↓
    'Private Shared ALLOWLIST_EIKO4_1 As New List(Of String) From {
    '    "K021116", "K010741", "K021437"
    '}
    Private Shared ALLOWLIST_EIKO4_1 As New List(Of String) From {
        "K021116", "K010741", "K021437", "K020381", "K020382", "K020383", "K020384", "K020385", "K021239", "K021240",
        "K021241", "K021243", "K021244", "K021245"
    }
    'REV_002↑
    Private Shared ALLOWRANGE_EIKO4_1_1 As New List(Of String) From {"K020301", "K020347"}
    Private Shared ALLOWRANGE_EIKO4_1_2 As New List(Of String) From {"K020601", "K020644"}

    'REV_002↓
    'Private Shared ALLOWLIST_EIKO4_2 As New List(Of String) From {
    '    "K010210"
    '}
    Private Shared ALLOWLIST_EIKO4_2 As New List(Of String) From {
        "K010210", "K020981", "K020982", "K020983", "K020984", "K020944"
    }
    'REV_002↑
    Private Shared ALLOWRANGE_EIKO4_2_1 As New List(Of String) From {"K020902", "K020943"}
    Private Shared ALLOWRANGE_EIKO4_2_2 As New List(Of String) From {"K021202", "K021220"}

    'REV_002↓
    'Private Shared ALLOWLIST_EIKO8 As New List(Of String) From {
    '    "K010509", "K010510", "K010511", "K010512", "K010513", "K010514", "K010515", "K010516", "K010517", "K010518",
    '    "K010519", "K010521", "K010522", "K010523", "K010524", "K010525", "K010532", "K010533", "K010538", "K010539",
    '    "K010540", "K010541", "K010542", "K010543", "K010544", "K010545", "K010547", "K010548", "K010549"
    '}
    Private Shared ALLOWLIST_EIKO8 As New List(Of String) From {
        "K010509", "K010510", "K010511", "K010512", "K010513", "K010514", "K010515", "K010516", "K010517", "K010518",
        "K010519", "K010521", "K010522", "K010523", "K010524", "K010525", "K010532", "K010533", "K010538", "K010539",
        "K010540", "K010541", "K010542", "K010543", "K010544", "K010545", "K010547", "K010548", "K010549", "K010550",
        "K010551"
    }
    'REV_002↑
#End Region

#Region "営農法人"
    Private Shared ALLOWLIST_EIHO2 As New List(Of String) From {
        "K021231", "K021232", "K021233", "K021234", "K021235", "K021236", "K021237", "K021238", "K021239", "K021240", "K021241",
        "K021242", "K021243", "K021244", "K021245", "K021246", "K021249"
    }
    'REV_002↓
    'Private Shared ALLOWLIST_EIHO3 As New List(Of String) From {
    '    "K010507", "K010508", "K010509", "K010510", "K010511", "K010512", "K010513", "K010514", "K010515", "K010516", "K010517",
    '    "K010518", "K010519", "K010520", "K010521", "K010522", "K010523", "K010525", "K010526", "K010527", "K010528", "K010529",
    '    "K010530", "K010531", "K010532", "K010533", "K010534", "K010535", "K010536", "K010537", "K010538", "K010539", "K010540",
    '    "K010542", "K010543", "K010544", "K010545", "K010546", "K010547", "K010548", "K010549", "K010551", "K010552", "K010553"
    '}
    Private Shared ALLOWLIST_EIHO3 As New List(Of String) From {
        "K010507", "K010508", "K010509", "K010510", "K010511", "K010512", "K010513", "K010514", "K010515", "K010516", "K010517",
        "K010518", "K010519", "K010520", "K010521", "K010522", "K010523", "K010525", "K010526", "K010527", "K010528", "K010529",
        "K010530", "K010531", "K010532", "K010533", "K010534", "K010535", "K010536", "K010537", "K010538", "K010539", "K010540",
        "K010542", "K010543", "K010544", "K010545", "K010546", "K010547", "K010548", "K010549", "K010551", "K010552", "K010553",
        "K010554", "K010555"
    }
    'REV_002↑
    'REV_002↓
    'Private Shared ALLOWLIST_EIHO4_1 As New List(Of String) From {
    '    "K021436", "K011236"
    '}
    Private Shared ALLOWLIST_EIHO4_1 As New List(Of String) From {
        "K021436", "K011236", "K020381", "K020382", "K020383", "K020384", "K020385", "K021448", "K021449", "K021450", "K021452",
        "K021453", "K021454"
    }
    'REV_002↑
    Private Shared ALLOWRANGE_EIHO4_1_1 As New List(Of String) From {"K020301", "K020346"}
    Private Shared ALLOWRANGE_EIHO4_1_2 As New List(Of String) From {"K020601", "K020646"}
    'REV_002↓
    'Private Shared ALLOWLIST_EIHO4_2 As New List(Of String) From {
    '    "K010833"
    '}
    Private Shared ALLOWLIST_EIHO4_2 As New List(Of String) From {
        "K010833", "K020981", "K020982", "K020983", "K020984", "K020944"
    }
    'REV_002↑
    Private Shared ALLOWRANGE_EIHO4_2_1 As New List(Of String) From {"K020902", "K020943"}
    Private Shared ALLOWRANGE_EIHO4_2_2 As New List(Of String) From {"K021202", "K021224"}
    'REV_002↓
    'Private Shared ALLOWLIST_EIHO8 As New List(Of String) From {
    '    "K010507", "K010508", "K010509", "K010510", "K010511", "K010512", "K010513", "K010514", "K010515", "K010516", "K010517",
    '    "K010518", "K010519", "K010520", "K010521", "K010522", "K010523", "K010525", "K010526", "K010527", "K010528", "K010529",
    '    "K010536", "K010537", "K010542", "K010543", "K010544", "K010545", "K010546", "K010547", "K010548", "K010549", "K010551",
    '    "K010552", "K010553"
    '}
    Private Shared ALLOWLIST_EIHO8 As New List(Of String) From {
        "K010507", "K010508", "K010509", "K010510", "K010511", "K010512", "K010513", "K010514", "K010515", "K010516", "K010517",
        "K010518", "K010519", "K010520", "K010521", "K010522", "K010523", "K010525", "K010526", "K010527", "K010528", "K010529",
        "K010536", "K010537", "K010542", "K010543", "K010544", "K010545", "K010546", "K010547", "K010548", "K010549", "K010551",
        "K010552", "K010553", "K010554", "K010555"
    }
    'REV_002↑
#End Region

#Region "生産費"
#Region "米生産費（個別）"
    Private Shared ALLOWLIST_SEISAN8_1_KOME_KO As New List(Of String) From {
        "K010706", "K010707", "K010708", "K010710", "K010711", "K010719", "K010720",
        "K010721", "K010722", "K010806", "K010815", "K010816", "K010817", "K010818", "K010822"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_KOME_KO As New List(Of List(Of String)) From {
        New List(Of String) From {"K010606", "K010622"},
        New List(Of String) From {"K010906", "K010930"}
    }

    Private Shared ALLOWLIST_SEISAN8_2_KOME_KO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_KOME_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010934", "K010946"}
    }

    Private Shared ALLOWLIST_SEISAN10_KOME_KO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN10_KOME_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010907", "K010918"}
    }
#End Region

#Region "麦類・大豆・そば・なたね・畑作物生産費（個別）"
    Private Shared ALLOWLIST_SEISAN8_1_HATAKE_KO As New List(Of String) From {
        "K010708", "K010709", "K010710", "K010712", "K010713", "K010721", "K010722", "K010723", "K010724",
        "K010808", "K010817", "K010818", "K010819", "K010820", "K010824"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_HATAKE_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010608", "K010624"},
         New List(Of String) From {"K010908", "K010932"}
    }
    Private Shared ALLOWLIST_SEISAN8_2_HATAKE_KO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_HATAKE_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010936", "K010951"}
    }
    Private Shared ALLOWLIST_SEISAN10_HATAKE_KO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN10_HATAKE_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010909", "K010920"}
    }
#End Region

#Region "牛乳生産費（個別）"
    Private Shared ALLOWLIST_SEISAN8_1_GYUNYU_KO As New List(Of String) From {
        "K010609", "K010610", "K010611", "K010612", "K010614", "K010615", "K010616", "K010617", "K010618", "K010619",
        "K010621", "K010622", "K010623", "K010624", "K010625", "K010626", "K010627", "K010628",
        "K010709", "K010710", "K010711", "K010712", "K010713", "K010714", "K010715", "K010716", "K010725", "K010726",
        "K010727", "K010728",
        "K010809", "K010820", "K010821", "K010822", "K010823", "K010824", "K010828"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_GYUNYU_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010909", "K010936"}
    }
    Private Shared ALLOWLIST_SEISAN8_2_GYUNYU_KO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_GYUNYU_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K020907", "K020910"}
    }
    Private Shared ALLOWLIST_SEISAN10_GYUNYU_KO As New List(Of String) From {
        "K010910", "K010911", "K010914", "K010915", "K010916", "K010917", "K010918", "K010919", "K010920", "K010921",
        "K010922", "K010923", "K010924"
    }
    Private Shared ALLOWRANGE_SEISAN10_GYUNYU_KO As New List(Of List(Of String)) From {
    }
#End Region
#Region "子牛生産費（個別）"
    Private Shared ALLOWLIST_SEISAN8_1_KOUSHI_KO As New List(Of String) From {
        "K010608", "K010609", "K010610", "K010611", "K010613", "K010614", "K010615", "K010616", "K010617", "K010618",
        "K010620", "K010621", "K010622", "K010623", "K010624", "K010625", "K010626", "K010627",
        "K010708", "K010709", "K010710", "K010711", "K010712", "K010713", "K010714", "K010715", "K010724", "K010725",
        "K010726", "K010727",
        "K010808", "K010819", "K010820", "K010821", "K010822", "K010823", "K010827"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_KOUSHI_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010908", "K010935"}
    }
    Private Shared ALLOWLIST_SEISAN8_2_KOUSHI_KO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_KOUSHI_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K020907", "K020909"}
    }
    Private Shared ALLOWLIST_SEISAN10_KOUSHI_KO As New List(Of String) From {
        "K010909", "K010910", "K010913", "K010914", "K010915", "K010916", "K010917", "K010918", "K010919", "K010920",
        "K010921", "K010922", "K010923"
    }
    Private Shared ALLOWRANGE_SEISAN10_KOUSHI_KO As New List(Of List(Of String)) From {
    }
#End Region
#Region "育成牛・肥育牛生産費（個別）"
    Private Shared ALLOWLIST_SEISAN8_1_NIKUGYU_KO As New List(Of String) From {
        "K010608", "K010609", "K010610", "K010611", "K010613", "K010614", "K010615", "K010616", "K010617", "K010618",
        "K010619", "K010620", "K010621", "K010622", "K010623", "K010624", "K010625", "K010626",
        "K010708", "K010709", "K010710", "K010711", "K010712", "K010713", "K010714", "K010715", "K010723", "K010724",
        "K010725", "K010726",
        "K010808", "K010819", "K010820", "K010821", "K010822", "K010826"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_NIKUGYU_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010908", "K010934"}
    }
    Private Shared ALLOWLIST_SEISAN8_2_NIKUGYU_KO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_NIKUGYU_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K020907", "K020909"}
    }
    Private Shared ALLOWLIST_SEISAN10_NIKUGYU_KO As New List(Of String) From {
        "K010909", "K010910", "K010913", "K010914", "K010915", "K010916", "K010917", "K010918", "K010919", "K010920",
        "K010921", "K010922"
    }
    Private Shared ALLOWRANGE_SEISAN10_NIKUGYU_KO As New List(Of List(Of String)) From {
    }
#End Region
#Region "肥育豚生産費（個別）"
    Private Shared ALLOWLIST_SEISAN8_1_BUTA_KO As New List(Of String) From {
        "K010609", "K010610", "K010611", "K010612", "K010613", "K010615", "K010616", "K010617", "K010618", "K010619",
        "K010620", "K010621", "K010622", "K010623", "K010624", "K010625", "K010626", "K010627", "K010628", "K010629",
        "K010630",
        "K010709", "K010712", "K010713", "K010714", "K010715", "K010716", "K010717", "K010727", "K010728", "K010729",
        "K010730",
        "K010809", "K010823", "K010824", "K010825", "K010826", "K010830"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_BUTA_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010909", "K010938"}
    }
    Private Shared ALLOWLIST_SEISAN8_2_BUTA_KO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_BUTA_KO As New List(Of List(Of String)) From {
         New List(Of String) From {"K021107", "K021109"}
    }
    Private Shared ALLOWLIST_SEISAN10_BUTA_KO As New List(Of String) From {
        "K010910", "K010911", "K010912", "K010915", "K010916", "K010917", "K010918", "K010919", "K010920", "K010921",
        "K010922", "K010923", "K010924", "K010925", "K010926"
    }
    Private Shared ALLOWRANGE_SEISAN10_BUTA_KO As New List(Of List(Of String)) From {
    }
#End Region
#Region "米生産費（組織法人）"
    Private Shared ALLOWLIST_SEISAN8_1_KOME_HO As New List(Of String) From {
        "K010708", "K010709", "K010710", "K010712", "K010713", "K010721", "K010722", "K010723", "K010724",
        "K010808", "K010817", "K010818", "K010819", "K010820", "K010824"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_KOME_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010608", "K010624"},
         New List(Of String) From {"K010908", "K010932"}
    }
    Private Shared ALLOWLIST_SEISAN8_2_KOME_HO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_KOME_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010535", "K010547"}
    }
    Private Shared ALLOWLIST_SEISAN10_KOME_HO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN10_KOME_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010909", "K010920"}
    }
#End Region
#Region "小麦・大豆生産費（組織法人）"
    Private Shared ALLOWLIST_SEISAN8_1_KOMUGI_HO As New List(Of String) From {
        "K010708", "K010709", "K010710", "K010712", "K010713", "K010721", "K010722", "K010723", "K010724",
        "K010808", "K010817", "K010818", "K010819", "K010820", "K010824"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_KOMUGI_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010608", "K010624"},
         New List(Of String) From {"K010908", "K010932"}
    }
    Private Shared ALLOWLIST_SEISAN8_2_KOMUGI_HO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_KOMUGI_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010536", "K010549"}
    }
    Private Shared ALLOWLIST_SEISAN10_KOMUGI_HO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN10_KOMUGI_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010909", "K010920"}
    }
#End Region
#Region "経営分析_麦類・そば・なたね・畑作物生産費"
    Private Shared ALLOWLIST_SEISAN8_1_HATAKE_HO As New List(Of String) From {
        "K010708", "K010709", "K010710", "K010712", "K010713", "K010721", "K010722", "K010723", "K010724",
        "K010808", "K010817", "K010818", "K010819", "K010820", "K010824"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_HATAKE_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010608", "K010624"},
         New List(Of String) From {"K010908", "K010932"}
    }
    Private Shared ALLOWLIST_SEISAN8_2_HATAKE_HO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_HATAKE_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010536", "K010551"}
    }
    Private Shared ALLOWLIST_SEISAN10_HATAKE_HO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN10_HATAKE_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010909", "K010920"}
    }
#End Region
#Region "経営分析_牛乳生産費"
    Private Shared ALLOWLIST_SEISAN8_1_GYUNYU_HO As New List(Of String) From {
        "K010609", "K010610", "K010611", "K010612", "K010614", "K010615", "K010616", "K010617", "K010618", "K010619",
        "K010621", "K010622", "K010623", "K010624", "K010625", "K010626", "K010627", "K010628",
        "K010709", "K010710", "K010711", "K010712", "K010713", "K010714", "K010715", "K010716", "K010725", "K010726",
        "K010727", "K010728",
        "K010809", "K010820", "K010821", "K010822", "K010823", "K010824", "K010828"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_GYUNYU_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010909", "K010936"}
    }
    Private Shared ALLOWLIST_SEISAN8_2_GYUNYU_HO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_GYUNYU_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K020507", "K020510"}
    }
    Private Shared ALLOWLIST_SEISAN10_GYUNYU_HO As New List(Of String) From {
        "K010910", "K010911", "K010914", "K010915", "K010916", "K010917", "K010918", "K010919", "K010920", "K010921",
        "K010922", "K010923", "K010924"
    }
    Private Shared ALLOWRANGE_SEISAN10_GYUNYU_HO As New List(Of List(Of String)) From {
    }
#End Region
#Region "経営分析_子牛生産費"
    Private Shared ALLOWLIST_SEISAN8_1_KOUSHI_HO As New List(Of String) From {
        "K010608", "K010609", "K010610", "K010611", "K010613", "K010614", "K010615", "K010616", "K010617", "K010618", "K010620",
        "K010621", "K010622", "K010623", "K010624", "K010625", "K010626", "K010627",
        "K010708", "K010709", "K010710", "K010711", "K010712", "K010713", "K010714", "K010715", "K010724", "K010725", "K010726",
        "K010727",
        "K010808", "K010819", "K010820", "K010821", "K010822", "K010823", "K010827"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_KOUSHI_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010908", "K010935"}
    }
    Private Shared ALLOWLIST_SEISAN8_2_KOUSHI_HO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_KOUSHI_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K020507", "K020509"}
    }
    Private Shared ALLOWLIST_SEISAN10_KOUSHI_HO As New List(Of String) From {
        "K010909", "K010910", "K010913", "K010914", "K010915", "K010916", "K010917", "K010918", "K010919", "K010920",
        "K010921", "K010922", "K010923"
    }
    Private Shared ALLOWRANGE_SEISAN10_KOUSHI_HO As New List(Of List(Of String)) From {
    }
#End Region
#Region "経営分析_育成牛・肥育牛生産費"
    Private Shared ALLOWLIST_SEISAN8_1_NIKUGYU_HO As New List(Of String) From {
        "K010608", "K010609", "K010610", "K010611", "K010613", "K010614", "K010615", "K010616", "K010617", "K010618",
        "K010619", "K010620", "K010621", "K010622", "K010623", "K010624", "K010625", "K010626",
        "K010708", "K010709", "K010710", "K010711", "K010712", "K010713", "K010714", "K010715", "K010723", "K010724",
        "K010725", "K010726",
        "K010808", "K010819", "K010820", "K010821", "K010822", "K010826"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_NIKUGYU_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010908", "K010934"}
    }
    Private Shared ALLOWLIST_SEISAN8_2_NIKUGYU_HO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_NIKUGYU_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K020507", "K020509"}
    }
    Private Shared ALLOWLIST_SEISAN10_NIKUGYU_HO As New List(Of String) From {
        "K010909", "K010910", "K010913", "K010914", "K010915", "K010916", "K010917", "K010918", "K010919", "K010920",
        "K010921", "K010922"
    }
    Private Shared ALLOWRANGE_SEISAN10_NIKUGYU_HO As New List(Of List(Of String)) From {
    }
#End Region
#Region "経営分析_肥育豚生産費"
    Private Shared ALLOWLIST_SEISAN8_1_BUTA_HO As New List(Of String) From {
        "K010609", "K010610", "K010611", "K010612", "K010613", "K010615", "K010616", "K010617", "K010618", "K010619", "K010620",
        "K010621", "K010622", "K010623", "K010624", "K010625", "K010626", "K010627", "K010628", "K010629", "K010630",
        "K010709", "K010712", "K010713", "K010714", "K010715", "K010716", "K010717", "K010727", "K010728", "K010729", "K010730",
        "K010809", "K010823", "K010824", "K010825", "K010826", "K010830"
    }
    Private Shared ALLOWRANGE_SEISAN8_1_BUTA_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K010909", "K010938"}
    }
    Private Shared ALLOWLIST_SEISAN8_2_BUTA_HO As New List(Of String) From {
    }
    Private Shared ALLOWRANGE_SEISAN8_2_BUTA_HO As New List(Of List(Of String)) From {
         New List(Of String) From {"K020507", "K020509"}
    }
    Private Shared ALLOWLIST_SEISAN10_BUTA_HO As New List(Of String) From {
        "K010910", "K010911", "K010912", "K010915", "K010916", "K010917", "K010918", "K010919", "K010920", "K010921", "K010922",
        "K010923", "K010924", "K010925", "K010926"
    }
    Private Shared ALLOWRANGE_SEISAN10_BUTA_HO As New List(Of List(Of String)) From {
    }
#End Region

#End Region

#End Region

    Private Const DEFAULT_CSNSUS As String = "0000000000000000"

    ''' <summary>調査年</summary>
    Private _chosaNen As String
    ''' <summary>局</summary>
    Private _kyoku As String
    ''' <summary>事務所</summary>
    Private _jimusho As String
    ''' <summary>拠点</summary>
    Private _kyoten As String
    ''' <summary>営農類型</summary>
    Private _einouRuikei As String
    ''' <summary>欠測値補完</summary>
    Private _kessokuchiHokan As String
    ''' <summary>貸借対照表</summary>
    Private _taishakuTaishohyo As String
    ''' <summary>規模階層</summary>
    Private _kiboKaiso As Integer

    ''' <summary>当年集計結果表テーブル</summary>
    Private shukeiTableThisYear As DataTable
    ''' <summary>前年集計結果表テーブル</summary>
    Private shukeiTablePrevYear As DataTable
    ''' <summary>還元資料設定値テーブルレイアウト</summary>
    Private settingTableLayout As DataTable
    ''' <summary>個別結果表マスタテーブル</summary>
    Private kobetsuItemTable As DataTable

    ''' <summary>「使用する集計結果表」コンボボックス配列</summary>
    Private shukeiComboArray(3, 2) As ComboBox

    ''' <summary>生産費チェックリスト</summary>
    Private seisan8_1CheckList As List(Of String)
    ''' <summary>生産費チェックリスト</summary>
    Private seisan8_1CheckRangeList As List(Of List(Of String))
    ''' <summary>生産費チェックリスト</summary>
    Private seisan8_2CheckList As List(Of String)
    ''' <summary>生産費チェックリスト</summary>
    Private seisan8_2CheckRangeList As List(Of List(Of String))
    ''' <summary>生産費チェックリスト</summary>
    Private seisan10CheckList As List(Of String)
    ''' <summary>生産費チェックリスト</summary>
    Private seisan10CheckRangeList As List(Of List(Of String))

#Region "【処理詳細仕様 1】初期表示"
    Public Sub New(pCensusNo As String, chosaNen As String, einouRuikei As String, kiboKaiso As Integer)

        InitializeComponent()   '呼び出しの後で初期化

        _censusNo = pCensusNo

        _chosaNen = chosaNen
        _einouRuikei = einouRuikei
        _kiboKaiso = kiboKaiso

        shukeiComboArray(0, 0) = cboShukeiData1
        shukeiComboArray(0, 1) = cboChiki1
        shukeiComboArray(0, 2) = cboKibokaiso1
        shukeiComboArray(1, 0) = cboShukeiData2
        shukeiComboArray(1, 1) = cboChiki2
        shukeiComboArray(1, 2) = cboKibokaiso2
        shukeiComboArray(2, 0) = cboShukeiData3
        shukeiComboArray(2, 1) = cboChiki3
        shukeiComboArray(2, 2) = cboKibokaiso3
        shukeiComboArray(3, 0) = cboShukeiDataZen
        shukeiComboArray(3, 1) = cboChikiZen
        shukeiComboArray(3, 2) = cboKibokaisoZen

        Select Case CommonInfo.Chosakubun
            Case ComConst.調査区分.米生産費統計_個別
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_KOME_KO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_KOME_KO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_KOME_KO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_KOME_KO
                seisan10CheckList = ALLOWLIST_SEISAN10_KOME_KO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_KOME_KO

            Case ComConst.調査区分.小麦生産費統計_個別,
                ComConst.調査区分.二条大麦生産費統計_個別,
                ComConst.調査区分.六条大麦生産費統計_個別,
                ComConst.調査区分.はだか麦生産費統計_個別,
                ComConst.調査区分.そば生産費統計_個別,
                ComConst.調査区分.大豆生産費統計_個別,
                ComConst.調査区分.原料用かんしょ生産費統計_個別,
                ComConst.調査区分.原料用ばれいしょ生産費統計_個別,
                ComConst.調査区分.なたね生産費統計_個別,
                ComConst.調査区分.てんさい生産費統計_個別,
                ComConst.調査区分.さとうきび生産費統計_個別
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_HATAKE_KO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_HATAKE_KO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_HATAKE_KO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_HATAKE_KO
                seisan10CheckList = ALLOWLIST_SEISAN10_HATAKE_KO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_HATAKE_KO

            Case ComConst.調査区分.牛乳生産費統計_個別
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_GYUNYU_KO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_GYUNYU_KO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_GYUNYU_KO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_GYUNYU_KO
                seisan10CheckList = ALLOWLIST_SEISAN10_GYUNYU_KO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_GYUNYU_KO

            Case ComConst.調査区分.子牛生産費統計_個別
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_KOUSHI_KO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_KOUSHI_KO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_KOUSHI_KO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_KOUSHI_KO
                seisan10CheckList = ALLOWLIST_SEISAN10_KOUSHI_KO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_KOUSHI_KO

            Case ComConst.調査区分.乳用雄育成牛生産費統計_個別,
                ComConst.調査区分.交雑種育成牛生産費統計_個別,
                ComConst.調査区分.去勢若齢肥育牛生産費統計_個別,
                ComConst.調査区分.乳用雄肥育牛生産費統計_個別,
                ComConst.調査区分.交雑種肥育牛生産費統計_個別
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_NIKUGYU_KO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_NIKUGYU_KO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_NIKUGYU_KO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_NIKUGYU_KO
                seisan10CheckList = ALLOWLIST_SEISAN10_NIKUGYU_KO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_NIKUGYU_KO

            Case ComConst.調査区分.肥育豚生産費統計_個別
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_BUTA_KO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_BUTA_KO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_BUTA_KO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_BUTA_KO
                seisan10CheckList = ALLOWLIST_SEISAN10_BUTA_KO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_BUTA_KO

            Case ComConst.調査区分.米生産費統計_組織法人
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_KOME_HO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_KOME_HO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_KOME_HO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_KOME_HO
                seisan10CheckList = ALLOWLIST_SEISAN10_KOME_HO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_KOME_HO

            Case ComConst.調査区分.小麦生産費統計_組織法人,
                ComConst.調査区分.大豆生産費統計_組織法人
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_KOMUGI_HO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_KOMUGI_HO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_KOMUGI_HO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_KOMUGI_HO
                seisan10CheckList = ALLOWLIST_SEISAN10_KOMUGI_HO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_KOMUGI_HO

            Case ComConst.調査区分.経営分析調査_二条大麦生産費,
                ComConst.調査区分.経営分析調査_六条大麦生産費,
                ComConst.調査区分.経営分析調査_はだか麦生産費,
                ComConst.調査区分.経営分析調査_そば生産費,
                ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費,
                ComConst.調査区分.経営分析調査_なたね生産費,
                ComConst.調査区分.経営分析調査_てんさい生産費,
                ComConst.調査区分.経営分析調査_さとうきび生産費
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_HATAKE_HO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_HATAKE_HO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_HATAKE_HO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_HATAKE_HO
                seisan10CheckList = ALLOWLIST_SEISAN10_HATAKE_HO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_HATAKE_HO

            Case ComConst.調査区分.経営分析調査_牛乳生産費
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_GYUNYU_HO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_GYUNYU_HO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_GYUNYU_HO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_GYUNYU_HO
                seisan10CheckList = ALLOWLIST_SEISAN10_GYUNYU_HO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_GYUNYU_HO

            Case ComConst.調査区分.経営分析調査_子牛生産費
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_KOUSHI_HO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_KOUSHI_HO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_KOUSHI_HO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_KOUSHI_HO
                seisan10CheckList = ALLOWLIST_SEISAN10_KOUSHI_HO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_KOUSHI_HO

            Case ComConst.調査区分.経営分析調査_乳用雄育成牛生産費,
                ComConst.調査区分.経営分析調査_交雑種育成牛生産費,
                ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費,
                ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費,
                ComConst.調査区分.経営分析調査_交雑種肥育牛生産費
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_NIKUGYU_HO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_NIKUGYU_HO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_NIKUGYU_HO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_NIKUGYU_HO
                seisan10CheckList = ALLOWLIST_SEISAN10_NIKUGYU_HO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_NIKUGYU_HO

            Case ComConst.調査区分.経営分析調査_肥育豚生産費
                seisan8_1CheckList = ALLOWLIST_SEISAN8_1_BUTA_HO
                seisan8_1CheckRangeList = ALLOWRANGE_SEISAN8_1_BUTA_HO
                seisan8_2CheckList = ALLOWLIST_SEISAN8_2_BUTA_HO
                seisan8_2CheckRangeList = ALLOWRANGE_SEISAN8_2_BUTA_HO
                seisan10CheckList = ALLOWLIST_SEISAN10_BUTA_HO
                seisan10CheckRangeList = ALLOWRANGE_SEISAN10_BUTA_HO

        End Select
    End Sub

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA7430F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Dim dt As DataTable
            Dim seisanFlg As Boolean = True

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                dt = DAOOther.GetKangenShiryoSetting(db, _censusNo)
                If dt.Rows.Count = 0 Then
                    Dim einouRuikei As Integer
                    Integer.TryParse(_einouRuikei, einouRuikei)

                    dt = DAOOther.GetKangenShiryoSetting(db, DEFAULT_CSNSUS, (CInt(CommonInfo.Chosakubun) + einouRuikei * 1000).ToString)
                End If

                'REV_002↓営農はバージョン区分を無視
                '' REV_001↓
                ''kobetsuItemTable = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun)
                'kobetsuItemTable = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
                '' REV_001↑
                If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 _
                    OrElse CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                    kobetsuItemTable = New DataTable()
                    For Each versionKbn In ComConst.バージョン区分.体系リスト.Keys
                        kobetsuItemTable.Merge(DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, versionKbn))
                    Next
                Else
                    kobetsuItemTable = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
                End If
                'REV_002↑
            End Using

            '設定読み出しに使用するため、テーブルレイアウトのみ保持
            settingTableLayout = dt.Clone()

            '「お取合せ先」の設定
            SetOtoiawaseSetting(dt)

            '「集計結果表」の設定
            InitShukeiKekkahyoComboBox()
            SetShukeiKekkahyoSetting(dt)

            '「個別結果表」の表示設定
            '調査区分が1(営農累計別経営統計(個人)の場合、下記を行う
            If (CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人) Then
                GroupEinoukojin.Visible = True
                GroupEinouhoujin.Visible = False
                GroupSeisanhi.Visible = False
            ElseIf (CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人) Then
                GroupEinoukojin.Visible = False
                GroupEinouhoujin.Visible = True
                GroupSeisanhi.Visible = False
            Else
                GroupEinoukojin.Visible = False
                GroupEinouhoujin.Visible = False
                GroupSeisanhi.Visible = True
            End If
            txtChohyoMei.Text = ComConst.還元資料.帳票名(CommonInfo.Chosakubun)

            SetKobetsuKekkahyoSetting(dt)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region

#Region "【処理詳細仕様 2】「集計結果表データ」選択"

#Region "【処理詳細仕様 2】「集計結果表データ」コンボボックス選択"
    Private Sub cboShukeiData1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboShukeiData1.SelectedIndexChanged
        ShukeiDataComboChanged(0)
    End Sub

    Private Sub cboShukeiData2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboShukeiData2.SelectedIndexChanged
        ShukeiDataComboChanged(1)
    End Sub

    Private Sub cboShukeiData3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboShukeiData3.SelectedIndexChanged
        ShukeiDataComboChanged(2)
    End Sub

    Private Sub cboShukeiDataZen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboShukeiDataZen.SelectedIndexChanged
        ShukeiDataComboChanged(3)
    End Sub

    Private Sub ShukeiDataComboChanged(cboIndex As Integer)
        Dim dt As DataTable

        If IsDBNull(shukeiComboArray(cboIndex, 0).SelectedValue) Then
            Exit Sub
        End If

        shukeiComboArray(cboIndex, 2).DataSource = Nothing
        shukeiComboArray(cboIndex, 2).Items.Clear()

        Dim shukeiTable As DataTable
        If cboIndex = 3 Then
            shukeiTable = shukeiTablePrevYear
        Else
            shukeiTable = shukeiTableThisYear
        End If

        Dim rows() As DataRow = shukeiTable.Select("集計番号 = '" + shukeiComboArray(cboIndex, 0).SelectedValue.ToString + "'")
        dt = shukeiTable.Clone

        For Each row As DataRow In rows
            dt.ImportRow(row)
        Next
        dt = dt.DefaultView.ToTable(True, "地域コード", "地域名")

        ' 空白はデフォルト(全国)とする
        Dim newRow As DataRow = dt.NewRow
        newRow.Item("地域コード") = CInt(ComConst.地域.全国)

        dt.Rows.InsertAt(newRow, 0)
        shukeiComboArray(cboIndex, 1).ValueMember = "地域コード"
        shukeiComboArray(cboIndex, 1).DisplayMember = "地域名"
        shukeiComboArray(cboIndex, 1).DataSource = dt

    End Sub

#End Region

#Region "【処理詳細仕様 2】「地域」コンボボックス選択"
    Private Sub cboChiki1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboChiki1.SelectedIndexChanged
        ChikiDataComboChanged(0)
    End Sub

    Private Sub cboChiki2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboChiki2.SelectedIndexChanged
        ChikiDataComboChanged(1)
    End Sub

    Private Sub cboChiki3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboChiki3.SelectedIndexChanged
        ChikiDataComboChanged(2)
    End Sub

    Private Sub cboChikiZen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboChikiZen.SelectedIndexChanged
        ChikiDataComboChanged(3)
    End Sub

    Private Sub ChikiDataComboChanged(cboIndex As Integer)
        Dim dt As DataTable

        If IsDBNull(shukeiComboArray(cboIndex, 1).SelectedValue) Then
            Exit Sub
        End If

        Dim shukeiTable As DataTable
        If cboIndex = 3 Then
            shukeiTable = shukeiTablePrevYear
        Else
            shukeiTable = shukeiTableThisYear
        End If

        Dim rows() As DataRow = shukeiTable.Select("集計番号 = '" + shukeiComboArray(cboIndex, 0).SelectedValue.ToString + "' AND 地域コード = " + shukeiComboArray(cboIndex, 1).SelectedValue.ToString)
        dt = shukeiTable.Clone
        For Each row As DataRow In rows
            dt.ImportRow(row)
        Next
        dt = dt.DefaultView.ToTable(True, "規模階層表示用", "規模階層")

        ' 空白はデフォルト(当該客体の規模階層)とする
        Dim newRow As DataRow = dt.NewRow
        newRow.Item("規模階層") = _kiboKaiso

        dt.Rows.InsertAt(newRow, 0)
        shukeiComboArray(cboIndex, 2).ValueMember = "規模階層"
        shukeiComboArray(cboIndex, 2).DisplayMember = "規模階層表示用"
        shukeiComboArray(cboIndex, 2).DataSource = dt

    End Sub

#End Region

#Region "【処理詳細仕様 2】「規模階層」コンボボックス選択"

#End Region


#End Region

#Region "【処理詳細仕様 3】「設定取込」ボタンクリック"
    Private Sub btnInput_Click(sender As Object, e As EventArgs) Handles btnInput.Click
        Try
            Dim sjisEnc As Encoding = Encoding.GetEncoding(CODEPAGE_SHIFT_JIS)

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of OpenFileDialog)(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainInPath, IniFileInfo.ExcelInPath), , "csvファイル(*.csv)|*.csv")

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            If System.IO.File.Exists(filePath) Then
                Dim lstLine As New List(Of String())

                Using parser As New TextFieldParser(filePath, sjisEnc)

                    parser.TextFieldType = FieldType.Delimited
                    parser.SetDelimiters(CSV_DELIMITER)

                    While Not parser.EndOfData
                        Dim arr As String() = parser.ReadFields()
                        lstLine.Add(arr)
                    End While
                End Using

                Dim details As New List(Of String)

                '3-1 エラーチェック
                If Not CheckCSV(lstLine, details) Then
                    'メッセージフォーム表示
                    Message.ShowMsgForm(Me, MessageID.MSG_E_010, {String.Join(vbCrLf, details)})
                    Exit Sub
                End If

                '3-2 画面表示
                Dim dt As DataTable = settingTableLayout.Clone
                For Each strAry As String() In lstLine
                    If strAry(0).Equals(ファイルタイトル) Then
                        Continue For
                    End If

                    Dim row As DataRow
                    row = dt.NewRow
                    row("項目番号") = strAry(1)
                    row("明細番号") = strAry(2)
                    row("設定値") = strAry(3)
                    dt.Rows.Add(row)
                Next

                '「お取合せ先」の設定
                SetOtoiawaseSetting(dt)
                '「集計結果表」の設定
                SetShukeiKekkahyoSetting(dt)
                '「個別結果表」の設定
                SetKobetsuKekkahyoSetting(dt)

                Message.ShowMsgBox(MessageID.MSG_I_032, MsgBoxStyle.OkOnly)

            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try

    End Sub
#End Region

#Region "【処理詳細仕様 4】「設定出力」ボタンクリック"
    Private Sub btnOutput_Click(sender As Object, e As EventArgs) Handles btnOutput.Click

        Try
            Dim sjisEnc As Encoding = Encoding.GetEncoding(CODEPAGE_SHIFT_JIS)

            '4-1 ファイル保存ダイアログ選択

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath),
                                                                            ファイルタイトル + "_" + CommonInfo.ChosakubunName + "_" + _censusNo + ".csv", "csvファイル(*.csv)|*.csv")

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            '4-2 還元資料パラメータファイル出力
            Using sw As New System.IO.StreamWriter(filePath, False, sjisEnc)

                '1行目
                sw.WriteLine(ファイルタイトル + CSV_DELIMITER + CommonInfo.Chosakubun)

                '2行目以降
                Dim listSetting As New List(Of Dictionary(Of String, String))
                listSetting = GetSettingListFromForm()

                For Each dc As Dictionary(Of String, String) In listSetting
                    Dim str As String = ""
                    If String.IsNullOrEmpty(dc("設定値").ToString) = False Then
                        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                            str = CType(dc("項目番号"), 営農個人_項目番号).ToString
                        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                            str = CType(dc("項目番号"), 営農法人_項目番号).ToString
                        Else
                            str = CType(dc("項目番号"), 生産費_項目番号).ToString
                        End If
                        str += CSV_DELIMITER + dc("項目番号") + CSV_DELIMITER + dc("明細番号") + CSV_DELIMITER + dc("設定値")

                        sw.WriteLine(str)
                    End If

                Next
            End Using

            '4-3 完了メッセージ表示
            Message.ShowMsgBox(MessageID.MSG_I_033, MsgBoxStyle.OkOnly)

        Catch ex As System.IO.IOException
            'ファイル出力失敗
            Message.ShowMsgBox(MessageID.MSG_E_058, MsgBoxStyle.OkOnly)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try

    End Sub

#End Region

#Region "【処理詳細仕様 5】「設定して戻る」ボタンクリック"



    ''' <summary>
    ''' 「設定して戻る」ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSetting_Click(sender As Object, e As EventArgs) Handles btnSetting.Click

        Try
            If Message.ShowMsgBox(MessageID.MSG_Q_001, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If

            Dim dcSetting As New Dictionary(Of String, String)
            Dim listSetting As New List(Of Dictionary(Of String, String))

            Dim list As New List(Of Dictionary(Of String, String))

            '集計結果表の集計戸数チェック
            For i = 0 To 3
                If IsDBNull(shukeiComboArray(i, 0).SelectedValue) Or IsNothing(shukeiComboArray(i, 0).SelectedValue) Then
                    ' 集計結果表データ空白の場合チェックをスキップ
                    Continue For
                End If

                Dim shukeiTable As DataTable = shukeiTableThisYear
                If i = 3 Then
                    shukeiTable = shukeiTablePrevYear
                End If

                Dim rows() As DataRow = shukeiTable.Select("集計番号 = '" + shukeiComboArray(i, 0).SelectedValue.ToString +
                                                                   "' AND 地域コード = " + shukeiComboArray(i, 1).SelectedValue.ToString +
                                                                   " AND 規模階層 = " + shukeiComboArray(i, 2).SelectedValue.ToString)

                If rows.Count = 0 Then
                    Message.ShowMsgBox(MessageID.MSG_E_067, {shukeiComboArray(i, 0).Text.ToString}, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                For Each row As DataRow In rows
                    Dim itemName As String
                    If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                        itemName = "営農集計戸数"
                    Else
                        itemName = "生産費集計経営体数"
                    End If

                    If CInt(rows(0).Item(itemName)) <= 2 Then
                        '集計戸数が2以下の場合はエラーとする
                        Message.ShowMsgBox(MessageID.MSG_E_066, {shukeiComboArray(i, 0).Text.ToString}, MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If
                Next
            Next

            '個別結果表設定値チェック
            Dim errList As List(Of String) = CheckKobetsuKobanAll()
            If errList.Count <> 0 Then
                errList.Insert(0, String.Empty)
                Message.ShowMsgForm(Me, MessageID.MSG_E_064, {String.Join(vbCrLf, errList)})
                Exit Sub
            End If

            listSetting = GetSettingListFromForm()

            Using db As New DBAccess(My.Settings.BRASConnectionString)
                Try
                    db.BeginTrans()

                    For Each dc As Dictionary(Of String, String) In listSetting
                        Try
                            DAOOther.UpdateKangenShiryoSetting(db, dc)
                        Catch ex As Exception
                            ' 更新対象がなかった場合、追加する
                            DAOOther.InsertKangenShiryoSetting(db, dc)
                        End Try
                    Next

                    db.CommitTrans()
                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using

            Me.Close()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region

#Region "【処理詳細仕様 6】「戻る」ボタンクリック"

    Private Sub btnReturnEx_Click(sender As Object, e As EventArgs) Handles btnReturnEx.Click
        If Message.ShowMsgBox(MessageID.MSG_Q_002, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Me.Close()
        End If

    End Sub

#End Region

#Region "処理全般"
    ''' <summary>
    ''' お問合せ先設定
    ''' </summary>
    ''' <param name=""></param>
    ''' <remarks></remarks>
    Private Sub SetOtoiawaseSetting(dt As DataTable)

        'REV_003↓
        txtOffice.Text = ""
        txtLocal.Text = ""
        txtOperator.Text = ""
        txtTele.Text = ""
        txtFax.Text = ""
        txtAddress.Text = ""
        'REV_003↑

        If dt.Rows.Count <> 0 Then
            For Each row As DataRow In dt.Select("項目番号 = 1")
                Dim statement As Integer
                If Integer.TryParse(row.Item("明細番号").ToString, statement) = True Then
                    Select Case statement
                        Case 1
                            txtOffice.Text = row.Item("設定値").ToString
                        Case 2
                            txtLocal.Text = row.Item("設定値").ToString
                        Case 3
                            txtOperator.Text = row.Item("設定値").ToString
                        Case 4
                            txtTele.Text = row.Item("設定値").ToString
                        Case 5
                            txtFax.Text = row.Item("設定値").ToString
                        Case 6
                            txtAddress.Text = row.Item("設定値").ToString
                    End Select

                End If
            Next
        End If

    End Sub

    ''' <summary>
    ''' 集計結果表コンボボックスにデータテーブルを関連付ける
    ''' </summary>
    ''' <param name="cb"></param>
    ''' <param name="dt"></param>
    ''' <remarks></remarks>
    Private Sub SetShukeiKekkahyoComboBox(cb As ComboBox, dt As DataTable)

        cb.ValueMember = "集計番号"
        cb.DisplayMember = "表示用"
        cb.DataSource = dt
    End Sub

    ''' <summary>
    ''' 集計結果表設定
    ''' </summary>
    ''' <param name=""></param>
    ''' <remarks></remarks>
    Private Sub InitShukeiKekkahyoComboBox()

        Dim dt As DataTable

        ' 集計結果表データコンボボックス設定
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

            Dim zennen As String = (Integer.Parse(_chosaNen) - 1).ToString

            ' 実査設置拠点に含まれる地域を取得
            Dim chikiList As New List(Of String)

            Dim pref As String = Integer.Parse(CommonInfo.Jimusyo).ToString    'ゼロサプレス
            '北海道の場合都道府県コードに直す
            Select Case pref
                Case "51", "52", "53", "54"
                    pref = ComConst.地域.北海道
            End Select
            For Each kv As KeyValuePair(Of String, ComConst.地域.詳細) In ComConst.地域.リスト
                If kv.Value.都道府県.Contains(pref) Then
                    chikiList.Add(kv.Key)
                End If
            Next

            shukeiTableThisYear = DAOSyukeiKekkahyo.GetListKangen(db, CommonInfo.Chosakubun, _chosaNen, chikiList.ToArray)
            shukeiTablePrevYear = DAOSyukeiKekkahyo.GetListKangen(db, CommonInfo.Chosakubun, zennen, chikiList.ToArray)


            'コンボボックス表示用のカラムを作成
            shukeiTableThisYear.Columns.Add("表示用")
            shukeiTableThisYear.Columns.Add("地域名")
            shukeiTableThisYear.Columns.Add("規模階層表示用")
            For Each row As DataRow In shukeiTableThisYear.Rows
                row.Item("表示用") = row.Item("集計番号").ToString + " " + row.Item("集計名称").ToString
                row.Item("地域名") = row.Item("地域コード").ToString + " " + ComConst.地域.リスト(row.Item("地域コード").ToString).名称
                row.Item("規模階層表示用") = row.Item("規模階層").ToString
            Next

            shukeiTablePrevYear.Columns.Add("表示用")
            shukeiTablePrevYear.Columns.Add("地域名")
            shukeiTablePrevYear.Columns.Add("規模階層表示用")
            For Each row As DataRow In shukeiTablePrevYear.Rows
                row.Item("表示用") = row.Item("集計番号").ToString + " " + row.Item("集計名称").ToString
                row.Item("地域名") = row.Item("地域コード").ToString + " " + ComConst.地域.リスト(row.Item("地域コード").ToString).名称
                row.Item("規模階層表示用") = row.Item("規模階層").ToString
            Next

            'コンボボックス初期化
            For Each cbo As ComboBox In shukeiComboArray
                cbo.Items.Clear()
            Next

            '重複の削除
            dt = shukeiTableThisYear.DefaultView.ToTable(True, "集計番号", "表示用")
            '空白行挿入
            dt.Rows.InsertAt(dt.NewRow(), 0)

            SetShukeiKekkahyoComboBox(cboShukeiData1, dt)
            SetShukeiKekkahyoComboBox(cboShukeiData2, dt.Copy)
            SetShukeiKekkahyoComboBox(cboShukeiData3, dt.Copy)

            '重複の削除
            dt = shukeiTablePrevYear.DefaultView.ToTable(True, "集計番号", "表示用")
            '空白行挿入
            dt.Rows.InsertAt(dt.NewRow(), 0)

            SetShukeiKekkahyoComboBox(cboShukeiDataZen, dt)
        End Using

    End Sub

    ''' <summary>
    ''' 集計結果表コンボボックスのデータ設定
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <remarks></remarks>
    Private Sub SetShukeiKekkahyoSetting(dt As DataTable)

        'REV_003↓
        For Each cbo As ComboBox In shukeiComboArray
            If cbo.Items.Count > 0 Then
                cbo.SelectedIndex = 0
            End If
        Next
        'REV_003↑

        For Each row As DataRow In dt.Select("項目番号 = 2")
            Dim values() As String = row.Item("設定値").ToString.Split("_"c)
            If values.Count <> 4 Then
                Continue For
            End If
            If shukeiComboArray(CInt(row.Item("明細番号")) - 1, 0).FindString(values(1)) <> -1 Then
                shukeiComboArray(CInt(row.Item("明細番号")) - 1, 0).SelectedValue = values(1)
            End If
            If shukeiComboArray(CInt(row.Item("明細番号")) - 1, 1).FindString(values(2)) <> -1 Then
                shukeiComboArray(CInt(row.Item("明細番号")) - 1, 1).SelectedValue = values(2)
            End If
            If shukeiComboArray(CInt(row.Item("明細番号")) - 1, 2).FindString(values(3)) <> -1 Then
                shukeiComboArray(CInt(row.Item("明細番号")) - 1, 2).SelectedValue = values(3)
            End If
        Next

    End Sub

    Private Sub AddDgvList(dgv As DataGridView, addCount As Integer)
        If dgv.Rows.Count = 0 Then
            dgv.Rows.Add(addCount)
        End If
    End Sub

    ''' <summary>
    ''' 個別結果表設定
    ''' </summary>
    ''' <param name=""></param>
    ''' <remarks></remarks>
    Private Sub SetKobetsuKekkahyoSetting(dt As DataTable)
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            '営農個人
            'REV_003↓
            dgvList_Eiko2.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            dgvList_Eiko3.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            txtEiko4_1.Text = ""
            txtEiko4_2.Text = ""
            dgvList_Eiko7.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            dgvList_Eiko8.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            dgvList_Eiko9.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            'REV_003↑

            '営個還２
            Me.AddDgvList(dgvList_Eiko2, 8)
            Me.SetDataToDgv(dgvList_Eiko2, 3, dt)
            '営個還３
            Me.AddDgvList(dgvList_Eiko3, 3)
            Me.SetDataToDgv(dgvList_Eiko3, 4, dt)
            '営個還４
            Dim rows() As DataRow = dt.Select("項目番号 = 5")
            For Each row As DataRow In rows
                Dim meisaiNo As Integer
                If Integer.TryParse(row.Item("明細番号").ToString, meisaiNo) Then
                    Select Case meisaiNo
                        Case ComConst.還元資料.還元資料項目_営農個人.営個還4_グラフ項目_面積_収入_明細番号.面積
                            txtEiko4_1.Text = row.Item("設定値").ToString
                        Case ComConst.還元資料.還元資料項目_営農個人.営個還4_グラフ項目_面積_収入_明細番号.収入
                            txtEiko4_2.Text = row.Item("設定値").ToString
                    End Select
                End If
            Next
            '営個還７
            Me.AddDgvList(dgvList_Eiko7, 7)
            Me.SetDataToDgv(dgvList_Eiko7, 6, dt)
            '営個還８
            Me.AddDgvList(dgvList_Eiko8, 4)
            Me.SetDataToDgv(dgvList_Eiko8, 7, dt)
            '営個還９
            Me.AddDgvList(dgvList_Eiko9, 18)
            Me.SetDataToDgv(dgvList_Eiko9, 8, dt)

        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            '営農法人
            'REV_003↓
            dgvList_Eiho2.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            dgvList_Eiho3.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            txtEiho4_1.Text = ""
            txtEiho4_2.Text = ""
            dgvList_Eiho7.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            dgvList_Eiho8.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            dgvList_Eiho9.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            'REV_003↑

            '営法還２
            Me.AddDgvList(dgvList_Eiho2, 8)
            Me.SetDataToDgv(dgvList_Eiho2, 3, dt)
            '営法還３
            Me.AddDgvList(dgvList_Eiho3, 3)
            Me.SetDataToDgv(dgvList_Eiho3, 4, dt)
            '営法還４
            Dim rows() As DataRow = dt.Select("項目番号 = 5")
            For Each row As DataRow In rows
                Dim meisaiNo As Integer
                If Integer.TryParse(row.Item("明細番号").ToString, meisaiNo) Then
                    Select Case meisaiNo
                        Case ComConst.還元資料.還元資料項目_営農法人.営法還4_グラフ項目_面積_収入_明細番号.面積
                            txtEiho4_1.Text = row.Item("設定値").ToString
                        Case ComConst.還元資料.還元資料項目_営農法人.営法還4_グラフ項目_面積_収入_明細番号.収入
                            txtEiho4_2.Text = row.Item("設定値").ToString
                    End Select
                End If
            Next
            '営法還７
            Me.AddDgvList(dgvList_Eiho7, 6)
            Me.SetDataToDgv(dgvList_Eiho7, 6, dt)
            '営法還８
            Me.AddDgvList(dgvList_Eiho8, 4)
            Me.SetDataToDgv(dgvList_Eiho8, 7, dt)
            '営法還９
            Me.AddDgvList(dgvList_Eiho9, 29)
            Me.SetDataToDgv(dgvList_Eiho9, 8, dt)

        Else
            '生産費
            'REV_003↓
            dgvList_Seisan4.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            dgvList_Seisan6.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            dgvList_Seisan8_1.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            dgvList_Seisan8_2.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            dgvList_Seisan10.Rows.Cast(Of DataGridViewRow).ToList.ForEach(Sub(x) x.Cells(1).Value = "")
            'REV_003↑

            '還元４、還元５
            Me.AddDgvList(dgvList_Seisan4, 29)
            Me.SetDataToDgv(dgvList_Seisan4, 3, dt)
            '還元６
            Me.AddDgvList(dgvList_Seisan6, 8)
            Me.SetDataToDgv(dgvList_Seisan6, 4, dt)
            '還元８
            Me.AddDgvList(dgvList_Seisan8_1, 5)
            Me.SetDataToDgv(dgvList_Seisan8_1, 5, dt)
            Me.AddDgvList(dgvList_Seisan8_2, 5)
            Me.SetDataToDgv(dgvList_Seisan8_2, 6, dt)
            '還元１０
            Me.AddDgvList(dgvList_Seisan10, 5)
            Me.SetDataToDgv(dgvList_Seisan10, 7, dt)

        End If
    End Sub

    ''' <summary>
    ''' 個別結果表項番チェック（範囲）
    ''' </summary>
    ''' <param name="checkDataList"></param>
    ''' <param name="allowFrom"></param>
    ''' <param name="allowTo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckKobetsuKoban(checkDataList As List(Of String), allowFrom As String, allowTo As String) As List(Of String)
        Dim fromNum As Integer
        Dim toNum As Integer
        Dim errList As New List(Of String)
        ' REV_001↓
        Dim dtItem As DataTable = kobetsuItemTable.DefaultView.ToTable(True, "項目番号")
        ' REV_001↑

        If Not Integer.TryParse(allowFrom.Substring(1), fromNum) Then
            'チェック不能につき空のリスト返却
            Return errList
        End If
        If Not Integer.TryParse(allowTo.Substring(1), toNum) Then
            'チェック不能につき空のリスト返却
            Return errList
        End If

        For Each checkData As String In checkDataList
            Dim value As String = checkData
            If String.IsNullOrEmpty(checkData) Then
                '空白は許容
                Continue For
            End If

            '文字列長チェック
            If value.Length <> 7 Then
                errList.Add(value)
                Continue For
            End If

            '1文字目チェック(FROM、TOで別の文字の指定がない前提で、FROMと比較)
            If value(0) <> allowFrom(0) Then
                errList.Add(value)
                Continue For
            End If

            '数値チェック
            Dim intValue As Integer
            If Not Integer.TryParse(value.Substring(1), intValue) Then
                errList.Add(value)
                Continue For
            End If

            '範囲チェック
            If fromNum > intValue Or intValue > toNum Then
                errList.Add(value)
                Continue For
            End If

            ' REV_001↓
            Dim rows() As DataRow = dtItem.Select("項目番号 = '" + checkData + "'")
            If rows.Length = 0 Then
                errList.Add(checkData)
                Continue For
            End If
            ' REV_001↑
        Next

        Return errList
    End Function

    ''' <summary>
    ''' 個別結果表項番チェック（リスト）
    ''' </summary>
    ''' <param name="checkDataList"></param>
    ''' <param name="allowList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckKobetsuKoban(checkDataList As List(Of String), allowList As List(Of String)) As List(Of String)
        Dim errList As New List(Of String)
        ' REV_001↓
        Dim dtItem As DataTable = kobetsuItemTable.DefaultView.ToTable(True, "項目番号")
        ' REV_001↑

        For Each checkData As String In checkDataList
            If String.IsNullOrEmpty(checkData) Then
                '空白は許容
                Continue For
            End If
            ' REV_001↓
            Dim rows() As DataRow = dtItem.Select("項目番号 = '" + checkData + "'")
            If rows.Length = 0 Then
                errList.Add(checkData)
                Continue For
            End If
            ' REV_001↑
            If Not allowList.Contains(checkData) Then
                errList.Add(checkData)
            End If
        Next

        Return errList
    End Function

    ''' <summary>
    ''' 個別結果表マスタに項番が存在するかチェック
    ''' </summary>
    ''' <param name="checkDataList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckKobetsuKobanMaster(checkDataList As List(Of String)) As List(Of String)

        Dim dtItem As DataTable = kobetsuItemTable.DefaultView.ToTable(True, "項目番号")
        Dim errList As New List(Of String)

        '大文字小文字を区別してチェック
        dtItem.CaseSensitive = True

        For Each value As String In checkDataList
            If String.IsNullOrEmpty(value) Then
                '空白は許容
                Continue For
            End If
            Dim rows() As DataRow = dtItem.Select("項目番号 = '" + value + "'")
            If rows.Length = 0 Then
                errList.Add(value)
            End If
        Next

        Return errList

    End Function

    ''' <summary>
    ''' 生産費の個別結果表項番チェックリスト取得
    ''' </summary>
    ''' <param name="itemNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSeisanCheckList(itemNo As String) As List(Of String)
        Dim checkList As New List(Of String)

        Select Case itemNo
            Case ComConst.還元資料.還元資料項目_生産費.還元8_主要項目
                Return seisan8_1CheckList
            Case ComConst.還元資料.還元資料項目_生産費.還元8_主要作業
                Return seisan8_2CheckList
            Case ComConst.還元資料.還元資料項目_生産費.還元10_物財費内訳
                Return seisan10CheckList
        End Select
        Return Nothing
    End Function

    ''' <summary>
    ''' 生産費の個別結果表項番チェック範囲リスト取得
    ''' </summary>
    ''' <param name="itemNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSeisanCheckRangeList(itemNo As String) As List(Of List(Of String))

        Select Case itemNo
            Case ComConst.還元資料.還元資料項目_生産費.還元8_主要項目
                Return seisan8_1CheckRangeList
            Case ComConst.還元資料.還元資料項目_生産費.還元8_主要作業
                Return seisan8_2CheckRangeList
            Case ComConst.還元資料.還元資料項目_生産費.還元10_物財費内訳
                Return seisan10CheckRangeList
        End Select
        Return Nothing
    End Function

    ''' <summary>
    ''' 個別結果表項番チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckKobetsuKobanAll() As List(Of String)

        Dim errList As New List(Of String)
        Dim listSetting As New List(Of Dictionary(Of String, String))
        Dim workErrList As New List(Of String)
        Dim workErrList2 As New List(Of String)

        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then

            '営個還２
            workErrList.Clear()
            workErrList = GetValueListFromDgv(dgvList_Eiko2)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWLIST_EIKO2)
            If workErrList.Count <> 0 Then
                errList.Add("[営個還２]")
                errList.AddRange(workErrList)
            End If

            '営個還３
            workErrList.Clear()
            workErrList = GetValueListFromDgv(dgvList_Eiko3)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWLIST_EIKO3)
            If workErrList.Count <> 0 Then
                errList.Add("[営個還３]")
                errList.AddRange(workErrList)
            End If

            '営個還４
            '面積・頭羽数
            workErrList.Clear()
            workErrList.Add(txtEiko4_1.Text)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWLIST_EIKO4_1)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWRANGE_EIKO4_1_1(0), ALLOWRANGE_EIKO4_1_1(1))
            workErrList = CheckKobetsuKoban(workErrList, ALLOWRANGE_EIKO4_1_2(0), ALLOWRANGE_EIKO4_1_2(1))
            '収入
            workErrList2.Clear()
            workErrList2.Add(txtEiko4_2.Text)
            workErrList2 = CheckKobetsuKoban(workErrList2, ALLOWLIST_EIKO4_2)
            workErrList2 = CheckKobetsuKoban(workErrList2, ALLOWRANGE_EIKO4_2_1(0), ALLOWRANGE_EIKO4_2_1(1))
            workErrList2 = CheckKobetsuKoban(workErrList2, ALLOWRANGE_EIKO4_2_2(0), ALLOWRANGE_EIKO4_2_2(1))
            If workErrList.Count <> 0 Or workErrList2.Count <> 0 Then
                errList.Add("[営個還４]")
                errList.AddRange(workErrList)
                errList.AddRange(workErrList2)
            End If

            '営個還７　(営個還４ 面積・頭羽数と同チェック)
            workErrList.Clear()
            workErrList = GetValueListFromDgv(dgvList_Eiko7)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWLIST_EIKO4_1)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWRANGE_EIKO4_1_1(0), ALLOWRANGE_EIKO4_1_1(1))
            workErrList = CheckKobetsuKoban(workErrList, ALLOWRANGE_EIKO4_1_2(0), ALLOWRANGE_EIKO4_1_2(1))
            If workErrList.Count <> 0 Then
                errList.Add("[営個還７]")
                errList.AddRange(workErrList)
            End If

            '営個還８
            workErrList.Clear()
            workErrList = GetValueListFromDgv(dgvList_Eiko8)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWLIST_EIKO8)
            If workErrList.Count <> 0 Then
                errList.Add("[営個還８]")
                errList.AddRange(workErrList)
            End If

            '営個還９は個別結果表全項目から選択可能であるため、個別結果表マスタに存在するかチェック
            workErrList.Clear()
            workErrList = CheckKobetsuKobanMaster(GetValueListFromDgv(dgvList_Eiko9))
            If workErrList.Count <> 0 Then
                errList.Add("[営個還９]")
                errList.AddRange(workErrList)
            End If

        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then

            '営法還２
            workErrList.Clear()
            workErrList = GetValueListFromDgv(dgvList_Eiho2)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWLIST_EIHO2)
            If workErrList.Count <> 0 Then
                errList.Add("[営法還２]")
                errList.AddRange(workErrList)
            End If

            '営法還３
            workErrList.Clear()
            workErrList = GetValueListFromDgv(dgvList_Eiho3)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWLIST_EIHO3)
            If workErrList.Count <> 0 Then
                errList.Add("[営法還３]")
                errList.AddRange(workErrList)
            End If

            '営法還４
            '面積・頭羽数
            workErrList.Clear()
            workErrList.Add(txtEiho4_1.Text)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWLIST_EIHO4_1)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWRANGE_EIHO4_1_1(0), ALLOWRANGE_EIHO4_1_1(1))
            workErrList = CheckKobetsuKoban(workErrList, ALLOWRANGE_EIHO4_1_2(0), ALLOWRANGE_EIHO4_1_2(1))
            '収入
            workErrList2.Clear()
            workErrList2.Add(txtEiho4_2.Text)
            workErrList2 = CheckKobetsuKoban(workErrList2, ALLOWLIST_EIHO4_2)
            workErrList2 = CheckKobetsuKoban(workErrList2, ALLOWRANGE_EIHO4_2_1(0), ALLOWRANGE_EIHO4_2_1(1))
            workErrList2 = CheckKobetsuKoban(workErrList2, ALLOWRANGE_EIHO4_2_2(0), ALLOWRANGE_EIHO4_2_2(1))
            If workErrList.Count <> 0 Or workErrList2.Count <> 0 Then
                errList.Add("[営法還４]")
                errList.AddRange(workErrList)
                errList.AddRange(workErrList2)
            End If

            '営法還７　(営法還４ 面積・頭羽数と同チェック)
            workErrList.Clear()
            workErrList = GetValueListFromDgv(dgvList_Eiho7)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWLIST_EIHO4_1)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWRANGE_EIHO4_1_1(0), ALLOWRANGE_EIHO4_1_1(1))
            workErrList = CheckKobetsuKoban(workErrList, ALLOWRANGE_EIHO4_1_2(0), ALLOWRANGE_EIHO4_1_2(1))
            If workErrList.Count <> 0 Then
                errList.Add("[営法還７]")
                errList.AddRange(workErrList)
            End If

            '営法還８
            workErrList.Clear()
            workErrList = GetValueListFromDgv(dgvList_Eiho8)
            workErrList = CheckKobetsuKoban(workErrList, ALLOWLIST_EIHO8)
            If workErrList.Count <> 0 Then
                errList.Add("[営法還８]")
                errList.AddRange(workErrList)
            End If

            '営法還９は個別結果表全項目から選択可能であるため、個別結果表マスタに存在するかチェック
            workErrList.Clear()
            workErrList = CheckKobetsuKobanMaster(GetValueListFromDgv(dgvList_Eiho9))
            If workErrList.Count <> 0 Then
                errList.Add("[営法還９]")
                errList.AddRange(workErrList)
            End If

        Else
            Dim checkList As New List(Of String)
            Dim checkRangeList As New List(Of List(Of String))

            '還元４、還元５は個別結果表全項目から選択可能であるため、個別結果表マスタに存在するかチェック
            workErrList.Clear()
            workErrList = CheckKobetsuKobanMaster(GetValueListFromDgv(dgvList_Seisan4))
            If workErrList.Count <> 0 Then
                errList.Add("[還元４、還元５]")
                errList.AddRange(workErrList)
            End If

            '還元６は個別結果表全項目から選択可能であるため、個別結果表マスタに存在するかチェック
            workErrList.Clear()
            workErrList = CheckKobetsuKobanMaster(GetValueListFromDgv(dgvList_Seisan6))
            If workErrList.Count <> 0 Then
                errList.Add("[還元６]")
                errList.AddRange(workErrList)
            End If

            '還元８
            checkList = GetSeisanCheckList(ComConst.還元資料.還元資料項目_生産費.還元8_主要項目)
            checkRangeList = GetSeisanCheckRangeList(ComConst.還元資料.還元資料項目_生産費.還元8_主要項目)
            workErrList.Clear()
            workErrList = GetValueListFromDgv(dgvList_Seisan8_1)
            workErrList = CheckKobetsuKoban(workErrList, checkList)
            For Each range As List(Of String) In checkRangeList
                workErrList = CheckKobetsuKoban(workErrList, range(0), range(1))
            Next
            If workErrList.Count <> 0 Then
                errList.Add("[還元８] 上")
                errList.AddRange(workErrList)
            End If

            checkList = GetSeisanCheckList(ComConst.還元資料.還元資料項目_生産費.還元8_主要作業)
            checkRangeList = GetSeisanCheckRangeList(ComConst.還元資料.還元資料項目_生産費.還元8_主要作業)
            workErrList.Clear()
            workErrList = GetValueListFromDgv(dgvList_Seisan8_2)
            workErrList = CheckKobetsuKoban(workErrList, checkList)
            For Each range As List(Of String) In checkRangeList
                workErrList = CheckKobetsuKoban(workErrList, range(0), range(1))
            Next
            If workErrList.Count <> 0 Then
                errList.Add("[還元８] 下")
                errList.AddRange(workErrList)
            End If

            '還元１０
            checkList = GetSeisanCheckList(ComConst.還元資料.還元資料項目_生産費.還元10_物財費内訳)
            checkRangeList = GetSeisanCheckRangeList(ComConst.還元資料.還元資料項目_生産費.還元10_物財費内訳)
            workErrList.Clear()
            workErrList = GetValueListFromDgv(dgvList_Seisan10)
            workErrList = CheckKobetsuKoban(workErrList, checkList)
            For Each range As List(Of String) In checkRangeList
                workErrList = CheckKobetsuKoban(workErrList, range(0), range(1))
            Next
            If workErrList.Count <> 0 Then
                errList.Add("[還元１０]")
                errList.AddRange(workErrList)
            End If

        End If

        Return errList
    End Function

    Private Function GetSettingListFromForm() As List(Of Dictionary(Of String, String))
        Dim listSetting As New List(Of Dictionary(Of String, String))
        Dim list As New List(Of Dictionary(Of String, String))

        'お問い合わせ先
        For i = 0 To 5
            Dim dc As New Dictionary(Of String, String)

            dc("センサス番号") = _censusNo
            dc("項目番号") = ComConst.還元資料.お問合せ先
            dc("明細番号") = (i + 1).ToString
            Select Case i
                Case 0
                    dc("設定値") = txtOffice.Text
                Case 1
                    dc("設定値") = txtLocal.Text
                Case 2
                    dc("設定値") = txtOperator.Text
                Case 3
                    dc("設定値") = txtTele.Text
                Case 4
                    dc("設定値") = txtFax.Text
                Case 5
                    dc("設定値") = txtAddress.Text
            End Select

            list.Add(dc)
        Next
        listSetting.AddRange(list)
        list.Clear()

        '使用する集計結果表
        For i As Integer = 0 To 3
            Dim dc As New Dictionary(Of String, String)

            Dim chosaNen As String = _chosaNen
            If i = 3 Then
                chosaNen = (CInt(chosaNen) - 1).ToString
            End If

            dc("センサス番号") = _censusNo
            dc("項目番号") = ComConst.還元資料.集計結果表
            dc("明細番号") = (i + 1).ToString
            If IsDBNull(shukeiComboArray(i, 0).SelectedValue) Then
                ' 集計結果表データ空白の場合
                dc("設定値") = String.Empty
            Else
                dc("設定値") =
                    chosaNen + "_" +
                    shukeiComboArray(i, 0).SelectedValue.ToString + "_" +
                    shukeiComboArray(i, 1).SelectedValue.ToString + "_" +
                    shukeiComboArray(i, 2).SelectedValue.ToString
            End If

            list.Add(dc)
        Next
        listSetting.AddRange(list)
        list.Clear()

        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            '営個還２
            listSetting.AddRange(GetDataFromDgv(dgvList_Eiko2, 3))
            '営個還３
            listSetting.AddRange(GetDataFromDgv(dgvList_Eiko3, 4))
            '営個還４
            listSetting.Add(GetDataFromTextBox(txtEiko4_1, 5, 1))
            listSetting.Add(GetDataFromTextBox(txtEiko4_2, 5, 2))
            '営個還７
            listSetting.AddRange(GetDataFromDgv(dgvList_Eiko7, 6))
            '営個還８
            listSetting.AddRange(GetDataFromDgv(dgvList_Eiko8, 7))
            '営個還９
            listSetting.AddRange(GetDataFromDgv(dgvList_Eiko9, 8))
        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            '営法還２
            listSetting.AddRange(GetDataFromDgv(dgvList_Eiho2, 3))
            '営法還３
            listSetting.AddRange(GetDataFromDgv(dgvList_Eiho3, 4))
            '営法還４
            listSetting.Add(GetDataFromTextBox(txtEiho4_1, 5, 1))
            listSetting.Add(GetDataFromTextBox(txtEiho4_2, 5, 2))
            '営法還７
            listSetting.AddRange(GetDataFromDgv(dgvList_Eiho7, 6))
            '営法還８
            listSetting.AddRange(GetDataFromDgv(dgvList_Eiho8, 7))
            '営法還９
            listSetting.AddRange(GetDataFromDgv(dgvList_Eiho9, 8))
        Else
            '還元４、還元５
            listSetting.AddRange(GetDataFromDgv(dgvList_Seisan4, 3))
            '還元６
            listSetting.AddRange(GetDataFromDgv(dgvList_Seisan6, 4))
            '還元８
            listSetting.AddRange(GetDataFromDgv(dgvList_Seisan8_1, 5))
            listSetting.AddRange(GetDataFromDgv(dgvList_Seisan8_2, 6))
            '還元１０
            listSetting.AddRange(GetDataFromDgv(dgvList_Seisan10, 7))
        End If

        Return listSetting
    End Function

    ''' <summary>
    ''' DataGridViewから設定値の一覧取得
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetValueListFromDgv(dgv As DataGridView) As List(Of String)
        Dim list As New List(Of String)

        For i = 0 To dgv.Rows.Count - 1
            Dim dc As New Dictionary(Of String, String)

            If dgv.Item(1, i).Value Is Nothing Then
                list.Add(String.Empty)
            Else
                list.Add(dgv.Item(1, i).Value.ToString)
            End If
        Next

        Return list
    End Function

    ''' <summary>
    ''' DataGridViewからデータ取得
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <param name="itemNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetDataFromDgv(dgv As DataGridView, itemNo As Integer) As List(Of Dictionary(Of String, String))
        Dim list As New List(Of Dictionary(Of String, String))

        For i = 0 To dgv.Rows.Count - 1
            Dim dc As New Dictionary(Of String, String)

            dc("センサス番号") = _censusNo
            dc("項目番号") = itemNo.ToString
            dc("明細番号") = (i + 1).ToString
            If dgv.Item(1, i).Value Is Nothing Then
                dc("設定値") = String.Empty
            Else
                dc("設定値") = dgv.Item(1, i).Value.ToString
            End If

            list.Add(dc)
        Next

        Return list
    End Function

    ''' <summary>
    ''' DataGridViewへのデータ設定
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <param name="itemNo"></param>
    ''' <param name="dt"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetDataToDgv(dgv As DataGridView, itemNo As Integer, dt As DataTable) As Boolean

        For i = 0 To dgv.Rows.Count - 1
            dgv.Item(0, i).Value = (i + 1).ToString
        Next

        For Each row As DataRow In dt.Select("項目番号 = " + itemNo.ToString)
            Dim dgvRow As Integer
            If Integer.TryParse(row.Item("明細番号").ToString, dgvRow) Then
                dgv.Item(1, dgvRow - 1).Value = row.Item("設定値").ToString
            End If
        Next

        dgv.Columns(0).Resizable = DataGridViewTriState.False
        dgv.Columns(1).Resizable = DataGridViewTriState.False

        Return True
    End Function

    Private Function GetDataFromTextBox(tb As TextBox, item As Integer, detail As Integer) As Dictionary(Of String, String)
        Dim dc As New Dictionary(Of String, String)

        dc("センサス番号") = _censusNo
        dc("項目番号") = item.ToString
        dc("明細番号") = detail.ToString
        dc("設定値") = tb.Text

        Return dc
    End Function

    ''' <summary>
    ''' 還元資料パラメータファイルチェック
    ''' </summary>
    ''' <param name=""></param>
    ''' <remarks></remarks>
    Private Function CheckCSV(pCsvList As List(Of String()), ByRef pErrList As List(Of String)) As Boolean
        Dim ret As Boolean = True

        Const max As Integer = ComConst.ERR_MESSAGE_MAX

        Dim msg As String
        Dim msg2 As String
        Dim cnt As Integer = 1
        Dim distinct As New Dictionary(Of String, String)

        For Each csvLine As String() In pCsvList
            If cnt = 1 Then
                '1行目のチェック
                '①項目の桁数チェック
                msg = "ファイルの1行目のデータ項目数を確認してください。"
                If csvLine.Count <> 2 Then
                    pErrList.Add(msg)
                End If

                '②ファイルタイトルが「還元資料パラメータ」であるか
                msg = "ファイルタイトルが正しくありません。"
                If Not ファイルタイトル.Equals(csvLine(0)) Then
                    pErrList.Add(msg)
                End If

                '③調査区分が実施している調査区分と同じであるか
                msg = "調査区分が正しくありません。"
                If Not CommonInfo.Chosakubun.Equals(csvLine(1)) Then
                    pErrList.Add(msg)
                    ret = False
                End If

                'ここでエラーとなった場合、以降のチェックは行わない
                If pErrList.Count <> 0 Then
                    Exit For
                End If
            Else
                '2行目以降についてのチェックを行う
                '④項目番号がコード一覧に入っているか
                Dim intVal As Integer
                Dim isError As Boolean = False
                msg = "X行目:一覧に存在しない項目番号です。"
                msg2 = "X行目:存在しない明細番号です。"
                If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                    If Not ComConst.還元資料.還元資料項目_営農個人.明細番号一覧.ContainsKey(csvLine(1)) Then
                        msg = msg.Replace("X", CStr(cnt))
                        pErrList.Add(msg)
                        If pErrList.Count = max Then
                            Exit For
                        End If
                    End If

                    '明細番号チェック
                    If Not Integer.TryParse(csvLine(2), intVal) Then
                        isError = True
                    Else
                        If csvLine(1) = "1" Then
                            If intVal < 1 Or 6 < intVal Then
                                isError = True
                            End If
                        Else
                            If Not ComConst.還元資料.還元資料項目_営農個人.明細番号一覧(csvLine(1)).Contains(intVal) Then
                                isError = True
                            End If
                        End If
                    End If
                    If isError Then
                        msg2 = msg2.Replace("X", CStr(cnt))
                        pErrList.Add(msg2)
                        If pErrList.Count = max Then
                            Exit For
                        End If
                    End If

                ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                    If Not ComConst.還元資料.還元資料項目_営農法人.明細番号一覧.ContainsKey(csvLine(1)) Then
                        msg = msg.Replace("X", CStr(cnt))
                        pErrList.Add(msg)
                        If pErrList.Count = max Then
                            Exit For
                        End If
                    End If

                    '明細番号チェック
                    If Not Integer.TryParse(csvLine(2), intVal) Then
                        isError = True
                    Else
                        If csvLine(1) = "1" Then
                            If intVal < 1 Or 6 < intVal Then
                                isError = True
                            End If
                        Else
                            If Not ComConst.還元資料.還元資料項目_営農法人.明細番号一覧(csvLine(1)).Contains(intVal) Then
                                isError = True
                            End If
                        End If
                    End If
                    If isError Then
                        msg2 = msg2.Replace("X", CStr(cnt))
                        pErrList.Add(msg2)
                        If pErrList.Count = max Then
                            Exit For
                        End If
                    End If
                Else
                    If Not ComConst.還元資料.還元資料項目_生産費.明細番号一覧.ContainsKey(csvLine(1)) Then
                        msg = msg.Replace("X", CStr(cnt))
                        pErrList.Add(msg)
                        If pErrList.Count = max Then
                            Exit For
                        End If
                    End If

                    '明細番号チェック
                    If Not Integer.TryParse(csvLine(2), intVal) Then
                        isError = True
                    Else
                        If csvLine(1) = "1" Then
                            If intVal < 1 Or 6 < intVal Then
                                isError = True
                            End If
                        Else
                            If Not ComConst.還元資料.還元資料項目_生産費.明細番号一覧(csvLine(1)).Contains(intVal) Then
                                isError = True
                            End If
                        End If
                    End If
                    If isError Then
                        msg2 = msg2.Replace("X", CStr(cnt))
                        pErrList.Add(msg2)
                        If pErrList.Count = max Then
                            Exit For
                        End If
                    End If
                End If

                '⑤他の行と重複していないか
                msg = "X行目:項目が重複しています。"
                Dim key As String = csvLine(1) & csvLine(2) & csvLine(3)
                'ディクショナリを確認し、あればエラーとする。
                If distinct.ContainsKey(key) Then
                    msg = msg.Replace("X", CStr(cnt))
                    pErrList.Add(msg)
                    If pErrList.Count = max Then
                        Exit For
                    End If
                Else
                    distinct.Add(key, csvLine(0))
                End If

                '⑥項目番号が[2]以上の場合、設定値が[半角英数字1文字]+[半角数字6文字]となっているか
                msg = "X行目:設定値が不正です。"
                If csvLine(1) > "2" Then
                    '7桁であるか
                    If Not csvLine(3).Length = 7 Then
                        msg = msg.Replace("X", CStr(cnt))
                        pErrList.Add(msg)

                        If pErrList.Count = max Then
                            Exit For
                        End If

                        Continue For
                    End If
                    '1文字目が半角英数字であるか
                    If Not Regex.IsMatch(csvLine(3).Substring(1, 6), "^[a-zA-Z0-9]+$") Then
                        msg = msg.Replace("X", CStr(cnt))
                        pErrList.Add(msg)
                        If pErrList.Count = max Then
                            Exit For
                        End If
                    End If
                    '2文字目以降が半角数字であるか
                    If Not Regex.IsMatch(csvLine(3).Substring(1, 6), "^[0-9]+$") Then
                        msg = msg.Replace("X", CStr(cnt))
                        pErrList.Add(msg)
                        If pErrList.Count = max Then
                            Exit For
                        End If
                    End If

                End If

                '⑦項目番号が[1]の場合、最大桁数を超えていないか
                msg = "X行目:N桁以下で登録してください。"
                If csvLine(1) = "1" Then
                    Select Case (csvLine(2))
                        Case "1", "2", "3"
                            '[局・事務所・事務局],[県拠点等],[担当者]が10桁以下であるか
                            If csvLine(3).Length > 10 Then
                                msg = msg.Replace("X", CStr(cnt))
                                msg = msg.Replace("N", "10")
                                pErrList.Add(msg)
                                If pErrList.Count = max Then
                                    Exit For
                                End If
                            End If
                        Case "4", "5"
                            '[問合せ電話番号],[FAX]が15桁以下であるか
                            If csvLine(3).Length > 15 Then
                                msg = msg.Replace("X", CStr(cnt))
                                msg = msg.Replace("N", "15")
                                pErrList.Add(msg)
                                If pErrList.Count = max Then
                                    Exit For
                                End If
                            End If
                        Case "6"
                            '[住所]が30桁以下であるか
                            If csvLine(3).Length > 30 Then
                                msg = msg.Replace("X", CStr(cnt))
                                msg = msg.Replace("N", "30")
                                pErrList.Add(msg)
                                If pErrList.Count = max Then
                                    Exit For
                                End If
                            End If
                    End Select
                End If
            End If

            cnt += 1

        Next

        If pErrList.Count <> 0 Then
            ret = False
        End If

        Return ret
    End Function
#End Region

End Class
