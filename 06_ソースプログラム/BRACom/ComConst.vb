'------------------------------------------------------------------------------------------
'| REV | 変更年月日 | 変更者                   | 変更内容
'------------------------------------------------------------------------------------------
'| 001 | 2019/05/15 | Daiko                    | 新規作成
'| 002 | 2020/08/10 | TSP                      | 還元資料出力機能実装
'| 003 | 2020/11/12 | TSP                      | 要件No1,6,追加要件No14 対応
'| 004 | 2020/11/29 | 日本コンピュータシステム | 要件No1-③対応
'| 005 | 2021/12/16 | 日本コンピュータシステム | 要件No10
'| 006 | 2021/12/16 | 日本コンピュータシステム | 要件No1-⑤
'| 007 | 2022/01/18 | 日本コンピュータシステム | 要件No2
'| 008 | 2022/10/11 | Daiko                    | 要件No1,14
'| 009 | 2022/10/11 | 大興電子通信             | 変更要件No.5,No.6
'| 010 | 2022/12/15 | Daiko                    | 要件No4
'| 011 | 2022/12/20 | Daiko                    | 要件No15
'| 012 | 2023/01/04 | Daiko                    | 要件No.8
'| 013 | 2023/01/04 | Daiko                    | 要件No.19
'| 014 | 2023/03/15 | Daiko                    | 要件No6 現行不具合修正
'| 015 | 2023/11/27 | Daiko                    | 要件No20
'| 016 | 2023/12/18 | Daiko                    | BRA-10372対応
'| 017 | 2024/03/11 | Daiko                    | 変更要件No.20 連絡票No.210
'| 018 | 2024/05/31 | Daiko                    | 要件No.1
'| 019 | 2025/08/28 | GCU                      | 要件No2 継続区分追加
'| 020 | 2025/12/09 | GCU                      | 要件No1 出力内容区分、出力内容列追加
'------------------------------------------------------------------------------------------

''' <summary>
''' 定数クラス
''' </summary>
''' <remarks></remarks>
Public Class ComConst

    ''' <summary>項目番号区切文字</summary>
    Public Const ITEM_NO_DELIMITER As String = "_"

    ''' <summary>日時フォーマット</summary>
    Public Const DATETIME_FORMAT As String = "yyyy/MM/dd HH:mm"

    ''' <summary>エラーメッセージ最大</summary>
    Public Const ERR_MESSAGE_MAX As Integer = 50

    ''' <summary>チェックボックス接頭辞</summary>
    Public Const PRE_CHECKBOX As String = "chk"

    ''' <summary>数値２桁フォーマット</summary>
    Public Const DIGIT_2_FORMAT As String = "00"

    ''' <summary>
    ''' 区分１クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 区分１
        Public Const 農業経営統計調査 As String = "1"
        Public Const 組織法人経営体に関する経営分析調査 As String = "2"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {区分１.農業経営統計調査, "農業経営統計調査"} _
            , {区分１.組織法人経営体に関する経営分析調査, "組織法人経営体に関する経営分析調査"}
        }
    End Class

    ''' <summary>
    ''' 区分２クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 区分２
        Public Const 営農類型別経営統計 As String = "1"
        Public Const 農産物生産費 As String = "2"
        Public Const 畜産物生産費 As String = "3"

        Public Class 詳細
            Public 区分１ As String()                  '区分１
            Public 名称 As String                      '名称
        End Class

        Public Shared リスト As New Dictionary(Of String, 詳細) From {
              {区分２.営農類型別経営統計, New 詳細 With {.区分１ = {区分１.農業経営統計調査}, .名称 = "営農類型別経営統計"}} _
            , {区分２.農産物生産費, New 詳細 With {.区分１ = {区分１.農業経営統計調査, 区分１.組織法人経営体に関する経営分析調査}, .名称 = "農産物生産費"}} _
            , {区分２.畜産物生産費, New 詳細 With {.区分１ = {区分１.農業経営統計調査, 区分１.組織法人経営体に関する経営分析調査}, .名称 = "畜産物生産費"}}
        }
    End Class

    ''' <summary>
    ''' 調査区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 調査区分
        Public Const 営農類型別経営統計_個人 As String = "1"
        Public Const 営農類型別経営統計_法人 As String = "2"
        Public Const 米生産費統計_個別 As String = "3"
        Public Const 小麦生産費統計_個別 As String = "4"
        Public Const 二条大麦生産費統計_個別 As String = "5"
        Public Const 六条大麦生産費統計_個別 As String = "6"
        Public Const はだか麦生産費統計_個別 As String = "7"
        Public Const そば生産費統計_個別 As String = "8"
        Public Const 大豆生産費統計_個別 As String = "9"
        Public Const 原料用かんしょ生産費統計_個別 As String = "10"
        Public Const 原料用ばれいしょ生産費統計_個別 As String = "11"
        Public Const なたね生産費統計_個別 As String = "12"
        Public Const てんさい生産費統計_個別 As String = "13"
        Public Const さとうきび生産費統計_個別 As String = "14"
        Public Const 米生産費統計_組織法人 As String = "15"
        Public Const 小麦生産費統計_組織法人 As String = "16"
        Public Const 大豆生産費統計_組織法人 As String = "17"
        Public Const 牛乳生産費統計_個別 As String = "18"
        Public Const 子牛生産費統計_個別 As String = "19"
        Public Const 乳用雄育成牛生産費統計_個別 As String = "20"
        Public Const 交雑種育成牛生産費統計_個別 As String = "21"
        Public Const 去勢若齢肥育牛生産費統計_個別 As String = "22"
        Public Const 乳用雄肥育牛生産費統計_個別 As String = "23"
        Public Const 交雑種肥育牛生産費統計_個別 As String = "24"
        Public Const 肥育豚生産費統計_個別 As String = "25"
        Public Const 経営分析調査_二条大麦生産費 As String = "26"
        Public Const 経営分析調査_六条大麦生産費 As String = "27"
        Public Const 経営分析調査_はだか麦生産費 As String = "28"
        Public Const 経営分析調査_そば生産費 As String = "29"
        Public Const 経営分析調査_原料用ばれいしょ生産費 As String = "30"
        Public Const 経営分析調査_なたね生産費 As String = "31"
        Public Const 経営分析調査_てんさい生産費 As String = "32"
        Public Const 経営分析調査_さとうきび生産費 As String = "33"
        Public Const 経営分析調査_牛乳生産費 As String = "34"
        Public Const 経営分析調査_子牛生産費 As String = "35"
        Public Const 経営分析調査_乳用雄育成牛生産費 As String = "36"
        Public Const 経営分析調査_交雑種育成牛生産費 As String = "37"
        Public Const 経営分析調査_去勢若齢肥育牛生産費 As String = "38"
        Public Const 経営分析調査_乳用雄肥育牛生産費 As String = "39"
        Public Const 経営分析調査_交雑種肥育牛生産費 As String = "40"
        Public Const 経営分析調査_肥育豚生産費 As String = "41"

        Public Class 詳細
            Public 区分１ As String                    '区分１
            Public 区分２ As String                    '区分２
            Public 名称 As String                      '名称
        End Class

        Public Shared リスト As New Dictionary(Of String, 詳細) From {
              {調査区分.営農類型別経営統計_個人, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.営農類型別経営統計, .名称 = "営農類型別経営統計（個人）"}} _
            , {調査区分.営農類型別経営統計_法人, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.営農類型別経営統計, .名称 = "営農類型別経営統計（法人）"}} _
            , {調査区分.米生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "米生産費統計（個別）"}} _
            , {調査区分.小麦生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "小麦生産費統計（個別）"}} _
            , {調査区分.二条大麦生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "二条大麦生産費統計（個別）"}} _
            , {調査区分.六条大麦生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "六条大麦生産費統計（個別）"}} _
            , {調査区分.はだか麦生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "はだか麦生産費統計（個別）"}} _
            , {調査区分.そば生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "そば生産費統計（個別）"}} _
            , {調査区分.大豆生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "大豆生産費統計（個別）"}} _
            , {調査区分.原料用かんしょ生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "原料用かんしょ生産費統計（個別）"}} _
            , {調査区分.原料用ばれいしょ生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "原料用ばれいしょ生産費統計（個別）"}} _
            , {調査区分.なたね生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "なたね生産費統計（個別）"}} _
            , {調査区分.てんさい生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "てんさい生産費統計（個別）"}} _
            , {調査区分.さとうきび生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "さとうきび生産費統計（個別）"}} _
            , {調査区分.米生産費統計_組織法人, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "米生産費統計（組織法人）"}} _
            , {調査区分.小麦生産費統計_組織法人, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "小麦生産費統計（組織法人）"}} _
            , {調査区分.大豆生産費統計_組織法人, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.農産物生産費, .名称 = "大豆生産費統計（組織法人）"}} _
            , {調査区分.牛乳生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.畜産物生産費, .名称 = "牛乳生産費統計（個別）"}} _
            , {調査区分.子牛生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.畜産物生産費, .名称 = "子牛生産費統計（個別）"}} _
            , {調査区分.乳用雄育成牛生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.畜産物生産費, .名称 = "乳用雄育成牛生産費統計（個別）"}} _
            , {調査区分.交雑種育成牛生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.畜産物生産費, .名称 = "交雑種育成牛生産費統計（個別）"}} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.畜産物生産費, .名称 = "去勢若齢肥育牛生産費統計（個別）"}} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.畜産物生産費, .名称 = "乳用雄肥育牛生産費統計（個別）"}} _
            , {調査区分.交雑種肥育牛生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.畜産物生産費, .名称 = "交雑種肥育牛生産費統計（個別）"}} _
            , {調査区分.肥育豚生産費統計_個別, New 詳細 With {.区分１ = 区分１.農業経営統計調査, .区分２ = 区分２.畜産物生産費, .名称 = "肥育豚生産費統計（個別）"}} _
            , {調査区分.経営分析調査_二条大麦生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.農産物生産費, .名称 = "経営分析調査_二条大麦生産費"}} _
            , {調査区分.経営分析調査_六条大麦生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.農産物生産費, .名称 = "経営分析調査_六条大麦生産費"}} _
            , {調査区分.経営分析調査_はだか麦生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.農産物生産費, .名称 = "経営分析調査_はだか麦生産費"}} _
            , {調査区分.経営分析調査_そば生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.農産物生産費, .名称 = "経営分析調査_そば生産費"}} _
            , {調査区分.経営分析調査_原料用ばれいしょ生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.農産物生産費, .名称 = "経営分析調査_原料用ばれいしょ生産費"}} _
            , {調査区分.経営分析調査_なたね生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.農産物生産費, .名称 = "経営分析調査_なたね生産費"}} _
            , {調査区分.経営分析調査_てんさい生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.農産物生産費, .名称 = "経営分析調査_てんさい生産費"}} _
            , {調査区分.経営分析調査_さとうきび生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.農産物生産費, .名称 = "経営分析調査_さとうきび生産費"}} _
            , {調査区分.経営分析調査_牛乳生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.畜産物生産費, .名称 = "経営分析調査_牛乳生産費"}} _
            , {調査区分.経営分析調査_子牛生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.畜産物生産費, .名称 = "経営分析調査_子牛生産費"}} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.畜産物生産費, .名称 = "経営分析調査_乳用雄育成牛生産費"}} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.畜産物生産費, .名称 = "経営分析調査_交雑種育成牛生産費"}} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.畜産物生産費, .名称 = "経営分析調査_去勢若齢肥育牛生産費"}} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.畜産物生産費, .名称 = "経営分析調査_乳用雄肥育牛生産費"}} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.畜産物生産費, .名称 = "経営分析調査_交雑種肥育牛生産費"}} _
            , {調査区分.経営分析調査_肥育豚生産費, New 詳細 With {.区分１ = 区分１.組織法人経営体に関する経営分析調査, .区分２ = 区分２.畜産物生産費, .名称 = "経営分析調査_肥育豚生産費"}}
        }
    End Class

    ''' <summary>
    ''' 上り下り区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 上り下り区分
        Public Const 上位工程への送信 As String = "1"
        Public Const 下位工程への送信 As String = "2"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {上り下り区分.上位工程への送信, "上位工程への送信"} _
            , {上り下り区分.下位工程への送信, "下位工程への送信"}
        }
    End Class

    ''' <summary>
    ''' 審査論理データ種別クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 審査論理データ種別
        Public Const 調査票 As String = "1"
        Public Const 個別結果表 As String = "2"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {審査論理データ種別.調査票, "調査票"} _
            , {審査論理データ種別.個別結果表, "個別結果表"}
        }
    End Class

    ''' <summary>
    ''' 審査論理種別クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 審査論理種別
        Public Const 基本チェック As String = "1"
        Public Const 範囲チェック As String = "2"
        Public Const 追加チェック As String = "3"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {審査論理種別.基本チェック, "基本チェック"} _
            , {審査論理種別.範囲チェック, "範囲チェック"} _
            , {審査論理種別.追加チェック, "追加チェック"}
        }
    End Class

    ''' <summary>
    ''' 送受信データ種別クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 送受信データ種別
        Public Const 調査票_個別結果表 As String = "1"
        Public Const 集計結果表 As String = "2"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {送受信データ種別.調査票_個別結果表, "調査票、個別結果表"} _
            , {送受信データ種別.集計結果表, "集計結果表"}
        }
    End Class

    ''' <summary>
    ''' 排他種別クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 排他種別
        Public Const 個別結果表作成 As String = "1"
        Public Const 集計結果表作成 As String = "2"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {排他種別.個別結果表作成, "個別結果表作成"} _
            , {排他種別.集計結果表作成, "集計結果表作成"}
        }
    End Class

    ''' <summary>
    ''' はい・いいえ区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class はい_いいえ区分
        Public Const いいえ_または空欄 As String = "0"
        Public Const はい As String = "1"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {はい_いいえ区分.いいえ_または空欄, "いいえ（または空欄）"} _
            , {はい_いいえ区分.はい, "はい"}
        }
    End Class

    ''' <summary>
    ''' 単位区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 単位区分
        Public Const ａ As String = "1"
        Public Const m_2 As String = "2"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {単位区分.ａ, "ａ"} _
            , {単位区分.m_2, "㎡"}
        }
    End Class

    ''' <summary>
    ''' 男女区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 男女区分
        Public Const 男 As String = "1"
        Public Const 女 As String = "2"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {男女区分.男, "男"} _
            , {男女区分.女, "女"}
        }
    End Class

    ''' <summary>
    ''' 家族・雇用の別クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 家族_雇用の別
        Public Const 家族 As String = "1"
        Public Const _７ヶ月未満の雇用者 As String = "2"
        Public Const _７ヶ月以上の雇用者 As String = "3"
        Public Const 臨時雇用者 As String = "4"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {家族_雇用の別.家族, "家族"} _
            , {家族_雇用の別._７ヶ月未満の雇用者, "７ヶ月未満の雇用者"} _
            , {家族_雇用の別._７ヶ月以上の雇用者, "７ヶ月以上の雇用者"} _
            , {家族_雇用の別.臨時雇用者, "臨時雇用者"}
        }
    End Class

    ''' <summary>
    ''' チェック区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class チェック区分
        Public Const 空欄 As String = "0"
        Public Const 丸 As String = "1"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {チェック区分.空欄, "空欄"} _
            , {チェック区分.丸, "○"}
        }
    End Class

    ''' <summary>
    ''' 営農類型区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 営農類型区分
        Public Const 水田作 As String = "1"
        Public Const 畑作 As String = "2"
        Public Const 露地野菜作 As String = "3"
        Public Const 施設野菜作 As String = "4"
        Public Const 果樹作 As String = "5"
        Public Const 露地花き作 As String = "6"
        Public Const 施設花き作 As String = "7"
        Public Const 酪農 As String = "8"
        Public Const 繁殖牛 As String = "9"
        Public Const 肥育牛 As String = "10"
        Public Const 養豚 As String = "11"
        Public Const 採卵養鶏 As String = "12"
        Public Const ブロイラー養鶏 As String = "13"
        Public Const その他 As String = "14"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {営農類型区分.水田作, "水田作"} _
            , {営農類型区分.畑作, "畑作"} _
            , {営農類型区分.露地野菜作, "露地野菜作"} _
            , {営農類型区分.施設野菜作, "施設野菜作"} _
            , {営農類型区分.果樹作, "果樹作"} _
            , {営農類型区分.露地花き作, "露地花き作"} _
            , {営農類型区分.施設花き作, "施設花き作"} _
            , {営農類型区分.酪農, "酪農"} _
            , {営農類型区分.繁殖牛, "繁殖牛"} _
            , {営農類型区分.肥育牛, "肥育牛"} _
            , {営農類型区分.養豚, "養豚"} _
            , {営農類型区分.採卵養鶏, "採卵養鶏"} _
            , {営農類型区分.ブロイラー養鶏, "ブロイラー養鶏"} _
            , {営農類型区分.その他, "その他"}
        }
    End Class

    ''' <summary>
    ''' 指定品目区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 指定品目区分
        Public Const かんしょ作 As String = "1"
        Public Const ばれいしょ作 As String = "2"
        Public Const 露地きゅうり作 As String = "3"
        Public Const 露地大玉トマト作 As String = "4"
        Public Const 露地なす作 As String = "5"
        Public Const 露地キャベツ作 As String = "6"
        Public Const 露地ほうれんそう作 As String = "7"
        Public Const 露地たまねぎ作 As String = "8"
        Public Const 露地レタス作 As String = "9"
        Public Const 露地はくさい作 As String = "10"
        Public Const 露地白ねぎ作 As String = "11"
        Public Const 露地だいこん作 As String = "12"
        Public Const 露地にんじん作 As String = "13"
        Public Const 施設きゅうり作 As String = "14"
        Public Const 施設大玉トマト作 As String = "15"
        Public Const 施設ミニトマト作 As String = "16"
        Public Const 施設なす作 As String = "17"
        Public Const りんご作 As String = "18"
        Public Const 露地温州みかん作 As String = "19"
        Public Const 施設温州みかん作 As String = "20"
        Public Const 露地ぶどう作 As String = "21"
        Public Const 施設ぶどう作 As String = "22"
        Public Const 日本なし作 As String = "23"
        Public Const もも作 As String = "24"
        Public Const かき作 As String = "25"
        Public Const うめ作 As String = "26"
        Public Const おうとう作 As String = "27"
        Public Const キウイフルーツ作 As String = "28"
        Public Const すもも作 As String = "29"
        Public Const 施設ばら作 As String = "30"
        Public Const 茶作 As String = "31"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {指定品目区分.かんしょ作, "かんしょ作"} _
            , {指定品目区分.ばれいしょ作, "ばれいしょ作"} _
            , {指定品目区分.露地きゅうり作, "露地きゅうり作"} _
            , {指定品目区分.露地大玉トマト作, "露地大玉トマト作"} _
            , {指定品目区分.露地なす作, "露地なす作"} _
            , {指定品目区分.露地キャベツ作, "露地キャベツ作"} _
            , {指定品目区分.露地ほうれんそう作, "露地ほうれんそう作"} _
            , {指定品目区分.露地たまねぎ作, "露地たまねぎ作"} _
            , {指定品目区分.露地レタス作, "露地レタス作"} _
            , {指定品目区分.露地はくさい作, "露地はくさい作"} _
            , {指定品目区分.露地白ねぎ作, "露地白ねぎ作"} _
            , {指定品目区分.露地だいこん作, "露地だいこん作"} _
            , {指定品目区分.露地にんじん作, "露地にんじん作"} _
            , {指定品目区分.施設きゅうり作, "施設きゅうり作"} _
            , {指定品目区分.施設大玉トマト作, "施設大玉トマト作"} _
            , {指定品目区分.施設ミニトマト作, "施設ミニトマト作"} _
            , {指定品目区分.施設なす作, "施設なす作"} _
            , {指定品目区分.りんご作, "りんご作"} _
            , {指定品目区分.露地温州みかん作, "露地温州みかん作"} _
            , {指定品目区分.施設温州みかん作, "施設温州みかん作"} _
            , {指定品目区分.露地ぶどう作, "露地ぶどう作"} _
            , {指定品目区分.施設ぶどう作, "施設ぶどう作"} _
            , {指定品目区分.日本なし作, "日本なし作"} _
            , {指定品目区分.もも作, "もも作"} _
            , {指定品目区分.かき作, "かき作"} _
            , {指定品目区分.うめ作, "うめ作"} _
            , {指定品目区分.おうとう作, "おうとう作"} _
            , {指定品目区分.キウイフルーツ作, "キウイフルーツ作"} _
            , {指定品目区分.すもも作, "すもも作"} _
            , {指定品目区分.施設ばら作, "施設ばら作"} _
            , {指定品目区分.茶作, "茶作"}
        }
    End Class

    ''' <summary>
    ''' エラーレベルクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class エラーレベル
        Public Const エラー As String = "Z"
        Public Const ワーニング As String = "W"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {エラーレベル.エラー, "エラー"} _
            , {エラーレベル.ワーニング, "ワーニング"}
        }
    End Class

    ''' <summary>
    ''' 型区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 型区分
        Public Const 数値型 As String = "1"
        Public Const 文字型 As String = "2"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {型区分.数値型, "数値型"} _
            , {型区分.文字型, "文字型"}
        }
    End Class

    ''' <summary>
    ''' 数式区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 数式区分
        Public Const 数式ではない As String = "0"
        Public Const 数式である As String = "1"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {数式区分.数式ではない, "数式ではない"} _
            , {数式区分.数式である, "数式である"}
        }
    End Class

    ''' <summary>
    ''' 可変区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 可変区分
        Public Const 可変項目ではない As String = "0"
        Public Const 可変項目である As String = "1"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {可変区分.可変項目ではない, "可変項目ではない"} _
            , {可変区分.可変項目である, "可変項目である"}
        }
    End Class

    ''' <summary>
    ''' 可変方向クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 可変方向
        Public Const 縦 As String = "1"
        Public Const 横 As String = "2"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {可変方向.縦, "縦"} _
            , {可変方向.横, "横"}
        }
    End Class


    'REV 004 START ---------------

    ''' <summary>
    ''' バージョン区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class バージョン区分
        Public Const 調査票項目2015 As String = "1"
        Public Const 調査票項目2020 As String = "2"
        ' REV_008↓
        Public Const 結果表等項目2021 As String = "1"
        Public Const 結果表等項目2022 As String = "2"
        ' REV_008↑

        Public Shared リスト As New Dictionary(Of String, String) From {
              {バージョン区分.調査票項目2015, "2015調査票項目"} _
            , {バージョン区分.調査票項目2020, "2020調査票項目"}
        }

        ' REV_008↓
        Public Shared 体系リスト As New Dictionary(Of String, String) From {
              {バージョン区分.結果表等項目2021, "令和元年体系"} _
            , {バージョン区分.結果表等項目2022, "令和４年体系"}
        }
        ' REV_008↑
    End Class

    'REV 004 END ---------------

    ' REV_008↓
    ''' <summary>
    ''' 令和４年体系クラス
    ''' </summary>
    Public Class 令和４年体系
        Public Shared 対象調査区分2022 As New List(Of String) From
            {調査区分.営農類型別経営統計_個人 _
            , 調査区分.営農類型別経営統計_法人 _
            , 調査区分.米生産費統計_個別 _
            , 調査区分.そば生産費統計_個別 _
            , 調査区分.大豆生産費統計_個別 _
            , 調査区分.原料用かんしょ生産費統計_個別 _
            , 調査区分.原料用ばれいしょ生産費統計_個別 _
            , 調査区分.てんさい生産費統計_個別 _
            , 調査区分.さとうきび生産費統計_個別 _
            , 調査区分.米生産費統計_組織法人 _
            , 調査区分.大豆生産費統計_組織法人 _
            , 調査区分.牛乳生産費統計_個別 _
            , 調査区分.子牛生産費統計_個別 _
            , 調査区分.乳用雄育成牛生産費統計_個別 _
            , 調査区分.交雑種育成牛生産費統計_個別 _
            , 調査区分.去勢若齢肥育牛生産費統計_個別 _
            , 調査区分.乳用雄肥育牛生産費統計_個別 _
            , 調査区分.交雑種肥育牛生産費統計_個別 _
            , 調査区分.肥育豚生産費統計_個別 _
            , 営農経営体区分.農業経営体                  ' REV_010
            }

        Public Shared 対象調査区分2023 As New List(Of String) From
            {調査区分.小麦生産費統計_個別 _
            , 調査区分.二条大麦生産費統計_個別 _
            , 調査区分.六条大麦生産費統計_個別 _
            , 調査区分.はだか麦生産費統計_個別 _
            , 調査区分.なたね生産費統計_個別 _
            , 調査区分.小麦生産費統計_組織法人
            }
    End Class
    ' REV_008↑

    ''' <summary>
    ''' 生産費区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 生産費区分
        Public Const 米_畜産物等_当年１月_当年12月 As String = "1"
        Public Const 麦類_なたね_前年９月_当年８月 As String = "2"
        Public Const さとうきび_当年４月_翌年３月 As String = "3"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {生産費区分.米_畜産物等_当年１月_当年12月, "米、畜産物等（当年１月～当年12月）"} _
            , {生産費区分.麦類_なたね_前年９月_当年８月, "麦類、なたね（前年９月～当年８月）"} _
            , {生産費区分.さとうきび_当年４月_翌年３月, "さとうきび（当年４月～翌年３月）"}
        }
    End Class

    ''' <summary>
    ''' 欠測値補完クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 欠測値補完
        Public Const 無 As String = "0"
        Public Const 有 As String = "1"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {欠測値補完.無, "無"} _
            , {欠測値補完.有, "有"}
        }
    End Class

    ''' <summary>
    ''' 貸借対照表区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 貸借対照表区分
        Public Const _1 As String = "1"
        Public Const _2 As String = "2"
        Public Const _3 As String = "3"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {貸借対照表区分._1, "1"} _
            , {貸借対照表区分._2, "2"} _
            , {貸借対照表区分._3, "3"}
        }
    End Class

    ''' <summary>
    ''' 営農経営体区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 営農経営体区分
        Public Const 個人経営体 As String = 調査区分.営農類型別経営統計_個人
        Public Const 法人経営体 As String = 調査区分.営農類型別経営統計_法人
        Public Const 農業経営体 As String = "42"

        Public Class 詳細
            Public 調査区分 As String()                '調査区分
            Public 名称 As String                      '名称
            Public 名称２ As String                      '名称２
        End Class

        Public Shared リスト As New Dictionary(Of String, 詳細) From {
              {営農経営体区分.個人経営体, New 詳細 With {.調査区分 = {調査区分.営農類型別経営統計_個人}, .名称 = "個人経営体", .名称２ = 調査区分.リスト(営農経営体区分.個人経営体).名称}} _
            , {営農経営体区分.法人経営体, New 詳細 With {.調査区分 = {調査区分.営農類型別経営統計_法人}, .名称 = "法人経営体", .名称２ = 調査区分.リスト(営農経営体区分.法人経営体).名称}} _
            , {営農経営体区分.農業経営体, New 詳細 With {.調査区分 = {調査区分.営農類型別経営統計_個人, 調査区分.営農類型別経営統計_法人}, .名称 = "農業経営体", .名称２ = "営農類型別経営統計（農業経営体）"}}
        }
    End Class

    ''' <summary>
    ''' 平均種類クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 平均種類
        Public Const 加重 As String = "1"
        Public Const 総和 As String = "2"
        Public Const 累積 As String = "3"

        Public Class 詳細
            Public 区分２ As String()                  '区分２
            Public 名称 As String                      '名称
        End Class

        Public Shared リスト As New Dictionary(Of String, 詳細) From {
              {平均種類.加重, New 詳細 With {.区分２ = {区分２.営農類型別経営統計, 区分２.農産物生産費, 区分２.畜産物生産費}, .名称 = "加重"}} _
            , {平均種類.総和, New 詳細 With {.区分２ = {区分２.営農類型別経営統計, 区分２.農産物生産費, 区分２.畜産物生産費}, .名称 = "総和"}} _
            , {平均種類.累積, New 詳細 With {.区分２ = {区分２.営農類型別経営統計}, .名称 = "累積"}}
        }
    End Class

    ''' <summary>
    ''' 任意階層利用クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 任意階層利用
        Public Const 無 As String = "1"
        Public Const 有 As String = "2"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {任意階層利用.無, "無"} _
            , {任意階層利用.有, "有"}
        }
    End Class

    ''' <summary>
    ''' 規模階層クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 規模階層
        Public Const 規模階層含 As String = "1"
        Public Const 平均のみ As String = "2"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {規模階層.規模階層含, "規模階層含"} _
            , {規模階層.平均のみ, "平均値のみ"}
        }
    End Class

    ''' <summary>
    ''' 集計区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 集計区分
        Public Const 集計対象 As String = "2"
        Public Const 全調査対象 As String = "1"

        Public Shared リスト As New Dictionary(Of String, String) From {
               {集計区分.集計対象, "集計対象"} _
             , {集計区分.全調査対象, "全調査対象"}
        }
    End Class

    ''' <summary>
    ''' 田畑区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 田畑区分
        Public Const 全て As String = "4"
        Public Const 田 As String = "1"
        Public Const 畑 As String = "2"
        Public Const 田畑計 As String = "3"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {田畑区分.全て, "全て"} _
            , {田畑区分.田, "田"} _
            , {田畑区分.畑, "畑"} _
            , {田畑区分.田畑計, "田畑計"}
        }
    End Class

    ''' <summary>
    ''' ビール麦販売区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ビール麦販売区分
        Public Const 全て As String = "5"
        Public Const ビール麦100 As String = "1"
        Public Const ビール麦100未満 As String = "2"
        Public Const ビール麦無 As String = "3"
        Public Const 指定しない As String = "4"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {ビール麦販売区分.全て, "全て"} _
            , {ビール麦販売区分.ビール麦100, "ビール麦100％"} _
            , {ビール麦販売区分.ビール麦100未満, "ビール麦100％未満"} _
            , {ビール麦販売区分.ビール麦無, "ビール麦無"} _
            , {ビール麦販売区分.指定しない, "指定しない"}
        }
    End Class

    ''' <summary>
    ''' てんさい栽培区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class てんさい栽培区分
        Public Const 全て As String = "4"
        Public Const 移植有 As String = "1"
        Public Const 移植無 As String = "2"
        Public Const 指定しない As String = "3"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {てんさい栽培区分.全て, "全て"} _
            , {てんさい栽培区分.移植有, "移植有"} _
            , {てんさい栽培区分.移植無, "移植無"} _
            , {てんさい栽培区分.指定しない, "指定しない"}
        }
    End Class

    ''' <summary>
    ''' 集計１クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 集計１
        Public Const 個人経営体_第１報用_分類別平均 As String = "1"
        Public Const 個人経営体 As String = "2"
        Public Const 個人経営体_経営形態別 As String = "3"
        Public Const 個人経営体_経営形態別_分類別平均 As String = "4"
        Public Const 個人経営体_分類別平均 As String = "5"
        Public Const 法人経営体_第１報用_分類別平均 As String = "6"
        Public Const 法人経営体 As String = "7"
        Public Const 法人経営体_経営形態別 As String = "8"
        Public Const 法人経営体_経営形態別_分類別平均 As String = "9"
        Public Const 法人経営体のうち組織法人経営 As String = "10"
        Public Const 法人経営体のうち組織法人経営_分類別平均 As String = "11"
        Public Const 法人経営体_分類別平均 As String = "12"
        Public Const 農業経営体_第１報用_分類別平均 As String = "13"
        Public Const 農業経営体 As String = "14"
        Public Const 農業経営体_経営形態別 As String = "15"
        Public Const 農業経営体_経営形態別_分類別平均 As String = "16"
        Public Const 農業経営体_分類別平均 As String = "17"

        Public Class 詳細
            Public 営農経営体区分 As String          '営農経営体区分
            Public 名称 As String                    '名称
        End Class

        '---REV_012↓
        Private Shared リスト_2021 As New Dictionary(Of String, 詳細) From {
              {集計１.個人経営体_第１報用_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.個人経営体, .名称 = "第１報用_分類別平均"}} _
            , {集計１.個人経営体_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.個人経営体, .名称 = "個人経営体（分類別平均）"}} _
            , {集計１.個人経営体, New 詳細 With {.営農経営体区分 = 営農経営体区分.個人経営体, .名称 = "個人経営体"}} _
            , {集計１.個人経営体_経営形態別_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.個人経営体, .名称 = "個人経営体　経営形態別（分類別平均）"}} _
            , {集計１.個人経営体_経営形態別, New 詳細 With {.営農経営体区分 = 営農経営体区分.個人経営体, .名称 = "個人経営体　経営形態別"}} _
            , {集計１.法人経営体_第１報用_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "第１報用（分類別平均）"}} _
            , {集計１.法人経営体_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "法人経営体（分類別平均）"}} _
            , {集計１.法人経営体, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "法人経営体"}} _
            , {集計１.法人経営体のうち組織法人経営_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "法人経営体のうち組織法人経営（分類別平均）"}} _
            , {集計１.法人経営体のうち組織法人経営, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "法人経営体のうち組織法人経営"}} _
            , {集計１.法人経営体_経営形態別_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "法人経営体 経営形態別（分類別平均）"}} _
            , {集計１.法人経営体_経営形態別, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "法人経営体 経営形態別"}} _
            , {集計１.農業経営体_第１報用_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.農業経営体, .名称 = "第１報用（分類別平均）"}} _
            , {集計１.農業経営体_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.農業経営体, .名称 = "農業経営体（分類別平均）"}} _
            , {集計１.農業経営体, New 詳細 With {.営農経営体区分 = 営農経営体区分.農業経営体, .名称 = "農業経営体"}} _
            , {集計１.農業経営体_経営形態別_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.農業経営体, .名称 = "農業経営体　経営形態別（分類別平均）"}} _
            , {集計１.農業経営体_経営形態別, New 詳細 With {.営農経営体区分 = 営農経営体区分.農業経営体, .名称 = "農業経営体 経営形態別"}}
          }
        Private Shared リスト_2022 As New Dictionary(Of String, 詳細) From {
              {集計１.個人経営体_第１報用_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.個人経営体, .名称 = "第１報用_分類別平均"}} _
            , {集計１.個人経営体_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.個人経営体, .名称 = "個人経営体（分類別平均）"}} _
            , {集計１.個人経営体, New 詳細 With {.営農経営体区分 = 営農経営体区分.個人経営体, .名称 = "個人経営体"}} _
            , {集計１.個人経営体_経営形態別_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.個人経営体, .名称 = "個人経営体　経営形態別（分類別平均）"}} _
            , {集計１.個人経営体_経営形態別, New 詳細 With {.営農経営体区分 = 営農経営体区分.個人経営体, .名称 = "個人経営体　経営形態別"}} _
            , {集計１.法人経営体_第１報用_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "第１報用（分類別平均）"}} _
            , {集計１.法人経営体_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "法人経営体（分類別平均）"}} _
            , {集計１.法人経営体, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "法人経営体"}} _
            , {集計１.法人経営体のうち組織法人経営_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "法人経営体のうち組織法人経営（分類別平均）"}} _
            , {集計１.法人経営体のうち組織法人経営, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "法人経営体のうち組織法人経営"}} _
            , {集計１.法人経営体_経営形態別_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "法人経営体 経営形態別（分類別平均）"}} _
            , {集計１.法人経営体_経営形態別, New 詳細 With {.営農経営体区分 = 営農経営体区分.法人経営体, .名称 = "法人経営体 経営形態別"}} _
            , {集計１.農業経営体_第１報用_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.農業経営体, .名称 = "第１報用（分類別平均）"}} _
            , {集計１.農業経営体_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.農業経営体, .名称 = "農業経営体（分類別平均）"}} _
            , {集計１.農業経営体, New 詳細 With {.営農経営体区分 = 営農経営体区分.農業経営体, .名称 = "農業経営体"}} _
            , {集計１.農業経営体_経営形態別_分類別平均, New 詳細 With {.営農経営体区分 = 営農経営体区分.農業経営体, .名称 = "農業経営体　経営形態別（分類別平均）"}} _
            , {集計１.農業経営体_経営形態別, New 詳細 With {.営農経営体区分 = 営農経営体区分.農業経営体, .名称 = "農業経営体 経営形態別"}}
          }

        Public Shared Function リスト(versionKbn As String) As Dictionary(Of String, 詳細)
            If ComUtil.IsEinou Then
                '営農は2021年以前と2022年以降で切り替える
                If versionKbn = ComConst.バージョン区分.結果表等項目2022 Then
                    Return リスト_2022
                Else
                    Return リスト_2021
                End If
            Else
                Return リスト_2021
            End If
        End Function
        '---REV_012↑
    End Class

    ''' <summary>
    ''' 集計２クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 集計２
        Public Const 項目なし As String = "0"
        Public Const 個人_ブロイラー養鶏経営 As String = "1"
        Public Const 法人_ブロイラー養鶏経営 As String = "2"
        Public Const 主副業別_経営 As String = "3"
        Public Const 個別法人経営 As String = "4"
        Public Const 平均 As String = "5"
        Public Const 個人_採卵養鶏経営 As String = "6"
        Public Const 法人_採卵養鶏経営 As String = "7"
        Public Const 個人_果樹作経営 As String = "8"
        Public Const 法人_果樹作経営 As String = "9"
        Public Const 果樹作経営_分類別平均 As String = "10"
        Public Const 個人_水田作経営 As String = "11"
        Public Const 法人_水田作経営 As String = "12"
        Public Const 水田作経営のうち集落営農組織法人 As String = "13"
        Public Const 水田作経営_分類別平均 As String = "14"
        Public Const 個人_畑作経営 As String = "15"
        Public Const 法人_畑作経営 As String = "16"
        Public Const 農業経営体_畑作経営 As String = "17"
        Public Const 畑作経営_分類別平均 As String = "18"
        Public Const 個人_肉用牛経営 As String = "19"
        Public Const 法人_肉用牛経営 As String = "20"
        Public Const 肉用牛経営_分類別平均 As String = "21"
        Public Const 個人_花き作経営 As String = "22"
        Public Const 法人_花き作経営 As String = "23"
        Public Const 花き作経営_分類別平均 As String = "24"
        Public Const 認定農業者のいる経営平均 As String = "25"
        Public Const 農業地域類型別 As String = "26"
        Public Const 農業生産関連事業あり As String = "27"
        Public Const 個人_酪農経営 As String = "28"
        Public Const 法人_酪農経営 As String = "29"
        Public Const 個人_野菜作経営 As String = "30"
        Public Const 法人_野菜作経営 As String = "31"
        Public Const 野菜作経営_分類別平均 As String = "32"
        Public Const 関連事業を行っている経営平均 As String = "33"
        Public Const 個人_養豚経営 As String = "34"
        Public Const 法人_養豚経営 As String = "35"
        '---REV_012↓
        Public Const 個人_水田作経営_分類別平均 As String = "36"
        Public Const 個人_畑作経営_分類別平均 As String = "37"
        Public Const 個人_野菜作経営_分類別平均 As String = "38"
        Public Const 個人_果樹作経営_分類別平均 As String = "39"
        Public Const 個人_花き作経営_分類別平均 As String = "40"
        Public Const 個人_肉用牛経営_分類別平均 As String = "41"
        Public Const 青色申告を行っている経営平均_基本項目集計 As String = "42"
        Public Const 青色申告を行っている経営平均_詳細項目集計 As String = "43"
        Public Const 経営平均 As String = "44"
        Public Const 平均_基本項目集計 As String = "45"
        Public Const 平均_詳細項目集計 As String = "46"
        Public Const 認定農業者のいる経営平均_基本項目集計 As String = "47"
        '---REV_012↑

        Public Class 詳細
            Public 営農経営体区分 As String()          '営農経営体区分
            Public 集計１ As String()                  '集計１
            Public 名称 As String                      '名称
        End Class

        '---REV_012↓
        Private Shared リスト_2021 As New Dictionary(Of String, 詳細) From {
               {集計２.項目なし, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体_第１報用_分類別平均, 集計１.個人経営体_分類別平均, 集計１.個人経営体_経営形態別_分類別平均, 集計１.法人経営体_第１報用_分類別平均, 集計１.法人経営体_分類別平均, 集計１.法人経営体のうち組織法人経営_分類別平均, 集計１.法人経営体_経営形態別_分類別平均, 集計１.農業経営体_第１報用_分類別平均, 集計１.農業経営体_分類別平均, 集計１.農業経営体_経営形態別_分類別平均}, .名称 = "‐"}} _
            , {集計２.水田作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.農業経営体}, .名称 = "水田作経営（分類別平均）"}} _
            , {集計２.個人_水田作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "水田作経営"}} _
            , {集計２.法人_水田作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "水田作経営"}} _
            , {集計２.水田作経営のうち集落営農組織法人, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体}, .名称 = "水田作経営のうち集落営農組織法人"}} _
            , {集計２.畑作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.農業経営体}, .名称 = "畑作経営（分類別平均）"}} _
            , {集計２.個人_畑作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "畑作経営"}} _
            , {集計２.法人_畑作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営}, .名称 = "畑作経営"}} _
            , {集計２.農業経営体_畑作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.農業経営体}, .集計１ = {集計１.農業経営体}, .名称 = "畑作経営"}} _
            , {集計２.野菜作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.農業経営体}, .名称 = "野菜作経営（分類別平均）"}} _
            , {集計２.個人_野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "野菜作経営"}} _
            , {集計２.法人_野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "野菜作経営"}} _
            , {集計２.果樹作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "果樹作経営（分類別平均）"}} _
            , {集計２.個人_果樹作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "果樹作経営"}} _
            , {集計２.法人_果樹作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "果樹作経営"}} _
            , {集計２.花き作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.農業経営体}, .名称 = "花き作経営（分類別平均）"}} _
            , {集計２.個人_花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "花き作経営"}} _
            , {集計２.法人_花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "花き作経営"}} _
            , {集計２.個人_酪農経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "酪農経営"}} _
            , {集計２.法人_酪農経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "酪農経営"}} _
            , {集計２.肉用牛経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.農業経営体}, .名称 = "肉用牛経営（分類別平均）"}} _
            , {集計２.個人_肉用牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "肉用牛経営"}} _
            , {集計２.法人_肉用牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "肉用牛経営"}} _
            , {集計２.個人_養豚経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "養豚経営"}} _
            , {集計２.法人_養豚経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "養豚経営"}} _
            , {集計２.個人_採卵養鶏経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "採卵養鶏経営"}} _
            , {集計２.法人_採卵養鶏経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "採卵養鶏経営"}} _
            , {集計２.個人_ブロイラー養鶏経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "ブロイラー養鶏経営"}} _
            , {集計２.法人_ブロイラー養鶏経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "ブロイラー養鶏経営"}} _
            , {集計２.平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体_経営形態別, 集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .名称 = "平均"}} _
            , {集計２.主副業別_経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .名称 = "主副業別（経営）"}} _
            , {集計２.農業地域類型別, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .名称 = "農業地域類型別"}} _
            , {集計２.認定農業者のいる経営平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .名称 = "認定農業者のいる経営平均"}} _
            , {集計２.農業生産関連事業あり, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .名称 = "農業生産関連事業あり"}} _
            , {集計２.個別法人経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .名称 = "個別法人経営"}} _
            , {集計２.関連事業を行っている経営平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .名称 = "関連事業を行っている経営平均"}}
       }

        Private Shared リスト_2022 As New Dictionary(Of String, 詳細) From {
              {集計２.項目なし, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体_第１報用_分類別平均, 集計１.個人経営体_分類別平均, 集計１.個人経営体_経営形態別_分類別平均, 集計１.法人経営体_第１報用_分類別平均, 集計１.法人経営体_分類別平均, 集計１.法人経営体のうち組織法人経営_分類別平均, 集計１.法人経営体_経営形態別_分類別平均, 集計１.農業経営体_第１報用_分類別平均, 集計１.農業経営体_分類別平均, 集計１.農業経営体_経営形態別_分類別平均}, .名称 = "‐"}} _
            , {集計２.個人_水田作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "水田作経営_分類別平均"}} _
            , {集計２.個人_水田作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "水田作経営"}} _
            , {集計２.個人_畑作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "畑作経営_分類別平均"}} _
            , {集計２.個人_畑作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "畑作経営"}} _
            , {集計２.個人_野菜作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "野菜作経営_分類別平均"}} _
            , {集計２.個人_野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "野菜作経営"}} _
            , {集計２.個人_果樹作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "果樹作経営_分類別平均"}} _
            , {集計２.個人_果樹作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "果樹作経営"}} _
            , {集計２.個人_花き作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "花き作経営_分類別平均"}} _
            , {集計２.個人_花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "花き作経営"}} _
            , {集計２.個人_酪農経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "酪農経営"}} _
            , {集計２.個人_肉用牛経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "肉用牛経営_分類別平均"}} _
            , {集計２.個人_肉用牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "肉用牛経営"}} _
            , {集計２.個人_養豚経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "養豚経営"}} _
            , {集計２.個人_採卵養鶏経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "採卵養鶏経営"}} _
            , {集計２.個人_ブロイラー養鶏経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .名称 = "ブロイラー養鶏経営"}} _
            , {集計２.平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .名称 = "平均"}} _
            , {集計２.平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .名称 = "平均（基本項目集計）"}} _
            , {集計２.平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .名称 = "平均（詳細項目集計）"}} _
            , {集計２.主副業別_経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .名称 = "主副業別（経営）"}} _
            , {集計２.農業地域類型別, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体_経営形態別, 集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .名称 = "農業地域類型別"}} _
            , {集計２.青色申告を行っている経営平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .名称 = "青色申告を行っている経営平均（基本項目集計）"}} _
            , {集計２.青色申告を行っている経営平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .名称 = "青色申告を行っている経営平均（詳細項目集計）"}} _
            , {集計２.認定農業者のいる経営平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .名称 = "認定農業者のいる経営平均（基本項目集計）"}} _
            , {集計２.水田作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.農業経営体}, .集計１ = {集計１.農業経営体}, .名称 = "水田作経営（分類別平均）"}} _
            , {集計２.法人_水田作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "水田作経営"}} _
            , {集計２.水田作経営のうち集落営農組織法人, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体}, .名称 = "水田作経営のうち集落営農組織法人"}} _
            , {集計２.畑作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.農業経営体}, .集計１ = {集計１.農業経営体}, .名称 = "畑作経営（分類別平均）"}} _
            , {集計２.法人_畑作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営}, .名称 = "畑作経営"}} _
            , {集計２.農業経営体_畑作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.農業経営体}, .集計１ = {集計１.農業経営体}, .名称 = "畑作経営"}} _
            , {集計２.野菜作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.農業経営体}, .名称 = "野菜作経営（分類別平均）"}} _
            , {集計２.法人_野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "野菜作経営"}} _
            , {集計２.法人_果樹作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "果樹作経営"}} _
            , {集計２.花き作経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.農業経営体}, .集計１ = {集計１.農業経営体}, .名称 = "花き作経営（分類別平均）"}} _
            , {集計２.法人_花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "花き作経営"}} _
            , {集計２.法人_酪農経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "酪農経営"}} _
            , {集計２.肉用牛経営_分類別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.農業経営体}, .集計１ = {集計１.農業経営体}, .名称 = "肉用牛経営（分類別平均）"}} _
            , {集計２.法人_肉用牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "肉用牛経営"}} _
            , {集計２.法人_養豚経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "養豚経営"}} _
            , {集計２.法人_採卵養鶏経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "採卵養鶏経営"}} _
            , {集計２.法人_ブロイラー養鶏経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .名称 = "ブロイラー養鶏経営"}} _
            , {集計２.経営平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体のうち組織法人経営}, .名称 = "経営平均"}} _
            , {集計２.関連事業を行っている経営平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体のうち組織法人経営, 集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .名称 = "関連事業を行っている経営平均"}} _
            , {集計２.農業生産関連事業あり, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .名称 = "農業生産関連事業あり"}} _
            , {集計２.個別法人経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .名称 = "個別法人経営"}}
        }

        Public Shared Function リスト(versionKbn As String) As Dictionary(Of String, 詳細)
            If ComUtil.IsEinou Then
                '営農は2021年以前と2022年以降で切り替える
                If versionKbn = ComConst.バージョン区分.結果表等項目2022 Then
                    Return リスト_2022
                Else
                    Return リスト_2021
                End If
            Else
                Return リスト_2021
            End If
        End Function
        '---REV_012↑
    End Class

    ''' <summary>
    ''' 集計３クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 集計３
        Public Const 項目なし As String = "0"
        Public Const うめ作部門 As String = "1"
        Public Const おうとう作部門 As String = "2"
        Public Const かき作部門 As String = "3"
        Public Const かんしょ作経営 As String = "4"
        Public Const かんしょ作部門 As String = "5"
        Public Const さとうきび作経営 As String = "6"
        Public Const すもも作部門 As String = "7"
        Public Const てんさい作経営 As String = "8"
        Public Const ばれいしょ作経営 As String = "9"
        Public Const ばれいしょ作部門 As String = "10"
        Public Const もも作部門 As String = "11"
        Public Const りんご作部門 As String = "12"
        Public Const キウイフルーツ作部門 As String = "13"
        Public Const 中山間農業地域 As String = "14"
        Public Const 中間農業地域 As String = "15"
        Public Const 主副業別平均 As String = "16"
        Public Const 主業 As String = "17"
        Public Const 主業平均 As String = "18"
        Public Const 主業経営 As String = "19"
        Public Const 副業的 As String = "20"
        Public Const 個人_水田作_大豆作経営 As String = "21"
        Public Const 個人_畑作_大豆作経営 As String = "22"
        Public Const 法人_大豆作経営 As String = "23"
        Public Const 大豆作１位経営 As String = "24"
        Public Const 山間農業地域 As String = "25"
        Public Const 平地農業地域 As String = "26"
        Public Const 施設ぶどう作部門 As String = "27"
        Public Const 施設温州みかん作部門 As String = "28"
        Public Const 個人_施設花き作経営 As String = "29"
        Public Const 法人_施設花き作経営 As String = "30"
        Public Const 個人_施設野菜作経営 As String = "31"
        Public Const 法人_施設野菜作経営 As String = "32"
        Public Const 日本なし作部門 As String = "33"
        Public Const 果樹作単一 As String = "34"
        Public Const 準主業 As String = "35"
        Public Const 稲作単一経営 As String = "36"
        Public Const 個人_稲作経営 As String = "37"
        Public Const 法人_稲作経営 As String = "38"
        Public Const 稲作１位経営 As String = "39"
        Public Const 稲作１位複合経営 As String = "40"
        Public Const 経営平均 As String = "41"
        Public Const 個人_繁殖牛経営 As String = "42"
        Public Const 法人_繁殖牛経営 As String = "43"
        Public Const 肉用牛経営 As String = "44"
        Public Const 個人_肥育牛経営 As String = "45"
        Public Const 法人_肥育牛経営 As String = "46"
        Public Const 花き作経営 As String = "47"
        Public Const 茶作単一経営 As String = "48"
        Public Const 茶作経営 As String = "49"
        Public Const 茶作部門 As String = "50"
        Public Const 観光農園 As String = "51"
        Public Const 観光農園を行っている経営 As String = "52"
        Public Const 貸し農園 As String = "53"
        Public Const 貸し農園を行っている経営 As String = "54"
        Public Const 農家レストラン As String = "55"
        Public Const 農家レストランを行っている経営 As String = "56"
        Public Const 農家民宿 As String = "57"
        Public Const 農家民宿を行っている経営 As String = "58"
        Public Const 農業地域類型別 As String = "59"
        Public Const 農業生産関連事業平均 As String = "60"
        Public Const 農産加工 As String = "61"
        Public Const 農産加工を行っている経営 As String = "62"
        Public Const 都市的地域 As String = "63"
        Public Const 野菜作経営 As String = "64"
        Public Const 露地ぶどう作部門 As String = "65"
        Public Const 露地温州みかん作部門 As String = "66"
        Public Const 個人_露地花き作経営 As String = "67"
        Public Const 法人_露地花き作経営 As String = "68"
        Public Const 個人_露地野菜作経営 As String = "69"
        Public Const 法人_露地野菜作経営 As String = "70"
        Public Const 個人_水田作_麦類作経営 As String = "71"
        Public Const 個人_畑作_麦類作経営 As String = "72"
        Public Const 法人_麦類作経営 As String = "73"
        Public Const 麦類作１位経営 As String = "74"
        '---REV_012↓
        Public Const 経営平均_基本項目集計 As String = "76"
        Public Const 経営平均_詳細項目集計 As String = "77"
        Public Const 主業平均_基本項目集計 As String = "78"
        Public Const 青色申告を行っている経営平均_詳細項目集計 As String = "79"
        Public Const ばれいしょ作部門_詳細項目集計 As String = "80"
        Public Const 茶作部門_詳細項目集計 As String = "81"
        Public Const かんしょ作部門_詳細項目集計 As String = "82"
        Public Const 個人_畑作_麦類作経営_基本項目集計 As String = "83"
        Public Const さとうきび作経営_基本項目集計 As String = "84"
        Public Const 個人_畑作_大豆作経営_基本項目集計 As String = "85"
        Public Const てんさい作経営_基本項目集計 As String = "86"
        Public Const ばれいしょ作経営_基本項目集計 As String = "87"
        Public Const 茶作経営_基本項目集計 As String = "88"
        Public Const かんしょ作経営_基本項目集計 As String = "89"
        Public Const りんご作部門_詳細項目集計 As String = "90"
        Public Const 日本なし作部門_詳細項目集計 As String = "91"
        Public Const もも作部門_詳細項目集計 As String = "92"
        Public Const 露地温州みかん作部門_詳細項目集計 As String = "93"
        Public Const 施設温州みかん作部門_詳細項目集計 As String = "94"
        Public Const 露地ぶどう作部門_詳細項目集計 As String = "95"
        Public Const 施設ぶどう作部門_詳細項目集計 As String = "96"
        Public Const かき作部門_詳細項目集計 As String = "97"
        Public Const うめ作部門_詳細項目集計 As String = "98"
        Public Const おうとう作部門_詳細項目集計 As String = "99"
        Public Const キウイフルーツ作部門_詳細項目集計 As String = "100"
        Public Const すもも作部門_詳細項目集計 As String = "101"
        Public Const 農業地域類型別_基本項目集計 As String = "102"
        Public Const 都市的地域_基本項目集計 As String = "103"
        Public Const 平地農業地域_基本項目集計 As String = "104"
        Public Const 中山間農業地域_基本項目集計 As String = "105"
        Public Const 中間農業地域_基本項目集計 As String = "106"
        Public Const 山間農業地域_基本項目集計 As String = "107"
        Public Const 関連事業を行っている経営平均 As String = "108"
        Public Const 部門別平均 As String = "109"
        Public Const 農業生産関連事業平均_基本項目集計 As String = "110"
        Public Const 主業経営_基本項目集計 As String = "111"
        Public Const 主副業別平均_基本項目集計 As String = "112"
        Public Const 果樹作単一_基本項目集計 As String = "113"
        '---REV_012↑

        Public Class 詳細
            Public 営農経営体区分 As String()          '営農経営体区分
            Public 集計１ As String()                  '集計１
            Public 集計２ As String()                  '集計２
            Public 名称 As String                      '名称
        End Class

        '---REV_012↓
        Private Shared リスト_2021 As New Dictionary(Of String, 詳細) From {
              {集計３.項目なし, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体_第１報用_分類別平均, 集計１.個人経営体, 集計１.個人経営体_経営形態別, 集計１.個人経営体_分類別平均, 集計１.個人経営体_経営形態別_分類別平均, 集計１.法人経営体_第１報用_分類別平均, 集計１.法人経営体, 集計１.法人経営体_経営形態別, 集計１.法人経営体_経営形態別_分類別平均, 集計１.法人経営体のうち組織法人経営, 集計１.法人経営体のうち組織法人経営_分類別平均, 集計１.法人経営体_分類別平均, 集計１.農業経営体, 集計１.農業経営体_第１報用_分類別平均, 集計１.農業経営体_分類別平均, 集計１.農業経営体_経営形態別, 集計１.農業経営体_経営形態別_分類別平均}, .集計２ = {集計２.法人_ブロイラー養鶏経営, 集計２.平均, 集計２.法人_採卵養鶏経営, 集計２.法人_果樹作経営, 集計２.果樹作経営_分類別平均, 集計２.水田作経営_分類別平均, 集計２.畑作経営_分類別平均, 集計２.肉用牛経営_分類別平均, 集計２.花き作経営_分類別平均, 集計２.認定農業者のいる経営平均, 集計２.野菜作経営_分類別平均, 集計２.法人_酪農経営, 集計２.法人_養豚経営, 集計２.項目なし}, .名称 = "‐"}} _
            , {集計３.経営平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.法人経営体_経営形態別, 集計１.農業経営体, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_畑作経営, 集計２.個人_野菜作経営, 集計２.個人_果樹作経営, 集計２.個人_花き作経営, 集計２.個人_酪農経営, 集計２.個人_肉用牛経営, 集計２.個人_養豚経営, 集計２.個人_採卵養鶏経営, 集計２.個人_ブロイラー養鶏経営, 集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人, 集計２.法人_畑作経営, 集計２.法人_水田作経営, 集計２.関連事業を行っている経営平均, 集計２.個別法人経営, 集計２.農業経営体_畑作経営}, .名称 = "経営平均"}} _
            , {集計３.主業平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_畑作経営, 集計２.個人_野菜作経営, 集計２.個人_果樹作経営}, .名称 = "主業平均"}} _
            , {集計３.主業経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .名称 = "主業経営"}} _
            , {集計３.個人_稲作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作経営"}} _
            , {集計３.法人_稲作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作経営"}} _
            , {集計３.稲作１位経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作１位経営"}} _
            , {集計３.稲作単一経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作単一経営"}} _
            , {集計３.稲作１位複合経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作１位複合経営"}} _
            , {集計３.個人_水田作_麦類作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .名称 = "麦類作経営"}} _
            , {集計３.法人_麦類作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.水田作経営のうち集落営農組織法人, 集計２.法人_水田作経営, 集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "麦類作経営"}} _
            , {集計３.麦類作１位経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "麦類作１位経営"}} _
            , {集計３.個人_水田作_大豆作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .名称 = "大豆作経営"}} _
            , {集計３.大豆作１位経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "大豆作１位経営"}} _
            , {集計３.ばれいしょ作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "ばれいしょ作部門"}} _
            , {集計３.茶作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "茶作部門"}} _
            , {集計３.かんしょ作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "かんしょ作部門"}} _
            , {集計３.個人_畑作_麦類作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "麦類作経営"}} _
            , {集計３.さとうきび作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.個人_畑作経営, 集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "さとうきび作経営"}} _
            , {集計３.個人_畑作_大豆作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "大豆作経営"}} _
            , {集計３.法人_大豆作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人, 集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "大豆作経営"}} _
            , {集計３.てんさい作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.個人_畑作経営, 集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "てんさい作経営"}} _
            , {集計３.ばれいしょ作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.個人_畑作経営, 集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "ばれいしょ作経営"}} _
            , {集計３.茶作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.個人_畑作経営, 集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "茶作経営"}} _
            , {集計３.かんしょ作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.個人_畑作経営, 集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "かんしょ作経営"}} _
            , {集計３.野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_野菜作経営}, .名称 = "野菜作経営"}} _
            , {集計３.個人_露地野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .名称 = "露地野菜作経営"}} _
            , {集計３.法人_露地野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_野菜作経営}, .名称 = "露地野菜作経営"}} _
            , {集計３.個人_施設野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .名称 = "施設野菜作経営"}} _
            , {集計３.法人_施設野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_野菜作経営}, .名称 = "施設野菜作経営"}} _
            , {集計３.果樹作単一, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "果樹作単一"}} _
            , {集計３.りんご作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "りんご作部門"}} _
            , {集計３.日本なし作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "日本なし作部門"}} _
            , {集計３.もも作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "もも作部門"}} _
            , {集計３.露地温州みかん作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "露地温州みかん作部門"}} _
            , {集計３.施設温州みかん作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "施設温州みかん作部門"}} _
            , {集計３.露地ぶどう作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "露地ぶどう作部門"}} _
            , {集計３.施設ぶどう作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "施設ぶどう作部門"}} _
            , {集計３.かき作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "かき作部門"}} _
            , {集計３.うめ作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "うめ作部門"}} _
            , {集計３.おうとう作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "おうとう作部門"}} _
            , {集計３.キウイフルーツ作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "キウイフルーツ作部門"}} _
            , {集計３.すもも作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "すもも作部門"}} _
            , {集計３.花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_花き作経営}, .名称 = "花き作経営"}} _
            , {集計３.個人_露地花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .名称 = "露地花き作経営"}} _
            , {集計３.法人_露地花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_花き作経営}, .名称 = "露地花き作経営"}} _
            , {集計３.個人_施設花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .名称 = "施設花き作経営"}} _
            , {集計３.法人_施設花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_花き作経営}, .名称 = "施設花き作経営"}} _
            , {集計３.肉用牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.個人_肉用牛経営, 集計２.法人_肉用牛経営}, .名称 = "肉用牛経営"}} _
            , {集計３.個人_繁殖牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_肉用牛経営}, .名称 = "繁殖牛経営"}} _
            , {集計３.法人_繁殖牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_肉用牛経営}, .名称 = "繁殖牛経営"}} _
            , {集計３.個人_肥育牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_肉用牛経営}, .名称 = "肥育牛経営"}} _
            , {集計３.法人_肥育牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_肉用牛経営}, .名称 = "肥育牛経営"}} _
            , {集計３.主副業別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .名称 = "主副業別平均"}} _
            , {集計３.主業, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .名称 = "主業"}} _
            , {集計３.準主業, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .名称 = "準主業"}} _
            , {集計３.副業的, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .名称 = "副業的"}} _
            , {集計３.農業地域類型別, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "農業地域類型別"}} _
            , {集計３.都市的地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "都市的地域"}} _
            , {集計３.平地農業地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "平地農業地域"}} _
            , {集計３.中山間農業地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "中山間農業地域"}} _
            , {集計３.中間農業地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "中間農業地域"}} _
            , {集計３.山間農業地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "山間農業地域"}} _
            , {集計３.農業生産関連事業平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業生産関連事業あり}, .名称 = "農業生産関連事業平均"}} _
            , {集計３.農産加工, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業生産関連事業あり}, .名称 = "農産加工"}} _
            , {集計３.観光農園, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業生産関連事業あり}, .名称 = "観光農園"}} _
            , {集計３.貸し農園, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業生産関連事業あり}, .名称 = "貸し農園"}} _
            , {集計３.農家民宿, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業生産関連事業あり}, .名称 = "農家民宿"}} _
            , {集計３.農家レストラン, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業生産関連事業あり}, .名称 = "農家レストラン"}} _
            , {集計３.茶作単一経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体のうち組織法人経営}, .集計２ = {集計２.法人_畑作経営}, .名称 = "茶作単一経営"}} _
            , {集計３.農産加工を行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "農産加工を行っている経営"}} _
            , {集計３.観光農園を行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "観光農園を行っている経営"}} _
            , {集計３.貸し農園を行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "貸し農園を行っている経営"}} _
            , {集計３.農家民宿を行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "農家民宿を行っている経営"}} _
            , {集計３.農家レストランを行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "農家レストランを行っている経営"}}
     }

        'REV_016↓
        'Private Shared リスト_2022 As New Dictionary(Of String, 詳細) From {
        '     {集計３.項目なし, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体_第１報用_分類別平均, 集計１.個人経営体_分類別平均, 集計１.個人経営体, 集計１.個人経営体_経営形態別_分類別平均, 集計１.個人経営体_経営形態別, 集計１.法人経営体_第１報用_分類別平均, 集計１.法人経営体_分類別平均, 集計１.法人経営体, 集計１.法人経営体のうち組織法人経営_分類別平均, 集計１.法人経営体のうち組織法人経営, 集計１.法人経営体_経営形態別_分類別平均, 集計１.法人経営体_経営形態別, 集計１.農業経営体_第１報用_分類別平均, 集計１.農業経営体_分類別平均, 集計１.農業経営体, 集計１.農業経営体_経営形態別_分類別平均, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.項目なし, 集計２.個人_水田作経営_分類別平均, 集計２.個人_畑作経営_分類別平均, 集計２.個人_野菜作経営_分類別平均, 集計２.個人_果樹作経営_分類別平均, 集計２.個人_花き作経営_分類別平均, 集計２.個人_肉用牛経営_分類別平均, 集計２.平均_基本項目集計, 集計２.平均_詳細項目集計, 集計２.青色申告を行っている経営平均_基本項目集計, 集計２.青色申告を行っている経営平均_詳細項目集計, 集計２.認定農業者のいる経営平均, 集計２.法人_果樹作経営, 集計２.法人_酪農経営, 集計２.法人_養豚経営, 集計２.法人_採卵養鶏経営, 集計２.法人_ブロイラー養鶏経営, 集計２.経営平均, 集計２.平均, 集計２.水田作経営_分類別平均, 集計２.畑作経営_分類別平均, 集計２.野菜作経営_分類別平均, 集計２.花き作経営_分類別平均, 集計２.肉用牛経営_分類別平均}, .名称 = "-"}} _
        '   , {集計３.経営平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.法人経営体_経営形態別, 集計１.農業経営体, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人, 集計２.法人_畑作経営, 集計２.関連事業を行っている経営平均, 集計２.個別法人経営, 集計２.農業経営体_畑作経営}, .名称 = "経営平均"}} _
        '   , {集計３.個人_稲作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .名称 = "稲作経営"}} _
        '   , {集計３.個人_水田作_麦類作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .名称 = "水田作_麦類作経営"}} _
        '   , {集計３.個人_水田作_大豆作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .名称 = "水田作_大豆作経営"}} _
        '   , {集計３.経営平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_畑作経営, 集計２.個人_野菜作経営, 集計２.個人_果樹作経営, 集計２.個人_花き作経営, 集計２.個人_酪農経営, 集計２.個人_肉用牛経営, 集計２.個人_養豚経営, 集計２.個人_採卵養鶏経営, 集計２.個人_ブロイラー養鶏経営}, .名称 = "経営平均（基本項目集計）"}} _
        '   , {集計３.経営平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_畑作経営, 集計２.個人_野菜作経営, 集計２.個人_果樹作経営, 集計２.個人_花き作経営, 集計２.個人_酪農経営, 集計２.個人_肉用牛経営, 集計２.個人_養豚経営, 集計２.個人_採卵養鶏経営, 集計２.個人_ブロイラー養鶏経営}, .名称 = "経営平均（詳細項目集計）"}} _
        '   , {集計３.主業平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_畑作経営, 集計２.個人_野菜作経営, 集計２.個人_果樹作経営}, .名称 = "主業平均（基本項目集計）"}} _
        '   , {集計３.青色申告を行っている経営平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_畑作経営, 集計２.個人_野菜作経営, 集計２.個人_果樹作経営}, .名称 = "青色申告を行っている経営平均（詳細項目集計）"}} _
        '   , {集計３.ばれいしょ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "ばれいしょ作部門（詳細項目集計）"}} _
        '   , {集計３.茶作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "茶作部門（詳細項目集計）"}} _
        '   , {集計３.かんしょ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "かんしょ作部門（詳細項目集計）"}} _
        '   , {集計３.個人_畑作_麦類作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "畑作_麦類作経営（基本項目集計）"}} _
        '   , {集計３.さとうきび作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "さとうきび作経営（基本項目集計）"}} _
        '   , {集計３.個人_畑作_大豆作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "畑作_大豆作経営（基本項目集計）"}} _
        '   , {集計３.てんさい作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "てんさい作経営（基本項目集計）"}} _
        '   , {集計３.ばれいしょ作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "ばれいしょ作経営（基本項目集計）"}} _
        '   , {集計３.茶作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "茶作経営（基本項目集計）"}} _
        '   , {集計３.かんしょ作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "かんしょ作経営（基本項目集計）"}} _
        '   , {集計３.個人_露地野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .名称 = "露地野菜作経営"}} _
        '   , {集計３.個人_施設野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .名称 = "施設野菜作経営"}} _
        '   , {集計３.果樹作単一_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "果樹作単一（基本項目集計）"}} _
        '   , {集計３.部門別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "部門別平均"}} _
        '   , {集計３.りんご作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "りんご作部門（詳細項目集計）"}} _
        '   , {集計３.日本なし作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "日本なし作部門（詳細項目集計）"}} _
        '   , {集計３.もも作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "もも作部門（詳細項目集計）"}} _
        '   , {集計３.露地温州みかん作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "露地温州みかん作部門（詳細項目集計）"}} _
        '   , {集計３.施設温州みかん作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "施設温州みかん作部門（詳細項目集計）"}} _
        '   , {集計３.露地ぶどう作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "露地ぶどう作部門（詳細項目集計）"}} _
        '   , {集計３.施設ぶどう作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "施設ぶどう作部門（詳細項目集計）"}} _
        '   , {集計３.かき作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "かき作部門（詳細項目集計）"}} _
        '   , {集計３.うめ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "うめ作部門（詳細項目集計）"}} _
        '   , {集計３.おうとう作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "おうとう作部門（詳細項目集計）"}} _
        '   , {集計３.キウイフルーツ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "キウイフルーツ作部門（詳細項目集計）"}} _
        '   , {集計３.すもも作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "すもも作部門（詳細項目集計）"}} _
        '   , {集計３.主業経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .名称 = "主業経営（基本項目集計）"}} _
        '   , {集計３.個人_露地花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .名称 = "露地花き作経営"}} _
        '   , {集計３.個人_施設花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .名称 = "施設花き作経営"}} _
        '   , {集計３.個人_繁殖牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_肉用牛経営}, .名称 = "繁殖牛経営"}} _
        '   , {集計３.個人_肥育牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_肉用牛経営}, .名称 = "肥育牛経営"}} _
        '   , {集計３.主副業別平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .名称 = "主副業別平均"}} _
        '   , {集計３.主業, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .名称 = "主業"}} _
        '   , {集計３.準主業, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .名称 = "準主業"}} _
        '   , {集計３.副業的, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .名称 = "副業的"}} _
        '   , {集計３.農業地域類型別, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "農業地域類型別"}} _
        '   , {集計３.農業地域類型別_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "農業地域類型別（基本項目集計）"}} _
        '   , {集計３.都市的地域_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "都市的地域（基本項目集計）"}} _
        '   , {集計３.平地農業地域_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "平地農業地域（基本項目集計）"}} _
        '   , {集計３.中山間農業地域_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "中山間農業地域（基本項目集計）"}} _
        '   , {集計３.中間農業地域_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "中間農業地域（基本項目集計）"}} _
        '   , {集計３.山間農業地域_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "山間農業地域（基本項目集計）"}} _
        '   , {集計３.農業生産関連事業平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業生産関連事業あり}, .名称 = "農業生産関連事業平均（基本項目集計）"}} _
        '   , {集計３.法人_稲作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作経営"}} _
        '   , {集計３.稲作１位経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作１位経営"}} _
        '   , {集計３.稲作単一経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作単一経営"}} _
        '   , {集計３.稲作１位複合経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作１位複合経営"}} _
        '   , {集計３.法人_麦類作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人, 集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "麦類作経営"}} _
        '   , {集計３.麦類作１位経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "麦類作１位経営"}} _
        '   , {集計３.法人_大豆作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人, 集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "大豆作経営"}} _
        '   , {集計３.大豆作１位経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "大豆作１位経営"}} _
        '   , {集計３.ばれいしょ作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "ばれいしょ作経営"}} _
        '   , {集計３.てんさい作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "てんさい作経営"}} _
        '   , {集計３.茶作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "茶作経営"}} _
        '   , {集計３.さとうきび作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "さとうきび作経営"}} _
        '   , {集計３.かんしょ作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "かんしょ作経営"}} _
        '   , {集計３.野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_野菜作経営}, .名称 = "野菜作経営"}} _
        '   , {集計３.法人_露地野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_野菜作経営}, .名称 = "露地野菜作経営"}} _
        '   , {集計３.法人_施設野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_野菜作経営}, .名称 = "施設野菜作経営"}} _
        '   , {集計３.花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_花き作経営}, .名称 = "花き作経営"}} _
        '   , {集計３.法人_露地花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_花き作経営}, .名称 = "露地花き作経営"}} _
        '   , {集計３.法人_施設花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_花き作経営}, .名称 = "施設花き作経営"}} _
        '   , {集計３.肉用牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_肉用牛経営}, .名称 = "肉用牛経営"}} _
        '   , {集計３.法人_繁殖牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_肉用牛経営}, .名称 = "繁殖牛経営"}} _
        '   , {集計３.法人_肥育牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_肉用牛経営}, .名称 = "肥育牛経営"}} _
        '   , {集計３.茶作単一経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体のうち組織法人経営}, .集計２ = {集計２.法人_畑作経営}, .名称 = "茶作単一経営"}} _
        '   , {集計３.農産加工を行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "農産加工を行っている経営"}} _
        '   , {集計３.観光農園を行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "観光農園を行っている経営"}} _
        '   , {集計３.貸し農園を行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "貸し農園を行っている経営"}} _
        '   , {集計３.農家民宿を行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "農家民宿を行っている経営"}} _
        '   , {集計３.農家レストランを行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "農家レストランを行っている経営"}} _
        '   , {集計３.関連事業を行っている経営平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.個別法人経営}, .名称 = "関連事業を行っている経営平均"}} _
        '   , {集計３.都市的地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "都市的地域"}} _
        '   , {集計３.平地農業地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "平地農業地域"}} _
        '   , {集計３.中山間農業地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "中山間農業地域"}} _
        '   , {集計３.中間農業地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "中間農業地域"}} _
        '   , {集計３.山間農業地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "山間農業地域"}}
        '}
        Private Shared リスト_2022 As New Dictionary(Of String, 詳細) From {
             {集計３.項目なし, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体_第１報用_分類別平均, 集計１.個人経営体_分類別平均, 集計１.個人経営体, 集計１.個人経営体_経営形態別_分類別平均, 集計１.個人経営体_経営形態別, 集計１.法人経営体_第１報用_分類別平均, 集計１.法人経営体_分類別平均, 集計１.法人経営体, 集計１.法人経営体のうち組織法人経営_分類別平均, 集計１.法人経営体のうち組織法人経営, 集計１.法人経営体_経営形態別_分類別平均, 集計１.法人経営体_経営形態別, 集計１.農業経営体_第１報用_分類別平均, 集計１.農業経営体_分類別平均, 集計１.農業経営体, 集計１.農業経営体_経営形態別_分類別平均, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.項目なし, 集計２.個人_水田作経営_分類別平均, 集計２.個人_畑作経営_分類別平均, 集計２.個人_野菜作経営_分類別平均, 集計２.個人_果樹作経営_分類別平均, 集計２.個人_花き作経営_分類別平均, 集計２.個人_肉用牛経営_分類別平均, 集計２.平均_基本項目集計, 集計２.平均_詳細項目集計, 集計２.青色申告を行っている経営平均_基本項目集計, 集計２.青色申告を行っている経営平均_詳細項目集計, 集計２.認定農業者のいる経営平均, 集計２.法人_果樹作経営, 集計２.法人_酪農経営, 集計２.法人_養豚経営, 集計２.法人_採卵養鶏経営, 集計２.法人_ブロイラー養鶏経営, 集計２.経営平均, 集計２.平均, 集計２.水田作経営_分類別平均, 集計２.畑作経営_分類別平均, 集計２.野菜作経営_分類別平均, 集計２.花き作経営_分類別平均, 集計２.肉用牛経営_分類別平均, 集計２.認定農業者のいる経営平均_基本項目集計}, .名称 = "-"}} _
           , {集計３.経営平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体, 集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.法人経営体_経営形態別, 集計１.農業経営体, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人, 集計２.法人_畑作経営, 集計２.関連事業を行っている経営平均, 集計２.個別法人経営, 集計２.農業経営体_畑作経営}, .名称 = "経営平均"}} _
           , {集計３.個人_稲作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .名称 = "稲作経営"}} _
           , {集計３.個人_水田作_麦類作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .名称 = "水田作_麦類作経営"}} _
           , {集計３.個人_水田作_大豆作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .名称 = "水田作_大豆作経営"}} _
           , {集計３.経営平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_畑作経営, 集計２.個人_野菜作経営, 集計２.個人_果樹作経営, 集計２.個人_花き作経営, 集計２.個人_酪農経営, 集計２.個人_肉用牛経営, 集計２.個人_養豚経営, 集計２.個人_採卵養鶏経営, 集計２.個人_ブロイラー養鶏経営}, .名称 = "経営平均（基本項目集計）"}} _
           , {集計３.経営平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_畑作経営, 集計２.個人_野菜作経営, 集計２.個人_果樹作経営, 集計２.個人_花き作経営, 集計２.個人_酪農経営, 集計２.個人_肉用牛経営, 集計２.個人_養豚経営, 集計２.個人_採卵養鶏経営, 集計２.個人_ブロイラー養鶏経営}, .名称 = "経営平均（詳細項目集計）"}} _
           , {集計３.主業平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_畑作経営, 集計２.個人_野菜作経営, 集計２.個人_果樹作経営}, .名称 = "主業平均（基本項目集計）"}} _
           , {集計３.青色申告を行っている経営平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_畑作経営, 集計２.個人_野菜作経営, 集計２.個人_果樹作経営}, .名称 = "青色申告を行っている経営平均（詳細項目集計）"}} _
           , {集計３.ばれいしょ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "ばれいしょ作部門（詳細項目集計）"}} _
           , {集計３.茶作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "茶作部門（詳細項目集計）"}} _
           , {集計３.かんしょ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "かんしょ作部門（詳細項目集計）"}} _
           , {集計３.個人_畑作_麦類作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "畑作_麦類作経営（基本項目集計）"}} _
           , {集計３.さとうきび作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "さとうきび作経営（基本項目集計）"}} _
           , {集計３.個人_畑作_大豆作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "畑作_大豆作経営（基本項目集計）"}} _
           , {集計３.てんさい作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "てんさい作経営（基本項目集計）"}} _
           , {集計３.ばれいしょ作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "ばれいしょ作経営（基本項目集計）"}} _
           , {集計３.茶作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "茶作経営（基本項目集計）"}} _
           , {集計３.かんしょ作経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_畑作経営}, .名称 = "かんしょ作経営（基本項目集計）"}} _
           , {集計３.個人_露地野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .名称 = "露地野菜作経営"}} _
           , {集計３.個人_施設野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .名称 = "施設野菜作経営"}} _
           , {集計３.果樹作単一_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "果樹作単一（基本項目集計）"}} _
           , {集計３.部門別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "部門別平均"}} _
           , {集計３.りんご作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "りんご作部門（詳細項目集計）"}} _
           , {集計３.日本なし作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "日本なし作部門（詳細項目集計）"}} _
           , {集計３.もも作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "もも作部門（詳細項目集計）"}} _
           , {集計３.露地温州みかん作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "露地温州みかん作部門（詳細項目集計）"}} _
           , {集計３.施設温州みかん作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "施設温州みかん作部門（詳細項目集計）"}} _
           , {集計３.露地ぶどう作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "露地ぶどう作部門（詳細項目集計）"}} _
           , {集計３.施設ぶどう作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "施設ぶどう作部門（詳細項目集計）"}} _
           , {集計３.かき作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "かき作部門（詳細項目集計）"}} _
           , {集計３.うめ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "うめ作部門（詳細項目集計）"}} _
           , {集計３.おうとう作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "おうとう作部門（詳細項目集計）"}} _
           , {集計３.キウイフルーツ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "キウイフルーツ作部門（詳細項目集計）"}} _
           , {集計３.すもも作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_果樹作経営}, .名称 = "すもも作部門（詳細項目集計）"}} _
           , {集計３.主業経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .名称 = "主業経営（基本項目集計）"}} _
           , {集計３.個人_露地花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .名称 = "露地花き作経営"}} _
           , {集計３.個人_施設花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .名称 = "施設花き作経営"}} _
           , {集計３.個人_繁殖牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_肉用牛経営}, .名称 = "繁殖牛経営"}} _
           , {集計３.個人_肥育牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_肉用牛経営}, .名称 = "肥育牛経営"}} _
           , {集計３.主副業別平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .名称 = "主副業別平均"}} _
           , {集計３.主業, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .名称 = "主業"}} _
           , {集計３.準主業, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .名称 = "準主業"}} _
           , {集計３.副業的, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .名称 = "副業的"}} _
           , {集計３.農業地域類型別, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "農業地域類型別"}} _
           , {集計３.農業地域類型別_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "農業地域類型別（基本項目集計）"}} _
           , {集計３.都市的地域_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "都市的地域（基本項目集計）"}} _
           , {集計３.平地農業地域_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "平地農業地域（基本項目集計）"}} _
           , {集計３.中山間農業地域_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "中山間農業地域（基本項目集計）"}} _
           , {集計３.中間農業地域_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "中間農業地域（基本項目集計）"}} _
           , {集計３.山間農業地域_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "山間農業地域（基本項目集計）"}} _
           , {集計３.農業生産関連事業平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.農業生産関連事業あり}, .名称 = "農業生産関連事業平均（基本項目集計）"}} _
           , {集計３.法人_稲作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作経営"}} _
           , {集計３.稲作１位経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作１位経営"}} _
           , {集計３.稲作単一経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作単一経営"}} _
           , {集計３.稲作１位複合経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "稲作１位複合経営"}} _
           , {集計３.法人_麦類作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人, 集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "麦類作経営"}} _
           , {集計３.麦類作１位経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "麦類作１位経営"}} _
           , {集計３.法人_大豆作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人, 集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "大豆作経営"}} _
           , {集計３.大豆作１位経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人}, .名称 = "大豆作１位経営"}} _
           , {集計３.ばれいしょ作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "ばれいしょ作経営"}} _
           , {集計３.てんさい作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "てんさい作経営"}} _
           , {集計３.茶作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "茶作経営"}} _
           , {集計３.さとうきび作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "さとうきび作経営"}} _
           , {集計３.かんしょ作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.農業経営体}, .集計２ = {集計２.法人_畑作経営, 集計２.農業経営体_畑作経営}, .名称 = "かんしょ作経営"}} _
           , {集計３.野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_野菜作経営}, .名称 = "野菜作経営"}} _
           , {集計３.法人_露地野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_野菜作経営}, .名称 = "露地野菜作経営"}} _
           , {集計３.法人_施設野菜作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_野菜作経営}, .名称 = "施設野菜作経営"}} _
           , {集計３.花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_花き作経営}, .名称 = "花き作経営"}} _
           , {集計３.法人_露地花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_花き作経営}, .名称 = "露地花き作経営"}} _
           , {集計３.法人_施設花き作経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_花き作経営}, .名称 = "施設花き作経営"}} _
           , {集計３.肉用牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_肉用牛経営}, .名称 = "肉用牛経営"}} _
           , {集計３.法人_繁殖牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_肉用牛経営}, .名称 = "繁殖牛経営"}} _
           , {集計３.法人_肥育牛経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体, 集計１.法人経営体のうち組織法人経営, 集計１.農業経営体}, .集計２ = {集計２.法人_肉用牛経営}, .名称 = "肥育牛経営"}} _
           , {集計３.茶作単一経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体のうち組織法人経営}, .集計２ = {集計２.法人_畑作経営}, .名称 = "茶作単一経営"}} _
           , {集計３.農産加工を行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "農産加工を行っている経営"}} _
           , {集計３.観光農園を行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "観光農園を行っている経営"}} _
           , {集計３.貸し農園を行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "貸し農園を行っている経営"}} _
           , {集計３.農家民宿を行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "農家民宿を行っている経営"}} _
           , {集計３.農家レストランを行っている経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.関連事業を行っている経営平均}, .名称 = "農家レストランを行っている経営"}} _
           , {集計３.関連事業を行っている経営平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.個別法人経営}, .名称 = "関連事業を行っている経営平均"}} _
           , {集計３.都市的地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "都市的地域"}} _
           , {集計３.平地農業地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "平地農業地域"}} _
           , {集計３.中山間農業地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "中山間農業地域"}} _
           , {集計３.中間農業地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "中間農業地域"}} _
           , {集計３.山間農業地域, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.法人経営体_経営形態別, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.農業地域類型別}, .名称 = "山間農業地域"}}
        }
        'REV_016↑

        Public Shared Function リスト(versionKbn As String) As Dictionary(Of String, 詳細)
            If ComUtil.IsEinou Then
                '営農は2021年以前と2022年以降で切り替える
                If versionKbn = ComConst.バージョン区分.結果表等項目2022 Then
                    Return リスト_2022
                Else
                    Return リスト_2021
                End If
            Else
                Return リスト_2021
            End If
        End Function
        '---REV_012↑
    End Class

    ''' <summary>
    ''' 集計４クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 集計４
        Public Const 項目なし As String = "0"
        Public Const 専従者がいない_65歳未満 As String = "1"
        Public Const 専従者がいる_65歳未満 As String = "2"
        Public Const 主業平均 As String = "3"
        Public Const 乳用種が主 As String = "4"
        Public Const 切り花が主 As String = "5"
        Public Const 副業的平均 As String = "6"
        Public Const 大豆作１位 As String = "7"
        Public Const 施設きゅうり作部門 As String = "8"
        Public Const 施設なす作部門 As String = "9"
        Public Const 施設ばら作部門 As String = "10"
        Public Const 施設ミニトマト作部門 As String = "11"
        Public Const 施設大玉トマト作部門 As String = "12"
        Public Const 施設花き作単一 As String = "13"
        Public Const 施設野菜作単一 As String = "14"
        Public Const 施設野菜作１位複合 As String = "15"
        Public Const 準主業平均 As String = "16"
        Public Const 稲作単一 As String = "17"
        Public Const 稲作１位 As String = "18"
        Public Const 稲作１位複合 As String = "19"
        Public Const 経営平均 As String = "20"
        Public Const 肉専用種が主 As String = "21"
        Public Const 部門別平均 As String = "22"
        Public Const 鉢物が主 As String = "23"
        Public Const 露地きゅうり作部門 As String = "24"
        Public Const 露地たまねぎ作部門 As String = "25"
        Public Const 露地だいこん作部門 As String = "26"
        Public Const 露地なす作部門 As String = "27"
        Public Const 露地にんじん作部門 As String = "28"
        Public Const 露地はくさい作部門 As String = "29"
        Public Const 露地ほうれんそう作部門 As String = "30"
        Public Const 露地キャベツ作部門 As String = "31"
        Public Const 露地レタス作部門 As String = "32"
        Public Const 露地大玉トマト作部門 As String = "33"
        Public Const 露地白ねぎ作部門 As String = "34"
        Public Const 露地花き作単一 As String = "35"
        Public Const 露地野菜作単一経営 As String = "36"
        Public Const 露地野菜作１位複合 As String = "37"
        Public Const 麦類作１位 As String = "38"
        '---REV_012↓
        Public Const 経営平均_基本項目集計 As String = "39"
        Public Const 経営平均_詳細項目集計 As String = "40"
        Public Const 主業平均_基本項目集計 As String = "41"
        Public Const 主業平均_詳細項目集計 As String = "42"
        Public Const 準主業平均_基本項目集計 As String = "43"
        Public Const 準主業平均_詳細項目集計 As String = "44"
        Public Const 青色申告を行っている経営平均_詳細項目集計 As String = "45"
        Public Const 稲作１位_基本項目集計 As String = "46"
        Public Const 稲作単一_基本項目集計 As String = "47"
        Public Const 稲作１位複合_基本項目集計 As String = "48"
        Public Const 麦類作１位_基本項目集計 As String = "49"
        Public Const 大豆作１位_基本項目集計 As String = "50"
        Public Const 露地野菜作単一経営_基本項目集計 As String = "51"
        Public Const 露地野菜作１位複合_基本項目集計 As String = "52"
        Public Const 露地キャベツ作部門_詳細項目集計 As String = "53"
        Public Const 露地ほうれんそう作部門_詳細項目集計 As String = "54"
        Public Const 露地レタス作部門_詳細項目集計 As String = "55"
        Public Const 露地白ねぎ作部門_詳細項目集計 As String = "56"
        Public Const 露地だいこん作部門_詳細項目集計 As String = "57"
        Public Const 露地にんじん作部門_詳細項目集計 As String = "58"
        Public Const 露地きゅうり作部門_詳細項目集計 As String = "59"
        Public Const 露地大玉トマト作部門_詳細項目集計 As String = "60"
        Public Const 露地なす作部門_詳細項目集計 As String = "61"
        Public Const 露地たまねぎ作部門_詳細項目集計 As String = "62"
        Public Const 露地はくさい作部門_詳細項目集計 As String = "63"
        Public Const 施設野菜作単一_基本項目集計 As String = "64"
        Public Const 施設野菜作１位複合_基本項目集計 As String = "65"
        Public Const 施設大玉トマト作部門_詳細項目集計 As String = "66"
        Public Const 施設きゅうり作部門_詳細項目集計 As String = "67"
        Public Const 施設なす作部門_詳細項目集計 As String = "68"
        Public Const 施設ミニトマト作部門_詳細項目集計 As String = "69"
        Public Const 露地花き作単一_基本項目集計 As String = "70"
        Public Const 切り花が主_基本項目集計 As String = "71"
        Public Const 施設花き作単一_基本項目集計 As String = "72"
        Public Const 鉢物が主_基本項目集計 As String = "73"
        Public Const 施設ばら作部門_詳細項目集計 As String = "74"
        Public Const 肉専用種が主_基本項目集計 As String = "75"
        Public Const 乳用種が主_基本項目集計 As String = "76"
        Public Const 副業的平均_基本項目集計 As String = "77"
        '---REV_012↑

        Public Class 詳細
            Public 営農経営体区分 As String()                 '営農経営体区分
            Public 集計１ As String()                   '集計１
            Public 集計２ As String()                   '集計２
            Public 集計３ As String()                   '集計３
            Public 名称 As String                       '名称
        End Class

        '---REV_012↓
        Private Shared リスト_2021 As New Dictionary(Of String, 詳細) From {
          {集計４.項目なし, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体_第１報用_分類別平均, 集計１.個人経営体, 集計１.個人経営体_経営形態別, 集計１.個人経営体_経営形態別_分類別平均, 集計１.個人経営体_分類別平均, 集計１.法人経営体_第１報用_分類別平均, 集計１.法人経営体, 集計１.法人経営体_経営形態別, 集計１.法人経営体_経営形態別_分類別平均, 集計１.法人経営体のうち組織法人経営, 集計１.法人経営体のうち組織法人経営_分類別平均, 集計１.法人経営体_分類別平均, 集計１.農業経営体_第１報用_分類別平均, 集計１.農業経営体, 集計１.農業経営体_経営形態別, 集計１.農業経営体_経営形態別_分類別平均, 集計１.農業経営体_分類別平均}, .集計２ = {集計２.関連事業を行っている経営平均, 集計２.個人_ブロイラー養鶏経営, 集計２.主副業別_経営, 集計２.平均, 集計２.個人_採卵養鶏経営, 集計２.個人_果樹作経営, 集計２.果樹作経営_分類別平均, 集計２.個人_水田作経営, 集計２.水田作経営_分類別平均, 集計２.個人_畑作経営, 集計２.畑作経営_分類別平均, 集計２.個人_肉用牛経営, 集計２.肉用牛経営_分類別平均, 集計２.個人_花き作経営, 集計２.花き作経営_分類別平均, 集計２.認定農業者のいる経営平均, 集計２.農業地域類型別, 集計２.農業生産関連事業あり, 集計２.個人_酪農経営, 集計２.個人_野菜作経営, 集計２.野菜作経営_分類別平均, 集計２.個人_養豚経営, 集計２.法人_ブロイラー養鶏経営, 集計２.個別法人経営, 集計２.平均, 集計２.法人_採卵養鶏経営, 集計２.法人_果樹作経営, 集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人, 集計２.法人_畑作経営, 集計２.法人_肉用牛経営, 集計２.法人_花き作経営, 集計２.法人_酪農経営, 集計２.法人_野菜作経営, 集計２.法人_養豚経営, 集計２.農業経営体_畑作経営, 集計２.項目なし}, .集計３ = {集計３.項目なし, 集計３.稲作１位経営, 集計３.稲作１位複合経営, 集計３.法人_稲作経営, 集計３.稲作単一経営, 集計３.うめ作部門, 集計３.おうとう作部門, 集計３.かき作部門, 集計３.果樹作単一, 集計３.観光農園, 集計３.貸し農園, 集計３.観光農園を行っている経営, 集計３.貸し農園を行っている経営, 集計３.かんしょ作部門, 集計３.かんしょ作経営, 集計３.キウイフルーツ作部門, 集計３.経営平均, 集計３.さとうきび作経営, 集計３.施設温州みかん作部門, 集計３.すもも作部門, 集計３.てんさい作経営, 集計３.ばれいしょ作経営, 集計３.ばれいしょ作部門, 集計３.もも作部門, 集計３.りんご作部門, 集計３.中山間農業地域, 集計３.中間農業地域, 集計３.主副業別平均, 集計３.主業平均, 集計３.主業経営, 集計３.個人_畑作_大豆作経営, 集計３.法人_大豆作経営, 集計３.大豆作１位経営, 集計３.山間農業地域, 集計３.平地農業地域, 集計３.施設ぶどう作部門, 集計３.法人_施設花き作経営, 集計３.法人_施設野菜作経営, 集計３.日本なし作部門, 集計３.法人_繁殖牛経営, 集計３.肉用牛経営, 集計３.法人_肥育牛経営, 集計３.花き作経営, 集計３.茶作単一経営, 集計３.茶作経営, 集計３.茶作部門, 集計３.農家レストラン, 集計３.農家レストランを行っている経営, 集計３.農家民宿, 集計３.農家民宿を行っている経営, 集計３.農業地域類型別, 集計３.農業生産関連事業平均, 集計３.農産加工, 集計３.農産加工を行っている経営, 集計３.都市的地域, 集計３.野菜作経営, 集計３.露地ぶどう作部門, 集計３.露地温州みかん作部門, 集計３.法人_露地花き作経営, 集計３.法人_露地野菜作経営, 集計３.個人_畑作_麦類作経営, 集計３.法人_麦類作経営, 集計３.麦類作１位経営}, .名称 = "‐"}} _
        , {集計４.経営平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_野菜作経営, 集計２.個人_花き作経営, 集計２.個人_肉用牛経営}, .集計３ = {集計３.個人_稲作経営, 集計３.個人_水田作_麦類作経営, 集計３.個人_水田作_大豆作経営, 集計３.個人_露地野菜作経営, 集計３.個人_施設野菜作経営, 集計３.個人_露地花き作経営, 集計３.個人_施設花き作経営, 集計３.個人_繁殖牛経営, 集計３.個人_肥育牛経営}, .名称 = "経営平均"}} _
        , {集計４.稲作１位, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_稲作経営}, .名称 = "稲作１位"}} _
        , {集計４.稲作単一, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_稲作経営}, .名称 = "稲作単一"}} _
        , {集計４.稲作１位複合, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_稲作経営}, .名称 = "稲作１位複合"}} _
        , {集計４.麦類作１位, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_水田作_麦類作経営}, .名称 = "麦類作１位"}} _
        , {集計４.大豆作１位, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_水田作_大豆作経営}, .名称 = "大豆作１位"}} _
        , {集計４.露地野菜作単一経営, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地野菜作単一経営"}} _
        , {集計４.露地野菜作１位複合, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地野菜作１位複合"}} _
        , {集計４.露地キャベツ作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地キャベツ作部門"}} _
        , {集計４.露地ほうれんそう作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地ほうれんそう作部門"}} _
        , {集計４.露地レタス作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地レタス作部門"}} _
        , {集計４.露地白ねぎ作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地白ねぎ作部門"}} _
        , {集計４.露地だいこん作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地だいこん作部門"}} _
        , {集計４.露地にんじん作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地にんじん作部門"}} _
        , {集計４.露地きゅうり作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地きゅうり作部門"}} _
        , {集計４.露地大玉トマト作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地大玉トマト作部門"}} _
        , {集計４.露地なす作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地なす作部門"}} _
        , {集計４.露地たまねぎ作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地たまねぎ作部門"}} _
        , {集計４.露地はくさい作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地はくさい作部門"}} _
        , {集計４.施設野菜作単一, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設野菜作単一"}} _
        , {集計４.施設野菜作１位複合, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設野菜作１位複合"}} _
        , {集計４.施設大玉トマト作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設大玉トマト作部門"}} _
        , {集計４.施設きゅうり作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設きゅうり作部門"}} _
        , {集計４.施設なす作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設なす作部門"}} _
        , {集計４.施設ミニトマト作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設ミニトマト作部門"}} _
        , {集計４.露地花き作単一, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_露地花き作経営}, .名称 = "露地花き作単一"}} _
        , {集計４.切り花が主, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_露地花き作経営, 集計３.個人_施設花き作経営}, .名称 = "切り花が主"}} _
        , {集計４.施設花き作単一, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_施設花き作経営}, .名称 = "施設花き作単一"}} _
        , {集計４.鉢物が主, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_施設花き作経営}, .名称 = "鉢物が主"}} _
        , {集計４.施設ばら作部門, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_施設花き作経営}, .名称 = "施設ばら作部門"}} _
        , {集計４.肉専用種が主, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_肉用牛経営}, .集計３ = {集計３.個人_肥育牛経営}, .名称 = "肉専用種が主"}} _
        , {集計４.乳用種が主, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_肉用牛経営}, .集計３ = {集計３.個人_肥育牛経営}, .名称 = "乳用種が主"}} _
        , {集計４.主業平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体, 集計１.個人経営体_経営形態別}, .集計２ = {集計２.個人_野菜作経営, 集計２.個人_花き作経営, 集計２.主副業別_経営}, .集計３ = {集計３.個人_露地野菜作経営, 集計３.個人_施設野菜作経営, 集計３.個人_露地花き作経営, 集計３.個人_施設花き作経営, 集計３.主業}, .名称 = "主業平均"}} _
        , {集計４.準主業平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .集計３ = {集計３.準主業}, .名称 = "準主業平均"}} _
        , {集計４.専従者がいる_65歳未満, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .集計３ = {集計３.主業, 集計３.準主業}, .名称 = "65歳未満専従者がいる"}} _
        , {集計４.専従者がいない_65歳未満, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .集計３ = {集計３.主業, 集計３.準主業}, .名称 = "65歳未満専従者がいない"}} _
        , {集計４.副業的平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .集計３ = {集計３.副業的}, .名称 = "副業的平均"}}
     }

        'REV_016↓
        'Private Shared リスト_2022 As New Dictionary(Of String, 詳細) From {
        '     {集計４.項目なし, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体_第１報用_分類別平均, 集計１.個人経営体_分類別平均, 集計１.個人経営体, 集計１.個人経営体_経営形態別_分類別平均, 集計１.個人経営体_経営形態別, 集計１.法人経営体_第１報用_分類別平均, 集計１.法人経営体_分類別平均, 集計１.法人経営体, 集計１.法人経営体のうち組織法人経営_分類別平均, 集計１.法人経営体のうち組織法人経営, 集計１.法人経営体_経営形態別_分類別平均, 集計１.法人経営体_経営形態別, 集計１.農業経営体_第１報用_分類別平均, 集計１.農業経営体_分類別平均, 集計１.農業経営体, 集計１.農業経営体_経営形態別_分類別平均, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.項目なし, 集計２.個人_水田作経営_分類別平均, 集計２.個人_水田作経営, 集計２.個人_畑作経営_分類別平均, 集計２.個人_畑作経営, 集計２.個人_野菜作経営_分類別平均, 集計２.個人_野菜作経営, 集計２.個人_果樹作経営_分類別平均, 集計２.個人_果樹作経営, 集計２.個人_花き作経営_分類別平均, 集計２.個人_花き作経営, 集計２.個人_酪農経営, 集計２.個人_肉用牛経営_分類別平均, 集計２.個人_肉用牛経営, 集計２.個人_養豚経営, 集計２.個人_採卵養鶏経営, 集計２.個人_ブロイラー養鶏経営, 集計２.平均_基本項目集計, 集計２.平均_詳細項目集計, 集計２.主副業別_経営, 集計２.農業地域類型別, 集計２.青色申告を行っている経営平均_基本項目集計, 集計２.青色申告を行っている経営平均_詳細項目集計, 集計２.認定農業者のいる経営平均, 集計２.農業生産関連事業あり, 集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人, 集計２.法人_畑作経営, 集計２.法人_野菜作経営, 集計２.法人_果樹作経営, 集計２.法人_花き作経営, 集計２.法人_酪農経営, 集計２.法人_肉用牛経営, 集計２.法人_養豚経営, 集計２.法人_採卵養鶏経営, 集計２.法人_ブロイラー養鶏経営, 集計２.経営平均, 集計２.関連事業を行っている経営平均, 集計２.平均, 集計２.個別法人経営, 集計２.水田作経営_分類別平均, 集計２.畑作経営_分類別平均, 集計２.農業経営体_畑作経営, 集計２.野菜作経営_分類別平均, 集計２.花き作経営_分類別平均, 集計２.肉用牛経営_分類別平均}, .集計３ = {集計３.項目なし, 集計３.経営平均_基本項目集計, 集計３.経営平均_詳細項目集計, 集計３.主業平均_基本項目集計, 集計３.青色申告を行っている経営平均_詳細項目集計, 集計３.ばれいしょ作部門_詳細項目集計, 集計３.茶作部門_詳細項目集計, 集計３.かんしょ作部門_詳細項目集計, 集計３.個人_畑作_麦類作経営_基本項目集計, 集計３.さとうきび作経営_基本項目集計, 集計３.個人_畑作_大豆作経営_基本項目集計, 集計３.てんさい作経営_基本項目集計, 集計３.ばれいしょ作経営_基本項目集計, 集計３.茶作経営_基本項目集計, 集計３.かんしょ作経営_基本項目集計, 集計３.果樹作単一_基本項目集計, 集計３.部門別平均, 集計３.りんご作部門_詳細項目集計, 集計３.日本なし作部門_詳細項目集計, 集計３.もも作部門_詳細項目集計, 集計３.露地温州みかん作部門_詳細項目集計, 集計３.施設温州みかん作部門_詳細項目集計, 集計３.露地ぶどう作部門_詳細項目集計, 集計３.施設ぶどう作部門_詳細項目集計, 集計３.かき作部門_詳細項目集計, 集計３.うめ作部門_詳細項目集計, 集計３.おうとう作部門_詳細項目集計, 集計３.キウイフルーツ作部門_詳細項目集計, 集計３.すもも作部門_詳細項目集計, 集計３.主業経営_基本項目集計, 集計３.主副業別平均_基本項目集計, 集計３.農業地域類型別_基本項目集計, 集計３.都市的地域_基本項目集計, 集計３.平地農業地域_基本項目集計, 集計３.中山間農業地域_基本項目集計, 集計３.中間農業地域_基本項目集計, 集計３.山間農業地域_基本項目集計, 集計３.農業生産関連事業平均_基本項目集計, 集計３.経営平均, 集計３.法人_稲作経営, 集計３.稲作１位経営, 集計３.稲作単一経営, 集計３.稲作１位複合経営, 集計３.法人_麦類作経営, 集計３.麦類作１位経営, 集計３.法人_大豆作経営, 集計３.大豆作１位経営, 集計３.ばれいしょ作経営, 集計３.てんさい作経営, 集計３.茶作経営, 集計３.さとうきび作経営, 集計３.かんしょ作経営, 集計３.野菜作経営, 集計３.法人_露地野菜作経営, 集計３.法人_施設野菜作経営, 集計３.花き作経営, 集計３.法人_露地花き作経営, 集計３.法人_施設花き作経営, 集計３.肉用牛経営, 集計３.法人_繁殖牛経営, 集計３.法人_肥育牛経営, 集計３.茶作単一経営, 集計３.農業地域類型別, 集計３.農産加工を行っている経営, 集計３.観光農園を行っている経営, 集計３.貸し農園を行っている経営, 集計３.農家民宿を行っている経営, 集計３.農家レストランを行っている経営, 集計３.都市的地域, 集計３.平地農業地域, 集計３.中山間農業地域, 集計３.中間農業地域, 集計３.山間農業地域}, .名称 = "-"}} _
        '   , {集計４.経営平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_野菜作経営, 集計２.個人_花き作経営, 集計２.個人_肉用牛経営}, .集計３ = {集計３.経営平均, 集計３.個人_稲作経営, 集計３.個人_水田作_麦類作経営, 集計３.個人_水田作_大豆作経営, 集計３.個人_露地野菜作経営, 集計３.個人_施設野菜作経営, 集計３.個人_露地花き作経営, 集計３.個人_施設花き作経営, 集計３.個人_繁殖牛経営, 集計３.個人_肥育牛経営}, .名称 = "経営平均（基本項目集計）"}} _
        '   , {集計４.経営平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_野菜作経営, 集計２.個人_花き作経営, 集計２.個人_肉用牛経営}, .集計３ = {集計３.経営平均, 集計３.個人_露地野菜作経営, 集計３.個人_施設野菜作経営, 集計３.個人_露地花き作経営, 集計３.個人_施設花き作経営, 集計３.個人_繁殖牛経営, 集計３.個人_肥育牛経営}, .名称 = "経営平均（詳細項目集計）"}} _
        '   , {集計４.主業平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体, 集計１.個人経営体_経営形態別}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_野菜作経営, 集計２.個人_花き作経営, 集計２.主副業別_経営}, .集計３ = {集計３.主業平均, 集計３.個人_露地野菜作経営, 集計３.個人_施設野菜作経営, 集計３.個人_露地花き作経営, 集計３.個人_施設花き作経営, 集計３.主業}, .名称 = "主業平均（基本項目集計）"}} _
        '   , {集計４.青色申告を行っている経営平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営, 集計３.個人_施設野菜作経営}, .名称 = "青色申告を行っている経営平均（詳細項目集計）"}} _
        '   , {集計４.稲作１位_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_稲作経営}, .名称 = "稲作１位（基本項目集計）"}} _
        '   , {集計４.稲作単一_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_稲作経営}, .名称 = "稲作単一（基本項目集計）"}} _
        '   , {集計４.稲作１位複合_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_稲作経営}, .名称 = "稲作１位複合（基本項目集計）"}} _
        '   , {集計４.麦類作１位_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_水田作_麦類作経営}, .名称 = "麦類作１位（基本項目集計）"}} _
        '   , {集計４.大豆作１位_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_水田作_大豆作経営}, .名称 = "大豆作１位（基本項目集計）"}} _
        '   , {集計４.露地野菜作単一経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地野菜作単一経営（基本項目集計）"}} _
        '   , {集計４.露地野菜作１位複合_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地野菜作１位複合（基本項目集計）"}} _
        '   , {集計４.部門別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営, 集計３.個人_施設野菜作経営}, .名称 = "部門別平均"}} _
        '   , {集計４.露地キャベツ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地キャベツ作部門（詳細項目集計）"}} _
        '   , {集計４.露地ほうれんそう作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地ほうれんそう作部門（詳細項目集計）"}} _
        '   , {集計４.露地レタス作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地レタス作部門（詳細項目集計）"}} _
        '   , {集計４.露地白ねぎ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地白ねぎ作部門（詳細項目集計）"}} _
        '   , {集計４.露地だいこん作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地だいこん作部門（詳細項目集計）"}} _
        '   , {集計４.露地にんじん作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地にんじん作部門（詳細項目集計）"}} _
        '   , {集計４.露地きゅうり作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地きゅうり作部門（詳細項目集計）"}} _
        '   , {集計４.露地大玉トマト作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地大玉トマト作部門（詳細項目集計）"}} _
        '   , {集計４.露地なす作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地なす作部門（詳細項目集計）"}} _
        '   , {集計４.露地たまねぎ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地たまねぎ作部門（詳細項目集計）"}} _
        '   , {集計４.露地はくさい作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地はくさい作部門（詳細項目集計）"}} _
        '   , {集計４.施設野菜作単一_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設野菜作単一（基本項目集計）"}} _
        '   , {集計４.施設野菜作１位複合_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設野菜作１位複合（基本項目集計）"}} _
        '   , {集計４.施設大玉トマト作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設大玉トマト作部門（詳細項目集計）"}} _
        '   , {集計４.施設きゅうり作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設きゅうり作部門（詳細項目集計）"}} _
        '   , {集計４.施設なす作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設なす作部門（詳細項目集計）"}} _
        '   , {集計４.施設ミニトマト作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設ミニトマト作部門（詳細項目集計）"}} _
        '   , {集計４.露地花き作単一_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_露地花き作経営}, .名称 = "露地花き作単一（基本項目集計）"}} _
        '   , {集計４.切り花が主_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_露地花き作経営, 集計３.個人_施設花き作経営}, .名称 = "切り花が主（基本項目集計）"}} _
        '   , {集計４.施設花き作単一_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_施設花き作経営}, .名称 = "施設花き作単一（基本項目集計）"}} _
        '   , {集計４.鉢物が主_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_施設花き作経営}, .名称 = "鉢物が主（基本項目集計）"}} _
        '   , {集計４.施設ばら作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_施設花き作経営}, .名称 = "施設ばら作部門（詳細項目集計）"}} _
        '   , {集計４.肉専用種が主_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_肉用牛経営}, .集計３ = {集計３.個人_肥育牛経営}, .名称 = "肉専用種が主（基本項目集計）"}} _
        '   , {集計４.乳用種が主_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_肉用牛経営}, .集計３ = {集計３.個人_肥育牛経営}, .名称 = "乳用種が主（基本項目集計）"}} _
        '   , {集計４.主業平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .集計３ = {集計３.主業}, .名称 = "主業平均（詳細項目集計）"}} _
        '   , {集計４.準主業平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .集計３ = {集計３.準主業}, .名称 = "準主業平均（基本項目集計）"}} _
        '   , {集計４.準主業平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .集計３ = {集計３.準主業}, .名称 = "準主業平均（詳細項目集計）"}} _
        '   , {集計４.副業的平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .集計３ = {集計３.副業的}, .名称 = "副業的平均（基本項目集計）"}} _
        '   , {集計４.経営平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.個別法人経営}, .集計３ = {集計３.関連事業を行っている経営平均}, .名称 = "経営平均"}}
        '}
        Private Shared リスト_2022 As New Dictionary(Of String, 詳細) From {
             {集計４.項目なし, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体, 営農経営体区分.法人経営体, 営農経営体区分.農業経営体}, .集計１ = {集計１.個人経営体_第１報用_分類別平均, 集計１.個人経営体_分類別平均, 集計１.個人経営体, 集計１.個人経営体_経営形態別_分類別平均, 集計１.個人経営体_経営形態別, 集計１.法人経営体_第１報用_分類別平均, 集計１.法人経営体_分類別平均, 集計１.法人経営体, 集計１.法人経営体のうち組織法人経営_分類別平均, 集計１.法人経営体のうち組織法人経営, 集計１.法人経営体_経営形態別_分類別平均, 集計１.法人経営体_経営形態別, 集計１.農業経営体_第１報用_分類別平均, 集計１.農業経営体_分類別平均, 集計１.農業経営体, 集計１.農業経営体_経営形態別_分類別平均, 集計１.農業経営体_経営形態別}, .集計２ = {集計２.項目なし, 集計２.個人_水田作経営_分類別平均, 集計２.個人_水田作経営, 集計２.個人_畑作経営_分類別平均, 集計２.個人_畑作経営, 集計２.個人_野菜作経営_分類別平均, 集計２.個人_野菜作経営, 集計２.個人_果樹作経営_分類別平均, 集計２.個人_果樹作経営, 集計２.個人_花き作経営_分類別平均, 集計２.個人_花き作経営, 集計２.個人_酪農経営, 集計２.個人_肉用牛経営_分類別平均, 集計２.個人_肉用牛経営, 集計２.個人_養豚経営, 集計２.個人_採卵養鶏経営, 集計２.個人_ブロイラー養鶏経営, 集計２.平均_基本項目集計, 集計２.平均_詳細項目集計, 集計２.主副業別_経営, 集計２.農業地域類型別, 集計２.青色申告を行っている経営平均_基本項目集計, 集計２.青色申告を行っている経営平均_詳細項目集計, 集計２.認定農業者のいる経営平均, 集計２.農業生産関連事業あり, 集計２.法人_水田作経営, 集計２.水田作経営のうち集落営農組織法人, 集計２.法人_畑作経営, 集計２.法人_野菜作経営, 集計２.法人_果樹作経営, 集計２.法人_花き作経営, 集計２.法人_酪農経営, 集計２.法人_肉用牛経営, 集計２.法人_養豚経営, 集計２.法人_採卵養鶏経営, 集計２.法人_ブロイラー養鶏経営, 集計２.経営平均, 集計２.関連事業を行っている経営平均, 集計２.平均, 集計２.個別法人経営, 集計２.水田作経営_分類別平均, 集計２.畑作経営_分類別平均, 集計２.農業経営体_畑作経営, 集計２.野菜作経営_分類別平均, 集計２.花き作経営_分類別平均, 集計２.肉用牛経営_分類別平均, 集計２.認定農業者のいる経営平均_基本項目集計}, .集計３ = {集計３.項目なし, 集計３.経営平均_基本項目集計, 集計３.経営平均_詳細項目集計, 集計３.主業平均_基本項目集計, 集計３.青色申告を行っている経営平均_詳細項目集計, 集計３.ばれいしょ作部門_詳細項目集計, 集計３.茶作部門_詳細項目集計, 集計３.かんしょ作部門_詳細項目集計, 集計３.個人_畑作_麦類作経営_基本項目集計, 集計３.さとうきび作経営_基本項目集計, 集計３.個人_畑作_大豆作経営_基本項目集計, 集計３.てんさい作経営_基本項目集計, 集計３.ばれいしょ作経営_基本項目集計, 集計３.茶作経営_基本項目集計, 集計３.かんしょ作経営_基本項目集計, 集計３.果樹作単一_基本項目集計, 集計３.部門別平均, 集計３.りんご作部門_詳細項目集計, 集計３.日本なし作部門_詳細項目集計, 集計３.もも作部門_詳細項目集計, 集計３.露地温州みかん作部門_詳細項目集計, 集計３.施設温州みかん作部門_詳細項目集計, 集計３.露地ぶどう作部門_詳細項目集計, 集計３.施設ぶどう作部門_詳細項目集計, 集計３.かき作部門_詳細項目集計, 集計３.うめ作部門_詳細項目集計, 集計３.おうとう作部門_詳細項目集計, 集計３.キウイフルーツ作部門_詳細項目集計, 集計３.すもも作部門_詳細項目集計, 集計３.主業経営_基本項目集計, 集計３.主副業別平均_基本項目集計, 集計３.農業地域類型別_基本項目集計, 集計３.都市的地域_基本項目集計, 集計３.平地農業地域_基本項目集計, 集計３.中山間農業地域_基本項目集計, 集計３.中間農業地域_基本項目集計, 集計３.山間農業地域_基本項目集計, 集計３.農業生産関連事業平均_基本項目集計, 集計３.経営平均, 集計３.法人_稲作経営, 集計３.稲作１位経営, 集計３.稲作単一経営, 集計３.稲作１位複合経営, 集計３.法人_麦類作経営, 集計３.麦類作１位経営, 集計３.法人_大豆作経営, 集計３.大豆作１位経営, 集計３.ばれいしょ作経営, 集計３.てんさい作経営, 集計３.茶作経営, 集計３.さとうきび作経営, 集計３.かんしょ作経営, 集計３.野菜作経営, 集計３.法人_露地野菜作経営, 集計３.法人_施設野菜作経営, 集計３.花き作経営, 集計３.法人_露地花き作経営, 集計３.法人_施設花き作経営, 集計３.肉用牛経営, 集計３.法人_繁殖牛経営, 集計３.法人_肥育牛経営, 集計３.茶作単一経営, 集計３.農業地域類型別, 集計３.農産加工を行っている経営, 集計３.観光農園を行っている経営, 集計３.貸し農園を行っている経営, 集計３.農家民宿を行っている経営, 集計３.農家レストランを行っている経営, 集計３.都市的地域, 集計３.平地農業地域, 集計３.中山間農業地域, 集計３.中間農業地域, 集計３.山間農業地域}, .名称 = "-"}} _
           , {集計４.経営平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_野菜作経営, 集計２.個人_花き作経営, 集計２.個人_肉用牛経営}, .集計３ = {集計３.経営平均, 集計３.個人_稲作経営, 集計３.個人_水田作_麦類作経営, 集計３.個人_水田作_大豆作経営, 集計３.個人_露地野菜作経営, 集計３.個人_施設野菜作経営, 集計３.個人_露地花き作経営, 集計３.個人_施設花き作経営, 集計３.個人_繁殖牛経営, 集計３.個人_肥育牛経営}, .名称 = "経営平均（基本項目集計）"}} _
           , {集計４.経営平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_野菜作経営, 集計２.個人_花き作経営, 集計２.個人_肉用牛経営}, .集計３ = {集計３.経営平均, 集計３.個人_露地野菜作経営, 集計３.個人_施設野菜作経営, 集計３.個人_露地花き作経営, 集計３.個人_施設花き作経営, 集計３.個人_繁殖牛経営, 集計３.個人_肥育牛経営}, .名称 = "経営平均（詳細項目集計）"}} _
           , {集計４.主業平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体, 集計１.個人経営体_経営形態別}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_野菜作経営, 集計２.個人_花き作経営, 集計２.主副業別_経営}, .集計３ = {集計３.主業平均, 集計３.個人_露地野菜作経営, 集計３.個人_施設野菜作経営, 集計３.個人_露地花き作経営, 集計３.個人_施設花き作経営, 集計３.主業}, .名称 = "主業平均（基本項目集計）"}} _
           , {集計４.青色申告を行っている経営平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営, 集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営, 集計３.個人_施設野菜作経営}, .名称 = "青色申告を行っている経営平均（詳細項目集計）"}} _
           , {集計４.稲作１位_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_稲作経営}, .名称 = "稲作１位（基本項目集計）"}} _
           , {集計４.稲作単一_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_稲作経営}, .名称 = "稲作単一（基本項目集計）"}} _
           , {集計４.稲作１位複合_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_稲作経営}, .名称 = "稲作１位複合（基本項目集計）"}} _
           , {集計４.麦類作１位_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_水田作_麦類作経営}, .名称 = "麦類作１位（基本項目集計）"}} _
           , {集計４.大豆作１位_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_水田作経営}, .集計３ = {集計３.個人_水田作_大豆作経営}, .名称 = "大豆作１位（基本項目集計）"}} _
           , {集計４.露地野菜作単一経営_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地野菜作単一経営（基本項目集計）"}} _
           , {集計４.露地野菜作１位複合_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地野菜作１位複合（基本項目集計）"}} _
           , {集計４.部門別平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営, 集計３.個人_施設野菜作経営}, .名称 = "部門別平均"}} _
           , {集計４.露地キャベツ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地キャベツ作部門（詳細項目集計）"}} _
           , {集計４.露地ほうれんそう作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地ほうれんそう作部門（詳細項目集計）"}} _
           , {集計４.露地レタス作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地レタス作部門（詳細項目集計）"}} _
           , {集計４.露地白ねぎ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地白ねぎ作部門（詳細項目集計）"}} _
           , {集計４.露地だいこん作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地だいこん作部門（詳細項目集計）"}} _
           , {集計４.露地にんじん作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地にんじん作部門（詳細項目集計）"}} _
           , {集計４.露地きゅうり作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地きゅうり作部門（詳細項目集計）"}} _
           , {集計４.露地大玉トマト作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地大玉トマト作部門（詳細項目集計）"}} _
           , {集計４.露地なす作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地なす作部門（詳細項目集計）"}} _
           , {集計４.露地たまねぎ作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地たまねぎ作部門（詳細項目集計）"}} _
           , {集計４.露地はくさい作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_露地野菜作経営}, .名称 = "露地はくさい作部門（詳細項目集計）"}} _
           , {集計４.施設野菜作単一_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設野菜作単一（基本項目集計）"}} _
           , {集計４.施設野菜作１位複合_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設野菜作１位複合（基本項目集計）"}} _
           , {集計４.施設大玉トマト作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設大玉トマト作部門（詳細項目集計）"}} _
           , {集計４.施設きゅうり作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設きゅうり作部門（詳細項目集計）"}} _
           , {集計４.施設なす作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設なす作部門（詳細項目集計）"}} _
           , {集計４.施設ミニトマト作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_野菜作経営}, .集計３ = {集計３.個人_施設野菜作経営}, .名称 = "施設ミニトマト作部門（詳細項目集計）"}} _
           , {集計４.露地花き作単一_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_露地花き作経営}, .名称 = "露地花き作単一（基本項目集計）"}} _
           , {集計４.切り花が主_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_露地花き作経営, 集計３.個人_施設花き作経営}, .名称 = "切り花が主（基本項目集計）"}} _
           , {集計４.施設花き作単一_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_施設花き作経営}, .名称 = "施設花き作単一（基本項目集計）"}} _
           , {集計４.鉢物が主_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_施設花き作経営}, .名称 = "鉢物が主（基本項目集計）"}} _
           , {集計４.施設ばら作部門_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_花き作経営}, .集計３ = {集計３.個人_施設花き作経営}, .名称 = "施設ばら作部門（詳細項目集計）"}} _
           , {集計４.肉専用種が主_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_肉用牛経営}, .集計３ = {集計３.個人_肥育牛経営}, .名称 = "肉専用種が主（基本項目集計）"}} _
           , {集計４.乳用種が主_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体}, .集計２ = {集計２.個人_肉用牛経営}, .集計３ = {集計３.個人_肥育牛経営}, .名称 = "乳用種が主（基本項目集計）"}} _
           , {集計４.主業平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .集計３ = {集計３.主業}, .名称 = "主業平均（詳細項目集計）"}} _
           , {集計４.準主業平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .集計３ = {集計３.準主業}, .名称 = "準主業平均（基本項目集計）"}} _
           , {集計４.準主業平均_詳細項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .集計３ = {集計３.準主業}, .名称 = "準主業平均（詳細項目集計）"}} _
           , {集計４.副業的平均_基本項目集計, New 詳細 With {.営農経営体区分 = {営農経営体区分.個人経営体}, .集計１ = {集計１.個人経営体_経営形態別}, .集計２ = {集計２.主副業別_経営}, .集計３ = {集計３.副業的}, .名称 = "副業的平均（基本項目集計）"}} _
           , {集計４.経営平均, New 詳細 With {.営農経営体区分 = {営農経営体区分.法人経営体}, .集計１ = {集計１.法人経営体_経営形態別}, .集計２ = {集計２.個別法人経営}, .集計３ = {集計３.関連事業を行っている経営平均}, .名称 = "経営平均"}}
        }
        'REV_016↑

        Public Shared Function リスト(versionKbn As String) As Dictionary(Of String, 詳細)
            If ComUtil.IsEinou Then
                '営農は2021年以前と2022年以降で切り替える
                If versionKbn = ComConst.バージョン区分.結果表等項目2022 Then
                    Return リスト_2022
                Else
                    Return リスト_2021
                End If
            Else
                Return リスト_2021
            End If
        End Function
        '---REV_012↑
    End Class

    '---REV_012↓
    ''' <summary>
    ''' 基本詳細項目集計クラス
    ''' </summary>
    Public Class 基本詳細項目集計
        Public Const 基本詳細項目集計 As String = "1"
        Public Const 詳細項目集計 As String = "2"

        Public Shared Function getValue(versionkbn As String, syukeiCode As String) As String
            If IsShosaiOnly(versionkbn, syukeiCode) Then
                '詳細項目集計
                Return 詳細項目集計
            Else
                '基本詳細項目集計
                Return 基本詳細項目集計
            End If
        End Function

        Public Shared Function IsShosaiOnly(versionkbn As String, syukeiCode As String) As Boolean
            If Not versionkbn = ComConst.バージョン区分.結果表等項目2022 Then
                Return False
            End If

            Return リスト.Contains(syukeiCode)
        End Function

        Private Shared リスト As New List(Of String)({
         "a-119",
        "a-120",
        "a-121",
        "a-122",
        "a-018",
        "a-019",
        "a-020",
        "a-123",
        "a-124",
        "a-125",
        "a-126",
        "a-033",
        "a-034",
        "a-035",
        "a-036",
        "a-037",
        "a-038",
        "a-039",
        "a-040",
        "a-041",
        "a-042",
        "a-043",
        "a-128",
        "a-129",
        "a-049",
        "a-050",
        "a-051",
        "a-052",
        "a-131",
        "a-132",
        "a-058",
        "a-059",
        "a-060",
        "a-061",
        "a-062",
        "a-063",
        "a-064",
        "a-065",
        "a-066",
        "a-067",
        "a-068",
        "a-069",
        "a-134",
        "a-135",
        "a-136",
        "a-082",
        "a-137",
        "a-138",
        "a-139",
        "a-140",
        "a-141",
        "a-142",
        "a-143",
        "a-144",
        "a-145",
        "a-146",
        "a-148"
        })

    End Class
    '---REV_012↑

    ''' <summary>
    ''' 地域区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 地域区分
        Public Const 全国 As String = "1"
        Public Const 大地域 As String = "2"
        Public Const 小地域 As String = "3"
        Public Const 農政局 As String = "4"
        Public Const 都府県 As String = "5"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {地域区分.全国, "全国"} _
            , {地域区分.大地域, "農業地域（大地域）"} _
            , {地域区分.小地域, "農業地域（小地域）"} _
            , {地域区分.農政局, "農政局"} _
            , {地域区分.都府県, "都府県"}
        }
    End Class

    ''' <summary>
    ''' 地域クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 地域
        Public Const 全国 As String = "0"
        Public Const 都府県平均 As String = "60"
        Public Const 北海道 As String = "1"
        Public Const 東北 As String = "61"
        Public Const 北陸 As String = "62"
        Public Const 関東_東山 As String = "63"
        Public Const 東海 As String = "67"
        Public Const 近畿 As String = "68"
        Public Const 中国 As String = "71"
        Public Const 四国 As String = "74"
        Public Const 九州 As String = "75"
        Public Const 沖縄 As String = "47"
        Public Const 北関東 As String = "64"
        Public Const 南関東 As String = "65"
        Public Const 東山 As String = "66"
        Public Const 北東北 As String = "69"
        Public Const 南東北 As String = "70"
        Public Const 山陰 As String = "72"
        Public Const 山陽 As String = "73"
        Public Const 北九州 As String = "76"
        Public Const 南九州 As String = "77"
        Public Const 関東農政局 As String = "93"
        Public Const 東海農政局 As String = "94"
        Public Const 中国四国局 As String = "95"
        Public Const 青森 As String = "2"
        Public Const 岩手 As String = "3"
        Public Const 宮城 As String = "4"
        Public Const 秋田 As String = "5"
        Public Const 山形 As String = "6"
        Public Const 福島 As String = "7"
        Public Const 茨城 As String = "8"
        Public Const 栃木 As String = "9"
        Public Const 群馬 As String = "10"
        Public Const 埼玉 As String = "11"
        Public Const 千葉 As String = "12"
        Public Const 東京 As String = "13"
        Public Const 神奈川 As String = "14"
        Public Const 山梨 As String = "19"
        Public Const 長野 As String = "20"
        Public Const 静岡 As String = "22"
        Public Const 新潟 As String = "15"
        Public Const 富山 As String = "16"
        Public Const 石川 As String = "17"
        Public Const 福井 As String = "18"
        Public Const 岐阜 As String = "21"
        Public Const 愛知 As String = "23"
        Public Const 三重 As String = "24"
        Public Const 滋賀 As String = "25"
        Public Const 京都 As String = "26"
        Public Const 大阪 As String = "27"
        Public Const 兵庫 As String = "28"
        Public Const 奈良 As String = "29"
        Public Const 和歌山 As String = "30"
        Public Const 鳥取 As String = "31"
        Public Const 島根 As String = "32"
        Public Const 岡山 As String = "33"
        Public Const 広島 As String = "34"
        Public Const 山口 As String = "35"
        Public Const 徳島 As String = "36"
        Public Const 香川 As String = "37"
        Public Const 愛媛 As String = "38"
        Public Const 高知 As String = "39"
        Public Const 福岡 As String = "40"
        Public Const 佐賀 As String = "41"
        Public Const 長崎 As String = "42"
        Public Const 熊本 As String = "43"
        Public Const 大分 As String = "44"
        Public Const 宮崎 As String = "45"
        Public Const 鹿児島 As String = "46"

        Public Class 詳細
            Public 地域区分 As String                  '地域区分
            Public 局 As String                        '局
            Public 名称 As String                      '名称
            Public 都道府県 As String()                '都道府県
        End Class

        Public Shared リスト As New Dictionary(Of String, 詳細) From {
              {地域.全国, New 詳細 With {.地域区分 = 地域区分.全国, .局 = "0", .名称 = "全国", .都道府県 = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47"}}} _
            , {地域.都府県平均, New 詳細 With {.地域区分 = 地域区分.全国, .局 = "0", .名称 = "都府県平均", .都道府県 = {"2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47"}}} _
            , {地域.北海道, New 詳細 With {.地域区分 = 地域区分.大地域, .局 = "1", .名称 = "北海道", .都道府県 = {"1"}}} _
            , {地域.東北, New 詳細 With {.地域区分 = 地域区分.大地域, .局 = "2", .名称 = "東北", .都道府県 = {"2", "3", "4", "5", "6", "7"}}} _
            , {地域.北陸, New 詳細 With {.地域区分 = 地域区分.大地域, .局 = "3", .名称 = "北陸", .都道府県 = {"15", "16", "17", "18"}}} _
            , {地域.関東_東山, New 詳細 With {.地域区分 = 地域区分.大地域, .局 = "4", .名称 = "関東・東山", .都道府県 = {"8", "9", "10", "11", "12", "13", "14", "19", "20"}}} _
            , {地域.東海, New 詳細 With {.地域区分 = 地域区分.大地域, .局 = "0", .名称 = "東海", .都道府県 = {"21", "22", "23", "24"}}} _
            , {地域.近畿, New 詳細 With {.地域区分 = 地域区分.大地域, .局 = "6", .名称 = "近畿", .都道府県 = {"25", "26", "27", "28", "29", "30"}}} _
            , {地域.中国, New 詳細 With {.地域区分 = 地域区分.大地域, .局 = "7", .名称 = "中国", .都道府県 = {"31", "32", "33", "34", "35"}}} _
            , {地域.四国, New 詳細 With {.地域区分 = 地域区分.大地域, .局 = "7", .名称 = "四国", .都道府県 = {"36", "37", "38", "39"}}} _
            , {地域.九州, New 詳細 With {.地域区分 = 地域区分.大地域, .局 = "8", .名称 = "九州", .都道府県 = {"40", "41", "42", "43", "44", "45", "46"}}} _
            , {地域.沖縄, New 詳細 With {.地域区分 = 地域区分.大地域, .局 = "9", .名称 = "沖縄", .都道府県 = {"47"}}} _
            , {地域.北関東, New 詳細 With {.地域区分 = 地域区分.小地域, .局 = "4", .名称 = "北関東", .都道府県 = {"8", "9", "10"}}} _
            , {地域.南関東, New 詳細 With {.地域区分 = 地域区分.小地域, .局 = "4", .名称 = "南関東", .都道府県 = {"11", "12", "13", "14"}}} _
            , {地域.東山, New 詳細 With {.地域区分 = 地域区分.小地域, .局 = "4", .名称 = "東山", .都道府県 = {"19", "20"}}} _
            , {地域.北東北, New 詳細 With {.地域区分 = 地域区分.小地域, .局 = "2", .名称 = "北東北", .都道府県 = {"2", "3", "5"}}} _
            , {地域.南東北, New 詳細 With {.地域区分 = 地域区分.小地域, .局 = "2", .名称 = "南東北", .都道府県 = {"4", "6", "7"}}} _
            , {地域.山陰, New 詳細 With {.地域区分 = 地域区分.小地域, .局 = "7", .名称 = "山陰", .都道府県 = {"31", "32"}}} _
            , {地域.山陽, New 詳細 With {.地域区分 = 地域区分.小地域, .局 = "7", .名称 = "山陽", .都道府県 = {"33", "34", "35"}}} _
            , {地域.北九州, New 詳細 With {.地域区分 = 地域区分.小地域, .局 = "8", .名称 = "北九州", .都道府県 = {"40", "41", "42", "43", "44"}}} _
            , {地域.南九州, New 詳細 With {.地域区分 = 地域区分.小地域, .局 = "8", .名称 = "南九州", .都道府県 = {"45", "46"}}} _
            , {地域.関東農政局, New 詳細 With {.地域区分 = 地域区分.農政局, .局 = "4", .名称 = "関東農政局", .都道府県 = {"8", "9", "10", "11", "12", "13", "14", "19", "20", "22"}}} _
            , {地域.東海農政局, New 詳細 With {.地域区分 = 地域区分.農政局, .局 = "5", .名称 = "東海農政局", .都道府県 = {"21", "23", "24"}}} _
            , {地域.中国四国局, New 詳細 With {.地域区分 = 地域区分.農政局, .局 = "7", .名称 = "中国四国局", .都道府県 = {"31", "32", "33", "34", "35", "36", "37", "38", "39"}}} _
            , {地域.青森, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "2", .名称 = "青森", .都道府県 = {"2"}}} _
            , {地域.岩手, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "2", .名称 = "岩手", .都道府県 = {"3"}}} _
            , {地域.宮城, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "2", .名称 = "宮城", .都道府県 = {"4"}}} _
            , {地域.秋田, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "2", .名称 = "秋田", .都道府県 = {"5"}}} _
            , {地域.山形, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "2", .名称 = "山形", .都道府県 = {"6"}}} _
            , {地域.福島, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "2", .名称 = "福島", .都道府県 = {"7"}}} _
            , {地域.茨城, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "4", .名称 = "茨城", .都道府県 = {"8"}}} _
            , {地域.栃木, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "4", .名称 = "栃木", .都道府県 = {"9"}}} _
            , {地域.群馬, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "4", .名称 = "群馬", .都道府県 = {"10"}}} _
            , {地域.埼玉, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "4", .名称 = "埼玉", .都道府県 = {"11"}}} _
            , {地域.千葉, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "4", .名称 = "千葉", .都道府県 = {"12"}}} _
            , {地域.東京, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "4", .名称 = "東京", .都道府県 = {"13"}}} _
            , {地域.神奈川, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "4", .名称 = "神奈川", .都道府県 = {"14"}}} _
            , {地域.山梨, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "4", .名称 = "山梨", .都道府県 = {"19"}}} _
            , {地域.長野, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "4", .名称 = "長野", .都道府県 = {"20"}}} _
            , {地域.静岡, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "4", .名称 = "静岡", .都道府県 = {"22"}}} _
            , {地域.新潟, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "3", .名称 = "新潟", .都道府県 = {"15"}}} _
            , {地域.富山, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "3", .名称 = "富山", .都道府県 = {"16"}}} _
            , {地域.石川, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "3", .名称 = "石川", .都道府県 = {"17"}}} _
            , {地域.福井, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "3", .名称 = "福井", .都道府県 = {"18"}}} _
            , {地域.岐阜, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "5", .名称 = "岐阜", .都道府県 = {"21"}}} _
            , {地域.愛知, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "5", .名称 = "愛知", .都道府県 = {"23"}}} _
            , {地域.三重, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "5", .名称 = "三重", .都道府県 = {"24"}}} _
            , {地域.滋賀, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "6", .名称 = "滋賀", .都道府県 = {"25"}}} _
            , {地域.京都, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "6", .名称 = "京都", .都道府県 = {"26"}}} _
            , {地域.大阪, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "6", .名称 = "大阪", .都道府県 = {"27"}}} _
            , {地域.兵庫, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "6", .名称 = "兵庫", .都道府県 = {"28"}}} _
            , {地域.奈良, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "6", .名称 = "奈良", .都道府県 = {"29"}}} _
            , {地域.和歌山, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "6", .名称 = "和歌山", .都道府県 = {"30"}}} _
            , {地域.鳥取, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "7", .名称 = "鳥取", .都道府県 = {"31"}}} _
            , {地域.島根, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "7", .名称 = "島根", .都道府県 = {"32"}}} _
            , {地域.岡山, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "7", .名称 = "岡山", .都道府県 = {"33"}}} _
            , {地域.広島, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "7", .名称 = "広島", .都道府県 = {"34"}}} _
            , {地域.山口, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "7", .名称 = "山口", .都道府県 = {"35"}}} _
            , {地域.徳島, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "7", .名称 = "徳島", .都道府県 = {"36"}}} _
            , {地域.香川, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "7", .名称 = "香川", .都道府県 = {"37"}}} _
            , {地域.愛媛, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "7", .名称 = "愛媛", .都道府県 = {"38"}}} _
            , {地域.高知, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "7", .名称 = "高知", .都道府県 = {"39"}}} _
            , {地域.福岡, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "8", .名称 = "福岡", .都道府県 = {"40"}}} _
            , {地域.佐賀, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "8", .名称 = "佐賀", .都道府県 = {"41"}}} _
            , {地域.長崎, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "8", .名称 = "長崎", .都道府県 = {"42"}}} _
            , {地域.熊本, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "8", .名称 = "熊本", .都道府県 = {"43"}}} _
            , {地域.大分, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "8", .名称 = "大分", .都道府県 = {"44"}}} _
            , {地域.宮崎, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "8", .名称 = "宮崎", .都道府県 = {"45"}}} _
            , {地域.鹿児島, New 詳細 With {.地域区分 = 地域区分.都府県, .局 = "8", .名称 = "鹿児島", .都道府県 = {"46"}}}
        }
    End Class

    ''' <summary>
    ''' 部門クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 部門
        Public Const かんしょ作 As String = "3"
        Public Const ばれいしょ作 As String = "4"
        Public Const 茶作 As String = "55"
        Public Const 露地きゅうり作 As String = "9"
        Public Const 露地大玉トマト作 As String = "10"
        Public Const 露地なす作 As String = "11"
        Public Const 露地キャベツ作 As String = "14"
        Public Const 露地ほうれんそう作 As String = "15"
        Public Const 露地たまねぎ作 As String = "16"
        Public Const 露地レタス作 As String = "17"
        Public Const 露地はくさい作 As String = "18"
        Public Const 露地白ねぎ作 As String = "19"
        Public Const 露地だいこん作 As String = "20"
        Public Const 露地にんじん作 As String = "21"
        Public Const 施設きゅうり作 As String = "29"
        Public Const 施設大玉トマト作 As String = "30"
        Public Const 施設ミニトマト作 As String = "31"
        Public Const 施設なす作 As String = "32"
        Public Const りんご作 As String = "37"
        Public Const 露地温州みかん作 As String = "38"
        Public Const 施設温州みかん作 As String = "39"
        Public Const 露地ぶどう作 As String = "40"
        Public Const 施設ぶどう作 As String = "41"
        Public Const 日本なし作 As String = "42"
        Public Const もも作 As String = "43"
        Public Const かき作 As String = "44"
        Public Const うめ作 As String = "45"
        Public Const おうとう作 As String = "46"
        Public Const キウイフルーツ作 As String = "48"
        Public Const すもも作 As String = "49"
        Public Const 施設ばら作 As String = "54"

        Public Class 詳細
            Public 営農類型 As String                  '営農類型
            Public 名称 As String                      '名称
        End Class

        Public Shared リスト As New Dictionary(Of String, 詳細) From {
              {部門.かんしょ作, New 詳細 With {.営農類型 = 営農類型区分.畑作, .名称 = "かんしょ作"}} _
            , {部門.ばれいしょ作, New 詳細 With {.営農類型 = 営農類型区分.畑作, .名称 = "ばれいしょ作"}} _
            , {部門.茶作, New 詳細 With {.営農類型 = 営農類型区分.畑作, .名称 = "茶作"}} _
            , {部門.露地きゅうり作, New 詳細 With {.営農類型 = 営農類型区分.露地野菜作, .名称 = "露地きゅうり作"}} _
            , {部門.露地大玉トマト作, New 詳細 With {.営農類型 = 営農類型区分.露地野菜作, .名称 = "露地大玉トマト作 "}} _
            , {部門.露地なす作, New 詳細 With {.営農類型 = 営農類型区分.露地野菜作, .名称 = "露地なす作"}} _
            , {部門.露地キャベツ作, New 詳細 With {.営農類型 = 営農類型区分.露地野菜作, .名称 = "露地キャベツ作"}} _
            , {部門.露地ほうれんそう作, New 詳細 With {.営農類型 = 営農類型区分.露地野菜作, .名称 = "露地ほうれんそう作"}} _
            , {部門.露地たまねぎ作, New 詳細 With {.営農類型 = 営農類型区分.露地野菜作, .名称 = "露地たまねぎ作"}} _
            , {部門.露地レタス作, New 詳細 With {.営農類型 = 営農類型区分.露地野菜作, .名称 = "露地レタス作"}} _
            , {部門.露地はくさい作, New 詳細 With {.営農類型 = 営農類型区分.露地野菜作, .名称 = "露地はくさい作"}} _
            , {部門.露地白ねぎ作, New 詳細 With {.営農類型 = 営農類型区分.露地野菜作, .名称 = "露地白ねぎ作"}} _
            , {部門.露地だいこん作, New 詳細 With {.営農類型 = 営農類型区分.露地野菜作, .名称 = "露地だいこん作"}} _
            , {部門.露地にんじん作, New 詳細 With {.営農類型 = 営農類型区分.露地野菜作, .名称 = "露地にんじん作"}} _
            , {部門.施設きゅうり作, New 詳細 With {.営農類型 = 営農類型区分.施設野菜作, .名称 = "施設きゅうり作"}} _
            , {部門.施設大玉トマト作, New 詳細 With {.営農類型 = 営農類型区分.施設野菜作, .名称 = "施設大玉トマト作"}} _
            , {部門.施設ミニトマト作, New 詳細 With {.営農類型 = 営農類型区分.施設野菜作, .名称 = "施設ミニトマト作"}} _
            , {部門.施設なす作, New 詳細 With {.営農類型 = 営農類型区分.施設野菜作, .名称 = "施設なす作"}} _
            , {部門.りんご作, New 詳細 With {.営農類型 = 営農類型区分.果樹作, .名称 = "りんご作"}} _
            , {部門.露地温州みかん作, New 詳細 With {.営農類型 = 営農類型区分.果樹作, .名称 = "露地温州みかん作"}} _
            , {部門.施設温州みかん作, New 詳細 With {.営農類型 = 営農類型区分.果樹作, .名称 = "施設温州みかん作"}} _
            , {部門.露地ぶどう作, New 詳細 With {.営農類型 = 営農類型区分.果樹作, .名称 = "露地ぶどう作"}} _
            , {部門.施設ぶどう作, New 詳細 With {.営農類型 = 営農類型区分.果樹作, .名称 = "施設ぶどう作"}} _
            , {部門.日本なし作, New 詳細 With {.営農類型 = 営農類型区分.果樹作, .名称 = "日本なし作"}} _
            , {部門.もも作, New 詳細 With {.営農類型 = 営農類型区分.果樹作, .名称 = "もも作"}} _
            , {部門.かき作, New 詳細 With {.営農類型 = 営農類型区分.果樹作, .名称 = "かき作"}} _
            , {部門.うめ作, New 詳細 With {.営農類型 = 営農類型区分.果樹作, .名称 = "うめ作"}} _
            , {部門.おうとう作, New 詳細 With {.営農類型 = 営農類型区分.果樹作, .名称 = "おうとう作"}} _
            , {部門.キウイフルーツ作, New 詳細 With {.営農類型 = 営農類型区分.果樹作, .名称 = "キウイフルーツ作"}} _
            , {部門.すもも作, New 詳細 With {.営農類型 = 営農類型区分.果樹作, .名称 = "すもも作"}} _
            , {部門.施設ばら作, New 詳細 With {.営農類型 = 営農類型区分.施設花き作, .名称 = "施設ばら作"}}
        }
    End Class

    ''' <summary>
    ''' 生産費平均値種類クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 生産費平均値種類
        Public Const 営農類型 As String = "0"
        Public Const 生産費_総数 As String = "1"
        Public Const 生産費_総数以外 As String = "9"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {生産費平均値種類.営農類型, ""} _
            , {生産費平均値種類.生産費_総数, "総数"} _
            , {生産費平均値種類.生産費_総数以外, "総数以外"}
        }
    End Class

    ''' <summary>
    ''' 生産費平均値種類Weightクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 生産費平均値種類Weight
        Public Const 農産物_10ａ当たり = "1"
        Public Const 搾乳牛１頭当たり = "2"
        Public Const 子牛１頭当たり = "3"
        Public Const 育成牛_肥育牛１頭当たり = "4"
        Public Const 肥育豚１頭当たり = "5"
        Public Const 米_単位数量当たり = "6"
        Public Const 小麦_単位数量当たり = "7"
        Public Const 小麦組織_単位数量当たり = "8"
        Public Const 二条大麦_単位数量当たり = "9"
        Public Const 二条大麦経営分析_単位数量当たり = "10"
        Public Const 六条大麦_単位数量当たり = "11"
        Public Const 六条大麦経営分析_単位数量当たり = "12"
        Public Const はだか麦_単位数量当たり = "13"
        Public Const はだか麦経営分析_単位数量当たり = "14"
        Public Const そば_単位数量当たり = "15"
        Public Const そば経営分析_単位数量当たり = "16"
        Public Const 大豆_単位数量当たり = "17"
        Public Const 大豆組織_単位数量当たり = "18"
        Public Const 原料用かんしょ_単位数量当たり = "19"
        Public Const 原料用ばれいしょ_単位数量当たり = "20"
        Public Const 原料用ばれいしょ経営分析_単位数量当たり = "21"
        Public Const なたね_単位数量当たり = "22"
        Public Const なたね経営分析_単位数量当たり = "23"
        Public Const てんさい_単位数量当たり = "24"
        Public Const てんさい経営分析_単位数量当たり = "25"
        Public Const さとうきび_単位数量当たり = "26"
        Public Const さとうきび経営分析_単位数量当たり = "27"
        Public Const 牛乳_換算乳量100kg当たり = "28"
        Public Const 子牛_繁殖雌牛１頭当たり = "29"
        Public Const 肥育牛生体100kg当たり = "30"
        Public Const 肥育豚生体100kg当たり = "31"
        Public Const 経営体当たり = "32"
        Public Const 牛乳_実搾乳量100kg当たり = "33"

        Public Class 詳細
            Public 調査区分 As String()                '営農類型
            Public 平均値種類 As String              '名称
            Public 割る数 As String                  '割る数
        End Class

        Public Shared リスト As New Dictionary(Of String, 詳細) From {
            {生産費平均値種類Weight.農産物_10ａ当たり, New 詳細 With {.調査区分 = {調査区分.米生産費統計_個別, 調査区分.米生産費統計_組織法人, 調査区分.小麦生産費統計_個別, 調査区分.小麦生産費統計_組織法人, 調査区分.二条大麦生産費統計_個別, 調査区分.経営分析調査_二条大麦生産費, 調査区分.六条大麦生産費統計_個別, 調査区分.経営分析調査_六条大麦生産費, 調査区分.はだか麦生産費統計_個別, 調査区分.経営分析調査_はだか麦生産費, 調査区分.そば生産費統計_個別, 調査区分.経営分析調査_そば生産費, 調査区分.大豆生産費統計_個別, 調査区分.大豆生産費統計_組織法人, 調査区分.原料用かんしょ生産費統計_個別, 調査区分.原料用ばれいしょ生産費統計_個別, 調査区分.経営分析調査_原料用ばれいしょ生産費, 調査区分.なたね生産費統計_個別, 調査区分.経営分析調査_なたね生産費, 調査区分.てんさい生産費統計_個別, 調査区分.経営分析調査_てんさい生産費, 調査区分.さとうきび生産費統計_個別, 調査区分.経営分析調査_さとうきび生産費}, .平均値種類 = "2", .割る数 = "([S011416_1] / 10)"}} _
          , {生産費平均値種類Weight.搾乳牛１頭当たり, New 詳細 With {.調査区分 = {調査区分.牛乳生産費統計_個別, 調査区分.経営分析調査_牛乳生産費}, .平均値種類 = "2", .割る数 = "([S011239_1] / 13)"}} _
          , {生産費平均値種類Weight.子牛１頭当たり, New 詳細 With {.調査区分 = {調査区分.子牛生産費統計_個別, 調査区分.経営分析調査_子牛生産費}, .平均値種類 = "2", .割る数 = "[S010438_1]"}} _
          , {生産費平均値種類Weight.育成牛_肥育牛１頭当たり, New 詳細 With {.調査区分 = {調査区分.乳用雄育成牛生産費統計_個別, 調査区分.経営分析調査_乳用雄育成牛生産費, 調査区分.交雑種育成牛生産費統計_個別, 調査区分.経営分析調査_交雑種育成牛生産費, 調査区分.去勢若齢肥育牛生産費統計_個別, 調査区分.経営分析調査_去勢若齢肥育牛生産費, 調査区分.乳用雄肥育牛生産費統計_個別, 調査区分.経営分析調査_乳用雄肥育牛生産費, 調査区分.交雑種肥育牛生産費統計_個別, 調査区分.経営分析調査_交雑種肥育牛生産費}, .平均値種類 = "2", .割る数 = "[S010439_1]"}} _
          , {生産費平均値種類Weight.肥育豚１頭当たり, New 詳細 With {.調査区分 = {調査区分.肥育豚生産費統計_個別, 調査区分.経営分析調査_肥育豚生産費}, .平均値種類 = "2", .割る数 = "[S010238_1]"}} _
          , {生産費平均値種類Weight.米_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.米生産費統計_個別, 調査区分.米生産費統計_組織法人}, .平均値種類 = "3", .割る数 = "([S011149_1] / 60)"}} _
          , {生産費平均値種類Weight.小麦_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.小麦生産費統計_個別}, .平均値種類 = "3", .割る数 = "([S011149_1] / 60)"}} _
          , {生産費平均値種類Weight.小麦組織_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.小麦生産費統計_組織法人}, .平均値種類 = "3", .割る数 = "([S010749_1] / 60)"}} _
          , {生産費平均値種類Weight.二条大麦_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.二条大麦生産費統計_個別}, .平均値種類 = "3", .割る数 = "([S011149_1] / 50)"}} _
          , {生産費平均値種類Weight.二条大麦経営分析_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.経営分析調査_二条大麦生産費}, .平均値種類 = "3", .割る数 = "([S010749_1] / 50)"}} _
          , {生産費平均値種類Weight.六条大麦_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.六条大麦生産費統計_個別}, .平均値種類 = "3", .割る数 = "([S011149_1] / 50)"}} _
          , {生産費平均値種類Weight.六条大麦経営分析_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.経営分析調査_六条大麦生産費}, .平均値種類 = "3", .割る数 = "([S010749_1] / 50)"}} _
          , {生産費平均値種類Weight.はだか麦_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.はだか麦生産費統計_個別}, .平均値種類 = "3", .割る数 = "([S011149_1] / 60)"}} _
          , {生産費平均値種類Weight.はだか麦経営分析_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.経営分析調査_はだか麦生産費}, .平均値種類 = "3", .割る数 = "([S010749_1] / 60)"}} _
          , {生産費平均値種類Weight.そば_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.そば生産費統計_個別}, .平均値種類 = "3", .割る数 = "([S011149_1] / 45)"}} _
          , {生産費平均値種類Weight.そば経営分析_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.経営分析調査_そば生産費}, .平均値種類 = "3", .割る数 = "([S010749_1] / 45)"}} _
          , {生産費平均値種類Weight.大豆_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.大豆生産費統計_個別}, .平均値種類 = "3", .割る数 = "([S011149_1] / 60)"}} _
          , {生産費平均値種類Weight.大豆組織_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.大豆生産費統計_組織法人}, .平均値種類 = "3", .割る数 = "([S010749_1] / 60)"}} _
          , {生産費平均値種類Weight.原料用かんしょ_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.原料用かんしょ生産費統計_個別}, .平均値種類 = "3", .割る数 = "([S011149_1] / 100)"}} _
          , {生産費平均値種類Weight.原料用ばれいしょ_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.原料用ばれいしょ生産費統計_個別}, .平均値種類 = "3", .割る数 = "([S011149_1] / 100)"}} _
          , {生産費平均値種類Weight.原料用ばれいしょ経営分析_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.経営分析調査_原料用ばれいしょ生産費}, .平均値種類 = "3", .割る数 = "([S010749_1] / 100)"}} _
          , {生産費平均値種類Weight.なたね_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.なたね生産費統計_個別}, .平均値種類 = "3", .割る数 = "([S011149_1] / 60)"}} _
          , {生産費平均値種類Weight.なたね経営分析_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.経営分析調査_なたね生産費}, .平均値種類 = "3", .割る数 = "([S010749_1] / 60)"}} _
          , {生産費平均値種類Weight.てんさい_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.てんさい生産費統計_個別}, .平均値種類 = "3", .割る数 = "([S011149_1] / 1000)"}} _
          , {生産費平均値種類Weight.てんさい経営分析_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.経営分析調査_てんさい生産費}, .平均値種類 = "3", .割る数 = "([S010749_1] / 1000)"}} _
          , {生産費平均値種類Weight.さとうきび_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.さとうきび生産費統計_個別}, .平均値種類 = "3", .割る数 = "([S011149_1] / 1000)"}} _
          , {生産費平均値種類Weight.さとうきび経営分析_単位数量当たり, New 詳細 With {.調査区分 = {調査区分.経営分析調査_さとうきび生産費}, .平均値種類 = "3", .割る数 = "([S010749_1] / 1000)"}} _
          , {生産費平均値種類Weight.牛乳_換算乳量100kg当たり, New 詳細 With {.調査区分 = {調査区分.牛乳生産費統計_個別, 調査区分.経営分析調査_牛乳生産費}, .平均値種類 = "3", .割る数 = "([S010241_1] / 0.035 / 100)"}} _
          , {生産費平均値種類Weight.子牛_繁殖雌牛１頭当たり, New 詳細 With {.調査区分 = {調査区分.子牛生産費統計_個別, 調査区分.経営分析調査_子牛生産費}, .平均値種類 = "3", .割る数 = "[S010849_1]"}} _
          , {生産費平均値種類Weight.肥育牛生体100kg当たり, New 詳細 With {.調査区分 = {調査区分.乳用雄育成牛生産費統計_個別, 調査区分.経営分析調査_乳用雄育成牛生産費, 調査区分.交雑種育成牛生産費統計_個別, 調査区分.経営分析調査_交雑種育成牛生産費, 調査区分.去勢若齢肥育牛生産費統計_個別, 調査区分.経営分析調査_去勢若齢肥育牛生産費, 調査区分.乳用雄肥育牛生産費統計_個別, 調査区分.経営分析調査_乳用雄肥育牛生産費, 調査区分.交雑種肥育牛生産費統計_個別, 調査区分.経営分析調査_交雑種肥育牛生産費}, .平均値種類 = "3", .割る数 = "([S010441_1] / 100)"}} _
          , {生産費平均値種類Weight.肥育豚生体100kg当たり, New 詳細 With {.調査区分 = {調査区分.肥育豚生産費統計_個別, 調査区分.経営分析調査_肥育豚生産費}, .平均値種類 = "3", .割る数 = "([S010240_1] / 100)"}} _
          , {生産費平均値種類Weight.経営体当たり, New 詳細 With {.調査区分 = {調査区分.米生産費統計_個別, 調査区分.米生産費統計_組織法人, 調査区分.小麦生産費統計_個別, 調査区分.小麦生産費統計_組織法人, 調査区分.二条大麦生産費統計_個別, 調査区分.経営分析調査_二条大麦生産費, 調査区分.六条大麦生産費統計_個別, 調査区分.経営分析調査_六条大麦生産費, 調査区分.はだか麦生産費統計_個別, 調査区分.経営分析調査_はだか麦生産費, 調査区分.そば生産費統計_個別, 調査区分.経営分析調査_そば生産費, 調査区分.大豆生産費統計_個別, 調査区分.大豆生産費統計_組織法人, 調査区分.原料用かんしょ生産費統計_個別, 調査区分.原料用ばれいしょ生産費統計_個別, 調査区分.経営分析調査_原料用ばれいしょ生産費, 調査区分.なたね生産費統計_個別, 調査区分.経営分析調査_なたね生産費, 調査区分.てんさい生産費統計_個別, 調査区分.経営分析調査_てんさい生産費, 調査区分.さとうきび生産費統計_個別, 調査区分.経営分析調査_さとうきび生産費, 調査区分.牛乳生産費統計_個別, 調査区分.経営分析調査_牛乳生産費, 調査区分.子牛生産費統計_個別, 調査区分.経営分析調査_子牛生産費, 調査区分.乳用雄育成牛生産費統計_個別, 調査区分.経営分析調査_乳用雄育成牛生産費, 調査区分.交雑種育成牛生産費統計_個別, 調査区分.経営分析調査_交雑種育成牛生産費, 調査区分.去勢若齢肥育牛生産費統計_個別, 調査区分.経営分析調査_去勢若齢肥育牛生産費, 調査区分.乳用雄肥育牛生産費統計_個別, 調査区分.経営分析調査_乳用雄肥育牛生産費, 調査区分.交雑種肥育牛生産費統計_個別, 調査区分.経営分析調査_交雑種肥育牛生産費, 調査区分.肥育豚生産費統計_個別, 調査区分.経営分析調査_肥育豚生産費}, .平均値種類 = "4", .割る数 = "[S000008_1]"}} _
          , {生産費平均値種類Weight.牛乳_実搾乳量100kg当たり, New 詳細 With {.調査区分 = {調査区分.牛乳生産費統計_個別, 調査区分.経営分析調査_牛乳生産費}, .平均値種類 = "5", .割る数 = "([S010240_1] / 100)"}}
            }
    End Class

    ''' <summary>
    ''' 調査票クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 調査票
        ''' <summary>テーブル名称</summary>
        Public Shared テーブル名称 As New Dictionary(Of String, String()) From {
              {調査区分.営農類型別経営統計_個人, {"調査票＿農業経営＿営農類型＿個人", "調査票＿農業経営＿営農類型＿個人＿可変"}} _
            , {調査区分.営農類型別経営統計_法人, {"調査票＿農業経営＿営農類型＿法人"}} _
            , {調査区分.米生産費統計_個別, {"調査票＿農業経営＿米＿個別", "調査票＿農業経営＿米＿個別＿職員", "調査票＿農業経営＿米＿個別＿可変"}} _
            , {調査区分.小麦生産費統計_個別, {"調査票＿農業経営＿小麦＿個別", "調査票＿農業経営＿小麦＿個別＿職員", "調査票＿農業経営＿小麦＿個別＿可変"}} _
            , {調査区分.二条大麦生産費統計_個別, {"調査票＿農業経営＿二条大麦＿個別", "調査票＿農業経営＿二条大麦＿個別＿職員", "調査票＿農業経営＿二条大麦＿個別＿可変"}} _
            , {調査区分.六条大麦生産費統計_個別, {"調査票＿農業経営＿六条大麦＿個別", "調査票＿農業経営＿六条大麦＿個別＿職員", "調査票＿農業経営＿六条大麦＿個別＿可変"}} _
            , {調査区分.はだか麦生産費統計_個別, {"調査票＿農業経営＿はだか麦＿個別", "調査票＿農業経営＿はだか麦＿個別＿職員", "調査票＿農業経営＿はだか麦＿個別＿可変"}} _
            , {調査区分.そば生産費統計_個別, {"調査票＿農業経営＿そば＿個別", "調査票＿農業経営＿そば＿個別＿職員", "調査票＿農業経営＿そば＿個別＿可変"}} _
            , {調査区分.大豆生産費統計_個別, {"調査票＿農業経営＿大豆＿個別", "調査票＿農業経営＿大豆＿個別＿職員", "調査票＿農業経営＿大豆＿個別＿可変"}} _
            , {調査区分.原料用かんしょ生産費統計_個別, {"調査票＿農業経営＿原料用かんしょ＿個別", "調査票＿農業経営＿原料用かんしょ＿個別＿職員", "調査票＿農業経営＿原料用かんしょ＿個別＿可変"}} _
            , {調査区分.原料用ばれいしょ生産費統計_個別, {"調査票＿農業経営＿原料用ばれいしょ＿個別", "調査票＿農業経営＿原料用ばれいしょ＿個別＿職員", "調査票＿農業経営＿原料用ばれいしょ＿個別＿可変"}} _
            , {調査区分.なたね生産費統計_個別, {"調査票＿農業経営＿なたね＿個別", "調査票＿農業経営＿なたね＿個別＿職員", "調査票＿農業経営＿なたね＿個別＿可変"}} _
            , {調査区分.てんさい生産費統計_個別, {"調査票＿農業経営＿てんさい＿個別", "調査票＿農業経営＿てんさい＿個別＿職員", "調査票＿農業経営＿てんさい＿個別＿可変"}} _
            , {調査区分.さとうきび生産費統計_個別, {"調査票＿農業経営＿さとうきび＿個別", "調査票＿農業経営＿さとうきび＿個別＿職員", "調査票＿農業経営＿さとうきび＿個別＿可変"}} _
            , {調査区分.米生産費統計_組織法人, {"調査票＿農業経営＿米＿組織", "調査票＿農業経営＿米＿組織＿職員", "調査票＿農業経営＿米＿組織＿可変"}} _
            , {調査区分.小麦生産費統計_組織法人, {"調査票＿農業経営＿小麦＿組織", "調査票＿農業経営＿小麦＿組織＿職員", "調査票＿農業経営＿小麦＿組織＿可変"}} _
            , {調査区分.大豆生産費統計_組織法人, {"調査票＿農業経営＿大豆＿組織", "調査票＿農業経営＿大豆＿組織＿職員", "調査票＿農業経営＿大豆＿組織＿可変"}} _
            , {調査区分.牛乳生産費統計_個別, {"調査票＿農業経営＿牛乳＿個別", "調査票＿農業経営＿牛乳＿個別＿職員", "調査票＿農業経営＿牛乳＿個別＿可変"}} _
            , {調査区分.子牛生産費統計_個別, {"調査票＿農業経営＿子牛＿個別", "調査票＿農業経営＿子牛＿個別＿職員", "調査票＿農業経営＿子牛＿個別＿可変"}} _
            , {調査区分.乳用雄育成牛生産費統計_個別, {"調査票＿農業経営＿乳用雄育成牛＿個別", "調査票＿農業経営＿乳用雄育成牛＿個別＿職員", "調査票＿農業経営＿乳用雄育成牛＿個別＿可変"}} _
            , {調査区分.交雑種育成牛生産費統計_個別, {"調査票＿農業経営＿交雑種育成牛＿個別", "調査票＿農業経営＿交雑種育成牛＿個別＿職員", "調査票＿農業経営＿交雑種育成牛＿個別＿可変"}} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, {"調査票＿農業経営＿去勢若齢肥育牛＿個別", "調査票＿農業経営＿去勢若齢肥育牛＿個別＿職員", "調査票＿農業経営＿去勢若齢肥育牛＿個別＿可変"}} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, {"調査票＿農業経営＿乳用雄肥育牛＿個別", "調査票＿農業経営＿乳用雄肥育牛＿個別＿職員", "調査票＿農業経営＿乳用雄肥育牛＿個別＿可変"}} _
            , {調査区分.交雑種肥育牛生産費統計_個別, {"調査票＿農業経営＿交雑種肥育牛＿個別", "調査票＿農業経営＿交雑種肥育牛＿個別＿職員", "調査票＿農業経営＿交雑種肥育牛＿個別＿可変"}} _
            , {調査区分.肥育豚生産費統計_個別, {"調査票＿農業経営＿肥育豚＿個別", "調査票＿農業経営＿肥育豚＿個別＿職員", "調査票＿農業経営＿肥育豚＿個別＿可変"}} _
            , {調査区分.経営分析調査_二条大麦生産費, {"調査票＿経営分析＿二条大麦", "調査票＿経営分析＿二条大麦＿職員", "調査票＿経営分析＿二条大麦＿可変"}} _
            , {調査区分.経営分析調査_六条大麦生産費, {"調査票＿経営分析＿六条大麦", "調査票＿経営分析＿六条大麦＿職員", "調査票＿経営分析＿六条大麦＿可変"}} _
            , {調査区分.経営分析調査_はだか麦生産費, {"調査票＿経営分析＿はだか麦", "調査票＿経営分析＿はだか麦＿職員", "調査票＿経営分析＿はだか麦＿可変"}} _
            , {調査区分.経営分析調査_そば生産費, {"調査票＿経営分析＿そば", "調査票＿経営分析＿そば＿職員", "調査票＿経営分析＿そば＿可変"}} _
            , {調査区分.経営分析調査_原料用ばれいしょ生産費, {"調査票＿経営分析＿原料用ばれいしょ", "調査票＿経営分析＿原料用ばれいしょ＿職員", "調査票＿経営分析＿原料用ばれいしょ＿可変"}} _
            , {調査区分.経営分析調査_なたね生産費, {"調査票＿経営分析＿なたね", "調査票＿経営分析＿なたね＿職員", "調査票＿経営分析＿なたね＿可変"}} _
            , {調査区分.経営分析調査_てんさい生産費, {"調査票＿経営分析＿てんさい", "調査票＿経営分析＿てんさい＿職員", "調査票＿経営分析＿てんさい＿可変"}} _
            , {調査区分.経営分析調査_さとうきび生産費, {"調査票＿経営分析＿さとうきび", "調査票＿経営分析＿さとうきび＿職員", "調査票＿経営分析＿さとうきび＿可変"}} _
            , {調査区分.経営分析調査_牛乳生産費, {"調査票＿経営分析＿牛乳", "調査票＿経営分析＿牛乳＿職員", "調査票＿経営分析＿牛乳＿可変"}} _
            , {調査区分.経営分析調査_子牛生産費, {"調査票＿経営分析＿子牛", "調査票＿経営分析＿子牛＿職員", "調査票＿経営分析＿子牛＿可変"}} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, {"調査票＿経営分析＿乳用雄育成牛", "調査票＿経営分析＿乳用雄育成牛＿職員", "調査票＿経営分析＿乳用雄育成牛＿可変"}} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, {"調査票＿経営分析＿交雑種育成牛", "調査票＿経営分析＿交雑種育成牛＿職員", "調査票＿経営分析＿交雑種育成牛＿可変"}} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, {"調査票＿経営分析＿去勢若齢肥育牛", "調査票＿経営分析＿去勢若齢肥育牛＿職員", "調査票＿経営分析＿去勢若齢肥育牛＿可変"}} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, {"調査票＿経営分析＿乳用雄肥育牛", "調査票＿経営分析＿乳用雄肥育牛＿職員", "調査票＿経営分析＿乳用雄肥育牛＿可変"}} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, {"調査票＿経営分析＿交雑種肥育牛", "調査票＿経営分析＿交雑種肥育牛＿職員", "調査票＿経営分析＿交雑種肥育牛＿可変"}} _
            , {調査区分.経営分析調査_肥育豚生産費, {"調査票＿経営分析＿肥育豚", "調査票＿経営分析＿肥育豚＿職員", "調査票＿経営分析＿肥育豚＿可変"}}
        }

        ''' <summary>項目番号クラス</summary>
        Public Class 項目番号
            Public Const 調査年 As String = "Q00000101"
            Public Const 都道府県 As String = "Q00001001"
            Public Const 市区町村 As String = "Q00001002"
            Public Const 旧市区町村 As String = "Q00001003"
            Public Const 農業集落 As String = "Q00001004"
            Public Const 調査区 As String = "Q00001005"
            Public Const 客体番号 As String = "Q00001006"

            Public Class 営農類型別経営統計
                Public Const 営農類型 As String = "Q00000201"
                Public Const 指定品目名 As String = "Q00000301"
            End Class

            Public Class 農産物生産費
                Public Const 対象品目 As String = "Q00001101"
                Public Const 経営種類 As String = "Q00001201"
            End Class

            Public Class 畜産物生産費
                Public Const 生産費区分 As String = "Q00000201"
            End Class

            '---REV_003 ADD START
            Public Shared 牛識別番号 As New Dictionary(Of String, String) From {
              {調査区分.牛乳生産費統計_個別, "Q11021801"} _
            , {調査区分.子牛生産費統計_個別, "Q02030301"} _
            , {調査区分.乳用雄育成牛生産費統計_個別, "Q02021401"} _
            , {調査区分.交雑種育成牛生産費統計_個別, "Q02021401"} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, "Q02021401"} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, "Q02021401"} _
            , {調査区分.交雑種肥育牛生産費統計_個別, "Q02021401"} _
            , {調査区分.経営分析調査_牛乳生産費, "Q11021801"} _
            , {調査区分.経営分析調査_子牛生産費, "Q02030301"} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, "Q02021401"} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, "Q02021401"} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "Q02021401"} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, "Q02021401"} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, "Q02021401"}
            }

            Public Shared 牛農家団体コード As New Dictionary(Of String, String) From {
              {調査区分.牛乳生産費統計_個別, "Q11010101"} _
            , {調査区分.子牛生産費統計_個別, "Q02010101"} _
            , {調査区分.乳用雄育成牛生産費統計_個別, "Q02010101"} _
            , {調査区分.交雑種育成牛生産費統計_個別, "Q02010101"} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, "Q02010101"} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, "Q02010101"} _
            , {調査区分.交雑種肥育牛生産費統計_個別, "Q02010101"} _
            , {調査区分.経営分析調査_牛乳生産費, "Q11010101"} _
            , {調査区分.経営分析調査_子牛生産費, "Q02010101"} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, "Q02010101"} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, "Q02010101"} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "Q02010101"} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, "Q02010101"} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, "Q02010101"}
            }
            '---REV_003 ADD END
        End Class

        ''' <summary>営農類型</summary>
        Public Shared 営農類型 As New Dictionary(Of String, String) From {
              {調査区分.営農類型別経営統計_個人, ComConst.調査票.項目番号.営農類型別経営統計.営農類型} _
            , {調査区分.営農類型別経営統計_法人, ComConst.調査票.項目番号.営農類型別経営統計.営農類型}
        }

        ''' <summary>入力用ファイル名称</summary>
        'REV-004 START-------------------------------------------------------------------------------------------
        Public Shared 入力用ファイル名称 As New Dictionary(Of Tuple(Of String, String), String) From {
              {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.調査票項目2015), "営農類型別経営統計（個人）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.調査票項目2015), "営農類型別経営統計（法人）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "米生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "小麦生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "二条大麦生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "六条大麦生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "はだか麦生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "そば生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "大豆生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "原料用かんしょ生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "原料用ばれいしょ生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "なたね生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "てんさい生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "さとうきび生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.調査票項目2015), "米生産費統計（組織法人）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.調査票項目2015), "小麦生産費統計（組織法人）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.調査票項目2015), "大豆生産費統計（組織法人）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "牛乳生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "子牛生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "乳用雄育成牛生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "交雑種育成牛生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "去勢若齢肥育牛生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "乳用雄肥育牛生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "交雑種肥育牛生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.調査票項目2015), "肥育豚生産費統計（個別）_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_二条大麦生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_六条大麦生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_はだか麦生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_そば生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_原料用ばれいしょ生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_なたね生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_てんさい生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_さとうきび生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_牛乳生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_子牛生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_乳用雄育成牛生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_交雑種育成牛生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_去勢若齢肥育牛生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_乳用雄肥育牛生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_交雑種肥育牛生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.調査票項目2015), "経営分析調査_肥育豚生産費_電子調査票_2015.xlsm"} _
            , {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.調査票項目2020), "営農類型別経営統計（個人）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.調査票項目2020), "営農類型別経営統計（法人）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "米生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "小麦生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "二条大麦生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "六条大麦生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "はだか麦生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "そば生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "大豆生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "原料用かんしょ生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "原料用ばれいしょ生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "なたね生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "てんさい生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "さとうきび生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.調査票項目2020), "米生産費統計（組織法人）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.調査票項目2020), "小麦生産費統計（組織法人）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.調査票項目2020), "大豆生産費統計（組織法人）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "牛乳生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "子牛生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "乳用雄育成牛生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "交雑種育成牛生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "去勢若齢肥育牛生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "乳用雄肥育牛生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "交雑種肥育牛生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.調査票項目2020), "肥育豚生産費統計（個別）_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_二条大麦生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_六条大麦生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_はだか麦生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_そば生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_原料用ばれいしょ生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_なたね生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_てんさい生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_さとうきび生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_牛乳生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_子牛生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_乳用雄育成牛生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_交雑種育成牛生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_去勢若齢肥育牛生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_乳用雄肥育牛生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_交雑種肥育牛生産費_電子調査票.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.調査票項目2020), "経営分析調査_肥育豚生産費_電子調査票.xlsm"}
        }
        'REV-004 END-------------------------------------------------------------------------------------------


        '---REV.003 ADD START
        Public Class 取り込みファイル判定
            Public Class 判定内容
                Public シート名 As String
                Public セル番号 As String
                Public セル内容 As String
            End Class

            Public Shared 判定内容一覧 As New Dictionary(Of String, 判定内容) From {
                 {調査区分.米生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.小麦生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.二条大麦生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.六条大麦生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.はだか麦生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.そば生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.大豆生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.原料用かんしょ生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.原料用ばれいしょ生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.なたね生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.てんさい生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.さとうきび生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.米生産費統計_組織法人, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.小麦生産費統計_組織法人, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.大豆生産費統計_組織法人, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.牛乳生産費統計_個別, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.子牛生産費統計_個別, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.乳用雄育成牛生産費統計_個別, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.交雑種育成牛生産費統計_個別, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.去勢若齢肥育牛生産費統計_個別, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.乳用雄肥育牛生産費統計_個別, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.交雑種肥育牛生産費統計_個別, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.肥育豚生産費統計_個別, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "調査対象畜負担割合（％）"}},
                 {調査区分.経営分析調査_二条大麦生産費, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.経営分析調査_六条大麦生産費, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.経営分析調査_はだか麦生産費, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.経営分析調査_そば生産費, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.経営分析調査_原料用ばれいしょ生産費, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.経営分析調査_なたね生産費, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.経営分析調査_てんさい生産費, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.経営分析調査_さとうきび生産費, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "Q10", .セル内容 = "対象品目負担割合（％）"}},
                 {調査区分.経営分析調査_牛乳生産費, New 判定内容 With {.シート名 = "【９】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.経営分析調査_子牛生産費, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.経営分析調査_乳用雄育成牛生産費, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.経営分析調査_交雑種育成牛生産費, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.経営分析調査_去勢若齢肥育牛生産費, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.経営分析調査_乳用雄肥育牛生産費, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.経営分析調査_交雑種肥育牛生産費, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "負担割合（％）"}},
                 {調査区分.経営分析調査_肥育豚生産費, New 判定内容 With {.シート名 = "【10】農業機械", .セル番号 = "P12", .セル内容 = "調査対象畜負担割合（％）"}}
                 }

        End Class
        '---REV.003 ADD END

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            Public Class Report
                'テンプレートファイル名
                Public tempFileName As String
                '帳票名
                Public reportName As String
            End Class

            'REV-004 START-------------------------------------------------------------------------------------------
            '連絡票No286対応　2022/03/03　START
            Public Shared リスト As New Dictionary(Of Tuple(Of String, String), Report) From {
                  {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.調査票項目2015)), .reportName = "営農類型別経営統計（個人）_電子調査票"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.調査票項目2015)), .reportName = "営農類型別経営統計（法人）_電子調査票"}} _
                , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "米生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "小麦生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "二条大麦生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "六条大麦生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "はだか麦生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "そば生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "大豆生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "原料用かんしょ生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "原料用ばれいしょ生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "なたね生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "てんさい生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "さとうきび生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.調査票項目2015)), .reportName = "米生産費統計（組織法人）_電子調査票"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.調査票項目2015)), .reportName = "小麦生産費統計（組織法人）_電子調査票"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.調査票項目2015)), .reportName = "大豆生産費統計（組織法人）_電子調査票"}} _
                , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "牛乳生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "子牛生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "乳用雄育成牛生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "交雑種育成牛生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "去勢若齢肥育牛生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "乳用雄肥育牛生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "交雑種肥育牛生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.調査票項目2015)), .reportName = "肥育豚生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_二条大麦生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_六条大麦生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_はだか麦生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_そば生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_原料用ばれいしょ生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_なたね生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_てんさい生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_さとうきび生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_牛乳生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_子牛生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_乳用雄育成牛生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_交雑種育成牛生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_去勢若齢肥育牛生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_乳用雄肥育牛生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_交雑種肥育牛生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.調査票項目2015), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.調査票項目2015)), .reportName = "経営分析調査_肥育豚生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.調査票項目2020)), .reportName = "営農類型別経営統計（個人）_電子調査票"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.調査票項目2020)), .reportName = "営農類型別経営統計（法人）_電子調査票"}} _
                , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "米生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "小麦生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "二条大麦生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "六条大麦生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "はだか麦生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "そば生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "大豆生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "原料用かんしょ生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "原料用ばれいしょ生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "なたね生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "てんさい生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "さとうきび生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.調査票項目2020)), .reportName = "米生産費統計（組織法人）_電子調査票"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.調査票項目2020)), .reportName = "小麦生産費統計（組織法人）_電子調査票"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.調査票項目2020)), .reportName = "大豆生産費統計（組織法人）_電子調査票"}} _
                , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "牛乳生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "子牛生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "乳用雄育成牛生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "交雑種育成牛生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "去勢若齢肥育牛生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "乳用雄肥育牛生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "交雑種肥育牛生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.調査票項目2020)), .reportName = "肥育豚生産費統計（個別）_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_二条大麦生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_六条大麦生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_はだか麦生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_そば生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_原料用ばれいしょ生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_なたね生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_てんさい生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_さとうきび生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_牛乳生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_子牛生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_乳用雄育成牛生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_交雑種育成牛生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_去勢若齢肥育牛生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_乳用雄肥育牛生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_交雑種肥育牛生産費_電子調査票"}} _
                , {Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.調査票項目2020), New Report With {.tempFileName = 調査票.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.調査票項目2020)), .reportName = "経営分析調査_肥育豚生産費_電子調査票"}}
            }

            '連絡票No286対応　2022/03/03　END
            'REV-004 END-------------------------------------------------------------------------------------------
        End Class

        ''' <summary>シートデータ範囲</summary>
        ''' REV-004 START------------------------
        ''' 　2020用のシートデータ範囲の追加
        ''' 　
        ''' ---REV_003 MOD
        '''  子牛生産費以下の通り定義変更
        ''' 　【２】2牛取引情報 　"A1:R2014"⇒"A1:AE2014"
        ''' 　【２】３牛取引情報　"A1:F510"⇒"A1:Z510"
        ''' 　【２】対象畜概要１　"A1:AS2020"⇒"A1:EZ2020"
        ''' 全生産費調査区分の「農業機械」シートの定義を以下の通り変更
        ''' 　・農産物・肥育豚→4列追加
        ''' 　・上記以外→6列追加
        ''' '20220203  調査票項目2015 の 「表紙_個別」→「表紙_組織」 に修正 (米生産費統計_組織法人,小麦生産費統計_組織法人,大豆生産費統計_組織法人)
        ''' '20220203  調査票項目2015 の 「表紙_個別」→「表紙_経営分析」 に修正 (経営分析調査_二条大麦生産費,経営分析調査_六条大麦生産費,経営分析調査_はだか麦生産費,経営分析調査_そば生産費,
        ''' '                                                                    経営分析調査_原料用ばれいしょ生産費,経営分析調査_なたね生産費,経営分析調査_てんさい生産費,経営分析調査_さとうきび生産費)
        ''' ---REV_009↓
        ''' 農産物・調査票項目2020の範囲変更
        ''' 【３】１種苗費_米、【３】１種苗費_米以外
        ''' 【11】土地_１所有地_米、【11】土地_１所有地_さとうきび、【11】土地_１所有地_米・さとうきび以外
        ''' 【11】土地_２借入地_米、【11】土地_２借入地_さとうきび、【11】土地_２借入地_米・さとうきび以外
        ''' ---REV_009↑
        ''' 
        ''' ---REV_017↓
        ''' 営農類型別経営統計_法人2020の不要シート削除
        '''  13_02_農産関連事業収支
        ''' ---REV_017↑
        Public Shared シートデータ範囲 As New Dictionary(Of Tuple(Of String, String), Dictionary(Of String, String)) From {
              {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"00_表紙", "A1:AC67"},
                    {"はじめに・01_現況", "A1:AR84"},
                    {"02_損益計算書", "A1:BJ44"},
                    {"03_貸借対照表（作成）", "A1:AR44"},
                    {"03_貸借対照表（未作成）", "A1:AR48"},
                    {"04_事業収支の概要・05_投資と資金", "A1:AQ48"},
                    {"06_固定資産・07土地面積", "A1:AO50"},
                    {"08_01_生産概況、農畜産物収入", "A1:CI58"},
                    {"08_02_生産概況、農畜産物収入", "A1:CI64"},
                    {"09_共済・補助金等", "A1:AF65"},
                    {"10_労働", "A1:AQ37"},
                    {"11_労働（指定品目）", "A1:AE48"},
                    {"12_01_農産関連事業収支", "A1:AR32"},
                    {"12_02_農産関連事業収支", "A1:AR20"}
                }
              } _
              , {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"00_表紙", "A1:AC75"},
                    {"はじめに・01_現況", "A1:AR93"},
                    {"02_損益計算書", "A1:BJ44"},
                    {"03_貸借対照表（作成）", "A1:AR44"},
                    {"03_貸借対照表（未作成）", "A1:AR48"},
                    {"04_事業収支の概要・05_投資と資金", "A1:AQ48"},
                    {"06_固定資産・07土地面積", "A1:AO31"},
                    {"08_01_生産概況、農畜産物収入", "A1:CI58"},
                    {"08_02_生産概況、農畜産物収入", "A1:CI71"},
                    {"09_共済・補助金等", "A1:AF113"},
                    {"10_労働", "A1:AQ125"},
                    {"11_労働（指定品目）", "A1:AE48"},
                    {"12_01_農産関連事業収支", "A1:AR32"}
                }
              } _
            , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"00_表紙", "A1:AC71"},
                    {"01_02_はじめに", "A1:AR48"},
                    {"02_01_貸借対照表", "A1:AJ45"},
                    {"02_02_貸借対照表", "A1:AD40"},
                    {"03_投資と資金", "A1:AR39"},
                    {"04_損益計算書", "A1:AQ35"},
                    {"05_01_製造原価、販管費", "A1:AQ39"},
                    {"05_02_製造原価、販管費", "A1:AQ36"},
                    {"06_給与の状況、07_土地面積", "A1:AO45"},
                    {"08_主要農業固定資産", "A1:AO34"},
                    {"09_01_生産概況+農畜産物収入", "A1:CA58"},
                    {"09_02_生産概況＋農畜産物収入", "A1:CN65"},
                    {"10_01_受託収入", "A1:AN41"},
                    {"10_02_受託収入", "A1:AO33"},
                    {"11_01_制度受取金・積立金", "A1:AD33"},
                    {"11_02_制度受取金・積立金", "A1:AD36"},
                    {"12_01_労働の概況", "A1:AQ45"},
                    {"12_02_労働の概況", "A1:AR52"},
                    {"13_01_生産関連事業収支", "A1:AR42"},
                    {"13_02_農産関連事業収支", "A1:AR27"}
                }
              } _
            , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"00_表紙", "A1:AC71"},
                    {"01_02_はじめに", "A1:AR48"},
                    {"02_01_貸借対照表", "A1:AJ45"},
                    {"02_02_貸借対照表", "A1:AD40"},
                    {"03_投資と資金", "A1:AR39"},
                    {"04_損益計算書", "A1:AQ35"},
                    {"05_01_製造原価、販管費", "A1:AQ42"},
                    {"05_02_製造原価、販管費", "A1:AQ36"},
                    {"06_給与の状況、07_土地面積", "A1:AO46"},
                    {"08_主要農業固定資産", "A1:AO23"},
                    {"09_01_生産概況+農畜産物収入", "A1:CA58"},
                    {"09_02_生産概況＋農畜産物収入", "A1:CN65"},
                    {"10_01_受託収入", "A1:AN41"},
                    {"10_02_受託収入", "A1:AO33"},
                    {"11_01_制度受取金・積立金", "A1:AD44"},
                    {"11_02_制度受取金・積立金", "A1:AD61"},
                    {"12_01_労働の概況", "A1:AQ45"},
                    {"12_02_労働の概況", "A1:AR52"},
                    {"13_01_生産関連事業収支", "A1:AR38"}
                }
              } _
            , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況②_個別_米", "A1:AU49"},
                    {"【１】経営の概況③_個別_米", "A1:AR44"},
                    {"【２】生産物の販売等の状況①_米", "A1:AG81"},
                    {"【２】生産物の販売等の状況②_米", "A1:AL39"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米", "A1:AB73"},
                    {"【３】２肥料費_米", "A1:AH124"},
                    {"【３】３農業薬剤費_米", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米", "A1:AT95"},
                    {"【３】６賃借料及び料金_米", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米", "A1:Q243"},
                    {"【11】土地_２借入地_米", "A1:Q241"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"【12】作業別労働時間②_米", "A1:Z9"},
                    {"【13】飼料用米①_米", "A1:AM37"},
                    {"【13】飼料用米②_米", "A1:Z39"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況②_個別_米", "A1:AU49"},
                    {"【１】経営の概況③_個別_米", "A1:AR108"},
                    {"【２】生産物の販売等の状況①_米", "A1:AG81"},
                    {"【２】生産物の販売等の状況②_米", "A1:AL39"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米", "A1:AC73"},
                    {"【３】２肥料費_米", "A1:AJ124"},
                    {"【３】３農業薬剤費_米", "A1:AC80"},
                    {"【３】４光熱動力費・５諸材料費_米", "A1:AU95"},
                    {"【３】６賃借料及び料金_米", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ132"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米", "A1:Q1093"},
                    {"【11】土地_２借入地_米", "A1:Q1091"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"【12】作業別労働時間②_米", "A1:Z9"},
                    {"【13】飼料用米①_米", "A1:AM37"},
                    {"【13】飼料用米②_米", "A1:Z41"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_麦・大豆", "A1:AN45"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AI85"},
                    {"【２】生産物の販売等の状況③_個別_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:I24"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_麦・大豆", "A1:AN113"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AI85"},
                    {"【２】生産物の販売等の状況③_個別_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AK124"},
                    {"【３】３農業薬剤費_米以外", "A1:AC80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AU95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T1093"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q1091"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_麦・大豆", "A1:AN45"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AI85"},
                    {"【２】生産物の販売等の状況②_二条大麦", "A1:AD83"},
                    {"【２】生産物の販売等の状況③_個別_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_麦・大豆", "A1:AN114"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AI85"},
                    {"【２】生産物の販売等の状況②_二条大麦", "A1:AD83"},
                    {"【２】生産物の販売等の状況③_個別_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AK124"},
                    {"【３】３農業薬剤費_米以外", "A1:AC80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AU95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T1093"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q1091"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_麦・大豆", "A1:AN45"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AI85"},
                    {"【２】生産物の販売等の状況②_二条大麦", "A1:AD83"},
                    {"【２】生産物の販売等の状況③_個別_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_麦・大豆", "A1:AN114"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AI85"},
                    {"【２】生産物の販売等の状況②_二条大麦", "A1:AD83"},
                    {"【２】生産物の販売等の状況③_個別_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AK124"},
                    {"【３】３農業薬剤費_米以外", "A1:AC80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AU95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T1093"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q1091"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_麦・大豆", "A1:AN45"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AI85"},
                    {"【２】生産物の販売等の状況②_二条大麦", "A1:AD83"},
                    {"【２】生産物の販売等の状況③_個別_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_麦・大豆", "A1:AN114"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AI85"},
                    {"【２】生産物の販売等の状況②_二条大麦", "A1:AD83"},
                    {"【２】生産物の販売等の状況③_個別_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AK124"},
                    {"【３】３農業薬剤費_米以外", "A1:AC80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AU95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T1093"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q1091"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_なたね・そば", "A1:AN44"},
                    {"【２】生産物の販売等の状況①_そば", "A1:AD77"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_なたね・そば", "A1:AN114"},
                    {"【２】生産物の販売等の状況①_そば", "A1:AE77"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AK124"},
                    {"【３】３農業薬剤費_米以外", "A1:AC80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AU95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T1093"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q1091"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_麦・大豆", "A1:AN45"},
                    {"【２】生産物の販売等の状況①_大豆", "A1:AF94"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_麦・大豆", "A1:AN114"},
                    {"【２】生産物の販売等の状況①_大豆", "A1:AG94"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AK124"},
                    {"【３】３農業薬剤費_米以外", "A1:AC80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AU95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T1093"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q1091"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_原かん", "A1:AN15"},
                    {"【２】生産物の販売等の状況①_原かん", "A1:AD57"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_原かん", "A1:AN99"},
                    {"【２】生産物の販売等の状況①_原かん", "A1:AD57"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AI124"},
                    {"【３】３農業薬剤費_米以外", "A1:AB80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T1093"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q1091"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_原ばれ", "A1:AN29"},
                    {"【２】生産物の販売等の状況①_原ばれ", "A1:AF57"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_原ばれ", "A1:AN99"},
                    {"【２】生産物の販売等の状況①_原ばれ", "A1:AF57"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AI124"},
                    {"【３】３農業薬剤費_米以外", "A1:AB80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T1093"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q1091"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_なたね・そば", "A1:AN44"},
                    {"【２】生産物の販売等の状況①_なたね", "A1:AE59"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_個別_なたね・そば", "A1:AN113"},
                    {"【２】生産物の販売等の状況①_なたね", "A1:AE59"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AK124"},
                    {"【３】３農業薬剤費_米以外", "A1:AC80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AU95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T1093"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q1091"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_てんさい", "A1:AN39"},
                    {"【２】生産物の販売等の状況①_てんさい", "A1:AI58"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_てんさい", "A1:AN109"},
                    {"【２】生産物の販売等の状況①_てんさい", "A1:AI58"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:AB80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T1093"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q1091"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_さとうきび", "A1:AN35"},
                    {"【２】生産物の販売等の状況①_さとうきび", "A1:AI59"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_さとうきび", "A1:T243"},
                    {"【11】土地_２借入地_さとうきび", "A1:Q241"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営の概況①_個別", "A1:AH47"},
                    {"【１】経営の概況③_さとうきび", "A1:AN111"},
                    {"【２】生産物の販売等の状況①_さとうきび", "A1:AI59"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:AF80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P64"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_さとうきび", "A1:T1093"},
                    {"【11】土地_２借入地_さとうきび", "A1:Q1091"},
                    {"【12】作業別労働時間①_個別", "A1:BS56"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_組織", "A1:BF63"},
                    {"【１】経営の概況①_組織_米", "A1:AH53"},
                    {"【１】経営の概況②_組織_米", "A1:AU58"},
                    {"【１】経営の概況③_組織_米", "A1:AR47"},
                    {"【２】生産物の販売等の状況①_米", "A1:AG81"},
                    {"【２】生産物の販売等の状況②_米", "A1:AL39"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米", "A1:AB73"},
                    {"【３】２肥料費_米", "A1:AH124"},
                    {"【３】３農業薬剤費_米", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米", "A1:AT95"},
                    {"【３】６賃借料及び料金_米", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米", "A1:Q243"},
                    {"【11】土地_２借入地_米", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"【12】作業別労働時間②_米", "A1:Z9"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_組織", "A1:BF64"},
                    {"【１】経営の概況①_組織_米", "A1:AH53"},
                    {"【１】経営の概況②_組織_米", "A1:AU58"},
                    {"【１】経営の概況③_組織_米", "A1:AR118"},
                    {"【２】生産物の販売等の状況①_米", "A1:AG81"},
                    {"【２】生産物の販売等の状況②_米", "A1:AL39"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米", "A1:AC73"},
                    {"【３】２肥料費_米", "A1:AI124"},
                    {"【３】３農業薬剤費_米", "A1:AB80"},
                    {"【３】４光熱動力費・５諸材料費_米", "A1:AT95"},
                    {"【３】６賃借料及び料金_米", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米", "A1:Q1093"},
                    {"【11】土地_２借入地_米", "A1:Q1091"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"【12】作業別労働時間②_米", "A1:Z9"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_組織", "A1:BF63"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN49"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AI85"},
                    {"【２】生産物の販売等の状況③_組織_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_組織", "A1:BF64"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN118"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AI85"},
                    {"【２】生産物の販売等の状況③_組織_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AK124"},
                    {"【３】３農業薬剤費_米以外", "A1:AC80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AU95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T1093"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q1091"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_組織", "A1:BF63"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN49"},
                    {"【２】生産物の販売等の状況①_大豆", "A1:AF94"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_組織", "A1:BF64"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN118"},
                    {"【２】生産物の販売等の状況①_大豆", "A1:AF94"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AF93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:AB80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T1093"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q1091"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"表紙_個別", "A1:BF63"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】１生乳", "A1:AT206"},
                    {"【２】２子牛、３きゅう肥", "A1:R46"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３種付料～９その他の資材等", "A1:S246"},
                    {"【４】物件税～【６】借入金", "A1:V89"},
                    {"減価償却費", "A1:J73"},
                    {"【７】建物", "A1:BH344"},
                    {"【8】自動車", "A1:BH328"},
                    {"【９】農業機械", "A1:BE333"},
                    {"【10】農具購入費", "A1:AD92"},
                    {"【11】農家団体コード", "A1:H20"},
                    {"【11】搾乳牛所有状況", "A1:AA2015"},
                    {"【11】搾乳牛所有状況２", "A1:AR2010"},
                    {"飼養頭数、負担割合", "A1:U36"},
                    {"子牛の概要", "A1:I12"},
                    {"搾乳牛の概要、資本額", "A1:L15"},
                    {"【12】労働時間", "A1:BF113"},
                    {"【13】地代・搾乳牛負担割合・自給牧草", "A1:BD64"}
                }
              } _
            , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BT67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】１生乳", "A1:AT206"},
                    {"【２】２子牛、３きゅう肥", "A1:R46"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３種付料～９その他の資材等", "A1:S246"},
                    {"【４】物件税～【６】借入金", "A1:V89"},
                    {"減価償却費", "A1:J73"},
                    {"【７】建物", "A1:BH344"},
                    {"【8】自動車", "A1:BH328"},
                    {"【９】農業機械", "A1:BE333"},
                    {"【10】農具購入費", "A1:AD92"},
                    {"【11】農家団体コード", "A1:H20"},
                    {"【11】搾乳牛所有状況", "A1:AD2015"},
                    {"【11】搾乳牛所有状況２", "A1:BI2010"},
                    {"飼養頭数、負担割合", "A1:U36"},
                    {"子牛の概要", "A1:I14"},
                    {"搾乳牛の概要、資本額", "A1:L15"},
                    {"【12】労働時間", "A1:BF113"},
                    {"【13】地代・搾乳牛負担割合・自給牧草", "A1:BD64"}
                }
              } _
            , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】農家団体コード", "A1:E18"},
                    {"【２】2牛取引情報", "A1:AE2014"},
                    {"【２】３牛取引情報", "A1:Z510"},
                    {"【２】対象畜概要１", "A1:EZ2020"},
                    {"【２】対象畜概要２", "A1:T49"},
                    {"【２】４きゅう肥", "A1:R34"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３種付料～９その他の資材等", "A1:S240"},
                    {"【４】物件税～【７】出荷経費", "A1:U96"},
                    {"減価償却費", "A1:V65"},
                    {"【８】建物", "A1:BH331"},
                    {"【９】自動車", "A1:BH326"},
                    {"【10】農業機械", "A1:BE332"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:BF113"},
                    {"【13】地代", "A1:AB38"}
                }
              } _
            , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】農家団体コード", "A1:E18"},
                    {"【２】2牛取引情報", "A1:AE2015"},
                    {"【２】３牛取引情報", "A1:Z510"},
                    {"【２】対象畜概要１", "A1:EZ2020"},
                    {"【２】対象畜概要２", "A1:T49"},
                    {"【２】４きゅう肥", "A1:R34"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３種付料～９その他の資材等", "A1:S240"},
                    {"【４】物件税～【７】出荷経費", "A1:U96"},
                    {"減価償却費", "A1:K65"},
                    {"【８】建物", "A1:BH331"},
                    {"【９】自動車", "A1:BH326"},
                    {"【10】農業機械", "A1:BE332"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:BF113"},
                    {"【13】地代", "A1:AB38"}
                }
              } _
            , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】農家団体コード", "A1:I18"},
                    {"【２】牛１", "A1:T10015"},
                    {"【２】牛２", "A1:AL8009"},
                    {"【２】牛３", "A1:R40"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"減価償却費", "A1:J64"},
                    {"【８】建物", "A1:BH331"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE333"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:BF115"},
                    {"【13】地代", "A1:AA47"}
                }
              } _
            , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】農家団体コード", "A1:I18"},
                    {"【２】牛１", "A1:T10015"},
                    {"【２】牛２", "A1:AS8009"},
                    {"【２】牛３", "A1:R46"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"減価償却費", "A1:J64"},
                    {"【８】建物", "A1:BH331"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE333"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:BF115"},
                    {"【13】地代", "A1:AA47"}
                }
              } _
            , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】農家団体コード", "A1:I18"},
                    {"【２】牛１", "A1:T10015"},
                    {"【２】牛２", "A1:AL8009"},
                    {"【２】牛３", "A1:R40"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"減価償却費", "A1:J64"},
                    {"【８】建物", "A1:BH331"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE333"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:BF115"},
                    {"【13】地代", "A1:AA47"}
                }
              } _
            , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】農家団体コード", "A1:I18"},
                    {"【２】牛１", "A1:T10015"},
                    {"【２】牛２", "A1:AS8009"},
                    {"【２】牛３", "A1:R46"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"減価償却費", "A1:J64"},
                    {"【８】建物", "A1:BH331"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE333"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:BF115"},
                    {"【13】地代", "A1:AA47"}
                }
              } _
            , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】農家団体コード", "A1:I18"},
                    {"【２】牛１", "A1:T10015"},
                    {"【２】牛２", "A1:AL8009"},
                    {"【２】牛３", "A1:R40"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"減価償却費", "A1:J64"},
                    {"【８】建物", "A1:BH331"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE333"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:BF115"},
                    {"【13】地代", "A1:AA47"}
                }
              } _
            , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】農家団体コード", "A1:I18"},
                    {"【２】牛１", "A1:T10015"},
                    {"【２】牛２", "A1:AN8009"},
                    {"【２】牛３", "A1:R46"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"減価償却費", "A1:J64"},
                    {"【８】建物", "A1:BH331"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE333"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:BF115"},
                    {"【13】地代", "A1:AA47"}
                }
              } _
            , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】農家団体コード", "A1:I18"},
                    {"【２】牛１", "A1:T10015"},
                    {"【２】牛２", "A1:AL8009"},
                    {"【２】牛３", "A1:R40"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"減価償却費", "A1:J64"},
                    {"【８】建物", "A1:BH331"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE333"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:BF115"},
                    {"【13】地代", "A1:AA47"}
                }
              } _
            , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】農家団体コード", "A1:I18"},
                    {"【２】牛１", "A1:T10015"},
                    {"【２】牛２", "A1:AS8009"},
                    {"【２】牛３", "A1:R46"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"減価償却費", "A1:J64"},
                    {"【８】建物", "A1:BH331"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE333"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:BF115"},
                    {"【13】地代", "A1:AA47"}
                }
              } _
            , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】農家団体コード", "A1:I18"},
                    {"【２】牛１", "A1:T10015"},
                    {"【２】牛２", "A1:AL8009"},
                    {"【２】牛３", "A1:R40"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"減価償却費", "A1:J64"},
                    {"【８】建物", "A1:BH331"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE333"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:BF115"},
                    {"【13】地代", "A1:AA47"}
                }
              } _
            , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】農家団体コード", "A1:I18"},
                    {"【２】牛１", "A1:T10015"},
                    {"【２】牛２", "A1:AS8009"},
                    {"【２】牛３", "A1:R46"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２自給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"減価償却費", "A1:J64"},
                    {"【８】建物", "A1:BH331"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE333"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:BF115"},
                    {"【13】地代", "A1:AA47"}
                }
              } _
            , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】生産物", "A1:P154"},
                    {"【３】１購入飼料～８その他の資材等", "A1:U261"},
                    {"【４】物件税～【７】出荷に要した経費", "A1:Q93"},
                    {"減価償却費", "A1:I53"},
                    {"【８】建物", "A1:BA341"},
                    {"【９】自動車", "A1:AX326"},
                    {"【10】農業機械", "A1:AU330"},
                    {"【11】農具購入費", "A1:AC91"},
                    {"【12】労働時間", "A1:BG114"},
                    {"【13】地代", "A1:L40"}
                }
              } _
            , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V31"},
                    {"【２】生産物", "A1:P154"},
                    {"【３】１購入飼料～８その他の資材等", "A1:U261"},
                    {"【４】物件税～【７】出荷に要した経費", "A1:Q93"},
                    {"減価償却費", "A1:I53"},
                    {"【８】建物", "A1:BA341"},
                    {"【９】自動車", "A1:AX326"},
                    {"【10】農業機械", "A1:AU330"},
                    {"【11】農具購入費", "A1:AC91"},
                    {"【12】労働時間", "A1:BG114"},
                    {"【13】地代", "A1:L40"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_経営分析", "A1:BF63"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN49"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AI85"},
                    {"【２】生産物の販売等の状況②_二条大麦", "A1:AD83"},
                    {"【２】生産物の販売等の状況③_組織_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_経営分析", "A1:BF64"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN118"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AI85"},
                    {"【２】生産物の販売等の状況②_二条大麦", "A1:AD83"},
                    {"【２】生産物の販売等の状況③_組織_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AE93"},
                    {"【３】２肥料費_米以外", "A1:AK124"},
                    {"【３】３農業薬剤費_米以外", "A1:AC80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AU95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_経営分析", "A1:BF63"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN49"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AD85"},
                    {"【２】生産物の販売等の状況③_組織_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_経営分析", "A1:BF64"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN118"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AX85"},
                    {"【２】生産物の販売等の状況③_組織_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AI124"},
                    {"【３】３農業薬剤費_米以外", "A1:AC80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_経営分析", "A1:BF63"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN49"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AD85"},
                    {"【２】生産物の販売等の状況③_組織_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_経営分析", "A1:BF64"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN118"},
                    {"【２】生産物の販売等の状況①_麦類", "A1:AD85"},
                    {"【２】生産物の販売等の状況③_組織_麦類", "A1:AG32"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:AB80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_経営分析", "A1:BF63"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN49"},
                    {"【２】生産物の販売等の状況①_そば", "A1:AD77"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_経営分析", "A1:BF64"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN118"},
                    {"【２】生産物の販売等の状況①_そば", "A1:AD77"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:AA80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_経営分析", "A1:BF63"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_原ばれ", "A1:AN29"},
                    {"【２】生産物の販売等の状況①_原ばれ", "A1:AF57"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_経営分析", "A1:BF64"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_原ばれ", "A1:AN99"},
                    {"【２】生産物の販売等の状況①_原ばれ", "A1:AF57"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AI124"},
                    {"【３】３農業薬剤費_米以外", "A1:AB80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_経営分析", "A1:BF63"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN49"},
                    {"【２】生産物の販売等の状況①_なたね", "A1:AE59"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_経営分析", "A1:BF64"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", "A1:AN118"},
                    {"【２】生産物の販売等の状況①_なたね", "A1:AE59"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AI124"},
                    {"【３】３農業薬剤費_米以外", "A1:AB80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_経営分析", "A1:BF63"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_てんさい", "A1:AN39"},
                    {"【２】生産物の販売等の状況①_てんさい", "A1:AI58"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_経営分析", "A1:BF64"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_てんさい", "A1:AN109"},
                    {"【２】生産物の販売等の状況①_てんさい", "A1:AI58"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:AB80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_米・さとうきび以外", "A1:T243"},
                    {"【11】土地_２借入地_米・さとうきび以外", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H19"},
                    {"表紙_経営分析", "A1:BF63"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_さとうきび", "A1:AN35"},
                    {"【２】生産物の販売等の状況①_さとうきび", "A1:AI59"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:Z80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P53"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_さとうきび", "A1:T243"},
                    {"【11】土地_２借入地_さとうきび", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"指標部入力", "A1:H24"},
                    {"表紙_経営分析", "A1:BF64"},
                    {"【１】経営の概況①_組織_米以外", "A1:AH48"},
                    {"【１】経営の概況③_さとうきび", "A1:AN110"},
                    {"【２】生産物の販売等の状況①_さとうきび", "A1:AI59"},
                    {"栽培の特徴", "A1:H29"},
                    {"【３】１種苗費_米以外", "A1:AD93"},
                    {"【３】２肥料費_米以外", "A1:AH124"},
                    {"【３】３農業薬剤費_米以外", "A1:AB80"},
                    {"【３】４光熱動力費・５諸材料費_米以外", "A1:AT95"},
                    {"【３】６賃借料及び料金_米以外", "A1:BB30"},
                    {"【４】物件税及び公課諸負担_米以外", "A1:AV44"},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", "A1:AK102"},
                    {"【７】建物及び構築物", "A1:AV129"},
                    {"【８】自動車", "A1:AX129"},
                    {"【９】農業機械", "A1:AZ129"},
                    {"減価償却費", "A1:P65"},
                    {"【10】農具の購入費等", "A1:CD115"},
                    {"【11】土地_１所有地_さとうきび", "A1:T243"},
                    {"【11】土地_２借入地_さとうきび", "A1:Q241"},
                    {"【12】作業別労働時間①_組織", "A1:BC49"},
                    {"任意項目", "A1:C14"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】１生乳", "A1:AS204"},
                    {"【２】２子牛、３きゅう肥", "A1:R46"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO63"},
                    {"【３】３種付料～９その他の資材等", "A1:S246"},
                    {"【４】物件税～【６】借入金", "A1:U89"},
                    {"【７】建物", "A1:BH1031"},
                    {"【8】自動車", "A1:BH327"},
                    {"【９】農業機械", "A1:BE1031"},
                    {"減価償却費", "A1:J73"},
                    {"【10】農具購入費", "A1:AD92"},
                    {"【11】農家団体コード", "A1:E28"},
                    {"【11】搾乳牛所有状況", "A1:AI13014"},
                    {"【11】搾乳牛所有状況２", "A1:CG13009"},
                    {"飼養頭数、負担割合", "A1:Y35"},
                    {"子牛の概要", "A1:I12"},
                    {"搾乳牛の概要、資本額", "A1:L17"},
                    {"【12】労働時間", "A1:AP35"},
                    {"【13】地代・搾乳牛負担割合・内給飼料", "A1:BE64"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】１生乳", "A1:AS204"},
                    {"【２】２子牛、３きゅう肥", "A1:R46"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO57"},
                    {"【３】３種付料～９その他の資材等", "A1:S246"},
                    {"【４】物件税～【６】借入金", "A1:U89"},
                    {"【７】建物", "A1:BH1031"},
                    {"【8】自動車", "A1:BH327"},
                    {"【９】農業機械", "A1:BE1031"},
                    {"減価償却費", "A1:J73"},
                    {"【10】農具購入費", "A1:AD92"},
                    {"【11】農家団体コード", "A1:E28"},
                    {"【11】搾乳牛所有状況", "A1:AI13014"},
                    {"【11】搾乳牛所有状況２", "A1:CG13009"},
                    {"飼養頭数、負担割合", "A1:Y35"},
                    {"子牛の概要", "A1:I15"},
                    {"搾乳牛の概要、資本額", "A1:L17"},
                    {"【12】労働時間", "A1:AP35"},
                    {"【13】地代・搾乳牛負担割合・内給飼料", "A1:BE64"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】農家団体コード", "A1:E28"},
                    {"【２】２牛取引情報", "A1:AE8014"},
                    {"【２】３牛取引情報", "A1:Z2009"},
                    {"【２】対象畜概要１", "A1:EZ8018"},
                    {"【２】対象畜概要２", "A1:T49"},
                    {"【２】４きゅう肥", "A1:R34"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO57"},
                    {"【３】３種付料～９その他の資材等", "A1:S240"},
                    {"【４】物件税～【７】出荷経費", "A1:U96"},
                    {"【８】建物", "A1:BH1031"},
                    {"【９】自動車", "A1:BH327"},
                    {"【10】農業機械", "A1:BE1030"},
                    {"減価償却費", "A1:V65"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:AQ32"},
                    {"【13】地代", "A1:AC37"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】農家団体コード", "A1:E28"},
                    {"【２】２牛取引情報", "A1:AE8014"},
                    {"【２】３牛取引情報", "A1:Z2009"},
                    {"【２】対象畜概要１", "A1:EZ8018"},
                    {"【２】対象畜概要２", "A1:T49"},
                    {"【２】４きゅう肥", "A1:R34"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO57"},
                    {"【３】３種付料～９その他の資材等", "A1:S240"},
                    {"【４】物件税～【７】出荷経費", "A1:U96"},
                    {"【８】建物", "A1:BH1031"},
                    {"【９】自動車", "A1:BH327"},
                    {"【10】農業機械", "A1:BE1030"},
                    {"減価償却費", "A1:V65"},
                    {"【11】農具購入費", "A1:AD92"},
                    {"【12】労働時間", "A1:AQ32"},
                    {"【13】地代", "A1:AC37"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】農家団体コード", "A1:H27"},
                    {"【２】牛１", "A1:T30015"},
                    {"【２】牛２", "A1:AL15009"},
                    {"【２】牛３", "A1:R40"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO63"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"【８】建物", "A1:BH1031"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE1030"},
                    {"減価償却費", "A1:J64"},
                    {"【11】農具購入費", "A1:AD93"},
                    {"【12】労働時間", "A1:AQ31"},
                    {"【13】地代", "A1:AC46"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】農家団体コード", "A1:H27"},
                    {"【２】牛１", "A1:T30015"},
                    {"【２】牛２", "A1:AN15009"},
                    {"【２】牛３", "A1:R46"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"【８】建物", "A1:BH1031"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE1030"},
                    {"減価償却費", "A1:J64"},
                    {"【11】農具購入費", "A1:AD93"},
                    {"【12】労働時間", "A1:AQ31"},
                    {"【13】地代", "A1:AC46"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】農家団体コード", "A1:H27"},
                    {"【２】牛１", "A1:T30015"},
                    {"【２】牛２", "A1:AL15009"},
                    {"【２】牛３", "A1:R40"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO63"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"【８】建物", "A1:BH1031"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE1030"},
                    {"減価償却費", "A1:J64"},
                    {"【11】農具購入費", "A1:AD93"},
                    {"【12】労働時間", "A1:AQ31"},
                    {"【13】地代", "A1:AC46"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】農家団体コード", "A1:H27"},
                    {"【２】牛１", "A1:T30015"},
                    {"【２】牛２", "A1:AS15009"},
                    {"【２】牛３", "A1:R46"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"【８】建物", "A1:BH1031"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE1030"},
                    {"減価償却費", "A1:J64"},
                    {"【11】農具購入費", "A1:AD93"},
                    {"【12】労働時間", "A1:AQ31"},
                    {"【13】地代", "A1:AC46"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】農家団体コード", "A1:H27"},
                    {"【２】牛１", "A1:T30015"},
                    {"【２】牛２", "A1:AL15009"},
                    {"【２】牛３", "A1:R40"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO63"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"【８】建物", "A1:BH1031"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE1030"},
                    {"減価償却費", "A1:J64"},
                    {"【11】農具購入費", "A1:AD93"},
                    {"【12】労働時間", "A1:AQ31"},
                    {"【13】地代", "A1:AC46"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】農家団体コード", "A1:H27"},
                    {"【２】牛１", "A1:T30015"},
                    {"【２】牛２", "A1:AS15009"},
                    {"【２】牛３", "A1:R46"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"【８】建物", "A1:BH1031"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE1030"},
                    {"減価償却費", "A1:J64"},
                    {"【11】農具購入費", "A1:AD93"},
                    {"【12】労働時間", "A1:AQ31"},
                    {"【13】地代", "A1:AC46"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】農家団体コード", "A1:H27"},
                    {"【２】牛１", "A1:T30015"},
                    {"【２】牛２", "A1:AL15009"},
                    {"【２】牛３", "A1:R40"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO63"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"【８】建物", "A1:BH1031"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE1030"},
                    {"減価償却費", "A1:J64"},
                    {"【11】農具購入費", "A1:AD93"},
                    {"【12】労働時間", "A1:AQ31"},
                    {"【13】地代", "A1:AC46"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】農家団体コード", "A1:H27"},
                    {"【２】牛１", "A1:T30015"},
                    {"【２】牛２", "A1:AS15009"},
                    {"【２】牛３", "A1:R46"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"【８】建物", "A1:BH1031"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE1030"},
                    {"減価償却費", "A1:J64"},
                    {"【11】農具購入費", "A1:AD93"},
                    {"【12】労働時間", "A1:AQ31"},
                    {"【13】地代", "A1:AC46"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】農家団体コード", "A1:H27"},
                    {"【２】牛１", "A1:T30015"},
                    {"【２】牛２", "A1:AL15009"},
                    {"【２】牛３", "A1:R40"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO63"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"【８】建物", "A1:BH1031"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE1030"},
                    {"減価償却費", "A1:J64"},
                    {"【11】農具購入費", "A1:AD93"},
                    {"【12】労働時間", "A1:AQ31"},
                    {"【13】地代", "A1:AC46"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】農家団体コード", "A1:H27"},
                    {"【２】牛１", "A1:V30015"},
                    {"【２】牛２", "A1:AS15009"},
                    {"【２】牛３", "A1:R46"},
                    {"【２】３きゅう肥", "A1:Q35"},
                    {"【３】１購入飼料", "A1:U38"},
                    {"【３】２内給飼料", "A1:AO57"},
                    {"【３】３敷料費～８その他の資材等", "A1:S231"},
                    {"【４】物件税～【７】出荷経費", "A1:U98"},
                    {"【８】建物", "A1:BH1031"},
                    {"【９】自動車", "A1:BH325"},
                    {"【10】農業機械", "A1:BE1030"},
                    {"減価償却費", "A1:J64"},
                    {"【11】農具購入費", "A1:AD93"},
                    {"【12】労働時間", "A1:AQ31"},
                    {"【13】地代", "A1:AC46"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.調査票項目2015), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】生産物", "A1:P154"},
                    {"【３】１購入飼料～８その他の資材等", "A1:U261"},
                    {"【４】物件税～【７】出荷に要した経費", "A1:Q93"},
                    {"【８】建物", "A1:BA1031"},
                    {"【９】自動車", "A1:AX326"},
                    {"【10】農業機械", "A1:AU1030"},
                    {"減価償却費", "A1:I53"},
                    {"【11】農具購入費", "A1:AC92"},
                    {"【12】労働時間", "A1:AQ32"},
                    {"【13】地代", "A1:M39"}
                }
              } _
            , {Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.調査票項目2020), New Dictionary(Of String, String) From {
                    {"表紙", "A1:BS67"},
                    {"【１】経営概要", "A1:V34"},
                    {"【２】生産物", "A1:P154"},
                    {"【３】１購入飼料～８その他の資材等", "A1:U261"},
                    {"【４】物件税～【７】出荷に要した経費", "A1:Q93"},
                    {"【８】建物", "A1:BA1031"},
                    {"【９】自動車", "A1:AX326"},
                    {"【10】農業機械", "A1:AU1030"},
                    {"減価償却費", "A1:I53"},
                    {"【11】農具購入費", "A1:AC92"},
                    {"【12】労働時間", "A1:AQ32"},
                    {"【13】地代", "A1:M39"}
                }
            }
        }
        '--- REV.003 ADD START
        Public Enum 非表示設定
            なし
            非表示
            折りたたみ
        End Enum
        ''' <summary>シート非表示設定</summary>
        Public Shared シート非表示設定 As New Dictionary(Of String, Dictionary(Of String, Integer)) From {
       {調査区分.米生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_個別", 非表示設定.なし},
                    {"約束事項_米_個別", 非表示設定.なし},
                    {"目次_個別_米", 非表示設定.なし},
                    {"【１】経営の概況①_個別", 非表示設定.折りたたみ},
                    {"【１】経営の概況②_個別_米", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_個別_米", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_米", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況②_米", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_個別", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間②_米", 非表示設定.折りたたみ},
                    {"【13】飼料用米①_米", 非表示設定.折りたたみ},
                    {"【13】飼料用米②_米", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.小麦生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_個別", 非表示設定.なし},
                    {"約束事項_麦類", 非表示設定.なし},
                    {"目次_個別_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_個別", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_個別_麦・大豆", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_麦類", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況③_個別_麦類", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_個別", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.二条大麦生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_個別", 非表示設定.なし},
                    {"約束事項_麦類", 非表示設定.なし},
                    {"目次_個別_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_個別", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_個別_麦・大豆", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_麦類", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況②_二条大麦", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況③_個別_麦類", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_個別", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                 }
              } _
            , {調査区分.六条大麦生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_個別", 非表示設定.なし},
                    {"約束事項_麦類", 非表示設定.なし},
                    {"目次_個別_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_個別", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_個別_麦・大豆", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_麦類", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況③_個別_麦類", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_個別", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.はだか麦生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_個別", 非表示設定.なし},
                    {"約束事項_麦類", 非表示設定.なし},
                    {"目次_個別_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_個別", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_個別_麦・大豆", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_麦類", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況③_個別_麦類", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_個別", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.そば生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_個別", 非表示設定.なし},
                    {"約束事項_そば", 非表示設定.なし},
                    {"目次_個別_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_個別", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_個別_なたね・そば", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_そば", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_個別", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                  }
              } _
            , {調査区分.大豆生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_個別", 非表示設定.なし},
                    {"約束事項_大豆", 非表示設定.なし},
                    {"目次_個別_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_個別", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_個別_麦・大豆", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_大豆", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_個別", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.原料用かんしょ生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_個別", 非表示設定.なし},
                    {"約束事項_原かん", 非表示設定.なし},
                    {"目次_個別_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_個別", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_原かん", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_原かん", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_個別", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.原料用ばれいしょ生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_個別", 非表示設定.なし},
                    {"約束事項_原ばれ", 非表示設定.なし},
                    {"目次_個別_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_個別", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_原ばれ", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_原ばれ", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_個別", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.なたね生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_個別", 非表示設定.なし},
                    {"約束事項_なたね", 非表示設定.なし},
                    {"目次_個別_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_個別", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_個別_なたね・そば", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_なたね", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_個別", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.てんさい生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_個別", 非表示設定.なし},
                    {"約束事項_てんさい", 非表示設定.なし},
                    {"目次_個別_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_個別", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_てんさい", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_てんさい", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_個別", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.さとうきび生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_個別", 非表示設定.なし},
                    {"約束事項_さとうきび", 非表示設定.なし},
                    {"目次_個別_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_個別", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_さとうきび", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_さとうきび", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_さとうきび", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_さとうきび", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_個別", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.米生産費統計_組織法人, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_組織", 非表示設定.なし},
                    {"約束事項_米_組織", 非表示設定.なし},
                    {"目次_組織_米", 非表示設定.なし},
                    {"【１】経営の概況①_組織_米", 非表示設定.折りたたみ},
                    {"【１】経営の概況②_組織_米", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_組織_米", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_米", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況②_米", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_組織", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間②_米", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.小麦生産費統計_組織法人, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_組織", 非表示設定.なし},
                    {"約束事項_麦類", 非表示設定.なし},
                    {"目次_組織_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_組織_米以外", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_麦類", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況③_組織_麦類", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_組織", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.大豆生産費統計_組織法人, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_組織", 非表示設定.なし},
                    {"約束事項_大豆", 非表示設定.なし},
                    {"目次_組織_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_組織_米以外", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_大豆", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_組織", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.牛乳生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】１生乳", 非表示設定.折りたたみ},
                    {"【２】２子牛、３きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２自給飼料", 非表示設定.折りたたみ},
                    {"【３】３種付料～９その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【６】借入金", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【７】建物", 非表示設定.折りたたみ},
                    {"【8】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"【10】農具購入費", 非表示設定.折りたたみ},
                    {"【11】農家団体コード", 非表示設定.なし},
                    {"【11】搾乳牛所有状況", 非表示設定.折りたたみ},
                    {"【11】搾乳牛所有状況２", 非表示設定.非表示},
                    {"飼養頭数、負担割合", 非表示設定.非表示},
                    {"子牛の概要", 非表示設定.非表示},
                    {"搾乳牛の概要、資本額", 非表示設定.非表示},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代・搾乳牛負担割合・自給牧草", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.子牛生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】農家団体コード", 非表示設定.なし},
                    {"【２】2牛取引情報", 非表示設定.折りたたみ},
                    {"【２】３牛取引情報", 非表示設定.なし},
                    {"【２】対象畜概要１", 非表示設定.非表示},
                    {"【２】対象畜概要２", 非表示設定.非表示},
                    {"【２】４きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２自給飼料", 非表示設定.折りたたみ},
                    {"【３】３種付料～９その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷経費", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.乳用雄育成牛生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】農家団体コード", 非表示設定.なし},
                    {"【２】牛１", 非表示設定.折りたたみ},
                    {"【２】牛２", 非表示設定.非表示},
                    {"【２】牛３", 非表示設定.非表示},
                    {"【２】３きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２自給飼料", 非表示設定.折りたたみ},
                    {"【３】３敷料費～８その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷経費", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.交雑種育成牛生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】農家団体コード", 非表示設定.なし},
                    {"【２】牛１", 非表示設定.折りたたみ},
                    {"【２】牛２", 非表示設定.非表示},
                    {"【２】牛３", 非表示設定.非表示},
                    {"【２】３きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２自給飼料", 非表示設定.折りたたみ},
                    {"【３】３敷料費～８その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷経費", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】農家団体コード", 非表示設定.なし},
                    {"【２】牛１", 非表示設定.折りたたみ},
                    {"【２】牛２", 非表示設定.非表示},
                    {"【２】牛３", 非表示設定.非表示},
                    {"【２】３きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２自給飼料", 非表示設定.折りたたみ},
                    {"【３】３敷料費～８その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷経費", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.乳用雄肥育牛生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】農家団体コード", 非表示設定.なし},
                    {"【２】牛１", 非表示設定.折りたたみ},
                    {"【２】牛２", 非表示設定.非表示},
                    {"【２】牛３", 非表示設定.非表示},
                    {"【２】３きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２自給飼料", 非表示設定.折りたたみ},
                    {"【３】３敷料費～８その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷経費", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.交雑種肥育牛生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】農家団体コード", 非表示設定.なし},
                    {"【２】牛１", 非表示設定.折りたたみ},
                    {"【２】牛２", 非表示設定.非表示},
                    {"【２】牛３", 非表示設定.非表示},
                    {"【２】３きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２自給飼料", 非表示設定.折りたたみ},
                    {"【３】３敷料費～８その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷経費", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.肥育豚生産費統計_個別, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】生産物", 非表示設定.折りたたみ},
                    {"【３】１購入飼料～８その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷に要した経費", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_二条大麦生産費, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_経営分析", 非表示設定.なし},
                    {"約束事項_麦類", 非表示設定.なし},
                    {"目次_組織_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_組織_米以外", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_麦類", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況②_二条大麦", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況③_組織_麦類", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_組織", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_六条大麦生産費, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_経営分析", 非表示設定.なし},
                    {"約束事項_麦類", 非表示設定.なし},
                    {"目次_組織_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_組織_米以外", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_麦類", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況③_組織_麦類", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_組織", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_はだか麦生産費, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_経営分析", 非表示設定.なし},
                    {"約束事項_麦類", 非表示設定.なし},
                    {"目次_組織_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_組織_米以外", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_麦類", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況③_組織_麦類", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_組織", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_そば生産費, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_経営分析", 非表示設定.なし},
                    {"約束事項_そば", 非表示設定.なし},
                    {"目次_組織_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_組織_米以外", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_そば", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_組織", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_原料用ばれいしょ生産費, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_経営分析", 非表示設定.なし},
                    {"約束事項_原ばれ", 非表示設定.なし},
                    {"目次_組織_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_組織_米以外", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_原ばれ", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_原ばれ", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_組織", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_なたね生産費, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_経営分析", 非表示設定.なし},
                    {"約束事項_なたね", 非表示設定.なし},
                    {"目次_組織_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_組織_米以外", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_組織_麦・大豆・なたね・そば", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_なたね", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_組織", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_てんさい生産費, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_経営分析", 非表示設定.なし},
                    {"約束事項_てんさい", 非表示設定.なし},
                    {"目次_組織_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_組織_米以外", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_てんさい", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_てんさい", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_米・さとうきび以外", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_組織", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_さとうきび生産費, New Dictionary(Of String, Integer) From {
                    {"指標部入力", 非表示設定.非表示},
                    {"表紙_経営分析", 非表示設定.なし},
                    {"約束事項_さとうきび", 非表示設定.なし},
                    {"目次_組織_米以外", 非表示設定.なし},
                    {"【１】経営の概況①_組織_米以外", 非表示設定.折りたたみ},
                    {"【１】経営の概況③_さとうきび", 非表示設定.折りたたみ},
                    {"【２】生産物の販売等の状況①_さとうきび", 非表示設定.折りたたみ},
                    {"栽培の特徴", 非表示設定.なし},
                    {"【３】１種苗費_米以外", 非表示設定.折りたたみ},
                    {"【３】２肥料費_米以外", 非表示設定.折りたたみ},
                    {"【３】３農業薬剤費_米以外", 非表示設定.折りたたみ},
                    {"【３】４光熱動力費・５諸材料費_米以外", 非表示設定.折りたたみ},
                    {"【３】６賃借料及び料金_米以外", 非表示設定.折りたたみ},
                    {"【４】物件税及び公課諸負担_米以外", 非表示設定.折りたたみ},
                    {"【５】土地改良及び水利費・【６】借入金_米以外", 非表示設定.折りたたみ},
                    {"【７】建物及び構築物", 非表示設定.折りたたみ},
                    {"【８】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具の購入費等", 非表示設定.折りたたみ},
                    {"【11】土地_１所有地_さとうきび", 非表示設定.折りたたみ},
                    {"【11】土地_２借入地_さとうきび", 非表示設定.折りたたみ},
                    {"【12】作業別労働時間①_組織", 非表示設定.折りたたみ},
                    {"任意項目", 非表示設定.非表示},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_牛乳生産費, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】１生乳", 非表示設定.折りたたみ},
                    {"【２】２子牛、３きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２内給飼料", 非表示設定.折りたたみ},
                    {"【３】３種付料～９その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【６】借入金", 非表示設定.折りたたみ},
                    {"【７】建物", 非表示設定.折りたたみ},
                    {"【8】自動車", 非表示設定.折りたたみ},
                    {"【９】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【10】農具購入費", 非表示設定.折りたたみ},
                    {"【11】農家団体コード", 非表示設定.なし},
                    {"【11】搾乳牛所有状況", 非表示設定.折りたたみ},
                    {"【11】搾乳牛所有状況２", 非表示設定.非表示},
                    {"飼養頭数、負担割合", 非表示設定.非表示},
                    {"子牛の概要", 非表示設定.非表示},
                    {"搾乳牛の概要、資本額", 非表示設定.非表示},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代・搾乳牛負担割合・内給飼料", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_子牛生産費, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】農家団体コード", 非表示設定.なし},
                    {"【２】２牛取引情報", 非表示設定.折りたたみ},
                    {"【２】３牛取引情報", 非表示設定.折りたたみ},
                    {"【２】対象畜概要１", 非表示設定.非表示},
                    {"【２】対象畜概要２", 非表示設定.非表示},
                    {"【２】４きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２内給飼料", 非表示設定.折りたたみ},
                    {"【３】３種付料～９その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷経費", 非表示設定.折りたたみ},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】農家団体コード", 非表示設定.なし},
                    {"【２】牛１", 非表示設定.折りたたみ},
                    {"【２】牛２", 非表示設定.非表示},
                    {"【２】牛３", 非表示設定.非表示},
                    {"【２】３きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２内給飼料", 非表示設定.折りたたみ},
                    {"【３】３敷料費～８その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷経費", 非表示設定.折りたたみ},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_交雑種育成牛生産費, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】農家団体コード", 非表示設定.なし},
                    {"【２】牛１", 非表示設定.折りたたみ},
                    {"【２】牛２", 非表示設定.非表示},
                    {"【２】牛３", 非表示設定.非表示},
                    {"【２】３きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２内給飼料", 非表示設定.折りたたみ},
                    {"【３】３敷料費～８その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷経費", 非表示設定.折りたたみ},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】農家団体コード", 非表示設定.なし},
                    {"【２】牛１", 非表示設定.折りたたみ},
                    {"【２】牛２", 非表示設定.非表示},
                    {"【２】牛３", 非表示設定.非表示},
                    {"【２】３きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２内給飼料", 非表示設定.折りたたみ},
                    {"【３】３敷料費～８その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷経費", 非表示設定.折りたたみ},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】農家団体コード", 非表示設定.なし},
                    {"【２】牛１", 非表示設定.折りたたみ},
                    {"【２】牛２", 非表示設定.非表示},
                    {"【２】牛３", 非表示設定.非表示},
                    {"【２】３きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２内給飼料", 非表示設定.折りたたみ},
                    {"【３】３敷料費～８その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷経費", 非表示設定.折りたたみ},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】農家団体コード", 非表示設定.なし},
                    {"【２】牛１", 非表示設定.折りたたみ},
                    {"【２】牛２", 非表示設定.非表示},
                    {"【２】牛３", 非表示設定.非表示},
                    {"【２】３きゅう肥", 非表示設定.折りたたみ},
                    {"【３】１購入飼料", 非表示設定.折りたたみ},
                    {"【３】２内給飼料", 非表示設定.折りたたみ},
                    {"【３】３敷料費～８その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷経費", 非表示設定.折りたたみ},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              } _
            , {調査区分.経営分析調査_肥育豚生産費, New Dictionary(Of String, Integer) From {
                    {"表紙", 非表示設定.折りたたみ},
                    {"説明", 非表示設定.なし},
                    {"【１】経営概要", 非表示設定.折りたたみ},
                    {"【２】生産物", 非表示設定.折りたたみ},
                    {"【３】１購入飼料～８その他の資材等", 非表示設定.折りたたみ},
                    {"【４】物件税～【７】出荷に要した経費", 非表示設定.折りたたみ},
                    {"【８】建物", 非表示設定.折りたたみ},
                    {"【９】自動車", 非表示設定.折りたたみ},
                    {"【10】農業機械", 非表示設定.折りたたみ},
                    {"減価償却費", 非表示設定.非表示},
                    {"【11】農具購入費", 非表示設定.折りたたみ},
                    {"【12】労働時間", 非表示設定.折りたたみ},
                    {"【13】地代", 非表示設定.折りたたみ},
                    {"リスト", 非表示設定.非表示}
                }
              }
        }

        ''' <summary>
        ''' 総括データのシート名
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared 牛総括データシート As New Dictionary(Of String, String) From {
            {調査区分.牛乳生産費統計_個別, "【11】搾乳牛所有状況２"} _
          , {調査区分.子牛生産費統計_個別, "【２】対象畜概要１"} _
          , {調査区分.乳用雄育成牛生産費統計_個別, "【２】牛２"} _
          , {調査区分.交雑種育成牛生産費統計_個別, "【２】牛２"} _
          , {調査区分.去勢若齢肥育牛生産費統計_個別, "【２】牛２"} _
          , {調査区分.乳用雄肥育牛生産費統計_個別, "【２】牛２"} _
          , {調査区分.交雑種肥育牛生産費統計_個別, "【２】牛２"} _
          , {調査区分.経営分析調査_牛乳生産費, "【11】搾乳牛所有状況２"} _
          , {調査区分.経営分析調査_子牛生産費, "【２】対象畜概要１"} _
          , {調査区分.経営分析調査_乳用雄育成牛生産費, "【２】牛２"} _
          , {調査区分.経営分析調査_交雑種育成牛生産費, "【２】牛２"} _
          , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "【２】牛２"} _
          , {調査区分.経営分析調査_乳用雄肥育牛生産費, "【２】牛２"} _
          , {調査区分.経営分析調査_交雑種肥育牛生産費, "【２】牛２"}
        }
        '--- REV.003 ADD END
    End Class

    ''' <summary>
    ''' 調査票エラーチェック一覧クラス
    ''' </summary>
    ''' <remarks>
    ''' 'REV004 Add
    ''' </remarks>
    Public Class 調査票エラーチェックリスト一覧
        Public Const delimiter As String = "/"
        Public Enum enm
            エラーシート名 = 1
            項目番号 = 7
            可変行番号 = 10
            エラーサイン = 13
            エラー内容 = 14
            エラーチェック種別 = 18
        End Enum

        Public Class 出力用ファイル名称
            Public Class Report
                ''' <summary>テンプレートファイル名</summary>
                Public Shared tempFileName As String = "調査票エラーチェックリスト.xlsx"

                ''' <summary>帳票名</summary>
                Public Const reportName As String = "調査票エラーチェックリスト"

                ''' <summary>シート名</summary>
                Public Const SheetName As String = "調査票エラーチェックリスト"

            End Class

            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 10
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 1004
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = Last - First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "T"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                      {1, "エラーシート名"} _
                    , {7, "項目番号"} _
                    , {10, "可変行番号"} _
                    , {13, "エラーサイン"} _
                    , {14, "エラー内容"} _
                    , {18, "エラーチェック種別"}
                }
        End Class
    End Class

    ''' <summary>
    ''' 調査票エラーチェックリスト一覧複数客体クラス
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    Public Class 調査票エラーチェックリスト一覧複数客体
        Public Const delimiter As String = "/"
        Public Enum enm
            No = 1
            都道府県
            市区町村
            旧市区町村
            農業集落
            調査区
            客体番号
            エラーシート名
            項目番号
            可変行番号
            エラーサイン
            エラー内容
            エラーチェック種別
        End Enum

        Public Class 出力用ファイル名称
            Public Class Report
                ''' <summary>テンプレートファイル名</summary>
                Public Shared tempFileName As String = "調査票エラーチェックリスト_複数客体.xlsx"

                ''' <summary>帳票名</summary>
                Public Shared reportName As New Dictionary(Of エラーチェック種別.enm, String) From {
                      {エラーチェック種別.enm.基本, "調査票基本エラーチェックリスト"} _
                    , {エラーチェック種別.enm.追加, "調査票追加エラーチェックリスト"} _
                    , {エラーチェック種別.enm.範囲, "調査票範囲エラーチェックリスト"}
                }

                ''' <summary>シート名</summary>
                Public Const SheetName As String = "チェックリスト"
                ''' <summary>タイトル欄</summary>
                Public Const Title As String = "B1"
                ''' <summary>調査年欄</summary>
                Public Const ChosaNen As String = "B4"
                ''' <summary>調査区分欄</summary>
                Public Const Chosakubun As String = "C4"
                ''' <summary>対象シート欄</summary>
                Public Const Sheet As String = "I4"
                ''' <summary>拠点欄</summary>
                Public Const Kyoten As String = "O6"

            End Class

            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 9
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 10008
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = Last - First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "O"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                      {1, "No"} _
                    , {2, "都道府県"} _
                    , {3, "市区町村"} _
                    , {4, "旧市区町村"} _
                    , {5, "農業集落"} _
                    , {6, "調査区"} _
                    , {7, "客体番号"} _
                    , {8, "エラーシート名"} _
                    , {9, "項目番号"} _
                    , {10, "可変行番号"} _
                    , {11, "エラーサイン"} _
                    , {12, "エラー内容"} _
                    , {13, "エラーチェック種別"}
                }
        End Class
    End Class

    ''' <summary>
    ''' 個別結果表クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 個別結果表
        ''' <summary>テーブル名称</summary>
        Public Shared テーブル名称 As New Dictionary(Of String, String()) From {
              {調査区分.営農類型別経営統計_個人, {"個別結果表＿農業経営＿営農類型＿個人１", "個別結果表＿農業経営＿営農類型＿個人２" _
                                                , "個別結果表＿農業経営＿営農類型＿個人３", "個別結果表＿農業経営＿営農類型＿個人４" _
                                                , "個別結果表＿農業経営＿営農類型＿個人５", "個別結果表＿農業経営＿営農類型＿個人６"}} _
            , {調査区分.営農類型別経営統計_法人, {"個別結果表＿農業経営＿営農類型＿法人１", "個別結果表＿農業経営＿営農類型＿法人２" _
                                                , "個別結果表＿農業経営＿営農類型＿法人３", "個別結果表＿農業経営＿営農類型＿法人４" _
                                                , "個別結果表＿農業経営＿営農類型＿法人５", "個別結果表＿農業経営＿営農類型＿法人６" _
                                                , "個別結果表＿農業経営＿営農類型＿法人７"}} _
            , {調査区分.米生産費統計_個別, {"個別結果表＿農業経営＿米＿個別１", "個別結果表＿農業経営＿米＿個別２" _
                                          , "個別結果表＿農業経営＿米＿個別３"}} _
            , {調査区分.小麦生産費統計_個別, {"個別結果表＿農業経営＿小麦＿個別１", "個別結果表＿農業経営＿小麦＿個別２" _
                                            , "個別結果表＿農業経営＿小麦＿個別３"}} _
            , {調査区分.二条大麦生産費統計_個別, {"個別結果表＿農業経営＿二条大麦＿個別１", "個別結果表＿農業経営＿二条大麦＿個別２" _
                                                , "個別結果表＿農業経営＿二条大麦＿個別３"}} _
            , {調査区分.六条大麦生産費統計_個別, {"個別結果表＿農業経営＿六条大麦＿個別１", "個別結果表＿農業経営＿六条大麦＿個別２" _
                                                , "個別結果表＿農業経営＿六条大麦＿個別３"}} _
            , {調査区分.はだか麦生産費統計_個別, {"個別結果表＿農業経営＿はだか麦＿個別１", "個別結果表＿農業経営＿はだか麦＿個別２" _
                                                , "個別結果表＿農業経営＿はだか麦＿個別３"}} _
            , {調査区分.そば生産費統計_個別, {"個別結果表＿農業経営＿そば＿個別１", "個別結果表＿農業経営＿そば＿個別２" _
                                            , "個別結果表＿農業経営＿そば＿個別３"}} _
            , {調査区分.大豆生産費統計_個別, {"個別結果表＿農業経営＿大豆＿個別１", "個別結果表＿農業経営＿大豆＿個別２" _
                                            , "個別結果表＿農業経営＿大豆＿個別３"}} _
            , {調査区分.原料用かんしょ生産費統計_個別, {"個別結果表＿農業経営＿原料用かんしょ＿個別１", "個別結果表＿農業経営＿原料用かんしょ＿個別２" _
                                                      , "個別結果表＿農業経営＿原料用かんしょ＿個別３"}} _
            , {調査区分.原料用ばれいしょ生産費統計_個別, {"個別結果表＿農業経営＿原料用ばれいしょ＿個別１", "個別結果表＿農業経営＿原料用ばれいしょ＿個別２" _
                                                        , "個別結果表＿農業経営＿原料用ばれいしょ＿個別３"}} _
            , {調査区分.なたね生産費統計_個別, {"個別結果表＿農業経営＿なたね＿個別１", "個別結果表＿農業経営＿なたね＿個別２" _
                                              , "個別結果表＿農業経営＿なたね＿個別３"}} _
            , {調査区分.てんさい生産費統計_個別, {"個別結果表＿農業経営＿てんさい＿個別１", "個別結果表＿農業経営＿てんさい＿個別２" _
                                                , "個別結果表＿農業経営＿てんさい＿個別３"}} _
            , {調査区分.さとうきび生産費統計_個別, {"個別結果表＿農業経営＿さとうきび＿個別１", "個別結果表＿農業経営＿さとうきび＿個別２" _
                                                  , "個別結果表＿農業経営＿さとうきび＿個別３"}} _
            , {調査区分.米生産費統計_組織法人, {"個別結果表＿農業経営＿米＿組織１", "個別結果表＿農業経営＿米＿組織２" _
                                              , "個別結果表＿農業経営＿米＿組織３"}} _
            , {調査区分.小麦生産費統計_組織法人, {"個別結果表＿農業経営＿小麦＿組織１", "個別結果表＿農業経営＿小麦＿組織２"}} _
            , {調査区分.大豆生産費統計_組織法人, {"個別結果表＿農業経営＿大豆＿組織１", "個別結果表＿農業経営＿大豆＿組織２"}} _
            , {調査区分.牛乳生産費統計_個別, {"個別結果表＿農業経営＿牛乳＿個別１", "個別結果表＿農業経営＿牛乳＿個別２"}} _
            , {調査区分.子牛生産費統計_個別, {"個別結果表＿農業経営＿子牛＿個別１", "個別結果表＿農業経営＿子牛＿個別２"}} _
            , {調査区分.乳用雄育成牛生産費統計_個別, {"個別結果表＿農業経営＿乳用雄育成牛＿個別１", "個別結果表＿農業経営＿乳用雄育成牛＿個別２"}} _
            , {調査区分.交雑種育成牛生産費統計_個別, {"個別結果表＿農業経営＿交雑種育成牛＿個別１", "個別結果表＿農業経営＿交雑種育成牛＿個別２"}} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, {"個別結果表＿農業経営＿去勢若齢肥育牛＿個別１", "個別結果表＿農業経営＿去勢若齢肥育牛＿個別２"}} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, {"個別結果表＿農業経営＿乳用雄肥育牛＿個別１", "個別結果表＿農業経営＿乳用雄肥育牛＿個別２"}} _
            , {調査区分.交雑種肥育牛生産費統計_個別, {"個別結果表＿農業経営＿交雑種肥育牛＿個別１", "個別結果表＿農業経営＿交雑種肥育牛＿個別２"}} _
            , {調査区分.肥育豚生産費統計_個別, {"個別結果表＿農業経営＿肥育豚＿個別１", "個別結果表＿農業経営＿肥育豚＿個別２"}} _
            , {調査区分.経営分析調査_二条大麦生産費, {"個別結果表＿経営分析＿二条大麦１", "個別結果表＿経営分析＿二条大麦２" _
                                                    , "個別結果表＿経営分析＿二条大麦３"}} _
            , {調査区分.経営分析調査_六条大麦生産費, {"個別結果表＿経営分析＿六条大麦１", "個別結果表＿経営分析＿六条大麦２" _
                                                    , "個別結果表＿経営分析＿六条大麦３"}} _
            , {調査区分.経営分析調査_はだか麦生産費, {"個別結果表＿経営分析＿はだか麦１", "個別結果表＿経営分析＿はだか麦２" _
                                                    , "個別結果表＿経営分析＿はだか麦３"}} _
            , {調査区分.経営分析調査_そば生産費, {"個別結果表＿経営分析＿そば１", "個別結果表＿経営分析＿そば２" _
                                                , "個別結果表＿経営分析＿そば３"}} _
            , {調査区分.経営分析調査_原料用ばれいしょ生産費, {"個別結果表＿経営分析＿原料用ばれいしょ１", "個別結果表＿経営分析＿原料用ばれいしょ２" _
                                                            , "個別結果表＿経営分析＿原料用ばれいしょ３"}} _
            , {調査区分.経営分析調査_なたね生産費, {"個別結果表＿経営分析＿なたね１", "個別結果表＿経営分析＿なたね２" _
                                                  , "個別結果表＿経営分析＿なたね３"}} _
            , {調査区分.経営分析調査_てんさい生産費, {"個別結果表＿経営分析＿てんさい１", "個別結果表＿経営分析＿てんさい２" _
                                                    , "個別結果表＿経営分析＿てんさい３"}} _
            , {調査区分.経営分析調査_さとうきび生産費, {"個別結果表＿経営分析＿さとうきび１", "個別結果表＿経営分析＿さとうきび２" _
                                                      , "個別結果表＿経営分析＿さとうきび３"}} _
            , {調査区分.経営分析調査_牛乳生産費, {"個別結果表＿経営分析＿牛乳１", "個別結果表＿経営分析＿牛乳２"}} _
            , {調査区分.経営分析調査_子牛生産費, {"個別結果表＿経営分析＿子牛１", "個別結果表＿経営分析＿子牛２"}} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, {"個別結果表＿経営分析＿乳用雄育成牛１", "個別結果表＿経営分析＿乳用雄育成牛２"}} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, {"個別結果表＿経営分析＿交雑種育成牛１", "個別結果表＿経営分析＿交雑種育成牛２"}} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, {"個別結果表＿経営分析＿去勢若齢肥育牛１", "個別結果表＿経営分析＿去勢若齢肥育牛２"}} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, {"個別結果表＿経営分析＿乳用雄肥育牛１", "個別結果表＿経営分析＿乳用雄肥育牛２"}} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, {"個別結果表＿経営分析＿交雑種肥育牛１", "個別結果表＿経営分析＿交雑種肥育牛２"}} _
            , {調査区分.経営分析調査_肥育豚生産費, {"個別結果表＿経営分析＿肥育豚１", "個別結果表＿経営分析＿肥育豚２"}} _
            , {営農経営体区分.農業経営体, {"個別結果表＿農業経営＿営農類型＿個人４", "個別結果表＿農業経営＿営農類型＿個人５", "個別結果表＿農業経営＿営農類型＿個人６" _
                                         , "個別結果表＿農業経営＿営農類型＿法人４", "個別結果表＿農業経営＿営農類型＿法人５", "個別結果表＿農業経営＿営農類型＿法人６"}}
        }

        ''' <summary>都道府県</summary>
        Public Shared 都道府県 As String = "K000002"

        ''' <summary>営農類型</summary>
        Public Shared 営農類型 As New Dictionary(Of String, String) From {
              {調査区分.営農類型別経営統計_個人, "K000005"} _
            , {調査区分.営農類型別経営統計_法人, "K000006"}
        }

        ''' <summary>貸借対照表</summary>
        Public Shared 貸借対照表 As New Dictionary(Of String, String) From {
              {調査区分.営農類型別経営統計_個人, "K000018"}
        }

        ''' <summary>前回センサス番号</summary>
        Public Shared 前回センサス番号 As New Dictionary(Of String, String) From {
              {調査区分.営農類型別経営統計_個人, "K000004"} _
            , {調査区分.営農類型別経営統計_法人, "K000004"} _
            , {調査区分.米生産費統計_個別, "K000011"} _
            , {調査区分.小麦生産費統計_個別, "K000014"} _
            , {調査区分.二条大麦生産費統計_個別, "K000014"} _
            , {調査区分.六条大麦生産費統計_個別, "K000014"} _
            , {調査区分.はだか麦生産費統計_個別, "K000014"} _
            , {調査区分.そば生産費統計_個別, "K000014"} _
            , {調査区分.大豆生産費統計_個別, "K000014"} _
            , {調査区分.原料用かんしょ生産費統計_個別, "K000014"} _
            , {調査区分.原料用ばれいしょ生産費統計_個別, "K000014"} _
            , {調査区分.なたね生産費統計_個別, "K000014"} _
            , {調査区分.てんさい生産費統計_個別, "K000014"} _
            , {調査区分.さとうきび生産費統計_個別, "K000014"} _
            , {調査区分.米生産費統計_組織法人, "K000015"} _
            , {調査区分.小麦生産費統計_組織法人, "K000015"} _
            , {調査区分.大豆生産費統計_組織法人, "K000015"} _
            , {調査区分.牛乳生産費統計_個別, "K000013"} _
            , {調査区分.子牛生産費統計_個別, "K000011"} _
            , {調査区分.乳用雄育成牛生産費統計_個別, "K000010"} _
            , {調査区分.交雑種育成牛生産費統計_個別, "K000010"} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, "K000010"} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, "K000010"} _
            , {調査区分.交雑種肥育牛生産費統計_個別, "K000010"} _
            , {調査区分.肥育豚生産費統計_個別, "K000014"} _
            , {調査区分.経営分析調査_二条大麦生産費, "K000016"} _
            , {調査区分.経営分析調査_六条大麦生産費, "K000016"} _
            , {調査区分.経営分析調査_はだか麦生産費, "K000016"} _
            , {調査区分.経営分析調査_そば生産費, "K000016"} _
            , {調査区分.経営分析調査_原料用ばれいしょ生産費, "K000016"} _
            , {調査区分.経営分析調査_なたね生産費, "K000016"} _
            , {調査区分.経営分析調査_てんさい生産費, "K000016"} _
            , {調査区分.経営分析調査_さとうきび生産費, "K000016"} _
            , {調査区分.経営分析調査_牛乳生産費, "K000014"} _
            , {調査区分.経営分析調査_子牛生産費, "K000011"} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, "K000010"} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, "K000010"} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "K000010"} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, "K000010"} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, "K000010"} _
            , {調査区分.経営分析調査_肥育豚生産費, "K000014"}
        }

        ''' <summary>集計倍率</summary>
        Public Shared 集計倍率 As New Dictionary(Of String, String) From {
              {調査区分.営農類型別経営統計_個人, "K000012"} _
            , {調査区分.営農類型別経営統計_法人, "K000015"} _
            , {調査区分.米生産費統計_個別, "K000006"} _
            , {調査区分.小麦生産費統計_個別, "K000007"} _
            , {調査区分.二条大麦生産費統計_個別, "K000007"} _
            , {調査区分.六条大麦生産費統計_個別, "K000007"} _
            , {調査区分.はだか麦生産費統計_個別, "K000007"} _
            , {調査区分.そば生産費統計_個別, "K000007"} _
            , {調査区分.大豆生産費統計_個別, "K000007"} _
            , {調査区分.原料用かんしょ生産費統計_個別, "K000007"} _
            , {調査区分.原料用ばれいしょ生産費統計_個別, "K000007"} _
            , {調査区分.なたね生産費統計_個別, "K000007"} _
            , {調査区分.てんさい生産費統計_個別, "K000007"} _
            , {調査区分.さとうきび生産費統計_個別, "K000007"} _
            , {調査区分.米生産費統計_組織法人, "K000006"} _
            , {調査区分.小麦生産費統計_組織法人, "K000007"} _
            , {調査区分.大豆生産費統計_組織法人, "K000007"} _
            , {調査区分.牛乳生産費統計_個別, "K000005"} _
            , {調査区分.子牛生産費統計_個別, "K000005"} _
            , {調査区分.乳用雄育成牛生産費統計_個別, "K000005"} _
            , {調査区分.交雑種育成牛生産費統計_個別, "K000005"} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, "K000005"} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, "K000005"} _
            , {調査区分.交雑種肥育牛生産費統計_個別, "K000005"} _
            , {調査区分.肥育豚生産費統計_個別, "K000005"} _
            , {調査区分.経営分析調査_二条大麦生産費, "K000007"} _
            , {調査区分.経営分析調査_六条大麦生産費, "K000007"} _
            , {調査区分.経営分析調査_はだか麦生産費, "K000007"} _
            , {調査区分.経営分析調査_そば生産費, "K000007"} _
            , {調査区分.経営分析調査_原料用ばれいしょ生産費, "K000007"} _
            , {調査区分.経営分析調査_なたね生産費, "K000007"} _
            , {調査区分.経営分析調査_てんさい生産費, "K000007"} _
            , {調査区分.経営分析調査_さとうきび生産費, "K000007"} _
            , {調査区分.経営分析調査_牛乳生産費, "K000005"} _
            , {調査区分.経営分析調査_子牛生産費, "K000005"} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, "K000005"} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, "K000005"} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "K000005"} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, "K000005"} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, "K000005"} _
            , {調査区分.経営分析調査_肥育豚生産費, "K000005"}
        }

        ''' <summary>部門集計倍率</summary>
        Public Shared 部門集計倍率 As New Dictionary(Of String, String) From {
            {調査区分.営農類型別経営統計_個人, "K000013"}
        }

        ''' <summary>指定部門</summary>
        Public Shared 指定部門 As New Dictionary(Of String, String) From {
            {調査区分.営農類型別経営統計_個人, "K000007"}
        }

        ''' <summary>農業地域区分１次</summary>
        Public Shared 農業地域区分１次 As New Dictionary(Of String, String) From {
            {調査区分.営農類型別経営統計_個人, "K000009"} _
           , {調査区分.営農類型別経営統計_法人, "K000013"}
        }

        ''' <summary>農業地域区分２次</summary>
        Public Shared 農業地域区分２次 As New Dictionary(Of String, String) From {
            {調査区分.営農類型別経営統計_個人, "K000010"} _
           , {調査区分.営農類型別経営統計_法人, "K000014"}
        }

        ''' <summary>認定農業者区分</summary>
        Public Shared 認定農業者区分 As New Dictionary(Of String, String) From {
            {調査区分.営農類型別経営統計_個人, "K000015"} _
           , {調査区分.営農類型別経営統計_法人, "K000017"}
        }

        ''' <summary>一位作目</summary>
        Public Shared 一位作目 As New Dictionary(Of String, String) From {
            {調査区分.営農類型別経営統計_個人, "K000019"} _
           , {調査区分.営農類型別経営統計_法人, "K000019"}
        }

        ''' <summary>二位作目</summary>
        Public Shared 二位作目 As New Dictionary(Of String, String) From {
            {調査区分.営農類型別経営統計_個人, "K000020"} _
           , {調査区分.営農類型別経営統計_法人, "K000020"}
        }

        ''' <summary>農業関連事業区分</summary>
        Public Shared 農業関連事業区分 As New Dictionary(Of String, String) From {
            {調査区分.営農類型別経営統計_個人, "K000024"} _
           , {調査区分.営農類型別経営統計_法人, "K000024"}
        }

        ''' <summary>集計対象区分</summary>
        Public Shared 集計対象区分 As New Dictionary(Of String, String) From {
            {調査区分.営農類型別経営統計_個人, "K000011"}
        }

        ''' <summary>任意項目１</summary>
        Public Shared 任意項目１ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020445"} _
           , {調査区分.営農類型別経営統計_法人, "K031553"}
        }

        ''' <summary>任意項目２</summary>
        Public Shared 任意項目２ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020446"} _
           , {調査区分.営農類型別経営統計_法人, "K031554"}
        }

        ''' <summary>任意項目３</summary>
        Public Shared 任意項目３ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020447"} _
           , {調査区分.営農類型別経営統計_法人, "K031555"}
        }

        ''' <summary>任意項目４</summary>
        Public Shared 任意項目４ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020545"} _
           , {調査区分.営農類型別経営統計_法人, "K031556"}
        }

        ''' <summary>任意項目５</summary>
        Public Shared 任意項目５ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020546"} _
           , {調査区分.営農類型別経営統計_法人, "K031557"}
        }

        ''' <summary>任意項目６</summary>
        Public Shared 任意項目６ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020547"} _
           , {調査区分.営農類型別経営統計_法人, "K031558"}
        }

        ''' <summary>任意項目７</summary>
        Public Shared 任意項目７ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020645"} _
           , {調査区分.営農類型別経営統計_法人, "K031559"}
        }

        ''' <summary>任意項目８</summary>
        Public Shared 任意項目８ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020646"} _
           , {調査区分.営農類型別経営統計_法人, "K031560"}
        }

        ''' <summary>任意項目９</summary>
        Public Shared 任意項目９ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020647"} _
           , {調査区分.営農類型別経営統計_法人, "K031561"}
        }

        ''' <summary>任意項目１０</summary>
        Public Shared 任意項目１０ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020744"} _
           , {調査区分.営農類型別経営統計_法人, "K031562"}
        }

        ''' <summary>任意項目１１</summary>
        Public Shared 任意項目１１ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020745"} _
           , {調査区分.営農類型別経営統計_法人, "K031653"}
        }

        ''' <summary>任意項目１２</summary>
        Public Shared 任意項目１２ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020746"} _
           , {調査区分.営農類型別経営統計_法人, "K031654"}
        }

        ''' <summary>任意項目１３</summary>
        Public Shared 任意項目１３ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020747"} _
           , {調査区分.営農類型別経営統計_法人, "K031655"}
        }

        ''' <summary>任意項目１４</summary>
        Public Shared 任意項目１４ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020844"} _
           , {調査区分.営農類型別経営統計_法人, "K031656"}
        }

        ''' <summary>任意項目１５</summary>
        Public Shared 任意項目１５ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020845"} _
           , {調査区分.営農類型別経営統計_法人, "K031657"}
        }

        ''' <summary>任意項目１６</summary>
        Public Shared 任意項目１６ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020846"} _
           , {調査区分.営農類型別経営統計_法人, "K031658"}
        }

        ''' <summary>任意項目１７</summary>
        Public Shared 任意項目１７ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020847"} _
           , {調査区分.営農類型別経営統計_法人, "K031659"}
        }

        ''' <summary>任意項目１８</summary>
        Public Shared 任意項目１８ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020944"} _
           , {調査区分.営農類型別経営統計_法人, "K031660"}
        }

        ''' <summary>任意項目１９</summary>
        Public Shared 任意項目１９ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020945"} _
           , {調査区分.営農類型別経営統計_法人, "K031661"}
        }

        ''' <summary>任意項目２０</summary>
        Public Shared 任意項目２０ As New Dictionary(Of String, String) From {
             {調査区分.営農類型別経営統計_個人, "K020946"} _
           , {調査区分.営農類型別経営統計_法人, "K031662"}
        }

        '---REV.002 ADD START
        ''' <summary>前回センサス番号</summary>
        Public Shared 規模階層 As New Dictionary(Of String, String) From {
              {調査区分.営農類型別経営統計_個人, "K000006"} _
            , {調査区分.営農類型別経営統計_法人, "K000007"} _
            , {調査区分.米生産費統計_個別, "K000007"} _
            , {調査区分.小麦生産費統計_個別, "K000008"} _
            , {調査区分.二条大麦生産費統計_個別, "K000008"} _
            , {調査区分.六条大麦生産費統計_個別, "K000008"} _
            , {調査区分.はだか麦生産費統計_個別, "K000008"} _
            , {調査区分.そば生産費統計_個別, "K000008"} _
            , {調査区分.大豆生産費統計_個別, "K000008"} _
            , {調査区分.原料用かんしょ生産費統計_個別, "K000008"} _
            , {調査区分.原料用ばれいしょ生産費統計_個別, "K000008"} _
            , {調査区分.なたね生産費統計_個別, "K000008"} _
            , {調査区分.てんさい生産費統計_個別, "K000008"} _
            , {調査区分.さとうきび生産費統計_個別, "K000008"} _
            , {調査区分.米生産費統計_組織法人, "K000007"} _
            , {調査区分.小麦生産費統計_組織法人, "K000008"} _
            , {調査区分.大豆生産費統計_組織法人, "K000008"} _
            , {調査区分.牛乳生産費統計_個別, "K000006"} _
            , {調査区分.子牛生産費統計_個別, "K000006"} _
            , {調査区分.乳用雄育成牛生産費統計_個別, "K000006"} _
            , {調査区分.交雑種育成牛生産費統計_個別, "K000006"} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, "K000006"} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, "K000006"} _
            , {調査区分.交雑種肥育牛生産費統計_個別, "K000006"} _
            , {調査区分.肥育豚生産費統計_個別, "K000006"} _
            , {調査区分.経営分析調査_二条大麦生産費, "K000008"} _
            , {調査区分.経営分析調査_六条大麦生産費, "K000008"} _
            , {調査区分.経営分析調査_はだか麦生産費, "K000008"} _
            , {調査区分.経営分析調査_そば生産費, "K000008"} _
            , {調査区分.経営分析調査_原料用ばれいしょ生産費, "K000008"} _
            , {調査区分.経営分析調査_なたね生産費, "K000008"} _
            , {調査区分.経営分析調査_てんさい生産費, "K000008"} _
            , {調査区分.経営分析調査_さとうきび生産費, "K000008"} _
            , {調査区分.経営分析調査_牛乳生産費, "K000006"} _
            , {調査区分.経営分析調査_子牛生産費, "K000006"} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, "K000006"} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, "K000006"} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "K000006"} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, "K000006"} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, "K000006"} _
            , {調査区分.経営分析調査_肥育豚生産費, "K000006"}
        }
        '---REV.002 ADD END

        ' REV_008↓
        '''' <summary>農業経営体の任意集計条件で置換する用の文字列</summary>
        'Public Shared 任意集計条件文字列 As New Dictionary(Of String, Dictionary(Of String, String)) From {
        '   {"【営農類型】", 営農類型} _
        '  , {"【農業地域区分１次】", 農業地域区分１次} _
        '  , {"【農業地域区分２次】", 農業地域区分２次} _
        '  , {"【認定農業者区分】", 認定農業者区分} _
        '  , {"【１位作目】", 一位作目} _
        '  , {"【２位作目】", 二位作目} _
        '  , {"【農業関連事業区分】", 農業関連事業区分} _
        '  , {"【任意項目１】", 任意項目１} _
        '  , {"【任意項目２】", 任意項目２} _
        '  , {"【任意項目３】", 任意項目３} _
        '  , {"【任意項目４】", 任意項目４} _
        '  , {"【任意項目５】", 任意項目５} _
        '  , {"【任意項目６】", 任意項目６} _
        '  , {"【任意項目７】", 任意項目７} _
        '  , {"【任意項目８】", 任意項目８} _
        '  , {"【任意項目９】", 任意項目９} _
        '  , {"【任意項目１０】", 任意項目１０} _
        '  , {"【任意項目１１】", 任意項目１１} _
        '  , {"【任意項目１２】", 任意項目１２} _
        '  , {"【任意項目１３】", 任意項目１３} _
        '  , {"【任意項目１４】", 任意項目１４} _
        '  , {"【任意項目１５】", 任意項目１５} _
        '  , {"【任意項目１６】", 任意項目１６} _
        '  , {"【任意項目１７】", 任意項目１７} _
        '  , {"【任意項目１８】", 任意項目１８} _
        '  , {"【任意項目１９】", 任意項目１９} _
        '  , {"【任意項目２０】", 任意項目２０}
        '}

        ''' <summary>農業経営体の任意集計条件で置換する用の文字列（令和元年体系）</summary>
        Public Shared 任意集計条件文字列2021 As New Dictionary(Of String, Dictionary(Of String, String)) From {
           {"【営農類型】", 営農類型} _
          , {"【農業地域区分１次】", 農業地域区分１次} _
          , {"【農業地域区分２次】", 農業地域区分２次} _
          , {"【認定農業者区分】", 認定農業者区分} _
          , {"【１位作目】", 一位作目} _
          , {"【２位作目】", 二位作目} _
          , {"【農業関連事業区分】", 農業関連事業区分} _
          , {"【任意項目１】", 任意項目１} _
          , {"【任意項目２】", 任意項目２} _
          , {"【任意項目３】", 任意項目３} _
          , {"【任意項目４】", 任意項目４} _
          , {"【任意項目５】", 任意項目５} _
          , {"【任意項目６】", 任意項目６} _
          , {"【任意項目７】", 任意項目７} _
          , {"【任意項目８】", 任意項目８} _
          , {"【任意項目９】", 任意項目９} _
          , {"【任意項目１０】", 任意項目１０} _
          , {"【任意項目１１】", 任意項目１１} _
          , {"【任意項目１２】", 任意項目１２} _
          , {"【任意項目１３】", 任意項目１３} _
          , {"【任意項目１４】", 任意項目１４} _
          , {"【任意項目１５】", 任意項目１５} _
          , {"【任意項目１６】", 任意項目１６} _
          , {"【任意項目１７】", 任意項目１７} _
          , {"【任意項目１８】", 任意項目１８} _
          , {"【任意項目１９】", 任意項目１９} _
          , {"【任意項目２０】", 任意項目２０}
        }

        ''' <summary>農業経営体の任意集計条件で置換する用の文字列（令和４年体系）</summary>
        Public Shared 任意集計条件文字列2022 As New Dictionary(Of String, Dictionary(Of String, String)) From {
           {"【営農類型】", 営農類型} _
          , {"【農業地域区分１次】", 農業地域区分１次} _
          , {"【農業地域区分２次】", 農業地域区分２次} _
          , {"【認定農業者区分】", 認定農業者区分} _
          , {"【１位作目】", 一位作目} _
          , {"【２位作目】", 二位作目} _
          , {"【農業関連事業区分】", 農業関連事業区分} _
          , {"【任意項目１】", 任意項目１} _
          , {"【任意項目２】", 任意項目２} _
          , {"【任意項目３】", 任意項目３} _
          , {"【任意項目４】", 任意項目４} _
          , {"【任意項目５】", 任意項目５} _
          , {"【任意項目６】", 任意項目６} _
          , {"【任意項目７】", 任意項目７} _
          , {"【任意項目８】", 任意項目８} _
          , {"【任意項目９】", 任意項目９} _
          , {"【任意項目１３】", 任意項目１３} _
          , {"【任意項目１５】", 任意項目１５} _
          , {"【任意項目１６】", 任意項目１６} _
          , {"【任意項目１７】", 任意項目１７}
        }

        ''' <summary>農業経営体の任意集計条件で置換する用の文字列</summary>
        Public Shared 任意集計条件文字列 As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, String))) From {
            {ComConst.バージョン区分.結果表等項目2021, 任意集計条件文字列2021} _
            , {ComConst.バージョン区分.結果表等項目2022, 任意集計条件文字列2022}}
        ' REV_008↑

        ''' <summary>修正不可項目</summary>
        Public Shared 修正不可項目 As New Dictionary(Of String, String()) From {
              {調査区分.営農類型別経営統計_個人, {"K000001", "K000002", "K000003"}} _
            , {調査区分.営農類型別経営統計_法人, {"K000001", "K000002", "K000003"}} _
            , {調査区分.米生産費統計_個別, {"K000001", "K000002", "K000003", "K000004"}} _
            , {調査区分.小麦生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000013"}} _
            , {調査区分.二条大麦生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000013"}} _
            , {調査区分.六条大麦生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000013"}} _
            , {調査区分.はだか麦生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000013"}} _
            , {調査区分.そば生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000013"}} _
            , {調査区分.大豆生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000013"}} _
            , {調査区分.原料用かんしょ生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000013"}} _
            , {調査区分.原料用ばれいしょ生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000013"}} _
            , {調査区分.なたね生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000013"}} _
            , {調査区分.てんさい生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000013"}} _
            , {調査区分.さとうきび生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000013"}} _
            , {調査区分.米生産費統計_組織法人, {"K000001", "K000002", "K000003", "K000004"}} _
            , {調査区分.小麦生産費統計_組織法人, {"K000001", "K000002", "K000003", "K000004"}} _
            , {調査区分.大豆生産費統計_組織法人, {"K000001", "K000002", "K000003", "K000004"}} _
            , {調査区分.牛乳生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000009", "K000010", "K000011"}} _
            , {調査区分.子牛生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.乳用雄育成牛生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.交雑種育成牛生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.交雑種肥育牛生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.肥育豚生産費統計_個別, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.経営分析調査_二条大麦生産費, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000014"}} _
            , {調査区分.経営分析調査_六条大麦生産費, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000014"}} _
            , {調査区分.経営分析調査_はだか麦生産費, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000014"}} _
            , {調査区分.経営分析調査_そば生産費, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000014"}} _
            , {調査区分.経営分析調査_原料用ばれいしょ生産費, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000014"}} _
            , {調査区分.経営分析調査_なたね生産費, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000014"}} _
            , {調査区分.経営分析調査_てんさい生産費, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000014"}} _
            , {調査区分.経営分析調査_さとうきび生産費, {"K000001", "K000002", "K000003", "K000004", "K000005", "K000009", "K000014"}} _
            , {調査区分.経営分析調査_牛乳生産費, {"K000001", "K000002", "K000003", "K000004", "K000009", "K000010", "K000011"}} _
            , {調査区分.経営分析調査_子牛生産費, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, {"K000001", "K000002", "K000003", "K000004", "K000008"}} _
            , {調査区分.経営分析調査_肥育豚生産費, {"K000001", "K000002", "K000003", "K000004", "K000008"}}
        }

        ''' <summary>営農個人修正可能項目</summary>
        Public Shared 営農個人修正可能項目 As String() = {"K010713", "K010720", "K010734", "K010909"}

        ' REV_008↓
        '''' <summary>入力用ファイル名称</summary>
        'Public Shared 入力用ファイル名称 As New Dictionary(Of String, String) From {
        '      {調査区分.営農類型別経営統計_個人, "農業経営_営農個人_個別結果表.xlsm"} _
        '    , {調査区分.営農類型別経営統計_法人, "農業経営_営農法人_個別結果表.xlsm"} _
        '    , {調査区分.米生産費統計_個別, "農業経営_米生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.小麦生産費統計_個別, "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.二条大麦生産費統計_個別, "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.六条大麦生産費統計_個別, "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.はだか麦生産費統計_個別, "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.そば生産費統計_個別, "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.大豆生産費統計_個別, "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.原料用かんしょ生産費統計_個別, "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.原料用ばれいしょ生産費統計_個別, "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.なたね生産費統計_個別, "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.てんさい生産費統計_個別, "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.さとうきび生産費統計_個別, "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.米生産費統計_組織法人, "農業経営_米生産費（組織法人）_個別結果表.xlsm"} _
        '    , {調査区分.小麦生産費統計_組織法人, "農業経営_小麦・大豆生産費（組織法人）_個別結果表.xlsm"} _
        '    , {調査区分.大豆生産費統計_組織法人, "農業経営_小麦・大豆生産費（組織法人）_個別結果表.xlsm"} _
        '    , {調査区分.牛乳生産費統計_個別, "農業経営_牛乳生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.子牛生産費統計_個別, "農業経営_子牛生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.乳用雄育成牛生産費統計_個別, "農業経営_育成牛・肥育牛生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.交雑種育成牛生産費統計_個別, "農業経営_育成牛・肥育牛生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.去勢若齢肥育牛生産費統計_個別, "農業経営_育成牛・肥育牛生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.乳用雄肥育牛生産費統計_個別, "農業経営_育成牛・肥育牛生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.交雑種肥育牛生産費統計_個別, "農業経営_育成牛・肥育牛生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.肥育豚生産費統計_個別, "農業経営_肥育豚生産費（個別）_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_二条大麦生産費, "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_六条大麦生産費, "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_はだか麦生産費, "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_そば生産費, "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_原料用ばれいしょ生産費, "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_なたね生産費, "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_てんさい生産費, "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_さとうきび生産費, "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_牛乳生産費, "経営分析_牛乳生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_子牛生産費, "経営分析_子牛生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_乳用雄育成牛生産費, "経営分析_育成牛・肥育牛生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_交雑種育成牛生産費, "経営分析_育成牛・肥育牛生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "経営分析_育成牛・肥育牛生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_乳用雄肥育牛生産費, "経営分析_育成牛・肥育牛生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_交雑種肥育牛生産費, "経営分析_育成牛・肥育牛生産費_個別結果表.xlsm"} _
        '    , {調査区分.経営分析調査_肥育豚生産費, "経営分析_肥育豚生産費_個別結果表.xlsm"}
        '}

        ''' <summary>入力用ファイル名称</summary>
        Public Shared 入力用ファイル名称 As New Dictionary(Of Tuple(Of String, String), String) From {
              {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.結果表等項目2021), "農業経営_営農個人_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.結果表等項目2021), "農業経営_営農法人_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_米生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), "農業経営_米生産費（組織法人）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), "農業経営_小麦・大豆生産費（組織法人）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), "農業経営_小麦・大豆生産費（組織法人）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_牛乳生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_子牛生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_育成牛・肥育牛生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_育成牛・肥育牛生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_育成牛・肥育牛生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_育成牛・肥育牛生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_育成牛・肥育牛生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), "農業経営_肥育豚生産費（個別）_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_麦類・そば・なたね・畑作物生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_牛乳生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_子牛生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_育成牛・肥育牛生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_育成牛・肥育牛生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_育成牛・肥育牛生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_育成牛・肥育牛生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_育成牛・肥育牛生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.結果表等項目2021), "経営分析_肥育豚生産費_個別結果表.xlsm"} _
            , {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.結果表等項目2022), "農業経営_営農個人_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.結果表等項目2022), "農業経営_営農法人_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_米生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), "農業経営_米生産費（組織法人）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), "農業経営_小麦・大豆生産費（組織法人）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), "農業経営_小麦・大豆生産費（組織法人）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_牛乳生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_子牛生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_育成牛・肥育牛生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_育成牛・肥育牛生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_育成牛・肥育牛生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_育成牛・肥育牛生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_育成牛・肥育牛生産費（個別）_個別結果表_2022.xlsm"} _
            , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), "農業経営_肥育豚生産費（個別）_個別結果表_2022.xlsm"}
        }
        ' REV_008↑

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            Public Class Report
                'テンプレートファイル名
                Public tempFileName As String
                '帳票名
                Public reportName As String
            End Class

            ' REV_008↓
            'Public Shared リスト As New Dictionary(Of String, Report) From {
            '      {調査区分.営農類型別経営統計_個人, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.営農類型別経営統計_個人), .reportName = "農業経営_営農個人_個別結果表"}} _
            '    , {調査区分.営農類型別経営統計_法人, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.営農類型別経営統計_法人), .reportName = "農業経営_営農法人_個別結果表"}} _
            '    , {調査区分.米生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.米生産費統計_個別), .reportName = "農業経営_米生産費（個別）_個別結果表"}} _
            '    , {調査区分.小麦生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.小麦生産費統計_個別), .reportName = "農業経営_小麦生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.二条大麦生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.二条大麦生産費統計_個別), .reportName = "農業経営_二条大麦生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.六条大麦生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.六条大麦生産費統計_個別), .reportName = "農業経営_六条大麦生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.はだか麦生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.はだか麦生産費統計_個別), .reportName = "農業経営_はだか麦生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.そば生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.そば生産費統計_個別), .reportName = "農業経営_そば生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.大豆生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.大豆生産費統計_個別), .reportName = "農業経営_大豆生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.原料用かんしょ生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.原料用かんしょ生産費統計_個別), .reportName = "農業経営_原料用かんしょ生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.原料用ばれいしょ生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.原料用ばれいしょ生産費統計_個別), .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.なたね生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.なたね生産費統計_個別), .reportName = "農業経営_なたね生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.てんさい生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.てんさい生産費統計_個別), .reportName = "農業経営_てんさい生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.さとうきび生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.さとうきび生産費統計_個別), .reportName = "農業経営_さとうきび生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.米生産費統計_組織法人, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.米生産費統計_組織法人), .reportName = "農業経営_米生産費（組織法人）_個別結果表"}} _
            '    , {調査区分.小麦生産費統計_組織法人, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.小麦生産費統計_組織法人), .reportName = "農業経営_小麦生産費統計（組織法人）_個別結果表"}} _
            '    , {調査区分.大豆生産費統計_組織法人, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.大豆生産費統計_組織法人), .reportName = "農業経営_大豆生産費統計（組織法人）_個別結果表"}} _
            '    , {調査区分.牛乳生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.牛乳生産費統計_個別), .reportName = "農業経営_牛乳生産費（個別）_個別結果表"}} _
            '    , {調査区分.子牛生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.子牛生産費統計_個別), .reportName = "農業経営_子牛生産費（個別）_個別結果表"}} _
            '    , {調査区分.乳用雄育成牛生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.乳用雄育成牛生産費統計_個別), .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.交雑種育成牛生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.交雑種育成牛生産費統計_個別), .reportName = "農業経営_交雑種育成牛生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.去勢若齢肥育牛生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.去勢若齢肥育牛生産費統計_個別), .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.乳用雄肥育牛生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.乳用雄肥育牛生産費統計_個別), .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.交雑種肥育牛生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.交雑種肥育牛生産費統計_個別), .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_個別結果表"}} _
            '    , {調査区分.肥育豚生産費統計_個別, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.肥育豚生産費統計_個別), .reportName = "農業経営_肥育豚生産費（個別）_個別結果表"}} _
            '    , {調査区分.経営分析調査_二条大麦生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_二条大麦生産費), .reportName = "経営分析_二条大麦生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_六条大麦生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_六条大麦生産費), .reportName = "経営分析_六条大麦生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_はだか麦生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_はだか麦生産費), .reportName = "経営分析_はだか麦生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_そば生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_そば生産費), .reportName = "経営分析_そば生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_原料用ばれいしょ生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_原料用ばれいしょ生産費), .reportName = "経営分析_原料用ばれいしょ生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_なたね生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_なたね生産費), .reportName = "経営分析_なたね生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_てんさい生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_てんさい生産費), .reportName = "経営分析_てんさい生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_さとうきび生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_さとうきび生産費), .reportName = "経営分析_さとうきび生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_牛乳生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_牛乳生産費), .reportName = "経営分析_牛乳生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_子牛生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_子牛生産費), .reportName = "経営分析_子牛生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_乳用雄育成牛生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_乳用雄育成牛生産費), .reportName = "経営分析_乳用雄育成牛生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_交雑種育成牛生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_交雑種育成牛生産費), .reportName = "経営分析_交雑種育成牛生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_去勢若齢肥育牛生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_去勢若齢肥育牛生産費), .reportName = "経営分析_去勢若齢肥育牛生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_乳用雄肥育牛生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_乳用雄肥育牛生産費), .reportName = "経営分析_乳用雄肥育牛生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_交雑種肥育牛生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_交雑種肥育牛生産費), .reportName = "経営分析_交雑種肥育牛生産費_個別結果表"}} _
            '    , {調査区分.経営分析調査_肥育豚生産費, New Report With {.tempFileName = 個別結果表.入力用ファイル名称(調査区分.経営分析調査_肥育豚生産費), .reportName = "経営分析_肥育豚生産費_個別結果表"}}
            '}
            Public Shared リスト As New Dictionary(Of Tuple(Of String, String), Report) From {
                  {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_営農個人_個別結果表"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_営農法人_個別結果表"}} _
                , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_米生産費（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_小麦生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_二条大麦生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_六条大麦生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_はだか麦生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_そば生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_大豆生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_原料用かんしょ生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_なたね生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_てんさい生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_さとうきび生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_米生産費（組織法人）_個別結果表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_小麦生産費統計（組織法人）_個別結果表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_大豆生産費統計（組織法人）_個別結果表"}} _
                , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_牛乳生産費（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_子牛生産費（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_交雑種育成牛生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2021)), .reportName = "農業経営_肥育豚生産費（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_二条大麦生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_六条大麦生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_はだか麦生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_そば生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_原料用ばれいしょ生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_なたね生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_てんさい生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_さとうきび生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_牛乳生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_子牛生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_乳用雄育成牛生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_交雑種育成牛生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_去勢若齢肥育牛生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_乳用雄肥育牛生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_交雑種肥育牛生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.結果表等項目2021)), .reportName = "経営分析_肥育豚生産費_個別結果表"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_営農個人_個別結果表"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_営農法人_個別結果表"}} _
                , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_米生産費（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_小麦生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_二条大麦生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_六条大麦生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_はだか麦生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_そば生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_大豆生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_原料用かんしょ生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_なたね生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_てんさい生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_さとうきび生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_米生産費（組織法人）_個別結果表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_小麦生産費統計（組織法人）_個別結果表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_大豆生産費統計（組織法人）_個別結果表"}} _
                , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_牛乳生産費（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_子牛生産費（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_交雑種育成牛生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_個別結果表"}} _
                , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = 個別結果表.入力用ファイル名称(Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2022)), .reportName = "農業経営_肥育豚生産費（個別）_個別結果表"}}
            }
        End Class
        ' REV_008↑

        ''' <summary>集計用テーブル付加名称</summary>
        Public Const 集計用テーブル付加名称 As String = "＿集計用"
    End Class

    ''' <summary>
    ''' 個別結果検討表クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 個別結果検討表

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            Public Class Report
                'テンプレートファイル名
                Public tempFileName As String
                '帳票名
                Public reportName As String
            End Class

            ' REV_008↓
            'Public Shared リスト As New Dictionary(Of String, Report) From {
            '      {調査区分.営農類型別経営統計_個人, New Report With {.tempFileName = "農業経営_営農個人_個別結果検討表.xlsx", .reportName = "農業経営_営農個人_個別結果検討表"}} _
            '    , {調査区分.営農類型別経営統計_法人, New Report With {.tempFileName = "農業経営_営農法人_個別結果検討表.xlsx", .reportName = "農業経営_営農法人_個別結果検討表"}} _
            '    , {調査区分.米生産費統計_個別, New Report With {.tempFileName = "農業経営_米生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_米生産費（個別）_個別結果検討表"}} _
            '    , {調査区分.小麦生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_小麦生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.二条大麦生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_二条大麦生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.六条大麦生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_六条大麦生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.はだか麦生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_はだか麦生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.そば生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_そば生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.大豆生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_大豆生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.原料用かんしょ生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_原料用かんしょ生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.原料用ばれいしょ生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.なたね生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_なたね生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.てんさい生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_てんさい生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.さとうきび生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_さとうきび生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.米生産費統計_組織法人, New Report With {.tempFileName = "農業経営_米生産費（組織法人）_個別結果検討表.xlsx", .reportName = "農業経営_米生産費（組織法人）_個別結果検討表"}} _
            '    , {調査区分.小麦生産費統計_組織法人, New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_個別結果検討表.xlsx", .reportName = "農業経営_小麦生産費統計（組織法人）_個別結果検討表"}} _
            '    , {調査区分.大豆生産費統計_組織法人, New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_個別結果検討表.xlsx", .reportName = "農業経営_大豆生産費統計（組織法人）_個別結果検討表"}} _
            '    , {調査区分.牛乳生産費統計_個別, New Report With {.tempFileName = "農業経営_牛乳生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_牛乳生産費（個別）_個別結果検討表"}} _
            '    , {調査区分.子牛生産費統計_個別, New Report With {.tempFileName = "農業経営_子牛生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_子牛生産費（個別）_個別結果検討表"}} _
            '    , {調査区分.乳用雄育成牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.交雑種育成牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_交雑種育成牛生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.去勢若齢肥育牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.乳用雄肥育牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.交雑種肥育牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_個別結果検討表"}} _
            '    , {調査区分.肥育豚生産費統計_個別, New Report With {.tempFileName = "農業経営_肥育豚生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_肥育豚生産費（個別）_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_二条大麦生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_二条大麦生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_六条大麦生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_六条大麦生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_はだか麦生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_はだか麦生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_そば生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_そば生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_原料用ばれいしょ生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_原料用ばれいしょ生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_なたね生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_なたね生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_てんさい生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_てんさい生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_さとうきび生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_さとうきび生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_牛乳生産費, New Report With {.tempFileName = "経営分析_牛乳生産費_個別結果検討表.xlsx", .reportName = "経営分析_牛乳生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_子牛生産費, New Report With {.tempFileName = "経営分析_子牛生産費_個別結果検討表.xlsx", .reportName = "経営分析_子牛生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_乳用雄育成牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_個別結果検討表.xlsx", .reportName = "経営分析_乳用雄育成牛生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_交雑種育成牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_個別結果検討表.xlsx", .reportName = "経営分析_交雑種育成牛生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_去勢若齢肥育牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_個別結果検討表.xlsx", .reportName = "経営分析_去勢若齢肥育牛生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_乳用雄肥育牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_個別結果検討表.xlsx", .reportName = "経営分析_乳用雄肥育牛生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_交雑種肥育牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_個別結果検討表.xlsx", .reportName = "経営分析_交雑種肥育牛生産費_個別結果検討表"}} _
            '    , {調査区分.経営分析調査_肥育豚生産費, New Report With {.tempFileName = "経営分析_肥育豚生産費_個別結果検討表.xlsx", .reportName = "経営分析_肥育豚生産費_個別結果検討表"}}
            '}
            Public Shared リスト As New Dictionary(Of Tuple(Of String, String), Report) From {
                  {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_営農個人_個別結果検討表.xlsx", .reportName = "農業経営_営農個人_個別結果検討表"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_営農法人_個別結果検討表.xlsx", .reportName = "農業経営_営農法人_個別結果検討表"}} _
                , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_米生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_米生産費（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_小麦生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_二条大麦生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_六条大麦生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_はだか麦生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_そば生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_大豆生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_原料用かんしょ生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_なたね生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_てんさい生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_さとうきび生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_米生産費（組織法人）_個別結果検討表.xlsx", .reportName = "農業経営_米生産費（組織法人）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_個別結果検討表.xlsx", .reportName = "農業経営_小麦生産費統計（組織法人）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_個別結果検討表.xlsx", .reportName = "農業経営_大豆生産費統計（組織法人）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_牛乳生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_牛乳生産費（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_子牛生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_子牛生産費（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_交雑種育成牛生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_肥育豚生産費（個別）_個別結果検討表.xlsx", .reportName = "農業経営_肥育豚生産費（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_二条大麦生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_六条大麦生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_はだか麦生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_そば生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_原料用ばれいしょ生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_なたね生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_てんさい生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_個別結果検討表.xlsx", .reportName = "経営分析_さとうきび生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_牛乳生産費_個別結果検討表.xlsx", .reportName = "経営分析_牛乳生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_子牛生産費_個別結果検討表.xlsx", .reportName = "経営分析_子牛生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_個別結果検討表.xlsx", .reportName = "経営分析_乳用雄育成牛生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_個別結果検討表.xlsx", .reportName = "経営分析_交雑種育成牛生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_個別結果検討表.xlsx", .reportName = "経営分析_去勢若齢肥育牛生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_個別結果検討表.xlsx", .reportName = "経営分析_乳用雄肥育牛生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_個別結果検討表.xlsx", .reportName = "経営分析_交雑種肥育牛生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_肥育豚生産費_個別結果検討表.xlsx", .reportName = "経営分析_肥育豚生産費_個別結果検討表"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_営農個人_個別結果検討表_2022.xlsx", .reportName = "農業経営_営農個人_個別結果検討表"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_営農法人_個別結果検討表_2022.xlsx", .reportName = "農業経営_営農法人_個別結果検討表"}} _
                , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_米生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_米生産費（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_小麦生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_二条大麦生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_六条大麦生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_はだか麦生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_そば生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_大豆生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_原料用かんしょ生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_なたね生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_てんさい生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_さとうきび生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_米生産費（組織法人）_個別結果検討表_2022.xlsx", .reportName = "農業経営_米生産費（組織法人）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_個別結果検討表_2022.xlsx", .reportName = "農業経営_小麦生産費統計（組織法人）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_個別結果検討表_2022.xlsx", .reportName = "農業経営_大豆生産費統計（組織法人）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_牛乳生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_牛乳生産費（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_子牛生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_子牛生産費（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_交雑種育成牛生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_個別結果検討表"}} _
                , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_肥育豚生産費（個別）_個別結果検討表_2022.xlsx", .reportName = "農業経営_肥育豚生産費（個別）_個別結果検討表"}}
            }
            ' REV_008↑
        End Class
    End Class

    ''' <summary>
    ''' 集計結果表クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 集計結果表
        ''' <summary>テーブル名称</summary>
        Public Shared テーブル名称 As New Dictionary(Of String, String()) From {
              {調査区分.営農類型別経営統計_個人, {"集計結果表＿農業経営＿営農類型＿個人１", "集計結果表＿農業経営＿営農類型＿個人２" _
                                                , "集計結果表＿農業経営＿営農類型＿個人３"}} _
            , {調査区分.営農類型別経営統計_法人, {"集計結果表＿農業経営＿営農類型＿法人１", "集計結果表＿農業経営＿営農類型＿法人２" _
                                                , "集計結果表＿農業経営＿営農類型＿法人３"}} _
            , {調査区分.米生産費統計_個別, {"集計結果表＿農業経営＿米＿個別１", "集計結果表＿農業経営＿米＿個別２" _
                                          , "集計結果表＿農業経営＿米＿個別３"}} _
            , {調査区分.小麦生産費統計_個別, {"集計結果表＿農業経営＿小麦＿個別１", "集計結果表＿農業経営＿小麦＿個別２" _
                                            , "集計結果表＿農業経営＿小麦＿個別３"}} _
            , {調査区分.二条大麦生産費統計_個別, {"集計結果表＿農業経営＿二条大麦＿個別１", "集計結果表＿農業経営＿二条大麦＿個別２" _
                                                , "集計結果表＿農業経営＿二条大麦＿個別３"}} _
            , {調査区分.六条大麦生産費統計_個別, {"集計結果表＿農業経営＿六条大麦＿個別１", "集計結果表＿農業経営＿六条大麦＿個別２" _
                                                , "集計結果表＿農業経営＿六条大麦＿個別３"}} _
            , {調査区分.はだか麦生産費統計_個別, {"集計結果表＿農業経営＿はだか麦＿個別１", "集計結果表＿農業経営＿はだか麦＿個別２" _
                                                , "集計結果表＿農業経営＿はだか麦＿個別３"}} _
            , {調査区分.そば生産費統計_個別, {"集計結果表＿農業経営＿そば＿個別１", "集計結果表＿農業経営＿そば＿個別２" _
                                            , "集計結果表＿農業経営＿そば＿個別３"}} _
            , {調査区分.大豆生産費統計_個別, {"集計結果表＿農業経営＿大豆＿個別１", "集計結果表＿農業経営＿大豆＿個別２" _
                                            , "集計結果表＿農業経営＿大豆＿個別３"}} _
            , {調査区分.原料用かんしょ生産費統計_個別, {"集計結果表＿農業経営＿原料用かんしょ＿個別１", "集計結果表＿農業経営＿原料用かんしょ＿個別２" _
                                                      , "集計結果表＿農業経営＿原料用かんしょ＿個別３"}} _
            , {調査区分.原料用ばれいしょ生産費統計_個別, {"集計結果表＿農業経営＿原料用ばれいしょ＿個別１", "集計結果表＿農業経営＿原料用ばれいしょ＿個別２" _
                                                        , "集計結果表＿農業経営＿原料用ばれいしょ＿個別３"}} _
            , {調査区分.なたね生産費統計_個別, {"集計結果表＿農業経営＿なたね＿個別１", "集計結果表＿農業経営＿なたね＿個別２" _
                                              , "集計結果表＿農業経営＿なたね＿個別３"}} _
            , {調査区分.てんさい生産費統計_個別, {"集計結果表＿農業経営＿てんさい＿個別１", "集計結果表＿農業経営＿てんさい＿個別２" _
                                                , "集計結果表＿農業経営＿てんさい＿個別３"}} _
            , {調査区分.さとうきび生産費統計_個別, {"集計結果表＿農業経営＿さとうきび＿個別１", "集計結果表＿農業経営＿さとうきび＿個別２" _
                                                  , "集計結果表＿農業経営＿さとうきび＿個別３"}} _
            , {調査区分.米生産費統計_組織法人, {"集計結果表＿農業経営＿米＿組織１", "集計結果表＿農業経営＿米＿組織２" _
                                              , "集計結果表＿農業経営＿米＿組織３"}} _
            , {調査区分.小麦生産費統計_組織法人, {"集計結果表＿農業経営＿小麦＿組織１", "集計結果表＿農業経営＿小麦＿組織２"}} _
            , {調査区分.大豆生産費統計_組織法人, {"集計結果表＿農業経営＿大豆＿組織１", "集計結果表＿農業経営＿大豆＿組織２"}} _
            , {調査区分.牛乳生産費統計_個別, {"集計結果表＿農業経営＿牛乳＿個別１", "集計結果表＿農業経営＿牛乳＿個別２", "集計結果表＿農業経営＿牛乳＿個別３"}} _
            , {調査区分.子牛生産費統計_個別, {"集計結果表＿農業経営＿子牛＿個別１", "集計結果表＿農業経営＿子牛＿個別２"}} _
            , {調査区分.乳用雄育成牛生産費統計_個別, {"集計結果表＿農業経営＿乳用雄育成牛＿個別１", "集計結果表＿農業経営＿乳用雄育成牛＿個別２"}} _
            , {調査区分.交雑種育成牛生産費統計_個別, {"集計結果表＿農業経営＿交雑種育成牛＿個別１", "集計結果表＿農業経営＿交雑種育成牛＿個別２"}} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, {"集計結果表＿農業経営＿去勢若齢肥育牛＿個別１", "集計結果表＿農業経営＿去勢若齢肥育牛＿個別２"}} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, {"集計結果表＿農業経営＿乳用雄肥育牛＿個別１", "集計結果表＿農業経営＿乳用雄肥育牛＿個別２"}} _
            , {調査区分.交雑種肥育牛生産費統計_個別, {"集計結果表＿農業経営＿交雑種肥育牛＿個別１", "集計結果表＿農業経営＿交雑種肥育牛＿個別２"}} _
            , {調査区分.肥育豚生産費統計_個別, {"集計結果表＿農業経営＿肥育豚＿個別１", "集計結果表＿農業経営＿肥育豚＿個別２"}} _
            , {調査区分.経営分析調査_二条大麦生産費, {"集計結果表＿経営分析＿二条大麦１", "集計結果表＿経営分析＿二条大麦２" _
                                                    , "集計結果表＿経営分析＿二条大麦３"}} _
            , {調査区分.経営分析調査_六条大麦生産費, {"集計結果表＿経営分析＿六条大麦１", "集計結果表＿経営分析＿六条大麦２" _
                                                    , "集計結果表＿経営分析＿六条大麦３"}} _
            , {調査区分.経営分析調査_はだか麦生産費, {"集計結果表＿経営分析＿はだか麦１", "集計結果表＿経営分析＿はだか麦２" _
                                                    , "集計結果表＿経営分析＿はだか麦３"}} _
            , {調査区分.経営分析調査_そば生産費, {"集計結果表＿経営分析＿そば１", "集計結果表＿経営分析＿そば２" _
                                                , "集計結果表＿経営分析＿そば３"}} _
            , {調査区分.経営分析調査_原料用ばれいしょ生産費, {"集計結果表＿経営分析＿原料用ばれいしょ１", "集計結果表＿経営分析＿原料用ばれいしょ２" _
                                                            , "集計結果表＿経営分析＿原料用ばれいしょ３"}} _
            , {調査区分.経営分析調査_なたね生産費, {"集計結果表＿経営分析＿なたね１", "集計結果表＿経営分析＿なたね２" _
                                                  , "集計結果表＿経営分析＿なたね３"}} _
            , {調査区分.経営分析調査_てんさい生産費, {"集計結果表＿経営分析＿てんさい１", "集計結果表＿経営分析＿てんさい２" _
                                                    , "集計結果表＿経営分析＿てんさい３"}} _
            , {調査区分.経営分析調査_さとうきび生産費, {"集計結果表＿経営分析＿さとうきび１", "集計結果表＿経営分析＿さとうきび２" _
                                                      , "集計結果表＿経営分析＿さとうきび３"}} _
            , {調査区分.経営分析調査_牛乳生産費, {"集計結果表＿経営分析＿牛乳１", "集計結果表＿経営分析＿牛乳２"}} _
            , {調査区分.経営分析調査_子牛生産費, {"集計結果表＿経営分析＿子牛１", "集計結果表＿経営分析＿子牛２"}} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, {"集計結果表＿経営分析＿乳用雄育成牛１", "集計結果表＿経営分析＿乳用雄育成牛２"}} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, {"集計結果表＿経営分析＿交雑種育成牛１", "集計結果表＿経営分析＿交雑種育成牛２"}} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, {"集計結果表＿経営分析＿去勢若齢肥育牛１", "集計結果表＿経営分析＿去勢若齢肥育牛２"}} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, {"集計結果表＿経営分析＿乳用雄肥育牛１", "集計結果表＿経営分析＿乳用雄肥育牛２"}} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, {"集計結果表＿経営分析＿交雑種肥育牛１", "集計結果表＿経営分析＿交雑種肥育牛２"}} _
            , {調査区分.経営分析調査_肥育豚生産費, {"集計結果表＿経営分析＿肥育豚１", "集計結果表＿経営分析＿肥育豚２"}} _
            , {営農経営体区分.農業経営体, {"集計結果表＿農業経営＿営農類型＿農業経営体４", "集計結果表＿農業経営＿営農類型＿農業経営体５" _
                                         , "集計結果表＿農業経営＿営農類型＿農業経営体６"}}}

        ''' <summary>集計戸数</summary>
        Public Shared 集計戸数 As New Dictionary(Of String, String) From {
              {調査区分.営農類型別経営統計_個人, "S000009"} _
            , {調査区分.営農類型別経営統計_法人, "S000009"} _
            , {調査区分.米生産費統計_個別, "S000007"} _
            , {調査区分.小麦生産費統計_個別, "S000007"} _
            , {調査区分.二条大麦生産費統計_個別, "S000007"} _
            , {調査区分.六条大麦生産費統計_個別, "S000007"} _
            , {調査区分.はだか麦生産費統計_個別, "S000007"} _
            , {調査区分.そば生産費統計_個別, "S000007"} _
            , {調査区分.大豆生産費統計_個別, "S000007"} _
            , {調査区分.原料用かんしょ生産費統計_個別, "S000007"} _
            , {調査区分.原料用ばれいしょ生産費統計_個別, "S000007"} _
            , {調査区分.なたね生産費統計_個別, "S000007"} _
            , {調査区分.てんさい生産費統計_個別, "S000007"} _
            , {調査区分.さとうきび生産費統計_個別, "S000007"} _
            , {調査区分.米生産費統計_組織法人, "S000007"} _
            , {調査区分.小麦生産費統計_組織法人, "S000007"} _
            , {調査区分.大豆生産費統計_組織法人, "S000007"} _
            , {調査区分.牛乳生産費統計_個別, "S000007"} _
            , {調査区分.子牛生産費統計_個別, "S000007"} _
            , {調査区分.乳用雄育成牛生産費統計_個別, "S000007"} _
            , {調査区分.交雑種育成牛生産費統計_個別, "S000007"} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, "S000007"} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, "S000007"} _
            , {調査区分.交雑種肥育牛生産費統計_個別, "S000007"} _
            , {調査区分.肥育豚生産費統計_個別, "S000007"} _
            , {調査区分.経営分析調査_二条大麦生産費, "S000007"} _
            , {調査区分.経営分析調査_六条大麦生産費, "S000007"} _
            , {調査区分.経営分析調査_はだか麦生産費, "S000007"} _
            , {調査区分.経営分析調査_そば生産費, "S000007"} _
            , {調査区分.経営分析調査_原料用ばれいしょ生産費, "S000007"} _
            , {調査区分.経営分析調査_なたね生産費, "S000007"} _
            , {調査区分.経営分析調査_てんさい生産費, "S000007"} _
            , {調査区分.経営分析調査_さとうきび生産費, "S000007"} _
            , {調査区分.経営分析調査_牛乳生産費, "S000007"} _
            , {調査区分.経営分析調査_子牛生産費, "S000007"} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, "S000007"} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, "S000007"} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "S000007"} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, "S000007"} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, "S000007"} _
            , {調査区分.経営分析調査_肥育豚生産費, "S000007"} _
            , {営農経営体区分.農業経営体, "S000009"}
        }

        ''' <summary>集計名称</summary>
        Public Shared 集計名称 As New Dictionary(Of String, String) From {
              {調査区分.営農類型別経営統計_個人, "S000002"} _
            , {調査区分.営農類型別経営統計_法人, "S000002"} _
            , {調査区分.牛乳生産費統計_個別, "S000010"} _
            , {調査区分.子牛生産費統計_個別, "S000010"} _
            , {調査区分.乳用雄育成牛生産費統計_個別, "S000010"} _
            , {調査区分.交雑種育成牛生産費統計_個別, "S000010"} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, "S000010"} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, "S000010"} _
            , {調査区分.交雑種肥育牛生産費統計_個別, "S000010"} _
            , {調査区分.肥育豚生産費統計_個別, "S000010"} _
            , {調査区分.経営分析調査_牛乳生産費, "S000010"} _
            , {調査区分.経営分析調査_子牛生産費, "S000010"} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, "S000010"} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, "S000010"} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "S000010"} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, "S000010"} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, "S000010"} _
            , {調査区分.経営分析調査_肥育豚生産費, "S000010"} _
            , {営農経営体区分.農業経営体, "S000002"}
        }

        ''' <summary>生産費平均値種類</summary>
        Public Shared 生産費平均値種類 As String = "S000006"

        ''' <summary>指標部</summary>
        Public Shared 指標部 As New Dictionary(Of String, String()) From {
              {調査区分.営農類型別経営統計_個人, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000014"}} _
            , {調査区分.営農類型別経営統計_法人, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000014"}} _
            , {調査区分.米生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000020"}} _
            , {調査区分.小麦生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.二条大麦生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.六条大麦生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.はだか麦生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.そば生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.大豆生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.原料用かんしょ生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.原料用ばれいしょ生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.なたね生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.てんさい生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.さとうきび生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.米生産費統計_組織法人, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000020"}} _
            , {調査区分.小麦生産費統計_組織法人, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000020"}} _
            , {調査区分.大豆生産費統計_組織法人, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000020"}} _
            , {調査区分.牛乳生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.子牛生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.乳用雄育成牛生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.交雑種育成牛生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.去勢若齢肥育牛生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.乳用雄肥育牛生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.交雑種肥育牛生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.肥育豚生産費統計_個別, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.経営分析調査_二条大麦生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.経営分析調査_六条大麦生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.経営分析調査_はだか麦生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.経営分析調査_そば生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.経営分析調査_原料用ばれいしょ生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.経営分析調査_なたね生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.経営分析調査_てんさい生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.経営分析調査_さとうきび生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000020"}} _
            , {調査区分.経営分析調査_牛乳生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.経営分析調査_子牛生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.経営分析調査_乳用雄育成牛生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.経営分析調査_交雑種育成牛生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.経営分析調査_去勢若齢肥育牛生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.経営分析調査_乳用雄肥育牛生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.経営分析調査_交雑種肥育牛生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {調査区分.経営分析調査_肥育豚生産費, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010"}} _
            , {営農経営体区分.農業経営体, {"S000001", "S000002", "S000003", "S000004", "S000005", "S000006", "S000007", "S000008", "S000009", "S000010", "S000011", "S000012", "S000013", "S000014"}}
        }

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            Public Class Report
                'テンプレートファイル名
                Public tempFileName As String
                '帳票名
                Public reportName As String
            End Class

            ' REV_010↓
            'Public Shared リスト As New Dictionary(Of String, Report) From {
            '      {調査区分.営農類型別経営統計_個人, New Report With {.tempFileName = "農業経営_営農個人_集計結果表.xlsx", .reportName = "農業経営_営農個人_集計結果表"}} _
            '    , {調査区分.営農類型別経営統計_法人, New Report With {.tempFileName = "農業経営_営農法人_集計結果表.xlsx", .reportName = "農業経営_営農法人_集計結果表"}} _
            '    , {調査区分.米生産費統計_個別, New Report With {.tempFileName = "農業経営_米生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_米生産費（個別）_集計結果表"}} _
            '    , {調査区分.小麦生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_小麦生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.二条大麦生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_二条大麦生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.六条大麦生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_六条大麦生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.はだか麦生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_はだか麦生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.そば生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_そば生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.大豆生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_大豆生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.原料用かんしょ生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_原料用かんしょ生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.原料用ばれいしょ生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.なたね生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_なたね生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.てんさい生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_てんさい生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.さとうきび生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_さとうきび生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.米生産費統計_組織法人, New Report With {.tempFileName = "農業経営_米生産費（組織法人）_集計結果表.xlsx", .reportName = "農業経営_米生産費（組織法人）_集計結果表"}} _
            '    , {調査区分.小麦生産費統計_組織法人, New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果表.xlsx", .reportName = "農業経営_小麦生産費統計（組織法人）_集計結果表"}} _
            '    , {調査区分.大豆生産費統計_組織法人, New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果表.xlsx", .reportName = "農業経営_大豆生産費統計（組織法人）_集計結果表"}} _
            '    , {調査区分.牛乳生産費統計_個別, New Report With {.tempFileName = "農業経営_牛乳生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_牛乳生産費（個別）_集計結果表"}} _
            '    , {調査区分.子牛生産費統計_個別, New Report With {.tempFileName = "農業経営_子牛生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_子牛生産費（個別）_集計結果表"}} _
            '    , {調査区分.乳用雄育成牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.交雑種育成牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_交雑種育成牛生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.去勢若齢肥育牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.乳用雄肥育牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.交雑種肥育牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_集計結果表"}} _
            '    , {調査区分.肥育豚生産費統計_個別, New Report With {.tempFileName = "農業経営_肥育豚生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_肥育豚生産費（個別）_集計結果表"}} _
            '    , {調査区分.経営分析調査_二条大麦生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_二条大麦生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_六条大麦生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_六条大麦生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_はだか麦生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_はだか麦生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_そば生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_そば生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_原料用ばれいしょ生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_原料用ばれいしょ生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_なたね生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_なたね生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_てんさい生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_てんさい生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_さとうきび生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_さとうきび生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_牛乳生産費, New Report With {.tempFileName = "経営分析_牛乳生産費_集計結果表.xlsx", .reportName = "経営分析_牛乳生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_子牛生産費, New Report With {.tempFileName = "経営分析_子牛生産費_集計結果表.xlsx", .reportName = "経営分析_子牛生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_乳用雄育成牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果表.xlsx", .reportName = "経営分析_乳用雄育成牛生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_交雑種育成牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果表.xlsx", .reportName = "経営分析_交雑種育成牛生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_去勢若齢肥育牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果表.xlsx", .reportName = "経営分析_去勢若齢肥育牛生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_乳用雄肥育牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果表.xlsx", .reportName = "経営分析_乳用雄肥育牛生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_交雑種肥育牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果表.xlsx", .reportName = "経営分析_交雑種肥育牛生産費_集計結果表"}} _
            '    , {調査区分.経営分析調査_肥育豚生産費, New Report With {.tempFileName = "経営分析_肥育豚生産費_集計結果表.xlsx", .reportName = "経営分析_肥育豚生産費_集計結果表"}} _
            '    , {営農経営体区分.農業経営体, New Report With {.tempFileName = "農業経営_営農経営体_集計結果表.xlsx", .reportName = "農業経営_営農経営体_集計結果表"}}
            '}
            Public Shared リスト As New Dictionary(Of Tuple(Of String, String), Report) From {
                  {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_営農個人_集計結果表.xlsx", .reportName = "農業経営_営農個人_集計結果表"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_営農法人_集計結果表.xlsx", .reportName = "農業経営_営農法人_集計結果表"}} _
                , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_米生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_米生産費（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_小麦生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_二条大麦生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_六条大麦生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_はだか麦生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_そば生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_大豆生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_原料用かんしょ生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_なたね生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_てんさい生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_さとうきび生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_米生産費（組織法人）_集計結果表.xlsx", .reportName = "農業経営_米生産費（組織法人）_集計結果表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果表.xlsx", .reportName = "農業経営_小麦生産費統計（組織法人）_集計結果表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果表.xlsx", .reportName = "農業経営_大豆生産費統計（組織法人）_集計結果表"}} _
                , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_牛乳生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_牛乳生産費（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_子牛生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_子牛生産費（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_交雑種育成牛生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_肥育豚生産費（個別）_集計結果表.xlsx", .reportName = "農業経営_肥育豚生産費（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_二条大麦生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_六条大麦生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_はだか麦生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_そば生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_原料用ばれいしょ生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_なたね生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_てんさい生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果表.xlsx", .reportName = "経営分析_さとうきび生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_牛乳生産費_集計結果表.xlsx", .reportName = "経営分析_牛乳生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_子牛生産費_集計結果表.xlsx", .reportName = "経営分析_子牛生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果表.xlsx", .reportName = "経営分析_乳用雄育成牛生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果表.xlsx", .reportName = "経営分析_交雑種育成牛生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果表.xlsx", .reportName = "経営分析_去勢若齢肥育牛生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果表.xlsx", .reportName = "経営分析_乳用雄肥育牛生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果表.xlsx", .reportName = "経営分析_交雑種肥育牛生産費_集計結果表"}} _
                , {Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_肥育豚生産費_集計結果表.xlsx", .reportName = "経営分析_肥育豚生産費_集計結果表"}} _
                , {Tuple.Create(営農経営体区分.農業経営体, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_営農経営体_集計結果表.xlsx", .reportName = "農業経営_営農経営体_集計結果表"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_営農個人_集計結果表_2022.xlsx", .reportName = "農業経営_営農個人_集計結果表"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_営農法人_集計結果表_2022.xlsx", .reportName = "農業経営_営農法人_集計結果表"}} _
                , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_米生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_米生産費（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_小麦生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_二条大麦生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_六条大麦生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_はだか麦生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_そば生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_大豆生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_原料用かんしょ生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_なたね生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_てんさい生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_さとうきび生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_米生産費（組織法人）_集計結果表_2022.xlsx", .reportName = "農業経営_米生産費（組織法人）_集計結果表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果表_2022.xlsx", .reportName = "農業経営_小麦生産費統計（組織法人）_集計結果表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果表_2022.xlsx", .reportName = "農業経営_大豆生産費統計（組織法人）_集計結果表"}} _
                , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_牛乳生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_牛乳生産費（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_子牛生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_子牛生産費（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_交雑種育成牛生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_集計結果表"}} _
                , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_肥育豚生産費（個別）_集計結果表_2022.xlsx", .reportName = "農業経営_肥育豚生産費（個別）_集計結果表"}} _
                , {Tuple.Create(営農経営体区分.農業経営体, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_営農経営体_集計結果表_2022.xlsx", .reportName = "農業経営_営農経営体_集計結果表"}}
            }
            ' REV_010↑
        End Class
    End Class

    ''' <summary>
    ''' 集計結果検討表クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 集計結果検討表

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            Public Class Report
                'テンプレートファイル名
                Public tempFileName As String
                '帳票名
                Public reportName As String
            End Class

            ' REV_010↓
            'Public Shared リスト As New Dictionary(Of String, Report) From {
            '      {調査区分.米生産費統計_個別, New Report With {.tempFileName = "農業経営_米生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_米生産費（個別）_集計結果検討表"}} _
            '    , {調査区分.小麦生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_小麦生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.二条大麦生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_二条大麦生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.六条大麦生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_六条大麦生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.はだか麦生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_はだか麦生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.そば生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_そば生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.大豆生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_大豆生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.原料用かんしょ生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_原料用かんしょ生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.原料用ばれいしょ生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.なたね生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_なたね生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.てんさい生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_てんさい生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.さとうきび生産費統計_個別, New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_さとうきび生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.米生産費統計_組織法人, New Report With {.tempFileName = "農業経営_米生産費（組織法人）_集計結果検討表.xlsx", .reportName = "農業経営_米生産費（組織法人）_集計結果検討表"}} _
            '    , {調査区分.小麦生産費統計_組織法人, New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果検討表.xlsx", .reportName = "農業経営_小麦生産費統計（組織法人）_集計結果検討表"}} _
            '    , {調査区分.大豆生産費統計_組織法人, New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果検討表.xlsx", .reportName = "農業経営_大豆生産費統計（組織法人）_集計結果検討表"}} _
            '    , {調査区分.牛乳生産費統計_個別, New Report With {.tempFileName = "農業経営_牛乳生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_牛乳生産費（個別）_集計結果検討表"}} _
            '    , {調査区分.子牛生産費統計_個別, New Report With {.tempFileName = "農業経営_子牛生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_子牛生産費（個別）_集計結果検討表"}} _
            '    , {調査区分.乳用雄育成牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.交雑種育成牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_交雑種育成牛生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.去勢若齢肥育牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.乳用雄肥育牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.交雑種肥育牛生産費統計_個別, New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_集計結果検討表"}} _
            '    , {調査区分.肥育豚生産費統計_個別, New Report With {.tempFileName = "農業経営_肥育豚生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_肥育豚生産費（個別）_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_二条大麦生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_二条大麦生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_六条大麦生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_六条大麦生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_はだか麦生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_はだか麦生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_そば生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_そば生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_原料用ばれいしょ生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_原料用ばれいしょ生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_なたね生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_なたね生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_てんさい生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_てんさい生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_さとうきび生産費, New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_さとうきび生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_牛乳生産費, New Report With {.tempFileName = "経営分析_牛乳生産費_集計結果検討表.xlsx", .reportName = "経営分析_牛乳生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_子牛生産費, New Report With {.tempFileName = "経営分析_子牛生産費_集計結果検討表.xlsx", .reportName = "経営分析_子牛生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_乳用雄育成牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果検討表.xlsx", .reportName = "経営分析_乳用雄育成牛生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_交雑種育成牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果検討表.xlsx", .reportName = "経営分析_交雑種育成牛生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_去勢若齢肥育牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果検討表.xlsx", .reportName = "経営分析_去勢若齢肥育牛生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_乳用雄肥育牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果検討表.xlsx", .reportName = "経営分析_乳用雄肥育牛生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_交雑種肥育牛生産費, New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果検討表.xlsx", .reportName = "経営分析_交雑種肥育牛生産費_集計結果検討表"}} _
            '    , {調査区分.経営分析調査_肥育豚生産費, New Report With {.tempFileName = "経営分析_肥育豚生産費_集計結果検討表.xlsx", .reportName = "経営分析_肥育豚生産費_集計結果検討表"}}
            '}
            Public Shared リスト As New Dictionary(Of Tuple(Of String, String), Report) From {
                  {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_米生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_米生産費（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_小麦生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_二条大麦生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_六条大麦生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_はだか麦生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_そば生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_大豆生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_原料用かんしょ生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_なたね生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_てんさい生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_さとうきび生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_米生産費（組織法人）_集計結果検討表.xlsx", .reportName = "農業経営_米生産費（組織法人）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果検討表.xlsx", .reportName = "農業経営_小麦生産費統計（組織法人）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果検討表.xlsx", .reportName = "農業経営_大豆生産費統計（組織法人）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_牛乳生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_牛乳生産費（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_子牛生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_子牛生産費（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_交雑種育成牛生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "農業経営_肥育豚生産費（個別）_集計結果検討表.xlsx", .reportName = "農業経営_肥育豚生産費（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_二条大麦生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_二条大麦生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_六条大麦生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_六条大麦生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_はだか麦生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_はだか麦生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_そば生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_そば生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_原料用ばれいしょ生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_なたね生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_なたね生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_てんさい生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_てんさい生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_さとうきび生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_麦類・そば・なたね・畑作物生産費_集計結果検討表.xlsx", .reportName = "経営分析_さとうきび生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_牛乳生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_牛乳生産費_集計結果検討表.xlsx", .reportName = "経営分析_牛乳生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_子牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_子牛生産費_集計結果検討表.xlsx", .reportName = "経営分析_子牛生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_乳用雄育成牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果検討表.xlsx", .reportName = "経営分析_乳用雄育成牛生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_交雑種育成牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果検討表.xlsx", .reportName = "経営分析_交雑種育成牛生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_去勢若齢肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果検討表.xlsx", .reportName = "経営分析_去勢若齢肥育牛生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_乳用雄肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果検討表.xlsx", .reportName = "経営分析_乳用雄肥育牛生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_交雑種肥育牛生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_育成牛・肥育牛生産費_集計結果検討表.xlsx", .reportName = "経営分析_交雑種肥育牛生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.経営分析調査_肥育豚生産費, ComConst.バージョン区分.結果表等項目2021), New Report With {.tempFileName = "経営分析_肥育豚生産費_集計結果検討表.xlsx", .reportName = "経営分析_肥育豚生産費_集計結果検討表"}} _
                , {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_米生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_米生産費（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_小麦生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_二条大麦生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_六条大麦生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_はだか麦生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_そば生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_大豆生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_原料用かんしょ生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_なたね生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_てんさい生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_さとうきび生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_米生産費（組織法人）_集計結果検討表_2022.xlsx", .reportName = "農業経営_米生産費（組織法人）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果検討表_2022.xlsx", .reportName = "農業経営_小麦生産費統計（組織法人）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果検討表_2022.xlsx", .reportName = "農業経営_大豆生産費統計（組織法人）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_牛乳生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_牛乳生産費（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_子牛生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_子牛生産費（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_交雑種育成牛生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_肥育豚生産費（個別）_集計結果検討表_2022.xlsx", .reportName = "農業経営_肥育豚生産費（個別）_集計結果検討表"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_個人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_営農個人_集計結果検討表_2022.xlsx", .reportName = "農業経営_営農個人_集計結果検討表"}} _
                , {Tuple.Create(調査区分.営農類型別経営統計_法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_営農法人_集計結果検討表_2022.xlsx", .reportName = "農業経営_営農法人_集計結果検討表"}} _
                , {Tuple.Create(営農経営体区分.農業経営体, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_営農経営体_集計結果検討表_2022.xlsx", .reportName = "農業経営_営農経営体_集計結果検討表"}}
            }
            ' REV_010↑
        End Class
    End Class

    ' REV_011↓
    ''' <summary>
    ''' 集計結果検討表_報告用クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 集計結果検討表_報告用

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            Public Class Report
                'テンプレートファイル名
                Public tempFileName As String
                '帳票名
                Public reportName As String
            End Class

            Public Shared リスト As New Dictionary(Of Tuple(Of String, String), Report) From {
                  {Tuple.Create(調査区分.米生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_米生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_米生産費（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_小麦生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.二条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_二条大麦生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.六条大麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_六条大麦生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.はだか麦生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_はだか麦生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.そば生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_そば生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_大豆生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.原料用かんしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_原料用かんしょ生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.原料用ばれいしょ生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_原料用ばれいしょ生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.なたね生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_なたね生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.てんさい生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_てんさい生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.さとうきび生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_麦類・大豆・そば・なたね・畑作物生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_さとうきび生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.米生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_米生産費（組織法人）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_米生産費（組織法人）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.小麦生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_小麦生産費統計（組織法人）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.大豆生産費統計_組織法人, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_小麦・大豆生産費（組織法人）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_大豆生産費統計（組織法人）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.牛乳生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_牛乳生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_牛乳生産費（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.子牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_子牛生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_子牛生産費（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.乳用雄育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_乳用雄育成牛生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.交雑種育成牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_交雑種育成牛生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_去勢若齢肥育牛生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.乳用雄肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_乳用雄肥育牛生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.交雑種肥育牛生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_育成牛・肥育牛生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_交雑種肥育牛生産費統計（個別）_集計結果検討表＿報告用"}} _
                , {Tuple.Create(調査区分.肥育豚生産費統計_個別, ComConst.バージョン区分.結果表等項目2022), New Report With {.tempFileName = "農業経営_肥育豚生産費（個別）_集計結果検討表＿報告用.xlsx", .reportName = "農業経営_肥育豚生産費（個別）_集計結果検討表＿報告用"}}
            }
        End Class
    End Class
    ' REV_011↑

    ''' <summary>
    ''' 調査票審査論理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 調査票審査論理
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "調査票審査論理（基本・追加）.xlsm"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 調査票審査論理.入力用ファイル名称
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "調査票審査論理（{0}）"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "審査論理"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 5
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 1004
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 調査票審査論理.出力用ファイル名称.Row.Last - 調査票審査論理.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "G"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "エラーサイン"} _
                , {2, "チェック項目名"} _
                , {3, "エラー内容"} _
                , {4, "エラーとなる条件"} _
                , {5, "繰り返し"} _
                , {6, "エラー区分"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 個別結果表作成論理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 個別結果表作成論理
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "個別結果表作成論理.xlsm"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 個別結果表作成論理.入力用ファイル名称
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "個別結果表作成論理"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "作成論理"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 5
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 3004
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 個別結果表作成論理.出力用ファイル名称.Row.Last - 個別結果表作成論理.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "H"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "項目番号"} _
                , {2, "項目名"} _
                , {3, "優先順位"} _
                , {4, "計算式"} _
                , {5, "表示単位"} _
                , {6, "再計算区分"} _
                , {7, "備考"}
            }
        End Class

        ''' <summary>表示単位クラス</summary>
        Public Class 表示単位
            Public Const 書式形式１ As Decimal = 9D
            Public Const 書式形式２ As Decimal = 9.9D
            Public Const 書式形式３ As Decimal = 9.99D
            Public Const 書式形式４ As Decimal = 9.999D

            Public Shared リスト As New Dictionary(Of Decimal, String) From {
                  {表示単位.書式形式１, "#,0"} _
                , {表示単位.書式形式２, "#,0.0"} _
                , {表示単位.書式形式３, "#,0.00"} _
                , {表示単位.書式形式４, "#,0.000"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 個別結果表作成論理_営農個人クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 個別結果表作成論理_営農個人
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "個別結果表作成論理＿営農個人.xlsm"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 個別結果表作成論理_営農個人.入力用ファイル名称
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "個別結果表作成論理"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "作成論理"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 5
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 3004
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 個別結果表作成論理_営農個人.出力用ファイル名称.Row.Last - 個別結果表作成論理_営農個人.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "I"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "項目番号"} _
                , {2, "貸借対照表区分"} _
                , {3, "項目名"} _
                , {4, "優先順位"} _
                , {5, "計算式"} _
                , {6, "表示単位"} _
                , {7, "再計算区分"} _
                , {8, "備考"}
            }
        End Class

        ''' <summary>表示単位クラス</summary>
        Public Class 表示単位
            Public Const 書式形式１ As Decimal = 9D
            Public Const 書式形式２ As Decimal = 9.9D
            Public Const 書式形式３ As Decimal = 9.99D
            Public Const 書式形式４ As Decimal = 9.999D

            Public Shared リスト As New Dictionary(Of Decimal, String) From {
                  {表示単位.書式形式１, "#,0"} _
                , {表示単位.書式形式２, "#,0.0"} _
                , {表示単位.書式形式３, "#,0.00"} _
                , {表示単位.書式形式４, "#,0.000"}
            }
        End Class

        ''' <summary>
        ''' 貸借対照表区分クラス
        ''' </summary>
        ''' <remarks></remarks>
        Public Class 貸借対照表区分
            Public Const _0 As String = "0"
            Public Const _1 As String = ComConst.貸借対照表区分._1
            Public Const _2 As String = ComConst.貸借対照表区分._2
            Public Const _3 As String = ComConst.貸借対照表区分._3

            Public Shared リスト As New Dictionary(Of String, String) From {
                  {貸借対照表区分._0, "0"} _
                , {貸借対照表区分._1, ComConst.貸借対照表区分.リスト(ComConst.貸借対照表区分._1)} _
                , {貸借対照表区分._2, ComConst.貸借対照表区分.リスト(ComConst.貸借対照表区分._2)} _
                , {貸借対照表区分._3, ComConst.貸借対照表区分.リスト(ComConst.貸借対照表区分._3)}
            }
        End Class
    End Class

    ''' <summary>
    ''' 個別結果検討表作成論理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 個別結果検討表作成論理
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "個別結果検討表作成論理.xlsm"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 個別結果検討表作成論理.入力用ファイル名称
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "個別結果検討表作成論理"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "作成論理"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 5
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 1004
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 個別結果検討表作成論理.出力用ファイル名称.Row.Last - 個別結果検討表作成論理.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "G"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "項目番号"} _
                , {2, "項目名"} _
                , {3, "優先順位"} _
                , {4, "計算式"} _
                , {5, "表示単位"} _
                , {6, "備考"}
            }
        End Class

        ''' <summary>表示単位クラス</summary>
        Public Class 表示単位
            Public Const 書式形式１ As Decimal = 9D
            Public Const 書式形式２ As Decimal = 9.9D
            Public Const 書式形式３ As Decimal = 9.99D
            Public Const 書式形式４ As Decimal = 9.999D

            Public Shared リスト As New Dictionary(Of Decimal, String) From {
                  {表示単位.書式形式１, "#,0"} _
                , {表示単位.書式形式２, "#,0.0"} _
                , {表示単位.書式形式３, "#,0.00"} _
                , {表示単位.書式形式４, "#,0.000"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 個別結果表審査論理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 個別結果表審査論理
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "個別結果表審査論理（基本・追加）.xlsm"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 個別結果表審査論理.入力用ファイル名称
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "個別結果表審査論理（{0}）"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "審査論理"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 5
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 1004
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 個別結果表審査論理.出力用ファイル名称.Row.Last - 個別結果表審査論理.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "F"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "エラーサイン"} _
                , {2, "チェック項目名"} _
                , {3, "エラー内容"} _
                , {4, "エラーとなる条件"} _
                , {5, "エラー区分"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 個別結果表審査論理範囲クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 個別結果表審査論理範囲
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "個別結果表審査論理（範囲）.xlsm"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 個別結果表審査論理範囲.入力用ファイル名称
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "個別結果表審査論理（範囲）"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "審査論理"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 6
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 1005
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 個別結果表審査論理範囲.出力用ファイル名称.Row.Last - 個別結果表審査論理範囲.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "C"
                ''' <summary>最終列</summary>
                Public Const Last As String = "H"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "チェック項目名"} _
                , {2, "項目番号１"} _
                , {3, "項目番号２"} _
                , {4, "値"} _
                , {5, "下限"} _
                , {6, "上限"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 個別結果表項目指定修正クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 個別結果表項目指定修正
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "個別結果表項目指定修正.xlsm"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 個別結果表項目指定修正.入力用ファイル名称
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "項目指定修正"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 5
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 5004
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 個別結果表項目指定修正.出力用ファイル名称.Row.Last - 個別結果表項目指定修正.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "G"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "No"} _
                , {2, "調査年（産）"} _
                , {3, "センサス番号"} _
                , {4, "項目番号"} _
                , {5, "修正前データ"} _
                , {6, "修正データ"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 個別結果表基本・追加チェックリストクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 個別結果表基本追加チェックリスト
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "個別結果表チェックリスト（基本・追加）.xlsx"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 個別結果表基本追加チェックリスト.入力用ファイル名称
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "個別結果表{0}チェックリスト"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "チェックリスト"
            ''' <summary>調査年欄</summary>
            Public Const ChosaNen As String = "B1"
            ''' <summary>調査区分欄</summary>
            Public Const Chosakubun As String = "B2"
            ''' <summary>拠点欄</summary>
            '--- REV.001 MOD START
            'Public Const Kyoten As String = "M3"
            Public Const Kyoten As String = "N3"
            '--- REV.001 MOD END
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 6
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 100005
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 個別結果表基本追加チェックリスト.出力用ファイル名称.Row.Last - 個別結果表基本追加チェックリスト.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                '--- REV.001 MOD START
                'Public Const Last As String = "M"
                Public Const Last As String = "N"
                '--- REV.001 MOD END
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "No"} _
                , {2, "都道府県"} _
                , {3, "市区町村"} _
                , {4, "旧市区町村"} _
                , {5, "農業集落"} _
                , {6, "調査区"} _
                , {7, "客体番号"} _
                , {8, "エラーサイン"} _
                , {9, "チェック項目名"} _
                , {10, "エラー内容"} _
                , {11, "エラーとなる条件"} _
                , {12, "エラーとなる要因"} _
                , {13, "区分"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 個別結果表範囲チェックリストクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 個別結果表範囲チェックリスト

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Const tempFileName As String = "個別結果表範囲チェックリスト.xlsx"
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "個別結果表範囲チェックリスト"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "チェックリスト"
            ''' <summary>調査年欄</summary>
            Public Const ChosaNen As String = "B1"
            ''' <summary>調査区分欄</summary>
            Public Const Chosakubun As String = "B2"
            ''' <summary>拠点欄</summary>
            '--- REV.001 MOD START
            'Public Const Kyoten As String = "Q3"
            Public Const Kyoten As String = "R3"
            '--- REV.001 MOD END
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 6
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                '--- REV.001 MOD START
                'Public Const Last As String = "Q"
                Public Const Last As String = "R"
                '--- REV.001 MOD END
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "No"} _
                , {2, "都道府県"} _
                , {3, "市区町村"} _
                , {4, "旧市区町村"} _
                , {5, "農業集落"} _
                , {6, "調査区"} _
                , {7, "客体番号"} _
                , {8, "チェック項目名"} _
                , {9, "下限"} _
                , {10, "上限"} _
                , {11, "値③"} _
                , {12, "項目①"} _
                , {13, "項目②"} _
                , {14, "範囲チェック結果①"} _
                , {15, "範囲チェック結果②"} _
                , {16, "単収等"} _
                , {17, "エラーとなる要因"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 集計結果表作成論理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 集計結果表作成論理
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "集計結果表作成論理.xlsm"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 集計結果表作成論理.入力用ファイル名称
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "集計結果表作成論理"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "作成論理"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 5
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 3004
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 集計結果表作成論理.出力用ファイル名称.Row.Last - 集計結果表作成論理.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "H"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "項目番号"} _
                , {2, "生産費平均値種類"} _
                , {3, "項目名"} _
                , {4, "優先順位"} _
                , {5, "計算式"} _
                , {6, "表示単位"} _
                , {7, "備考"}
            }
        End Class

        ''' <summary>表示単位クラス</summary>
        Public Class 表示単位
            Public Const 書式形式１ As Decimal = 9D
            Public Const 書式形式２ As Decimal = 9.9D
            Public Const 書式形式３ As Decimal = 9.99D
            Public Const 書式形式４ As Decimal = 9.999D

            Public Shared リスト As New Dictionary(Of Decimal, String) From {
                  {表示単位.書式形式１, "#,0"} _
                , {表示単位.書式形式２, "#,0.0"} _
                , {表示単位.書式形式３, "#,0.00"} _
                , {表示単位.書式形式４, "#,0.000"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 集計結果検討表作成論理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 集計結果検討表作成論理
        ' REV_011↓
        ''' <summary>論理種別</summary>
        Public Enum 論理種別
            集計結果検討表 = 1
            集計結果検討表_報告用
        End Enum

        ''' <summary>入力用ファイル名称</summary>
        'Public Const 入力用ファイル名称 As String = "集計結果検討表作成論理.xlsm"
        Public Shared 入力用ファイル名称 As New Dictionary(Of Integer, String) From {
              {論理種別.集計結果検討表, "集計結果検討表作成論理.xlsm"} _
            , {論理種別.集計結果検討表_報告用, "集計結果検討表＿報告用＿作成論理.xlsm"}
        }
        ' REV_011↑

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ' REV_011↓
            ''' <summary>テンプレートファイル名</summary>
            'Public Shared tempFileName As String = 集計結果検討表作成論理.入力用ファイル名称
            Public Shared tempFileName As New Dictionary(Of Integer, String) From {
              {論理種別.集計結果検討表, 集計結果検討表作成論理.入力用ファイル名称(論理種別.集計結果検討表)} _
            , {論理種別.集計結果検討表_報告用, 集計結果検討表作成論理.入力用ファイル名称(論理種別.集計結果検討表)}
            }
            ''' <summary>帳票名</summary>
            'Public Const reportName As String = "集計結果検討表作成論理"
            Public Shared reportName As New Dictionary(Of Integer, String) From {
              {論理種別.集計結果検討表, "集計結果検討表作成論理"} _
            , {論理種別.集計結果検討表_報告用, "集計結果検討表＿報告用＿作成論理"}
            }
            ' REV_011↑
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "作成論理"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 5
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 1004
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 集計結果検討表作成論理.出力用ファイル名称.Row.Last - 集計結果検討表作成論理.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "G"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "項目番号"} _
                , {2, "項目名"} _
                , {3, "優先順位"} _
                , {4, "計算式"} _
                , {5, "表示単位"} _
                , {6, "備考"}
            }
        End Class

        ''' <summary>表示単位クラス</summary>
        Public Class 表示単位
            Public Const 書式形式１ As Decimal = 9D
            Public Const 書式形式２ As Decimal = 9.9D
            Public Const 書式形式３ As Decimal = 9.99D
            Public Const 書式形式４ As Decimal = 9.999D

            Public Shared リスト As New Dictionary(Of Decimal, String) From {
                  {表示単位.書式形式１, "#,0"} _
                , {表示単位.書式形式２, "#,0.0"} _
                , {表示単位.書式形式３, "#,0.00"} _
                , {表示単位.書式形式４, "#,0.000"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 任意帳票クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 任意帳票
        ''' <summary>
        ''' 編成パラメータクラス
        ''' </summary>
        ''' <remarks></remarks>
        Public Class 編成パラメータ
            ''' <summary>タイトル欄</summary>
            Public Const title As String = "A1"
            ''' <summary>帳票レイアウト欄</summary>
            Public Const layout As String = "B4"
            ''' <summary>開始位置セル</summary>
            Public Const startCell As String = "B7"
            ''' <summary>生産費平均値種類欄</summary>
            Public Const seisanhiHeikinShurui As String = "B10"
            ''' <summary>田畑区分欄</summary>
            Public Const tahata As String = "B13"
            ''' <summary>ビール麦販売区分欄</summary>
            Public Const beerMugiHanbai As String = "B16"
            ''' <summary>てんさい栽培区分欄</summary>
            Public Const tensaiSaibai As String = "B19"
            ''' <summary>表頭・表側欄</summary>
            Public Const hyotoHyosoku As String = "B22"
            ''' <summary>項目番号欄</summary>
            Public Const itemNo As String = "C27: C226"
            ''' <summary>指標部欄</summary>
            Public Const shihyobu As String = "D27:E226"
        End Class

        ''' <summary>タイトル文字列</summary>
        Public Const TITLE As String = "任意帳票出力再編成パラメータ"

        ''' <summary>表頭・表側クラス</summary>
        Public Class 表頭_表側
            Public Enum enm
                表頭 = 1
                表側
            End Enum

            Public Shared リスト As New Dictionary(Of Integer, String) From {
                  {表頭_表側.enm.表頭, "表頭"} _
                , {表頭_表側.enm.表側, "表側"}
            }
        End Class

        ''' <summary>項目番号クラス</summary>
        Public Class 項目番号
            ''' <summary>タイトル文字列</summary>
            Public Const 前 As Char = "前"c
            Public Const SPACE As String = "SPACE"
        End Class

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "任意帳票"
            ''' <summary>置換文字</summary>
            Public Const replaceChr As String = "x"
            ''' <summary>置換空白</summary>
            Public Const replaceBlank As String = "-"
        End Class
    End Class

    ''' <summary>
    ''' 毎月勤労統計クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 毎月勤労統計
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "毎勤データ入力・修正.xlsm"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 毎月勤労統計.入力用ファイル名称
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "毎月勤労統計の月別データ"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "毎月勤労統計"
            ''' <summary>都道府県欄</summary>
            Public Const Todofuken As String = "V4"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 9
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 128
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 毎月勤労統計.出力用ファイル名称.Row.Last - 毎月勤労統計.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "A"
                ''' <summary>最終列</summary>
                Public Const Last As String = "V"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "年"} _
                , {2, "月"} _
                , {3, "男＿製造業＿現金給与総額"} _
                , {4, "男＿製造業＿労働時間"} _
                , {5, "男＿製造業＿常用労働者数"} _
                , {6, "男＿建設業＿現金給与総額"} _
                , {7, "男＿建設業＿労働時間"} _
                , {8, "男＿建設業＿常用労働者数"} _
                , {9, "男＿運輸業郵便業＿現金給与総額"} _
                , {10, "男＿運輸業郵便業＿労働時間"} _
                , {11, "男＿運輸業郵便業＿常用労働者数"} _
                , {12, "女＿製造業＿現金給与総額"} _
                , {13, "女＿製造業＿労働時間"} _
                , {14, "女＿製造業＿常用労働者数"} _
                , {15, "女＿建設業＿現金給与総額"} _
                , {16, "女＿建設業＿労働時間"} _
                , {17, "女＿建設業＿常用労働者数"} _
                , {18, "女＿運輸業郵便業＿現金給与総額"} _
                , {19, "女＿運輸業郵便業＿労働時間"} _
                , {20, "女＿運輸業郵便業＿常用労働者数"} _
                , {21, "男女平均＿全産業計＿賃金対比"} _
                , {22, "男女平均＿全産業計＿労働時間対比"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 労賃単価クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 労賃単価

        ''' <summary>期間上限</summary>
        Public Const TERM_UPPER As String = "upper"
        ''' <summary>期間下限</summary>
        Public Const TERM_LOWER As String = "lower"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Const tempFileName As String = "労賃単価結果表.xlsx"
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "労賃単価結果表"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "労賃単価結果表"
            ''' <summary>生産費欄</summary>
            Public Const Seisanhi As String = "D3"
            ''' <summary>調査年欄</summary>
            Public Const ChosaNen As String = "D5"
            ''' <summary>都道府県欄</summary>
            Public Const Todofuken As String = "D7"

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of String, String) From {
                  {"男＿５～２９人＿本年値", "C10"} _
                , {"男＿５～２９人＿前年値", "C11"} _
                , {"男＿５～２９人＿対前年比", "C12"} _
                , {"女＿５～２９人＿本年値", "E10"} _
                , {"女＿５～２９人＿前年値", "E11"} _
                , {"女＿５～２９人＿対前年比", "E12"} _
                , {"男女平均＿５～２９人＿本年値", "G10"} _
                , {"男女平均＿５～２９人＿前年値", "G11"} _
                , {"男女平均＿５～２９人＿対前年比", "G12"} _
                , {"男女平均＿採用単価＿本年値", "G13"} _
                , {"男女平均＿採用単価＿前年値", "G14"} _
                , {"男女平均＿採用単価＿対前年比", "G15"} _
                , {"男女平均＿通勤手当割合", "G16"} _
                , {"男女平均＿評価単価", "G17"} _
                , {"男女平均＿３０人以上＿産業計対比", "G19"} _
                , {"男女平均＿３０人以上＿対比単価", "G20"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 専門調査員及び担当調査客体一覧クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 専門調査員及び担当調査客体一覧

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = "専門調査員及び担当調査客体一覧.xlsx"
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "専門調査員及び担当調査客体一覧"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "専門調査員及び担当調査客体一覧"
            ''' <summary>実査設置拠点欄</summary>
            Public Const Kyoten As String = "H4"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 7
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "A"
                ''' <summary>最終列</summary>
                Public Const Last As String = "J"
            End Class
        End Class
    End Class

    ''' <summary>
    ''' エラーチェック種別クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class エラーチェック種別
        Public Enum enm
            基本 = 1
            追加 = 2
            範囲 = 3 'REV004 Add
        End Enum


        Public Shared リスト As New Dictionary(Of Integer, String) From {
              {enm.基本, "基本"} _
            , {enm.追加, "追加"} _
            , {enm.範囲, "範囲"}
          }
    End Class

    ''' <summary>
    ''' CSVファイルクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CSVファイル

        ''' <summary>開始付加文字</summary>
        Public Const START_ADDITION As String = """"
        ''' <summary>終了付加文字</summary>
        Public Const END_ADDITION As String = START_ADDITION
        ''' <summary>CSV区切文字</summary>
        Public Const CSV_DELIMITER As String = ","
        ''' <summary>コードページ_Shift_JIS</summary>
        Public Const CODEPAGE_SHIFT_JIS As String = "Shift_JIS"

        ''' <summary>ファイルダイアログCSVファイルフィルタ</summary>
        Public Const FILEDIALOG_CSV_FILEFILTER As String = "CSVファイル (*.csv)|*.csv"

        ''' <summary>出力処理実行結果</summary>
        Public Enum enmOutputReturn
            CANCEL
            OK
            ERR_SAVEAS
        End Enum
    End Class

    ''' <summary>
    ''' 中間集計表クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 中間集計表

        ''' <summary>
        ''' 中間集計表生産費区分クラス
        ''' </summary>
        ''' <remarks></remarks>
        Public Class 中間集計表生産費区分
            Public Const なたね As String = "75"
            Public Const そば As String = "76"
            Public Const 二条大麦 As String = "77"
            Public Const 六条大麦 As String = "78"
            Public Const はだか麦 As String = "79"
            Public Const 米 As String = "80"
            Public Const 小麦 As String = "81"
            Public Const 大豆 As String = "82"
            Public Const 原料用かんしょ As String = "83"
            Public Const 原料用ばれいしょ As String = "84"
            Public Const さとうきび As String = "85"
            Public Const てんさい As String = "86"
            Public Const 牛乳 As String = "90"
            Public Const 肉用子牛 As String = "91"
            Public Const 乳用雄育成牛 As String = "92"
            Public Const 交雑種育成牛 As String = "93"
            Public Const 去勢若齢肥育牛 As String = "94"
            Public Const 乳用雄肥育牛 As String = "95"
            Public Const 交雑種肥育牛 As String = "96"
            Public Const 肥育豚 As String = "97"

        End Class

        ''' <summary>
        ''' 中間集計表生産費区分クラス
        ''' </summary>
        ''' <remarks></remarks>
        Public Class 中間集計表生産費区分名称
            Public Const なたね As String = "なたね"
            Public Const そば As String = "そば"
            Public Const 二条大麦 As String = "二条大麦"
            Public Const 六条大麦 As String = "六条大麦"
            Public Const はだか麦 As String = "はだか麦"
            Public Const 米 As String = "米"
            Public Const 小麦 As String = "小麦"
            Public Const 大豆 As String = "大豆"
            Public Const 原料用かんしょ As String = "原料用かんしょ"
            Public Const 原料用ばれいしょ As String = "原料用ばれいしょ"
            Public Const さとうきび As String = "さとうきび"
            Public Const てんさい As String = "てんさい"
            Public Const 牛乳 As String = "牛乳"
            Public Const 肉用子牛 As String = "肉用子牛"
            Public Const 乳用雄育成牛 As String = "乳用雄育成牛"
            Public Const 交雑種育成牛 As String = "交雑種育成牛"
            Public Const 去勢若齢肥育牛 As String = "去勢若齢肥育牛"
            Public Const 乳用雄肥育牛 As String = "乳用雄肥育牛"
            Public Const 交雑種肥育牛 As String = "交雑種肥育牛"
            Public Const 肥育豚 As String = "肥育豚"

            ''' <summary>生産費区分名称</summary>
            Public Shared Seisanhi_NAME As String

        End Class

        ''' <summary>
        ''' 入力用シート取込用クラス
        ''' </summary>
        Public Class 入力用シート取込
            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "BN"
                ''' <summary>先頭列（数値）</summary>
                Public Const numFirst As Integer = 2
                ''' <summary>最終列（数値）</summary>
                Public Const numLast As Integer = 66
            End Class

            ''' <summary>列</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 11
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 10310
            End Class

            Public Enum 出力内容区分
                品名 = 1
                数量
                金額
                負担割合
                単位
                取引糖度
                自給牧草負担割合
                自給牧草数量
                自給牧草金額
                'REV 020↓追加項目
                単価
                自給飼料コード
                肥育豚自給飼料コード
                自給項目単位
                自給項目総数量
                利子農具金額
                'REV 020↑追加項目
            End Enum

            'REV 020↓単価～利子農具金額 項目追加
            Public Shared 出力内容列 As New Dictionary(Of Integer, Integer()) From {
                  {出力内容区分.品名, {4}} _
                , {出力内容区分.数量, {57, 59, 61}} _
                , {出力内容区分.金額, {58, 60, 62}} _
                , {出力内容区分.負担割合, {44, 45, 46}} _
                , {出力内容区分.単位, {15}} _
                , {出力内容区分.取引糖度, {18}} _
                , {出力内容区分.自給牧草負担割合, {47}} _
                , {出力内容区分.自給牧草数量, {63}} _
                , {出力内容区分.自給牧草金額, {64}} _
                , {出力内容区分.単価, {38}} _
                , {出力内容区分.自給飼料コード, {7}} _
                , {出力内容区分.肥育豚自給飼料コード, {4}} _
                , {出力内容区分.自給項目単位, {14}} _
                , {出力内容区分.自給項目総数量, {57, 59, 61}} _
                , {出力内容区分.利子農具金額, {58, 60, 62}}
            }
            'REV 020↑単価～利子農具金額 項目追加

            Public Enum 収支区分
                収入 = 1
                支出
            End Enum

        End Class

        ''' <summary>
        ''' 登録削除区分種別クラス
        ''' </summary>
        ''' <remarks></remarks>
        Public Class 登録削除区分
            Public Enum enm
                登録 = 1
                削除
            End Enum

        End Class
    End Class

    ''' <summary>
    ''' 欠測値クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 欠測値

        ''' <summary>営農欠測値適用平均値データテーブルの名称</summary>
        Public Const 営農欠測値適用平均値データ As String = "営農欠測値適用平均値データ"

        ''' <summary>営農欠測値平均値代入結果テーブルの名称</summary>
        Public Const 営農欠測値平均値代入結果 As String = "営農欠測値平均値代入結果"

        ''' <summary>個別結果表＿農業経営＿営農類型＿個人１＿集計用テーブルの名称</summary>
        Public Const 個別結果表＿農業経営＿営農類型＿個人１＿集計用 As String = "個別結果表＿農業経営＿営農類型＿個人１" & ComConst.個別結果表.集計用テーブル付加名称

        ''' <summary>個別結果表＿農業経営＿営農類型＿個人２＿集計用テーブルの名称</summary>
        Public Const 個別結果表＿農業経営＿営農類型＿個人２＿集計用 As String = "個別結果表＿農業経営＿営農類型＿個人２" & ComConst.個別結果表.集計用テーブル付加名称

        ''' <summary>個別結果表＿農業経営＿営農類型＿個人３＿集計用テーブルの名称</summary>
        Public Const 個別結果表＿農業経営＿営農類型＿個人３＿集計用 As String = "個別結果表＿農業経営＿営農類型＿個人３" & ComConst.個別結果表.集計用テーブル付加名称

        ''' <summary>個別結果表＿農業経営＿営農類型＿個人４＿集計用テーブルの名称</summary>
        Public Const 個別結果表＿農業経営＿営農類型＿個人４＿集計用 As String = "個別結果表＿農業経営＿営農類型＿個人４" & ComConst.個別結果表.集計用テーブル付加名称

        ''' <summary>個別結果表＿農業経営＿営農類型＿個人５＿集計用テーブルの名称</summary>
        Public Const 個別結果表＿農業経営＿営農類型＿個人５＿集計用 As String = "個別結果表＿農業経営＿営農類型＿個人５" & ComConst.個別結果表.集計用テーブル付加名称

        ''' <summary>個別結果表＿農業経営＿営農類型＿個人６＿集計用テーブルの名称</summary>
        Public Const 個別結果表＿農業経営＿営農類型＿個人６＿集計用 As String = "個別結果表＿農業経営＿営農類型＿個人６" & ComConst.個別結果表.集計用テーブル付加名称

        ''' <summary>
        ''' 農業地域コードの一覧
        ''' </summary>
        ''' <remarks>
        ''' key  ：農業地域コード
        ''' value：都道府県コード
        ''' </remarks>
        Public Shared 農業地域コード一覧 As New Dictionary(Of Integer, List(Of Integer)) From {
            {0, New List(Of Integer) From {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47}},
            {1, New List(Of Integer) From {1}},
            {61, New List(Of Integer) From {2, 3, 4, 5, 6, 7}},
            {62, New List(Of Integer) From {15, 16, 17, 18}},
            {63, New List(Of Integer) From {8, 9, 10, 11, 12, 13, 14, 19, 20}},
            {67, New List(Of Integer) From {21, 22, 23, 24}},
            {68, New List(Of Integer) From {25, 26, 27, 28, 29, 30}},
            {71, New List(Of Integer) From {31, 32, 33, 34, 35}},
            {74, New List(Of Integer) From {36, 37, 38, 39}},
            {75, New List(Of Integer) From {40, 41, 42, 43, 44, 45, 46}},
            {47, New List(Of Integer) From {47}}
        }

        ''' <summary>
        ''' 都道府県コードを基に農業地域コードを取得する
        ''' </summary>
        ''' <param name="都道府県コード"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get農業地域コード(ByVal 都道府県コード As Integer?) As Integer?

            If 都道府県コード Is Nothing Then
                Return Nothing
            End If

            ' REV_008↓
            ''農業地域コードが「全国 = 0」のデータを削除する
            'ComConst.欠測値.農業地域コード一覧.Remove(0)
            ' REV_014↑

            For Each 都道府県コード一覧 In ComConst.欠測値.農業地域コード一覧

                ' REV_008↓
                '農業地域コードが「全国 = 0」のデータは対象外とする
                If 都道府県コード一覧.Key = 0 Then
                    Continue For
                End If
                ' REV_014↑

                '都道府県コードを検索する
                Dim query = 都道府県コード一覧.Value.Where(Function(row) row = CInt(都道府県コード))

                '都道府県コードが検索できたらそれに対応する農業地域コードを返す
                If query.Count = 1 Then
                    Return 都道府県コード一覧.Key
                End If
            Next

            Return Nothing

        End Function

        ''' <summary>
        ''' 営農類型コードおよび営農規模コードの一覧
        ''' </summary>
        ''' <remarks>
        ''' key  ：営農類型コード
        ''' value：営農規模コード
        ''' </remarks>
        Public Shared 営農類型コード_営農規模コード一覧 As New Dictionary(Of Integer, List(Of Integer)) From {
            {1, New List(Of Integer) From {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}},
            {2, New List(Of Integer) From {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}},
            {3, New List(Of Integer) From {0, 1, 2, 3, 4, 5, 6, 7, 8}},
            {4, New List(Of Integer) From {0, 1, 2, 3, 4, 5, 6}},
            {5, New List(Of Integer) From {0, 1, 2, 3, 4, 5}},
            {6, New List(Of Integer) From {0, 1, 2, 3, 4}},
            {7, New List(Of Integer) From {0, 1, 2, 3, 4}},
            {8, New List(Of Integer) From {0, 1, 2, 3, 4, 5, 6}},
            {9, New List(Of Integer) From {0, 1, 2, 3, 4, 5, 6}},
            {10, New List(Of Integer) From {0, 1, 2, 3, 4, 5}},
            {11, New List(Of Integer) From {0, 1, 2, 3, 4, 5}},
            {12, New List(Of Integer) From {0, 1, 2, 3, 4, 5}},
            {13, New List(Of Integer) From {0, 1, 2, 3, 4}},
            {14, New List(Of Integer) From {0, 1, 2, 3, 4, 5}}
        }

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 適用データ一覧表出力用ファイル
            Public Class Report
                ''' <summary>テンプレートファイル名</summary>
                Public Shared tempFileName As String = "欠測値適用データ一覧表.xlsx"

                ''' <summary>帳票名</summary>
                Public Const reportName As String = "欠測値適用データ一覧表"

                ''' <summary>シート名</summary>
                Public Const SheetName As String = "欠測値適用データ一覧表"

            End Class

            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 6
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 1083
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = Last - First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "A"
                ''' <summary>最終列</summary>
                Public Const Last As String = "O"
            End Class

        End Class

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 適用状況一覧表出力用ファイル
            Public Class Report
                ''' <summary>テンプレートファイル名</summary>
                Public Shared tempFileName As String = "欠測値適用状況一覧表.xlsx"

                ''' <summary>帳票名</summary>
                Public Const reportName As String = "欠測値適用状況一覧表"

                ''' <summary>シート名</summary>
                Public Const SheetName As String = "欠測値適用状況一覧表"

            End Class

            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 5
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 2004
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = Last - First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "A"
                ''' <summary>最終列</summary>
                Public Const Last As String = "I"
            End Class

        End Class

    End Class

    '---REV.002 ADD START
    '還元資料出力クラス
    Public Class 還元資料
        Public Const お問合せ先 As String = "1"
        Public Enum お問合せ先_明細番号
            事務局_拠点 = 1
            住所
            電話番号
            FAX
            担当者名
        End Enum

        Public Const 集計結果表 As String = "2"
        Public Enum 集計結果表_明細番号
            集計結果表データ_当年_１ = 1
            集計結果表データ_当年_２
            集計結果表データ_当年_３
            集計結果表データ_前年
        End Enum

        Public Class 還元資料項目_営農個人

            Public Const 営個還2_レーダーチャート As String = "3"
            Public Enum 営個還2_レーダーチャート_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
                任意項目５
                任意項目６
                任意項目７
                任意項目８
            End Enum

            Public Const 営個還3_分析指標の経年変化 As String = "4"
            Public Enum 営個還3_分析指標の経年変化_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
            End Enum

            Public Const 営個還4_グラフ項目_面積_収入 As String = "5"
            Public Enum 営個還4_グラフ項目_面積_収入_明細番号
                面積 = 1
                収入
            End Enum

            Public Const 営個還7_生産概況内訳 As String = "6"
            Public Enum 営個還7_生産概況内訳_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
                任意項目５
                任意項目６
                任意項目７
            End Enum

            Public Const 営個還8_分析指標 As String = "7"
            Public Enum 営個還8_分析指標_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
            End Enum

            Public Const 営個還9_統計表 As String = "8"
            Public Enum 営個還9_統計表_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
                任意項目５
                任意項目６
                任意項目７
                任意項目８
                任意項目９
                任意項目１０
                任意項目１１
                任意項目１２
                任意項目１３
                任意項目１４
                任意項目１５
                任意項目１６
                任意項目１７
                任意項目１８
            End Enum


            Public Shared 明細番号一覧 As New Dictionary(Of String, List(Of Integer)) From {
                {還元資料.お問合せ先, New List(Of Integer) From {お問合せ先_明細番号.事務局_拠点, お問合せ先_明細番号.住所, お問合せ先_明細番号.電話番号, お問合せ先_明細番号.FAX, お問合せ先_明細番号.担当者名}},
                {還元資料.集計結果表, New List(Of Integer) From {集計結果表_明細番号.集計結果表データ_当年_１, 集計結果表_明細番号.集計結果表データ_当年_２, 集計結果表_明細番号.集計結果表データ_当年_３ _
                                                                             , 集計結果表_明細番号.集計結果表データ_前年}},
                {還元資料項目_営農個人.営個還2_レーダーチャート, New List(Of Integer) From {営個還2_レーダーチャート_明細番号.任意項目１, 営個還2_レーダーチャート_明細番号.任意項目２, 営個還2_レーダーチャート_明細番号.任意項目３ _
                                                                                         , 営個還2_レーダーチャート_明細番号.任意項目４, 営個還2_レーダーチャート_明細番号.任意項目５, 営個還2_レーダーチャート_明細番号.任意項目６ _
                                                                                         , 営個還2_レーダーチャート_明細番号.任意項目７, 営個還2_レーダーチャート_明細番号.任意項目８}},
                {還元資料項目_営農個人.営個還3_分析指標の経年変化, New List(Of Integer) From {営個還3_分析指標の経年変化_明細番号.任意項目１, 営個還3_分析指標の経年変化_明細番号.任意項目２, 営個還3_分析指標の経年変化_明細番号.任意項目３}},
                {還元資料項目_営農個人.営個還4_グラフ項目_面積_収入, New List(Of Integer) From {営個還4_グラフ項目_面積_収入_明細番号.面積, 営個還4_グラフ項目_面積_収入_明細番号.収入}},
                {還元資料項目_営農個人.営個還7_生産概況内訳, New List(Of Integer) From {営個還7_生産概況内訳_明細番号.任意項目１, 営個還7_生産概況内訳_明細番号.任意項目２, 営個還7_生産概況内訳_明細番号.任意項目３, 営個還7_生産概況内訳_明細番号.任意項目４ _
                                                                                      , 営個還7_生産概況内訳_明細番号.任意項目５, 営個還7_生産概況内訳_明細番号.任意項目６, 営個還7_生産概況内訳_明細番号.任意項目７}},
                {還元資料項目_営農個人.営個還8_分析指標, New List(Of Integer) From {営個還8_分析指標_明細番号.任意項目１, 営個還8_分析指標_明細番号.任意項目２, 営個還8_分析指標_明細番号.任意項目３, 営個還8_分析指標_明細番号.任意項目４}},
                {還元資料項目_営農個人.営個還9_統計表, New List(Of Integer) From {営個還9_統計表_明細番号.任意項目１, 営個還9_統計表_明細番号.任意項目２, 営個還9_統計表_明細番号.任意項目３, 営個還9_統計表_明細番号.任意項目４ _
                                                                                , 営個還9_統計表_明細番号.任意項目５, 営個還9_統計表_明細番号.任意項目６, 営個還9_統計表_明細番号.任意項目７, 営個還9_統計表_明細番号.任意項目８ _
                                                                                , 営個還9_統計表_明細番号.任意項目９, 営個還9_統計表_明細番号.任意項目１０, 営個還9_統計表_明細番号.任意項目１１, 営個還9_統計表_明細番号.任意項目１２ _
                                                                                , 営個還9_統計表_明細番号.任意項目１３, 営個還9_統計表_明細番号.任意項目１４, 営個還9_統計表_明細番号.任意項目１５, 営個還9_統計表_明細番号.任意項目１６ _
                                                                                , 営個還9_統計表_明細番号.任意項目１７, 営個還9_統計表_明細番号.任意項目１８}}
            }
        End Class

        Public Class 還元資料項目_営農法人
            Public Const 営法還2_レーダーチャート As String = "3"
            Public Enum 営法還2_レーダーチャート_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
                任意項目５
                任意項目６
                任意項目７
                任意項目８
            End Enum

            Public Const 営法還3_分析指標の経年変化 As String = "4"
            Public Enum 営法還3_分析指標の経年変化_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
            End Enum

            Public Const 営法還4_グラフ項目_面積_収入 As String = "5"
            Public Enum 営法還4_グラフ項目_面積_収入_明細番号
                面積 = 1
                収入
            End Enum

            Public Const 営法還7_生産概況内訳 As String = "6"
            Public Enum 営法還7_生産概況内訳_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
                任意項目５
                任意項目６
            End Enum

            Public Const 営法還8_分析指標 As String = "7"
            Public Enum 営法還8_分析指標_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
            End Enum

            Public Const 営法還9_統計表 As String = "8"
            Public Enum 営法還9_統計表_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
                任意項目５
                任意項目６
                任意項目７
                任意項目８
                任意項目９
                任意項目１０
                任意項目１１
                任意項目１２
                任意項目１３
                任意項目１４
                任意項目１５
                任意項目１６
                任意項目１７
                任意項目１８
                任意項目１９
                任意項目２０
                任意項目２１
                任意項目２２
                任意項目２３
                任意項目２４
                任意項目２５
                任意項目２６
                任意項目２７
                任意項目２８
                任意項目２９
            End Enum


            Public Shared 明細番号一覧 As New Dictionary(Of String, List(Of Integer)) From {
                {還元資料.お問合せ先, New List(Of Integer) From {お問合せ先_明細番号.事務局_拠点, お問合せ先_明細番号.住所, お問合せ先_明細番号.電話番号, お問合せ先_明細番号.FAX, お問合せ先_明細番号.担当者名}},
                {還元資料.集計結果表, New List(Of Integer) From {集計結果表_明細番号.集計結果表データ_当年_１, 集計結果表_明細番号.集計結果表データ_当年_２, 集計結果表_明細番号.集計結果表データ_当年_３ _
                                                                             , 集計結果表_明細番号.集計結果表データ_前年}},
                {還元資料項目_営農法人.営法還2_レーダーチャート, New List(Of Integer) From {営法還2_レーダーチャート_明細番号.任意項目１, 営法還2_レーダーチャート_明細番号.任意項目２, 営法還2_レーダーチャート_明細番号.任意項目３ _
                                                                                         , 営法還2_レーダーチャート_明細番号.任意項目４, 営法還2_レーダーチャート_明細番号.任意項目５, 営法還2_レーダーチャート_明細番号.任意項目６ _
                                                                                         , 営法還2_レーダーチャート_明細番号.任意項目７, 営法還2_レーダーチャート_明細番号.任意項目８}},
                {還元資料項目_営農法人.営法還3_分析指標の経年変化, New List(Of Integer) From {営法還3_分析指標の経年変化_明細番号.任意項目１, 営法還3_分析指標の経年変化_明細番号.任意項目２, 営法還3_分析指標の経年変化_明細番号.任意項目３}},
                {還元資料項目_営農法人.営法還4_グラフ項目_面積_収入, New List(Of Integer) From {営法還4_グラフ項目_面積_収入_明細番号.面積, 営法還4_グラフ項目_面積_収入_明細番号.収入}},
                {還元資料項目_営農法人.営法還7_生産概況内訳, New List(Of Integer) From {営法還7_生産概況内訳_明細番号.任意項目１, 営法還7_生産概況内訳_明細番号.任意項目２, 営法還7_生産概況内訳_明細番号.任意項目３, 営法還7_生産概況内訳_明細番号.任意項目４ _
                                                                                      , 営法還7_生産概況内訳_明細番号.任意項目５, 営法還7_生産概況内訳_明細番号.任意項目６}},
                {還元資料項目_営農法人.営法還8_分析指標, New List(Of Integer) From {営法還8_分析指標_明細番号.任意項目１, 営法還8_分析指標_明細番号.任意項目２, 営法還8_分析指標_明細番号.任意項目３, 営法還8_分析指標_明細番号.任意項目４}},
                {還元資料項目_営農法人.営法還9_統計表, New List(Of Integer) From {営法還9_統計表_明細番号.任意項目１, 営法還9_統計表_明細番号.任意項目２, 営法還9_統計表_明細番号.任意項目３, 営法還9_統計表_明細番号.任意項目４ _
                                                                                , 営法還9_統計表_明細番号.任意項目５, 営法還9_統計表_明細番号.任意項目６, 営法還9_統計表_明細番号.任意項目７, 営法還9_統計表_明細番号.任意項目８ _
                                                                                , 営法還9_統計表_明細番号.任意項目９, 営法還9_統計表_明細番号.任意項目１０, 営法還9_統計表_明細番号.任意項目１１, 営法還9_統計表_明細番号.任意項目１２ _
                                                                                , 営法還9_統計表_明細番号.任意項目１３, 営法還9_統計表_明細番号.任意項目１４, 営法還9_統計表_明細番号.任意項目１５, 営法還9_統計表_明細番号.任意項目１６ _
                                                                                , 営法還9_統計表_明細番号.任意項目１７, 営法還9_統計表_明細番号.任意項目１８, 営法還9_統計表_明細番号.任意項目１９, 営法還9_統計表_明細番号.任意項目２０ _
                                                                                , 営法還9_統計表_明細番号.任意項目２１, 営法還9_統計表_明細番号.任意項目２２, 営法還9_統計表_明細番号.任意項目２３, 営法還9_統計表_明細番号.任意項目２４ _
                                                                                , 営法還9_統計表_明細番号.任意項目２５, 営法還9_統計表_明細番号.任意項目２６, 営法還9_統計表_明細番号.任意項目２７, 営法還9_統計表_明細番号.任意項目２８ _
                                                                                , 営法還9_統計表_明細番号.任意項目２９}}
            }
        End Class

        Public Class 還元資料項目_生産費
            Public Const 還元4_統計表 As String = "3"
            Public Enum 還元4_統計表_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
                任意項目５
                任意項目６
                任意項目７
                任意項目８
                任意項目９
                任意項目１０
                任意項目１１
                任意項目１２
                任意項目１３
                任意項目１４
                任意項目１５
                任意項目１６
                任意項目１７
                任意項目１８
                任意項目１９
                任意項目２０
                任意項目２１
                任意項目２２
                任意項目２３
                任意項目２４
                任意項目２５
                任意項目２６
                任意項目２７
                任意項目２８
                任意項目２９
            End Enum

            Public Const 還元6_レーダーチャートの表示項目 As String = "4"
            Public Enum 還元6_レーダーチャート_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
                任意項目５
                任意項目６
                任意項目７
                任意項目８
            End Enum

            Public Const 還元8_主要項目 As String = "5"
            Public Enum 還元8_主要項目_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
                任意項目５
            End Enum

            Public Const 還元8_主要作業 As String = "6"
            Public Enum 還元8_主要作業_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
                任意項目５
            End Enum

            Public Const 還元10_物財費内訳 As String = "7"
            Public Enum 還元10_物財費内訳_明細番号
                任意項目１ = 1
                任意項目２
                任意項目３
                任意項目４
                任意項目５
            End Enum

            Public Shared 明細番号一覧 As New Dictionary(Of String, List(Of Integer)) From {
                {還元資料.お問合せ先, New List(Of Integer) From {お問合せ先_明細番号.事務局_拠点, お問合せ先_明細番号.住所, お問合せ先_明細番号.電話番号, お問合せ先_明細番号.FAX, お問合せ先_明細番号.担当者名}},
                {還元資料.集計結果表, New List(Of Integer) From {集計結果表_明細番号.集計結果表データ_当年_１, 集計結果表_明細番号.集計結果表データ_当年_２, 集計結果表_明細番号.集計結果表データ_当年_３ _
                                                                         , 集計結果表_明細番号.集計結果表データ_前年}},
                {還元資料項目_生産費.還元4_統計表, New List(Of Integer) From {還元4_統計表_明細番号.任意項目１, 還元4_統計表_明細番号.任意項目２, 還元4_統計表_明細番号.任意項目３, 還元4_統計表_明細番号.任意項目４ _
                                                                            , 還元4_統計表_明細番号.任意項目５, 還元4_統計表_明細番号.任意項目６, 還元4_統計表_明細番号.任意項目７, 還元4_統計表_明細番号.任意項目８ _
                                                                            , 還元4_統計表_明細番号.任意項目９, 還元4_統計表_明細番号.任意項目１０, 還元4_統計表_明細番号.任意項目１１, 還元4_統計表_明細番号.任意項目１２ _
                                                                            , 還元4_統計表_明細番号.任意項目１３, 還元4_統計表_明細番号.任意項目１４, 還元4_統計表_明細番号.任意項目１５, 還元4_統計表_明細番号.任意項目１６ _
                                                                            , 還元4_統計表_明細番号.任意項目１７, 還元4_統計表_明細番号.任意項目１８, 還元4_統計表_明細番号.任意項目１９, 還元4_統計表_明細番号.任意項目２０ _
                                                                            , 還元4_統計表_明細番号.任意項目２１, 還元4_統計表_明細番号.任意項目２２, 還元4_統計表_明細番号.任意項目２３, 還元4_統計表_明細番号.任意項目２４ _
                                                                            , 還元4_統計表_明細番号.任意項目２５, 還元4_統計表_明細番号.任意項目２６, 還元4_統計表_明細番号.任意項目２７, 還元4_統計表_明細番号.任意項目２８ _
                                                                            , 還元4_統計表_明細番号.任意項目２９}},
                {還元資料項目_生産費.還元6_レーダーチャートの表示項目, New List(Of Integer) From {還元6_レーダーチャート_明細番号.任意項目１, 還元6_レーダーチャート_明細番号.任意項目２, 還元6_レーダーチャート_明細番号.任意項目３ _
                                                                                     , 還元6_レーダーチャート_明細番号.任意項目４, 還元6_レーダーチャート_明細番号.任意項目５, 還元6_レーダーチャート_明細番号.任意項目６ _
                                                                                     , 還元6_レーダーチャート_明細番号.任意項目７, 還元6_レーダーチャート_明細番号.任意項目８}},
                {還元資料項目_生産費.還元8_主要項目, New List(Of Integer) From {還元8_主要項目_明細番号.任意項目１, 還元8_主要項目_明細番号.任意項目２, 還元8_主要項目_明細番号.任意項目３ _
                                                                                , 還元8_主要項目_明細番号.任意項目４, 還元8_主要項目_明細番号.任意項目５}},
                {還元資料項目_生産費.還元8_主要作業, New List(Of Integer) From {還元8_主要作業_明細番号.任意項目１, 還元8_主要作業_明細番号.任意項目２, 還元8_主要作業_明細番号.任意項目３ _
                                                                                , 還元8_主要作業_明細番号.任意項目４, 還元8_主要作業_明細番号.任意項目５}},
                {還元資料項目_生産費.還元10_物財費内訳, New List(Of Integer) From {還元10_物財費内訳_明細番号.任意項目１, 還元10_物財費内訳_明細番号.任意項目２, 還元10_物財費内訳_明細番号.任意項目３ _
                                                                                    , 還元10_物財費内訳_明細番号.任意項目４, 還元10_物財費内訳_明細番号.任意項目５}}
    }
        End Class

        'テンプレートファイル名
        Public Shared テンプレートファイル名称 As New Dictionary(Of String, String) From {
                     {ComConst.調査区分.営農類型別経営統計_個人, "11還元資料（農業経営の概況_営農個人）.xlsx"} _
                   , {ComConst.調査区分.営農類型別経営統計_法人, "12還元資料（農業経営の概況_営農法人）.xlsx"} _
                   , {ComConst.調査区分.米生産費統計_個別, "13還元資料（生産費の概況_米生産費（個別））.xlsx"} _
                   , {ComConst.調査区分.小麦生産費統計_個別, "14還元資料（生産費の概況_麦類・大豆・そば・なたね・畑作物生産費（個別）).xlsx"} _
                   , {ComConst.調査区分.二条大麦生産費統計_個別, "14還元資料（生産費の概況_麦類・大豆・そば・なたね・畑作物生産費（個別）).xlsx"} _
                   , {ComConst.調査区分.六条大麦生産費統計_個別, "14還元資料（生産費の概況_麦類・大豆・そば・なたね・畑作物生産費（個別）).xlsx"} _
                   , {ComConst.調査区分.はだか麦生産費統計_個別, "14還元資料（生産費の概況_麦類・大豆・そば・なたね・畑作物生産費（個別）).xlsx"} _
                   , {ComConst.調査区分.そば生産費統計_個別, "14還元資料（生産費の概況_麦類・大豆・そば・なたね・畑作物生産費（個別）).xlsx"} _
                   , {ComConst.調査区分.大豆生産費統計_個別, "14還元資料（生産費の概況_麦類・大豆・そば・なたね・畑作物生産費（個別）).xlsx"} _
                   , {ComConst.調査区分.原料用かんしょ生産費統計_個別, "14還元資料（生産費の概況_麦類・大豆・そば・なたね・畑作物生産費（個別）).xlsx"} _
                   , {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, "14還元資料（生産費の概況_麦類・大豆・そば・なたね・畑作物生産費（個別）).xlsx"} _
                   , {ComConst.調査区分.なたね生産費統計_個別, "14還元資料（生産費の概況_麦類・大豆・そば・なたね・畑作物生産費（個別）).xlsx"} _
                   , {ComConst.調査区分.てんさい生産費統計_個別, "14還元資料（生産費の概況_麦類・大豆・そば・なたね・畑作物生産費（個別）).xlsx"} _
                   , {ComConst.調査区分.さとうきび生産費統計_個別, "14還元資料（生産費の概況_麦類・大豆・そば・なたね・畑作物生産費（個別）).xlsx"} _
                   , {ComConst.調査区分.米生産費統計_組織法人, "19還元資料（生産費の概況_米生産費（組織法人））.xlsx"} _
                   , {ComConst.調査区分.小麦生産費統計_組織法人, "20還元資料（生産費の概況_小麦・大豆生産費（組織法人））.xlsx"} _
                   , {ComConst.調査区分.大豆生産費統計_組織法人, "20還元資料（生産費の概況_小麦・大豆生産費（組織法人））.xlsx"} _
                   , {ComConst.調査区分.牛乳生産費統計_個別, "15還元資料（生産費の概況_牛乳生産費（個別）) .xlsx"} _
                   , {ComConst.調査区分.子牛生産費統計_個別, "16還元資料（生産費の概況_子牛生産費（個別）) .xlsx"} _
                   , {ComConst.調査区分.乳用雄育成牛生産費統計_個別, "17還元資料（生産費の概況_育成牛・肥育牛生産費（個別）)  .xlsx"} _
                   , {ComConst.調査区分.交雑種育成牛生産費統計_個別, "17還元資料（生産費の概況_育成牛・肥育牛生産費（個別）)  .xlsx"} _
                   , {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, "17還元資料（生産費の概況_育成牛・肥育牛生産費（個別）)  .xlsx"} _
                   , {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, "17還元資料（生産費の概況_育成牛・肥育牛生産費（個別）)  .xlsx"} _
                   , {ComConst.調査区分.交雑種肥育牛生産費統計_個別, "17還元資料（生産費の概況_育成牛・肥育牛生産費（個別）)  .xlsx"} _
                   , {ComConst.調査区分.肥育豚生産費統計_個別, "18還元資料（生産費の概況_肥育豚生産費（個別）).xlsx"} _
                   , {ComConst.調査区分.経営分析調査_二条大麦生産費, "21還元資料（生産費の概況_麦類・そば・なたね・畑作物生産費（組織法人））.xlsx"} _
                   , {ComConst.調査区分.経営分析調査_六条大麦生産費, "21還元資料（生産費の概況_麦類・そば・なたね・畑作物生産費（組織法人））.xlsx"} _
                   , {ComConst.調査区分.経営分析調査_はだか麦生産費, "21還元資料（生産費の概況_麦類・そば・なたね・畑作物生産費（組織法人））.xlsx"} _
                   , {ComConst.調査区分.経営分析調査_そば生産費, "21還元資料（生産費の概況_麦類・そば・なたね・畑作物生産費（組織法人））.xlsx"} _
                   , {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, "21還元資料（生産費の概況_麦類・そば・なたね・畑作物生産費（組織法人））.xlsx"} _
                   , {ComConst.調査区分.経営分析調査_なたね生産費, "21還元資料（生産費の概況_麦類・そば・なたね・畑作物生産費（組織法人））.xlsx"} _
                   , {ComConst.調査区分.経営分析調査_てんさい生産費, "21還元資料（生産費の概況_麦類・そば・なたね・畑作物生産費（組織法人））.xlsx"} _
                   , {ComConst.調査区分.経営分析調査_さとうきび生産費, "21還元資料（生産費の概況_麦類・そば・なたね・畑作物生産費（組織法人））.xlsx"} _
                   , {ComConst.調査区分.経営分析調査_牛乳生産費, "22還元資料（生産費の概況_牛乳生産費（組織法人）).xlsx"} _
                   , {ComConst.調査区分.経営分析調査_子牛生産費, "23還元資料（生産費の概況_子牛生産費（組織法人）).xlsx"} _
                   , {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, "24還元資料（生産費の概況_育成牛・肥育牛生産費（組織法人）).xlsx"} _
                   , {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, "24還元資料（生産費の概況_育成牛・肥育牛生産費（組織法人）).xlsx"} _
                   , {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, "24還元資料（生産費の概況_育成牛・肥育牛生産費（組織法人）).xlsx"} _
                   , {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, "24還元資料（生産費の概況_育成牛・肥育牛生産費（組織法人）).xlsx"} _
                   , {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, "24還元資料（生産費の概況_育成牛・肥育牛生産費（組織法人）).xlsx"} _
                   , {ComConst.調査区分.経営分析調査_肥育豚生産費, "25還元資料（生産費の概況_肥育豚生産費（組織法人）).xlsx"}
               }

        '帳票名
        Public Shared 帳票名 As New Dictionary(Of String, String) From {
                     {ComConst.調査区分.営農類型別経営統計_個人, "営農類型別経営統計（個人）"} _
                   , {ComConst.調査区分.営農類型別経営統計_法人, "営農類型別経営統計（法人）"} _
                   , {ComConst.調査区分.米生産費統計_個別, "米生産費統計（個別）"} _
                   , {ComConst.調査区分.小麦生産費統計_個別, "小麦生産費統計（個別）"} _
                   , {ComConst.調査区分.二条大麦生産費統計_個別, "二条大麦生産費統計（個別）"} _
                   , {ComConst.調査区分.六条大麦生産費統計_個別, "六条大麦生産費統計（個別）"} _
                   , {ComConst.調査区分.はだか麦生産費統計_個別, "はだか麦生産費統計（個別）"} _
                   , {ComConst.調査区分.そば生産費統計_個別, "そば生産費統計（個別）"} _
                   , {ComConst.調査区分.大豆生産費統計_個別, "大豆生産費統計（個別）"} _
                   , {ComConst.調査区分.原料用かんしょ生産費統計_個別, "原料用かんしょ生産費統計（個別）"} _
                   , {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, "原料用ばれいしょ生産費統計（個別）"} _
                   , {ComConst.調査区分.なたね生産費統計_個別, "なたね生産費統計（個別）"} _
                   , {ComConst.調査区分.てんさい生産費統計_個別, "てんさい生産費統計（個別）"} _
                   , {ComConst.調査区分.さとうきび生産費統計_個別, "さとうきび生産費統計（個別）"} _
                   , {ComConst.調査区分.米生産費統計_組織法人, "米生産費統計（組織法人）"} _
                   , {ComConst.調査区分.小麦生産費統計_組織法人, "小麦生産費統計（組織法人）"} _
                   , {ComConst.調査区分.大豆生産費統計_組織法人, "大豆生産費統計（組織法人）"} _
                   , {ComConst.調査区分.牛乳生産費統計_個別, "牛乳生産費統計（個別）"} _
                   , {ComConst.調査区分.子牛生産費統計_個別, "子牛生産費統計（個別）"} _
                   , {ComConst.調査区分.乳用雄育成牛生産費統計_個別, "乳用雄育成牛生産費統計（個別）"} _
                   , {ComConst.調査区分.交雑種育成牛生産費統計_個別, "交雑種育成牛生産費統計（個別）"} _
                   , {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, "去勢若齢肥育牛生産費統計（個別）"} _
                   , {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, "乳用雄肥育牛生産費統計（個別）"} _
                   , {ComConst.調査区分.交雑種肥育牛生産費統計_個別, "交雑種肥育牛生産費統計（個別）"} _
                   , {ComConst.調査区分.肥育豚生産費統計_個別, "肥育豚生産費統計（個別）"} _
                   , {ComConst.調査区分.経営分析調査_二条大麦生産費, "経営分析調査_二条大麦生産費"} _
                   , {ComConst.調査区分.経営分析調査_六条大麦生産費, "経営分析調査_六条大麦生産費"} _
                   , {ComConst.調査区分.経営分析調査_はだか麦生産費, "経営分析調査_はだか麦生産費"} _
                   , {ComConst.調査区分.経営分析調査_そば生産費, "経営分析調査_そば生産費"} _
                   , {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, "経営分析調査_原料用ばれいしょ生産費"} _
                   , {ComConst.調査区分.経営分析調査_なたね生産費, "経営分析調査_なたね生産費"} _
                   , {ComConst.調査区分.経営分析調査_てんさい生産費, "経営分析調査_てんさい生産費"} _
                   , {ComConst.調査区分.経営分析調査_さとうきび生産費, "経営分析調査_さとうきび生産費"} _
                   , {ComConst.調査区分.経営分析調査_牛乳生産費, "経営分析調査_牛乳生産費"} _
                   , {ComConst.調査区分.経営分析調査_子牛生産費, "経営分析調査_子牛生産費"} _
                   , {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, "経営分析調査_乳用雄育成牛生産費"} _
                   , {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, "経営分析調査_交雑種育成牛生産費"} _
                   , {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, "経営分析調査_去勢若齢肥育牛生産費"} _
                   , {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, "経営分析調査_乳用雄肥育牛生産費"} _
                   , {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, "経営分析調査_交雑種肥育牛生産費"} _
                   , {ComConst.調査区分.経営分析調査_肥育豚生産費, "経営分析調査_肥育豚生産費"}
               }

        '表示名
        Public Shared 表示名 As New Dictionary(Of String, String) From {
                     {ComConst.調査区分.営農類型別経営統計_個人, "営農個人"} _
                   , {ComConst.調査区分.営農類型別経営統計_法人, "営農法人"} _
                   , {ComConst.調査区分.米生産費統計_個別, "米"} _
                   , {ComConst.調査区分.小麦生産費統計_個別, "小麦"} _
                   , {ComConst.調査区分.二条大麦生産費統計_個別, "二条大麦"} _
                   , {ComConst.調査区分.六条大麦生産費統計_個別, "六条大麦"} _
                   , {ComConst.調査区分.はだか麦生産費統計_個別, "はだか麦"} _
                   , {ComConst.調査区分.そば生産費統計_個別, "そば"} _
                   , {ComConst.調査区分.大豆生産費統計_個別, "大豆"} _
                   , {ComConst.調査区分.原料用かんしょ生産費統計_個別, "原料用かんしょ"} _
                   , {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, "原料用ばれいしょ"} _
                   , {ComConst.調査区分.なたね生産費統計_個別, "なたね"} _
                   , {ComConst.調査区分.てんさい生産費統計_個別, "てんさい"} _
                   , {ComConst.調査区分.さとうきび生産費統計_個別, "さとうきび"} _
                   , {ComConst.調査区分.米生産費統計_組織法人, "米"} _
                   , {ComConst.調査区分.小麦生産費統計_組織法人, "小麦"} _
                   , {ComConst.調査区分.大豆生産費統計_組織法人, "大豆"} _
                   , {ComConst.調査区分.牛乳生産費統計_個別, "牛乳"} _
                   , {ComConst.調査区分.子牛生産費統計_個別, "子牛"} _
                   , {ComConst.調査区分.乳用雄育成牛生産費統計_個別, "乳用雄育成牛"} _
                   , {ComConst.調査区分.交雑種育成牛生産費統計_個別, "交雑種育成牛"} _
                   , {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, "去勢若齢肥育牛"} _
                   , {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, "乳用雄肥育牛"} _
                   , {ComConst.調査区分.交雑種肥育牛生産費統計_個別, "交雑種肥育牛"} _
                   , {ComConst.調査区分.肥育豚生産費統計_個別, "肥育豚"} _
                   , {ComConst.調査区分.経営分析調査_二条大麦生産費, "二条大麦"} _
                   , {ComConst.調査区分.経営分析調査_六条大麦生産費, "六条大麦"} _
                   , {ComConst.調査区分.経営分析調査_はだか麦生産費, "はだか麦"} _
                   , {ComConst.調査区分.経営分析調査_そば生産費, "そば"} _
                   , {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, "原料用ばれいしょ"} _
                   , {ComConst.調査区分.経営分析調査_なたね生産費, "なたね"} _
                   , {ComConst.調査区分.経営分析調査_てんさい生産費, "てんさい"} _
                   , {ComConst.調査区分.経営分析調査_さとうきび生産費, "さとうきび"} _
                   , {ComConst.調査区分.経営分析調査_牛乳生産費, "牛乳"} _
                   , {ComConst.調査区分.経営分析調査_子牛生産費, "子牛"} _
                   , {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, "乳用雄育成牛"} _
                   , {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, "交雑種育成牛"} _
                   , {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, "去勢若齢肥育牛"} _
                   , {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, "乳用雄肥育牛"} _
                   , {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, "交雑種肥育牛"} _
                   , {ComConst.調査区分.経営分析調査_肥育豚生産費, "肥育豚"}
               }

        '単位当たり表示名
        Public Shared 単位当たり表示 As New Dictionary(Of String, String) From {
                     {ComConst.調査区分.小麦生産費統計_個別, "60kg"} _
                   , {ComConst.調査区分.二条大麦生産費統計_個別, "50kg"} _
                   , {ComConst.調査区分.六条大麦生産費統計_個別, "50kg"} _
                   , {ComConst.調査区分.はだか麦生産費統計_個別, "60kg"} _
                   , {ComConst.調査区分.そば生産費統計_個別, "45kg"} _
                   , {ComConst.調査区分.大豆生産費統計_個別, "60kg"} _
                   , {ComConst.調査区分.原料用かんしょ生産費統計_個別, "100kg"} _
                   , {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, "100kg"} _
                   , {ComConst.調査区分.なたね生産費統計_個別, "60kg"} _
                   , {ComConst.調査区分.てんさい生産費統計_個別, "1t"} _
                   , {ComConst.調査区分.さとうきび生産費統計_個別, "1t"} _
                   , {ComConst.調査区分.小麦生産費統計_組織法人, "60kg"} _
                   , {ComConst.調査区分.大豆生産費統計_組織法人, "60kg"} _
                   , {ComConst.調査区分.経営分析調査_二条大麦生産費, "50kg"} _
                   , {ComConst.調査区分.経営分析調査_六条大麦生産費, "50kg"} _
                   , {ComConst.調査区分.経営分析調査_はだか麦生産費, "60kg"} _
                   , {ComConst.調査区分.経営分析調査_そば生産費, "45kg"} _
                   , {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, "100kg"} _
                   , {ComConst.調査区分.経営分析調査_なたね生産費, "60kg"} _
                   , {ComConst.調査区分.経営分析調査_てんさい生産費, "1t"} _
                   , {ComConst.調査区分.経営分析調査_さとうきび生産費, "1t"}
         }

    End Class


    '---REV_002 ADD End

    '---REV.003 ADD START
    Public Enum 送受信区分
        本省_局から_受信
        本省_局へ_送信
        局_本省から_受信
        局_本省へ_送信
        局_実査設置拠点から_受信
        局_実査設置拠点へ_送信
        実査設置拠点_局から_受信
        実査設置拠点_局へ_送信
    End Enum

    Public Class 労賃単価反映
        Public Shared 労賃単価カラム As New Dictionary(Of String, String) From {
                {ComConst.調査区分.米生産費統計_個別, "Q12080101"},
                {ComConst.調査区分.小麦生産費統計_個別, "Q12080101"},
                {ComConst.調査区分.二条大麦生産費統計_個別, "Q12080101"},
                {ComConst.調査区分.六条大麦生産費統計_個別, "Q12080101"},
                {ComConst.調査区分.はだか麦生産費統計_個別, "Q12080101"},
                {ComConst.調査区分.そば生産費統計_個別, "Q12080101"},
                {ComConst.調査区分.大豆生産費統計_個別, "Q12080101"},
                {ComConst.調査区分.原料用かんしょ生産費統計_個別, "Q12080101"},
                {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, "Q12080101"},
                {ComConst.調査区分.なたね生産費統計_個別, "Q12080101"},
                {ComConst.調査区分.てんさい生産費統計_個別, "Q12080101"},
                {ComConst.調査区分.さとうきび生産費統計_個別, "Q12080101"},
                {ComConst.調査区分.米生産費統計_組織法人, "Q12070101"},
                {ComConst.調査区分.小麦生産費統計_組織法人, "Q12070101"},
                {ComConst.調査区分.大豆生産費統計_組織法人, "Q12070101"},
                {ComConst.調査区分.牛乳生産費統計_個別, "Q12011301"},
                {ComConst.調査区分.子牛生産費統計_個別, "Q12011101"},
                {ComConst.調査区分.乳用雄育成牛生産費統計_個別, "Q12011101"},
                {ComConst.調査区分.交雑種育成牛生産費統計_個別, "Q12011101"},
                {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, "Q12011101"},
                {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, "Q12011101"},
                {ComConst.調査区分.交雑種肥育牛生産費統計_個別, "Q12011101"},
                {ComConst.調査区分.肥育豚生産費統計_個別, "Q12011101"},
                {ComConst.調査区分.経営分析調査_二条大麦生産費, "Q12070101"},
                {ComConst.調査区分.経営分析調査_六条大麦生産費, "Q12070101"},
                {ComConst.調査区分.経営分析調査_はだか麦生産費, "Q12070101"},
                {ComConst.調査区分.経営分析調査_そば生産費, "Q12070101"},
                {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, "Q12070101"},
                {ComConst.調査区分.経営分析調査_なたね生産費, "Q12070101"},
                {ComConst.調査区分.経営分析調査_てんさい生産費, "Q12070101"},
                {ComConst.調査区分.経営分析調査_さとうきび生産費, "Q12070101"},
                {ComConst.調査区分.経営分析調査_牛乳生産費, "Q12010501"},
                {ComConst.調査区分.経営分析調査_子牛生産費, "Q12010501"},
                {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, "Q12010501"},
                {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, "Q12010501"},
                {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, "Q12010501"},
                {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, "Q12010501"},
                {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, "Q12010501"},
                {ComConst.調査区分.経営分析調査_肥育豚生産費, "Q12010501"}
            }
    End Class

    ''' <summary>
    ''' 牛トレサデータクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 牛トレサデータ

        ''' <summary>テーブル名称</summary>
        'Public Shared テーブル名称 As New Dictionary(Of String, String()) From { _

        Public Class 牛の識別CD
            Public Const ホルスタイン種 As Integer = 1
            Public Const ジャージー種 As Integer = 2
            Public Const 交雑種 As Integer = 3
            Public Const 黒毛和種 As Integer = 4
            Public Const 褐毛和種 As Integer = 5
            Public Const 日本短角種 As Integer = 6
            Public Const 無角和種 As Integer = 7
            Public Const 黒毛和種Ｘ褐毛和種 As Integer = 8
            Public Const 和牛間交雑種 As Integer = 10
            Public Const 肉専用種 As Integer = 11
            Public Const 乳用種 As Integer = 12
            Public Const 不明 As Integer = 91
            Public Const その他 As Integer = 99

        End Class

        Public Class 性別CD
            Public Const おす As Integer = 1
            Public Const めす As Integer = 2
            Public Const 去勢 As Integer = 3
            Public Const フリーマーチン As Integer = 4
            Public Const 不明 As Integer = 5
            Public Const その他 As Integer = 9
        End Class

        Public Class 異動フラグ
            Public Const 初期装着 As Integer = 1
            Public Const 出生 As Integer = 2
            Public Const 転出取引 As Integer = 3
            Public Const 転入搬入 As Integer = 4
            Public Const 死亡と畜 As Integer = 5
            Public Const と場 As Integer = 6
            Public Const 輸出 As Integer = 7
            Public Const 輸入 As Integer = 8
            Public Const とさつ As Integer = 9

        End Class

        Public Shared 異動名称 As New Dictionary(Of Integer, String) From {
            {異動フラグ.初期装着, "初期装着"},
            {異動フラグ.出生, "出生"},
            {異動フラグ.転出取引, "転出/取引"},
            {異動フラグ.転入搬入, "転入/搬入"},
            {異動フラグ.死亡と畜, "死亡/と畜"},
            {異動フラグ.と場, "と場"},
            {異動フラグ.輸出, "輸出"},
            {異動フラグ.輸入, "輸入"},
            {異動フラグ.とさつ, "とさつ"}
           }

        Public Enum 調査区分分類
            牛乳
            子牛
            乳用
            交雑
            去勢
        End Enum

        Public Shared 調査区分分類変換テーブル As New Dictionary(Of String, 調査区分分類) From {
            {ComConst.調査区分.牛乳生産費統計_個別, 調査区分分類.牛乳},
            {ComConst.調査区分.経営分析調査_牛乳生産費, 調査区分分類.牛乳},
            {ComConst.調査区分.子牛生産費統計_個別, 調査区分分類.子牛},
            {ComConst.調査区分.経営分析調査_子牛生産費, 調査区分分類.子牛},
            {ComConst.調査区分.乳用雄育成牛生産費統計_個別, 調査区分分類.乳用},
            {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, 調査区分分類.乳用},
            {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, 調査区分分類.乳用},
            {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, 調査区分分類.乳用},
            {ComConst.調査区分.交雑種育成牛生産費統計_個別, 調査区分分類.交雑},
            {ComConst.調査区分.交雑種肥育牛生産費統計_個別, 調査区分分類.交雑},
            {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, 調査区分分類.交雑},
            {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, 調査区分分類.交雑},
            {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, 調査区分分類.去勢},
            {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, 調査区分分類.去勢}
        }

        Public Shared 品種コード変換テーブル As New Dictionary(Of Integer, Dictionary(Of Integer, String)) From {
            {ComConst.牛トレサデータ.牛の識別CD.ホルスタイン種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "1"}, {ComConst.牛トレサデータ.性別CD.めす, "1"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.ジャージー種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "1"}, {ComConst.牛トレサデータ.性別CD.めす, "99"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.交雑種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "6"}, {ComConst.牛トレサデータ.性別CD.めす, "6"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.乳用種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "1"}, {ComConst.牛トレサデータ.性別CD.めす, "3"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.黒毛和種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "2"}, {ComConst.牛トレサデータ.性別CD.めす, "5"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.褐毛和種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "3"}, {ComConst.牛トレサデータ.性別CD.めす, "6"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.日本短角種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "4"}, {ComConst.牛トレサデータ.性別CD.めす, "7"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.無角和種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "5"}, {ComConst.牛トレサデータ.性別CD.めす, "8"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.黒毛和種Ｘ褐毛和種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "3"}, {ComConst.牛トレサデータ.性別CD.めす, "6"}
            }}
        }

        Public Shared 品種コード変換テーブル_牛乳 As New Dictionary(Of Integer, Dictionary(Of Integer, String)) From {
            {ComConst.牛トレサデータ.牛の識別CD.ホルスタイン種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "1"}, {ComConst.牛トレサデータ.性別CD.めす, "1"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.ジャージー種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "1"}, {ComConst.牛トレサデータ.性別CD.めす, "2"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.交雑種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "6"}, {ComConst.牛トレサデータ.性別CD.めす, "6"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.乳用種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "1"}, {ComConst.牛トレサデータ.性別CD.めす, "3"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.黒毛和種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "11"}, {ComConst.牛トレサデータ.性別CD.めす, "11"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.褐毛和種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "11"}, {ComConst.牛トレサデータ.性別CD.めす, "11"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.日本短角種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "11"}, {ComConst.牛トレサデータ.性別CD.めす, "11"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.無角和種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "11"}, {ComConst.牛トレサデータ.性別CD.めす, "11"}
            }},
            {ComConst.牛トレサデータ.牛の識別CD.黒毛和種Ｘ褐毛和種, New Dictionary(Of Integer, String) From {
                    {ComConst.牛トレサデータ.性別CD.おす, "3"}, {ComConst.牛トレサデータ.性別CD.めす, "6"}
            }}
        }

        Public Shared 性別コード変換テーブル As New Dictionary(Of Integer, String) From {
            {ComConst.牛トレサデータ.性別CD.おす, "1"},
            {ComConst.牛トレサデータ.性別CD.めす, "2"},
            {ComConst.牛トレサデータ.性別CD.去勢, "1"},
            {ComConst.牛トレサデータ.性別CD.フリーマーチン, "2"}
        }

        Public Shared 異動コード変換テーブル As New Dictionary(Of Integer, String) From {
            {ComConst.牛トレサデータ.異動フラグ.出生, "1：購入（生産）"},
            {ComConst.牛トレサデータ.異動フラグ.転出取引, "2：売却（売却）"},
            {ComConst.牛トレサデータ.異動フラグ.転入搬入, "1：購入（購入）"},
            {ComConst.牛トレサデータ.異動フラグ.死亡と畜, "2：売却（死亡）"}
        }
        Public Shared 異動コード変換テーブル_子牛 As New Dictionary(Of Integer, String) From {
            {ComConst.牛トレサデータ.異動フラグ.出生, "1：転入（生産）"},
            {ComConst.牛トレサデータ.異動フラグ.転出取引, "2：転出（売却）"},
            {ComConst.牛トレサデータ.異動フラグ.転入搬入, "1：転入（購入）"},
            {ComConst.牛トレサデータ.異動フラグ.死亡と畜, "2：転出（死亡）"}
        }

        ''' <summary>
        ''' 牛資産異動情報フォイルのシート構成
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared 異動ファイルシート As New Dictionary(Of String, String()) From {
            {調査区分.牛乳生産費統計_個別, {"表紙", "【11】搾乳牛所有状況"}} _
          , {調査区分.子牛生産費統計_個別, {"表紙", "【２】2牛取引情報"}} _
          , {調査区分.乳用雄育成牛生産費統計_個別, {"表紙", "【２】牛１"}} _
          , {調査区分.交雑種育成牛生産費統計_個別, {"表紙", "【２】牛１"}} _
          , {調査区分.去勢若齢肥育牛生産費統計_個別, {"表紙", "【２】牛１"}} _
          , {調査区分.乳用雄肥育牛生産費統計_個別, {"表紙", "【２】牛１"}} _
          , {調査区分.交雑種肥育牛生産費統計_個別, {"表紙", "【２】牛１"}} _
          , {調査区分.経営分析調査_牛乳生産費, {"表紙", "【11】搾乳牛所有状況"}} _
          , {調査区分.経営分析調査_子牛生産費, {"表紙", "【２】２牛取引情報"}} _
          , {調査区分.経営分析調査_乳用雄育成牛生産費, {"表紙", "【２】牛１"}} _
          , {調査区分.経営分析調査_交雑種育成牛生産費, {"表紙", "【２】牛１"}} _
          , {調査区分.経営分析調査_去勢若齢肥育牛生産費, {"表紙", "【２】牛１"}} _
          , {調査区分.経営分析調査_乳用雄肥育牛生産費, {"表紙", "【２】牛１"}} _
          , {調査区分.経営分析調査_交雑種肥育牛生産費, {"表紙", "【２】牛１"}}
        }

        ''' <summary>
        ''' 牛資産異動情報フォイルの各シートのデータ範囲
        ''' </summary>
        ''' <remarks></remarks>
        '20220201 MOD START 牛トレサの表示行数入力欄がAD列に異動したことの対応
        Public Shared シートデータ範囲 As New Dictionary(Of String, Dictionary(Of String, String)) From {
            {調査区分.牛乳生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【11】搾乳牛所有状況", "A1:AD2015"}
                    }
                } _
          , {調査区分.子牛生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【２】2牛取引情報", "A1:AA2014"}
                    }
                } _
          , {調査区分.乳用雄育成牛生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【２】牛１", "A1:T10015"}
                    }
                } _
          , {調査区分.交雑種育成牛生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【２】牛１", "A1:T10015"}
                    }
                } _
          , {調査区分.去勢若齢肥育牛生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【２】牛１", "A1:T10015"}
                    }
                } _
          , {調査区分.乳用雄肥育牛生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【２】牛１", "A1:T10015"}
                    }
                } _
          , {調査区分.交雑種肥育牛生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【２】牛１", "A1:T10015"}
                    }
                } _
          , {調査区分.経営分析調査_牛乳生産費, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【１】経営概要", "A1:V34"},
                        {"【11】搾乳牛所有状況", "A1:AI13014"}
                    }
                } _
          , {調査区分.経営分析調査_子牛生産費, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【１】経営概要", "A1:V34"},
                        {"【２】２牛取引情報", "A1:AA8014"}
                    }
                } _
          , {調査区分.経営分析調査_乳用雄育成牛生産費, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【１】経営概要", "A1:V34"},
                        {"【２】牛１", "A1:T30015"}
                    }
                } _
          , {調査区分.経営分析調査_交雑種育成牛生産費, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【１】経営概要", "A1:V34"},
                        {"【２】牛１", "A1:T30015"}
                    }
                } _
          , {調査区分.経営分析調査_去勢若齢肥育牛生産費, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【１】経営概要", "A1:V34"},
                        {"【２】牛１", "A1:T30015"}
                    }
                } _
          , {調査区分.経営分析調査_乳用雄肥育牛生産費, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【１】経営概要", "A1:V34"},
                        {"【２】牛１", "A1:T30015"}
                    }
                } _
          , {調査区分.経営分析調査_交雑種肥育牛生産費, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【１】経営概要", "A1:V34"},
                        {"【２】牛１", "A1:T30015"}
                    }
                }
            }
        '20220201 MOD END

        '---REV_013↓
        ''' <summary>
        ''' 牛資産総括情報フォイルのシート構成
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared 総括ファイルシート As New Dictionary(Of String, String()) From {
            {調査区分.牛乳生産費統計_個別, {"表紙", "【11】搾乳牛所有状況２"}} _
          , {調査区分.子牛生産費統計_個別, {"表紙", "【２】対象畜概要１"}} _
          , {調査区分.乳用雄育成牛生産費統計_個別, {"表紙", "【２】牛２"}} _
          , {調査区分.交雑種育成牛生産費統計_個別, {"表紙", "【２】牛２"}} _
          , {調査区分.去勢若齢肥育牛生産費統計_個別, {"表紙", "【２】牛２"}} _
          , {調査区分.乳用雄肥育牛生産費統計_個別, {"表紙", "【２】牛２"}} _
          , {調査区分.交雑種肥育牛生産費統計_個別, {"表紙", "【２】牛２"}}
        }

        ''' <summary>
        ''' 牛資産総括情報フォイルの各シートのデータ範囲
        ''' </summary>
        ''' <remarks></remarks>
        '20220201 MOD START 牛トレサの表示行数入力欄がAD列に異動したことの対応
        Public Shared シートデータ範囲_総括 As New Dictionary(Of String, Dictionary(Of String, String)) From {
            {調査区分.牛乳生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【11】搾乳牛所有状況２", "A1:CG2015"}
                    }
                } _
          , {調査区分.子牛生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【２】対象畜概要１", "A1:DF2019"}
                    }
                } _
          , {調査区分.乳用雄育成牛生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【２】牛２", "A1:AL8010"}
                    }
                } _
          , {調査区分.交雑種育成牛生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【２】牛２", "A1:AL8010"}
                    }
                } _
          , {調査区分.去勢若齢肥育牛生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【２】牛２", "A1:AL8010"}
                    }
                } _
          , {調査区分.乳用雄肥育牛生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【２】牛２", "A1:AL8010"}
                    }
                } _
          , {調査区分.交雑種肥育牛生産費統計_個別, New Dictionary(Of String, String) From {
                        {"表紙", "A1:BS67"},
                        {"【２】牛２", "A1:AL8010"}
                    }
                }
            }
        '---REV_013↑
        ''' <summary>
        ''' 牛資産異動情報の異動月を取得する項目番号
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared 異動月_項目番号 As New Dictionary(Of String, String) From {
            {調査区分.牛乳生産費統計_個別, "Q11020301"} _
          , {調査区分.子牛生産費統計_個別, "Q02020301"} _
          , {調査区分.乳用雄育成牛生産費統計_個別, "Q02020301"} _
          , {調査区分.交雑種育成牛生産費統計_個別, "Q02020301"} _
          , {調査区分.去勢若齢肥育牛生産費統計_個別, "Q02020301"} _
          , {調査区分.乳用雄肥育牛生産費統計_個別, "Q02020301"} _
          , {調査区分.交雑種肥育牛生産費統計_個別, "Q02020301"} _
          , {調査区分.経営分析調査_牛乳生産費, "Q11020301"} _
          , {調査区分.経営分析調査_子牛生産費, "Q02020301"} _
          , {調査区分.経営分析調査_乳用雄育成牛生産費, "Q02020301"} _
          , {調査区分.経営分析調査_交雑種育成牛生産費, "Q02020301"} _
          , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "Q02020301"} _
          , {調査区分.経営分析調査_乳用雄肥育牛生産費, "Q02020301"} _
          , {調査区分.経営分析調査_交雑種肥育牛生産費, "Q02020301"}
        }
        ''' <summary>
        ''' 牛資産異動情報の決算月を取得する項目番号
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared 決算月_項目番号 As New Dictionary(Of String, String) From {
            {調査区分.牛乳生産費統計_個別, ""} _
          , {調査区分.子牛生産費統計_個別, ""} _
          , {調査区分.乳用雄育成牛生産費統計_個別, ""} _
          , {調査区分.交雑種育成牛生産費統計_個別, ""} _
          , {調査区分.去勢若齢肥育牛生産費統計_個別, ""} _
          , {調査区分.乳用雄肥育牛生産費統計_個別, ""} _
          , {調査区分.交雑種肥育牛生産費統計_個別, ""} _
          , {調査区分.経営分析調査_牛乳生産費, "Q01040101"} _
          , {調査区分.経営分析調査_子牛生産費, "Q01040101"} _
          , {調査区分.経営分析調査_乳用雄育成牛生産費, "Q01040101"} _
          , {調査区分.経営分析調査_交雑種育成牛生産費, "Q01040101"} _
          , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "Q01040101"} _
          , {調査区分.経営分析調査_乳用雄肥育牛生産費, "Q01040101"} _
          , {調査区分.経営分析調査_交雑種肥育牛生産費, "Q01040101"}
        }
        ''' <summary>
        ''' 牛資産異動情報の異動年月を取得する項目番号
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared 異動年月_項目番号 As New Dictionary(Of String, String) From {
            {調査区分.牛乳生産費統計_個別, "Q11020801"} _
          , {調査区分.子牛生産費統計_個別, ""} _
          , {調査区分.乳用雄育成牛生産費統計_個別, ""} _
          , {調査区分.交雑種育成牛生産費統計_個別, ""} _
          , {調査区分.去勢若齢肥育牛生産費統計_個別, ""} _
          , {調査区分.乳用雄肥育牛生産費統計_個別, ""} _
          , {調査区分.交雑種肥育牛生産費統計_個別, ""} _
          , {調査区分.経営分析調査_牛乳生産費, "Q11020801"} _
          , {調査区分.経営分析調査_子牛生産費, ""} _
          , {調査区分.経営分析調査_乳用雄育成牛生産費, ""} _
          , {調査区分.経営分析調査_交雑種育成牛生産費, ""} _
          , {調査区分.経営分析調査_去勢若齢肥育牛生産費, ""} _
          , {調査区分.経営分析調査_乳用雄肥育牛生産費, ""} _
          , {調査区分.経営分析調査_交雑種肥育牛生産費, ""}
        }
        ''' <summary>
        ''' 牛資産異動情報を取り込む項目番号の範囲
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared 取込対象_項目番号 As New Dictionary(Of String, String()) From {
            {調査区分.牛乳生産費統計_個別, {"Q11020101", "Q11021701"}} _
          , {調査区分.子牛生産費統計_個別, {"Q02020101", "Q02021201"}} _
          , {調査区分.乳用雄育成牛生産費統計_個別, {"Q02020101", "Q02021301"}} _
          , {調査区分.交雑種育成牛生産費統計_個別, {"Q02020101", "Q02021301"}} _
          , {調査区分.去勢若齢肥育牛生産費統計_個別, {"Q02020101", "Q02021301"}} _
          , {調査区分.乳用雄肥育牛生産費統計_個別, {"Q02020101", "Q02021301"}} _
          , {調査区分.交雑種肥育牛生産費統計_個別, {"Q02020101", "Q02021301"}} _
          , {調査区分.経営分析調査_牛乳生産費, {"Q11020101", "Q11021701"}} _
          , {調査区分.経営分析調査_子牛生産費, {"Q02020101", "Q02021201"}} _
          , {調査区分.経営分析調査_乳用雄育成牛生産費, {"Q02020101", "Q02021301"}} _
          , {調査区分.経営分析調査_交雑種育成牛生産費, {"Q02020101", "Q02021301"}} _
          , {調査区分.経営分析調査_去勢若齢肥育牛生産費, {"Q02020101", "Q02021301"}} _
          , {調査区分.経営分析調査_乳用雄肥育牛生産費, {"Q02020101", "Q02021301"}} _
          , {調査区分.経営分析調査_交雑種肥育牛生産費, {"Q02020101", "Q02021301"}}
        }
        ''' <summary>
        ''' 成畜牛資産異動情報を取り込む項目番号の範囲
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared 成畜取込対象_項目番号 As New Dictionary(Of String, String()) From {
            {調査区分.牛乳生産費統計_個別, {"Q11020701", "Q11021701"}} _
          , {調査区分.経営分析調査_牛乳生産費, {"Q11020701", "Q11021701"}}
        }

        ''' <summary>
        ''' 牛資産異動情報を取り込む項目番号
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared 対象_項目番号 As New Dictionary(Of String, String()) From {
            {調査区分.牛乳生産費統計_個別, {"Q11020101", "Q11020201", "Q11020301", "Q11020401", "Q11020501", "Q11020601", "Q11020701", "Q11020801", "Q11020901", "Q11021001", "Q11021101", "Q11021201", "Q11021301", "Q11021401", "Q11021501", "Q11021601", "Q11021701"}} _
          , {調査区分.子牛生産費統計_個別, {"Q02020101", "Q02020201", "Q02020301", "Q02020401", "Q02020501", "Q02020601", "Q02020701", "Q02020801", "Q02020901", "Q02021001", "Q02021101", "Q02021201"}} _
          , {調査区分.乳用雄育成牛生産費統計_個別, {"Q02020101", "Q02020201", "Q02020301", "Q02020401", "Q02020501", "Q02020601", "Q02020701", "Q02020801", "Q02020901", "Q02021001", "Q02021101", "Q02021201", "Q02021301"}} _
          , {調査区分.交雑種育成牛生産費統計_個別, {"Q02020101", "Q02020201", "Q02020301", "Q02020401", "Q02020501", "Q02020601", "Q02020701", "Q02020801", "Q02020901", "Q02021001", "Q02021101", "Q02021201", "Q02021301"}} _
          , {調査区分.去勢若齢肥育牛生産費統計_個別, {"Q02020101", "Q02020201", "Q02020301", "Q02020401", "Q02020501", "Q02020601", "Q02020701", "Q02020801", "Q02020901", "Q02021001", "Q02021101", "Q02021201", "Q02021301"}} _
          , {調査区分.乳用雄肥育牛生産費統計_個別, {"Q02020101", "Q02020201", "Q02020301", "Q02020401", "Q02020501", "Q02020601", "Q02020701", "Q02020801", "Q02020901", "Q02021001", "Q02021101", "Q02021201", "Q02021301"}} _
          , {調査区分.交雑種肥育牛生産費統計_個別, {"Q02020101", "Q02020201", "Q02020301", "Q02020401", "Q02020501", "Q02020601", "Q02020701", "Q02020801", "Q02020901", "Q02021001", "Q02021101", "Q02021201", "Q02021301"}} _
          , {調査区分.経営分析調査_牛乳生産費, {"Q11020101", "Q11020201", "Q11020301", "Q11020401", "Q11020501", "Q11020601", "Q11020701", "Q11020801", "Q11020901", "Q11021001", "Q11021101", "Q11021201", "Q11021301", "Q11021401", "Q11021501", "Q11021601", "Q11021701"}} _
          , {調査区分.経営分析調査_子牛生産費, {"Q02020101", "Q02020201", "Q02020301", "Q02020401", "Q02020501", "Q02020601", "Q02020701", "Q02020801", "Q02020901", "Q02021001", "Q02021101", "Q02021201"}} _
          , {調査区分.経営分析調査_乳用雄育成牛生産費, {"Q02020101", "Q02020201", "Q02020301", "Q02020401", "Q02020501", "Q02020601", "Q02020701", "Q02020801", "Q02020901", "Q02021001", "Q02021101", "Q02021201", "Q02021301"}} _
          , {調査区分.経営分析調査_交雑種育成牛生産費, {"Q02020101", "Q02020201", "Q02020301", "Q02020401", "Q02020501", "Q02020601", "Q02020701", "Q02020801", "Q02020901", "Q02021001", "Q02021101", "Q02021201", "Q02021301"}} _
          , {調査区分.経営分析調査_去勢若齢肥育牛生産費, {"Q02020101", "Q02020201", "Q02020301", "Q02020401", "Q02020501", "Q02020601", "Q02020701", "Q02020801", "Q02020901", "Q02021001", "Q02021101", "Q02021201", "Q02021301"}} _
          , {調査区分.経営分析調査_乳用雄肥育牛生産費, {"Q02020101", "Q02020201", "Q02020301", "Q02020401", "Q02020501", "Q02020601", "Q02020701", "Q02020801", "Q02020901", "Q02021001", "Q02021101", "Q02021201", "Q02021301"}} _
          , {調査区分.経営分析調査_交雑種肥育牛生産費, {"Q02020101", "Q02020201", "Q02020301", "Q02020401", "Q02020501", "Q02020601", "Q02020701", "Q02020801", "Q02020901", "Q02021001", "Q02021101", "Q02021201", "Q02021301"}}
        }

        ''' <summary>
        ''' 成畜牛資産異動情報を取り込む項目番号
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared 成畜対象_項目番号 As New Dictionary(Of String, String()) From {
            {調査区分.牛乳生産費統計_個別, {"Q11020701", "Q11020801", "Q11020901", "Q11021001", "Q11021101", "Q11021201", "Q11021301", "Q11021401", "Q11021501", "Q11021601", "Q11021701"}} _
          , {調査区分.経営分析調査_牛乳生産費, {"Q11020701", "Q11020801", "Q11020901", "Q11021001", "Q11021101", "Q11021201", "Q11021301", "Q11021401", "Q11021501", "Q11021601", "Q11021701"}}
        }

    End Class
    '---REV_003 ADD END

    'REV 006 start-----------------
    ''' <summary>
    ''' 調査票審査論理範囲クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 調査票審査論理範囲
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "調査票審査論理（範囲）.xlsm"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 調査票審査論理範囲.入力用ファイル名称
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "調査票審査論理（範囲）"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "審査論理"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 6
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 1005
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 調査票審査論理範囲.出力用ファイル名称.Row.Last - 調査票審査論理範囲.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "C"
                ''' <summary>最終列</summary>
                Public Const Last As String = "I"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "チェック項目名"} _
                , {2, "項目番号１"} _
                , {3, "項目番号２"} _
                , {4, "値"} _
                , {5, "下限"} _
                , {6, "上限"} _
                , {7, "繰り返し"}
            }
        End Class
    End Class

    ''' <summary>
    ''' 労働時間整理ファイルクラス
    ''' </summary>
    ''' <remarks>
    ''' 'REV007 Add
    ''' </remarks>
    Public Class 労働時間整理ファイル
        'Public Const delimiter As String = "/"

        Public Class 労働時間整理ファイルシート名_営農
            ''' <summary>労働時間整理ファイルシート名（営農）</summary>
            Public Shared eiKoTsukiRoudou As String = "10_労働"
            Public Shared eiKoTsukiRoudouHinmoku As String = "11_労働（指定品目）"
            Public Shared eiKoHiRoudou As String = "10_労働"
            Public Shared eiKoHiRoudouHinmoku As String = "11_労働（指定品目）"
            Public Shared eiHoTsukiRoudou01 As String = "12_01_労働の概況"
            Public Shared eiHoTsukiRoudou02 As String = "12_02_労働の概況"
            Public Shared eiHoHiRoudou01 As String = "12_01_労働の概況"
            Public Shared eiHoHiRoudou02 As String = "12_02_労働の概況"
        End Class

        Public Class ベース用ファイル名称
            Public Class Report
                ''' <summary>テンプレートファイル名</summary>
                Public Shared tempFileNameEiKoTsuki As String = "労働時間整理ファイル（営個、月別）.xlsx"
                Public Shared tempFileNameEiKoHi As String = "労働時間整理ファイル（営個、日別）.xlsx"
                Public Shared tempFileNameEiHoTsuki As String = "労働時間整理ファイル（営法、月別）.xlsx"
                Public Shared tempFileNameEiHoHi As String = "労働時間整理ファイル（営法、日別）.xlsx"
                Public Shared tempFileNameTikuTsuki As String = "労働時間整理ファイル（畜生、月別）.xlsx"
                Public Shared tempFileNameTikuHi As String = "労働時間整理ファイル（畜生、日別）.xlsx"
                Public Shared tempFileNameNoTsuki As String = "労働時間整理ファイル（農生、月別）.xlsx"
                Public Shared tempFileNameNoHi As String = "労働時間整理ファイル（農生、日別）.xlsx"

            End Class

        End Class

        Public Class 取り込みファイル判定
            Public Class 判定内容
                Public シート名 As String
                Public セル番号 As String
                Public セル内容 As String
            End Class

            Public Shared センサス桁数判定内容一覧 As New Dictionary(Of String, 判定内容) From {
                 {"1", New 判定内容 With {.シート名 = "設定", .セル番号 = "C3", .セル内容 = "1"}},
                 {"2", New 判定内容 With {.シート名 = "設定", .セル番号 = "C3", .セル内容 = "2"}},
                 {"3", New 判定内容 With {.シート名 = "設定", .セル番号 = "B3", .セル内容 = "3"}},
                 {"4", New 判定内容 With {.シート名 = "設定", .セル番号 = "B3", .セル内容 = "4"}}
            }

        End Class

        Public Class 経営種類及び生産費区分
            ''' <summary>経営種類及び生産費区分　確認</summary>
            Public Shared 確認 As New Dictionary(Of String, String()) From {
                  {"経営種類", {"個別経営体", "組織法人経営体"}} _
                , {"農生生産費区分", {"なたね", "そば", "二条大麦", "六条大麦", "はだか麦", "米", "小麦", "大豆", "原料用かんしょ", "原料用ばれいしょ", "さとうきび", "てんさい"}} _
                , {"畜生生産費区分", {"牛乳", "子牛", "乳用雄育成牛", "交雑種育成牛", "去勢若齢肥育牛", "乳用雄肥育牛", "交雑種肥育牛", "肥育豚"}}
            }
        End Class

        Public Class 労働時間シート名
            ''' <summary>労働時間　シート名</summary>
            Public Shared シート名 As New Dictionary(Of String, String) From {
                    {調査区分.米生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.小麦生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.二条大麦生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.六条大麦生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.はだか麦生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.そば生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.大豆生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.原料用かんしょ生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.原料用ばれいしょ生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.なたね生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.てんさい生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.さとうきび生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.米生産費統計_組織法人, "【12】作業別労働時間①_組織"} _
                  , {調査区分.小麦生産費統計_組織法人, "【12】作業別労働時間①_組織"} _
                  , {調査区分.大豆生産費統計_組織法人, "【12】作業別労働時間①_組織"} _
                  , {調査区分.牛乳生産費統計_個別, "【12】労働時間（個別・牛乳）"} _
                  , {調査区分.子牛生産費統計_個別, "【12】労働時間（個別・牛乳以外）"} _
                  , {調査区分.乳用雄育成牛生産費統計_個別, "【12】労働時間（個別・牛乳以外）"} _
                  , {調査区分.交雑種育成牛生産費統計_個別, "【12】労働時間（個別・牛乳以外）"} _
                  , {調査区分.去勢若齢肥育牛生産費統計_個別, "【12】労働時間（個別・牛乳以外）"} _
                  , {調査区分.乳用雄肥育牛生産費統計_個別, "【12】労働時間（個別・牛乳以外）"} _
                  , {調査区分.交雑種肥育牛生産費統計_個別, "【12】労働時間（個別・牛乳以外）"} _
                  , {調査区分.肥育豚生産費統計_個別, "【12】労働時間（個別・牛乳以外）"} _
                  , {調査区分.経営分析調査_二条大麦生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_六条大麦生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_はだか麦生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_そば生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_原料用ばれいしょ生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_なたね生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_てんさい生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_さとうきび生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_牛乳生産費, "【12】労働時間（経営分析・牛乳）"} _
                  , {調査区分.経営分析調査_子牛生産費, "【12】労働時間（経営分析・牛乳以外）"} _
                  , {調査区分.経営分析調査_乳用雄育成牛生産費, "【12】労働時間（経営分析・牛乳以外）"} _
                  , {調査区分.経営分析調査_交雑種育成牛生産費, "【12】労働時間（経営分析・牛乳以外）"} _
                  , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "【12】労働時間（経営分析・牛乳以外）"} _
                  , {調査区分.経営分析調査_乳用雄肥育牛生産費, "【12】労働時間（経営分析・牛乳以外）"} _
                  , {調査区分.経営分析調査_交雑種肥育牛生産費, "【12】労働時間（経営分析・牛乳以外）"} _
                  , {調査区分.経営分析調査_肥育豚生産費, "【12】労働時間（経営分析・牛乳以外）"}
                }

        End Class

        Public Class 調査票シート名
            ''' <summary>調査票シート名</summary>
            Public Shared 調査票シート名 As New Dictionary(Of String, String) From {
                    {調査区分.米生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.小麦生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.二条大麦生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.六条大麦生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.はだか麦生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.そば生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.大豆生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.原料用かんしょ生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.原料用ばれいしょ生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.なたね生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.てんさい生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.さとうきび生産費統計_個別, "【12】作業別労働時間①_個別"} _
                  , {調査区分.米生産費統計_組織法人, "【12】作業別労働時間①_組織"} _
                  , {調査区分.小麦生産費統計_組織法人, "【12】作業別労働時間①_組織"} _
                  , {調査区分.大豆生産費統計_組織法人, "【12】作業別労働時間①_組織"} _
                  , {調査区分.牛乳生産費統計_個別, "【12】労働時間"} _
                  , {調査区分.子牛生産費統計_個別, "【12】労働時間"} _
                  , {調査区分.乳用雄育成牛生産費統計_個別, "【12】労働時間"} _
                  , {調査区分.交雑種育成牛生産費統計_個別, "【12】労働時間"} _
                  , {調査区分.去勢若齢肥育牛生産費統計_個別, "【12】労働時間"} _
                  , {調査区分.乳用雄肥育牛生産費統計_個別, "【12】労働時間"} _
                  , {調査区分.交雑種肥育牛生産費統計_個別, "【12】労働時間"} _
                  , {調査区分.肥育豚生産費統計_個別, "【12】労働時間"} _
                  , {調査区分.経営分析調査_二条大麦生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_六条大麦生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_はだか麦生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_そば生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_原料用ばれいしょ生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_なたね生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_てんさい生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_さとうきび生産費, "【12】作業別労働時間①_組織"} _
                  , {調査区分.経営分析調査_牛乳生産費, "【12】労働時間"} _
                  , {調査区分.経営分析調査_子牛生産費, "【12】労働時間"} _
                  , {調査区分.経営分析調査_乳用雄育成牛生産費, "【12】労働時間"} _
                  , {調査区分.経営分析調査_交雑種育成牛生産費, "【12】労働時間"} _
                  , {調査区分.経営分析調査_去勢若齢肥育牛生産費, "【12】労働時間"} _
                  , {調査区分.経営分析調査_乳用雄肥育牛生産費, "【12】労働時間"} _
                  , {調査区分.経営分析調査_交雑種肥育牛生産費, "【12】労働時間"} _
                  , {調査区分.経営分析調査_肥育豚生産費, "【12】労働時間"}
                }
        End Class

        Public Class シートデータ範囲
            ''' <summary>ファイル名毎のシートのリスト取得</summary>
            Public Shared 範囲 As New Dictionary(Of String, Dictionary(Of String, String)) From {
                    {ベース用ファイル名称.Report.tempFileNameEiKoHi, New Dictionary(Of String, String) From {
                          {"設定", "A1:W111"},
                          {"日別入力", "A1:AP681"},
                          {"10_労働", "A1:AH123"},
                          {"11_労働（指定品目）", "A1:Z47"}
                      }
                    } _
                    , {ベース用ファイル名称.Report.tempFileNameEiKoTsuki, New Dictionary(Of String, String) From {
                          {"設定", "A1:W111"},
                          {"月別入力", "A1:AP681"},
                          {"10_労働", "A1:AH123"},
                          {"11_労働（指定品目）", "A1:Z47"}
                      }
                    } _
                    , {ベース用ファイル名称.Report.tempFileNameEiHoHi, New Dictionary(Of String, String) From {
                          {"設定", "A1:N108"},
                          {"日別入力", "A1:J681"},
                          {"12_01_労働の概況", "A1:V44"},
                          {"12_02_労働の概況", "A1:AM51"}
                      }
                    } _
                    , {ベース用ファイル名称.Report.tempFileNameEiHoTsuki, New Dictionary(Of String, String) From {
                          {"設定", "A1:N108"},
                          {"月別入力", "A1:J681"},
                          {"12_01_労働の概況", "A1:V44"},
                          {"12_02_労働の概況", "A1:AM51"}
                      }
                    } _
                    , {ベース用ファイル名称.Report.tempFileNameNoHi, New Dictionary(Of String, String) From {
                          {"設定", "A1:X109"},
                          {"日別入力", "A1:AO106"},
                          {"【12】作業別労働時間①_個別", "A1:BR55"},
                          {"【12】作業別労働時間①_組織", "A1:BB48"}
                      }
                    } _
                    , {ベース用ファイル名称.Report.tempFileNameNoTsuki, New Dictionary(Of String, String) From {
                          {"設定", "A1:X109"},
                          {"月別入力", "A1:I106"},
                          {"【12】作業別労働時間①_個別", "A1:BR55"},
                          {"【12】作業別労働時間①_組織", "A1:BB48"}
                      }
                    } _
                    , {ベース用ファイル名称.Report.tempFileNameTikuHi, New Dictionary(Of String, String) From {
                          {"設定", "A1:R108"},
                          {"日別入力", "A1:AO537"},
                          {"【12】労働時間（個別・牛乳）", "A1:BF113"},
                          {"【12】労働時間（個別・牛乳以外）", "A1:BF113"},
                          {"【12】労働時間（経営分析・牛乳）", "A1:AQ36"},
                          {"【12】労働時間（経営分析・牛乳以外）", "A1:AQ32"}
                      }
                    } _
                    , {ベース用ファイル名称.Report.tempFileNameTikuTsuki, New Dictionary(Of String, String) From {
                          {"設定", "A1:R108"},
                          {"月別入力", "A1:I537"},
                          {"【12】労働時間（個別・牛乳）", "A1:BF113"},
                          {"【12】労働時間（個別・牛乳以外）", "A1:BF113"},
                          {"【12】労働時間（経営分析・牛乳）", "A1:AQ36"},
                          {"【12】労働時間（経営分析・牛乳以外）", "A1:AQ32"}
                      }
                    }
                }
        End Class

        Public Class シートリスト
            ''' <summary>ファイル名毎のシートのリスト取得</summary>
            Public Shared リスト As New Dictionary(Of String, String()) From {
                        {ベース用ファイル名称.Report.tempFileNameEiKoHi, {"設定", "日別入力", "10_労働", "11_労働（指定品目）"}},
                        {ベース用ファイル名称.Report.tempFileNameEiKoTsuki, {"設定", "月別入力", "10_労働", "11_労働（指定品目）"}},
                        {ベース用ファイル名称.Report.tempFileNameEiHoHi, {"設定", "日別入力", "12_01_労働の概況", "12_02_労働の概況"}},
                        {ベース用ファイル名称.Report.tempFileNameEiHoTsuki, {"設定", "月別入力", "12_01_労働の概況", "12_02_労働の概況"}},
                        {ベース用ファイル名称.Report.tempFileNameNoHi, {"設定", "日別入力", "【12】作業別労働時間①_個別", "【12】作業別労働時間①_組織"}},
                        {ベース用ファイル名称.Report.tempFileNameNoTsuki, {"設定", "月別入力", "【12】作業別労働時間①_個別", "【12】作業別労働時間①_組織"}},
                        {ベース用ファイル名称.Report.tempFileNameTikuHi, {"設定", "日別入力", "【12】労働時間（個別・牛乳）", "【12】労働時間（個別・牛乳以外）", "【12】労働時間（経営分析・牛乳）", "【12】労働時間（経営分析・牛乳以外）"}},
                        {ベース用ファイル名称.Report.tempFileNameTikuTsuki, {"設定", "月別入力", "【12】労働時間（個別・牛乳）", "【12】労働時間（個別・牛乳以外）", "【12】労働時間（経営分析・牛乳）", "【12】労働時間（経営分析・牛乳以外）"}}
                }
        End Class

    End Class

    '↓ 2022/02/01 調査票項目指定修正クラス追加
    ''' <summary>
    ''' 調査票項目指定修正クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 調査票項目指定修正
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "調査票項目指定修正.xlsm"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 調査票項目指定修正.入力用ファイル名称
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "項目指定修正"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 5
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 5004
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 調査票項目指定修正.出力用ファイル名称.Row.Last - 調査票項目指定修正.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "G"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "No"} _
                , {2, "調査年（産）"} _
                , {3, "センサス番号"} _
                , {4, "項目番号"} _
                , {5, "修正前データ"} _
                , {6, "修正データ"}
            }
        End Class
    End Class
    '↑ 2022/02/01 調査票項目指定修正クラス追加

#Region "制度受取金・積立金等項目"

    ''' <summary>
    ''' 制度受取金・積立金等項目クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 制度受取金積立金等項目
        ''' <summary>入力用ファイル名称</summary>
        Public Const 入力用ファイル名称 As String = "制度受取金・積立金等項目.xlsm"

        ''' <summary>出力用ファイル名称クラス</summary>
        Public Class 出力用ファイル名称
            ''' <summary>テンプレートファイル名</summary>
            Public Shared tempFileName As String = 制度受取金積立金等項目.入力用ファイル名称
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "制度受取金・積立金等項目"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "制度受取金・積立金等項目"
            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>先頭行</summary>q
                Public Const First As Integer = 5
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 80
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 制度受取金積立金等項目.出力用ファイル名称.Row.Last - 制度受取金積立金等項目.出力用ファイル名称.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>先頭列</summary>
                Public Const First As String = "B"
                ''' <summary>最終列</summary>
                Public Const Last As String = "E"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "項番"} _
                , {2, "出力項目名"} _
                , {3, "制度受取金等項番"} _
                , {4, "制度積立金等項番"}
            }
        End Class

        '調査年取得
        Public Class 制度受取金調査年
            Public Const 調査年2022 As String = "1"
            Public Const 調査年2023 As String = "2"
            Public Const 調査年2024 As String = "3"
            Public Const 調査年2025 As String = "4"
            Public Const 調査年2026 As String = "5"
            Public Const 調査年2027 As String = "6"
            Public Const 調査年2028 As String = "7"
            Public Const 調査年2029 As String = "8"
            Public Const 調査年2030 As String = "9"
            Public Const 調査年2031 As String = "10"

            Public Class 詳細
                Public 調査年 As String
            End Class

            Public Shared リスト As New Dictionary(Of String, 詳細) From {
              {制度受取金調査年.調査年2022, New 詳細 With {.調査年 = "2022"}} _
            , {制度受取金調査年.調査年2023, New 詳細 With {.調査年 = "2023"}} _
            , {制度受取金調査年.調査年2024, New 詳細 With {.調査年 = "2024"}} _
            , {制度受取金調査年.調査年2025, New 詳細 With {.調査年 = "2025"}} _
            , {制度受取金調査年.調査年2026, New 詳細 With {.調査年 = "2026"}} _
            , {制度受取金調査年.調査年2027, New 詳細 With {.調査年 = "2027"}} _
            , {制度受取金調査年.調査年2028, New 詳細 With {.調査年 = "2028"}} _
            , {制度受取金調査年.調査年2029, New 詳細 With {.調査年 = "2029"}} _
            , {制度受取金調査年.調査年2030, New 詳細 With {.調査年 = "2030"}} _
            , {制度受取金調査年.調査年2031, New 詳細 With {.調査年 = "2031"}}
        }
        End Class

        '帳票名
        Public Shared 制度受取金シート名 As New Dictionary(Of String, String) From {
                     {ComConst.調査区分.営農類型別経営統計_個人, "09_共済・補助金等"} _
                   , {ComConst.調査区分.営農類型別経営統計_法人, "11_01_制度受取金・積立金"} _
                   , {ComConst.調査区分.米生産費統計_個別, "【１】経営の概況③_個別_米"} _
                   , {ComConst.調査区分.小麦生産費統計_個別, "【１】経営の概況③_個別_麦・大豆"} _
                   , {ComConst.調査区分.二条大麦生産費統計_個別, "【１】経営の概況③_個別_麦・大豆"} _
                   , {ComConst.調査区分.六条大麦生産費統計_個別, "【１】経営の概況③_個別_麦・大豆"} _
                   , {ComConst.調査区分.はだか麦生産費統計_個別, "【１】経営の概況③_個別_麦・大豆"} _
                   , {ComConst.調査区分.そば生産費統計_個別, "【１】経営の概況③_個別_なたね・そば"} _
                   , {ComConst.調査区分.大豆生産費統計_個別, "【１】経営の概況③_個別_麦・大豆"} _
                   , {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, "【１】経営の概況③_原ばれ"} _
                   , {ComConst.調査区分.なたね生産費統計_個別, "【１】経営の概況③_個別_なたね・そば"} _
                   , {ComConst.調査区分.てんさい生産費統計_個別, "【１】経営の概況③_てんさい"} _
                   , {ComConst.調査区分.さとうきび生産費統計_個別, "【１】経営の概況③_さとうきび"} _
                   , {ComConst.調査区分.米生産費統計_組織法人, "【１】経営の概況③_組織_米"} _
                   , {ComConst.調査区分.小麦生産費統計_組織法人, "【１】経営の概況③_組織_麦・大豆・なたね・そば"} _
                   , {ComConst.調査区分.大豆生産費統計_組織法人, "【１】経営の概況③_組織_麦・大豆・なたね・そば"} _
                   , {ComConst.調査区分.経営分析調査_二条大麦生産費, "【１】経営の概況③_組織_麦・大豆・なたね・そば"} _
                   , {ComConst.調査区分.経営分析調査_六条大麦生産費, "【１】経営の概況③_組織_麦・大豆・なたね・そば"} _
                   , {ComConst.調査区分.経営分析調査_はだか麦生産費, "【１】経営の概況③_組織_麦・大豆・なたね・そば"} _
                   , {ComConst.調査区分.経営分析調査_そば生産費, "【１】経営の概況③_組織_麦・大豆・なたね・そば"} _
                   , {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, "【１】経営の概況③_原ばれ"} _
                   , {ComConst.調査区分.経営分析調査_なたね生産費, "【１】経営の概況③_組織_麦・大豆・なたね・そば"} _
                   , {ComConst.調査区分.経営分析調査_てんさい生産費, "【１】経営の概況③_てんさい"} _
                   , {ComConst.調査区分.経営分析調査_さとうきび生産費, "【１】経営の概況③_さとうきび"}
               }
        '>>>2022/1/27 項番変更
        Public Shared Seidouketorikin_sheetcheck() As String = {"Q11010141", "Q11010142", "Q11010143", "Q11010144", "Q11010145", "Q11010146", "Q11010147", "Q11010148", "Q11010149", "Q11010150", "Q11010151", "Q11010152", "Q11010153", "Q11010154", "Q11010155", "Q11010156", "Q11010157", "Q11010158", "Q11010159", "Q11010160", "Q11010161", "Q11010162", "Q11010163", "Q11010164", "Q11010165", "Q11010166", "Q11010167", "Q11010168", "Q11010169", "Q11010170", "Q11010171", "Q11010172", "Q11010173", "Q11010174", "Q11010175", "Q11010176", "Q11010177", "Q11010178", "Q11010179", "Q11010180", "Q11010181", "Q11010182", "Q11010183", "Q11010184", "Q11010185", "Q11010186", "Q11010241", "Q11010242", "Q11010243", "Q11010244", "Q11010245", "Q11010246", "Q11010247", "Q11010248", "Q11010249", "Q11010250", "Q11010251", "Q11010252", "Q11010253", "Q11010254", "Q11010255", "Q11010256", "Q11010257", "Q11010258", "Q11010259", "Q11010260", "Q11010261", "Q11010262", "Q11010263", "Q11010264", "Q11010265", "Q11010266", "Q11010267", "Q11010268", "Q11010269", "Q11010270", "Q11010271", "Q11010272", "Q11010273", "Q11010274", "Q11010275", "Q11010276", "Q11010277", "Q11010278", "Q11010279", "Q11010280", "Q11010281", "Q11010282", "Q11010283", "Q11010284", "Q11010285", "Q11010286"
}
        '<<<

        'REV 010 START ---------------
        Public Class 任意項目名称
            Public Const 調査票項番 As String = "Q090"
        End Class
        'REV 010 START ---------------

    End Class
#End Region

    'REV 018 ADD START ---------------
#Region "農業地域類型マスタ管理クラス"
    ''' <summary>
    ''' 農業地域類型マスタ管理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 農業地域類型マスタ管理
        Public Const テーブル名称 As String = "農業地域類型区分マスタ"

        Public Class ファイル情報
            ''' <summary>テンプレートファイル名</summary>
            Public Const tempFileName As String = "農業地域類型マスタ一覧.xlsx"
            ''' <summary>帳票名</summary>
            Public Const reportName As String = "農業地域類型マスタ一覧"
            ''' <summary>シート名</summary>
            Public Const SheetName As String = "農業地域類型マスタ一覧"

            ''' <summary>行</summary>
            Public Class Row
                ''' <summary>調査年行</summary>
                Public Const Year As Integer = 5
                ''' <summary>先頭行</summary>
                Public Const First As Integer = 7
                ''' <summary>最終行</summary>
                Public Const Last As Integer = 20006
                ''' <summary>最大行数</summary>
                Public Const Max As Integer = 農業地域類型マスタ管理.ファイル情報.Row.Last - 農業地域類型マスタ管理.ファイル情報.Row.First + 1
            End Class

            ''' <summary>列</summary>
            Public Class Col
                ''' <summary>調査年列</summary>
                Public Const Year As Integer = 7
                ''' <summary>先頭列</summary>
                Public Const First As String = "A"
                ''' <summary>最終列</summary>
                Public Const Last As String = "G"
            End Class

            ''' <summary>フィールド</summary>
            Public Shared Field As New Dictionary(Of Integer, String) From {
                  {1, "都道府県"} _
                , {2, "市区町村"} _
                , {3, "旧市区町村"} _
                , {4, "市区町村名"} _
                , {5, "旧市区町村名"} _
                , {6, "第１次分類"} _
                , {7, "第２次分類"}
            }

        End Class
        ''' <summary>農業地域類型マスタ調査年</summary>
        Public Class 農業地域類型マスタ調査年
            Public Const 農業地域類型マスタ2015 As String = "2015"
            Public Const 農業地域類型マスタ2020 As String = "2020"
            Public Const 農業地域類型マスタ2025 As String = "2025"
            Public Const 農業地域類型マスタ2030 As String = "2030"

            Public Shared リスト As New Dictionary(Of String, String) From {
              {農業地域類型マスタ調査年.農業地域類型マスタ2015, "農業地域類型マスタ2015"} _
            , {農業地域類型マスタ調査年.農業地域類型マスタ2020, "農業地域類型マスタ2020"} _
            , {農業地域類型マスタ調査年.農業地域類型マスタ2025, "農業地域類型マスタ2025"} _
            , {農業地域類型マスタ調査年.農業地域類型マスタ2030, "農業地域類型マスタ2030"}
        }
        End Class
    End Class

#End Region
    'REV 018 ADD END ---------------

    'REV 019↓
    ''' <summary>
    ''' 継続区分クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 継続区分
        ''' <summary>全て</summary>
        Public Const 全て As String = "1"
        ''' <summary>継続のみ</summary>
        Public Const 継続のみ As String = "2"

        Public Shared リスト As New Dictionary(Of String, String) From {
              {継続区分.全て, "全て"} _
            , {継続区分.継続のみ, "継続のみ"}
        }
    End Class
    'REV 019↑

End Class
