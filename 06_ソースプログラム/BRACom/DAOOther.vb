Imports System.Data.SqlClient

''' <summary>
''' その他テーブル操作
''' </summary>
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.10.26 |TSP                 | フェーズ3 要件No.1, 2修正
'//  REV_002   | 2020.11.06 |TSP                 | フェーズ3 要件No.6修正　
'//  REV_003   | 2021.02.16 |TSP                 | 牛トレサ取込にて母牛識別番号が空文字時の変換エラー対応(連絡票No.644）
'//  REV_004   | 2021.11.29 |日本ｺﾝﾋﾟｭｰﾀｼｽﾃﾑ     | 要件No.1-③
'//  REV_005   | 2022.01.05 |日本ｺﾝﾋﾟｭｰﾀｼｽﾃﾑ     | 要件No.2
'//  REV_006   | 2022.11.04 |Daiko               | 要件No.1 バージョン区分追加        
'//  REV_007   | 2022.11.21 |Daiko               | 要件No.1 制度受取金・積立金等項目名称取得処理追加
'//  REV_008   | 2022.12.14 |Daiko               | 要件No.4 バージョン区分追加  
'//  REV_009   | 2022.12.14 |Daiko               | 要件No.15 集計結果検討表（報告用）追加
'//  REV_010   | 2023.01.11 |Daiko               | 要件No.4 制度受取金・積立金等項目名称取得処理追加
'//  REV_011   | 2023.01.12 |Daiko               | 要件No.4 制度受取金・積立金等項目名称取得処理追加（集計結果検討表）
'//  REV_012   | 2023.01.12 |Daiko               | 要件No.6 欠測値補完機能に係る修正
'//  REV_013   | 2023.01.12 |Daiko               | 要件No.8 バージョン区分追加
'//  REV_014   | 2023.04.28 |Daiko               | 変更要件No.3対応
'//  REV_015   | 2023.08.07 |Daiko               | 要件No.13-① 成畜の条件に初回分娩を追加
'//  REV_016   | 2024.05.31 |Daiko               | 要件No.1対応
'//*************************************************************************************************
''' <remarks></remarks>
Public Class DAOOther

    ''' <summary>
    ''' 調査票項目マスタ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="suushikiKubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoItemMaster(db As DBAccess, chosaKubun As String, chosaNen As String, Optional suushikiKubun As String = Nothing) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        'REV_004 START---------------
        'バージョン区分を調査年から指定する
        Dim ver_kubun As String
        ver_kubun = ComUtil.getVersionKubun(chosaNen, chosaKubun)
        'REV_004 END---------------

        Try

            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where Not val.Contains("＿可変")

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT     A.*")
                .AppendLine("         , B.precision AS 有効桁数 ")
                .AppendLine("         , B.scale AS 小数点以下桁数")
                .AppendLine("         , C.項目名 AS 出力項目名")
                .AppendLine("FROM      (SELECT * ")
                .AppendLine("           FROM   調査票項目マスタ ")
                .AppendLine("           WHERE  調査区分 = @調査区分 ")
                'REV_004 START-------------------
                .AppendLine("           AND  バージョン区分 = @バージョン区分 ")
                'REV_004 END-------------------
                If Not suushikiKubun Is Nothing Then
                    .AppendLine("           AND    数式区分 = @数式区分 ")
                End If
                .AppendLine("          ) A")
                .AppendLine("LEFT JOIN (SELECT * ")
                .AppendLine("           FROM   sys.columns")
                .AppendLine("           WHERE  object_id IN (SELECT object_id")
                .AppendLine("                                FROM   sys.tables")
                .AppendLine("                                WHERE  name IN ('" & String.Join("', '", query) & "')")
                .AppendLine("                               )")
                .AppendLine("          ) B")
                .AppendLine("ON A.項目番号 = B.name")
                .AppendLine(" LEFT JOIN 調査票項目名マスタ C ")
                .AppendLine("ON A.項目番号 = C.項番")
                .AppendLine(" AND C.調査区分 = @調査区分 ")     'INS:2022/01/26 NCS_TM Excelチェックがdictionaryの重複でエラーになる

            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            'REV_004 START-------------------
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, ver_kubun))
            'REV_004 END-------------------
            If Not suushikiKubun Is Nothing Then
                para.Add(db.CreateParameter("@数式区分", SqlDbType.Int, suushikiKubun))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票項目マスタ取得（可変のみ）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="suushikiKubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoItemMasterKahen(db As DBAccess, chosaKubun As String, chosaNen As String, Optional suushikiKubun As String = Nothing) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        'バージョン区分を調査年から指定する
        Dim ver_kubun As String
        ver_kubun = ComUtil.getVersionKubun(chosaNen, chosaKubun)

        Try

            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where Not val.Contains("＿可変")

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT     A.*")
                .AppendLine("         , B.precision AS 有効桁数 ")
                .AppendLine("         , B.scale AS 小数点以下桁数")
                .AppendLine("FROM      (SELECT * ")
                .AppendLine("           FROM   調査票項目マスタ ")
                .AppendLine("           WHERE  調査区分 = @調査区分 ")
                .AppendLine("           AND  バージョン区分 = @バージョン区分 ")
                .AppendLine("           AND  可変区分 = '1' ")
                If Not suushikiKubun Is Nothing Then
                    .AppendLine("           AND    数式区分 = @数式区分 ")
                End If
                .AppendLine("          ) A")
                .AppendLine("LEFT JOIN (SELECT * ")
                .AppendLine("           FROM   sys.columns")
                .AppendLine("           WHERE  object_id IN (SELECT object_id")
                .AppendLine("                                FROM   sys.tables")
                .AppendLine("                                WHERE  name IN ('" & String.Join("', '", query) & "')")
                .AppendLine("                               )")
                .AppendLine("          ) B")
                .AppendLine("ON A.項目番号 = B.name")

            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, ver_kubun))
            If Not suushikiKubun Is Nothing Then
                para.Add(db.CreateParameter("@数式区分", SqlDbType.Int, suushikiKubun))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票項目マスタ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="suushikiKubun"></param>
    ''' <returns></returns>
    ''' <remarks>REV 006 ADD</remarks>
    Public Shared Function GetChosahyoItemMasterSeidoUketori(db As DBAccess, chosaKubun As String, chosaNen As String, Optional suushikiKubun As String = Nothing) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        'バージョン区分を調査年から指定する
        Dim ver_kubun As String
        ver_kubun = ComUtil.getVersionKubun(chosaNen, chosaKubun)

        Try

            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where Not val.Contains("＿可変")

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT     A.*")
                .AppendLine("         , B.precision AS 有効桁数 ")
                .AppendLine("         , B.scale AS 小数点以下桁数")
                .AppendLine("         , C.出力項目名 AS 出力項目名")
                .AppendLine("FROM      (SELECT * ")
                .AppendLine("           FROM   調査票項目マスタ ")
                .AppendLine("           WHERE  調査区分 = @調査区分 ")
                .AppendLine("           AND  バージョン区分 = @バージョン区分 ")
                If Not suushikiKubun Is Nothing Then
                    .AppendLine("           AND    数式区分 = @数式区分 ")
                End If
                .AppendLine("          ) A")
                .AppendLine("LEFT JOIN (SELECT * ")
                .AppendLine("           FROM   sys.columns")
                .AppendLine("           WHERE  object_id IN (SELECT object_id")
                .AppendLine("                                FROM   sys.tables")
                .AppendLine("                                WHERE  name IN ('" & String.Join("', '", query) & "')")
                .AppendLine("                               )")
                .AppendLine("          ) B")
                .AppendLine("ON A.項目番号 = B.name")
                .AppendLine("LEFT JOIN (SELECT * ")
                .AppendLine("           FROM   制度受取金・積立金等項目")
                .AppendLine("           WHERE  調査年 = '" & chosaNen & "'")
                .AppendLine("           AND    調査区分 = '" & chosaKubun & "'")
                .AppendLine("          ) C")
                .AppendLine("ON A.項目番号 = C.項番")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, ver_kubun))
            If Not suushikiKubun Is Nothing Then
                para.Add(db.CreateParameter("@数式区分", SqlDbType.Int, suushikiKubun))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表項目マスタ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="isIncludeUra">裏項番を含める場合はTrue</param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkahyoItemMaster(db As DBAccess, chosaKubun As String, Optional isIncludeUra As Boolean = False) As DataTable
    Public Shared Function GetKobetsuKekkahyoItemMaster(db As DBAccess, chosaKubun As String, versionKbn As String, Optional isIncludeUra As Boolean = False, Optional chosaNen As String = Nothing) As DataTable
        ' REV_006 REV_007↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT     A.*")
                .AppendLine("         , B.precision AS 有効桁数 ")
                .AppendLine("         , B.scale AS 小数点以下桁数")
                .AppendLine("         , C.表示単位 AS 表示単位")
                .AppendLine("         , C.再計算区分 AS 再計算区分")
                'REV_007 START---------------
                '出力項目名を取得する
                If chosaNen <> "" Then
                    .AppendLine("         , D.出力項目名 AS 出力項目名")
                End If
                'REV_007 END---------------
                .AppendLine("FROM      (SELECT * ")
                .AppendLine("           FROM   個別結果表項目マスタ ")
                .AppendLine("           WHERE  調査区分 = @調査区分 ")
                .AppendLine("           AND  バージョン区分 = @バージョン区分 ")        ' REV_006
                If Not isIncludeUra Then
                    .AppendLine("           AND    裏項番区分 = 0 ")
                End If
                .AppendLine("          ) A")
                .AppendLine("LEFT JOIN (SELECT * ")
                .AppendLine("           FROM   sys.columns")
                .AppendLine("           WHERE  object_id IN (SELECT object_id")
                .AppendLine("                                FROM   sys.tables")
                .AppendLine("                                WHERE  name IN ('" & String.Join("', '", ComConst.個別結果表.テーブル名称(chosaKubun)) & "')")
                .AppendLine("                               )")
                .AppendLine("          ) B")
                .AppendLine("ON  A.項目番号 = B.name")
                .AppendLine("LEFT JOIN (SELECT * ")
                If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                    .AppendLine("           FROM   個別結果表作成論理＿営農個人 ")
                    .AppendLine("           WHERE  調査区分 = @調査区分 ")
                    .AppendLine("           AND    貸借対照表区分 IN (0, 1) ")
                Else
                    .AppendLine("           FROM   個別結果表作成論理 ")
                    .AppendLine("           WHERE  調査区分 = @調査区分 ")
                End If
                .AppendLine("           AND  バージョン区分 = @バージョン区分 ")        ' REV_006
                .AppendLine("          ) C")
                .AppendLine("ON  A.調査区分 = C.調査区分")
                .AppendLine("AND A.項目番号 = C.項目番号")
                'REV_007 START---------------
                '制度受取金・積立金等項目を取得する
                If chosaNen <> "" Then
                    .AppendLine("LEFT JOIN (SELECT * ")
                    .AppendLine("FROM   制度受取金・積立金等項目")
                    .AppendLine("WHERE  調査年 = @調査年")
                    .AppendLine("AND    調査区分 = @調査区分")
                    .AppendLine(") D")
                    .AppendLine("ON C.計算式 = '[' + D.項番 + ']'")
                    'REV_007 END---------------
                End If


            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, chosaNen))   ' REV_007

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表項目マスタ取得(農業経営体の項番存在チェックに使用)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetNougyouKeieitaiKobetsuKekkahyoItemMaster(db As DBAccess) As DataTable
    Public Shared Function GetNougyouKeieitaiKobetsuKekkahyoItemMaster(db As DBAccess, versionKbn As String) As DataTable
        ' REV_006↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   個別結果表項目マスタ ")
                .AppendLine("WHERE  調査区分 = 1 ")
                .AppendLine("AND    シート名 in ('No4', 'No5', 'No6') ")
                .AppendLine("AND    裏項番区分 = 0 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
            End With

            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006
            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果検討表項目マスタ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkaKentohyoItemMaster(db As DBAccess, chosaKubun As String) As DataTable
    Public Shared Function GetKobetsuKekkaKentohyoItemMaster(db As DBAccess, chosaKubun As String, versionKbn As String) As DataTable
        ' REV_006↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   個別結果検討表項目マスタ ")
                .AppendLine("WHERE  調査区分 = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 集計結果表項目マスタ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetSyukeiItemMaster(db As DBAccess, chosaKubun As String) As DataTable
    Public Shared Function GetSyukeiItemMaster(db As DBAccess, chosaKubun As String, versionKbn As String) As DataTable
        ' REV_008↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   集計結果表項目マスタ ")
                .AppendLine("WHERE  調査区分 = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_008
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 集計結果表項目マスタ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="seisanhiHeikin"></param>
    ''' <param name="isIncludeUra">裏項番を含める場合はTrue</param>
    ''' <param name="isIncludeOut">出力しないを含める場合はTrue</param>
    ''' <returns></returns>
    'Public Shared Function GetSyukeiKekkahyoItemMaster(db As DBAccess, chosaKubun As String, seisanhiHeikin As String, Optional isIncludeUra As Boolean = False, Optional isIncludeOut As Boolean = False) As DataTable
    Public Shared Function GetSyukeiKekkahyoItemMaster(db As DBAccess, chosaKubun As String, versionKbn As String, seisanhiHeikin As String, Optional isIncludeUra As Boolean = False, Optional isIncludeOut As Boolean = False, Optional chosaNen As String = Nothing) As DataTable
        ' REV_008,REV_010↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT     A.*")
                .AppendLine("         , B.表示単位 AS 表示単位")
                'REV_010 START---------------
                '出力項目名を取得する
                If chosaNen <> "" Then
                    .AppendLine("         , C.出力項目名 AS 出力項目名")
                End If
                'REV_010 END---------------
                .AppendLine("FROM      (SELECT * ")
                .AppendLine("           FROM   集計結果表項目マスタ ")
                .AppendLine("           WHERE  調査区分 = @調査区分 ")
                .AppendLine("           AND    バージョン区分 = @バージョン区分 ")        ' REV_008
                If Not isIncludeUra Then
                    .AppendLine("           AND    裏項番区分 = 0 ")
                End If
                If Not isIncludeOut Then
                    .AppendLine("           AND    出力区分   = 0 ")
                End If
                .AppendLine("          ) A")
                .AppendLine("LEFT JOIN (SELECT * ")
                .AppendLine("           FROM   集計結果表作成論理 ")
                .AppendLine("           WHERE  調査区分       = @調査区分 ")
                .AppendLine("           AND    バージョン区分 = @バージョン区分 ")        ' REV_008
                .AppendLine("           AND    生産費平均値種類 = @生産費平均値種類 ")
                .AppendLine("          ) B")
                .AppendLine("ON  A.調査区分 = B.調査区分")
                .AppendLine("AND A.項目番号 = B.項目番号")
                'REV_010 START---------------
                '制度受取金・積立金等項目を取得する
                If chosaNen <> "" Then
                    .AppendLine("LEFT JOIN (SELECT * ")
                    .AppendLine("FROM   制度受取金・積立金等項目")
                    .AppendLine("WHERE  調査年 = @調査年")
                    .AppendLine("AND    調査区分 = @調査区分")
                    .AppendLine(") C")
                    .AppendLine("ON TRIM(REPLACE(B.計算式,'''','')) = C.項番")
                End If
                'REV_010 END---------------
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, chosaNen))   ' REV_010

            Dim sHeikin As Integer = 0

            If Not (ComConst.営農経営体区分.リスト.ContainsKey(chosaKubun)) Then
                Select Case seisanhiHeikin
                    Case "1"
                        sHeikin = 1
                    Case Else
                        sHeikin = 9
                End Select
            End If
            para.Add(db.CreateParameter("@生産費平均値種類", SqlDbType.Int, sHeikin))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 集計結果検討表項目マスタ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetSyukeiKekkaKentohyoItemMaster(db As DBAccess, chosaKubun As String) As DataTable
    Public Shared Function GetSyukeiKekkaKentohyoItemMaster(db As DBAccess, chosaKubun As String, versionKbn As String) As DataTable
        ' REV_008↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   集計結果検討表項目マスタ ")
                .AppendLine("WHERE  調査区分 = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_008
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_009↓
    ''' <summary>
    ''' 集計結果検討表＿報告用＿項目マスタ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    Public Shared Function GetSyukeiKekkaKentohyoHoukokuyoItemMaster(db As DBAccess, chosaKubun As String, versionKbn As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   集計結果検討表＿報告用＿項目マスタ ")
                .AppendLine("WHERE  調査区分 = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function
    ' REV_009↑

    ''' <summary>
    ''' 調査票審査論理更新日時取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="errType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoShinsaRonriUpdateDate(db As DBAccess, errType As ComConst.エラーチェック種別.enm) As String

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine(String.Format("FROM   調査票審査論理＿{0}", ComConst.エラーチェック種別.リスト(errType)))
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                If errType = ComConst.エラーチェック種別.enm.追加 Then
                    .AppendLine("AND    農政局         = @農政局 ")
                    .AppendLine("AND    都道府県       = @都道府県 ")
                    .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                End If
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            If errType = ComConst.エラーチェック種別.enm.追加 Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            End If

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn
    End Function

    ''' <summary>
    ''' 調査票審査論理取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="errType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoShinsaRonri(db As DBAccess, errType As ComConst.エラーチェック種別.enm) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine(String.Format("FROM   調査票審査論理＿{0}", ComConst.エラーチェック種別.リスト(errType)))
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                If errType = ComConst.エラーチェック種別.enm.追加 Then
                    .AppendLine("AND    農政局         = @農政局 ")
                    .AppendLine("AND    都道府県       = @都道府県 ")
                    .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                End If
                .AppendLine("ORDER BY エラーサイン")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            If errType = ComConst.エラーチェック種別.enm.追加 Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票審査論理追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertChosahyoShinsaRonri(db As DBAccess, dc As Dictionary(Of String, String), errType As ComConst.エラーチェック種別.enm) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine(String.Format("INSERT INTO 調査票審査論理＿{0} ", ComConst.エラーチェック種別.リスト(errType)))
                .AppendLine("( ")
                .AppendLine("   調査区分 ")
                If errType = ComConst.エラーチェック種別.enm.追加 Then
                    .AppendLine("  ,農政局 ")
                    .AppendLine("  ,都道府県 ")
                    .AppendLine("  ,実査設置拠点 ")
                End If
                .AppendLine("  ,エラーサイン ")
                .AppendLine("  ,チェック項目名 ")
                .AppendLine("  ,エラー内容 ")
                .AppendLine("  ,エラーとなる条件 ")
                .AppendLine("  ,繰り返し ")
                .AppendLine("  ,エラー区分 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @調査区分 ")
                If errType = ComConst.エラーチェック種別.enm.追加 Then
                    .AppendLine("  ,@農政局 ")
                    .AppendLine("  ,@都道府県 ")
                    .AppendLine("  ,@実査設置拠点 ")
                End If
                .AppendLine("  ,@エラーサイン ")
                .AppendLine("  ,@チェック項目名 ")
                .AppendLine("  ,@エラー内容 ")
                .AppendLine("  ,@エラーとなる条件 ")
                .AppendLine("  ,@繰り返し ")
                .AppendLine("  ,@エラー区分 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            If errType = ComConst.エラーチェック種別.enm.追加 Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            End If
            para.Add(db.CreateParameter("@エラーサイン", SqlDbType.VarChar, dc("エラーサイン")))
            para.Add(db.CreateParameter("@チェック項目名", SqlDbType.VarChar, dc("チェック項目名")))
            para.Add(db.CreateParameter("@エラー内容", SqlDbType.VarChar, dc("エラー内容")))
            para.Add(db.CreateParameter("@エラーとなる条件", SqlDbType.VarChar, dc("エラーとなる条件")))
            para.Add(db.CreateParameter("@繰り返し", SqlDbType.VarChar, dc("繰り返し")))
            para.Add(db.CreateParameter("@エラー区分", SqlDbType.VarChar, dc("エラー区分")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("調査票審査論理追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票審査論理削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteChosahyoShinsaRonri(db As DBAccess, errType As ComConst.エラーチェック種別.enm) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine(String.Format("FROM   調査票審査論理＿{0}", ComConst.エラーチェック種別.リスト(errType)))
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                If errType = ComConst.エラーチェック種別.enm.追加 Then
                    .AppendLine("AND    農政局         = @農政局 ")
                    .AppendLine("AND    都道府県       = @都道府県 ")
                    .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                End If
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            If errType = ComConst.エラーチェック種別.enm.追加 Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            End If

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("調査票審査論理削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表作成論理更新日時取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkahyoSakuseiRonriUpdateDate(db As DBAccess) As String
    Public Shared Function GetKobetsuKekkahyoSakuseiRonriUpdateDate(db As DBAccess, versionKbn As String) As String
        ' REV_006↑

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine("FROM   個別結果表作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表作成論理(営農個人、貸借対照表区分)を取得する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Private Shared Function GetKobetsuKekkahyoSakuseiRonriEinokojin(db As DBAccess) As DataTable
    Private Shared Function GetKobetsuKekkahyoSakuseiRonriEinokojin(db As DBAccess, versionKbn As String) As DataTable
        ' REV_006↑
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   個別結果表作成論理＿営農個人")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    項目番号       = @項目番号 ")
                .AppendLine("AND    貸借対照表区分 = 0 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, ComConst.調査区分.営農類型別経営統計_個人))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, ComConst.個別結果表.貸借対照表(ComConst.調査区分.営農類型別経営統計_個人)))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 営農個人の貸借対照表区分を取得する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="dtChoItemMst"></param>
    ''' <param name="dtKobetsuItemMst"></param>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTaisyakuKubunEinokojin(db As DBAccess, chosaNen As String, dtChoItemMst As DataTable, dtKobetsuItemMst As DataTable, censusNo As String) As String
        Dim ret As String = Nothing
        Dim dtCreateRonri As DataTable
        Dim itemInfoList As List(Of CreateKobetsu.ItemInfo)
        Dim ItemInfoListKobetsu As List(Of CreateKobetsu.ItemInfo)

        '個別結果表作成論理(営農個人、貸借対照表区分)を取得する
        ' REV_006↓
        'dtCreateRonri = GetKobetsuKekkahyoSakuseiRonriEinokojin(db)
        dtCreateRonri = GetKobetsuKekkahyoSakuseiRonriEinokojin(db, ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun))
        ' REV_006↑
        '個別結果表・個別結果検討表作成クラス
        Dim kobetsu As CreateKobetsu = New CreateKobetsu(db, ComConst.調査区分.営農類型別経営統計_個人,
                                                         chosaNen,
                                                         CreateKobetsu.enmCreateType.個別結果表作成,
                                                         dtChoItemMst,
                                                         dtKobetsuItemMst,
                                                         Nothing,
                                                         dtCreateRonri,
                                                         Nothing,
                                                         Nothing)
        '個別結果表作成実行
        itemInfoList = kobetsu.Execute(censusNo, True) 'REV_016 CHG DAIKO
        '個別結果表(当年データ、裏項番以外)で抽出
        ItemInfoListKobetsu = (From n In itemInfoList Where n.ItemType = CreateKobetsu.enmItemType.個別結果表 _
                                                      And n.ItemNo = ComConst.個別結果表.貸借対照表(ComConst.調査区分.営農類型別経営統計_個人) _
                                                      And Not n.ItemNo.Contains("前") _
                                                      And n.IsHidden = False).ToList
        If ItemInfoListKobetsu.Count > 0 Then
            ret = If(ItemInfoListKobetsu(0).Value Is Nothing, "0", CStr(ItemInfoListKobetsu(0).Value))
        End If

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表作成論理取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="isOnlyReCalc">再計算する論理のみ取得する場合にTrueを指定する</param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkahyoSakuseiRonri(db As DBAccess, Optional ByVal isOnlyReCalc As Boolean = False) As DataTable
    Public Shared Function GetKobetsuKekkahyoSakuseiRonri(db As DBAccess, versionKbn As String, Optional ByVal isOnlyReCalc As Boolean = False) As DataTable
        ' REV_006↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   個別結果表作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                If isOnlyReCalc Then
                    .AppendLine("AND    再計算区分     = '○' ")
                End If
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
                .AppendLine("ORDER BY 項目番号")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表作成論理追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    'Public Shared Function InsertKobetsuKekkahyoSakuseiRonri(db As DBAccess, dc As Dictionary(Of String, String)) As Boolean
    Public Shared Function InsertKobetsuKekkahyoSakuseiRonri(db As DBAccess, versionKbn As String, dc As Dictionary(Of String, String)) As Boolean
        ' REV_006↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 個別結果表作成論理 ")
                .AppendLine("( ")
                .AppendLine("   バージョン区分 ")        ' REV_006
                .AppendLine("  ,調査区分 ")
                .AppendLine("  ,項目番号 ")
                .AppendLine("  ,項目名 ")
                .AppendLine("  ,優先順位 ")
                .AppendLine("  ,計算式 ")
                .AppendLine("  ,表示単位 ")
                .AppendLine("  ,再計算区分 ")
                .AppendLine("  ,備考 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @バージョン区分 ")        ' REV_006
                .AppendLine("  ,@調査区分 ")
                .AppendLine("  ,@項目番号 ")
                .AppendLine("  ,@項目名 ")
                .AppendLine("  ,@優先順位 ")
                .AppendLine("  ,@計算式 ")
                .AppendLine("  ,@表示単位 ")
                .AppendLine("  ,@再計算区分 ")
                .AppendLine("  ,@備考 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))        ' REV_006
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, dc("項目番号")))
            para.Add(db.CreateParameter("@項目名", SqlDbType.VarChar, dc("項目名")))
            para.Add(db.CreateParameter("@優先順位", SqlDbType.Decimal, dc("優先順位")))
            para.Add(db.CreateParameter("@計算式", SqlDbType.VarChar, dc("計算式")))
            para.Add(db.CreateParameter("@表示単位", SqlDbType.Decimal, dc("表示単位")))
            para.Add(db.CreateParameter("@再計算区分", SqlDbType.VarChar, dc("再計算区分")))
            para.Add(db.CreateParameter("@備考", SqlDbType.VarChar, dc("備考")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("個別結果表作成論理追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表作成論理削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function DeleteKobetsuKekkahyoSakuseiRonri(db As DBAccess) As Boolean
    Public Shared Function DeleteKobetsuKekkahyoSakuseiRonri(db As DBAccess, versionKbn As String) As Boolean
        ' REV_006↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   個別結果表作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("個別結果表作成論理削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表作成論理＿営農個人更新日時取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkahyoSakuseiRonriEinouKobetsuUpdateDate(db As DBAccess) As String
    Public Shared Function GetKobetsuKekkahyoSakuseiRonriEinouKobetsuUpdateDate(db As DBAccess, versionKbn As String) As String
        ' REV_006↑

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine("FROM   個別結果表作成論理＿営農個人")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表作成論理＿営農個人取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="isOnlyReCalc">再計算する論理のみ取得する場合にTrueを指定する</param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db As DBAccess, Optional ByVal isOnlyReCalc As Boolean = False) As DataTable
    Public Shared Function GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db As DBAccess, versionKbn As String, Optional ByVal isOnlyReCalc As Boolean = False) As DataTable
        ' REV_006↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   個別結果表作成論理＿営農個人")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                If isOnlyReCalc Then
                    .AppendLine("AND    再計算区分     = '○' ")
                End If
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
                .AppendLine("ORDER BY 項目番号")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表作成論理＿営農個人取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="taisyakuKubun">貸借対照表区分</param>
    ''' <param name="versionKbn"></param>
    ''' <param name="isOnlyReCalc">再計算する論理のみ取得する場合にTrueを指定する</param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db As DBAccess, ByVal taisyakuKubun As String, Optional ByVal isOnlyReCalc As Boolean = False) As DataTable
    Public Shared Function GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db As DBAccess, ByVal taisyakuKubun As String, versionKbn As String, Optional ByVal isOnlyReCalc As Boolean = False) As DataTable
        ' REV_006↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   個別結果表作成論理＿営農個人")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                If taisyakuKubun Is Nothing OrElse Not Integer.TryParse(taisyakuKubun, New Integer) Then
                    .AppendLine("AND    貸借対照表区分 IN (0)")
                Else
                    .AppendLine("AND    貸借対照表区分 IN (0," & taisyakuKubun & ")")
                End If
                If isOnlyReCalc Then
                    .AppendLine("AND    再計算区分     = '○' ")
                End If
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
                .AppendLine("ORDER BY 項目番号")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表作成論理＿営農個人追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    'Public Shared Function InsertKobetsuKekkahyoSakuseiRonriEinouKobetsu(db As DBAccess, dc As Dictionary(Of String, String)) As Boolean
    Public Shared Function InsertKobetsuKekkahyoSakuseiRonriEinouKobetsu(db As DBAccess, versionKbn As String, dc As Dictionary(Of String, String)) As Boolean
        ' REV_006↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 個別結果表作成論理＿営農個人 ")
                .AppendLine("( ")
                .AppendLine("   バージョン区分 ")        ' REV_006
                .AppendLine("  ,調査区分 ")
                .AppendLine("  ,項目番号 ")
                .AppendLine("  ,貸借対照表区分 ")
                .AppendLine("  ,項目名 ")
                .AppendLine("  ,優先順位 ")
                .AppendLine("  ,計算式 ")
                .AppendLine("  ,表示単位 ")
                .AppendLine("  ,再計算区分 ")
                .AppendLine("  ,備考 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @バージョン区分 ")        ' REV_006
                .AppendLine("  ,@調査区分 ")
                .AppendLine("  ,@項目番号 ")
                .AppendLine("  ,@貸借対照表区分 ")
                .AppendLine("  ,@項目名 ")
                .AppendLine("  ,@優先順位 ")
                .AppendLine("  ,@計算式 ")
                .AppendLine("  ,@表示単位 ")
                .AppendLine("  ,@再計算区分 ")
                .AppendLine("  ,@備考 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))        ' REV_006
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, dc("項目番号")))
            para.Add(db.CreateParameter("@貸借対照表区分", SqlDbType.Int, dc("貸借対照表区分")))
            para.Add(db.CreateParameter("@項目名", SqlDbType.VarChar, dc("項目名")))
            para.Add(db.CreateParameter("@優先順位", SqlDbType.Decimal, dc("優先順位")))
            para.Add(db.CreateParameter("@計算式", SqlDbType.VarChar, dc("計算式")))
            para.Add(db.CreateParameter("@表示単位", SqlDbType.Decimal, dc("表示単位")))
            para.Add(db.CreateParameter("@再計算区分", SqlDbType.VarChar, dc("再計算区分")))
            para.Add(db.CreateParameter("@備考", SqlDbType.VarChar, dc("備考")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("個別結果表作成論理＿営農個人追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表作成論理＿営農個人削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function DeleteKobetsuKekkahyoSakuseiRonriEinouKobetsu(db As DBAccess) As Boolean
    Public Shared Function DeleteKobetsuKekkahyoSakuseiRonriEinouKobetsu(db As DBAccess, versionKbn As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   個別結果表作成論理＿営農個人")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("個別結果表作成論理＿営農個人削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果検討表作成論理更新日時取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkaKentohyoSakuseiRonriUpdateDate(db As DBAccess) As String
    Public Shared Function GetKobetsuKekkaKentohyoSakuseiRonriUpdateDate(db As DBAccess, versionKbn As String) As String
        ' REV_006↑

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine("FROM   個別結果検討表作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果検討表作成論理取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkaKentohyoSakuseiRonri(db As DBAccess) As DataTable
    Public Shared Function GetKobetsuKekkaKentohyoSakuseiRonri(db As DBAccess, versionKbn As String) As DataTable
        ' REV_006↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   個別結果検討表作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
                .AppendLine("ORDER BY 項目番号")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果検討表作成論理追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    'Public Shared Function InsertKobetsuKekkaKentohyoSakuseiRonri(db As DBAccess, dc As Dictionary(Of String, String)) As Boolean
    Public Shared Function InsertKobetsuKekkaKentohyoSakuseiRonri(db As DBAccess, versionKbn As String, dc As Dictionary(Of String, String)) As Boolean
        ' REV_006↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 個別結果検討表作成論理 ")
                .AppendLine("( ")
                .AppendLine("   バージョン区分 ")        ' REV_006
                .AppendLine("  ,調査区分 ")
                .AppendLine("  ,項目番号 ")
                .AppendLine("  ,項目名 ")
                .AppendLine("  ,優先順位 ")
                .AppendLine("  ,計算式 ")
                .AppendLine("  ,表示単位 ")
                .AppendLine("  ,備考 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @バージョン区分 ")        ' REV_006
                .AppendLine("  ,@調査区分 ")
                .AppendLine("  ,@項目番号 ")
                .AppendLine("  ,@項目名 ")
                .AppendLine("  ,@優先順位 ")
                .AppendLine("  ,@計算式 ")
                .AppendLine("  ,@表示単位 ")
                .AppendLine("  ,@備考 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))        ' REV_006
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, dc("項目番号")))
            para.Add(db.CreateParameter("@項目名", SqlDbType.VarChar, dc("項目名")))
            para.Add(db.CreateParameter("@優先順位", SqlDbType.Decimal, dc("優先順位")))
            para.Add(db.CreateParameter("@計算式", SqlDbType.VarChar, dc("計算式")))
            para.Add(db.CreateParameter("@表示単位", SqlDbType.Decimal, dc("表示単位")))
            para.Add(db.CreateParameter("@備考", SqlDbType.VarChar, dc("備考")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("個別結果検討表作成論理追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果検討表作成論理削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    'Public Shared Function DeleteKobetsuKekkaKentohyoSakuseiRonri(db As DBAccess) As Boolean
    Public Shared Function DeleteKobetsuKekkaKentohyoSakuseiRonri(db As DBAccess, versionKbn As String) As Boolean
        ' REV_006↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   個別結果検討表作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("個別結果検討表作成論理削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表審査論理更新日時取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="errType"></param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkahyoShinsaRonriUpdateDate(db As DBAccess, errType As ComConst.エラーチェック種別.enm) As String
    Public Shared Function GetKobetsuKekkahyoShinsaRonriUpdateDate(db As DBAccess, versionKbn As String, errType As ComConst.エラーチェック種別.enm) As String
        ' REV_006↑

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine(String.Format("FROM   個別結果表審査論理＿{0}", ComConst.エラーチェック種別.リスト(errType)))
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
                If errType = ComConst.エラーチェック種別.enm.追加 Then
                    .AppendLine("AND    農政局         = @農政局 ")
                    .AppendLine("AND    都道府県       = @都道府県 ")
                    .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                End If
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006
            If errType = ComConst.エラーチェック種別.enm.追加 Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            End If

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表審査論理取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="errType"></param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkahyoShinsaRonri(db As DBAccess, errType As ComConst.エラーチェック種別.enm) As DataTable
    Public Shared Function GetKobetsuKekkahyoShinsaRonri(db As DBAccess, versionKbn As String, errType As ComConst.エラーチェック種別.enm) As DataTable
        ' REV_006↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine(String.Format("FROM   個別結果表審査論理＿{0}", ComConst.エラーチェック種別.リスト(errType)))
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
                If errType = ComConst.エラーチェック種別.enm.追加 Then
                    .AppendLine("AND    農政局         = @農政局 ")
                    .AppendLine("AND    都道府県       = @都道府県 ")
                    .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                End If
                .AppendLine("ORDER BY エラーサイン")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006
            If errType = ComConst.エラーチェック種別.enm.追加 Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表審査論理追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertKobetsuKekkahyoShinsaRonri(db As DBAccess, versionKbn As String, dc As Dictionary(Of String, String), errType As ComConst.エラーチェック種別.enm) As Boolean
        ' REV_006↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine(String.Format("INSERT INTO 個別結果表審査論理＿{0} ", ComConst.エラーチェック種別.リスト(errType)))
                .AppendLine("( ")
                .AppendLine("   バージョン区分 ")        ' REV_006
                .AppendLine("  ,調査区分 ")
                If errType = ComConst.エラーチェック種別.enm.追加 Then
                    .AppendLine("  ,農政局 ")
                    .AppendLine("  ,都道府県 ")
                    .AppendLine("  ,実査設置拠点 ")
                End If
                .AppendLine("  ,エラーサイン ")
                .AppendLine("  ,チェック項目名 ")
                .AppendLine("  ,エラー内容 ")
                .AppendLine("  ,エラーとなる条件 ")
                .AppendLine("  ,エラー区分 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @バージョン区分 ")        ' REV_006
                .AppendLine("  ,@調査区分 ")
                If errType = ComConst.エラーチェック種別.enm.追加 Then
                    .AppendLine("  ,@農政局 ")
                    .AppendLine("  ,@都道府県 ")
                    .AppendLine("  ,@実査設置拠点 ")
                End If
                .AppendLine("  ,@エラーサイン ")
                .AppendLine("  ,@チェック項目名 ")
                .AppendLine("  ,@エラー内容 ")
                .AppendLine("  ,@エラーとなる条件 ")
                .AppendLine("  ,@エラー区分 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            If errType = ComConst.エラーチェック種別.enm.追加 Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            End If
            para.Add(db.CreateParameter("@エラーサイン", SqlDbType.VarChar, dc("エラーサイン")))
            para.Add(db.CreateParameter("@チェック項目名", SqlDbType.VarChar, dc("チェック項目名")))
            para.Add(db.CreateParameter("@エラー内容", SqlDbType.VarChar, dc("エラー内容")))
            para.Add(db.CreateParameter("@エラーとなる条件", SqlDbType.VarChar, dc("エラーとなる条件")))
            para.Add(db.CreateParameter("@エラー区分", SqlDbType.VarChar, dc("エラー区分")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("個別結果表審査論理追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表審査論理削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="errType"></param>
    ''' <returns></returns>
    'Public Shared Function DeleteKobetsuKekkahyoShinsaRonri(db As DBAccess, errType As ComConst.エラーチェック種別.enm) As Boolean
    Public Shared Function DeleteKobetsuKekkahyoShinsaRonri(db As DBAccess, versionKbn As String, errType As ComConst.エラーチェック種別.enm) As Boolean
        ' REV_006↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine(String.Format("FROM   個別結果表審査論理＿{0}", ComConst.エラーチェック種別.リスト(errType)))
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
                If errType = ComConst.エラーチェック種別.enm.追加 Then
                    .AppendLine("AND    農政局         = @農政局 ")
                    .AppendLine("AND    都道府県       = @都道府県 ")
                    .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                End If
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006
            If errType = ComConst.エラーチェック種別.enm.追加 Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            End If

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("個別結果表審査論理削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表審査論理範囲更新日時取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkahyoShinsaRonriRangeUpdateDate(db As DBAccess) As String
    Public Shared Function GetKobetsuKekkahyoShinsaRonriRangeUpdateDate(db As DBAccess, versionKbn As String) As String
        ' REV_006↑

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine("FROM   個別結果表審査論理＿範囲")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
                .AppendLine("AND    農政局         = @農政局 ")
                .AppendLine("AND    都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表審査論理範囲取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetKobetsuKekkahyoShinsaRonriRange(db As DBAccess) As DataTable
    Public Shared Function GetKobetsuKekkahyoShinsaRonriRange(db As DBAccess, versionKbn As String) As DataTable
        ' REV_006↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   個別結果表審査論理＿範囲")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
                .AppendLine("AND    農政局         = @農政局 ")
                .AppendLine("AND    都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                .AppendLine("ORDER BY 連番")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表審査論理範囲追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    'Public Shared Function InsertKobetsuKekkahyoShinsaRonriRange(db As DBAccess, dc As Dictionary(Of String, String)) As Boolean
    Public Shared Function InsertKobetsuKekkahyoShinsaRonriRange(db As DBAccess, versionKbn As String, dc As Dictionary(Of String, String)) As Boolean
        ' REV_006↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 個別結果表審査論理＿範囲 ")
                .AppendLine("( ")
                .AppendLine("   バージョン区分 ")        ' REV_006
                .AppendLine("  ,調査区分 ")
                .AppendLine("  ,農政局 ")
                .AppendLine("  ,都道府県 ")
                .AppendLine("  ,実査設置拠点 ")
                .AppendLine("  ,連番 ")
                .AppendLine("  ,チェック項目名 ")
                .AppendLine("  ,項目番号１ ")
                .AppendLine("  ,項目番号２ ")
                .AppendLine("  ,値 ")
                .AppendLine("  ,上限 ")
                .AppendLine("  ,下限 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @バージョン区分 ")        ' REV_006
                .AppendLine("  ,@調査区分 ")
                .AppendLine("  ,@農政局 ")
                .AppendLine("  ,@都道府県 ")
                .AppendLine("  ,@実査設置拠点 ")
                .AppendLine("  ,@連番 ")
                .AppendLine("  ,@チェック項目名 ")
                .AppendLine("  ,@項目番号１ ")
                .AppendLine("  ,@項目番号２ ")
                .AppendLine("  ,@値 ")
                .AppendLine("  ,@上限 ")
                .AppendLine("  ,@下限 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            para.Add(db.CreateParameter("@連番", SqlDbType.Int, dc("連番")))
            para.Add(db.CreateParameter("@チェック項目名", SqlDbType.VarChar, dc("チェック項目名")))
            para.Add(db.CreateParameter("@項目番号１", SqlDbType.VarChar, dc("項目番号１")))
            para.Add(db.CreateParameter("@項目番号２", SqlDbType.VarChar, dc("項目番号２")))
            para.Add(db.CreateParameter("@値", SqlDbType.Decimal, dc("値")))
            para.Add(db.CreateParameter("@上限", SqlDbType.Decimal, dc("上限")))
            para.Add(db.CreateParameter("@下限", SqlDbType.Decimal, dc("下限")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("個別結果表審査論理＿範囲追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_006↓
    ''' <summary>
    ''' 個別結果表審査論理範囲削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function DeleteKobetsuKekkahyoShinsaRonriRange(db As DBAccess) As Boolean
    Public Shared Function DeleteKobetsuKekkahyoShinsaRonriRange(db As DBAccess, versionKbn As String) As Boolean
        ' REV_006↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   個別結果表審査論理＿範囲")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_006
                .AppendLine("AND    農政局         = @農政局 ")
                .AppendLine("AND    都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_006
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("個別結果表審査論理＿範囲削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 集計結果表作成論理更新日時取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    'Public Shared Function GetSyukeiKekkahyoSakuseiRonriUpdateDate(db As DBAccess, chosakubun As String) As String
    Public Shared Function GetSyukeiKekkahyoSakuseiRonriUpdateDate(db As DBAccess, chosakubun As String, versionKbn As String) As String
        ' REV_008↑

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine("FROM   集計結果表作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_008
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 集計結果表作成論理取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="seisanhiHeikin"></param>
    ''' <returns></returns>
    'Public Shared Function GetSyukeiKekkahyoSakuseiRonri(db As DBAccess, chosakubun As String, Optional seisanhiHeikin As String = Nothing) As DataTable
    Public Shared Function GetSyukeiKekkahyoSakuseiRonri(db As DBAccess, chosakubun As String, versionKbn As String, Optional seisanhiHeikin As String = Nothing) As DataTable
        ' REV_008↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   集計結果表作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_008
                If Not seisanhiHeikin Is Nothing Then
                    .AppendLine("AND    生産費平均値種類       = @生産費平均値種類 ")
                End If
                .AppendLine("ORDER BY 項目番号")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008
            If Not seisanhiHeikin Is Nothing Then
                para.Add(db.CreateParameter("@生産費平均値種類", SqlDbType.Int, seisanhiHeikin))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 集計結果表作成論理追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    'Public Shared Function InsertSyukeiKekkahyoSakuseiRonri(db As DBAccess, chosakubun As String, dc As Dictionary(Of String, String)) As Boolean
    Public Shared Function InsertSyukeiKekkahyoSakuseiRonri(db As DBAccess, chosakubun As String, versionKbn As String, dc As Dictionary(Of String, String)) As Boolean
        ' REV_008↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 集計結果表作成論理 ")
                .AppendLine("( ")
                .AppendLine("   バージョン区分 ")        ' REV_008
                .AppendLine("  ,調査区分 ")
                .AppendLine("  ,項目番号 ")
                .AppendLine("  ,生産費平均値種類 ")
                .AppendLine("  ,項目名 ")
                .AppendLine("  ,優先順位 ")
                .AppendLine("  ,計算式 ")
                .AppendLine("  ,表示単位 ")
                .AppendLine("  ,備考 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @バージョン区分 ")        ' REV_008
                .AppendLine("  ,@調査区分 ")
                .AppendLine("  ,@項目番号 ")
                .AppendLine("  ,@生産費平均値種類 ")
                .AppendLine("  ,@項目名 ")
                .AppendLine("  ,@優先順位 ")
                .AppendLine("  ,@計算式 ")
                .AppendLine("  ,@表示単位 ")
                .AppendLine("  ,@備考 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, dc("項目番号")))
            para.Add(db.CreateParameter("@生産費平均値種類", SqlDbType.Int, If(CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計,
                                                                               ComConst.生産費平均値種類.営農類型,
                                                                               (From n In ComConst.生産費平均値種類.リスト Where n.Value = dc("生産費平均値種類")).FirstOrDefault.Key)))
            para.Add(db.CreateParameter("@項目名", SqlDbType.VarChar, dc("項目名")))
            para.Add(db.CreateParameter("@優先順位", SqlDbType.Decimal, dc("優先順位")))
            para.Add(db.CreateParameter("@計算式", SqlDbType.VarChar, dc("計算式")))
            para.Add(db.CreateParameter("@表示単位", SqlDbType.Decimal, dc("表示単位")))
            para.Add(db.CreateParameter("@備考", SqlDbType.VarChar, dc("備考")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("集計結果表作成論理追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 集計結果表作成論理削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function DeleteSyukeiKekkahyoSakuseiRonri(db As DBAccess, chosakubun As String) As Boolean
    Public Shared Function DeleteSyukeiKekkahyoSakuseiRonri(db As DBAccess, chosakubun As String, versionKbn As String) As Boolean
        ' REV_008↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   集計結果表作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_008
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("集計結果表作成論理削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 集計結果検討表作成論理更新日時取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetSyukeiKekkaKentohyoSakuseiRonriUpdateDate(db As DBAccess, chosakubun As String) As String
    Public Shared Function GetSyukeiKekkaKentohyoSakuseiRonriUpdateDate(db As DBAccess, chosakubun As String, versionKbn As String) As String
        ' REV_008↑

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine("FROM   集計結果検討表作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_008
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 集計結果検討表作成論理取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    'Public Shared Function GetSyukeiKekkaKentohyoSakuseiRonri(db As DBAccess, chosakubun As String) As DataTable
    Public Shared Function GetSyukeiKekkaKentohyoSakuseiRonri(db As DBAccess, chosakubun As String, versionKbn As String, Optional chosaNen As String = Nothing) As DataTable
        ' REV_008↑

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                'REV_011 START---------------
                '出力項目名がある場合、任意項目名称を計算式に設定する
                If chosaNen <> "" Then
                    .AppendLine("SELECT A.*")
                    .AppendLine("		,B.出力項目名")
                    .AppendLine("FROM (")
                End If
                'REV_0011 END---------------
                .AppendLine("SELECT *")
                .AppendLine("FROM   集計結果検討表作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_008
                'REV_011 START---------------
                '出力項目名がある場合、任意項目名称を計算式に設定する
                If chosaNen <> "" Then
                    .AppendLine(") A")
                    .AppendLine("LEFT JOIN (SELECT * ")
                    .AppendLine("FROM   制度受取金・積立金等項目")
                    .AppendLine("WHERE  調査年 = @調査年")
                    .AppendLine("AND    調査区分 = @調査区分")
                    .AppendLine(") B")
                    .AppendLine("ON TRIM(REPLACE(A.計算式,'''','')) = B.項番")
                End If
                'REV_011 END---------------
                .AppendLine("ORDER BY 項目番号")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))   ' REV_011 

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 集計結果検討表作成論理追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    'Public Shared Function InsertSyukeiKekkaKentohyoSakuseiRonri(db As DBAccess, chosakubun As String, dc As Dictionary(Of String, String)) As Boolean
    Public Shared Function InsertSyukeiKekkaKentohyoSakuseiRonri(db As DBAccess, chosakubun As String, versionKbn As String, dc As Dictionary(Of String, String)) As Boolean
        ' REV_008↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 集計結果検討表作成論理 ")
                .AppendLine("( ")
                .AppendLine("   バージョン区分 ")        ' REV_008
                .AppendLine("  ,調査区分 ")
                .AppendLine("  ,項目番号 ")
                .AppendLine("  ,項目名 ")
                .AppendLine("  ,優先順位 ")
                .AppendLine("  ,計算式 ")
                .AppendLine("  ,表示単位 ")
                .AppendLine("  ,備考 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @バージョン区分 ")        ' REV_008
                .AppendLine("  ,@調査区分 ")
                .AppendLine("  ,@項目番号 ")
                .AppendLine("  ,@項目名 ")
                .AppendLine("  ,@優先順位 ")
                .AppendLine("  ,@計算式 ")
                .AppendLine("  ,@表示単位 ")
                .AppendLine("  ,@備考 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, dc("項目番号")))
            para.Add(db.CreateParameter("@項目名", SqlDbType.VarChar, dc("項目名")))
            para.Add(db.CreateParameter("@優先順位", SqlDbType.Decimal, dc("優先順位")))
            para.Add(db.CreateParameter("@計算式", SqlDbType.VarChar, dc("計算式")))
            para.Add(db.CreateParameter("@表示単位", SqlDbType.Decimal, dc("表示単位")))
            para.Add(db.CreateParameter("@備考", SqlDbType.VarChar, dc("備考")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("集計結果検討表作成論理追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 集計結果検討表作成論理削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function DeleteSyukeiKekkaKentohyoSakuseiRonri(db As DBAccess, chosakubun As String) As Boolean
    Public Shared Function DeleteSyukeiKekkaKentohyoSakuseiRonri(db As DBAccess, chosakubun As String, versionKbn As String) As Boolean
        ' REV_008↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   集計結果検討表作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_008
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("集計結果検討表作成論理削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_009↓
    ''' <summary>
    ''' 集計結果検討表＿報告用＿作成論理更新日時取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    Public Shared Function GetSyukeiKekkaKentohyoHoukokuyoSakuseiRonriUpdateDate(db As DBAccess, chosakubun As String, versionKbn As String) As String

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine("FROM   集計結果検討表＿報告用＿作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")
                .AppendLine("AND    農政局         = @農政局 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn
    End Function

    ''' <summary>
    ''' 集計結果検討表＿報告用＿作成論理取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSyukeiKekkaKentohyoHoukokuyoSakuseiRonri(db As DBAccess, chosakubun As String, versionKbn As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   集計結果検討表＿報告用＿作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")
                .AppendLine("AND    農政局         = @農政局 ")
                .AppendLine("ORDER BY 項目番号")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計結果検討表＿報告用＿作成論理追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    Public Shared Function InsertSyukeiKekkaKentohyoHoukokuyoSakuseiRonri(db As DBAccess, chosakubun As String, versionKbn As String, dc As Dictionary(Of String, String)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 集計結果検討表＿報告用＿作成論理 ")
                .AppendLine("( ")
                .AppendLine("   バージョン区分 ")
                .AppendLine("  ,調査区分 ")
                .AppendLine("  ,農政局 ")
                .AppendLine("  ,項目番号 ")
                .AppendLine("  ,項目名 ")
                .AppendLine("  ,優先順位 ")
                .AppendLine("  ,計算式 ")
                .AppendLine("  ,表示単位 ")
                .AppendLine("  ,備考 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @バージョン区分 ")
                .AppendLine("  ,@調査区分 ")
                .AppendLine("  ,@農政局 ")
                .AppendLine("  ,@項目番号 ")
                .AppendLine("  ,@項目名 ")
                .AppendLine("  ,@優先順位 ")
                .AppendLine("  ,@計算式 ")
                .AppendLine("  ,@表示単位 ")
                .AppendLine("  ,@備考 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, dc("項目番号")))
            para.Add(db.CreateParameter("@項目名", SqlDbType.VarChar, dc("項目名")))
            para.Add(db.CreateParameter("@優先順位", SqlDbType.Decimal, dc("優先順位")))
            para.Add(db.CreateParameter("@計算式", SqlDbType.VarChar, dc("計算式")))
            para.Add(db.CreateParameter("@表示単位", SqlDbType.Decimal, dc("表示単位")))
            para.Add(db.CreateParameter("@備考", SqlDbType.VarChar, dc("備考")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("集計結果検討表＿報告用＿作成論理追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計結果検討表＿報告用＿作成論理削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    Public Shared Function DeleteSyukeiKekkaKentohyoHoukokuyoSakuseiRonri(db As DBAccess, chosakubun As String, versionKbn As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   集計結果検討表＿報告用＿作成論理")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")
                .AppendLine("AND    農政局         = @農政局 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("集計結果検討表＿報告用＿作成論理削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function
    ' REV_009↑

    ''' <summary>
    ''' 送受信管理調査年取得（受信）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="upLow"></param>
    ''' <param name="dataType"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosaNenJushin(db As DBAccess, upLow As String, dataType As String, kyoku As String, jimusho As String, kyoten As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT DISTINCT 調査年 ")
                .AppendLine("FROM   送受信管理")
                .AppendLine("WHERE  調査区分         = @調査区分 ")
                .AppendLine("AND    上り下り区分     = @上り下り区分 ")
                .AppendLine("AND    送受信データ種別 = @送受信データ種別 ")
                If Not kyoku Is Nothing Then
                    .AppendLine("AND    農政局           = @農政局 ")
                End If
                If Not jimusho Is Nothing Then
                    .AppendLine("AND    都道府県         = @都道府県 ")
                End If
                If Not kyoten Is Nothing Then
                    .AppendLine("AND    実査設置拠点     = @実査設置拠点 ")
                End If
                .AppendLine("ORDER BY 調査年 DESC")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, upLow))
            para.Add(db.CreateParameter("@送受信データ種別", SqlDbType.Int, dataType))
            If Not kyoku Is Nothing Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kyoku))
            End If
            If Not jimusho Is Nothing Then
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))
            End If
            If Not kyoten Is Nothing Then
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kyoten))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 送受信管理取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="upLow"></param>
    ''' <param name="dataType"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSendReceiveManagement(db As DBAccess, chosaNen As String, upLow As String, dataType As String, kyoku As String, jimusho As String, kyoten As String, Optional chosaKoutei As String = Nothing) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT センサス番号 ")
                .AppendLine("     , 送信日時 ")
                .AppendLine("     , 受信日時 ")
                .AppendLine("     , 農政局 ")
                .AppendLine("     , 都道府県 ")
                .AppendLine("     , 実査設置拠点 ")
                If chosaKoutei Is Nothing Then
                    .AppendLine("FROM   送受信管理")
                Else
                    .AppendLine(String.Format("FROM   [{0}].[dbo].[送受信管理]", chosaKoutei))
                End If
                .AppendLine("WHERE  調査区分         = @調査区分 ")
                .AppendLine("AND    調査年           = @調査年 ")
                .AppendLine("AND    上り下り区分     = @上り下り区分 ")
                .AppendLine("AND    送受信データ種別 = @送受信データ種別 ")
                If Not kyoku Is Nothing Then
                    .AppendLine("AND    農政局           = @農政局 ")
                End If
                If Not jimusho Is Nothing Then
                    .AppendLine("AND    都道府県         = @都道府県 ")
                End If
                If Not kyoten Is Nothing Then
                    .AppendLine("AND    実査設置拠点     = @実査設置拠点 ")
                End If
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, upLow))
            para.Add(db.CreateParameter("@送受信データ種別", SqlDbType.Int, dataType))
            If Not kyoku Is Nothing Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kyoku))
            End If
            If Not jimusho Is Nothing Then
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))
            End If
            If Not kyoten Is Nothing Then
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kyoten))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 送受信管理データ存在チェック
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="dc"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CheckSendReceiveManagementExist(db As DBAccess, dc As Dictionary(Of String, String), chosaKoutei As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine(String.Format("FROM   [{0}].[dbo].[送受信管理]", chosaKoutei))
                .AppendLine("WHERE  調査区分         = @調査区分 ")
                .AppendLine("AND    調査年           = @調査年 ")
                .AppendLine("AND    センサス番号     = @センサス番号 ")
                .AppendLine("AND    上り下り区分     = @上り下り区分 ")
                .AppendLine("AND    送受信データ種別 = @送受信データ種別 ")
                .AppendLine("AND    農政局           = @農政局 ")
                .AppendLine("AND    都道府県         = @都道府県 ")
                .AppendLine("AND    実査設置拠点     = @実査設置拠点 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, dc("調査年")))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, dc("センサス番号")))
            para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, dc("上り下り区分")))
            para.Add(db.CreateParameter("@送受信データ種別", SqlDbType.Int, dc("送受信データ種別")))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, dc("農政局")))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, dc("都道府県")))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, dc("実査設置拠点")))

            dr = db.ExecuteReader(sb.ToString, para)

            ret = dr.HasRows
        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 送受信管理追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="dc"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertSendReceiveManagement(db As DBAccess, dc As Dictionary(Of String, String), chosaKoutei As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine(String.Format("INSERT INTO [{0}].[dbo].[送受信管理] ", chosaKoutei))
                .AppendLine("( ")
                .AppendLine("   調査区分 ")
                .AppendLine("  ,調査年 ")
                .AppendLine("  ,センサス番号 ")
                .AppendLine("  ,上り下り区分 ")
                .AppendLine("  ,送受信データ種別 ")
                .AppendLine("  ,農政局 ")
                .AppendLine("  ,都道府県 ")
                .AppendLine("  ,実査設置拠点 ")
                .AppendLine("  ,送信日時 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @調査区分 ")
                .AppendLine("  ,@調査年 ")
                .AppendLine("  ,@センサス番号 ")
                .AppendLine("  ,@上り下り区分 ")
                .AppendLine("  ,@送受信データ種別 ")
                .AppendLine("  ,@農政局 ")
                .AppendLine("  ,@都道府県 ")
                .AppendLine("  ,@実査設置拠点 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, dc("調査年")))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, dc("センサス番号")))
            para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, dc("上り下り区分")))
            para.Add(db.CreateParameter("@送受信データ種別", SqlDbType.Int, dc("送受信データ種別")))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, dc("農政局")))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, dc("都道府県")))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, dc("実査設置拠点")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("送受信管理追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 送受信管理送信日時更新
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="dc"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function UpdateSendReceiveManagementSendDate(db As DBAccess, dc As Dictionary(Of String, String), chosaKoutei As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine(String.Format("UPDATE [{0}].[dbo].[送受信管理] ", chosaKoutei))
                .AppendLine("SET    送信日時 = GETDATE() ")
                .AppendLine("      ,更新日付 = GETDATE() ")
                .AppendLine("      ,更新者ID = @更新者ID ")
                .AppendLine("WHERE  調査区分         = @調査区分 ")
                .AppendLine("AND    調査年           = @調査年 ")
                .AppendLine("AND    センサス番号     = @センサス番号 ")
                .AppendLine("AND    上り下り区分     = @上り下り区分 ")
                .AppendLine("AND    送受信データ種別 = @送受信データ種別 ")
                .AppendLine("AND    農政局           = @農政局 ")
                .AppendLine("AND    都道府県         = @都道府県 ")
                .AppendLine("AND    実査設置拠点     = @実査設置拠点 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, dc("調査年")))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, dc("センサス番号")))
            para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, dc("上り下り区分")))
            para.Add(db.CreateParameter("@送受信データ種別", SqlDbType.Int, dc("送受信データ種別")))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, dc("農政局")))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, dc("都道府県")))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, dc("実査設置拠点")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("送受信管理送信日時更新失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 送受信管理受信日時更新
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function UpdateSendReceiveManagementReceiveDate(db As DBAccess, dc As Dictionary(Of String, String)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("UPDATE 送受信管理 ")
                .AppendLine("SET    受信日時 = GETDATE() ")
                .AppendLine("      ,更新日付 = GETDATE() ")
                .AppendLine("      ,更新者ID = @更新者ID ")
                .AppendLine("WHERE  調査区分         = @調査区分 ")
                .AppendLine("AND    調査年           = @調査年 ")
                .AppendLine("AND    センサス番号     = @センサス番号 ")
                .AppendLine("AND    上り下り区分     = @上り下り区分 ")
                .AppendLine("AND    送受信データ種別 = @送受信データ種別 ")
                .AppendLine("AND    農政局           = @農政局 ")
                .AppendLine("AND    都道府県         = @都道府県 ")
                .AppendLine("AND    実査設置拠点     = @実査設置拠点 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, dc("調査年")))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, dc("センサス番号")))
            para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, dc("上り下り区分")))
            para.Add(db.CreateParameter("@送受信データ種別", SqlDbType.Int, dc("送受信データ種別")))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, dc("農政局")))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, dc("都道府県")))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, dc("実査設置拠点")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("送受信管理受信日時更新失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 毎月勤労統計取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetMaitsukiKinrouToukei(db As DBAccess, jimusho As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   毎月勤労統計")
                .AppendLine("WHERE  都道府県       = @都道府県 ")
                .AppendLine("ORDER BY 年, 月")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 毎月勤労統計取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="yyyymmL"></param>
    ''' <param name="yyyymmU"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetMaitsukiKinrouToukei(db As DBAccess, jimusho As String, yyyymmLower As String, yyyymmUpper As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   毎月勤労統計")
                .AppendLine("WHERE  都道府県       = @都道府県 ")
                .AppendLine("AND    年 * 100 + 月  BETWEEN @年月下限 AND @年月上限 ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))
            para.Add(db.CreateParameter("@年月下限", SqlDbType.Int, yyyymmLower))
            para.Add(db.CreateParameter("@年月上限", SqlDbType.Int, yyyymmUpper))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 毎月勤労統計追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertMaitsukiKinrouToukei(db As DBAccess, jimusho As String, dc As Dictionary(Of String, String)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 毎月勤労統計 ")
                .AppendLine("( ")
                .AppendLine("   都道府県 ")
                .AppendLine("  ,年 ")
                .AppendLine("  ,月 ")
                .AppendLine("  ,男＿製造業＿現金給与総額 ")
                .AppendLine("  ,男＿製造業＿労働時間 ")
                .AppendLine("  ,男＿製造業＿常用労働者数 ")
                .AppendLine("  ,男＿建設業＿現金給与総額 ")
                .AppendLine("  ,男＿建設業＿労働時間 ")
                .AppendLine("  ,男＿建設業＿常用労働者数 ")
                .AppendLine("  ,男＿運輸業郵便業＿現金給与総額 ")
                .AppendLine("  ,男＿運輸業郵便業＿労働時間 ")
                .AppendLine("  ,男＿運輸業郵便業＿常用労働者数 ")
                .AppendLine("  ,男女平均＿全産業計＿賃金対比 ")
                .AppendLine("  ,男女平均＿全産業計＿労働時間対比 ")
                .AppendLine("  ,女＿製造業＿現金給与総額 ")
                .AppendLine("  ,女＿製造業＿労働時間 ")
                .AppendLine("  ,女＿製造業＿常用労働者数 ")
                .AppendLine("  ,女＿建設業＿現金給与総額 ")
                .AppendLine("  ,女＿建設業＿労働時間 ")
                .AppendLine("  ,女＿建設業＿常用労働者数 ")
                .AppendLine("  ,女＿運輸業郵便業＿現金給与総額 ")
                .AppendLine("  ,女＿運輸業郵便業＿労働時間 ")
                .AppendLine("  ,女＿運輸業郵便業＿常用労働者数 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @都道府県 ")
                .AppendLine("  ,@年 ")
                .AppendLine("  ,@月 ")
                .AppendLine("  ,@男＿製造業＿現金給与総額 ")
                .AppendLine("  ,@男＿製造業＿労働時間 ")
                .AppendLine("  ,@男＿製造業＿常用労働者数 ")
                .AppendLine("  ,@男＿建設業＿現金給与総額 ")
                .AppendLine("  ,@男＿建設業＿労働時間 ")
                .AppendLine("  ,@男＿建設業＿常用労働者数 ")
                .AppendLine("  ,@男＿運輸業郵便業＿現金給与総額 ")
                .AppendLine("  ,@男＿運輸業郵便業＿労働時間 ")
                .AppendLine("  ,@男＿運輸業郵便業＿常用労働者数 ")
                .AppendLine("  ,@男女平均＿全産業計＿賃金対比 ")
                .AppendLine("  ,@男女平均＿全産業計＿労働時間対比 ")
                .AppendLine("  ,@女＿製造業＿現金給与総額 ")
                .AppendLine("  ,@女＿製造業＿労働時間 ")
                .AppendLine("  ,@女＿製造業＿常用労働者数 ")
                .AppendLine("  ,@女＿建設業＿現金給与総額 ")
                .AppendLine("  ,@女＿建設業＿労働時間 ")
                .AppendLine("  ,@女＿建設業＿常用労働者数 ")
                .AppendLine("  ,@女＿運輸業郵便業＿現金給与総額 ")
                .AppendLine("  ,@女＿運輸業郵便業＿労働時間 ")
                .AppendLine("  ,@女＿運輸業郵便業＿常用労働者数 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))
            para.Add(db.CreateParameter("@年", SqlDbType.Int, dc("年")))
            para.Add(db.CreateParameter("@月", SqlDbType.Int, dc("月")))
            para.Add(db.CreateParameter("@男＿製造業＿現金給与総額", SqlDbType.Decimal, dc("男＿製造業＿現金給与総額")))
            para.Add(db.CreateParameter("@男＿製造業＿労働時間", SqlDbType.Decimal, dc("男＿製造業＿労働時間")))
            para.Add(db.CreateParameter("@男＿製造業＿常用労働者数", SqlDbType.Decimal, dc("男＿製造業＿常用労働者数")))
            para.Add(db.CreateParameter("@男＿建設業＿現金給与総額", SqlDbType.Decimal, dc("男＿建設業＿現金給与総額")))
            para.Add(db.CreateParameter("@男＿建設業＿労働時間", SqlDbType.Decimal, dc("男＿建設業＿労働時間")))
            para.Add(db.CreateParameter("@男＿建設業＿常用労働者数", SqlDbType.Decimal, dc("男＿建設業＿常用労働者数")))
            para.Add(db.CreateParameter("@男＿運輸業郵便業＿現金給与総額", SqlDbType.Decimal, dc("男＿運輸業郵便業＿現金給与総額")))
            para.Add(db.CreateParameter("@男＿運輸業郵便業＿労働時間", SqlDbType.Decimal, dc("男＿運輸業郵便業＿労働時間")))
            para.Add(db.CreateParameter("@男＿運輸業郵便業＿常用労働者数", SqlDbType.Decimal, dc("男＿運輸業郵便業＿常用労働者数")))
            para.Add(db.CreateParameter("@男女平均＿全産業計＿賃金対比", SqlDbType.Decimal, dc("男女平均＿全産業計＿賃金対比")))
            para.Add(db.CreateParameter("@男女平均＿全産業計＿労働時間対比", SqlDbType.Decimal, dc("男女平均＿全産業計＿労働時間対比")))
            para.Add(db.CreateParameter("@女＿製造業＿現金給与総額", SqlDbType.Decimal, dc("女＿製造業＿現金給与総額")))
            para.Add(db.CreateParameter("@女＿製造業＿労働時間", SqlDbType.Decimal, dc("女＿製造業＿労働時間")))
            para.Add(db.CreateParameter("@女＿製造業＿常用労働者数", SqlDbType.Decimal, dc("女＿製造業＿常用労働者数")))
            para.Add(db.CreateParameter("@女＿建設業＿現金給与総額", SqlDbType.Decimal, dc("女＿建設業＿現金給与総額")))
            para.Add(db.CreateParameter("@女＿建設業＿労働時間", SqlDbType.Decimal, dc("女＿建設業＿労働時間")))
            para.Add(db.CreateParameter("@女＿建設業＿常用労働者数", SqlDbType.Decimal, dc("女＿建設業＿常用労働者数")))
            para.Add(db.CreateParameter("@女＿運輸業郵便業＿現金給与総額", SqlDbType.Decimal, dc("女＿運輸業郵便業＿現金給与総額")))
            para.Add(db.CreateParameter("@女＿運輸業郵便業＿労働時間", SqlDbType.Decimal, dc("女＿運輸業郵便業＿労働時間")))
            para.Add(db.CreateParameter("@女＿運輸業郵便業＿常用労働者数", SqlDbType.Decimal, dc("女＿運輸業郵便業＿常用労働者数")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("毎月勤労統計追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 毎月勤労統計削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteMaitsukiKinrouToukei(db As DBAccess, jimusho As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   毎月勤労統計")
                .AppendLine("WHERE  都道府県       = @都道府県 ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("毎月勤労統計削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 労賃単価取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="seisanhi"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="jimusho"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetRouchinTanka(db As DBAccess, seisanhi As String, chosaNen As String, jimusho As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   労賃単価")
                .AppendLine("WHERE  生産費区分     = @生産費区分 ")
                .AppendLine("AND    調査年         = @調査年 ")
                .AppendLine("AND    都道府県       = @都道府県 ")
            End With

            para.Add(db.CreateParameter("@生産費区分", SqlDbType.Int, seisanhi))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 労賃単価追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertRouchinTanka(db As DBAccess, seisanhi As String, chosaNen As String, jimusho As String, dc As Dictionary(Of String, String)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 労賃単価 ")
                .AppendLine("( ")
                .AppendLine("   生産費区分 ")
                .AppendLine("  ,調査年 ")
                .AppendLine("  ,都道府県 ")
                .AppendLine("  ,[男＿５～２９人＿本年値] ")
                .AppendLine("  ,[男＿５～２９人＿前年値] ")
                .AppendLine("  ,[男＿５～２９人＿対前年比] ")
                .AppendLine("  ,[女＿５～２９人＿本年値] ")
                .AppendLine("  ,[女＿５～２９人＿前年値] ")
                .AppendLine("  ,[女＿５～２９人＿対前年比] ")
                .AppendLine("  ,[男女平均＿５～２９人＿本年値] ")
                .AppendLine("  ,[男女平均＿５～２９人＿前年値] ")
                .AppendLine("  ,[男女平均＿５～２９人＿対前年比] ")
                .AppendLine("  ,男女平均＿採用単価＿本年値 ")
                .AppendLine("  ,男女平均＿採用単価＿前年値 ")
                .AppendLine("  ,男女平均＿採用単価＿対前年比 ")
                .AppendLine("  ,男女平均＿通勤手当割合 ")
                .AppendLine("  ,男女平均＿評価単価 ")
                .AppendLine("  ,男女平均＿３０人以上＿産業計対比 ")
                .AppendLine("  ,男女平均＿３０人以上＿対比単価 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @生産費区分 ")
                .AppendLine("  ,@調査年 ")
                .AppendLine("  ,@都道府県 ")
                .AppendLine("  ,@男＿５＿２９人＿本年値 ")
                .AppendLine("  ,@男＿５＿２９人＿前年値 ")
                .AppendLine("  ,@男＿５＿２９人＿対前年比 ")
                .AppendLine("  ,@女＿５＿２９人＿本年値 ")
                .AppendLine("  ,@女＿５＿２９人＿前年値 ")
                .AppendLine("  ,@女＿５＿２９人＿対前年比 ")
                .AppendLine("  ,@男女平均＿５＿２９人＿本年値 ")
                .AppendLine("  ,@男女平均＿５＿２９人＿前年値 ")
                .AppendLine("  ,@男女平均＿５＿２９人＿対前年比 ")
                .AppendLine("  ,@男女平均＿採用単価＿本年値 ")
                .AppendLine("  ,@男女平均＿採用単価＿前年値 ")
                .AppendLine("  ,@男女平均＿採用単価＿対前年比 ")
                .AppendLine("  ,@男女平均＿通勤手当割合 ")
                .AppendLine("  ,@男女平均＿評価単価 ")
                .AppendLine("  ,@男女平均＿３０人以上＿産業計対比 ")
                .AppendLine("  ,@男女平均＿３０人以上＿対比単価 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@生産費区分", SqlDbType.Int, seisanhi))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))
            para.Add(db.CreateParameter("@男＿５＿２９人＿本年値", SqlDbType.Decimal, dc("男＿５～２９人＿本年値")))
            para.Add(db.CreateParameter("@男＿５＿２９人＿前年値", SqlDbType.Decimal, dc("男＿５～２９人＿前年値")))
            para.Add(db.CreateParameter("@男＿５＿２９人＿対前年比", SqlDbType.Decimal, dc("男＿５～２９人＿対前年比")))
            para.Add(db.CreateParameter("@女＿５＿２９人＿本年値", SqlDbType.Decimal, dc("女＿５～２９人＿本年値")))
            para.Add(db.CreateParameter("@女＿５＿２９人＿前年値", SqlDbType.Decimal, dc("女＿５～２９人＿前年値")))
            para.Add(db.CreateParameter("@女＿５＿２９人＿対前年比", SqlDbType.Decimal, dc("女＿５～２９人＿対前年比")))
            para.Add(db.CreateParameter("@男女平均＿５＿２９人＿本年値", SqlDbType.Decimal, dc("男女平均＿５～２９人＿本年値")))
            para.Add(db.CreateParameter("@男女平均＿５＿２９人＿前年値", SqlDbType.Decimal, dc("男女平均＿５～２９人＿前年値")))
            para.Add(db.CreateParameter("@男女平均＿５＿２９人＿対前年比", SqlDbType.Decimal, dc("男女平均＿５～２９人＿対前年比")))
            para.Add(db.CreateParameter("@男女平均＿採用単価＿本年値", SqlDbType.Decimal, dc("男女平均＿採用単価＿本年値")))
            para.Add(db.CreateParameter("@男女平均＿採用単価＿前年値", SqlDbType.Decimal, dc("男女平均＿採用単価＿前年値")))
            para.Add(db.CreateParameter("@男女平均＿採用単価＿対前年比", SqlDbType.Decimal, dc("男女平均＿採用単価＿対前年比")))
            para.Add(db.CreateParameter("@男女平均＿通勤手当割合", SqlDbType.Decimal, dc("男女平均＿通勤手当割合")))
            para.Add(db.CreateParameter("@男女平均＿評価単価", SqlDbType.Decimal, dc("男女平均＿評価単価")))
            para.Add(db.CreateParameter("@男女平均＿３０人以上＿産業計対比", SqlDbType.Decimal, dc("男女平均＿３０人以上＿産業計対比")))
            para.Add(db.CreateParameter("@男女平均＿３０人以上＿対比単価", SqlDbType.Decimal, dc("男女平均＿３０人以上＿対比単価")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("労賃単価追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 労賃単価削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteRouchinTanka(db As DBAccess, seisanhi As String, chosaNen As String, jimusho As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   労賃単価")
                .AppendLine("WHERE  生産費区分     = @生産費区分 ")
                .AppendLine("AND    調査年         = @調査年 ")
                .AppendLine("AND    都道府県       = @都道府県 ")
            End With

            para.Add(db.CreateParameter("@生産費区分", SqlDbType.Int, seisanhi))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("労賃単価削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 専門調査員一覧取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSenmonChosainList(db As DBAccess) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT A.都道府県 ")
                .AppendLine("     , A.実査設置拠点 ")
                .AppendLine("     , A.ユーザーID ")
                .AppendLine("     , A.氏名 ")
                .AppendLine("     , A.更新日付 ")
                .AppendLine("     , COALESCE(担当客体数,0) AS 担当客体数 ")
                .AppendLine("FROM 専門調査員管理 A ")
                .AppendLine("LEFT JOIN (SELECT 都道府県 ")
                .AppendLine("                , 実査設置拠点 ")
                .AppendLine("                , ユーザーID ")
                .AppendLine("                , COUNT(*) AS 担当客体数 ")
                .AppendLine("           FROM 専門調査員担当調査客体 ")
                .AppendLine("           GROUP BY 都道府県 ")
                .AppendLine("                  , 実査設置拠点 ")
                .AppendLine("                  , ユーザーID) B ")
                .AppendLine("ON    A.都道府県     = B.都道府県 ")
                .AppendLine("AND   A.実査設置拠点 = B.実査設置拠点 ")
                .AppendLine("AND   A.ユーザーID   = B.ユーザーID ")
                .AppendLine("WHERE A.都道府県     = @都道府県 ")
                .AppendLine("AND   A.実査設置拠点 = @実査設置拠点 ")
                .AppendLine("ORDER BY A.ユーザーID ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 専門調査員管理取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="userID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSenmonChosain(db As DBAccess, userID As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   専門調査員管理")
                .AppendLine("WHERE  都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                .AppendLine("AND    ユーザーID     = @ユーザーID ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, userID))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 専門調査員管理存在チェック
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="userID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CheckSenmonChosainExist(db As DBAccess, Optional userID As String = Nothing) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   専門調査員管理")
                .AppendLine("WHERE  都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                If Not userID Is Nothing Then
                    .AppendLine("AND    ユーザーID     = @ユーザーID ")
                End If
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            If Not userID Is Nothing Then
                para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, userID))
            End If

            dr = db.ExecuteReader(sb.ToString, para)

            ret = dr.HasRows
        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 専門調査員管理追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="userID"></param>
    ''' <param name="shimei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertSenmonChosain(db As DBAccess, userID As String, shimei As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 専門調査員管理 ")
                .AppendLine("( ")
                .AppendLine("   都道府県 ")
                .AppendLine("  ,実査設置拠点 ")
                .AppendLine("  ,ユーザーID ")
                .AppendLine("  ,氏名 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @都道府県 ")
                .AppendLine("  ,@実査設置拠点 ")
                .AppendLine("  ,@ユーザーID ")
                .AppendLine("  ,@氏名 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, userID))
            para.Add(db.CreateParameter("@氏名", SqlDbType.VarChar, shimei))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("専門調査員管理追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 専門調査員管理更新
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="userID"></param>
    ''' <param name="shimei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function UpdateSenmonChosain(db As DBAccess, userID As String, shimei As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("UPDATE 専門調査員管理 ")
                .AppendLine("SET    氏名           = @氏名 ")
                .AppendLine("      ,更新日付       = GETDATE() ")
                .AppendLine("      ,更新者ID       = @更新者ID")
                .AppendLine("WHERE  都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                .AppendLine("AND    ユーザーID     = @ユーザーID ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, userID))
            para.Add(db.CreateParameter("@氏名", SqlDbType.VarChar, shimei))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("専門調査員管理更新失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 専門調査員管理削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="userID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteSenmonChosain(db As DBAccess, userID As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   専門調査員管理")
                .AppendLine("WHERE  都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                .AppendLine("AND    ユーザーID     = @ユーザーID ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, userID))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("専門調査員管理削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 専門調査員氏名取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="userID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSenmonChosainShimei(db As DBAccess, userID As String) As String

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT 氏名")
                .AppendLine("FROM   専門調査員管理")
                .AppendLine("WHERE  都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                .AppendLine("AND    ユーザーID     = @ユーザーID ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, userID))

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("氏名").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn
    End Function

    ''' <summary>
    ''' 専門調査員担当調査客体取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="userID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSenmonChosainKyakutai(db As DBAccess, userID As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   専門調査員担当調査客体")
                .AppendLine("WHERE  都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                .AppendLine("AND    ユーザーID     = @ユーザーID ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, userID))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 専門調査員担当調査客体ユーザーID取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSenmonChosainKyakutaiUserID(db As DBAccess, censusNo As String) As List(Of String)

        Dim ret As New List(Of String)
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT ユーザーID")
                .AppendLine("FROM   専門調査員担当調査客体")
                .AppendLine("WHERE  都道府県       = @都道府県 ")
                .AppendLine("AND    センサス番号   = @センサス番号 ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, censusNo))

            dr = db.ExecuteReader(sb.ToString, para)
            While dr.Read()
                ret.Add(dr("ユーザーID").ToString)
            End While

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 専門調査員担当調査客体存在チェック
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="userID"></param>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CheckSenmonChosainKyakutaiExist(db As DBAccess, userID As String, censusNo As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   専門調査員担当調査客体")
                .AppendLine("WHERE  都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                .AppendLine("AND    ユーザーID     = @ユーザーID ")
                .AppendLine("AND    センサス番号   = @センサス番号 ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, userID))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, censusNo))

            dr = db.ExecuteReader(sb.ToString, para)

            ret = dr.HasRows
        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 専門調査員担当調査客体追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="userID"></param>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertSenmonChosainKyakutai(db As DBAccess, userID As String, censusNo As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 専門調査員担当調査客体 ")
                .AppendLine("( ")
                .AppendLine("   都道府県 ")
                .AppendLine("  ,実査設置拠点 ")
                .AppendLine("  ,ユーザーID ")
                .AppendLine("  ,センサス番号 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @都道府県 ")
                .AppendLine("  ,@実査設置拠点 ")
                .AppendLine("  ,@ユーザーID ")
                .AppendLine("  ,@センサス番号 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, userID))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, censusNo))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("専門調査員担当調査客体追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 専門調査員担当調査客体削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="userID"></param>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteSenmonChosainKyakutai(db As DBAccess, userID As String, Optional censusNo As String = Nothing) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   専門調査員担当調査客体")
                .AppendLine("WHERE  都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                .AppendLine("AND    ユーザーID     = @ユーザーID ")
                If Not censusNo Is Nothing Then
                    .AppendLine("AND    センサス番号   = @センサス番号 ")
                End If
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, userID))
            If Not censusNo Is Nothing Then
                para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, censusNo))
            End If

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("専門調査員担当調査客体削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 専門調査員及び担当調査客体一覧取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSenmonChosainKyakutaiList(db As DBAccess) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT A.ユーザーID ")
                .AppendLine("     , A.氏名 ")
                .AppendLine("     , B.更新日付 ")
                .AppendLine("     , B.センサス番号 ")
                .AppendLine("FROM   専門調査員管理 A ")
                .AppendLine("LEFT JOIN 専門調査員担当調査客体 B ")
                .AppendLine("ON     A.都道府県       = B.都道府県 ")
                .AppendLine("AND    A.実査設置拠点   = B.実査設置拠点 ")
                .AppendLine("AND    A.ユーザーID     = B.ユーザーID ")
                .AppendLine("WHERE  A.都道府県       = @都道府県 ")
                .AppendLine("AND    A.実査設置拠点   = @実査設置拠点 ")
                .AppendLine("ORDER BY ユーザーID ")
                .AppendLine("       , センサス番号 ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' CSV一括読み込み
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="tableName"></param>
    ''' <param name="lstLine"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function BulkCopyCSV(db As DBAccess, tableName As String, lstLine As List(Of String())) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing

        Try
            sb = New System.Text.StringBuilder
            sb.AppendLine(String.Format("SELECT TOP 0 * FROM {0} ", tableName))

            Dim dt As DataTable = db.GetDataTable(sb.ToString)

            Dim row As DataRow = Nothing

            For Each line As String() In lstLine
                row = dt.NewRow()
                row.BeginEdit()

                For i As Integer = LBound(line) To UBound(line)
                    If Not String.IsNullOrEmpty(line(i)) Then
                        row(i) = line(i)
                    End If
                Next

                row.EndEdit()
                dt.Rows.Add(row)

                If dt.Rows.Count = 10000 Then
                    Using bc As New Data.SqlClient.SqlBulkCopy(db.Connection, Data.SqlClient.SqlBulkCopyOptions.KeepIdentity, db.Transaction)
                        bc.DestinationTableName = tableName
                        bc.WriteToServer(dt)
                        bc.Close()
                    End Using

                    dt.Clear()
                End If
            Next

            If dt.Rows.Count > 0 Then
                Using bc As New Data.SqlClient.SqlBulkCopy(db.Connection, Data.SqlClient.SqlBulkCopyOptions.KeepIdentity, db.Transaction)
                    bc.DestinationTableName = tableName
                    bc.WriteToServer(dt)
                    bc.Close()
                End Using
            End If

            ret = True
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 農業地域区分取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="seisanhi"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="year"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
        Public Shared Function GetNogyoChikiKbn(db As DBAccess, jimusho As String, shikuchoson As String, kyushikuchoson As String, year As String) As DataTable 'REV_016 CNG DAIKO

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   農業地域類型区分マスタ")
                .AppendLine("WHERE  調査年   = @調査年 ") 'REV_016 CNG DAIKO
                .AppendLine("AND    都道府県   = @都道府県 ") 'REV_016 CNG DAIKO
                .AppendLine("AND    市区町村   = @市区町村 ")
                .AppendLine("AND    旧市区町村 = @旧市区町村 ")
            End With

            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))
            para.Add(db.CreateParameter("@市区町村", SqlDbType.Int, shikuchoson))
            para.Add(db.CreateParameter("@旧市区町村", SqlDbType.Int, kyushikuchoson))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, year)) 'REV_016 CNG DAIKO

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    'REV_016 ADD START
    ''' <summary>
    ''' 農業地域類型マスタ調査年 更新日時取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="year"></param>
    ''' <returns></returns>
    Public Shared Function GetNogyoChikiMstYearUpdateDate(db As DBAccess, year As String) As String

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            'SQLの設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine("FROM 農業地域類型区分マスタ")
                .AppendLine("WHERE 調査年 = @調査年")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, year))

            dr = db.ExecuteReader(sb.ToString, para)

            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn

    End Function

    ''' <summary>
    ''' 農業地域類型マスタ調査年一覧取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="year"></param>
    ''' <returns></returns>
    Public Shared Function GetNogyoChikiMstYearList(db As DBAccess, year As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            'SQLの設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM 農業地域類型区分マスタ")
                .AppendLine("WHERE 調査年 = @調査年")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, year))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret

    End Function

    ''' <summary>
    ''' 農業地域類型マスタ調査年データ削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="year"></param>
    ''' <returns></returns>
    Public Shared Function DeleteNogyoChikiMstYear(db As DBAccess, year As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM 農業地域類型区分マスタ")
                .AppendLine("WHERE 調査年 = @調査年 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, year))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("農業地域類型マスタ調査年削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 農業地域類型マスタ調査年データ登録
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="tableName"></param>
    ''' <param name="year"></param>
    ''' <param name="lstLine"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertNogyoChikiMstYearList(db As DBAccess, tableName As String, year As String, lstDc As List(Of Dictionary(Of String, String))) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing

        Try
            sb = New System.Text.StringBuilder
            sb.AppendLine(String.Format("SELECT TOP 0 * FROM " & tableName))
            Dim dt As DataTable = db.GetDataTable(sb.ToString)

            Dim row As DataRow = Nothing
            For Each dc As Dictionary(Of String, String) In lstDc
                row = dt.NewRow()
                row.BeginEdit()

                row("調査年") = year
                row("都道府県") = dc("都道府県")
                row("市区町村") = dc("市区町村")
                row("旧市区町村") = dc("旧市区町村")
                row("市区町村名") = dc("市区町村名")
                row("旧市区町村名") = dc("旧市区町村名")
                row("第１次分類") = dc("第１次分類")
                row("第２次分類") = dc("第２次分類")
                row("更新日付") = Now()
                row("更新者ID") = CommonInfo.UserId

                row.EndEdit()
                dt.Rows.Add(row)

                If dt.Rows.Count = 10000 Then
                    Using bc As New Data.SqlClient.SqlBulkCopy(db.Connection, Data.SqlClient.SqlBulkCopyOptions.KeepIdentity, db.Transaction)
                        bc.DestinationTableName = tableName
                        bc.WriteToServer(dt)
                        bc.Close()
                    End Using

                    dt.Clear()
                End If
            Next

            If dt.Rows.Count > 0 Then
                Using bc As New Data.SqlClient.SqlBulkCopy(db.Connection, Data.SqlClient.SqlBulkCopyOptions.KeepIdentity, db.Transaction)
                    bc.DestinationTableName = tableName
                    bc.WriteToServer(dt)
                    bc.Close()
                End Using
            End If

            ret = True

        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function
    'REV_016 ADD END

    ''' <summary>
    ''' 中間集計表、調査票項目マスタ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="tourokusakujoKubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTyukanSyukeiItemMaster(db As DBAccess, chosaKubun As String, tourokusakujoKubun As Integer, chosaNen As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim ver_kubun As String
        Dim vk_check As Integer = 2022

        Try

            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where Not val.Contains("＿可変")

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ver_kubun = ComUtil.getVersionKubun(chosaNen, chosaKubun)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT     A.項目番号,A.可変区分,A.型区分,C.シート名,C.行位置,C.列位置,C.ラウンド桁数,C.乗算値,C.追加項目,C.登録削除区分,C.明細番号")
                .AppendLine("         , B.precision AS 有効桁数 ")
                .AppendLine("         , B.scale AS 小数点以下桁数")
                .AppendLine("FROM      (SELECT * ")
                .AppendLine("           FROM   調査票項目マスタ ")
                .AppendLine("           WHERE  調査区分 = @調査区分 ")
                .AppendLine("           AND  バージョン区分 = @バージョン区分 ")
                .AppendLine("          ) A")
                .AppendLine("LEFT JOIN (SELECT * ")
                .AppendLine("           FROM   sys.columns")
                .AppendLine("           WHERE  object_id IN (SELECT object_id")
                .AppendLine("                                FROM   sys.tables")
                .AppendLine("                                WHERE  name IN ('" & String.Join("', '", query) & "')")
                .AppendLine("                               )")
                .AppendLine("          ) B")
                .AppendLine("ON A.項目番号 = B.name")
                .AppendLine("LEFT JOIN (SELECT * ")
                .AppendLine("           FROM   中間集計表項目マスタ ")
                .AppendLine("           WHERE  調査区分 = @調査区分 ")
                .AppendLine("           AND  登録削除区分 =  " & tourokusakujoKubun)
                .AppendLine("           AND  バージョン区分 = @バージョン区分 ")
                .AppendLine("          ) C")
                .AppendLine("ON A.項目番号 = C.項目番号")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, ver_kubun))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_013↓
    ''' <summary>
    ''' 営農集計条件マスタを取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="param"></param>
    ''' <param name="isBumon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getEinouSyukeiJoukenMaster(db As DBAccess, chosakubun As String, param As String, versionKbn As String, Optional isBumon As Boolean = False) As DataTable
        ' REV_013↑
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   営農集計条件マスタ")
                .AppendLine("WHERE  調査区分   = @調査区分 ")
                .AppendLine("AND  バージョン区分 = @バージョン区分 ")        ' REV_013
                If isBumon Then
                    .AppendLine("AND    部門コード = @部門コード ")
                Else
                    .AppendLine("AND    集計コード = @集計コード ")
                End If
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_013
            If isBumon Then
                para.Add(db.CreateParameter("@部門コード", SqlDbType.VarChar, param))
            Else
                para.Add(db.CreateParameter("@集計コード", SqlDbType.VarChar, param))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_013↓
    ''' <summary>
    ''' 営農集計条件マスタを取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="syukei1"></param>
    ''' <param name="syukei2"></param>
    ''' <param name="syukei3"></param>
    ''' <param name="syukei4"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getEinouSyukeiJoukenMaster(db As DBAccess, chosakubun As String, syukei1 As String, syukei2 As String, syukei3 As String, syukei4 As String, versionKbn As String) As DataTable
        ' REV_013↑
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   営農集計条件マスタ")
                .AppendLine("WHERE  調査区分   = @調査区分 ")
                .AppendLine("AND  バージョン区分 = @バージョン区分 ")        ' REV_013
                .AppendLine("AND    集計１     = @集計１ ")
                .AppendLine("AND    集計２     = @集計２ ")
                .AppendLine("AND    集計３     = @集計３ ")
                .AppendLine("AND    集計４     = @集計４ ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_013
            para.Add(db.CreateParameter("@集計１", SqlDbType.Int, syukei1))
            para.Add(db.CreateParameter("@集計２", SqlDbType.Int, syukei2))
            para.Add(db.CreateParameter("@集計３", SqlDbType.Int, syukei3))
            para.Add(db.CreateParameter("@集計４", SqlDbType.Int, syukei4))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 規模階層マスタを取得する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="kaisouCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getKiboKaisouMaster(db As DBAccess, chosaKubun As String, Optional kaisouCd As String = Nothing) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   規模階層マスタ")
                .AppendLine("WHERE  調査区分   = @調査区分 ")
                If Not kaisouCd Is Nothing Then
                    .AppendLine("AND    階層コード = @階層コード ")
                End If
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            If Not kaisouCd Is Nothing Then
                para.Add(db.CreateParameter("@階層コード", SqlDbType.VarChar, kaisouCd))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 任意階層項目番号取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetNiniKaisouItemNo(db As DBAccess, chosakubun As String) As DataTable
    Public Shared Function GetNiniKaisouItemNo(db As DBAccess, chosakubun As String, versionKbn As String) As DataTable
        ' REV_008↑
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT DISTINCT 項目番号")
                .AppendLine("FROM   任意階層")
                .AppendLine("WHERE  調査区分   = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_008
                .AppendLine("AND    農政局     = @農政局 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 任意階層取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetNiniKaisou(db As DBAccess, chosakubun As String) As DataTable
    Public Shared Function GetNiniKaisou(db As DBAccess, chosakubun As String, versionKbn As String) As DataTable
        ' REV_008↑
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   任意階層")
                .AppendLine("WHERE  調査区分   = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_008
                .AppendLine("AND    農政局     = @農政局 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 任意階層追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    'Public Shared Function InsertNiniKaisou(db As DBAccess, chosakubun As String, dc As Dictionary(Of String, String)) As Boolean
    Public Shared Function InsertNiniKaisou(db As DBAccess, chosakubun As String, versionKbn As String, dc As Dictionary(Of String, String)) As Boolean
        ' REV_008↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 任意階層 ")
                .AppendLine("( ")
                .AppendLine("   バージョン区分 ")        ' REV_008
                .AppendLine("  ,調査区分 ")
                .AppendLine("  ,農政局 ")
                .AppendLine("  ,項目番号 ")
                .AppendLine("  ,規模階層 ")
                .AppendLine("  ,上限 ")
                .AppendLine("  ,下限 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @バージョン区分 ")        ' REV_008
                .AppendLine("  ,@調査区分 ")
                .AppendLine("  ,@農政局 ")
                .AppendLine("  ,@項目番号 ")
                .AppendLine("  ,@規模階層 ")
                .AppendLine("  ,@上限 ")
                .AppendLine("  ,@下限 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, dc("項目番号")))
            para.Add(db.CreateParameter("@規模階層", SqlDbType.Int, dc("規模階層")))
            para.Add(db.CreateParameter("@上限", SqlDbType.Decimal, dc("上限")))
            para.Add(db.CreateParameter("@下限", SqlDbType.Decimal, dc("下限")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("任意階層追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ' REV_008↓
    ''' <summary>
    ''' 任意階層削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    'Public Shared Function DeleteNiniKaisou(db As DBAccess, chosakubun As String) As Boolean
    Public Shared Function DeleteNiniKaisou(db As DBAccess, chosakubun As String, versionKbn As String) As Boolean
        ' REV_008↑
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   任意階層")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_008
                .AppendLine("AND    農政局         = @農政局 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("任意階層削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計用センサス番号管理追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertSyukeiCensusNo(db As DBAccess, chosaNen As String, censusNo As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 集計用センサス番号管理 ")
                .AppendLine("( ")
                .AppendLine("   調査年 ")
                .AppendLine("  ,センサス番号 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @調査年 ")
                .AppendLine("  ,@センサス番号 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, censusNo))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("集計用センサス番号管理追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計用センサス番号管理削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteSyukeiCensusNo(db As DBAccess) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   集計用センサス番号管理")
                .AppendLine("WHERE  更新者ID       = @更新者ID ")
            End With

            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("集計用センサス番号管理削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計用センサス番号管理取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSyukeiCensusNo(db As DBAccess, chosaNen As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   集計用センサス番号管理")
                .AppendLine("WHERE  調査年         = @調査年 ")
                .AppendLine("AND    更新者ID       = @更新者ID ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret

    End Function

    ' REV_008↓
    ''' <summary>
    ''' 生産費集計結果表出力用編集マスタ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function GetSeisanhiSyutsuryokuHensyu(db As DBAccess, chosakubun As String) As DataTable
    Public Shared Function GetSeisanhiSyutsuryokuHensyu(db As DBAccess, chosakubun As String, versionKbn As String) As DataTable
        ' REV_008↑
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   生産費集計結果表出力用編集マスタ")
                .AppendLine("WHERE  調査区分   = @調査区分 ")
                .AppendLine("AND    バージョン区分 = @バージョン区分 ")        ' REV_008
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_008

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret

    End Function

    ''' <summary>
    ''' 更新日時設定
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="tableName"></param>
    ''' <remarks></remarks>
    Public Shared Function GetUpdateDate(ByVal db As DBAccess, ByVal chosaNen As String, ByVal tableName As String) As String

        If db Is Nothing Then
            Return Nothing
        End If

        If String.IsNullOrWhiteSpace(chosaNen) Then
            Return Nothing
        End If

        If String.IsNullOrWhiteSpace(tableName) Then
            Return Nothing
        End If

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine(String.Format("FROM   ""{0}""", tableName))
                .AppendLine("WHERE  調査年       = @調査年 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn

    End Function

    ''' <summary>
    ''' 調査年取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosaNen(ByVal db As DBAccess, ByVal tableName As String) As DataTable

        If db Is Nothing Then
            Return Nothing
        End If

        If String.IsNullOrWhiteSpace(tableName) Then
            Return Nothing
        End If

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT DISTINCT 調査年 ")
                .AppendLine(String.Format("FROM   ""{0}""", tableName))
                .AppendLine("ORDER BY 調査年 DESC")
            End With

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret

    End Function

    ''' <summary>
    ''' 「営農類型別経営統計_個人」テーブルから農業地域、営農類型、営農規模でグループ化したレコード群の平均値を算出し抽出する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SelectKobetsuKekkaHyoRowListOfAverage(ByVal db As DBAccess, ByVal chosaNen As String) As List(Of ComUtil.営農欠測値適用平均値データ)

        If db Is Nothing Then
            Return Nothing
        End If

        If Not ComUtil.TryParseToInteger(chosaNen) Then
            Return Nothing
        End If

        Dim rowList As New List(Of ComUtil.営農欠測値適用平均値データ)
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each 農業地域コード In ComConst.欠測値.農業地域コード一覧

                For Each 営農類型コード_営農規模コード In ComConst.欠測値.営農類型コード_営農規模コード一覧

                    For Each 営農規模コード In 営農類型コード_営農規模コード.Value

                        sb = New System.Text.StringBuilder

                        'REV_008 START---------------
                        ' SQL文の設定
                        With sb
                            .AppendLine("SELECT COUNT(*) AS レコード数, CASE WHEN AVG(K010712) IS NULL THEN NULL ELSE ROUND(AVG(COALESCE(K010712,0)),0) END AS 現金, CASE WHEN AVG(K010713) IS NULL THEN NULL ELSE ROUND(AVG(COALESCE(K010713,0)),0) END AS 預金等, CASE WHEN AVG(K010720) IS NULL THEN NULL ELSE ROUND(AVG(COALESCE(K010720,0)),0) END AS 売掛金・未収金, CASE WHEN AVG(K010732) IS NULL THEN NULL ELSE ROUND(AVG(COALESCE(K010732,0)),0) END AS 自動車・農機具, CASE WHEN AVG(K010733) IS NULL THEN NULL ELSE ROUND(AVG(COALESCE(K010733,0)),0) END AS 建物・構築物, CASE WHEN AVG(K010734) IS NULL THEN NULL ELSE ROUND(AVG(COALESCE(K010734,0)),0) END AS 土地, CASE WHEN AVG(K010737) IS NULL THEN NULL ELSE ROUND(AVG(COALESCE(K010737,0)),0) END AS 果樹・牛馬等, CASE WHEN AVG(K010909) IS NULL THEN NULL ELSE ROUND(AVG(COALESCE(K010909,0)),0) END AS 流動負債, CASE WHEN AVG(K010911) IS NULL THEN NULL ELSE ROUND(AVG(COALESCE(K010911,0)),0) END AS 買掛金, CASE WHEN AVG(K010913) IS NULL THEN NULL ELSE ROUND(AVG(COALESCE(K010913,0)),0) END AS 短期借入金, CASE WHEN AVG(K010919) IS NULL THEN NULL ELSE ROUND(AVG(COALESCE(K010919,0)),0) END AS 長期借入金 ")
                            .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(ComConst.調査区分.営農類型別経営統計_個人)(0)))
                            .AppendLine("WHERE  調査年 = @調査年 ")
                            .AppendLine("AND    K000018 IN (1,2) ")
                            .AppendLine("AND    K000012 > 0 ")
                            .AppendLine(String.Format("AND    K000002 IN ({0}) ", String.Join(",", 農業地域コード.Value)))
                            .AppendLine("AND    K000005 = @営農類型 ")
                            .AppendLine(String.Format("AND    K000006 IN ({0}) ", String.Join(",", If(営農規模コード = 0, String.Join(",", 営農類型コード_営農規模コード.Value), CStr(営農規模コード)))))
                        End With

                        para = New List(Of DBAccess.Parameter)
                        para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
                        para.Add(db.CreateParameter("@営農類型", SqlDbType.Int, 営農類型コード_営農規模コード.Key))

                        Dim dataTable = db.GetDataTable(sb.ToString, para)

                        For Each tableRow As DataRow In dataTable.Rows

                            Dim row As New ComUtil.営農欠測値適用平均値データ With {
                                .調査年 = CInt(chosaNen),
                                .農業地域 = 農業地域コード.Key,
                                .営農類型 = 営農類型コード_営農規模コード.Key,
                                .営農規模 = 営農規模コード,
                                .集計数 = ComUtil.GetZeroOrDec(dataTable.Rows(0).Item("レコード数")),
                                .現金 = ComUtil.GetNullOrDec(tableRow.Item("現金")),
                                .預金等 = ComUtil.GetNullOrDec(tableRow.Item("預金等")),
                                .売掛金＿未収金 = ComUtil.GetNullOrDec(tableRow.Item("売掛金・未収金")),
                                .自動車＿農機具 = ComUtil.GetNullOrDec(tableRow.Item("自動車・農機具")),
                                .建物＿構築物 = ComUtil.GetNullOrDec(tableRow.Item("建物・構築物")),
                                .土地 = ComUtil.GetNullOrDec(tableRow.Item("土地")),
                                .果樹＿牛馬等 = ComUtil.GetNullOrDec(tableRow.Item("果樹・牛馬等")),
                                .流動負債 = ComUtil.GetNullOrDec(tableRow.Item("流動負債")),
                                .買掛金 = ComUtil.GetNullOrDec(tableRow.Item("買掛金")),
                                .短期借入金 = ComUtil.GetNullOrDec(tableRow.Item("短期借入金")),
                                .長期借入金 = ComUtil.GetNullOrDec(tableRow.Item("長期借入金"))
                            }

                            rowList.Add(row)
                        Next
                        'REV_008 END-----------------
                    Next
                Next
            Next

        Catch ex As Exception
            Throw ex
        End Try

        Return rowList

    End Function

    ''' <summary>
    ''' 「営農欠測値適用平均値データ」テーブルからレコードを抽出する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="whereAnd"></param>
    ''' <param name="orderBy"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function SelectAverageDataTable(ByVal db As DBAccess, ByVal chosaNen As String, Optional ByVal whereAnd As String = "", Optional ByVal orderBy As String = "") As DataTable

        If db Is Nothing Then
            Return Nothing
        End If

        If String.IsNullOrWhiteSpace(chosaNen) Then
            Return Nothing
        End If

        Dim dataTable As DataTable = Nothing
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            'SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   営農欠測値適用平均値データ")
                .AppendLine("WHERE  調査年 = @調査年")

                If Not whereAnd = "" Then
                    .AppendLine(whereAnd)
                End If

                If Not orderBy = "" Then
                    .AppendLine(orderBy)
                End If

            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))

            dataTable = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return dataTable

    End Function

    ''' <summary>
    ''' 「営農欠測値平均値代入結果」テーブルからレコードを抽出する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="whereAnd"></param>
    ''' <param name="orderBy"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function SelectAverageResultTable(ByVal db As DBAccess, ByVal chosaNen As String, Optional ByVal whereAnd As String = "", Optional ByVal orderBy As String = "") As DataTable

        If db Is Nothing Then
            Return Nothing
        End If

        If String.IsNullOrWhiteSpace(chosaNen) Then
            Return Nothing
        End If

        Dim dataTable As DataTable = Nothing
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            'SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   営農欠測値平均値代入結果")
                .AppendLine("WHERE  調査年 = @調査年")

                If Not whereAnd = "" Then
                    .AppendLine(whereAnd)
                End If

                If Not orderBy = "" Then
                    .AppendLine(orderBy)
                End If

            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))

            dataTable = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return dataTable

    End Function

    ''' <summary>
    ''' データ削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteTable(ByVal db As DBAccess, ByVal chosaNen As String, ByVal tableName As String) As Boolean

        If db Is Nothing Then
            Return False
        End If

        If Not ComUtil.TryParseToInteger(chosaNen) Then
            Return False
        End If

        If String.IsNullOrWhiteSpace(tableName) Then
            Return False
        End If

        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine(String.Format("FROM   ""{0}""", tableName))
                .AppendLine("WHERE  調査年         = @調査年 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception(String.Format("{0}削除失敗", tableName))
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret

    End Function

    ''' <summary>
    ''' データ挿入
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="rowList"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertAverageTable(ByVal db As DBAccess, ByVal rowList As List(Of ComUtil.営農欠測値適用平均値データ))

        If db Is Nothing Then
            Return
        End If

        If rowList Is Nothing Then
            Return
        End If

        Try
            rowList.ForEach(
                Sub(row)
                    Dim sb As New System.Text.StringBuilder
                    Dim para As New List(Of DBAccess.Parameter)

                    'SQL文の設定
                    With sb
                        .AppendLine(String.Format("INSERT INTO {0} ", ComConst.欠測値.営農欠測値適用平均値データ))
                        .AppendLine("VALUES (@調査年, @農業地域, @営農類型, @営農規模, @集計数, @現金, @預金等, @売掛金＿未収金, @自動車＿農機具, @建物＿構築物, @土地, @果樹＿牛馬等, @流動負債, @買掛金, @短期借入金, @長期借入金, GETDATE(), @更新者ID)")
                    End With

                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, row.調査年))
                    para.Add(db.CreateParameter("@農業地域", SqlDbType.Int, row.農業地域))
                    para.Add(db.CreateParameter("@営農類型", SqlDbType.Int, row.営農類型))
                    para.Add(db.CreateParameter("@営農規模", SqlDbType.Int, row.営農規模))
                    para.Add(db.CreateParameter("@集計数", SqlDbType.Decimal, row.集計数))
                    para.Add(db.CreateParameter("@現金", SqlDbType.Decimal, row.現金))
                    para.Add(db.CreateParameter("@預金等", SqlDbType.Decimal, row.預金等))
                    para.Add(db.CreateParameter("@売掛金＿未収金", SqlDbType.Decimal, row.売掛金＿未収金))
                    para.Add(db.CreateParameter("@自動車＿農機具", SqlDbType.Decimal, row.自動車＿農機具))
                    para.Add(db.CreateParameter("@建物＿構築物", SqlDbType.Decimal, row.建物＿構築物))
                    para.Add(db.CreateParameter("@土地", SqlDbType.Decimal, row.土地))
                    para.Add(db.CreateParameter("@果樹＿牛馬等", SqlDbType.Decimal, row.果樹＿牛馬等))
                    para.Add(db.CreateParameter("@流動負債", SqlDbType.Decimal, row.流動負債))
                    para.Add(db.CreateParameter("@買掛金", SqlDbType.Decimal, row.買掛金))
                    para.Add(db.CreateParameter("@短期借入金", SqlDbType.Decimal, row.短期借入金))
                    para.Add(db.CreateParameter("@長期借入金", SqlDbType.Decimal, row.長期借入金))
                    para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

                    '更新処理
                    db.ExecuteNonQuery(sb.ToString, para)
                End Sub)
        Catch ex As Exception
            Throw
        End Try

        Return

    End Sub

    ''' <summary>
    ''' データ挿入
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertAverageResultTable(ByVal db As DBAccess, ByVal row As ComUtil.営農欠測値平均値代入結果)

        If db Is Nothing Then
            Return
        End If

        If row Is Nothing Then
            Return
        End If

        Try
            Dim sb As New System.Text.StringBuilder
            Dim para As New List(Of DBAccess.Parameter)

            'SQL文の設定
            With sb
                .AppendLine(String.Format("INSERT INTO {0} ", ComConst.欠測値.営農欠測値平均値代入結果))
                .AppendLine("VALUES (@調査年, @センサス番号, @営農類型, @営農規模, @適用＿営農類型, @適用＿営農規模, @適用＿農業地域, @適用＿不可, GETDATE(), @更新者ID)")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, row.調査年))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, row.センサス番号))
            para.Add(db.CreateParameter("@営農類型", SqlDbType.Decimal, row.営農類型))
            para.Add(db.CreateParameter("@営農規模", SqlDbType.Decimal, row.営農規模))
            para.Add(db.CreateParameter("@適用＿営農類型", SqlDbType.Decimal, row.適用＿営農類型))
            para.Add(db.CreateParameter("@適用＿営農規模", SqlDbType.Decimal, row.適用＿営農規模))
            para.Add(db.CreateParameter("@適用＿農業地域", SqlDbType.Decimal, row.適用＿農業地域))
            para.Add(db.CreateParameter("@適用＿不可", SqlDbType.Decimal, row.適用＿不可))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            '更新処理
            db.ExecuteNonQuery(sb.ToString, para)

        Catch ex As Exception
            Throw
        End Try

        Return

    End Sub

    ''' <summary>
    ''' 還元管理取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetReduceManagement(db As DBAccess, chosakubun As String, chosaNen As String, kyoku As String, jimusho As String, kyoten As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT 集計番号 ")
                .AppendLine("     , 還元日時 ")
                .AppendLine("     , 農政局 ")
                .AppendLine("     , 都道府県 ")
                .AppendLine("     , 実査設置拠点 ")
                .AppendLine("FROM   還元管理")
                .AppendLine("WHERE  調査区分         = @調査区分 ")
                .AppendLine("AND    調査年           = @調査年 ")
                If Not kyoku Is Nothing Then
                    .AppendLine("AND    農政局           = @農政局 ")
                End If
                If Not jimusho Is Nothing Then
                    .AppendLine("AND    都道府県         = @都道府県 ")
                End If
                If Not kyoten Is Nothing Then
                    .AppendLine("AND    実査設置拠点     = @実査設置拠点 ")
                End If
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            If Not kyoku Is Nothing Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kyoku))
            End If
            If Not jimusho Is Nothing Then
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))
            End If
            If Not kyoten Is Nothing Then
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kyoten))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 還元管理データ存在チェック
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="dc"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CheckReduceManagementExist(db As DBAccess, chosakubun As String, dc As Dictionary(Of String, String)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   還元管理")
                .AppendLine("WHERE  調査区分         = @調査区分 ")
                .AppendLine("AND    調査年           = @調査年 ")
                .AppendLine("AND    集計番号         = @集計番号 ")
                .AppendLine("AND    農政局           = @農政局 ")
                .AppendLine("AND    都道府県         = @都道府県 ")
                .AppendLine("AND    実査設置拠点     = @実査設置拠点 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, dc("調査年")))
            para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, dc("集計番号")))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, dc("農政局")))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, dc("都道府県")))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, dc("実査設置拠点")))

            dr = db.ExecuteReader(sb.ToString, para)

            ret = dr.HasRows
        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 還元管理追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertReduceManagement(db As DBAccess, chosakubun As String, dc As Dictionary(Of String, String)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 還元管理 ")
                .AppendLine("( ")
                .AppendLine("   調査区分 ")
                .AppendLine("  ,調査年 ")
                .AppendLine("  ,集計番号 ")
                .AppendLine("  ,農政局 ")
                .AppendLine("  ,都道府県 ")
                .AppendLine("  ,実査設置拠点 ")
                .AppendLine("  ,還元日時 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @調査区分 ")
                .AppendLine("  ,@調査年 ")
                .AppendLine("  ,@集計番号 ")
                .AppendLine("  ,@農政局 ")
                .AppendLine("  ,@都道府県 ")
                .AppendLine("  ,@実査設置拠点 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, dc("調査年")))
            para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, dc("集計番号")))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, dc("農政局")))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, dc("都道府県")))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, dc("実査設置拠点")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("還元管理追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 還元管理還元日時更新
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function UpdateReduceManagementSendDate(db As DBAccess, chosakubun As String, dc As Dictionary(Of String, String)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("UPDATE 還元管理 ")
                .AppendLine("SET    還元日時 = GETDATE() ")
                .AppendLine("      ,更新日付 = GETDATE() ")
                .AppendLine("      ,更新者ID = @更新者ID ")
                .AppendLine("WHERE  調査区分         = @調査区分 ")
                .AppendLine("AND    調査年           = @調査年 ")
                .AppendLine("AND    集計番号         = @集計番号 ")
                .AppendLine("AND    農政局           = @農政局 ")
                .AppendLine("AND    都道府県         = @都道府県 ")
                .AppendLine("AND    実査設置拠点     = @実査設置拠点 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, dc("調査年")))
            para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, dc("集計番号")))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, dc("農政局")))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, dc("都道府県")))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, dc("実査設置拠点")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("還元管理還元日時更新失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 還元資料設定値取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <param name="itemNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetKangenShiryoSetting(db As DBAccess, censusNo As String, Optional chosakubun As String = "") As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT 項目番号, 明細番号, 設定値 ")
                .AppendLine("FROM   還元資料設定値")
                .AppendLine("WHERE  調査区分         = @調査区分 ")
                .AppendLine("AND    センサス番号     = @センサス番号 ")
                '.AppendLine("AND    項目番号         = @項目番号 ")
                .AppendLine("ORDER BY 明細番号")
            End With

            If String.IsNullOrEmpty(chosakubun) Then
                chosakubun = CommonInfo.Chosakubun
            End If

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, censusNo))
            'para.Add(db.CreateParameter("@項目番号", SqlDbType.Int, itemNo))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 還元資料設定値追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <param name="itemNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertKangenShiryoSetting(db As DBAccess, dc As Dictionary(Of String, String)) As Boolean

        Dim ret As Boolean
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 還元資料設定値 ")
                .AppendLine("( ")
                .AppendLine("   調査区分 ")
                .AppendLine("  ,センサス番号 ")
                .AppendLine("  ,項目番号 ")
                .AppendLine("  ,明細番号 ")
                .AppendLine("  ,設定値 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @調査区分 ")
                .AppendLine("  ,@センサス番号 ")
                .AppendLine("  ,@項目番号 ")
                .AppendLine("  ,@明細番号 ")
                .AppendLine("  ,@設定値 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, dc("センサス番号")))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.Int, dc("項目番号")))
            para.Add(db.CreateParameter("@明細番号", SqlDbType.Int, dc("明細番号")))
            para.Add(db.CreateParameter("@設定値", SqlDbType.NVarChar, dc("設定値")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("還元資料設定値追加失敗")
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 還元資料設定値更新
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <param name="itemNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function UpdateKangenShiryoSetting(db As DBAccess, dc As Dictionary(Of String, String)) As Boolean

        Dim ret As Boolean
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("UPDATE 還元資料設定値 ")
                .AppendLine("SET 設定値 = @設定値 ")
                .AppendLine("   ,更新日付 = GETDATE() ")
                .AppendLine("   ,更新者ID = @更新者ID ")
                .AppendLine("WHERE   調査区分 = @調査区分 ")
                .AppendLine("AND     センサス番号 = @センサス番号 ")
                .AppendLine("AND     項目番号 = @項目番号 ")
                .AppendLine("AND     明細番号 = @明細番号 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, dc("センサス番号")))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.Int, dc("項目番号")))
            para.Add(db.CreateParameter("@明細番号", SqlDbType.Int, dc("明細番号")))
            para.Add(db.CreateParameter("@設定値", SqlDbType.NVarChar, dc("設定値")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("還元資料設定値更新失敗")
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    '--- REV.001 ADD START
    ''' <summary>
    ''' 同一期間/同一取込名称牛トレサ取込履歴データ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="fromDate"></param>
    ''' <param name="toDate"></param>
    ''' <param name="importName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSamePeriodTresaData(db As DBAccess, fromDate As String, toDate As String, importName As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("SELECT * FROM 牛トレサ取込履歴 ")
                .AppendLine("WHERE 期間＿開始 = @期間_開始 ")
                .AppendLine(" AND 期間＿終了 = @期間_終了 ")
                .AppendLine(" AND 取込データ名称 = @取込データ名称")
            End With

            para.Add(db.CreateParameter("@期間_開始", SqlDbType.VarChar, fromDate))
            para.Add(db.CreateParameter("@期間_終了", SqlDbType.VarChar, toDate))
            para.Add(db.CreateParameter("@取込データ名称", SqlDbType.VarChar, importName))

            ret = db.GetDataTable(sb.ToString, para)
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 重複期間牛トレサ取込履歴データ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="fromDate"></param>
    ''' <param name="toDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetDuplicatePeriodTresaData(db As DBAccess, fromDate As String, toDate As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("SELECT * FROM 牛トレサ取込履歴 ")
                .AppendLine("WHERE (期間＿開始 <= @期間_開始 AND @期間_開始 <= 期間＿終了) ")
                .AppendLine(" OR (期間＿開始 <= @期間_終了 AND @期間_終了 <= 期間＿終了)")
            End With

            para.Add(db.CreateParameter("@期間_開始", SqlDbType.VarChar, fromDate))
            para.Add(db.CreateParameter("@期間_終了", SqlDbType.VarChar, toDate))

            ret = db.GetDataTable(sb.ToString, para)
        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ取込履歴追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="txtFrom"></param>
    ''' <param name="txtTo"></param>
    ''' <param name="txtImportName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertTresaImportHistory(db As DBAccess, txtFrom As String, txtTo As String, txtImportName As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("INSERT INTO 牛トレサ取込履歴 ")
                .AppendLine("( ")
                .AppendLine(" 期間＿開始 ")
                .AppendLine(" , 期間＿終了 ")
                .AppendLine(" , 取込データ名称 ")
                .AppendLine(" , 更新日付 ")
                .AppendLine(" , 更新者ID ")
                .AppendLine(") ")
                .AppendLine(" VALUES ")
                .AppendLine("( ")
                .AppendLine(" @期間_開始 ")
                .AppendLine(" , @期間_終了 ")
                .AppendLine(" , @取込データ名称 ")
                .AppendLine(" , GETDATE() ")
                .AppendLine(" , @更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@期間_開始", SqlDbType.VarChar, txtFrom))
            para.Add(db.CreateParameter("@期間_終了", SqlDbType.VarChar, txtTo))
            para.Add(db.CreateParameter("@取込データ名称", SqlDbType.VarChar, txtImportName))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("牛トレサ取込履歴追加失敗")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ個体情報取込履歴追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="lstKotai"></param>
    ''' <param name="iRirekiNum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertTresaKotaiHistory(db As DBAccess, lstKotai As String(), iRirekiNum As Integer) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("INSERT INTO 牛トレサ個体情報取込履歴 ")
                .AppendLine("( ")
                .AppendLine(" 履歴番号 ")
                .AppendLine(" , 農家団体コード ")
                .AppendLine(" , 個体識別番号 ")
                .AppendLine(" , 農家団体区分CD ")
                .AppendLine(" , 牛の識別CD ")
                .AppendLine(" , 性別コード ")
                .AppendLine(" , 生年月日 ")
                .AppendLine(" , 母牛個体識別番号 ")
                .AppendLine(" , 輸入国CD ")
                .AppendLine(" , 輸入年月日 ")
                .AppendLine(" , 検疫所コード ")
                .AppendLine(" , アクティブ牛フラグ ")
                .AppendLine(" , 照会日 ")
                .AppendLine(") ")
                .AppendLine(" VALUES ")
                .AppendLine("( ")
                .AppendLine(" @履歴番号 ")
                .AppendLine(" , @農家団体コード ")
                .AppendLine(" , @個体識別番号 ")
                .AppendLine(" , @農家団体区分CD ")
                .AppendLine(" , @牛の識別CD ")
                .AppendLine(" , @性別コード ")
                .AppendLine(" , @生年月日 ")
                .AppendLine(" , @母牛個体識別番号 ")
                .AppendLine(" , @輸入国CD ")
                .AppendLine(" , @輸入年月日 ")
                .AppendLine(" , @検疫所コード ")
                .AppendLine(" , @アクティブ牛フラグ ")
                .AppendLine(" , @照会日 ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@履歴番号", SqlDbType.Int, iRirekiNum))
            para.Add(db.CreateParameter("@農家団体コード", SqlDbType.VarChar, lstKotai(1).ToString.TrimStart({"0"c})))
            para.Add(db.CreateParameter("@個体識別番号", SqlDbType.VarChar, lstKotai(0).ToString))
            para.Add(db.CreateParameter("@農家団体区分CD", SqlDbType.Int, lstKotai(2).ToString))
            para.Add(db.CreateParameter("@牛の識別CD", SqlDbType.Int, lstKotai(3).ToString))
            para.Add(db.CreateParameter("@性別コード", SqlDbType.Int, lstKotai(4).ToString))
            para.Add(db.CreateParameter("@生年月日", SqlDbType.Int, lstKotai(5).ToString))
            para.Add(db.CreateParameter("@母牛個体識別番号", SqlDbType.VarChar, lstKotai(6).ToString))
            para.Add(db.CreateParameter("@輸入国CD", SqlDbType.VarChar, lstKotai(7).ToString))
            para.Add(db.CreateParameter("@輸入年月日", SqlDbType.Int, lstKotai(8).ToString))
            para.Add(db.CreateParameter("@検疫所コード", SqlDbType.VarChar, lstKotai(9).ToString))
            para.Add(db.CreateParameter("@アクティブ牛フラグ", SqlDbType.Int, lstKotai(10).ToString))
            para.Add(db.CreateParameter("@照会日", SqlDbType.Int, lstKotai(11).ToString))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("牛トレサ個体情報取込履歴追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ異動情報取込履歴追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="lstIdo"></param>
    ''' <param name="iRirekiNum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertTresaIdoHistory(db As DBAccess, lstIdo As String(), iRirekiNum As Integer) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("INSERT INTO 牛トレサ異動情報取込履歴 ")
                .AppendLine("( ")
                .AppendLine(" 履歴番号 ")
                .AppendLine(" , 農家団体コード ")
                .AppendLine(" , 個体識別番号 ")
                .AppendLine(" , 異動年月日 ")
                .AppendLine(" , 異動フラグ ")
                .AppendLine(" , 農家団体区分CD ")
                .AppendLine(" , 農家団体名称 ")
                .AppendLine(" , 飼養地都道府県名 ")
                .AppendLine(" , 飼養所在地 ")
                .AppendLine(") ")
                .AppendLine(" VALUES ")
                .AppendLine("( ")
                .AppendLine(" @履歴番号 ")
                .AppendLine(" , @農家団体コード ")
                .AppendLine(" , @個体識別番号 ")
                .AppendLine(" , @異動年月日 ")
                .AppendLine(" , @異動フラグ ")
                .AppendLine(" , @農家団体区分CD ")
                .AppendLine(" , @農家団体名称 ")
                .AppendLine(" , @飼養地都道府県名 ")
                .AppendLine(" , @飼養所在地 ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@履歴番号", SqlDbType.Int, iRirekiNum))
            para.Add(db.CreateParameter("@農家団体コード", SqlDbType.VarChar, lstIdo(3).ToString.TrimStart({"0"c})))
            para.Add(db.CreateParameter("@個体識別番号", SqlDbType.VarChar, lstIdo(0).ToString))
            para.Add(db.CreateParameter("@異動年月日", SqlDbType.Int, lstIdo(1).ToString))
            para.Add(db.CreateParameter("@異動フラグ", SqlDbType.Int, lstIdo(2).ToString))
            para.Add(db.CreateParameter("@農家団体区分CD", SqlDbType.Int, lstIdo(4).ToString))
            para.Add(db.CreateParameter("@農家団体名称", SqlDbType.VarChar, lstIdo(5).ToString))
            para.Add(db.CreateParameter("@飼養地都道府県名", SqlDbType.VarChar, lstIdo(6).ToString))
            para.Add(db.CreateParameter("@飼養所在地", SqlDbType.VarChar, lstIdo(7).ToString))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("牛トレサ異動情報取込履歴追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ個体情報削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteTresaKotaiInfo(db As DBAccess) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("DELETE FROM 牛トレサ個体情報 ")
            End With

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("牛トレサ個体情報削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ個体情報追加(更新)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertTresaKotaiInfo(db As DBAccess) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("INSERT INTO 牛トレサ個体情報 ")
                .AppendLine("( ")
                .AppendLine(" 農家団体コード ")
                .AppendLine(" , 個体識別番号 ")
                .AppendLine(" , 農家団体区分CD ")
                .AppendLine(" , 牛の識別CD ")
                .AppendLine(" , 性別コード ")
                .AppendLine(" , 生年月日 ")
                .AppendLine(" , 母牛個体識別番号 ")
                .AppendLine(" , 輸入国CD ")
                .AppendLine(" , 輸入年月日 ")
                .AppendLine(" , 検疫所コード ")
                .AppendLine(" , アクティブ牛フラグ ")
                .AppendLine(" , 照会日 ")
                .AppendLine(" , 履歴番号 ")
                .AppendLine(") ")
                .AppendLine("SELECT ")
                .AppendLine(" 牛トレサ個体情報取込履歴.農家団体コード ")
                .AppendLine(" , CONVERT(varchar, CONVERT(decimal, SUBSTRING(牛トレサ個体情報取込履歴.個体識別番号, 6, 10))) ")
                .AppendLine(" , 農家団体区分CD ")
                .AppendLine(" , 牛の識別CD ")
                .AppendLine(" , 性別コード ")
                .AppendLine(" , 生年月日 ")
                'REV_003 Mod Start
                '.AppendLine(" , CONVERT(varchar, CONVERT(decimal, SUBSTRING(母牛個体識別番号, 6, 10))) ")
                .AppendLine(" , IIF(LEN(母牛個体識別番号)>=15,CONVERT(varchar, CONVERT(decimal, SUBSTRING(母牛個体識別番号, 6, 10))),母牛個体識別番号)  ")
                'REV_003 Mod Start
                .AppendLine(" , 輸入国CD ")
                .AppendLine(" , 輸入年月日 ")
                .AppendLine(" , 検疫所コード ")
                .AppendLine(" , アクティブ牛フラグ ")
                .AppendLine(" , 照会日 ")
                .AppendLine(" , 牛トレサ個体情報取込履歴.履歴番号 ")
                .AppendLine(" FROM 牛トレサ個体情報取込履歴 ")
                .AppendLine(" INNER JOIN ( ")
                .AppendLine(" SELECT MAX(履歴番号) AS 履歴番号 ")
                .AppendLine(" ,農家団体コード ")
                .AppendLine(" ,個体識別番号 ")
                .AppendLine(" FROM 牛トレサ個体情報取込履歴 ")
                .AppendLine(" GROUP BY 農家団体コード,個体識別番号) 最大履歴牛トレサ個体情報取込履歴 ")
                .AppendLine(" ON ")
                .AppendLine(" 牛トレサ個体情報取込履歴.農家団体コード = 最大履歴牛トレサ個体情報取込履歴.農家団体コード ")
                .AppendLine(" AND 牛トレサ個体情報取込履歴.個体識別番号 = 最大履歴牛トレサ個体情報取込履歴.個体識別番号 ")
                .AppendLine(" AND 牛トレサ個体情報取込履歴.履歴番号 = 最大履歴牛トレサ個体情報取込履歴.履歴番号 ")
            End With

            If db.ExecuteNonQuery(sb.ToString, para) >= 1 Then
                ret = True
            Else
                Throw New Exception("牛トレサ個体情報更新失敗")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ異動情報削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteTresaIdoInfo(db As DBAccess) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("DELETE FROM 牛トレサ異動情報 ")
            End With

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("牛トレサ異動情報削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ異動情報追加(更新)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertTresaIdoInfo(db As DBAccess) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("INSERT INTO 牛トレサ異動情報 ")
                .AppendLine("( ")
                .AppendLine(" 農家団体コード ")
                .AppendLine(" , 個体識別番号 ")
                .AppendLine(" , 異動年月日 ")
                .AppendLine(" , 異動フラグ ")
                .AppendLine(" , 農家団体区分CD ")
                .AppendLine(" , 農家団体名称 ")
                .AppendLine(" , 飼養地都道府県名 ")
                .AppendLine(" , 飼養所在地 ")
                .AppendLine(" , 履歴番号 ")
                .AppendLine(") ")
                .AppendLine("SELECT ")
                .AppendLine(" 牛トレサ異動情報取込履歴.農家団体コード ")
                .AppendLine(" , CONVERT(varchar, CONVERT(decimal, SUBSTRING(牛トレサ異動情報取込履歴.個体識別番号, 6, 10))) ")
                .AppendLine(" , 牛トレサ異動情報取込履歴.異動年月日 ")
                .AppendLine(" , 牛トレサ異動情報取込履歴.異動フラグ ")
                .AppendLine(" , 農家団体区分CD ")
                .AppendLine(" , 農家団体名称 ")
                .AppendLine(" , 飼養地都道府県名 ")
                .AppendLine(" , 飼養所在地 ")
                .AppendLine(" , 牛トレサ異動情報取込履歴.履歴番号 ")
                .AppendLine(" FROM 牛トレサ異動情報取込履歴 ")
                .AppendLine(" INNER JOIN ( ")
                .AppendLine(" SELECT MAX(履歴番号) AS 履歴番号 ")
                .AppendLine(" ,農家団体コード ")
                .AppendLine(" ,個体識別番号 ")
                .AppendLine(" ,異動年月日 ")
                .AppendLine(" ,異動フラグ ")
                .AppendLine(" FROM 牛トレサ異動情報取込履歴 ")
                .AppendLine(" GROUP BY 農家団体コード,個体識別番号, 異動年月日,異動フラグ) 最大履歴牛トレサ異動情報取込履歴 ")
                .AppendLine(" ON ")
                .AppendLine(" 牛トレサ異動情報取込履歴.農家団体コード = 最大履歴牛トレサ異動情報取込履歴.農家団体コード ")
                .AppendLine(" AND 牛トレサ異動情報取込履歴.個体識別番号 = 最大履歴牛トレサ異動情報取込履歴.個体識別番号 ")
                .AppendLine(" AND 牛トレサ異動情報取込履歴.異動年月日 = 最大履歴牛トレサ異動情報取込履歴.異動年月日 ")
                .AppendLine(" AND 牛トレサ異動情報取込履歴.異動フラグ = 最大履歴牛トレサ異動情報取込履歴.異動フラグ ")
                .AppendLine(" AND 牛トレサ異動情報取込履歴.履歴番号 = 最大履歴牛トレサ異動情報取込履歴.履歴番号 ")
            End With

            If db.ExecuteNonQuery(sb.ToString, para) >= 1 Then
                ret = True
            Else
                Throw New Exception("牛トレサ個体情報更新失敗")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ取込履歴取得(牛トレサデータ管理のリスト表示用)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="fromDate"></param>
    ''' <param name="toDate"></param>
    ''' <param name="keika"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTresaData(db As DBAccess, fromDate As String, toDate As String, keika As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim strWHERE As String


        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("SELECT ")
                .AppendLine(" SUBSTRING(牛トレサ管理データ.期間＿開始,1,4) AS 期間開始年 ")
                .AppendLine(" ,SUBSTRING(牛トレサ管理データ.期間＿開始,5,2) AS 期間開始月 ")
                .AppendLine(" ,SUBSTRING(牛トレサ管理データ.期間＿終了,1,4) AS 期間終了年 ")
                .AppendLine(" ,SUBSTRING(牛トレサ管理データ.期間＿終了,5,2) AS 期間終了月 ")
                .AppendLine(" ,牛トレサ管理データ.取込データ名称 ")
                .AppendLine(" ,FORMAT(牛トレサ管理データ.更新日付, 'yyyyMMdd') AS データ登録日 ")
                .AppendLine(" ,IIF(牛トレサ管理データ.更新日付 <= CONVERT(date,DATEADD(year,-2,GETDATE())),'2年以上経過','－') AS データ登録以降2年経過情報 ")
                .AppendLine(" ,COUNT(牛トレサ管理データ.個体履歴番号) AS 個体情報レコード数 ")
                .AppendLine(" ,COUNT(牛トレサ管理データ.異動履歴番号) AS 異動情報レコード数 ")
                .AppendLine(" ,牛トレサ管理データ.履歴番号 ")
                .AppendLine(" FROM ( ")
                .AppendLine(" ( ")
                .AppendLine(" SELECT ")
                .AppendLine(" 牛トレサ取込履歴.履歴番号 ")
                .AppendLine(" ,牛トレサ取込履歴.期間＿開始 ")
                .AppendLine(" ,牛トレサ取込履歴.期間＿終了 ")
                .AppendLine(" ,牛トレサ取込履歴.取込データ名称 ")
                .AppendLine(" ,牛トレサ取込履歴.更新者ID ")
                .AppendLine(" ,牛トレサ取込履歴.更新日付 ")
                .AppendLine(" , 牛トレサ個体情報取込履歴.履歴番号 AS 個体履歴番号 ")
                .AppendLine(" , '異動履歴番号' = null ")
                .AppendLine(" FROM 牛トレサ取込履歴 ")
                .AppendLine(" LEFT JOIN 牛トレサ個体情報取込履歴 ON ")
                .AppendLine(" 牛トレサ取込履歴.履歴番号 = 牛トレサ個体情報取込履歴.履歴番号")
                .AppendLine(" ) ")
                .AppendLine(" UNION ALL ")
                .AppendLine(" ( ")
                .AppendLine(" SELECT ")
                .AppendLine(" 牛トレサ取込履歴.履歴番号 ")
                .AppendLine(" ,牛トレサ取込履歴.期間＿開始 ")
                .AppendLine(" ,牛トレサ取込履歴.期間＿終了 ")
                .AppendLine(" ,牛トレサ取込履歴.取込データ名称 ")
                .AppendLine(" ,牛トレサ取込履歴.更新者ID ")
                .AppendLine(" ,牛トレサ取込履歴.更新日付 ")
                .AppendLine(" , '個体履歴番号'=null ")
                .AppendLine(" , 牛トレサ異動情報取込履歴.履歴番号 AS 異動履歴番号 ")
                .AppendLine(" FROM 牛トレサ取込履歴 ")
                .AppendLine(" LEFT JOIN 牛トレサ異動情報取込履歴 ON ")
                .AppendLine(" 牛トレサ取込履歴.履歴番号 = 牛トレサ異動情報取込履歴.履歴番号")
                .AppendLine(" ) ")
                .AppendLine(" ) AS 牛トレサ管理データ ")

                strWHERE = ""

                If Not (fromDate = "00" And toDate = "00" And keika = "") Then
                    .AppendLine(" WHERE ")
                    If fromDate <> "00" Then
                        strWHERE = " (( @期間_開始 <= 牛トレサ管理データ.期間＿開始) "
                    End If

                    If toDate <> "00" Then
                        If strWHERE = "" Then
                            strWHERE = strWHERE & " ((牛トレサ管理データ.期間＿終了 <= @期間_終了)) "
                        Else
                            strWHERE = strWHERE & " OR (牛トレサ管理データ.期間＿終了 <= @期間_終了)) "
                        End If
                    Else
                        If strWHERE <> "" Then
                            strWHERE = strWHERE & ")"
                        End If
                    End If

                    If keika IsNot Nothing Then
                        If strWHERE = "" Then
                            If keika = "2年以上経過" Then
                                strWHERE = strWHERE & " (牛トレサ管理データ.更新日付 <= CONVERT(date,DATEADD(year,-2,GETDATE()))) "
                            Else
                                strWHERE = strWHERE & " (牛トレサ管理データ.更新日付 > CONVERT(date,DATEADD(year,-2,GETDATE()))) "
                            End If
                        Else
                            If keika = "2年以上経過" Then
                                strWHERE = strWHERE & " AND (牛トレサ管理データ.更新日付 <= CONVERT(date,DATEADD(year,-2,GETDATE()))) "
                            Else
                                strWHERE = strWHERE & " AND (牛トレサ管理データ.更新日付 > CONVERT(date,DATEADD(year,-2,GETDATE()))) "
                            End If
                        End If
                    End If
                End If

                If strWHERE <> "" Then
                    .AppendLine(strWHERE)
                End If
                .AppendLine(" GROUP BY 牛トレサ管理データ.履歴番号,牛トレサ管理データ.期間＿開始,牛トレサ管理データ.期間＿終了,牛トレサ管理データ.取込データ名称,牛トレサ管理データ.更新日付 ")
                .AppendLine(" ORDER BY 牛トレサ管理データ.履歴番号")
            End With

            para.Add(db.CreateParameter("@期間_開始", SqlDbType.VarChar, fromDate))
            para.Add(db.CreateParameter("@期間_終了", SqlDbType.VarChar, toDate))

            ret = db.GetDataTable(sb.ToString, para)
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ個体情報取込履歴データ取得(CSV出力用)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="iRirekiNum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTresaKotaiHistory(db As DBAccess, iRirekiNum As Integer) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("SELECT ")
                .AppendLine(" 個体識別番号 ")
                .AppendLine(" , 農家団体コード ")
                .AppendLine(" , 農家団体区分CD ")
                .AppendLine(" , 牛の識別CD ")
                .AppendLine(" , 性別コード ")
                .AppendLine(" , 生年月日 ")
                .AppendLine(" , 母牛個体識別番号 ")
                .AppendLine(" , 輸入国CD ")
                .AppendLine(" , 輸入年月日 ")
                .AppendLine(" , 検疫所コード ")
                .AppendLine(" , アクティブ牛フラグ ")
                .AppendLine(" , 照会日 ")
                .AppendLine(" FROM 牛トレサ個体情報取込履歴 ")
                .AppendLine(" WHERE 履歴番号 = @履歴番号 ")
            End With
            para.Add(db.CreateParameter("@履歴番号", SqlDbType.Int, iRirekiNum))

            ret = db.GetDataTable(sb.ToString, para)
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ異動情報取込履歴データ取得(CSV出力用)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="iRirekiNum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTresaIdoHistory(db As DBAccess, iRirekiNum As Integer) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("SELECT ")
                .AppendLine(" 個体識別番号 ")
                .AppendLine(" , 異動年月日 ")
                .AppendLine(" , 異動フラグ ")
                .AppendLine(" , 農家団体コード ")
                .AppendLine(" , 農家団体区分CD ")
                .AppendLine(" , 農家団体名称 ")
                .AppendLine(" , 飼養地都道府県名 ")
                .AppendLine(" , 飼養所在地 ")
                .AppendLine(" FROM 牛トレサ異動情報取込履歴 ")
                .AppendLine(" WHERE 履歴番号 = @履歴番号 ")
            End With
            para.Add(db.CreateParameter("@履歴番号", SqlDbType.Int, iRirekiNum))

            ret = db.GetDataTable(sb.ToString, para)
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ取込履歴更新処理
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="fromDate"></param>
    ''' <param name="toDate"></param>
    ''' <param name="importName"></param>
    ''' <param name="iRirekiNum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function UpdateTresaData(db As DBAccess, fromDate As String, toDate As String, importName As String, iRirekiNum As Integer) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("UPDATE 牛トレサ取込履歴 ")
                .AppendLine(" SET ")
                .AppendLine(" 期間＿開始 = @期間_開始 ")
                .AppendLine(" , 期間＿終了 = @期間_終了 ")
                .AppendLine(" , 取込データ名称 = @取込データ名称 ")
                .AppendLine(" WHERE ")
                .AppendLine(" 履歴番号 = @履歴番号 ")
            End With

            para.Add(db.CreateParameter("@期間＿開始", SqlDbType.VarChar, fromDate))
            para.Add(db.CreateParameter("@期間＿終了", SqlDbType.VarChar, toDate))
            para.Add(db.CreateParameter("@取込データ名称", SqlDbType.VarChar, importName))
            para.Add(db.CreateParameter("@履歴番号", SqlDbType.Int, iRirekiNum))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("牛トレサ取込履歴更新失敗")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ関連データ削除処理
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="iRirekiNum"></param>
    ''' <param name="strTableName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteTresaData(db As DBAccess, iRirekiNum As Integer, strTableName As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("DELETE FROM " & strTableName)
                .AppendLine(" WHERE ")
                .AppendLine(" 履歴番号 = @履歴番号")
            End With

            para.Add(db.CreateParameter("@履歴番号", SqlDbType.Int, iRirekiNum))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception(strTableName & "削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサ取込履歴取得処理
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTresaHistory(db As DBAccess) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("SELECT * FROM ")
                .AppendLine(" 牛トレサ取込履歴")
            End With

            ret = db.GetDataTable(sb.ToString, para)
        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛トレサデータの取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="farmCodeList">農家団体コード</param>
    ''' <param name="idoDateFrom">異動年月日(これ以降のデータを取得)</param>
    ''' <param name="idoDateTo">異動年月日(これ以前のデータを取得)</param>
    ''' <returns>牛トレサデータ</returns>
    ''' <remarks></remarks>
    Public Shared Function GetTresa(db As DBAccess, farmCodeList As List(Of String), idoDateFrom As Integer, idoDateTo As Integer, Optional isIdoOrderAsc As Boolean = False) As DataTable

        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * ")
                .AppendLine("FROM   牛トレサ異動情報 AS I")
                .AppendLine(" LEFT JOIN ( ")
                .AppendLine("   SELECT DISTINCT ")
                .AppendLine("   個体情報.* ")
                .AppendLine("   FROM 牛トレサ個体情報 As 個体情報 ")
                .AppendLine("   INNER JOIN ( ")
                .AppendLine("       SELECT ")
                .AppendLine("           個体識別番号, MAX(照会日) As 照会日 ")
                .AppendLine("       FROM 牛トレサ個体情報 ")
                .AppendLine("       GROUP BY 個体識別番号 ")
                .AppendLine("   ) AS 個体情報＿最新照会日 ")
                .AppendLine("   ON 個体情報.個体識別番号 = 個体情報＿最新照会日.個体識別番号 ")
                .AppendLine("   AND 個体情報.照会日 = 個体情報＿最新照会日.照会日 ")
                .AppendLine(" ) AS K")
                .AppendLine("ON I.個体識別番号 = K.個体識別番号")
                .AppendLine("WHERE  I.農家団体コード IN ('" & String.Join("', '", farmCodeList) & "')")
                .AppendLine("AND    (異動年月日 BETWEEN @異動年月日FROM AND @異動年月日TO) ")
                If isIdoOrderAsc Then
                    .AppendLine("ORDER BY 異動年月日, I.個体識別番号 ")
                Else
                    .AppendLine("ORDER BY I.個体識別番号, 異動年月日 DESC ")
                End If
            End With

            para.Add(db.CreateParameter("@異動年月日FROM", SqlDbType.Int, idoDateFrom))
            para.Add(db.CreateParameter("@異動年月日TO", SqlDbType.Int, idoDateTo))

            dt = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return dt
    End Function

    ''' <summary>
    ''' 成畜牛トレサデータの取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="farmCodeList">農家団体コード</param>
    ''' <param name="idoDateFrom">異動年月日(これ以降のデータを取得)</param>
    ''' <param name="idoDateTo">異動年月日(これ以前のデータを取得)</param>
    ''' <returns>牛トレサデータ</returns>
    ''' <remarks></remarks>
    Public Shared Function GetSeitikuTresa(db As DBAccess, farmCodeList As List(Of String), idoDateFrom As Integer, idoDateTo As Integer, idoDateFrom_y As String, idoDateFrom_m As String, Optional isIdoOrderAsc As Boolean = False) As DataTable

        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * ")
                .AppendLine("FROM   牛トレサ異動情報 AS I")
                .AppendLine(" LEFT JOIN ( ")
                .AppendLine("   SELECT DISTINCT ")
                .AppendLine("   個体情報.* ")
                .AppendLine("   FROM 牛トレサ個体情報 As 個体情報 ")
                .AppendLine("   INNER JOIN ( ")
                .AppendLine("       SELECT ")
                .AppendLine("           個体識別番号, MAX(照会日) As 照会日 ")
                .AppendLine("       FROM 牛トレサ個体情報 ")
                .AppendLine("       GROUP BY 個体識別番号 ")
                .AppendLine("   ) AS 個体情報＿最新照会日 ")
                .AppendLine("   ON 個体情報.個体識別番号 = 個体情報＿最新照会日.個体識別番号 ")
                .AppendLine("   AND 個体情報.照会日 = 個体情報＿最新照会日.照会日 ")
                .AppendLine(" ) AS K")
                .AppendLine("ON I.個体識別番号 = K.個体識別番号")
                .AppendLine("WHERE  I.農家団体コード IN ('" & String.Join("', '", farmCodeList) & "')")
                '↓成畜用追加条件
                .AppendLine(" AND 異動フラグ = '2' ")
                '↓REV_015
                .AppendLine(" AND NOT EXISTS(SELECT * FROM 牛トレサ個体情報 X")
                .AppendLine(" WHERE X.母牛個体識別番号 = K.母牛個体識別番号")
                .AppendLine(" AND X.生年月日 < K.生年月日)")
                '↑REV_015
                '↑成畜用追加条件
                .AppendLine("AND    (異動年月日 BETWEEN @異動年月日FROM AND @異動年月日TO) ")
                If isIdoOrderAsc Then
                    .AppendLine("ORDER BY 異動年月日, I.個体識別番号 ")
                Else
                    .AppendLine("ORDER BY I.個体識別番号, 異動年月日 DESC ")
                End If
            End With

            para.Add(db.CreateParameter("@異動年月日FROM", SqlDbType.Int, idoDateFrom))
            para.Add(db.CreateParameter("@異動年月日TO", SqlDbType.Int, idoDateTo))
            para.Add(db.CreateParameter("@開始年", SqlDbType.Int, CInt(idoDateFrom_y)))
            para.Add(db.CreateParameter("@開始月", SqlDbType.Int, CInt(idoDateFrom_m)))

            dt = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return dt
    End Function

    ''' <summary>
    ''' 母牛個体識別番号が一致する牛トレサデータを生年月日順に取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="cowID">個体識別番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTresaChild(db As DBAccess, cowID As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * ")
                .AppendLine("FROM   牛トレサ個体情報")
                .AppendLine("WHERE  母牛個体識別番号         = @母牛個体識別番号 ")
                .AppendLine("ORDER BY 生年月日")
            End With

            para.Add(db.CreateParameter("@母牛個体識別番号", SqlDbType.VarChar, cowID))

            ret = db.GetDataTable(sb.ToString, para)
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票項目マスタ取得(牛資産異動情報)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="suushikiKubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoItemMasterIdoTresa(db As DBAccess, chosaKubun As String, Optional suushikiKubun As String = Nothing) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where Not val.Contains("＿可変")
            Dim sheetList = ComConst.牛トレサデータ.異動ファイルシート(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT     A.*")
                .AppendLine("         , B.precision AS 有効桁数 ")
                .AppendLine("         , B.scale AS 小数点以下桁数")
                .AppendLine("FROM      (SELECT * ")
                .AppendLine("           FROM   調査票項目マスタ ")
                .AppendLine("           WHERE  調査区分 = @調査区分 ")
                If Not suushikiKubun Is Nothing Then
                    .AppendLine("           AND    数式区分 = @数式区分 ")
                End If
                .AppendLine("          ) A")
                .AppendLine("LEFT JOIN (SELECT * ")
                .AppendLine("           FROM   sys.columns")
                .AppendLine("           WHERE  object_id IN (SELECT object_id")
                .AppendLine("                                FROM   sys.tables")
                .AppendLine("                                WHERE  name IN ('" & String.Join("', '", query) & "')")
                .AppendLine("                               )")
                .AppendLine("          ) B")
                .AppendLine("ON A.項目番号 = B.name")
                .AppendLine(" WHERE A.シート名 IN ('" & String.Join("', '", sheetList) & "') ")
                .AppendLine(" ORDER BY シート名 DESC,項目番号 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            If Not suushikiKubun Is Nothing Then
                para.Add(db.CreateParameter("@数式区分", SqlDbType.Int, suushikiKubun))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 同一期間(異動年月)調査票データの登録済みチェック
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strYearFrom"></param>
    ''' <param name="strMonthFrom"></param>
    ''' <param name="strYearTo"></param>
    ''' <param name="strMonthTo"></param>
    ''' <param name="strCensusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSameIdoKikanChosahyo(db As DBAccess, chosanen As String, chosaKubun As String, strYearFrom As String, strMonthFrom As String, strYearTo As String, strMonthTo As String, strKessan As String, strCensusNo As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Dim tStrMonrhFrom As String
        Dim tStrMonrhTo As String

        If CInt(strMonthFrom) <= CInt(strKessan) Then
            tStrMonrhFrom = CStr(CInt(strMonthFrom) + 12)
        Else
            tStrMonrhFrom = strMonthFrom
        End If

        If CInt(strMonthTo) <= CInt(strKessan) Then
            tStrMonrhTo = CStr(CInt(strMonthTo) + 12)
        Else
            tStrMonrhTo = strMonthTo
        End If

        'EXCELとSQLではシリアル値が2日ずれるため、3から開始する
        Dim dtfrom As New DateTime(CInt(strYearFrom), CInt(strMonthFrom), 3)
        Dim dtto As DateTime
        Dim tmpYearTo As String
        Dim tmpMonthTo As String
        If CInt(strMonthTo) = 12 Then
            tmpYearTo = CStr(CInt(strYearTo) + 1)
            tmpMonthTo = "1"
        Else
            tmpYearTo = strYearTo
            tmpMonthTo = CStr(CInt(strMonthTo) + 1)
        End If
        dtto = New DateTime(CInt(tmpYearTo), CInt(tmpMonthTo), 3)

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                If chosaKubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse chosaKubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
                    .AppendLine("SELECT ")
                    .AppendLine("CASE WHEN A.調査年 Is null THEN B.調査年 ELSE A.調査年  END AS 調査年")
                    .AppendLine(",CASE WHEN A.明細番号 Is null THEN B.明細番号 ELSE A.明細番号  END AS 明細番号")
                    .AppendLine(",CASE WHEN ISNUMERIC(A.値) = 0 THEN CASE WHEN LEN(CONVERT(float,B.値)) = 5 THEN MONTH(CONVERT(DateTime,(CONVERT(float,B.値))) - 2) ELSE A.値 END ELSE A.値 END As 異動月")
                    .AppendLine(" FROM ")
                    .AppendLine(" ( ")
                End If

                .AppendLine("SELECT * FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年 ")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND 項目番号 = @異動月項目番号 ")
                .AppendLine(" AND CASE WHEN ISNUMERIC(値) = 1 THEN ")
                .AppendLine("   CASE WHEN  CAST(値 as integer) <= @決算月 THEN ")
                .AppendLine("    CAST(値 as integer) + 12")
                .AppendLine("   ELSE CAST(値 as integer) END  ")
                .AppendLine("  ELSE 0 END ")
                .AppendLine("between @開始月 and @終了月 ")

                If chosaKubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse chosaKubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
                    .AppendLine(" ) AS A")
                    .AppendLine(" FULL OUTER JOIN ")
                    .AppendLine(" ( ")
                    .AppendLine("SELECT * FROM " & query(0).ToString)
                    .AppendLine(" WHERE ")
                    .AppendLine(" 調査年 = @調査年 ")
                    .AppendLine(" AND センサス番号 = @センサス番号 ")
                    .AppendLine(" AND 項目番号 = @異動年月項目番号 ")
                    .AppendLine("  AND (値 >= CONVERT(varchar, (CONVERT(float, CONVERT(datetime,@開始年月)))) ")
                    .AppendLine("  AND 値 < CONVERT(varchar,CONVERT(float, CONVERT(datetime,@終了年月))))")
                    .AppendLine(")")
                    .AppendLine(" As B ON A.明細番号 = B.明細番号 ")
                End If
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, chosanen))
            para.Add(db.CreateParameter("@開始月", SqlDbType.VarChar, tStrMonrhFrom))
            para.Add(db.CreateParameter("@終了月", SqlDbType.VarChar, tStrMonrhTo))
            para.Add(db.CreateParameter("@決算月", SqlDbType.VarChar, strKessan))
            para.Add(db.CreateParameter("@異動月項目番号", SqlDbType.VarChar, ComConst.牛トレサデータ.異動月_項目番号(chosaKubun)))
            If chosaKubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse chosaKubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
                para.Add(db.CreateParameter("@開始年月", SqlDbType.VarChar, dtfrom.ToString))
                para.Add(db.CreateParameter("@終了年月", SqlDbType.VarChar, dtto.ToString))
                para.Add(db.CreateParameter("@異動年月項目番号", SqlDbType.VarChar, ComConst.牛トレサデータ.異動年月_項目番号(chosaKubun)))
            End If
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 同一期間(異動年月)調査票データ削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strYearFrom"></param>
    ''' <param name="strMonthFrom"></param>
    ''' <param name="strYearTo"></param>
    ''' <param name="strMonthTo"></param>
    ''' <param name="strCensusNo"></param>
    'REV_005 START---------------
    ''' <param name="lngFileType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    'Public Shared Function DeleteSameIdoKikanChosahyo(db As DBAccess, chosaKubun As String, strYear As String, strMeisaiNo As String, strCensusNo As String) As Boolean
    Public Shared Function DeleteSameIdoKikanChosahyo(db As DBAccess, chosaKubun As String, strYear As String, strMeisaiNo As String, strCensusNo As String, lngFileType As Long) As Boolean
        'REV_005 END---------------
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim importNum = ComConst.牛トレサデータ.取込対象_項目番号(chosaKubun)

            ' SQL文の設定（削除対象判定用）
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)
            With sb
                .AppendLine("SELECT * FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年 ")
                .AppendLine(" AND 明細番号 = @明細番号 ")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND 項目番号 = 'Q11020101' ")  '個体識別番号
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, strYear))
            para.Add(db.CreateParameter("@明細番号", SqlDbType.VarChar, strMeisaiNo))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))

            If 0 < db.GetDataTable(sb.ToString, para).Rows.Count And 1 = lngFileType Then
                '対象明細行のQ11020101（個体識別番号）にデータがあり、成畜処理の場合、処理をスキップ（削除対象外）する
                Return True
            ElseIf 0 >= db.GetDataTable(sb.ToString, para).Rows.Count And 0 = lngFileType And (ComConst.調査区分.牛乳生産費統計_個別 = CommonInfo.Chosakubun Or ComConst.調査区分.経営分析調査_牛乳生産費 = CommonInfo.Chosakubun) Then   'REV ADD 2022/05/20
                '対象明細行のQ11020101（個体識別番号）にデータがなく、通常処理かつ調査区分が牛乳の場合、処理をスキップ（削除対象外）する
                Return True
            End If

            ' SQL文の設定
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)
            With sb
                .AppendLine("DELETE FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年 ")
                .AppendLine(" AND 明細番号 = @明細番号 ")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND (項目番号 BETWEEN @下限項目番号 AND @上限項目番号) ")
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, strYear))
            para.Add(db.CreateParameter("@明細番号", SqlDbType.VarChar, strMeisaiNo))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@下限項目番号", SqlDbType.VarChar, importNum(0)))
            para.Add(db.CreateParameter("@上限項目番号", SqlDbType.VarChar, importNum(1)))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("調査票データ削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 電子調査票の登録済みチェック
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetImportChosahyo(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, strChosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 取込時使用する明細番号取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetMaxMeisaiNo(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim importNum = ComConst.牛トレサデータ.取込対象_項目番号(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(明細番号) AS 明細番号最大値 FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND (項目番号 BETWEEN @下限項目番号 AND @上限項目番号) ")
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, strChosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@下限項目番号", SqlDbType.VarChar, importNum(0)))
            para.Add(db.CreateParameter("@上限項目番号", SqlDbType.VarChar, importNum(1)))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 成畜取込時使用する明細番号取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSeitikuMaxMeisaiNo(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim importNum = ComConst.牛トレサデータ.成畜取込対象_項目番号(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(明細番号) AS 明細番号最大値 FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND (項目番号 BETWEEN @下限項目番号 AND @上限項目番号) ")
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, strChosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@下限項目番号", SqlDbType.VarChar, importNum(0)))
            para.Add(db.CreateParameter("@上限項目番号", SqlDbType.VarChar, importNum(1)))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛資産異動情報データ追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="maxMesaiNo"></param>
    ''' <param name="dc"></param>
    ''' <param name="lstImpTarget"></param>
    ''' <param name="clearList">REV_014</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertIdoTresaChosahyoTable(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, chosaKubun As String, maxMesaiNo As String, dc As Dictionary(Of String, DAOChosahyo.調査票項目), clearList As List(Of String)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)
                If tableName.Contains("＿可変") Then
                    Dim dt As DataTable

                    sb = New System.Text.StringBuilder
                    sb.AppendLine(String.Format("SELECT TOP 0 * FROM {0} ", tableName))

                    dt = db.GetDataTable(sb.ToString)

                    Dim row As DataRow = Nothing
                    Dim query = From dr In dc Where dr.Key.Contains(ComConst.ITEM_NO_DELIMITER) Select dr
                    For Each kv As KeyValuePair(Of String, DAOChosahyo.調査票項目) In query
                        Dim no As String() = kv.Key.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))


                        row = dt.NewRow()
                        row.BeginEdit()

                        row("調査年") = pKey.chosaNen
                        row("センサス番号") = pKey.censusNo
                        row("項目番号") = no(0)
                        row("明細番号") = (CInt(maxMesaiNo) + CInt(no(1))).ToString
                        row("値") = kv.Value.値
                        row("更新日付") = Now()
                        row("更新者ID") = CommonInfo.UserId

                        row.EndEdit()
                        dt.Rows.Add(row)

                        If dt.Rows.Count = 10000 Then
                            Using bc As New Data.SqlClient.SqlBulkCopy(db.Connection, Data.SqlClient.SqlBulkCopyOptions.KeepIdentity, db.Transaction)
                                bc.DestinationTableName = tableName
                                bc.WriteToServer(dt)
                                bc.Close()
                            End Using

                            dt.Clear()
                        End If

                    Next

                    If dt.Rows.Count > 0 Then
                        Using bc As New Data.SqlClient.SqlBulkCopy(db.Connection, Data.SqlClient.SqlBulkCopyOptions.KeepIdentity, db.Transaction)
                            bc.DestinationTableName = tableName
                            bc.WriteToServer(dt)
                            bc.Close()
                        End Using
                    End If
                Else
                    '可変以外のテーブルは更新日付、更新者IDのみ更新

                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)

                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("UPDATE ""{0}"" ", tableName))
                        .AppendLine(" SET ")
                        'REV_014↓ clearListがあり、職員テーブルの場合にNULLに更新
                        If clearList.Count > 0 AndAlso tableName.Contains("＿職員") Then
                            For Each koban In clearList
                                .AppendLine(String.Format("  {0} = NULL, ", koban))
                            Next
                        End If
                        'REV_014↑
                        .AppendLine("  更新日付 = GETDATE() ")
                        .AppendLine(" ,更新者ID = @更新者ID ")
                        .AppendLine(" WHERE ")
                        .AppendLine(" 調査年 = @調査年")
                        .AppendLine(" AND ")
                        .AppendLine(" センサス番号 = @センサス番号")
                    End With

                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                    para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                    para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

                    If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                        ret = True
                    Else
                        Throw New Exception(String.Format("{0}更新失敗", tableName))
                    End If
                End If
            Next

            ret = True
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 並び替え用ダミーデータ追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertDummyIdoTresaChosahyoTable(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim importNum = ComConst.牛トレサデータ.取込対象_項目番号(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO " & query(0).ToString)
                .AppendLine(" (")
                .AppendLine(" 調査年")
                .AppendLine(", センサス番号")
                .AppendLine(", 項目番号")
                .AppendLine(", 明細番号")
                .AppendLine(", 値")
                .AppendLine(", 更新日付")
                .AppendLine(", 更新者ID")
                .AppendLine(") ")
                .AppendLine("SELECT ")
                .AppendLine(" 調査年+100000 AS 調査年")
                .AppendLine(", センサス番号")
                .AppendLine(", 項目番号")
                .AppendLine(", 明細番号")
                .AppendLine(", 値")
                .AppendLine(", 更新日付")
                .AppendLine(", 更新者ID")
                .AppendLine(" FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND (項目番号 BETWEEN @下限項目番号 AND @上限項目番号) ")
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, strChosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@下限項目番号", SqlDbType.VarChar, importNum(0)))
            para.Add(db.CreateParameter("@上限項目番号", SqlDbType.VarChar, importNum(1)))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("ダミー調査票データ追加失敗")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 成畜並び替え用ダミーデータ追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertSeitikuDummyIdoTresaChosahyoTable(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim importNum = ComConst.牛トレサデータ.成畜取込対象_項目番号(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO " & query(0).ToString)
                .AppendLine(" (")
                .AppendLine(" 調査年")
                .AppendLine(", センサス番号")
                .AppendLine(", 項目番号")
                .AppendLine(", 明細番号")
                .AppendLine(", 値")
                .AppendLine(", 更新日付")
                .AppendLine(", 更新者ID")
                .AppendLine(") ")
                .AppendLine("SELECT ")
                .AppendLine(" 調査年+100000 AS 調査年")
                .AppendLine(", センサス番号")
                .AppendLine(", 項目番号")
                .AppendLine(", 明細番号")
                .AppendLine(", 値")
                .AppendLine(", 更新日付")
                .AppendLine(", 更新者ID")
                .AppendLine(" FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND (項目番号 BETWEEN @下限項目番号 AND @上限項目番号) ")
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, strChosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@下限項目番号", SqlDbType.VarChar, importNum(0)))
            para.Add(db.CreateParameter("@上限項目番号", SqlDbType.VarChar, importNum(1)))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("ダミー調査票データ追加失敗")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' データ削除(牛資産異動情報の並び替え時使用)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteChosahyo(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim importNum = ComConst.牛トレサデータ.取込対象_項目番号(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" (調査年 = @調査年) ")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND (項目番号 BETWEEN @下限項目番号 AND @上限項目番号) ")
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, strChosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@下限項目番号", SqlDbType.VarChar, importNum(0)))
            para.Add(db.CreateParameter("@上限項目番号", SqlDbType.VarChar, importNum(1)))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("調査票データ削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 成畜データ削除(牛資産異動情報の並び替え時使用
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <returns></returns>
    Public Shared Function SeitikuDeleteChosahyo(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim importNum = ComConst.牛トレサデータ.成畜取込対象_項目番号(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" (調査年 = @調査年) ")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND (項目番号 BETWEEN @下限項目番号 AND @上限項目番号) ")
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, strChosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@下限項目番号", SqlDbType.VarChar, importNum(0)))
            para.Add(db.CreateParameter("@上限項目番号", SqlDbType.VarChar, importNum(1)))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("調査票データ削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 牛資産異動情報データ(ソート)取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <param name="strKessan"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSortDummyChosahyo(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String, strKessan As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim targetNum = ComConst.牛トレサデータ.対象_項目番号(chosaKubun)
            Dim importNum = ComConst.牛トレサデータ.取込対象_項目番号(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * FROM ( ")
                .AppendLine(" SELECT ")
                .AppendLine(" 調査年")
                .AppendLine(" ,センサス番号")
                .AppendLine(" ,明細番号")
                For Each strNum In targetNum
                    .AppendLine(", MAX(CASE WHEN 項目番号 = '" & strNum & "' THEN 値 END) AS " & strNum)
                Next
                .AppendLine(",CASE WHEN MAX(CASE WHEN 項目番号 = @異動月項目番号 THEN 値 END) Is NULL ")
                If chosaKubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse chosaKubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
                    .AppendLine("THEN CASE WHEN LEN(MAX(CASE WHEN 項目番号 =   @異動年月項目番号 THEN 値 END)) = 1 ")
                    .AppendLine("THEN NULL ELSE MONTH(CAST(MAX(CASE WHEN 項目番号 =  @異動年月項目番号 THEN 値 END) - 2 as DateTime)) END")
                    .AppendLine("ELSE MAX(CASE WHEN 項目番号 = @異動月項目番号 THEN 値 END) end As 並び替え ")
                Else
                    .AppendLine("THEN NULL ")
                    .AppendLine("ELSE MAX(CASE WHEN 項目番号 = @異動月項目番号 THEN 値 END) end As 並び替え ")
                End If
                .AppendLine(" FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND (項目番号 BETWEEN @下限項目番号 AND @上限項目番号) ")
                .AppendLine(" GROUP BY ")
                .AppendLine(" 調査年")
                .AppendLine(" ,センサス番号")
                .AppendLine(" ,明細番号")
                .AppendLine(" ) A ")
                .AppendLine(" ORDER BY ")
                .AppendLine(" IIF(A.並び替え <= @決算月")
                .AppendLine(" , A.並び替え +12+1")
                .AppendLine(" , A.並び替え + 1")
                .AppendLine(")")
                If chosaKubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse chosaKubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
                    .AppendLine(",Q11020801")
                End If
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, (CInt(strChosaNen) + 100000).ToString))
            para.Add(db.CreateParameter("@異動月項目番号", SqlDbType.VarChar, ComConst.牛トレサデータ.異動月_項目番号(chosaKubun)))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@下限項目番号", SqlDbType.VarChar, importNum(0)))
            para.Add(db.CreateParameter("@上限項目番号", SqlDbType.VarChar, importNum(1)))
            para.Add(db.CreateParameter("@決算月", SqlDbType.Int, CInt(strKessan)))
            If chosaKubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse chosaKubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
                para.Add(db.CreateParameter("@異動年月項目番号", SqlDbType.VarChar, ComConst.牛トレサデータ.異動年月_項目番号(chosaKubun)))
            End If
            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 成畜牛資産異動情報データ(ソート)取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <param name="strKessan"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSeitikuSortDummyChosahyo(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String, strKessan As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim targetNum = ComConst.牛トレサデータ.対象_項目番号(chosaKubun)
            Dim importNum = ComConst.牛トレサデータ.成畜取込対象_項目番号(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * FROM ( ")
                .AppendLine(" SELECT ")
                .AppendLine(" 調査年")
                .AppendLine(" ,センサス番号")
                .AppendLine(" ,明細番号")
                For Each strNum In targetNum
                    .AppendLine(", MAX(CASE WHEN 項目番号 = '" & strNum & "' THEN 値 END) AS " & strNum)
                Next
                .AppendLine(",CASE WHEN MAX(CASE WHEN 項目番号 = @異動月項目番号 THEN 値 END) Is NULL ")
                If chosaKubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse chosaKubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
                    .AppendLine("THEN CASE WHEN LEN(MAX(CASE WHEN 項目番号 =   @異動年月項目番号 THEN 値 END)) = 1 ")
                    .AppendLine("THEN NULL ELSE MONTH(CAST(MAX(CASE WHEN 項目番号 =  @異動年月項目番号 THEN 値 END) - 2 as DateTime)) END")
                    .AppendLine("ELSE MAX(CASE WHEN 項目番号 = @異動月項目番号 THEN 値 END) end As 並び替え ")
                Else
                    .AppendLine("THEN NULL ")
                    .AppendLine("ELSE MAX(CASE WHEN 項目番号 = @異動月項目番号 THEN 値 END) end As 並び替え ")
                End If
                .AppendLine(" FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND (項目番号 BETWEEN @下限項目番号 AND @上限項目番号) ")
                .AppendLine(" GROUP BY ")
                .AppendLine(" 調査年")
                .AppendLine(" ,センサス番号")
                .AppendLine(" ,明細番号")
                .AppendLine(" ) A ")
                .AppendLine(" ORDER BY ")
                .AppendLine(" IIF(A.並び替え <= @決算月")
                .AppendLine(" , A.並び替え +12+1")
                .AppendLine(" , A.並び替え + 1")
                .AppendLine(")")
                If chosaKubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse chosaKubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
                    .AppendLine(",Q11020801")
                End If
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, (CInt(strChosaNen) + 100000).ToString))
            para.Add(db.CreateParameter("@異動月項目番号", SqlDbType.VarChar, ComConst.牛トレサデータ.異動月_項目番号(chosaKubun)))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@下限項目番号", SqlDbType.VarChar, importNum(0)))
            para.Add(db.CreateParameter("@上限項目番号", SqlDbType.VarChar, importNum(1)))
            para.Add(db.CreateParameter("@決算月", SqlDbType.Int, CInt(strKessan)))
            If chosaKubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse chosaKubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
                para.Add(db.CreateParameter("@異動年月項目番号", SqlDbType.VarChar, ComConst.牛トレサデータ.異動年月_項目番号(chosaKubun)))
            End If
            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' ソート後のデータを登録するための基データ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <param name="strMeisaiNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSortTaget(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String, strMeisaiNo As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim importNum = ComConst.牛トレサデータ.取込対象_項目番号(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年")
                .AppendLine(" AND (項目番号 BETWEEN @下限項目番号 AND @上限項目番号) ")
                .AppendLine(" AND センサス番号 = @センサス番号")
                .AppendLine(" AND 明細番号 = @明細番号")
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, (CInt(strChosaNen) + 100000).ToString))
            para.Add(db.CreateParameter("@明細番号", SqlDbType.VarChar, strMeisaiNo))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@下限項目番号", SqlDbType.VarChar, importNum(0)))
            para.Add(db.CreateParameter("@上限項目番号", SqlDbType.VarChar, importNum(1)))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' ソート後の成畜データを登録するための基データ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <param name="strMeisaiNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSeitikuSortTaget(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String, strMeisaiNo As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim importNum = ComConst.牛トレサデータ.取込対象_項目番号(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年")
                .AppendLine(" AND (項目番号 BETWEEN @下限項目番号 AND @上限項目番号) ")
                .AppendLine(" AND センサス番号 = @センサス番号")
                .AppendLine(" AND 明細番号 = @明細番号")
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, (CInt(strChosaNen) + 100000).ToString))
            para.Add(db.CreateParameter("@明細番号", SqlDbType.VarChar, strMeisaiNo))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@下限項目番号", SqlDbType.VarChar, importNum(0)))
            para.Add(db.CreateParameter("@上限項目番号", SqlDbType.VarChar, importNum(1)))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' ソートデータの追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strMeisaiNo"></param>
    ''' <param name="dtSort"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertSortChosahyoTable(db As DBAccess, chosaKubun As String, strChosaNen As String, strMeisaiNo As String, dtSort As DataTable) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim tableName = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim dt As DataTable

            sb = New System.Text.StringBuilder
            sb.AppendLine(String.Format("SELECT TOP 0 * FROM {0} ", tableName(0).ToString))

            dt = db.GetDataTable(sb.ToString)

            Dim row As DataRow = Nothing
            For Each dr As DataRow In dtSort.Rows
                row = dt.NewRow()
                row.BeginEdit()

                row("調査年") = strChosaNen
                row("センサス番号") = dr.Item("センサス番号")
                row("項目番号") = dr.Item("項目番号")
                row("明細番号") = strMeisaiNo
                row("値") = dr.Item("値")
                row("更新日付") = dr.Item("更新日付")
                row("更新者ID") = dr.Item("更新者ID")

                row.EndEdit()
                dt.Rows.Add(row)

                If dt.Rows.Count = 10000 Then
                    Using bc As New Data.SqlClient.SqlBulkCopy(db.Connection, Data.SqlClient.SqlBulkCopyOptions.KeepIdentity, db.Transaction)
                        bc.DestinationTableName = tableName(0).ToString
                        bc.WriteToServer(dt)
                        bc.Close()
                    End Using

                    dt.Clear()
                End If
            Next

            If dt.Rows.Count > 0 Then
                Using bc As New Data.SqlClient.SqlBulkCopy(db.Connection, Data.SqlClient.SqlBulkCopyOptions.KeepIdentity, db.Transaction)
                    bc.DestinationTableName = tableName(0).ToString
                    bc.WriteToServer(dt)
                    bc.Close()
                End Using
            End If

            ret = True

        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' ソート成畜データの追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strMeisaiNo"></param>
    ''' <param name="dtSort"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertSeitikuSortChosahyoTable(db As DBAccess, chosaKubun As String, strChosaNen As String, dtSort As DataTable) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim tableName = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim dt As DataTable

            sb = New System.Text.StringBuilder
            sb.AppendLine(String.Format("SELECT TOP 0 * FROM {0} ", tableName(0).ToString))

            dt = db.GetDataTable(sb.ToString)

            Dim row As DataRow = Nothing
            For Each dr As DataRow In dtSort.Rows
                row = dt.NewRow()
                row.BeginEdit()

                row("調査年") = strChosaNen
                row("センサス番号") = dr.Item("センサス番号")
                row("項目番号") = dr.Item("項目番号")
                row("明細番号") = dr.Item("明細番号")
                row("値") = dr.Item("値")
                row("更新日付") = dr.Item("更新日付")
                row("更新者ID") = dr.Item("更新者ID")

                row.EndEdit()
                dt.Rows.Add(row)

                If dt.Rows.Count = 10000 Then
                    Using bc As New Data.SqlClient.SqlBulkCopy(db.Connection, Data.SqlClient.SqlBulkCopyOptions.KeepIdentity, db.Transaction)
                        bc.DestinationTableName = tableName(0).ToString
                        bc.WriteToServer(dt)
                        bc.Close()
                    End Using

                    dt.Clear()
                End If
            Next

            If dt.Rows.Count > 0 Then
                Using bc As New Data.SqlClient.SqlBulkCopy(db.Connection, Data.SqlClient.SqlBulkCopyOptions.KeepIdentity, db.Transaction)
                    bc.DestinationTableName = tableName(0).ToString
                    bc.WriteToServer(dt)
                    bc.Close()
                End Using
            End If

            ret = True

        Catch ex As Exception
            Throw
        End Try
        Return ret
    End Function

    Public Shared Function GetPrePrintTargetData(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String, strFromKikan As String, strToKikan As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim strNoukaCodeNum = ComConst.調査票.項目番号.牛農家団体コード(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT ")
                .AppendLine(" 異動情報.農家団体コード ")
                .AppendLine(" ,異動情報.個体識別番号 ")
                .AppendLine(" ,異動情報.異動年月日 ")
                .AppendLine(" ,異動情報.異動フラグ ")
                .AppendLine(" ,個体情報.牛の識別CD ")
                .AppendLine(" ,個体情報.性別コード ")
                .AppendLine(" ,個体情報.生年月日 ")
                .AppendLine(" ,個体情報.母牛個体識別番号 ")
                .AppendLine(" FROM 牛トレサ異動情報 AS 異動情報 ")
                .AppendLine(" LEFT JOIN ( ")
                .AppendLine("   SELECT DISTINCT ")
                .AppendLine("       個体情報.個体識別番号")
                .AppendLine("       ,牛の識別CD")
                .AppendLine("       ,性別コード")
                .AppendLine("       ,生年月日")
                .AppendLine("       ,母牛個体識別番号")
                .AppendLine("   FROM 牛トレサ個体情報 As 個体情報 ")
                .AppendLine("   INNER JOIN ( ")
                .AppendLine("       SELECT ")
                .AppendLine("           個体識別番号, MAX(照会日) As 照会日 ")
                .AppendLine("       FROM 牛トレサ個体情報 ")
                .AppendLine("       GROUP BY 個体識別番号 ")
                .AppendLine("   ) AS 個体情報＿最新照会日 ")
                .AppendLine("   ON 個体情報.個体識別番号 = 個体情報＿最新照会日.個体識別番号 ")
                .AppendLine("   AND 個体情報.照会日 = 個体情報＿最新照会日.照会日 ")
                .AppendLine(" ) AS 個体情報 ON")
                .AppendLine(" 異動情報.個体識別番号 = 個体情報.個体識別番号 ")
                .AppendLine(" WHERE 異動情報.農家団体コード IN ( ")
                .AppendLine(" SELECT 値 FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年 ")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND 項目番号 = @項目番号 ")
                .AppendLine(" ) ")
                .AppendLine(" AND (異動情報.異動年月日 BETWEEN @開始年月日 AND @終了年月日) ")
                .AppendLine(" ORDER BY 異動情報.異動年月日 ")
            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, strChosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, strNoukaCodeNum))
            para.Add(db.CreateParameter("@開始年月日", SqlDbType.Int, CInt(strFromKikan)))
            para.Add(db.CreateParameter("@終了年月日", SqlDbType.Int, CInt(strToKikan)))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function


    ''' <summary>
    ''' プレプリント出力対象牛トレサ情報の取得（成畜）''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <param name="strFromKikan"></param>
    ''' <param name="strToKikan"></param>
    ''' <returns></returns>
    Public Shared Function GetSeitikuPrePrintTargetData(db As DBAccess, chosaKubun As String, strChosaNen As String, strCensusNo As String, strFromKikan As String, strFromKikan_y As String, strFromKikan_m As String, strToKikan As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        If (chosaKubun = ComConst.調査区分.経営分析調査_牛乳生産費) Then
            Dim KessanNum = ComConst.牛トレサデータ.決算月_項目番号(CommonInfo.Chosakubun)
            Dim kessan As String
            kessan = DAOChosahyo.GettoresaKessan(db, strChosaNen, strCensusNo, CommonInfo.Chosakubun, KessanNum)
            '経営分析：調査票から取得する決算月（「【１】経営概要」シートの「４ 決算月」）の12か月前の1日
            strFromKikan_y = CStr(CInt(strChosaNen) - 1)
            strFromKikan_m = kessan
        Else
            '個別：画面.調査年の1月1日　から年月を取得
            strFromKikan_y = strChosaNen
            strFromKikan_m = "01"
        End If

        Try
            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
            Dim strNoukaCodeNum = ComConst.調査票.項目番号.牛農家団体コード(chosaKubun)

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT ")
                .AppendLine(" 異動情報.農家団体コード ")
                .AppendLine(" ,異動情報.個体識別番号 ")
                .AppendLine(" ,異動情報.異動年月日 ")
                .AppendLine(" ,異動情報.異動フラグ ")
                .AppendLine(" ,個体情報.牛の識別CD ")
                .AppendLine(" ,個体情報.性別コード ")
                .AppendLine(" ,個体情報.生年月日 ")
                .AppendLine(" ,個体情報.母牛個体識別番号 ")
                .AppendLine(" FROM 牛トレサ異動情報 AS 異動情報 ")
                .AppendLine(" LEFT JOIN ( ")
                .AppendLine("   SELECT DISTINCT ")
                .AppendLine("       個体情報.個体識別番号")
                .AppendLine("       ,牛の識別CD")
                .AppendLine("       ,性別コード")
                .AppendLine("       ,生年月日")
                .AppendLine("       ,母牛個体識別番号")
                .AppendLine("   FROM 牛トレサ個体情報 As 個体情報 ")
                .AppendLine("   INNER JOIN ( ")
                .AppendLine("       SELECT ")
                .AppendLine("           個体識別番号, MAX(照会日) As 照会日 ")
                .AppendLine("       FROM 牛トレサ個体情報 ")
                .AppendLine("       GROUP BY 個体識別番号 ")
                .AppendLine("   ) AS 個体情報＿最新照会日 ")
                .AppendLine("   ON 個体情報.個体識別番号 = 個体情報＿最新照会日.個体識別番号 ")
                .AppendLine("   AND 個体情報.照会日 = 個体情報＿最新照会日.照会日 ")
                .AppendLine(" ) AS 個体情報 ON")
                .AppendLine(" 異動情報.個体識別番号 = 個体情報.個体識別番号 ")
                .AppendLine(" WHERE 異動情報.農家団体コード IN ( ")
                .AppendLine(" SELECT 値 FROM " & query(0).ToString)
                .AppendLine(" WHERE ")
                .AppendLine(" 調査年 = @調査年 ")
                .AppendLine(" AND センサス番号 = @センサス番号 ")
                .AppendLine(" AND 項目番号 = @項目番号 ")
                .AppendLine(" ) ")
                '条件3-1（start）
                .AppendLine("  AND ( ")
                .AppendLine("    ( 牛の識別CD = 1  AND 性別コード = 2 ) ")
                .AppendLine(" OR ( 牛の識別CD = 2  AND 性別コード = 2 ) ")
                .AppendLine(" OR ( 牛の識別CD = 12 AND 性別コード = 2 ) ")
                .AppendLine(" ) ")
                '条件3-1（end）
                '条件3-2（start）
                '.AppendLine("  AND (( @開始年 - ( LEFT( 生年月日 ,4 ))) * 12 + ( @開始月 -(left((right(生年月日,4)),2))) < 30  ")
                '.AppendLine(" ) ")
                '条件3-2（end）
                '条件3-3（start）
                '.AppendLine(" AND @開始年月日 > (select　生年月日　from 牛トレサ個体情報　where　牛トレサ個体情報.個体識別番号 = 牛トレサ個体情報.母牛個体識別番号) ")
                '.AppendLine(" AND @開始年月日 < 生年月日 ")
                '条件3-3（end）
                .AppendLine(" AND 異動情報.異動フラグ = '2' ")
                .AppendLine(" AND (異動情報.異動年月日 BETWEEN @開始年月日 AND @終了年月日) ")

            End With
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, strChosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, strNoukaCodeNum))
            para.Add(db.CreateParameter("@開始年月日", SqlDbType.Int, CInt(strFromKikan)))
            para.Add(db.CreateParameter("@開始年", SqlDbType.Int, CInt(strFromKikan_y)))
            para.Add(db.CreateParameter("@開始月", SqlDbType.Int, CInt(strFromKikan_m)))
            para.Add(db.CreateParameter("@終了年月日", SqlDbType.Int, CInt(strToKikan)))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function


    '--- REV.001 ADD END
    '--- REV.002 ADD START
    ''' <summary>
    ''' 差分データ比較テーブル削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pChosaNen"></param>
    ''' <param name="pKyoku"></param>
    ''' <param name="pJimusho"></param>
    ''' <param name="pKyoten"></param>
    ''' <returns></returns>
    Public Shared Function DeleteSabunTable(db As DBAccess) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   差分比較データ")
            End With

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("差分比較テーブル削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    Public Shared Function InsertSabunTable(db As DBAccess, pSabunList As List(Of BRA7510F.差分比較項目)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each dc As BRA7510F.差分比較項目 In pSabunList
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("INSERT INTO 差分比較データ ")
                    .AppendLine("( ")
                    .AppendLine("   調査年 ")
                    .AppendLine("  ,調査区分 ")
                    .AppendLine("  ,農政局 ")
                    .AppendLine("  ,都道府県 ")
                    .AppendLine("  ,実査設置拠点 ")
                    .AppendLine("  ,センサス番号 ")
                    .AppendLine("  ,データ識別区分 ")
                    .AppendLine("  ,項目番号 ")
                    .AppendLine("  ,明細番号 ")
                    .AppendLine("  ,前回送受信データ ")
                    .AppendLine("  ,最新データ ")
                    .AppendLine("  ,更新日付 ")
                    .AppendLine("  ,更新者ID ")
                    .AppendLine(") ")
                    .AppendLine("VALUES ")
                    .AppendLine("( ")
                    .AppendLine("   @調査年 ")
                    .AppendLine("  ,@調査区分 ")
                    .AppendLine("  ,@農政局 ")
                    .AppendLine("  ,@都道府県 ")
                    .AppendLine("  ,@実査設置拠点 ")
                    .AppendLine("  ,@センサス番号 ")
                    .AppendLine("  ,@データ識別区分 ")
                    .AppendLine("  ,@項目番号 ")
                    .AppendLine("  ,@明細番号 ")
                    .AppendLine("  ,@前回送受信データ ")
                    .AppendLine("  ,@最新データ ")
                    .AppendLine("  ,GETDATE() ")
                    .AppendLine("  ,@更新者ID ")
                    .AppendLine(") ")
                End With

                para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, dc.調査年))
                para.Add(db.CreateParameter("@調査区分", SqlDbType.VarChar, dc.調査区分))
                para.Add(db.CreateParameter("@農政局", SqlDbType.VarChar, dc.農政局))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.VarChar, dc.都道府県))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.VarChar, dc.実査設置拠点))
                para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, dc.センサス番号))
                para.Add(db.CreateParameter("@データ識別区分", SqlDbType.VarChar, dc.データ識別区分))
                para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, dc.項目番号))
                para.Add(db.CreateParameter("@明細番号", SqlDbType.VarChar, dc.明細番号))
                para.Add(db.CreateParameter("@前回送受信データ", SqlDbType.VarChar, dc.前回送受信データ))
                para.Add(db.CreateParameter("@最新データ", SqlDbType.VarChar, dc.最新データ))
                para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

                If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                    ret = True
                Else
                    Throw New Exception("差分データ追加失敗")
                End If
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 差分データ比較一覧取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pChosaNen"></param>
    ''' <param name="pKyoku"></param>
    ''' <param name="pJimusho"></param>
    ''' <param name="pKyoten"></param>
    ''' <returns></returns>
    Public Shared Function GetSabunIchiran(db As DBAccess, pChosaNen As String, pKyoku As String, pJimusho As String, pKyoten As String, pChosaKoutei As String, upLow As String, pChosakubun As String, Optional Jushin As Boolean = False) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT DISTINCT SA.調査年 ")
                .AppendLine(", SA.調査区分 ")
                .AppendLine(", SA.センサス番号 ")
                .AppendLine(", SO.送信日時 ")
                .AppendLine(", SO.受信日時 ")
                .AppendLine(", CH.更新日付 As 調査票更新日時")
                .AppendLine(", KB.更新日付 As 個別結果表更新日時")
                .AppendLine("FROM   差分比較データ As SA")
                .AppendLine(String.Format("LEFT JOIN [{0}].[dbo].[送受信管理] As SO", pChosaKoutei))
                .AppendLine("ON SA.調査年 = SO.調査年 ")
                .AppendLine("AND SA.センサス番号 = SO.センサス番号 ")
                .AppendLine("AND SA.調査区分 = SO.調査区分 ")
                .AppendLine("AND   ( SO.上り下り区分  = @上り下り区分 ")
                .AppendLine("OR    SO.上り下り区分  IS NULL )")
                .AppendLine("AND   ( SO.送受信データ種別  = @送受信データ種別 ")
                .AppendLine("OR    SO.送受信データ種別  IS NULL )")
                .AppendLine(String.Format("LEFT JOIN {0} As CH", ComConst.調査票.テーブル名称(pChosakubun)(0)))
                .AppendLine("ON SA.調査年 = CH.調査年 ")
                .AppendLine("AND SA.センサス番号 = CH.センサス番号 ")
                .AppendLine(String.Format("LEFT JOIN {0} As KB", ComConst.個別結果表.テーブル名称(pChosakubun)(0)))
                .AppendLine("ON SA.調査年 = KB.調査年 ")
                .AppendLine("AND SA.センサス番号 = KB.センサス番号 ")
                .AppendLine("WHERE SA.調査年  = @調査年 ")
                .AppendLine("AND    SA.調査区分  = @調査区分 ")
                If Not pKyoku Is Nothing Then
                    .AppendLine("AND     SA.農政局           = @農政局 ")
                End If
                If Not pJimusho Is Nothing Then
                    .AppendLine("AND     SA.都道府県         = @都道府県 ")
                End If
                If Not pKyoten Is Nothing Then
                    .AppendLine("AND     SA.実査設置拠点     = @実査設置拠点 ")
                End If
                .AppendLine("ORDER BY センサス番号 ASC ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pChosaNen))
            para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, upLow))
            para.Add(db.CreateParameter("@送受信データ種別", SqlDbType.Int, "1"))
            If Not pKyoku Is Nothing Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, pKyoku))
            End If
            If Not pJimusho Is Nothing Then
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, pJimusho))
            End If
            If Not pKyoten Is Nothing Then
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, pKyoten))
            End If

            ret = db.GetDataTable(sb.ToString, para)
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 差分データ登録確認
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pChosaNen"></param>
    ''' <param name="pKyoku"></param>
    ''' <param name="pJimusho"></param>
    ''' <param name="pKyoten"></param>
    ''' <returns></returns>
    Public Shared Function sabunKensuCheck(db As DBAccess, pChosaNen As String, pKyoku As String, pJimusho As String, pKyoten As String) As Boolean
        Dim ret As Boolean = False
        Dim table As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *  ")
                .AppendLine("FROM   差分比較データ ")
                .AppendLine("WHERE 調査年  = @調査年 ")
                .AppendLine("AND 調査区分  = @調査区分 ")
                If Not pKyoku Is Nothing Then
                    .AppendLine("AND     農政局           = @農政局 ")
                End If
                If Not pJimusho Is Nothing Then
                    .AppendLine("AND     都道府県         = @都道府県 ")
                End If
                If Not pKyoten Is Nothing Then
                    .AppendLine("AND     実査設置拠点     = @実査設置拠点 ")
                End If
                .AppendLine("ORDER BY センサス番号 DESC ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pChosaNen))
            If Not pKyoku Is Nothing Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, pKyoku))
            End If
            If Not pJimusho Is Nothing Then
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, pJimusho))
            End If
            If Not pKyoten Is Nothing Then
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, pKyoten))
            End If

            table = db.GetDataTable(sb.ToString, para)

            If table.Rows.Count = 0 Then
                ret = False
            Else
                ret = True
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 差分データテーブル取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pChosaNen"></param>
    ''' <param name="pKyoku"></param>
    ''' <param name="pJimusho"></param>
    ''' <param name="pKyoten"></param>
    ''' <returns></returns>
    Public Shared Function getSabunDataTable(db As DBAccess, pChosaNen As String, pKyoku As String, pJimusho As String, pKyoten As String) As DataTable
        Dim ret As DataTable

        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *  ")
                .AppendLine("FROM   差分比較データ ")
                .AppendLine("WHERE 調査年  = @調査年 ")
                .AppendLine("AND 調査区分  = @調査区分 ")
                If Not pKyoku Is Nothing Then
                    .AppendLine("AND     農政局           = @農政局 ")
                End If
                If Not pJimusho Is Nothing Then
                    .AppendLine("AND     都道府県         = @都道府県 ")
                End If
                If Not pKyoten Is Nothing Then
                    .AppendLine("AND     実査設置拠点     = @実査設置拠点 ")
                End If
                .AppendLine("ORDER BY センサス番号 ASC ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pChosaNen))
            If Not pKyoku Is Nothing Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, pKyoku))
            End If
            If Not pJimusho Is Nothing Then
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, pJimusho))
            End If
            If Not pKyoten Is Nothing Then
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, pKyoten))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function
    '--- REV.002 ADD END


    ''' <summary>
    ''' 調査票審査論理範囲取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoShinsaRonriRange(db As DBAccess) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM   調査票審査論理＿範囲")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    農政局         = @農政局 ")
                .AppendLine("AND    都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                .AppendLine("ORDER BY 連番")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function



    ''' <summary>
    ''' 調査票審査論理範囲削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteChosahyoShinsaRonriRange(db As DBAccess) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine("FROM   調査票審査論理＿範囲")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    農政局         = @農政局 ")
                .AppendLine("AND    都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("調査票審査論理＿範囲削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function


    ''' <summary>
    ''' 調査票審査論理範囲追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertChosahyoShinsaRonriRange(db As DBAccess, dc As Dictionary(Of String, String)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO 調査票審査論理＿範囲 ")
                .AppendLine("( ")
                .AppendLine("   調査区分 ")
                .AppendLine("  ,農政局 ")
                .AppendLine("  ,都道府県 ")
                .AppendLine("  ,実査設置拠点 ")
                .AppendLine("  ,連番 ")
                .AppendLine("  ,チェック項目名 ")
                .AppendLine("  ,項目番号１ ")
                .AppendLine("  ,項目番号２ ")
                .AppendLine("  ,値 ")
                .AppendLine("  ,上限 ")
                .AppendLine("  ,下限 ")
                .AppendLine("  ,繰り返し ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @調査区分 ")
                .AppendLine("  ,@農政局 ")
                .AppendLine("  ,@都道府県 ")
                .AppendLine("  ,@実査設置拠点 ")
                .AppendLine("  ,@連番 ")
                .AppendLine("  ,@チェック項目名 ")
                .AppendLine("  ,@項目番号１ ")
                .AppendLine("  ,@項目番号２ ")
                .AppendLine("  ,@値 ")
                .AppendLine("  ,@上限 ")
                .AppendLine("  ,@下限 ")
                .AppendLine("  ,@繰り返し ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            para.Add(db.CreateParameter("@連番", SqlDbType.Int, dc("連番")))
            para.Add(db.CreateParameter("@チェック項目名", SqlDbType.VarChar, dc("チェック項目名")))
            para.Add(db.CreateParameter("@項目番号１", SqlDbType.VarChar, dc("項目番号１")))
            para.Add(db.CreateParameter("@項目番号２", SqlDbType.VarChar, dc("項目番号２")))
            para.Add(db.CreateParameter("@値", SqlDbType.Decimal, dc("値")))
            para.Add(db.CreateParameter("@上限", SqlDbType.Decimal, dc("上限")))
            para.Add(db.CreateParameter("@下限", SqlDbType.Decimal, dc("下限")))
            para.Add(db.CreateParameter("@繰り返し", SqlDbType.VarChar, dc("繰り返し")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("調査票審査論理＿範囲追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票審査論理範囲更新日時取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoShinsaRonriRangeUpdateDate(db As DBAccess) As String

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine("FROM   調査票審査論理＿範囲")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND    農政局         = @農政局 ")
                .AppendLine("AND    都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, CommonInfo.Jimusyo))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn
    End Function

    ''' <summary>
    ''' 労働時間整理ファイル項目クラス
    ''' </summary>
    ''' REV_005 ADD
    ''' <remarks></remarks>
    Public Class 労働時間整理ファイル項目
        Public シート名 As String                  'シート名
        Public 行位置 As Integer                   '行位置
        Public 列位置 As Integer                   '列位置
        Public 値 As String                        '値
        Public 型区分 As String                    '型区分
        Public 有効桁数 As Integer                 '有効桁数
        Public 小数点以下桁数 As Integer           '小数点以下桁数
        Public 可変範囲 As String                  '可変範囲
    End Class

    ''' <summary>
    ''' 労働時間整理ファイル拠点クラス
    ''' </summary>
    ''' REV_005 ADD
    ''' <remarks></remarks>
    Public Class RoudouKyotenKey
        ''' <summary>農政局</summary>
        Public Property kyoku As String
        ''' <summary>都道府県</summary>
        Public Property jimusho As String
        ''' <summary>実査設置拠点</summary>
        Public Property kyoten As String

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Sub New()

        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="kyoku"></param>
        ''' <param name="jimusho"></param>
        ''' <param name="kyoten"></param>
        ''' <remarks></remarks>
        Sub New(kyoku As String, jimusho As String, kyoten As String)
            _kyoku = kyoku
            _jimusho = jimusho
            _kyoten = kyoten
        End Sub

    End Class

#Region "要件No.3"
    ''' <summary>
    ''' 制度受取金・積立金等項目取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="errType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSeidouketorikin(db As DBAccess) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim chosanenselect As String

        Try
            If ComUtil.Seidouketorikin.Chosanen = "1" Then
                chosanenselect = "2022"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "2" Then
                chosanenselect = "2023"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "3" Then
                chosanenselect = "2024"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "4" Then
                chosanenselect = "2025"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "5" Then
                chosanenselect = "2026"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "6" Then
                chosanenselect = "2027"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "7" Then
                chosanenselect = "2028"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "8" Then
                chosanenselect = "2029"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "9" Then
                chosanenselect = "2030"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "10" Then
                chosanenselect = "2031"
            End If

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine("FROM  制度受取金・積立金等項目 ")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND  調査年       = @調査年 ")
                .AppendLine("ORDER BY 項番")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanenselect))


            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 制度受取金・積立金等項目追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertSeidouketorikin(db As DBAccess, dc As Dictionary(Of String, String)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim chosanenselect As String



        Try
            If ComUtil.Seidouketorikin.Chosanen = "1" Then
                chosanenselect = "2022"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "2" Then
                chosanenselect = "2023"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "3" Then
                chosanenselect = "2024"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "4" Then
                chosanenselect = "2025"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "5" Then
                chosanenselect = "2026"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "6" Then
                chosanenselect = "2027"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "7" Then
                chosanenselect = "2028"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "8" Then
                chosanenselect = "2029"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "9" Then
                chosanenselect = "2030"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "10" Then
                chosanenselect = "2031"
            End If
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine(String.Format("INSERT INTO 制度受取金・積立金等項目"))
                .AppendLine("( ")
                .AppendLine("   調査区分 ")
                .AppendLine("  ,調査年 ")
                .AppendLine("  ,項番 ")
                .AppendLine("  ,出力項目名 ")
                .AppendLine("  ,制度受取金等項番 ")
                .AppendLine("  ,制度積立金等項番 ")
                .AppendLine("  ,更新日付 ")
                .AppendLine("  ,更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine("   @調査区分 ")
                .AppendLine("  ,@調査年 ")
                .AppendLine("  ,@項番 ")
                .AppendLine("  ,@出力項目名 ")
                .AppendLine("  ,@制度受取金等項番 ")
                .AppendLine("  ,@制度積立金等項番 ")
                .AppendLine("  ,GETDATE() ")
                .AppendLine("  ,@更新者ID ")
                .AppendLine(") ")
            End With


            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanenselect))
            para.Add(db.CreateParameter("@項番", SqlDbType.VarChar, dc("項番")))
            para.Add(db.CreateParameter("@出力項目名", SqlDbType.VarChar, dc("出力項目名")))
            para.Add(db.CreateParameter("@制度受取金等項番", SqlDbType.VarChar, dc("制度受取金等項番")))
            para.Add(db.CreateParameter("@制度積立金等項番", SqlDbType.VarChar, dc("制度積立金等項番")))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("制度受取金・積立金等項目追加失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 制度受取金・積立金等項目調査票マスタ反映
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertOtherSeidouketorikin(db As DBAccess, dc As Dictionary(Of String, String), Kouban As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing


        Try
            Dim query = From val In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun) Where Not val.Contains("＿可変")
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb

                .AppendLine("UPDATE " & query(0).ToString)
                .AppendLine(" SET ")
                .AppendLine(Kouban & " = @出力項目名 ")



            End With
            para.Add(db.CreateParameter("@出力項目名", SqlDbType.VarChar, dc("出力項目名")))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("制度受取金・積立金等項目更新失敗")
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 制度受取金等項番チェック
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Seidokincheck(db As DBAccess, chosaKubun As String, ChosaNen As String, Seidokinkouban As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * from 制度受取金・積立金等項目")
                .AppendLine(" WHERE ")
                .AppendLine(" 調査区分 = @調査区分")
                .AppendLine(" AND 調査年 = @調査年 ")
                .AppendLine(" AND 制度受取金等項番 = @制度受取金等項番 ")
            End With
            para.Add(db.CreateParameter("@調査区分", SqlDbType.VarChar, chosaKubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, ChosaNen))
            para.Add(db.CreateParameter("@制度受取金等項番", SqlDbType.VarChar, Seidokinkouban))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 制度積立金等項番チェック
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="strChosaNen"></param>
    ''' <param name="strCensusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Seidokincheck2(db As DBAccess, chosaKubun As String, ChosaNen As String, Seidokinkouban As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * from 制度受取金・積立金等項目")
                .AppendLine(" WHERE ")
                .AppendLine(" 調査区分 = @調査区分")
                .AppendLine(" AND 調査年 = @調査年 ")
                .AppendLine(" AND 制度積立金等項番 = @制度積立金等項番 ")
            End With
            para.Add(db.CreateParameter("@調査区分", SqlDbType.VarChar, chosaKubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, ChosaNen))
            para.Add(db.CreateParameter("@制度積立金等項番", SqlDbType.VarChar, Seidokinkouban))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 制度受取金・積立金等項目削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteSeidouketorikin(db As DBAccess) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim chosanenselect As String
        Try
            If ComUtil.Seidouketorikin.Chosanen = "1" Then
                chosanenselect = "2022"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "2" Then
                chosanenselect = "2023"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "3" Then
                chosanenselect = "2024"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "4" Then
                chosanenselect = "2025"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "5" Then
                chosanenselect = "2026"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "6" Then
                chosanenselect = "2027"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "7" Then
                chosanenselect = "2028"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "8" Then
                chosanenselect = "2029"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "9" Then
                chosanenselect = "2030"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "10" Then
                chosanenselect = "2031"
            End If
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE ")
                .AppendLine(String.Format("FROM   制度受取金・積立金等項目"))
                .AppendLine("WHERE  調査区分       = @調査区分 ")
                .AppendLine("AND  調査年       = @調査年 ")

            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanenselect))


            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("制度受取金・積立金等項目削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 制度受取金・積立金等項目更新日時取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSeidouketorikinDate(db As DBAccess) As String

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT MAX(更新日付) AS 更新日時")
                .AppendLine("FROM   制度受取金・積立金等項目")
                .AppendLine("WHERE  調査区分       = @調査区分 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("更新日時").ToString
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn
    End Function

    ''' <summary>
    ''' 制度受取金・積立金等項目削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="jimusho"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteSeidouketorikinNum(db As DBAccess, address As String()) As Boolean
        Dim Seidouketorikincheck As Boolean = False
        For Each a As String In address
            If Not a Is Nothing Then
                Seidouketorikincheck = True
            End If
        Next

        If Seidouketorikincheck = False Then
            Return True
        End If



        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim chosanenselect As String
        Try
            Dim query = From val In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun) Where Not val.Contains("＿可変")

            If ComUtil.Seidouketorikin.Chosanen = "1" Then
                chosanenselect = "2022"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "2" Then
                chosanenselect = "2023"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "3" Then
                chosanenselect = "2024"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "4" Then
                chosanenselect = "2025"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "5" Then
                chosanenselect = "2026"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "6" Then
                chosanenselect = "2027"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "7" Then
                chosanenselect = "2028"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "8" Then
                chosanenselect = "2029"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "9" Then
                chosanenselect = "2030"
            ElseIf ComUtil.Seidouketorikin.Chosanen = "10" Then
                chosanenselect = "2031"
            End If
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            Dim isTop As Boolean = True
            With sb
                .AppendLine("UPDATE " & query(0).ToString & " SET")
                For i = 0 To address.Count - 1
                    If String.IsNullOrEmpty(address(i)) Then
                        Continue For
                    End If

                    If isTop Then
                        .AppendLine(address(i) & "= NULL")
                        isTop = False
                    Else
                        .AppendLine(", " & address(i) & "= NULL")
                    End If
                Next
                .AppendLine("Where  調査年 = @調査年 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanenselect))


            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("制度受取金・積立金等項目削除失敗")
            End If
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

#End Region

End Class
