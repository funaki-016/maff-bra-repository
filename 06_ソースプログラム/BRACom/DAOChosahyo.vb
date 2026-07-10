Imports System.Data.SqlClient

''' <summary>
''' 調査票テーブル操作
''' </summary>
''' <remarks></remarks>
'''
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.11.06 |TSP)                | フェーズ3 要件No.1,6, 追加要件No12修正　
'//  REV_002   | 2021.03.17 |TSP)                | 連絡票No.687 テーブル結合条件不正等対応　
'//  REV_003   | 2022.01.05 |日本ｺﾝﾋﾟｭｰﾀｼｽﾃﾑ     | 要件No.2対応　
'//  REV_004   | 2023.01.13 |DAiKO               | 変更要件No.7対応
'//  REV_005   | 2023.04.21 |DAiKO               | 変更要件No.3対応
'//  REV_006   | 2023.08.09 |DAiKO               | 要件No.11 調査票項目拡張クラス（数式区分追加）を追加
'//  REV_007   | 2023.11.27 |DAIKO               | 変更要件No16 CSVから出力を行わない項目の対応
'//  REV_008   | 2025.09.11 |GCU                 | 変更要件No.9 調査対象経営体と担当名称を取得
'//*************************************************************************************************
Public Class DAOChosahyo

    ''' <summary>
    ''' 調査票項目クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 調査票項目
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
    ''' 調査票項目拡張クラス
    ''' </summary>
    ''' <remarks>REV_006</remarks>
    Public Class 調査票項目EX : Inherits 調査票項目
        Public 数式区分 As String                  '数式区分
    End Class

    ''' <summary>
    ''' 中間集計表項目クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 中間集計表項目
        Inherits 調査票項目
        Public 乗算値 As Integer                   '中間集計表乗算値
        Public 追加項目 As String                  '中間集計表追加項目
        Public ラウンド桁数 As Integer             '中間集計表ラウンド桁数
        Public 登録削除区分 As Integer             '中間集計表登録削除区分
        Public 明細番号 As Integer                 '中間集計表明細番号
    End Class

    ''' <summary>主キークラス</summary>
    Public Class PrimaryKey
        ''' <summary>調査年</summary>
        Public Property chosaNen As String
        ''' <summary>センサス番号</summary>
        Public Property censusNo As String

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="chosaNen"></param>
        ''' <remarks></remarks>
        Sub New(chosaNen As String)
            _chosaNen = chosaNen
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="chosaNen"></param>
        ''' <param name="censusNo"></param>
        ''' <remarks></remarks>
        Sub New(chosaNen As String, censusNo As String)
            _chosaNen = chosaNen
            _censusNo = censusNo
        End Sub

    End Class

    ''' <summary>拠点クラス</summary>
    Public Class KotenKey
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

    ''' <summary>
    ''' 調査票調査年取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosaNen(db As DBAccess, kyoku As String, jimusho As String, kyoten As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT DISTINCT 調査年 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine("WHERE  1         = 1 ")
                If Not kyoku Is Nothing Then
                    .AppendLine("AND    農政局       = @農政局 ")
                End If
                If Not jimusho Is Nothing Then
                    .AppendLine("AND    都道府県       = @都道府県 ")
                End If
                If Not kyoten Is Nothing Then
                    .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                End If
                If CommonInfo.SenmonChosain Then
                    .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                    .AppendLine("                          FROM   専門調査員担当調査客体 ")
                    .AppendLine("                          WHERE  ユーザーID = @ユーザーID ")
                    .AppendLine("                         ) ")
                End If
                .AppendLine("ORDER BY 調査年 DESC")
            End With

            If Not kyoku Is Nothing Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kyoku))
            End If
            If Not jimusho Is Nothing Then
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))
            End If
            If Not kyoten Is Nothing Then
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kyoten))
            End If
            If CommonInfo.SenmonChosain Then
                para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, CommonInfo.UserId))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function
    '↑2022/02/01 調査票調査年取得追加

    ''' <summary>
    ''' センサス番号取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="censusNo"></param>
    ''' <param name="kotenKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCensusNo(db As DBAccess, chosakubun As String, chosaNen As String, kyoku As String, jimusho As String, kyoten As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            With sb
                .AppendLine("SELECT センサス番号 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine("WHERE  調査年         = @調査年 ")
                If Not kyoku Is Nothing Then
                    .AppendLine("AND    農政局         = @農政局 ")
                End If
                If Not jimusho Is Nothing Then
                    .AppendLine("AND    都道府県       = @都道府県 ")
                End If
                If Not kyoten Is Nothing Then
                    .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                End If
            End With

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
    ''' 調査票取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="censusNo"></param>
    ''' <param name="kotenKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyo(ByVal db As DBAccess,
                                       ByVal chosaNen As String,
                                       ByVal censusNo As String,
                                       Optional ByVal kotenKey As DAOChosahyo.KotenKey = Nothing) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine("WHERE  調査年         = @調査年 ")
                .AppendLine("AND    センサス番号   = @センサス番号")
                If Not String.IsNullOrWhiteSpace(kotenKey.kyoku) Then
                    .AppendLine("AND    農政局       = @農政局 ")
                End If
                If Not String.IsNullOrWhiteSpace(kotenKey.jimusho) Then
                    .AppendLine("AND    都道府県       = @都道府県 ")
                End If
                If Not String.IsNullOrWhiteSpace(kotenKey.kyoten) Then
                    .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                End If
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, censusNo))
            If Not String.IsNullOrWhiteSpace(kotenKey.kyoku) Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kotenKey.kyoku))
            End If
            If Not String.IsNullOrWhiteSpace(kotenKey.jimusho) Then
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kotenKey.jimusho))
            End If
            If Not String.IsNullOrWhiteSpace(kotenKey.kyoten) Then
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kotenKey.kyoten))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票リスト表示取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="einouRuikei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoList(db As DBAccess, chosaNen As String, kyoku As String, jimusho As String, kyoten As String, einouRuikei As String, chosaSoshiki As String, tantoMeisho As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT 農政局 ")
                .AppendLine("     , 都道府県 ")
                .AppendLine("     , 実査設置拠点 ")
                .AppendLine("     , センサス番号 ")
                .AppendLine("     , 更新日付 ")
                .AppendLine("     , Q00000701 ")
                .AppendLine("     , Q00000801 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine("WHERE  調査年         = @調査年 ")
                If Not kyoku Is Nothing Then
                    .AppendLine("AND    農政局       = @農政局 ")
                End If
                If Not jimusho Is Nothing Then
                    .AppendLine("AND    都道府県       = @都道府県 ")
                End If
                If Not kyoten Is Nothing Then
                    .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                End If
                If Not einouRuikei Is Nothing Then
                    .AppendLine(String.Format("AND    SUBSTRING({0}, 1, CHARINDEX( '：', {0} ) - 1 )   = @営農類型 ", ComConst.調査票.営農類型(CommonInfo.Chosakubun)))
                End If
                If CommonInfo.SenmonChosain Then
                    .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                    .AppendLine("                          FROM   専門調査員担当調査客体 ")
                    .AppendLine("                          WHERE  ユーザーID = @ユーザーID ")
                    .AppendLine("                         ) ")
                End If
                If chosaSoshiki <> "" Then
                    .AppendLine("AND    Q00000701   LIKE '%" + chosaSoshiki + "%' ")
                End If
                If tantoMeisho <> "" Then
                    .AppendLine("AND    Q00000801   LIKE '%" + tantoMeisho + "%' ")
                End If
            End With

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
            If Not einouRuikei Is Nothing Then
                para.Add(db.CreateParameter("@営農類型", SqlDbType.VarChar, einouRuikei))
            End If
            If CommonInfo.SenmonChosain Then
                para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, CommonInfo.UserId))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票テーブル取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoTable(db As DBAccess, pKey As DAOChosahyo.PrimaryKey) As Dictionary(Of String, DataTable)
        Dim ret As New Dictionary(Of String, DataTable)
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT *")
                    .AppendLine(String.Format("FROM   ""{0}""", tableName))
                    .AppendLine("WHERE  調査年         = @調査年 ")
                    .AppendLine("AND    センサス番号   = @センサス番号 ")
                End With

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))

                dt = db.GetDataTable(sb.ToString, para)

                ret.Add(tableName, dt)
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票テーブル取得（調査対象経営体・担当名称のみ）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoTable_KeieiTanto(db As DBAccess, pKey As DAOChosahyo.PrimaryKey) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT Q00000701 ")
                .AppendLine("     , Q00000801 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine("WHERE  調査年         = @調査年 ")
                .AppendLine("AND    センサス番号   = @センサス番号 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function


    '--- REV_001 ADD START
    ''' <summary>
    ''' 調査票テーブル取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoTable(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, dcKouban As Dictionary(Of String, String())) As Dictionary(Of String, DataTable)
        Dim ret As New Dictionary(Of String, DataTable)
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)

                If Not dcKouban.ContainsKey(tableName) Then
                    '対象項番が存在しない場合はスキップ
                    Continue For
                End If

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT ")
                    If tableName.Contains("＿可変") Then
                        .AppendLine("* ")
                    Else
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        For Each colName As String In dcKouban(tableName)
                            .AppendLine(String.Format("  ,{0}", colName))
                        Next

                    End If
                    .AppendLine(String.Format("FROM   ""{0}""", tableName))
                    .AppendLine("WHERE  調査年         = @調査年 ")
                    .AppendLine("AND    センサス番号   = @センサス番号 ")
                    If tableName.Contains("＿可変") Then
                        .AppendLine("AND    項目番号   IN ('" & String.Join("', '", dcKouban(tableName)) & "') ")
                    End If
                End With

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))

                dt = db.GetDataTable(sb.ToString, para)

                ret.Add(tableName, dt)
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function
    '--- REV_001 ADD END

    '--- REV_003 ADD START
    ''' <summary>
    ''' 調査票テーブル取得（調査区分／調査年／センサス番号指定）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="censusNo"></param>
    ''' <param name="chosakubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoTable(db As DBAccess, chosaNen As Integer, censusNo As String, chosakubun As String) As Integer
        Dim ret As New Integer
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(chosakubun)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT *")
                    .AppendLine(String.Format("FROM   ""{0}""", tableName))
                    .AppendLine("WHERE  調査年         = @調査年 ")
                    .AppendLine("AND    センサス番号   = @センサス番号 ")
                End With

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
                para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, censusNo))

                dt = db.GetDataTable(sb.ToString, para)

                ret = ret + dt.Rows.Count
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function
    '--- REV_003 ADD END

    ''' <summary>
    ''' 調査票データ追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertChosahyoTable(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, kKey As DAOChosahyo.KotenKey, dc As Dictionary(Of String, DAOChosahyo.調査票項目), chosaSoshiki As String, tantoMeisho As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing



        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)
                If Not tableName.Contains("＿可変") Then
                    '列名称取得
                    Dim colArr As List(Of String) = GetColumns(db, tableName)

                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)

                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("INSERT INTO ""{0}"" ", tableName))
                        .AppendLine("( ")
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        If tableName = ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0) Then
                            .AppendLine("  ,農政局 ")
                            .AppendLine("  ,都道府県 ")
                            .AppendLine("  ,実査設置拠点 ")
                        End If
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,{0} ", col))
                        Next
                        .AppendLine("  ,更新日付 ")
                        .AppendLine("  ,更新者ID ")
                        .AppendLine(") ")
                        .AppendLine("VALUES ")
                        .AppendLine("( ")
                        .AppendLine("   @調査年 ")
                        .AppendLine("  ,@センサス番号 ")
                        If tableName = ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0) Then
                            .AppendLine("  ,@農政局 ")
                            .AppendLine("  ,@都道府県 ")
                            .AppendLine("  ,@実査設置拠点 ")
                        End If
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,@{0} ", col))
                        Next
                        .AppendLine("  ,GETDATE() ")
                        .AppendLine("  ,@更新者ID ")
                        .AppendLine(") ")
                    End With

                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                    para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                    If tableName = ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0) Then
                        para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                        para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                        para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))
                    End If
                    For Each col As String In colArr
                        '--- REV_001 ADD Start
                        If dc.ContainsKey(col) Then
                            '--- REV_001 ADD End
                            para.Add(db.CreateParameter(String.Format("@{0}", col), If(dc(col).型区分 = ComConst.型区分.数値型, SqlDbType.Decimal, SqlDbType.VarChar), dc(col).値))
                            '--- REV_001 ADD Start
                            '>>>2022/01/27
                        ElseIf col = "Q00000701" Then
                            para.Add(db.CreateParameter("@Q00000701", SqlDbType.VarChar, chosaSoshiki))

                        ElseIf col = "Q00000801" Then
                            para.Add(db.CreateParameter("@Q00000801", SqlDbType.VarChar, tantoMeisho))
                            '<<<2022/01/27
                        Else
                            para.Add(db.CreateParameter(String.Format("@{0}", col), SqlDbType.Decimal, Nothing))
                        End If
                        '--- REV_001 ADD End
                    Next
                    para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

                    If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                        ret = True
                    Else
                        Throw New Exception(String.Format("{0}追加失敗", tableName))
                    End If
                Else
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
                        row("明細番号") = no(1)
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

                    ret = True
                End If
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票データ追加（労働時間整理ファイル用）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="dc"></param>
    ''' <param name="dr"></param>
    ''' <param name="dtPrev"></param>
    ''' <returns></returns>
    ''' <remarks>'REV003 Add</remarks>
    Public Shared Function InsertChosahyoTable_Roudou(db As DBAccess,
                                                      pKey As DAOChosahyo.PrimaryKey,
                                                      kKey As DAOOther.RoudouKyotenKey,
                                                      dc As Dictionary(Of String, DAOChosahyo.調査票項目),
                                                      dcr As Dictionary(Of String, DAOOther.労働時間整理ファイル項目),
                                                      chosakbn As String,
                                                      dtPrev As Dictionary(Of String, DataTable),
                                                      dtChosa As Dictionary(Of String, DataTable),
                                                      chosaSoshiki As String,
                                                      tantoMeisho As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(chosakbn)
                If Not tableName.Contains("＿可変") Then
                    '列名称取得
                    Dim colArr As List(Of String) = GetColumns(db, tableName)

                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)

                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("INSERT INTO ""{0}"" ", tableName))
                        .AppendLine("( ")
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        If tableName = ComConst.調査票.テーブル名称(chosakbn)(0) Then
                            .AppendLine("  ,農政局 ")
                            .AppendLine("  ,都道府県 ")
                            .AppendLine("  ,実査設置拠点 ")
                        End If
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,{0} ", col))
                        Next
                        .AppendLine("  ,更新日付 ")
                        .AppendLine("  ,更新者ID ")
                        .AppendLine(") ")
                        .AppendLine("VALUES ")
                        .AppendLine("( ")
                        .AppendLine("   @調査年 ")
                        .AppendLine("  ,@センサス番号 ")
                        If tableName = ComConst.調査票.テーブル名称(chosakbn)(0) Then
                            .AppendLine("  ,@農政局 ")
                            .AppendLine("  ,@都道府県 ")
                            .AppendLine("  ,@実査設置拠点 ")
                        End If
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,@{0} ", col))
                        Next
                        .AppendLine("  ,GETDATE() ")
                        .AppendLine("  ,@更新者ID ")
                        .AppendLine(") ")
                    End With

                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                    para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                    If tableName = ComConst.調査票.テーブル名称(chosakbn)(0) Then
                        para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                        para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                        para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))
                    End If
                    For Each col As String In colArr
                        If dc.ContainsKey(col) Then
                            If dcr.ContainsKey(col) Then
                                '労働時間整理ファイルとマッチする場合はそちらからセット
                                para.Add(db.CreateParameter(String.Format("@{0}", col), If(dcr(col).型区分 = ComConst.型区分.数値型, SqlDbType.Decimal, SqlDbType.VarChar), dcr(col).値))
                            Else
                                '↓MOD MS 2022/01/25
                                'para.Add(db.CreateParameter(String.Format("@{0}", col), If(dc(col).型区分 = ComConst.型区分.数値型, SqlDbType.Decimal, SqlDbType.VarChar), dc(col).値))
                                Dim targetDt As DataTable = dtPrev(tableName)
                                If dc(col).型区分 = ComConst.型区分.数値型 Then
                                    If targetDt.Rows(0).Item(col) Is DBNull.Value OrElse String.IsNullOrEmpty(targetDt.Rows(0).Item(col).ToString()) Then
                                        para.Add(db.CreateParameter(String.Format("@{0}", col), SqlDbType.Decimal, Nothing))
                                    Else
                                        para.Add(db.CreateParameter(String.Format("@{0}", col), SqlDbType.Decimal, Decimal.Parse(targetDt.Rows(0).Item(col).ToString())))
                                    End If
                                Else
                                    If targetDt.Rows(0).Item(col) Is DBNull.Value OrElse String.IsNullOrEmpty(targetDt.Rows(0).Item(col).ToString()) Then
                                        para.Add(db.CreateParameter(String.Format("@{0}", col), SqlDbType.VarChar, Nothing))
                                    Else
                                        para.Add(db.CreateParameter(String.Format("@{0}", col), SqlDbType.VarChar, targetDt.Rows(0).Item(col).ToString()))
                                    End If
                                End If
                                '↑MOD MS 2022/01/25
                            End If
                            '↓ADD 2022/03/01
                        ElseIf col = "Q00000701" Then
                            para.Add(db.CreateParameter("@Q00000701", SqlDbType.VarChar, chosaSoshiki))

                        ElseIf col = "Q00000801" Then
                            para.Add(db.CreateParameter("@Q00000801", SqlDbType.VarChar, tantoMeisho))
                            '↑ADD 2022/03/01
                        Else
                            para.Add(db.CreateParameter(String.Format("@{0}", col), SqlDbType.Decimal, Nothing))
                        End If
                    Next
                    para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

                    If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                        ret = True
                    Else
                        Throw New Exception(String.Format("{0}追加失敗", tableName))
                    End If
                Else
                    Dim dt As DataTable

                    sb = New System.Text.StringBuilder
                    sb.AppendLine(String.Format("SELECT TOP 0 * FROM {0} ", tableName))

                    dt = db.GetDataTable(sb.ToString)

                    Dim row As DataRow = Nothing

                    'DEL MS 2022/01/25  Dim query = From dr In dc Where dr.Key.Contains(ComConst.ITEM_NO_DELIMITER) Select dr
                    'DEL MS 2022/01/25  For Each kv As KeyValuePair(Of String, DAOChosahyo.調査票項目) In query

                    Dim query = From dr In dcr Where dr.Key.Contains(ComConst.ITEM_NO_DELIMITER) Select dr
                    For Each kv As KeyValuePair(Of String, DAOOther.労働時間整理ファイル項目) In query
                        Dim no As String() = kv.Key.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))

                        row = dt.NewRow()
                        row.BeginEdit()

                        row("調査年") = pKey.chosaNen
                        row("センサス番号") = pKey.censusNo
                        row("項目番号") = no(0)
                        row("明細番号") = no(1)
                        '↓MOD MS 2022/01/25
                        'Dim key = no(0) & ComConst.ITEM_NO_DELIMITER & no(1)
                        'If dcr.ContainsKey(key) Then
                        '    '労働時間整理ファイル側のキーにヒットしたので取得
                        '    row("値") = dcr(key)
                        'Else
                        '    row("値") = kv.Value.値
                        'End If
                        row("値") = kv.Value.値
                        '↑MOD MS 2022/01/25
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

                    '↓INS MS 2022/03/09  調査票項目のKeyにもひっかけないといけないから追加
                    Dim chosaDt As DataTable = dtChosa(tableName)
                    For Each dtrow As DataRow In chosaDt.Rows
                        Dim key = dtrow("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & dtrow("明細番号").ToString
                        If Not dcr.ContainsKey(key) Then
                            '労働時間整理ファイル側のキーにヒットしない場合のみ値をセット
                            row = dt.NewRow()
                            row.BeginEdit()

                            row("調査年") = pKey.chosaNen
                            row("センサス番号") = pKey.censusNo
                            row("項目番号") = dtrow("項目番号")
                            row("明細番号") = dtrow("明細番号")
                            row("値") = dtrow("値")
                            row("更新日付") = Now()
                            row("更新者ID") = CommonInfo.UserId

                            row.EndEdit()
                            dt.Rows.Add(row)
                        End If

                        If dt.Rows.Count = 10000 Then
                            Using bc As New Data.SqlClient.SqlBulkCopy(db.Connection, Data.SqlClient.SqlBulkCopyOptions.KeepIdentity, db.Transaction)
                                bc.DestinationTableName = tableName
                                bc.WriteToServer(dt)
                                bc.Close()
                            End Using

                            dt.Clear()
                        End If
                    Next
                    '↑INS MS 2022/03/09  調査票項目のKeyにもひっかけないといけないから追加

                    If dt.Rows.Count > 0 Then
                        Using bc As New Data.SqlClient.SqlBulkCopy(db.Connection, Data.SqlClient.SqlBulkCopyOptions.KeepIdentity, db.Transaction)
                            bc.DestinationTableName = tableName
                            bc.WriteToServer(dt)
                            bc.Close()
                        End Using
                    End If

                    ret = True
                End If
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票データ追加（労働時間整理ファイル用、再計算時）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks>'REV003 Add</remarks>
    Public Shared Function InsertChosahyoTable_RoudouRe(db As DBAccess,
                                                        pKey As DAOChosahyo.PrimaryKey,
                                                        kKey As DAOOther.RoudouKyotenKey,
                                                        dc As Dictionary(Of String, DAOChosahyo.調査票項目),
                                                        chosakbn As String,
                                                        chosaSoshiki As String,
                                                        tantoMeisho As String) As Boolean

        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(chosakbn)
                If Not tableName.Contains("＿可変") Then
                    '列名称取得
                    Dim colArr As List(Of String) = GetColumns(db, tableName)

                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)

                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("INSERT INTO ""{0}"" ", tableName))
                        .AppendLine("( ")
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        If tableName = ComConst.調査票.テーブル名称(chosakbn)(0) Then
                            .AppendLine("  ,農政局 ")
                            .AppendLine("  ,都道府県 ")
                            .AppendLine("  ,実査設置拠点 ")
                        End If
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,{0} ", col))
                        Next
                        .AppendLine("  ,更新日付 ")
                        .AppendLine("  ,更新者ID ")
                        .AppendLine(") ")
                        .AppendLine("VALUES ")
                        .AppendLine("( ")
                        .AppendLine("   @調査年 ")
                        .AppendLine("  ,@センサス番号 ")
                        If tableName = ComConst.調査票.テーブル名称(chosakbn)(0) Then
                            .AppendLine("  ,@農政局 ")
                            .AppendLine("  ,@都道府県 ")
                            .AppendLine("  ,@実査設置拠点 ")
                        End If
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,@{0} ", col))
                        Next
                        .AppendLine("  ,GETDATE() ")
                        .AppendLine("  ,@更新者ID ")
                        .AppendLine(") ")
                    End With

                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                    para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                    If tableName = ComConst.調査票.テーブル名称(chosakbn)(0) Then
                        para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                        para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                        para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))
                    End If
                    For Each col As String In colArr
                        If dc.ContainsKey(col) Then
                            para.Add(db.CreateParameter(String.Format("@{0}", col), If(dc(col).型区分 = ComConst.型区分.数値型, SqlDbType.Decimal, SqlDbType.VarChar), dc(col).値))
                            '↓ADD 2022/03/01
                        ElseIf col = "Q00000701" Then
                            para.Add(db.CreateParameter("@Q00000701", SqlDbType.VarChar, chosaSoshiki))

                        ElseIf col = "Q00000801" Then
                            para.Add(db.CreateParameter("@Q00000801", SqlDbType.VarChar, tantoMeisho))
                            '↑ADD 2022/03/01
                        Else
                            para.Add(db.CreateParameter(String.Format("@{0}", col), SqlDbType.Decimal, Nothing))
                        End If
                    Next
                    para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

                    If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                        ret = True
                    Else
                        Throw New Exception(String.Format("{0}追加失敗", tableName))
                    End If
                Else
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
                        row("明細番号") = no(1)
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

                    ret = True
                End If
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    '↓2022/02/01 調査票データ追加処理(可変テーブル以外)追加
    ''' <summary>
    ''' 調査票データ追加(可変テーブル以外)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertChosahyoTableKahenIgai(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, kKey As DAOChosahyo.KotenKey, dc As Dictionary(Of String, DAOChosahyo.調査票項目), chosaSoshiki As String, tantoMeisho As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing


        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)
                If Not tableName.Contains("＿可変") Then
                    '列名称取得
                    Dim colArr As List(Of String) = GetColumns(db, tableName)

                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)

                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("INSERT INTO ""{0}"" ", tableName))
                        .AppendLine("( ")
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        If tableName = ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0) Then
                            .AppendLine("  ,農政局 ")
                            .AppendLine("  ,都道府県 ")
                            .AppendLine("  ,実査設置拠点 ")
                        End If
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,{0} ", col))
                        Next
                        .AppendLine("  ,更新日付 ")
                        .AppendLine("  ,更新者ID ")
                        .AppendLine(") ")
                        .AppendLine("VALUES ")
                        .AppendLine("( ")
                        .AppendLine("   @調査年 ")
                        .AppendLine("  ,@センサス番号 ")
                        If tableName = ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0) Then
                            .AppendLine("  ,@農政局 ")
                            .AppendLine("  ,@都道府県 ")
                            .AppendLine("  ,@実査設置拠点 ")
                        End If
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,@{0} ", col))
                        Next
                        .AppendLine("  ,GETDATE() ")
                        .AppendLine("  ,@更新者ID ")
                        .AppendLine(") ")
                    End With

                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                    para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                    If tableName = ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0) Then
                        para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                        para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                        para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))

                    End If
                    For Each col As String In colArr
                        '--- REV_001 ADD Start
                        If dc.ContainsKey(col) Then
                            '--- REV_001 ADD End
                            para.Add(db.CreateParameter(String.Format("@{0}", col), If(dc(col).型区分 = ComConst.型区分.数値型, SqlDbType.Decimal, SqlDbType.VarChar), dc(col).値))
                            '--- REV_001 ADD Start

                            '>>>2022/01/27
                        ElseIf col = "Q00000701" Then
                            para.Add(db.CreateParameter("@Q00000701", SqlDbType.VarChar, chosaSoshiki))

                        ElseIf col = "Q00000801" Then
                            para.Add(db.CreateParameter("@Q00000801", SqlDbType.VarChar, tantoMeisho))
                            '<<<2022/01/27
                        Else
                            para.Add(db.CreateParameter(String.Format("@{0}", col), SqlDbType.Decimal, Nothing))
                        End If
                        '--- REV_001 ADD End
                    Next
                    para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

                    If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                        ret = True
                    Else
                        Throw New Exception(String.Format("{0}追加失敗", tableName))
                    End If

                End If
            Next
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function
    '↑2022/02/01 調査票データ追加処理(可変テーブル以外)追加

    '--- REV_001 ADD START
    ''' <summary>
    ''' 調査票データ追加(可変テーブルのみ)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="dc"></param>
    ''' <param name="clearList">REV_005</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertChosahyoTableKahen(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, kKey As DAOChosahyo.KotenKey, dc As Dictionary(Of String, DAOChosahyo.調査票項目), clearList As List(Of String)) As Boolean
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
                        row("明細番号") = no(1)
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

                    ret = True
                Else
                    '可変以外のテーブルは更新日付、更新者IDのみ更新

                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)

                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("UPDATE ""{0}"" ", tableName))
                        .AppendLine(" SET ")
                        'REV_005↓ clearListがあり、職員テーブルの場合にNULLに更新
                        If clearList.Count > 0 AndAlso tableName.Contains("＿職員") Then
                            For Each koban In clearList
                                .AppendLine(String.Format("  {0} = NULL, ", koban))
                            Next
                        End If
                        'REV_005↑
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

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function
    '--- REV_001 ADD END

    ''' <summary>
    ''' 調査票データ追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertUpdateChosahyoTable_TyukanSyukei(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, kKey As DAOChosahyo.KotenKey, dc As Dictionary(Of String, DAOChosahyo.中間集計表項目)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)
                If Not tableName.Contains("＿可変") Then
                    '列名称取得
                    Dim colArr As List(Of String) = GetColumns(db, tableName)

                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)

                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("UPDATE ""{0}"" ", tableName))
                        .AppendLine(" SET ")
                        For Each col As String In colArr
                            If dc.ContainsKey(col) AndAlso dc(col).登録削除区分 = ComConst.中間集計表.登録削除区分.enm.登録 _
                                AndAlso dc(col).値 IsNot Nothing Then
                                .AppendLine(String.Format("  {0} = @{0}", col))
                                .AppendLine(String.Format(" , "))
                            End If
                        Next
                        .AppendLine("  更新日付 = GETDATE() ")
                        .AppendLine(" ,更新者ID = @更新者ID ")
                        .AppendLine(" WHERE ")
                        .AppendLine(" 調査年 = @調査年")
                        .AppendLine(" AND ")
                        .AppendLine(" センサス番号 = @センサス番号")
                    End With

                    para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                    para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))
                    For Each col As String In colArr
                        If dc.ContainsKey(col) AndAlso dc(col).登録削除区分 = ComConst.中間集計表.登録削除区分.enm.登録 _
                            AndAlso dc(col).値 IsNot Nothing Then
                            para.Add(db.CreateParameter(String.Format("@{0}", col), If(dc(col).型区分 = ComConst.型区分.数値型, SqlDbType.Decimal, SqlDbType.VarChar), dc(col).値))
                        End If
                    Next

                    If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                        ret = True
                    Else
                        Throw New Exception(String.Format("{0}更新失敗", tableName))
                    End If
                Else
                    Dim dt As DataTable

                    sb = New System.Text.StringBuilder
                    sb.AppendLine(String.Format("SELECT TOP 0 * FROM {0} ", tableName))

                    dt = db.GetDataTable(sb.ToString)

                    Dim row As DataRow = Nothing

                    Dim query = From dr In dc Where dr.Key.Contains(ComConst.ITEM_NO_DELIMITER) Select dr
                    For Each kv As KeyValuePair(Of String, DAOChosahyo.中間集計表項目) In query
                        Dim no As String() = kv.Key.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))

                        row = dt.NewRow()
                        row.BeginEdit()

                        row("調査年") = pKey.chosaNen
                        row("センサス番号") = pKey.censusNo
                        row("項目番号") = no(0)
                        row("明細番号") = no(1)
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

                    ret = True
                End If
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票データ追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertChosahyoTable(db As DBAccess, dc As Dictionary(Of String, String)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)
                '列名称取得
                Dim colArr As List(Of String) = GetColumnsJushin(db, tableName, True)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)) Then
                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("INSERT INTO ""{0}"" ", tableName))
                        .AppendLine("( ")
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,{0} ", col))
                        Next
                        .AppendLine("  ,更新日付 ")
                        .AppendLine("  ,更新者ID ")
                        .AppendLine(") ")
                        .AppendLine("SELECT ")
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,{0} ", col))
                        Next
                        .AppendLine("  ,GETDATE() ")
                        .AppendLine("  ,@更新者ID ")
                        .AppendLine(String.Format("FROM   受信＿{0}", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   = @センサス番号 ")
                        .AppendLine("AND    上り下り区分   = @上り下り区分 ")
                        .AppendLine("AND    農政局         = @農政局 ")
                        .AppendLine("AND    都道府県       = @都道府県 ")
                        .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                    End With
                Else
                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("INSERT INTO ""{0}"" ", tableName))
                        .AppendLine("( ")
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,{0} ", col))
                        Next
                        .AppendLine("  ,更新日付 ")
                        .AppendLine("  ,更新者ID ")
                        .AppendLine(") ")
                        .AppendLine("SELECT ")
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,{0} ", col))
                        Next
                        .AppendLine("  ,GETDATE() ")
                        .AppendLine("  ,@更新者ID ")
                        .AppendLine(String.Format("FROM   受信＿{0}", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                        .AppendLine(String.Format("                          FROM   受信＿{0}", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        .AppendLine("                          AND    センサス番号   = @センサス番号 ")
                        .AppendLine("                          AND    上り下り区分   = @上り下り区分 ")
                        .AppendLine("                          AND    農政局         = @農政局 ")
                        .AppendLine("                          AND    都道府県       = @都道府県 ")
                        .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        .AppendLine("                         ) ")
                        .AppendLine("AND    上り下り区分   = @上り下り区分 ")
                    End With
                End If

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, dc("調査年")))
                para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, dc("センサス番号")))
                para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, dc("上り下り区分")))
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, dc("農政局")))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, dc("都道府県")))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, dc("実査設置拠点")))
                para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

                If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                    ret = True
                Else
                    Throw New Exception(String.Format("{0}追加失敗", tableName))
                End If
            Next
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 列名称取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetColumns(db As DBAccess, tableName As String) As List(Of String)
        Dim ret As New List(Of String)

        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing

        Dim comCol As String() = {"調査年", "センサス番号", "農政局", "都道府県", "実査設置拠点", "更新日付", "更新者ID"}

        Try
            sb = New System.Text.StringBuilder

            ' SQL文の設定
            With sb
                .AppendLine("SELECT TOP 0 *")
                .AppendLine(String.Format("FROM   ""{0}""", tableName))
            End With

            dt = db.GetDataTable(sb.ToString)

            For Each col As DataColumn In dt.Columns
                If Not comCol.Contains(col.ColumnName) Then
                    ret.Add(col.ColumnName)
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票データ実査設置拠点取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosahyoKoten(db As DBAccess, pKey As DAOChosahyo.PrimaryKey) As KotenKey

        Dim ret As KotenKey = Nothing
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT 農政局 ")
                .AppendLine("     , 都道府県 ")
                .AppendLine("     , 実査設置拠点 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine("WHERE  調査年         = @調査年 ")
                .AppendLine("AND    センサス番号   = @センサス番号 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))

            dr = db.ExecuteReader(sb.ToString, para)

            If dr.Read Then
                ret = New KotenKey(dr("農政局").ToString, dr("都道府県").ToString, dr("実査設置拠点").ToString)
            End If
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
    ''' 調査票データ削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteChosahyoTable(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, kKey As DAOChosahyo.KotenKey) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun).Reverse().ToArray()
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)) Then
                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   = @センサス番号 ")
                        .AppendLine("AND    農政局         = @農政局 ")
                        .AppendLine("AND    都道府県       = @都道府県 ")
                        .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                    End With
                Else
                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        .AppendLine("                          AND    センサス番号   = @センサス番号 ")
                        .AppendLine("                          AND    農政局         = @農政局 ")
                        .AppendLine("                          AND    都道府県       = @都道府県 ")
                        .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        .AppendLine("                         ) ")
                    End With
                End If

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))

                If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                    ret = True
                Else
                    Throw New Exception(String.Format("{0}削除失敗", tableName))
                End If
            Next
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    'REV_008
    ''' <summary>
    ''' 調査票調査対象経営体と担当名称取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Shared Function GetKeieitaiAndTantoName(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, kKey As DAOChosahyo.KotenKey) As DataTable
        Dim ret As DataTable = Nothing
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT Q00000701, Q00000801")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine("WHERE  調査年 = @調査年")
                .AppendLine("AND    センサス番号 = @センサス番号")
                .AppendLine("AND    農政局 = @農政局")
                .AppendLine("AND    都道府県 = @都道府県")
                .AppendLine("AND    実査設置拠点 = @実査設置拠点")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    '↓調査票データ削除処理（可変テーブル以外）追加
    ''' <summary>
    ''' 調査票データ削除（可変テーブル以外）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteChosahyoTableKahenIgai(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, kKey As DAOChosahyo.KotenKey) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun).Reverse().ToArray()
                If Not tableName.Contains("＿可変") Then
                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)

                    If tableName.Equals(ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)) Then
                        ' SQL文の設定
                        With sb
                            .AppendLine("DELETE ")
                            .AppendLine(String.Format("FROM   ""{0}""", tableName))
                            .AppendLine("WHERE  調査年         = @調査年 ")
                            .AppendLine("And    センサス番号   = @センサス番号 ")
                            .AppendLine("And    農政局         = @農政局 ")
                            .AppendLine("And    都道府県       = @都道府県 ")
                            .AppendLine("And    実査設置拠点   = @実査設置拠点 ")
                        End With
                    Else
                        ' SQL文の設定
                        With sb
                            .AppendLine("DELETE ")
                            .AppendLine(String.Format("FROM   ""{0}""", tableName))
                            .AppendLine("WHERE  調査年         = @調査年 ")
                            .AppendLine("And    センサス番号   In (Select センサス番号 ")
                            .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                            .AppendLine("                          WHERE  調査年         = @調査年 ")
                            .AppendLine("                          And    センサス番号   = @センサス番号 ")
                            .AppendLine("                          And    農政局         = @農政局 ")
                            .AppendLine("                          And    都道府県       = @都道府県 ")
                            .AppendLine("                          And    実査設置拠点   = @実査設置拠点 ")
                            .AppendLine("                         ) ")
                        End With
                    End If

                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                    para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                    para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                    para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                    para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))

                    If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                        ret = True
                    Else
                        Throw New Exception(String.Format("{0}削除失敗", tableName))
                    End If
                End If
            Next
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function
    '↑調査票データ削除処理（可変テーブル以外）追加

    ''' <summary>
    ''' 調査票データ削除（労働時間整理ファイル用）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="chosakbn"></param>
    ''' <param name="allDelete">True:全て削除、False:可変のみ削除</param>
    ''' <returns></returns>
    ''' <remarks>'REV003 Add</remarks>
    Public Shared Function DeleteChosahyoTable_Roudou(db As DBAccess,
                                                      pKey As DAOChosahyo.PrimaryKey,
                                                      kKey As DAOOther.RoudouKyotenKey,
                                                      chosakbn As String,
                                                      allDelete As Boolean) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(chosakbn).Reverse().ToArray()
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                '↓MOD MS 2022/01/25
                'If tableName.Equals(ComConst.調査票.テーブル名称(chosakbn)(0)) Then
                '' SQL文の設定
                'With sb
                '    .AppendLine("DELETE ")
                '    .AppendLine(String.Format("FROM   ""{0}""", tableName))
                '    .AppendLine("WHERE  調査年         = @調査年 ")
                '    .AppendLine("AND    センサス番号   = @センサス番号 ")
                '    .AppendLine("AND    農政局         = @農政局 ")
                '    .AppendLine("AND    都道府県       = @都道府県 ")
                '    .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                'End With
                If Not tableName.Contains("＿可変") Then
                    If allDelete Then
                        ' SQL文の設定
                        With sb
                            .AppendLine("DELETE ")
                            .AppendLine(String.Format("FROM   ""{0}""", tableName))
                            .AppendLine("WHERE  調査年         = @調査年 ")
                            .AppendLine("AND    センサス番号   = @センサス番号 ")
                            If tableName = ComConst.調査票.テーブル名称(chosakbn)(0) Then
                                .AppendLine("AND    農政局         = @農政局 ")
                                .AppendLine("AND    都道府県       = @都道府県 ")
                                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                            End If
                        End With

                        para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                        para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                        If tableName = ComConst.調査票.テーブル名称(chosakbn)(0) Then
                            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))
                        End If

                        If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                            ret = True
                        Else
                            Throw New Exception(String.Format("{0}削除失敗", tableName))
                        End If
                    End If
                    '↑MOD MS 2022/01/25
                Else
                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.調査票.テーブル名称(chosakbn)(0)))
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        .AppendLine("                          AND    センサス番号   = @センサス番号 ")
                        .AppendLine("                          AND    農政局         = @農政局 ")
                        .AppendLine("                          AND    都道府県       = @都道府県 ")
                        .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        .AppendLine("                         ) ")
                    End With

                    '↓INS MS 2022/01/25
                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                    para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                    para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                    para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                    para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))

                    If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                        ret = True
                    Else
                        Throw New Exception(String.Format("{0}削除失敗", tableName))
                    End If
                    '↑INS MS 2022/01/25
                End If
                '↓DEL MS 2022/01/25
                'para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                'para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                'para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                'para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                'para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))

                'If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                '    ret = True
                'Else
                '    Throw New Exception(String.Format("{0}削除失敗", tableName))
                'End If
                '↑DEL MS 2022/01/25
            Next
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票データ削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteChosahyoTable_TyukanSyukei(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, kKey As DAOChosahyo.KotenKey) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun).Reverse().ToArray()
                If tableName.Contains("＿可変") Then
                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)

                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        .AppendLine("                          AND    センサス番号   = @センサス番号 ")
                        .AppendLine("                          AND    農政局         = @農政局 ")
                        .AppendLine("                          AND    都道府県       = @都道府県 ")
                        .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        .AppendLine("                         ) ")
                    End With

                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                    para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                    para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                    para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                    para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))

                    If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                        ret = True
                    Else
                        Throw New Exception(String.Format("{0}削除失敗", tableName))
                    End If
                End If
            Next
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票データ削除更新
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteUpdateChosahyoTable_TyukanSyukei(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, kKey As DAOChosahyo.KotenKey, dc As Dictionary(Of String, DAOChosahyo.中間集計表項目)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)
                If Not tableName.Contains("＿可変") Then
                    '列名称取得
                    Dim colArr As List(Of String) = GetColumns(db, tableName)

                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)

                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("UPDATE ""{0}"" ", tableName))
                        .AppendLine(" SET ")
                        For Each col As String In colArr
                            If dc.ContainsKey(col) AndAlso dc(col).登録削除区分 = ComConst.中間集計表.登録削除区分.enm.削除 Then
                                .AppendLine(String.Format("  {0} = @{0}", col))
                                .AppendLine(String.Format(" , "))
                            End If
                        Next
                        .AppendLine("  更新日付 = GETDATE() ")
                        .AppendLine(" ,更新者ID = @更新者ID ")
                        .AppendLine(" WHERE ")
                        .AppendLine(" 調査年 = @調査年")
                        .AppendLine(" AND ")
                        .AppendLine(" センサス番号 = @センサス番号")
                    End With

                    para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                    para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))
                    For Each col As String In colArr
                        If dc.ContainsKey(col) AndAlso dc(col).登録削除区分 = ComConst.中間集計表.登録削除区分.enm.削除 Then
                            para.Add(db.CreateParameter(String.Format("@{0}", col), If(dc(col).型区分 = ComConst.型区分.数値型, SqlDbType.Decimal, SqlDbType.VarChar), dc(col).値))
                        End If
                    Next

                    If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                        ret = True
                    Else
                        Throw New Exception(String.Format("{0}更新失敗", tableName))
                    End If
                Else
                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)
                    Dim colArr As List(Of String) = GetColumns(db, tableName)

                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        .AppendLine("                          AND    センサス番号   = @センサス番号 ")
                        .AppendLine("                          AND    農政局         = @農政局 ")
                        .AppendLine("                          AND    都道府県       = @都道府県 ")
                        .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        .AppendLine(String.Format(" AND "))
                        .AppendLine(String.Format("  項目番号 in ("))
                        Dim query = From dr In dc Where dr.Key.Contains(ComConst.ITEM_NO_DELIMITER) Select dr
                        Dim number As Integer = 0
                        For Each kv As KeyValuePair(Of String, DAOChosahyo.中間集計表項目) In query
                            Dim no As String() = kv.Key.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))
                            If dc.ContainsKey(no(0) & ComConst.ITEM_NO_DELIMITER & 1) AndAlso dc(no(0) & ComConst.ITEM_NO_DELIMITER & 1).登録削除区分 = ComConst.中間集計表.登録削除区分.enm.削除 Then
                                If number > 0 Then
                                    .AppendLine(String.Format("  ,  ", no(0)))
                                End If
                                number = number + 1
                                .AppendLine(String.Format("  '{0}'  ", no(0)))
                            End If
                        Next
                        .AppendLine("                         ) ")
                        .AppendLine("                         ) ")
                    End With

                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                    para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                    para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                    para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                    para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))

                    If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                        ret = True
                    Else
                        Throw New Exception(String.Format("{0}削除失敗", tableName))
                    End If

                End If
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 列名称取得（受信）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="tableName"></param>
    ''' <param name="update"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetColumnsJushin(db As DBAccess, tableName As String, Optional update As Boolean = False) As List(Of String)
        Dim ret As New List(Of String)

        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing

        Dim comCol As String() = {}
        If update Then
            comCol = {"調査年", "センサス番号", "上り下り区分", "更新日付", "更新者ID"}
        Else

            comCol = {"調査年", "センサス番号", "上り下り区分"}
        End If

        Try
            sb = New System.Text.StringBuilder

            ' SQL文の設定
            With sb
                .AppendLine("SELECT TOP 0 *")
                .AppendLine(String.Format("FROM   ""受信＿{0}""", tableName))
            End With

            dt = db.GetDataTable(sb.ToString)

            For Each col As DataColumn In dt.Columns
                If Not comCol.Contains(col.ColumnName) Then
                    ret.Add(col.ColumnName)
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票データ追加（受信）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="dc"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertChosahyoTableJushin(db As DBAccess, dc As Dictionary(Of String, String), chosaKoutei As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)
                '列名称取得
                Dim colArr As List(Of String) = GetColumnsJushin(db, tableName)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)) Then
                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("INSERT INTO [{0}].[dbo].[受信＿{1}] ", chosaKoutei, tableName))
                        .AppendLine("( ")
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        .AppendLine("  ,上り下り区分 ")
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,{0} ", col))
                        Next
                        .AppendLine(") ")
                        .AppendLine("SELECT ")
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        .AppendLine("  ,@上り下り区分 ")
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,{0} ", col))
                        Next
                        .AppendLine(String.Format("FROM   {0}", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   = @センサス番号 ")
                        .AppendLine("AND    農政局         = @農政局 ")
                        .AppendLine("AND    都道府県       = @都道府県 ")
                        .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                    End With
                Else
                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("INSERT INTO [{0}].[dbo].[受信＿{1}] ", chosaKoutei, tableName))
                        .AppendLine("( ")
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        .AppendLine("  ,上り下り区分 ")
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,{0} ", col))
                        Next
                        .AppendLine(") ")
                        .AppendLine("SELECT ")
                        .AppendLine("   調査年 ")
                        .AppendLine("  ,センサス番号 ")
                        .AppendLine("  ,@上り下り区分 ")
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,{0} ", col))
                        Next
                        .AppendLine(String.Format("FROM   {0}", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        .AppendLine("                          AND    センサス番号   = @センサス番号 ")
                        .AppendLine("                          AND    農政局         = @農政局 ")
                        .AppendLine("                          AND    都道府県       = @都道府県 ")
                        .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        .AppendLine("                         ) ")
                    End With
                End If

                para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, dc("上り下り区分")))

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, dc("調査年")))
                para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, dc("センサス番号")))
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, dc("農政局")))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, dc("都道府県")))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, dc("実査設置拠点")))

                If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                    ret = True
                Else
                    Throw New Exception(String.Format("受信＿{0}追加失敗", tableName))
                End If
            Next
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票データ削除（受信）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="dc"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteChosahyoTableJushin(db As DBAccess, dc As Dictionary(Of String, String), Optional chosaKoutei As String = Nothing) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun).Reverse().ToArray()
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)) Then
                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        If chosaKoutei Is Nothing Then
                            .AppendLine(String.Format("FROM   受信＿{0}", tableName))
                        Else
                            .AppendLine(String.Format("FROM   [{0}].[dbo].[受信＿{1}]", chosaKoutei, tableName))
                        End If
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   = @センサス番号 ")
                        .AppendLine("AND    上り下り区分   = @上り下り区分 ")
                        .AppendLine("AND    農政局         = @農政局 ")
                        .AppendLine("AND    都道府県       = @都道府県 ")
                        .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                    End With
                Else
                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        If chosaKoutei Is Nothing Then
                            .AppendLine(String.Format("FROM   受信＿{0}", tableName))
                        Else
                            .AppendLine(String.Format("FROM   [{0}].[dbo].[受信＿{1}]", chosaKoutei, tableName))
                        End If
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                        If chosaKoutei Is Nothing Then
                            .AppendLine(String.Format("                          FROM   受信＿{0}", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                        Else
                            .AppendLine(String.Format("                          FROM   [{0}].[dbo].[受信＿{1}]", chosaKoutei, ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                        End If
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        .AppendLine("                          AND    センサス番号   = @センサス番号 ")
                        .AppendLine("                          AND    上り下り区分   = @上り下り区分 ")
                        .AppendLine("                          AND    農政局         = @農政局 ")
                        .AppendLine("                          AND    都道府県       = @都道府県 ")
                        .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        .AppendLine("                         ) ")
                        .AppendLine("AND    上り下り区分   = @上り下り区分 ")
                    End With
                End If

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, dc("調査年")))
                para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, dc("センサス番号")))
                para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, dc("上り下り区分")))
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, dc("農政局")))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, dc("都道府県")))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, dc("実査設置拠点")))

                If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                    ret = True
                Else
                    Throw New Exception(String.Format("受信＿{0}削除失敗", tableName))
                End If
            Next
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票テーブル取得（バックアップ）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetBackUpChosahyoTable(db As DBAccess, chosaKubun As String, chosaNen As String, kyoku As String, jimusho As String, kyoten As String) As Dictionary(Of String, DataTable)
        Dim ret As New Dictionary(Of String, DataTable)
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(chosaKubun)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.調査票.テーブル名称(chosaKubun)(0)) Then
                    ' SQL文の設定
                    With sb
                        .AppendLine("SELECT * ")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        If Not kyoku Is Nothing Then
                            .AppendLine("AND    農政局       = @農政局 ")
                        End If
                        If Not jimusho Is Nothing Then
                            .AppendLine("AND    都道府県       = @都道府県 ")
                        End If
                        If Not kyoten Is Nothing Then
                            .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                        End If
                    End With

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
                Else
                    ' SQL文の設定
                    With sb
                        .AppendLine("SELECT *")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.調査票.テーブル名称(chosaKubun)(0)))
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        If Not kyoku Is Nothing Then
                            .AppendLine("                          AND    農政局       = @農政局 ")
                        End If
                        If Not jimusho Is Nothing Then
                            .AppendLine("                          AND    都道府県       = @都道府県 ")
                        End If
                        If Not kyoten Is Nothing Then
                            .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        End If
                        .AppendLine("                         ) ")
                    End With

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
                End If

                dt = db.GetDataTable(sb.ToString, para)

                ret.Add(tableName, dt)
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票データ削除（レストア）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteRestoreChosahyoTable(db As DBAccess, chosaKubun As String, chosaNen As String, kyoku As String, jimusho As String, kyoten As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(chosaKubun).Reverse().ToArray()
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.調査票.テーブル名称(chosaKubun)(0)) Then
                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        If Not kyoku Is Nothing Then
                            .AppendLine("AND    農政局       = @農政局 ")
                        End If
                        If Not jimusho Is Nothing Then
                            .AppendLine("AND    都道府県       = @都道府県 ")
                        End If
                        If Not kyoten Is Nothing Then
                            .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                        End If
                    End With

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
                Else
                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.調査票.テーブル名称(chosaKubun)(0)))
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        If Not kyoku Is Nothing Then
                            .AppendLine("                          AND    農政局       = @農政局 ")
                        End If
                        If Not jimusho Is Nothing Then
                            .AppendLine("                          AND    都道府県       = @都道府県 ")
                        End If
                        If Not kyoten Is Nothing Then
                            .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        End If
                        .AppendLine("                         ) ")
                    End With

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
                End If

                If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                    ret = True
                Else
                    Throw New Exception(String.Format("{0}削除失敗", tableName))
                End If
            Next
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' プレプリント削除対象データ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="ronriKubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetPrePrintDelData(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, ronriKubun As String) As List(Of String)

        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim dt As DataTable
        Dim ret As List(Of String) = New List(Of String)
        Dim table As String = If(CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別, "調査票＿農業経営＿牛乳＿個別＿可変", "調査票＿経営分析＿牛乳＿可変")

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)
            Select Case ronriKubun
                Case "ア"
                    ' SQL文の設定
                    With sb
                        .AppendLine("SELECT DISTINCT row1.* ")
                        .AppendLine("FROM " & table & " AS row1")
                        .AppendLine("INNER JOIN " & table & " AS row2")
                        .AppendLine("ON row1.調査年 = row2.調査年")
                        .AppendLine("AND row1.センサス番号 = row2.センサス番号")
                        .AppendLine("AND row1.値 = row2.値")
                        .AppendLine("INNER JOIN " & table & " AS row3")
                        .AppendLine("ON row2.調査年 = row3.調査年")
                        .AppendLine("AND row2.センサス番号 = row3.センサス番号")
                        .AppendLine("AND row2.明細番号 = row3.明細番号")
                        '--- REV_002 DEL START
                        '.AppendLine("INNER JOIN " & table & " AS row4")
                        '.AppendLine("ON row2.調査年 = row4.調査年")
                        '.AppendLine("AND row2.センサス番号 = row4.センサス番号")
                        '.AppendLine("AND row2.明細番号 = row4.明細番号")
                        '--- REV_002 DEL END
                        .AppendLine("LEFT JOIN (")
                        'REV_004↓
                        '.AppendLine("SELECT * FROM " & table & " WHERE センサス番号 = @センサス番号  AND 項目番号 = 'Q11020601' AND (値 = '1：出（入牧）' or 値 = '2：戻（下牧）')")
                        .AppendLine("SELECT * FROM " & table & " WHERE 調査年 = @調査年 AND センサス番号 = @センサス番号  AND 項目番号 = 'Q11020601' AND (値 = '1：出（入牧）' or 値 = '2：戻（下牧）')")
                        'REV_004↑
                        .AppendLine(") row5")
                        '--- REV_002 MOD START
                        '.AppendLine("ON row1.明細番号 = row5.明細番号")
                        .AppendLine("ON row2.明細番号 = row5.明細番号")
                        '--- REV_002 MOD END
                        .AppendLine("WHERE row1.センサス番号 = @センサス番号")
                        .AppendLine("AND row1.調査年 = @調査年")
                        .AppendLine("AND row1.項目番号 = 'Q11021801'")
                        .AppendLine("AND row2.項目番号 = 'Q11020101'")
                        .AppendLine("AND row3.項目番号 = 'Q11020401' ")
                        .AppendLine("AND (row3.値 = '２：売却（売却）' OR row3.値 = '２：売却（死亡）')")
                        .AppendLine("AND row5.明細番号 IS NULL")
                        .AppendLine("ORDER BY row1.項目番号,row1.明細番号")
                    End With
                Case "イ"
                    ' SQL文の設定
                    With sb
                        .AppendLine("SELECT DISTINCT row1.* ")
                        .AppendLine("FROM " & table & " AS row1")
                        .AppendLine("INNER JOIN " & table & " AS row2")
                        .AppendLine("ON row1.調査年 = row2.調査年")
                        .AppendLine("AND row1.センサス番号 = row2.センサス番号")
                        .AppendLine("AND row1.値 = row2.値")
                        .AppendLine("INNER JOIN " & table & " AS row3")
                        .AppendLine("ON row2.調査年 = row3.調査年")
                        .AppendLine("AND row2.センサス番号 = row3.センサス番号")
                        .AppendLine("AND row2.明細番号 = row3.明細番号")
                        '--- REV_002 DEL START
                        '.AppendLine("INNER JOIN " & table & " AS row4")
                        '.AppendLine("ON row1.調査年 = row4.調査年")
                        '.AppendLine("AND row1.センサス番号 = row4.センサス番号")
                        '.AppendLine("AND row1.値 = row4.値")
                        '--- REV_002 DEL END
                        .AppendLine("LEFT JOIN (")
                        '--- REV_002 MOD START
                        '.AppendLine("	SELECT * FROM " & table & " WHERE センサス番号 = @センサス番号 AND 項目番号 = 'Q11020601' AND 値 = '2：戻（下牧）'")
                        .AppendLine("	SELECT B.明細番号, B.値 FROM " & table & " AS A")
                        .AppendLine("	INNER JOIN " & table & " AS B")
                        .AppendLine("	ON A.調査年 = B.調査年 AND A.センサス番号 = B.センサス番号 AND A.明細番号 = B.明細番号")
                        .AppendLine("	WHERE A.調査年 = @調査年 AND A.センサス番号 = @センサス番号 AND A.項目番号 = 'Q11020601' AND A.値 = '2：戻（下牧）' AND B.項目番号 =  'Q11020101'")
                        '--- REV_002 MOD END
                        .AppendLine("	) row5")
                        '--- REV_002 MOD START
                        '.AppendLine("ON row1.明細番号 = row5.明細番号")
                        .AppendLine("ON row1.値 = row5.値")
                        '--- REV_002 MOD END
                        .AppendLine("WHERE row1.センサス番号 = @センサス番号")
                        .AppendLine("AND row1.調査年 = @調査年")
                        .AppendLine("AND row1.項目番号 = 'Q11021801'")
                        .AppendLine("AND row2.項目番号 = 'Q11020101' ")
                        .AppendLine("AND row3.項目番号 = 'Q11020601' ")
                        .AppendLine("AND row3.値 = '1：出（入牧）'")
                        '--- REV_002 DEL START
                        '.AppendLine("AND row4.項目番号 = 'Q11020101' ")
                        '--- REV_002 DEL END
                        .AppendLine("AND row5.明細番号 IS NULL")
                    End With
            End Select

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))

            dt = db.GetDataTable(sb.ToString, para)

            For Each row As DataRow In dt.Rows
                ret.Add(ComUtil.Chosahyo.GetEdaNo(row.Item("明細番号").ToString))
            Next

            Return ret

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    '--- REV_001 ADD START
    ''' <summary>
    ''' 比較用調査テーブル取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="einouRuikei"></param>
    ''' <param name="ChosaKoutei"></param>
    ''' <param name="upLow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetHikakuChosahyoTable(db As DBAccess, chosaNen As String, kyoku As String, jimusho As String, kyoten As String, einouRuikei As String, upLow As String,
                                                  ChosaKouteiTo As String, ChosaKouteiFrom As String, Optional jushin As Boolean = False) As Dictionary(Of String, DataTable)
        Dim ret As Dictionary(Of String, DataTable) = New Dictionary(Of String, DataTable)
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim kobetsuTableName As String
        Dim otherTableName As String
        Dim mainTableName As String
        Dim dt As DataTable
        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)
                If jushin Then
                    kobetsuTableName = String.Format("[{0}].[dbo].[受信＿{1}]", ChosaKouteiTo, ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0))
                    mainTableName = String.Format("[{0}].[dbo].[受信＿{1}]", ChosaKouteiTo, tableName)
                    otherTableName = String.Format("[{0}].[dbo].[{1}]", ChosaKouteiFrom, ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0))
                Else
                    kobetsuTableName = String.Format("[{0}].[dbo].[{1}]", ChosaKouteiFrom, ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0))
                    mainTableName = String.Format("[{0}].[dbo].[{1}]", ChosaKouteiFrom, tableName)
                    otherTableName = String.Format("[{0}].[dbo].[受信＿{1}]", ChosaKouteiTo, ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0))
                End If

                With sb
                    If tableName = ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0) Then
                        .AppendLine("SELECT A.* ")
                        .AppendLine(String.Format("FROM {0} As A", kobetsuTableName))
                    Else
                        .AppendLine("SELECT B.* ")
                        .AppendLine("  ,A.農政局 ")
                        .AppendLine("  ,A.都道府県 ")
                        .AppendLine("  ,A.実査設置拠点 ")
                        .AppendLine(String.Format("FROM {0} As A", kobetsuTableName))
                        If Not tableName.Contains("＿可変") Then
                            .AppendLine(String.Format("INNER JOIN {0} As B", mainTableName))
                            .AppendLine(" ON A.調査年 = B.調査年 ")
                            .AppendLine(" AND A.センサス番号 = B.センサス番号 ")
                            If jushin Then
                                .AppendLine(" AND A.上り下り区分 = B.上り下り区分 ")
                            End If
                        Else
                            .AppendLine(String.Format("INNER JOIN {0} As B", mainTableName))
                            .AppendLine(" ON A.調査年 = B.調査年 ")
                            .AppendLine(" AND A.センサス番号 = B.センサス番号 ")
                            If jushin Then
                                .AppendLine(" AND A.上り下り区分 = B.上り下り区分 ")
                            End If
                            .AppendLine(String.Format("INNER JOIN {0} As C", otherTableName))
                            .AppendLine(" ON A.調査年 = C.調査年 ")
                            .AppendLine(" AND A.センサス番号 = C.センサス番号 ")
                        End If
                    End If

                    .AppendLine("WHERE  A.調査年         = @調査年 ")

                    If Not kyoku Is Nothing Then
                        .AppendLine("                          AND    A.農政局       = @農政局 ")
                    End If
                    If Not jimusho Is Nothing Then
                        .AppendLine("                          AND    A.都道府県       = @都道府県 ")
                    End If
                    If Not kyoten Is Nothing Then
                        .AppendLine("                          AND    A.実査設置拠点   = @実査設置拠点 ")
                    End If
                    If Not einouRuikei Is Nothing Then
                        .AppendLine(String.Format("AND    SUBSTRING(A.{0}, 1, CHARINDEX( '：', A.{0} ) - 1 )   = @営農類型 ", ComConst.調査票.営農類型(CommonInfo.Chosakubun)))
                    End If
                    If jushin Then
                        .AppendLine("                          AND    A.上り下り区分   = @上り下り区分 ")
                    Else
                        If tableName.Contains("＿可変") Then
                            .AppendLine("                          AND    C.上り下り区分   = @上り下り区分 ")
                        End If
                    End If
                    .AppendLine("ORDER BY センサス番号 ASC")
                End With

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

                If Not einouRuikei Is Nothing Then
                    para.Add(db.CreateParameter("@営農類型", SqlDbType.VarChar, einouRuikei))
                End If

                para.Add(db.CreateParameter("@上り下り区分", SqlDbType.VarChar, upLow))

                dt = db.GetDataTable(sb.ToString, para)

                ret.Add(tableName, dt)

            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票項目名マスタから取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getChosahyoColName(db As DBAccess) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * ")
                .AppendLine("FROM   調査票項目名マスタ")
                .AppendLine("WHERE  調査区分 = @調査区分")
                .AppendLine("ORDER BY 項番 ASC")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.VarChar, CommonInfo.Chosakubun))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票テーブル(CSV出力用)取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCsvChosahyoTable(db As DBAccess, pKey As DAOChosahyo.PrimaryKey) As Dictionary(Of String, DataTable)
        Dim ret As New Dictionary(Of String, DataTable)
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    If tableName = ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0) Then
                        .AppendLine("SELECT A.* ")
                        .AppendLine(String.Format("FROM {0} As A", tableName))
                    Else
                        '列名称取得
                        Dim colArr As List(Of String) = GetColumns(db, tableName)

                        .AppendLine("SELECT B.調査年 ")
                        .AppendLine("  ,B.センサス番号 ")
                        .AppendLine("  ,A.農政局 ")
                        .AppendLine("  ,A.都道府県 ")
                        .AppendLine("  ,A.実査設置拠点 ")
                        For Each col As String In colArr
                            .AppendLine(String.Format("  ,B.{0} ", col))
                        Next
                        .AppendLine("  ,B.更新日付 ")
                        .AppendLine("  ,B.更新者ID ")
                        .AppendLine(String.Format("FROM {0} As A", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                        .AppendLine(String.Format("INNER JOIN {0} As B", tableName))
                        .AppendLine(" ON A.調査年 = B.調査年 ")
                        .AppendLine(" AND A.センサス番号 = B.センサス番号 ")
                    End If
                    .AppendLine("WHERE  A.調査年         = @調査年 ")
                    .AppendLine("AND    A.センサス番号   = @センサス番号 ")
                End With

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))

                dt = db.GetDataTable(sb.ToString, para)

                '2023/11/27 DAIKO ADD START REV_007
                '下記項目はDB・項番はそのままとし、電子調査票CSVに出力を行わない。
                '育成牛・肥育牛（職員）：Q02024802
                If tableName = ComConst.調査票.テーブル名称("20")(1) Or tableName = ComConst.調査票.テーブル名称("21")(1) Or tableName = ComConst.調査票.テーブル名称("22")(1) Or tableName = ComConst.調査票.テーブル名称("23")(1) Or tableName = ComConst.調査票.テーブル名称("24")(1) Then
                    dt.Columns.Remove("Q02024802")
                End If
                '2023/11/27 DAIKO ADD START REV_007

                ret.Add(tableName, dt)
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 労賃単価の反映
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function setRouchinTanka(db As DBAccess, seisanhiKbn As String, chosanen As String, kyoku As String, jimusho As String, kyoten As String) As Boolean
        Dim ret As Boolean
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("UPDATE SH ")
                .AppendLine(String.Format("Set {0} =  RO.男女平均＿評価単価 ", ComConst.労賃単価反映.労賃単価カラム(CommonInfo.Chosakubun)))
                .AppendLine(String.Format("FROM {0} As CH ", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine(String.Format("INNER JOIN {0} As SH", ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(1)))
                .AppendLine(" ON CH.調査年 = SH.調査年 ")
                .AppendLine(" AND CH.センサス番号 = SH.センサス番号 ")
                .AppendLine(String.Format(" INNER JOIN BRAN.dbo.労賃単価 As RO "))
                .AppendLine(" ON RO.調査年 = CH.調査年 ")
                .AppendLine(" AND RO.都道府県 = CH.都道府県 ")
                .AppendLine(" AND RO.生産費区分 = @生産費区分 ")
                .AppendLine("WHERE  CH.調査年         = @調査年 ")
                If Not kyoku Is Nothing Then
                    .AppendLine("                          AND    CH.農政局       = @農政局 ")
                End If
                If Not jimusho Is Nothing Then
                    .AppendLine("                          AND    CH.都道府県       = @都道府県 ")
                End If
                If Not kyoten Is Nothing Then
                    .AppendLine("                          AND    CH.実査設置拠点   = @実査設置拠点 ")
                End If
            End With

            para.Add(db.CreateParameter("@生産費区分", SqlDbType.VarChar, seisanhiKbn))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, jimusho))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kyoten))


            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("更新失敗")
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票に牛トレサ総括データが存在するかチェックする
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="censusNo"></param>
    ''' <param name="chosakubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Shared Function CheckTresaSummaryExist(db As DBAccess, chosaNen As String, censusNo As String, Optional chosakubun As String = "") As Boolean

        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.調査票.テーブル名称(chosakubun)(2)))
                .AppendLine("WHERE  調査年          = @調査年 ")
                .AppendLine("AND    センサス番号    = @センサス番号 ")
                .AppendLine("AND    項目番号        = @項目番号 ")
            End With

            If String.IsNullOrEmpty(chosakubun) Then
                chosakubun = CommonInfo.Chosakubun
            End If

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, censusNo))
            para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, ComConst.調査票.項目番号.牛識別番号(chosakubun)))

            dt = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return If(dt.Rows.Count > 0, True, False)
    End Function

    ''' <summary>
    ''' 牛総括データの削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen">調査年</param>
    ''' <param name="censusNo">センサス番号</param>
    ''' <param name="chosakubun">調査区分</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteTresaSoukatsu(db As DBAccess, chosaNen As String, censusNo As String, chosakubun As String) As Boolean

        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("DELETE C ")
                .AppendLine("FROM " & ComConst.調査票.テーブル名称(chosakubun)(2) & " AS C ")
                .AppendLine("  INNER JOIN 調査票項目マスタ AS M ")
                .AppendLine("    ON C.項目番号 = M.項目番号 ")
                .AppendLine("WHERE  調査年        = @調査年 ")
                .AppendLine("  AND  センサス番号  = @センサス番号 ")
                .AppendLine("  AND  調査区分      = @調査区分 ")
                .AppendLine("  AND  シート名      = @シート名 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, censusNo))
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            para.Add(db.CreateParameter("@シート名", SqlDbType.NVarChar, ComConst.調査票.牛総括データシート(chosakubun)))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception("調査票データ削除失敗")
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function


    ''' <summary>
    ''' 牛トレサ決算月取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="censusNo"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="pkessanNum"></param>
    ''' <returns></returns>
    Public Shared Function GettoresaKessan(db As DBAccess, chosaNen As String, censusNo As String, chosakubun As String, pkessanNum As String) As String

        Dim ret As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim dt As DataTable

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT " & pkessanNum)
                .AppendLine(" FROM " & ComConst.調査票.テーブル名称(chosakubun)(0))
                .AppendLine("WHERE  調査年        = @調査年 ")
                .AppendLine("  AND  センサス番号  = @センサス番号 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, censusNo))
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))

            dt = db.GetDataTable(sb.ToString, para)

            If dt.Rows.Count = 0 Then
                ret = String.Empty
            Else
                If IsDBNull(dt.Rows(0)(0)) Then
                    ret = "12"
                Else
                    ret = CStr(dt.Rows(0)(0))
                End If

            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    '--- REV_001 ADD END

    Public Shared Function UpdateChosahyoTable_SoshikiTanto(db As DBAccess, pKey As DAOChosahyo.PrimaryKey, kKey As DAOChosahyo.KotenKey, chosaSoshiki As String, tantoMeisho As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)
                If Not tableName.Contains("＿可変") And Not tableName.Contains("＿職員") Then
                    '列名称取得
                    Dim colArr As List(Of String) = GetColumns(db, tableName)

                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)

                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("UPDATE ""{0}"" ", tableName))
                        .AppendLine(" SET ")
                    End With
                    If chosaSoshiki <> "" And tantoMeisho <> "" Then
                        With sb
                            .AppendLine("  Q00000701 = @調査対象組織体 ,")
                            .AppendLine("  Q00000801 = @担当名称 ")
                        End With

                        para.Add(db.CreateParameter("@調査対象組織体", SqlDbType.VarChar, chosaSoshiki))
                        para.Add(db.CreateParameter("@担当名称", SqlDbType.VarChar, tantoMeisho))
                    ElseIf chosaSoshiki <> "" Then
                        With sb
                            .AppendLine("  Q00000701 = @調査対象組織体 ,")
                            .AppendLine("  Q00000801 = NULL ")
                        End With

                        para.Add(db.CreateParameter("@調査対象組織体", SqlDbType.VarChar, chosaSoshiki))
                    ElseIf tantoMeisho <> "" Then
                        With sb
                            .AppendLine("  Q00000701 = NULL ,")
                            .AppendLine("  Q00000801 = @担当名称 ")
                        End With

                        para.Add(db.CreateParameter("@担当名称", SqlDbType.VarChar, tantoMeisho))
                    Else
                        With sb
                            .AppendLine("  Q00000701 = NULL ,")
                            .AppendLine("  Q00000801 = NULL ")
                        End With
                    End If

                    With sb
                        .AppendLine(" WHERE ")
                        .AppendLine(" センサス番号 = @センサス番号")
                    End With

                    para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))

                    If db.ExecuteNonQuery(sb.ToString, para) >= 1 Then
                        ret = True
                    Else
                        Throw New Exception(String.Format("{0}更新失敗", tableName))
                    End If

                End If
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    Public Shared Function GetSenmonChosainName(db As DBAccess) As String
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim name As String = ""

        Try
            If CommonInfo.SenmonChosain Then
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT 氏名 ")
                    .AppendLine("FROM   専門調査員管理")
                    .AppendLine("WHERE  ユーザーID = @ユーザーID")
                End With

                para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, CommonInfo.UserId))

                ret = db.GetDataTable(sb.ToString, para)

                For Each row As DataRow In ret.Rows
                    '1件の想定だが…
                    name = row("氏名").ToString
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try

        Return name
    End Function


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
