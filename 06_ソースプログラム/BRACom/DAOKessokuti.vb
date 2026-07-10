Imports System.Data.SqlClient

''' <summary>
''' 個別結果表テーブル操作
''' </summary>
''' <remarks></remarks>
Public Class DAOKobetsuKekkahyo

    ''' <summary>
    ''' 個別結果表項目クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 個別結果表項目
        Public シート名 As String                  'シート名
        Public 行位置 As Integer                   '行位置
        Public 列位置 As Integer                   '列位置
        Public 値 As String                        '値
        Public 型区分 As String                    '型区分
        Public 有効桁数 As Integer                 '有効桁数
        Public 小数点以下桁数 As Integer           '小数点以下桁数
        Public 表示単位 As String                  '表示単位
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
    Public Class KyotenKey
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
    ''' 個別結果表集計条件クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SyukeiInfo
        ''' <summary>集計番号</summary>
        Public syukeiNo As String
        ''' <summary>連番</summary>
        Public groupKey As Integer
        ''' <summary>集計名称</summary>
        Public syukeiName As String
        ''' <summary>平均種類</summary>
        Public heikinSyurui As String
        ''' <summary>規模階層</summary>
        Public kiboKaisou As String
        ''' <summary>集計パターン</summary>
        Public syukeipattern As String
        ''' <summary>集計１（生産費は集計区分）</summary>
        Public syukei1 As String
        ''' <summary>集計２（生産費は田畑区分）</summary>
        Public syukei2 As String
        ''' <summary>集計３（生産費はビール麦販売区分）</summary>
        Public syukei3 As String
        ''' <summary>集計４（生産費はてんさい栽培区分）</summary>
        Public syukei4 As String
        ''' <summary>個別結果表抽出条件(集計１～４の組み合わせ) 営農のみ使用</summary>
        Public einouSyukeiJouken As String
        ''' <summary>地域</summary>
        Public chiiki As String()
        ''' <summary>部門</summary>
        Public bumon As String
        ''' <summary>任意集計条件</summary>
        Public niniSyukei As String
        ''' <summary>生産費平均値種類（営農は0）</summary>
        Public seisanhiHeikin As String
    End Class

    ''' <summary>
    ''' 営農経営体抽出条件クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EinouKeieitaiChusyutsu
        ''' <summary>集計コード</summary>
        Public Property syukeiCd As String
        ''' <summary>集計１</summary>
        Public Property syukei1 As String
        ''' <summary>集計２</summary>
        Public Property syukei2 As String
        ''' <summary>集計３</summary>
        Public Property syukei3 As String
        ''' <summary>集計４</summary>
        Public Property syukei4 As String
        ''' <summary>抽出条件</summary>
        Public Property jouken As String
        ''' <summary>階層コード</summary>
        Public Property kaisouCd As String
        ''' <summary>部門</summary>
        Public Property bumonCd As String

        Sub New(syukeiCd As String, syukei1 As String, syukei2 As String, syukei3 As String, syukei4 As String, jouken As String, kaisouCd As String, bumonCd As String)
            _syukei1 = syukei1
            _syukei2 = syukei2
            _syukei3 = syukei3
            _syukei4 = syukei4
            _jouken = jouken
            _kaisouCd = kaisouCd
            _bumonCd = bumonCd
        End Sub
    End Class

    ''' <summary>
    ''' 生産費抽出条件クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SeisanhiChusyutsu
        ''' <summary>集計区分</summary>
        Public Property syukeiKbn As String
        ''' <summary>田畑区分</summary>
        Public Property tahataKbn As String
        ''' <summary>ビール麦販売区分</summary>
        Public Property beerKbn As String
        ''' <summary>てんさい栽培区分</summary>
        Public Property tensaiKbn As String
        ''' <summary>抽出条件</summary>
        Public Property jouken As String

        Sub New(syukeiKbn As String, tahataKbn As String, beerKbn As String, tensaiKbn As String, jouken As String)
            _syukeiKbn = syukeiKbn
            _tahataKbn = tahataKbn
            _beerKbn = beerKbn
            _tensaiKbn = tensaiKbn
            _jouken = jouken
        End Sub
    End Class


    ''' <summary>
    ''' 規模階層クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Kibokaisou
        ''' <summary>規模階層</summary>
        Public Property kaisouNo As String
        ''' <summary>階層データキー（個別結果表の項番）</summary>
        Public Property jouken As String
        ''' <summary>上限</summary>
        Public Property max As String
        ''' <summary>下限</summary>
        Public Property min As String

        Sub New(kaisouNo As String, jouken As String, max As String, min As String)
            _kaisouNo = kaisouNo
            _jouken = jouken
            _max = max
            _min = min
        End Sub
    End Class

    ''' <summary>
    ''' 個別結果表調査年取得
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
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
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

    ''' <summary>
    ''' 個別結果表調査年取得（農業経営体対応）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosaNen(db As DBAccess, chosakubun As String, kyoku As String, jimusho As String, kyoten As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT DISTINCT 調査年 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosakubun)(0)))
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

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 個別結果表調査年取得（受信）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="upLow"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosaNenJushin(db As DBAccess, upLow As String, kyoku As String, jimusho As String, kyoten As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT DISTINCT 調査年 ")
                .AppendLine(String.Format("FROM   ""{0}""", "受信＿" & ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine("WHERE  上り下り区分     = @上り下り区分 ")
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

            para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, upLow))
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

    ''' <summary>
    ''' 個別結果表リスト表示取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="einouRuikei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetList(db As DBAccess, chosaNen As String, kyoku As String, jimusho As String, kyoten As String, einouRuikei As String) As DataTable

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
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
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
                    .AppendLine(String.Format("AND    ""{0}""   = @営農類型 ", ComConst.個別結果表.営農類型(CommonInfo.Chosakubun)))
                End If
                If CommonInfo.SenmonChosain Then
                    .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                    .AppendLine("                          FROM   専門調査員担当調査客体 ")
                    .AppendLine("                          WHERE  ユーザーID = @ユーザーID ")
                    .AppendLine("                         ) ")
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
                para.Add(db.CreateParameter("@営農類型", SqlDbType.Decimal, einouRuikei))
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
    ''' センサス番号取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCensusNo(db As DBAccess, chosakubun As String, chosaNen As String, kyoku As String, jimusho As String, kyoten As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT センサス番号 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosakubun)(0) & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(chosakubun)))
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
    ''' 個別結果表データ存在チェック
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CheckExist(db As DBAccess, pKey As DAOKobetsuKekkahyo.PrimaryKey, kKey As DAOKobetsuKekkahyo.KyotenKey) As Boolean
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
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine("WHERE  調査年         = @調査年 ")
                .AppendLine("AND    センサス番号   = @センサス番号 ")
                .AppendLine("AND    農政局         = @農政局 ")
                .AppendLine("AND    都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))

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
    ''' 個別結果表テーブル取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTable(db As DBAccess, pKey As DAOKobetsuKekkahyo.PrimaryKey) As Dictionary(Of String, DataTable)
        Dim ret As New Dictionary(Of String, DataTable)
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)

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
    ''' 個別結果表データ削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteTable(db As DBAccess, pKey As DAOKobetsuKekkahyo.PrimaryKey, kKey As DAOKobetsuKekkahyo.KyotenKey) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun).Reverse().ToArray()
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)) Then
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
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
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

    ''' <summary>
    ''' 個別結果表データ追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertTable(db As DBAccess, pKey As DAOKobetsuKekkahyo.PrimaryKey, kKey As DAOKobetsuKekkahyo.KyotenKey, dc As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)
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
                    If tableName = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0) Then
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
                    If tableName = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0) Then
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
                If tableName = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0) Then
                    para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                    para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                    para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))
                End If
                For Each col As String In colArr
                    para.Add(db.CreateParameter(String.Format("@{0}", col), If(dc(col).型区分 = ComConst.型区分.数値型, SqlDbType.Decimal, SqlDbType.VarChar), dc(col).値))
                Next
                para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

                If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
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
    ''' 個別結果表データ追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertTable(db As DBAccess, dc As Dictionary(Of String, String)) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)
                '列名称取得
                Dim colArr As List(Of String) = GetColumnsJushin(db, tableName, True)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)) Then
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
                        .AppendLine(String.Format("                          FROM   受信＿{0}", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
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
    ''' 個別結果表データ追加（受信）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="dc"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertTableJushin(db As DBAccess, dc As Dictionary(Of String, String), chosaKoutei As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)
                '列名称取得
                Dim colArr As List(Of String) = GetColumnsJushin(db, tableName)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)) Then
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
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
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
    ''' 個別結果表データ削除（受信）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="dc"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteTableJushin(db As DBAccess, dc As Dictionary(Of String, String), Optional chosaKoutei As String = Nothing) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun).Reverse().ToArray()
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)) Then
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
                            .AppendLine(String.Format("                          FROM   受信＿{0}", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
                        Else
                            .AppendLine(String.Format("                          FROM   [{0}].[dbo].[受信＿{1}]", chosaKoutei, ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
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
    ''' 個別結果表テーブル取得（バックアップ）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetBackUpTable(db As DBAccess, chosaKubun As String, chosaNen As String, kyoku As String, jimusho As String, kyoten As String) As Dictionary(Of String, DataTable)
        Dim ret As New Dictionary(Of String, DataTable)
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(chosaKubun)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.個別結果表.テーブル名称(chosaKubun)(0)) Then
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
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosaKubun)(0)))
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
    ''' 個別結果表データ削除（レストア）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteRestoreTable(db As DBAccess, chosaKubun As String, chosaNen As String, kyoku As String, jimusho As String, kyoten As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(chosaKubun).Reverse().ToArray()
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.個別結果表.テーブル名称(chosaKubun)(0)) Then
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
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosaKubun)(0)))
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
    ''' 個別結果表テーブル取得（集計用）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetKobetsuTable(db As DBAccess, _
                                           chosakubun As String, _
                                           pKey As DAOKobetsuKekkahyo.PrimaryKey, _
                                           kKey As DAOKobetsuKekkahyo.KyotenKey, _
                                           pSyukeiInfo As SyukeiInfo, _
                                           pKibokaisou As Kibokaisou, _
                                           pEinouKeieitaiChusyutsu As EinouKeieitaiChusyutsu, _
                                           pSeisanhiChusyutsu As SeisanhiChusyutsu, _
                                           Optional pNougyouKeieitaiJouken As String = Nothing) As Dictionary(Of String, DataTable)

        Dim ret As New Dictionary(Of String, DataTable)
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Dim censusNoList As New List(Of String)
            Dim censusNoTable As DataTable = getCensusNoUseSyukei(db, chosakubun, pKey, kKey, pSyukeiInfo, pKibokaisou, pEinouKeieitaiChusyutsu, pSeisanhiChusyutsu, pNougyouKeieitaiJouken)

            For Each dr As DataRow In censusNoTable.Rows
                censusNoList.Add(dr.Item("センサス番号").ToString)
            Next


            If censusNoList.Count > 0 Then
                Dim chosaNen As String = censusNoTable.Rows(0).Item("調査年").ToString

                For Each tableName As String In ComConst.個別結果表.テーブル名称(chosakubun)

                    sb = New System.Text.StringBuilder
                    para = New List(Of DBAccess.Parameter)

                    ' SQL文の設定
                    With sb
                        .AppendLine("SELECT *")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号  IN(")
                        For i As Integer = 0 To censusNoList.Count - 1
                            If i = 0 Then
                                .Append(String.Format("'{0}'", censusNoList(i)))
                            Else
                                .Append(String.Format(", '{0}'", censusNoList(i)))
                            End If
                        Next
                        .AppendLine(" )")
                    End With

                    para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
                    dt = db.GetDataTable(sb.ToString, para)

                    ret.Add(tableName, dt)
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 個別結果表取得SQLが実行できるか検証する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="param"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function isExecutableSQL(db As DBAccess, param As String) As Boolean
        Dim ret As Boolean = True
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing

        Try

            Try
                sb = New System.Text.StringBuilder
                ' SQL文の設定
                With sb
                    .AppendLine("SELECT 調査年, センサス番号")
                    .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
                    .AppendLine(String.Format("WHERE    センサス番号   IN (SELECT ""{0}"".センサス番号 ", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
                    .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
                    For i As Integer = 1 To ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun).Count - 1
                        .AppendLine(String.Format("                          INNER JOIN   ""{0}""", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(i)))
                        .AppendLine(String.Format("                          ON   ""{0}"".センサス番号 = ""{1}"".センサス番号 ", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(i - 1), ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(i)))
                    Next
                    .AppendLine(String.Format("                          WHERE  {0}", param))
                    .Append(String.Format(" )"))
                End With

                dt = db.GetDataTable(sb.ToString)


            Catch ex As Exception
                ret = False
            End Try

        Catch ex As Exception
            Throw ex
        End Try

        Return ret

    End Function

    ''' <summary>
    ''' 集計対象となる個別結果表のセンサス番号を取得する。（個別結果表テーブルを連結して条件を指定する）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="pSyukeiInfo"></param>
    ''' <param name="pEinouKeieitaiChusyutsu"></param>
    ''' <param name="pKibokaisou"></param>
    ''' <param name="pNougyouKeieitaiJouken"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getCensusNoUseSyukei(db As DBAccess, _
                                           chosakubun As String, _
                                           pKey As DAOKobetsuKekkahyo.PrimaryKey, _
                                           kKey As DAOKobetsuKekkahyo.KyotenKey, _
                                           pSyukeiInfo As SyukeiInfo, _
                                           pKibokaisou As Kibokaisou, _
                                           pEinouKeieitaiChusyutsu As EinouKeieitaiChusyutsu, _
                                           pSeisanhiChusyutsu As SeisanhiChusyutsu, _
                                           pNougyouKeieitaiJouken As String) As DataTable


        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim censusNoList As New List(Of String)

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine(String.Format("SELECT  ""{0}"".調査年, ""{0}"".センサス番号 ", ComConst.個別結果表.テーブル名称(chosakubun)(0)))
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosakubun)(0)))
                For i As Integer = 1 To ComConst.個別結果表.テーブル名称(chosakubun).Count - 1
                    .AppendLine(String.Format("                          INNER JOIN   ""{0}""", ComConst.個別結果表.テーブル名称(chosakubun)(i)))
                    .AppendLine(String.Format("                          ON   ""{0}"".センサス番号 = ""{1}"".センサス番号 ", ComConst.個別結果表.テーブル名称(chosakubun)(i - 1), ComConst.個別結果表.テーブル名称(chosakubun)(i)))
                Next
                .AppendLine(String.Format("                          WHERE  ""{0}"".調査年         = @調査年 ", ComConst.個別結果表.テーブル名称(chosakubun)(0)))
                If Not kKey Is Nothing Then
                    .AppendLine(String.Format("                          AND    ""{0}"".農政局         = @農政局 ", ComConst.個別結果表.テーブル名称(chosakubun)(0)))
                End If

                '個別結果表抽出条件：地域
                If pSyukeiInfo.chiiki.Length > 0 Then
                    .Append("AND    ")
                    .Append(ComConst.個別結果表.都道府県)
                    .Append("    In( ")
                    For i As Integer = 0 To pSyukeiInfo.chiiki.Length - 1
                        If i = 0 Then
                            .Append(String.Format("'{0}'", pSyukeiInfo.chiiki(i)))
                        Else
                            .Append(String.Format(", '{0}'", pSyukeiInfo.chiiki(i)))
                        End If
                    Next
                    .AppendLine(")")
                End If

                .Append("                          AND ")

                '個別結果表抽出条件：集計倍率＞0
                If chosakubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) AndAlso (Not pEinouKeieitaiChusyutsu Is Nothing) AndAlso (Not pEinouKeieitaiChusyutsu.bumonCd Is Nothing) Then
                    .Append(ComConst.個別結果表.部門集計倍率(chosakubun))
                Else
                    .Append(ComConst.個別結果表.集計倍率(chosakubun))
                End If
                .AppendLine(" > 0")

                If (Not String.IsNullOrEmpty(pSyukeiInfo.niniSyukei)) Or (Not pKibokaisou Is Nothing) Or (Not pEinouKeieitaiChusyutsu Is Nothing) Or (Not pSeisanhiChusyutsu Is Nothing) Then

                    '個別結果表抽出条件：任意集計
                    If Not String.IsNullOrEmpty(pSyukeiInfo.niniSyukei) Then
                        .Append("                          AND ")
                        .AppendLine(pSyukeiInfo.niniSyukei)
                    End If

                    '個別結果表抽出条件：営農経営体抽出条件
                    If (Not pEinouKeieitaiChusyutsu Is Nothing) AndAlso (Not pEinouKeieitaiChusyutsu.jouken Is Nothing) Then
                        .Append("                          AND ")
                        .AppendLine(If(pNougyouKeieitaiJouken Is Nothing, pEinouKeieitaiChusyutsu.jouken, pNougyouKeieitaiJouken))
                    End If

                    '個別結果表抽出条件：規模階層
                    If (Not pKibokaisou.jouken Is Nothing) AndAlso (Not pKibokaisou.jouken Is Nothing) Then
                        If (Not pKibokaisou.min Is Nothing) And (Not pKibokaisou.max Is Nothing) Then
                            '上限下限が設定されている通常の規模階層
                            .Append("                          AND ")
                            .Append(pKibokaisou.jouken)
                            .Append(" >= ")
                            .AppendLine(pKibokaisou.min)
                            .Append("                          AND ")
                            .Append(pKibokaisou.max)
                            .Append(" >= ")
                            .AppendLine(pKibokaisou.jouken)
                        Else
                            '階層に集計コードが指定されている規模階層
                            .Append("                          AND ")
                            .AppendLine(pKibokaisou.jouken)
                        End If
                    End If

                    '個別結果表抽出条件：生産費
                    If (Not pSeisanhiChusyutsu Is Nothing) Then
                        .Append("                          AND ")
                        .AppendLine(pSeisanhiChusyutsu.jouken)
                    End If
                End If

            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
            If Not kKey Is Nothing Then
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
            End If

            dt = db.GetDataTable(sb.ToString, para)


        Catch ex As Exception
            Throw ex
        End Try

        Return dt

    End Function


End Class
