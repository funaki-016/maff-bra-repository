'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.08.20 |TSP)Nakamura        | R2 フェーズ1 要件No.4修正　
'//  REV_002   | 2023.01.10 |Daiko               | 要件No.4
'//            |            |                    |
'//*************************************************************************************************

Imports System.Data.SqlClient

''' <summary>
''' 集計結果表テーブル操作
''' </summary>
''' <remarks></remarks>
Public Class DAOSyukeiKekkahyo

    ''' <summary>
    ''' 集計結果表項目クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 集計結果表項目
        Public シート名 As String                  'シート名
        Public 行位置 As Integer                   '行位置
        Public 列位置 As Integer                   '列位置
        Public 値 As String                        '値
        Public 型区分 As String                    '型区分
        Public 表示単位 As String                  '表示単位
    End Class

    ''' <summary>主キークラス</summary>
    Public Class PrimaryKey
        ''' <summary>調査年</summary>
        Public Property chosaNen As String
        ''' <summary>集計番号</summary>
        Public Property syukeiNo As String
        ''' <summary>連番</summary>
        Public Property groupKey As String

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="chosaNen"></param>
        ''' <param name="syukeiNo"></param>
        ''' <remarks></remarks>
        Sub New(chosaNen As String, syukeiNo As String)
            _chosaNen = chosaNen
            _syukeiNo = syukeiNo
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="chosaNen"></param>
        ''' <param name="syukeiNo"></param>
        ''' <param name="groupKey"></param>
        ''' <remarks></remarks>
        Sub New(chosaNen As String, syukeiNo As String, groupKey As String)
            _chosaNen = chosaNen
            _syukeiNo = syukeiNo
            _groupKey = groupKey
        End Sub

    End Class

    ''' <summary>項目キークラス</summary>
    Public Class KomokuKey
        ''' <summary>農政局</summary>
        Public Property kyoku As String
        ''' <summary>都道府県</summary>
        Public Property jimusho As String
        ''' <summary>実査設置拠点</summary>
        Public Property kyoten As String
        ''' <summary>平均種類</summary>
        Public Property heikinSyurui As String
        ''' <summary>規模階層</summary>
        Public Property kiboKaisou As String
        ''' <summary>生産費平均値種類</summary>
        Public Property seisanhiHeikin As String
        ''' <summary>田畑区分</summary>
        Public Property tahataKbn As String
        ''' <summary>ビール麦販売区分</summary>
        Public Property beerKbn As String
        ''' <summary>てんさい栽培区分</summary>
        Public Property tensaiKbn As String
        ''' <summary>集計名称</summary>
        Public Property syukeiName As String
        ''' <summary>集計条件</summary>
        Public Property syukeiJouken As String
        ''' <summary>地域コード</summary>
        Public Property chiikiCd As String
        ''' <summary>部門コード</summary>
        Public Property bumonCd As String

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

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="kyoku"></param>
        ''' <param name="jimusho"></param>
        ''' <param name="kyoten"></param>
        ''' <param name="heikinSyurui"></param>
        ''' <param name="kiboKaisou"></param>
        ''' <param name="seisanhiHeikin"></param>
        ''' <param name="tahataKbn"></param>
        ''' <param name="beerKbn"></param>
        ''' <param name="tensaiKbn"></param>
        ''' <param name="syukeiName"></param>
        ''' <param name="syukeiJouken"></param>
        ''' <param name="chiikiCd"></param>
        ''' <remarks></remarks>
        Sub New(kyoku As String, jimusho As String, kyoten As String, heikinSyurui As String, kiboKaisou As String, seisanhiHeikin As String, tahataKbn As String, beerKbn As String, tensaiKbn As String, syukeiName As String, syukeiJouken As String, chiikiCd As String, bumonCd As String)
            _kyoku = kyoku
            _jimusho = jimusho
            _kyoten = kyoten
            _heikinSyurui = heikinSyurui
            _kiboKaisou = kiboKaisou
            _seisanhiHeikin = seisanhiHeikin
            _tahataKbn = tahataKbn
            _beerKbn = beerKbn
            _tensaiKbn = tensaiKbn
            _syukeiName = syukeiName
            _syukeiJouken = syukeiJouken
            _chiikiCd = chiikiCd
            _bumonCd = bumonCd
        End Sub

    End Class

    ''' <summary>
    ''' 集計結果表調査年取得
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
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
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
    ''' 集計結果表リスト表示取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetList(db As DBAccess, chosakubun As String, chosaNen As String, kyoku As String, jimusho As String, kyoten As String) As DataTable

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
                .AppendLine("     , 集計番号 ")
                .AppendLine("     , 集計名称")
                .AppendLine("     , 集計条件 ")
                .AppendLine("     , MAX(更新日付) AS 更新日付 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
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
                .AppendLine("GROUP BY 農政局 ")
                .AppendLine("       , 都道府県 ")
                .AppendLine("       , 実査設置拠点 ")
                .AppendLine("       , 集計番号 ")
                .AppendLine("       , 集計名称")
                .AppendLine("       , 集計条件 ")
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
    ''' 集計結果表リスト表示取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="kiboKaisou"></param>
    ''' <param name="chiiki"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetList(db As DBAccess, chosakubun As String, pKey As DAOSyukeiKekkahyo.PrimaryKey, kKey As DAOSyukeiKekkahyo.KomokuKey, Optional kiboKaisou As String = Nothing, Optional chiiki As String = Nothing) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT 連番 AS 本年連番 ")
                .AppendLine("     , 規模階層 ")
                .AppendLine("     , 生産費平均値種類 ")
                .AppendLine("     , 地域コード ")
                .AppendLine(String.Format("     , ""{0}"" AS 本年集計戸数 ", ComConst.集計結果表.集計戸数(chosakubun)))
                .AppendLine("     , 田畑区分 ")
                .AppendLine("     , ビール麦販売区分 ")
                .AppendLine("     , てんさい栽培区分 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
                .AppendLine("WHERE  調査年         = @調査年 ")
                .AppendLine("AND    集計番号       = @集計番号 ")
                .AppendLine("AND    農政局         = @農政局 ")
                .AppendLine("AND    都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                If Not kiboKaisou Is Nothing Then
                    .AppendLine("AND    規模階層       = @規模階層 ")
                End If
                If Not chiiki Is Nothing Then
                    .AppendLine("AND    地域コード     = @地域コード ")
                End If
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
            para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, pKey.syukeiNo))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))
            If Not kiboKaisou Is Nothing Then
                para.Add(db.CreateParameter("@規模階層", SqlDbType.Int, kiboKaisou))
            End If
            If Not chiiki Is Nothing Then
                para.Add(db.CreateParameter("@地域コード", SqlDbType.Int, chiiki))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計結果表リスト表示取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="pKeyThis"></param>
    ''' <param name="kKeyThis"></param>
    ''' <param name="pKeyLast"></param>
    ''' <param name="kKeyLast"></param>
    ''' <param name="kiboKaisou"></param>
    ''' <param name="chiiki"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetList(db As DBAccess, chosakubun As String, pKeyThis As DAOSyukeiKekkahyo.PrimaryKey, kKeyThis As DAOSyukeiKekkahyo.KomokuKey, pKeyLast As DAOSyukeiKekkahyo.PrimaryKey, kKeyLast As DAOSyukeiKekkahyo.KomokuKey, Optional kiboKaisou As String = Nothing, Optional chiiki As String = Nothing) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT A.連番 AS 本年連番 ")
                .AppendLine("     , A.規模階層 ")
                .AppendLine("     , A.生産費平均値種類 ")
                .AppendLine("     , A.地域コード ")
                .AppendLine("     , A.集計戸数 AS 本年集計戸数")
                .AppendLine("     , B.連番 AS 前年連番 ")
                .AppendLine("     , B.集計戸数 AS 前年集計戸数")
                .AppendLine("     , A.田畑区分 ")
                .AppendLine("     , A.ビール麦販売区分 ")
                .AppendLine("     , A.てんさい栽培区分 ")
                .AppendLine("FROM ")
                .AppendLine("  (SELECT 連番 ")
                .AppendLine("        , 規模階層 ")
                .AppendLine("        , 生産費平均値種類 ")
                .AppendLine("        , 田畑区分 ")
                .AppendLine("        , ビール麦販売区分 ")
                .AppendLine("        , てんさい栽培区分 ")
                .AppendLine("        , 農政局 ")
                .AppendLine("        , 都道府県 ")
                .AppendLine("        , 実査設置拠点 ")
                .AppendLine("        , 地域コード ")
                .AppendLine("        , 部門コード ")
                .AppendLine(String.Format("        , ""{0}"" AS 集計戸数 ", ComConst.集計結果表.集計戸数(chosakubun)))
                .AppendLine(String.Format("   FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
                .AppendLine("   WHERE  調査年         = @本年調査年 ")
                .AppendLine("   AND    集計番号       = @本年集計番号 ")
                .AppendLine("   AND    農政局         = @本年農政局 ")
                .AppendLine("   AND    都道府県       = @本年都道府県 ")
                .AppendLine("   AND    実査設置拠点   = @本年実査設置拠点 ")
                If Not kiboKaisou Is Nothing Then
                    .AppendLine("   AND    規模階層       = @規模階層 ")
                End If
                If Not chiiki Is Nothing Then
                    .AppendLine("   AND    地域コード     = @地域コード ")
                End If
                .AppendLine("  ) A ")
                .AppendLine("LEFT JOIN ")
                .AppendLine("  (SELECT 連番 ")
                .AppendLine("        , 規模階層 ")
                .AppendLine("        , 生産費平均値種類 ")
                .AppendLine("        , 田畑区分 ")
                .AppendLine("        , ビール麦販売区分 ")
                .AppendLine("        , てんさい栽培区分 ")
                .AppendLine("        , 農政局 ")
                .AppendLine("        , 都道府県 ")
                .AppendLine("        , 実査設置拠点 ")
                .AppendLine("        , 地域コード ")
                .AppendLine("        , 部門コード ")
                .AppendLine(String.Format("        , ""{0}"" AS 集計戸数 ", ComConst.集計結果表.集計戸数(chosakubun)))
                .AppendLine(String.Format("   FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
                .AppendLine("   WHERE  調査年         = @前年調査年 ")
                .AppendLine("   AND    集計番号       = @前年集計番号 ")
                .AppendLine("   AND    農政局         = @前年農政局 ")
                .AppendLine("   AND    都道府県       = @前年都道府県 ")
                .AppendLine("   AND    実査設置拠点   = @前年実査設置拠点 ")
                If Not kiboKaisou Is Nothing Then
                    .AppendLine("   AND    規模階層       = @規模階層 ")
                End If
                If Not chiiki Is Nothing Then
                    .AppendLine("   AND    地域コード     = @地域コード ")
                End If
                .AppendLine("  ) B ")
                .AppendLine("ON     A.規模階層         = B.規模階層 ")
                .AppendLine("AND    A.生産費平均値種類 = B.生産費平均値種類 ")
                .AppendLine("AND    A.田畑区分         = B.田畑区分 ")
                .AppendLine("AND    A.ビール麦販売区分 = B.ビール麦販売区分 ")
                .AppendLine("AND    A.てんさい栽培区分 = B.てんさい栽培区分 ")
                .AppendLine("AND    A.農政局           = B.農政局 ")
                .AppendLine("AND    A.都道府県         = B.都道府県 ")
                .AppendLine("AND    A.実査設置拠点     = B.実査設置拠点 ")
                .AppendLine("AND    A.地域コード       = B.地域コード ")
                .AppendLine("AND    A.部門コード       = B.部門コード ")
            End With

            para.Add(db.CreateParameter("@本年調査年", SqlDbType.Int, pKeyThis.chosaNen))
            para.Add(db.CreateParameter("@本年集計番号", SqlDbType.VarChar, pKeyThis.syukeiNo))
            para.Add(db.CreateParameter("@本年農政局", SqlDbType.Int, kKeyThis.kyoku))
            para.Add(db.CreateParameter("@本年都道府県", SqlDbType.Int, kKeyThis.jimusho))
            para.Add(db.CreateParameter("@本年実査設置拠点", SqlDbType.Int, kKeyThis.kyoten))
            para.Add(db.CreateParameter("@前年調査年", SqlDbType.Int, pKeyLast.chosaNen))
            para.Add(db.CreateParameter("@前年集計番号", SqlDbType.VarChar, pKeyLast.syukeiNo))
            para.Add(db.CreateParameter("@前年農政局", SqlDbType.Int, kKeyLast.kyoku))
            para.Add(db.CreateParameter("@前年都道府県", SqlDbType.Int, kKeyLast.jimusho))
            para.Add(db.CreateParameter("@前年実査設置拠点", SqlDbType.Int, kKeyLast.kyoten))
            If Not kiboKaisou Is Nothing Then
                para.Add(db.CreateParameter("@規模階層", SqlDbType.Int, kiboKaisou))
            End If
            If Not chiiki Is Nothing Then
                para.Add(db.CreateParameter("@地域コード", SqlDbType.Int, chiiki))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    '---REV_001_ADD START
    ''' <summary>
    ''' 集計結果表リスト表示取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetListKangen(db As DBAccess, chosakubun As String, chosaNen As String, chikiArray As String()) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT 集計番号 ")
                .AppendLine("     , 集計名称")
                .AppendLine("     , 地域コード ")
                .AppendLine("     , 規模階層 ")
                .AppendLine("     , 生産費平均値種類 ")
                .AppendLine("     , S000009 AS 営農集計戸数 ")
                .AppendLine("     , S000007 AS 生産費集計経営体数 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
                .AppendLine("WHERE  調査年         = @調査年 ")
                .AppendLine("AND    地域コード     IN ('" + String.Join("', '", chikiArray) + "')")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計結果表リスト表示取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="shukeiNo"></param>
    ''' <param name="chiikiCd"></param>
    ''' <param name="kiboKaiso"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetListKangen(db As DBAccess, chosakubun As String, chosaNen As String, shukeiNo As String, chiikiCd As String, kiboKaiso As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Dim isKomugi As Boolean = False
        Dim isNijo As Boolean = False
        Dim isTensai As Boolean = False

        Select Case chosakubun
            Case ComConst.調査区分.小麦生産費統計_個別,
                ComConst.調査区分.二条大麦生産費統計_個別,
                ComConst.調査区分.六条大麦生産費統計_個別,
                ComConst.調査区分.はだか麦生産費統計_個別,
                ComConst.調査区分.大豆生産費統計_個別,
                ComConst.調査区分.小麦生産費統計_組織法人,
                ComConst.調査区分.大豆生産費統計_組織法人,
                ComConst.調査区分.経営分析調査_二条大麦生産費,
                ComConst.調査区分.経営分析調査_六条大麦生産費,
                ComConst.調査区分.経営分析調査_はだか麦生産費
                isKomugi = True
        End Select
        Select Case chosakubun
            Case ComConst.調査区分.二条大麦生産費統計_個別,
                 ComConst.調査区分.経営分析調査_二条大麦生産費
                isNijo = True
        End Select
        Select Case chosakubun
            Case ComConst.調査区分.てんさい生産費統計_個別,
                 ComConst.調査区分.経営分析調査_てんさい生産費
                isTensai = True
        End Select


        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT 調査年 ")
                .AppendLine("     , 集計番号 ")
                .AppendLine("     , 連番 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
                .AppendLine("WHERE  調査年           = @調査年 ")
                .AppendLine("AND    集計番号         = @集計番号 ")
                .AppendLine("AND    地域コード       = @地域コード ")
                .AppendLine("AND    規模階層         = @規模階層 ")
                If isKomugi Then
                    .AppendLine("AND    田畑区分         = @田畑区分 ")
                End If
                If isNijo Then
                    .AppendLine("AND    ビール麦販売区分 = @ビール麦販売区分 ")
                End If
                If isTensai Then
                    .AppendLine("AND    てんさい栽培区分 = @てんさい栽培区分 ")
                End If
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, shukeiNo))
            para.Add(db.CreateParameter("@地域コード", SqlDbType.Int, chiikiCd))
            para.Add(db.CreateParameter("@規模階層", SqlDbType.Int, kiboKaiso))
            If isKomugi Then
                para.Add(db.CreateParameter("@田畑区分", SqlDbType.Int, ComConst.田畑区分.田畑計))
            End If
            If isNijo Then
                para.Add(db.CreateParameter("@ビール麦販売区分", SqlDbType.Int, ComConst.ビール麦販売区分.指定しない))
            End If
            If isTensai Then
                para.Add(db.CreateParameter("@てんさい栽培区分", SqlDbType.Int, ComConst.てんさい栽培区分.指定しない))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function
    '---REV_001_ADD END

    ''' <summary>
    ''' 集計結果表テーブル取得（集計番号単位）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSyukeiNoTable(db As DBAccess, chosaKubun As String, pKey As DAOSyukeiKekkahyo.PrimaryKey, kKey As DAOSyukeiKekkahyo.KomokuKey) As Dictionary(Of String, DataTable)
        Dim ret As New Dictionary(Of String, DataTable)
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.集計結果表.テーブル名称(chosaKubun)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.集計結果表.テーブル名称(chosaKubun)(0)) Then
                    ' SQL文の設定
                    With sb
                        .AppendLine("SELECT * ")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    集計番号       = @集計番号 ")
                        .AppendLine("AND    農政局         = @農政局 ")
                        .AppendLine("AND    都道府県       = @都道府県 ")
                        .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                    End With
                Else
                    ' SQL文の設定
                    With sb
                        .AppendLine("SELECT *")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    集計番号       IN (SELECT 集計番号 ")
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosaKubun)(0)))
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        .AppendLine("                          AND    集計番号       = @集計番号 ")
                        .AppendLine("                          AND    農政局         = @農政局 ")
                        .AppendLine("                          AND    都道府県       = @都道府県 ")
                        .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        .AppendLine("                         ) ")
                    End With
                End If

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, pKey.syukeiNo))
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))

                dt = db.GetDataTable(sb.ToString, para)

                ret.Add(tableName, dt)
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計結果表テーブル取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="pKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTable(db As DBAccess, chosaKubun As String, pKey As DAOSyukeiKekkahyo.PrimaryKey) As Dictionary(Of String, DataTable)
        Dim ret As New Dictionary(Of String, DataTable)
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.集計結果表.テーブル名称(chosaKubun)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT *")
                    .AppendLine(String.Format("FROM   ""{0}""", tableName))
                    .AppendLine("WHERE  調査年         = @調査年 ")
                    .AppendLine("AND    集計番号       = @集計番号 ")
                    .AppendLine("AND    連番           = @連番 ")
                End With

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, pKey.syukeiNo))
                para.Add(db.CreateParameter("@連番", SqlDbType.VarChar, pKey.groupKey))

                dt = db.GetDataTable(sb.ToString, para)

                ret.Add(tableName, dt)
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計結果表テーブル取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="pKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTableMain(db As DBAccess, chosaKubun As String, pKey As DAOSyukeiKekkahyo.PrimaryKey) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosaKubun)(0)))
                .AppendLine("WHERE  調査年         = @調査年 ")
                .AppendLine("AND    集計番号       = @集計番号 ")
                .AppendLine("AND    連番           = @連番 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
            para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, pKey.syukeiNo))
            para.Add(db.CreateParameter("@連番", SqlDbType.VarChar, pKey.groupKey))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計結果表テーブル取得（生産費平均値種類０以外）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="pKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTableAllHeikin(db As DBAccess, chosaKubun As String, pKey As DAOSyukeiKekkahyo.PrimaryKey, kkey As DAOSyukeiKekkahyo.KomokuKey) As Dictionary(Of String, DataTable)
        Dim ret As New Dictionary(Of String, DataTable)
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.集計結果表.テーブル名称(chosaKubun)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                If tableName.Equals(ComConst.集計結果表.テーブル名称(chosaKubun)(0)) Then
                    With sb
                        .AppendLine("SELECT *")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年           = @調査年 ")
                        .AppendLine("AND    集計番号         = @集計番号 ")
                        .AppendLine("AND    平均種類         = @平均種類 ")
                        .AppendLine("AND    規模階層         = @規模階層 ")
                        .AppendLine("AND    生産費平均値種類 <> 0 ")
                        .AppendLine("AND    田畑区分         = @田畑区分 ")
                        .AppendLine("AND    ビール麦販売区分 = @ビール麦販売区分 ")
                        .AppendLine("AND    てんさい栽培区分 = @てんさい栽培区分 ")
                        .AppendLine("AND    農政局           = @農政局 ")
                        .AppendLine("AND    都道府県         = @都道府県 ")
                        .AppendLine("AND    実査設置拠点     = @実査設置拠点 ")
                        .AppendLine("AND    地域コード       = @地域コード ")
                        .AppendLine("AND    部門コード       = @部門コード ")
                    End With
                Else
                    Dim colArr As List(Of String) = GetColumns(db, tableName)
                    With sb
                        .AppendLine("SELECT t1.* , 生産費平均値種類")
                        .AppendLine(String.Format("FROM   ""{0}"" as t1", tableName))
                        .AppendLine(String.Format("INNER JOIN   ""{0}"" as t2", ComConst.集計結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
                        .AppendLine("ON   t1.調査年 = t2.調査年 ")
                        .AppendLine("AND  t1.集計番号 = t2.集計番号 ")
                        .AppendLine("AND  t1.連番 = t2.連番 ")
                        .AppendLine("WHERE  t1.調査年           = @調査年 ")
                        .AppendLine("AND    t1.集計番号         = @集計番号 ")
                        .AppendLine("AND    t2.平均種類         = @平均種類 ")
                        .AppendLine("AND    t2.規模階層         = @規模階層 ")
                        .AppendLine("AND    t2.生産費平均値種類 <> 0 ")
                        .AppendLine("AND    t2.田畑区分         = @田畑区分 ")
                        .AppendLine("AND    t2.ビール麦販売区分 = @ビール麦販売区分 ")
                        .AppendLine("AND    t2.てんさい栽培区分 = @てんさい栽培区分 ")
                        .AppendLine("AND    t2.農政局           = @農政局 ")
                        .AppendLine("AND    t2.都道府県         = @都道府県 ")
                        .AppendLine("AND    t2.実査設置拠点     = @実査設置拠点 ")
                        .AppendLine("AND    t2.地域コード       = @地域コード ")
                        .AppendLine("AND    t2.部門コード       = @部門コード ")
                    End With
                End If
                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, pKey.syukeiNo))
                para.Add(db.CreateParameter("@平均種類", SqlDbType.Int, If(kkey Is Nothing, "NULL", kkey.heikinSyurui)))
                para.Add(db.CreateParameter("@規模階層", SqlDbType.Int, If(kkey Is Nothing, "NULL", kkey.kiboKaisou)))
                para.Add(db.CreateParameter("@田畑区分", SqlDbType.Int, If(kkey Is Nothing, "NULL", kkey.tahataKbn)))
                para.Add(db.CreateParameter("@ビール麦販売区分", SqlDbType.Int, If(kkey Is Nothing, "NULL", kkey.beerKbn)))
                para.Add(db.CreateParameter("@てんさい栽培区分", SqlDbType.Int, If(kkey Is Nothing, "NULL", kkey.tensaiKbn)))
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, If(kkey Is Nothing, "NULL", kkey.kyoku)))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, If(kkey Is Nothing, "NULL", kkey.jimusho)))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, If(kkey Is Nothing, "NULL", kkey.kyoten)))
                para.Add(db.CreateParameter("@地域コード", SqlDbType.Int, If(kkey Is Nothing, "NULL", kkey.chiikiCd)))
                para.Add(db.CreateParameter("@部門コード", SqlDbType.Int, If(kkey Is Nothing, "NULL", kkey.bumonCd)))

                dt = db.GetDataTable(sb.ToString, para)

                ret.Add(tableName, dt)
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function


    ''' <summary>
    ''' 集計結果表テーブル取得（任意帳票）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTableNiniChohyo(db As DBAccess, chosaKubun As String, pKey As DAOSyukeiKekkahyo.PrimaryKey, kKey As DAOSyukeiKekkahyo.KomokuKey) As Dictionary(Of String, DataTable)
        Dim ret As New Dictionary(Of String, DataTable)
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.集計結果表.テーブル名称(chosaKubun)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.集計結果表.テーブル名称(chosaKubun)(0)) Then
                    ' SQL文の設定
                    With sb
                        .AppendLine("SELECT * ")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年           = @調査年 ")
                        .AppendLine("AND    集計番号         = @集計番号 ")
                        .AppendLine("AND    農政局           = @農政局 ")
                        .AppendLine("AND    都道府県         = @都道府県 ")
                        .AppendLine("AND    実査設置拠点     = @実査設置拠点 ")
                        .AppendLine("AND    規模階層         = @規模階層 ")
                        .AppendLine("AND    生産費平均値種類 = @生産費平均値種類 ")
                        .AppendLine("AND    田畑区分         = @田畑区分 ")
                        .AppendLine("AND    ビール麦販売区分 = @ビール麦販売区分 ")
                        .AppendLine("AND    てんさい栽培区分 = @てんさい栽培区分 ")
                        .AppendLine("AND    地域コード       = @地域コード ")
                    End With
                Else
                    ' SQL文の設定
                    With sb
                        .AppendLine("SELECT *")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    集計番号       = @集計番号 ")
                        .AppendLine("AND    連番           IN (SELECT 連番 ")
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosaKubun)(0)))
                        .AppendLine("                          WHERE  調査年           = @調査年 ")
                        .AppendLine("                          AND    集計番号         = @集計番号 ")
                        .AppendLine("                          AND    農政局           = @農政局 ")
                        .AppendLine("                          AND    都道府県         = @都道府県 ")
                        .AppendLine("                          AND    実査設置拠点     = @実査設置拠点 ")
                        .AppendLine("                          AND    規模階層         = @規模階層 ")
                        .AppendLine("                          AND    生産費平均値種類 = @生産費平均値種類 ")
                        .AppendLine("                          AND    田畑区分         = @田畑区分 ")
                        .AppendLine("                          AND    ビール麦販売区分 = @ビール麦販売区分 ")
                        .AppendLine("                          AND    てんさい栽培区分 = @てんさい栽培区分 ")
                        .AppendLine("                          AND    地域コード       = @地域コード ")
                        .AppendLine("                         ) ")
                    End With
                End If

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, pKey.syukeiNo))
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))
                para.Add(db.CreateParameter("@規模階層", SqlDbType.Int, kKey.kiboKaisou))
                para.Add(db.CreateParameter("@生産費平均値種類", SqlDbType.Int, kKey.seisanhiHeikin))
                para.Add(db.CreateParameter("@田畑区分", SqlDbType.Int, kKey.tahataKbn))
                para.Add(db.CreateParameter("@ビール麦販売区分", SqlDbType.Int, kKey.beerKbn))
                para.Add(db.CreateParameter("@てんさい栽培区分", SqlDbType.Int, kKey.tensaiKbn))
                para.Add(db.CreateParameter("@地域コード", SqlDbType.Int, kKey.chiikiCd))

                dt = db.GetDataTable(sb.ToString, para)

                ret.Add(tableName, dt)
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計名称取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="pkey"></param>
    ''' <param name="kkey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSyukeiName(db As DBAccess, chosakubun As String, pkey As DAOSyukeiKekkahyo.PrimaryKey, kkey As DAOSyukeiKekkahyo.KomokuKey) As String

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT 集計名称")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
                .AppendLine("WHERE  調査年         = @調査年 ")
                .AppendLine("AND    集計番号       = @集計番号 ")
                .AppendLine("AND    農政局         = @農政局 ")
                .AppendLine("AND    都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                .AppendLine("GROUP BY 集計番号 ")
                .AppendLine("       , 集計名称")
            End With


            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pkey.chosaNen))
            para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, pkey.syukeiNo))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kkey.kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kkey.jimusho))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kkey.kyoten))

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("集計名称").ToString
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
    ''' 集計結果表データ追加
    ''' </summary>
    ''' <param name="chosakubun"></param>
    ''' <param name="vtList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertTable(db As DBAccess, chosakubun As String, vtList As List(Of ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)))) As Boolean
        Dim ret As Boolean = False
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder

        Try
            For Each tableName As String In ComConst.集計結果表.テーブル名称(chosakubun)
                '列名称取得
                Dim colArr As List(Of String) = GetColumns(db, tableName)

                sb = New System.Text.StringBuilder
                sb.AppendLine(String.Format("SELECT TOP 0 * FROM {0} ", tableName))

                dt = db.GetDataTable(sb.ToString)
                Dim row As DataRow = Nothing

                For Each vt As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)) In vtList
                    row = dt.NewRow()
                    row.BeginEdit()

                    row("調査年") = vt.Item1.chosaNen
                    row("集計番号") = vt.Item1.syukeiNo
                    row("連番") = vt.Item1.groupKey
                    If tableName = ComConst.集計結果表.テーブル名称(chosakubun)(0) Then
                        row("平均種類") = vt.Item2.heikinSyurui
                        row("規模階層") = vt.Item2.kiboKaisou
                        row("生産費平均値種類") = vt.Item2.seisanhiHeikin
                        row("田畑区分") = vt.Item2.tahataKbn
                        row("ビール麦販売区分") = vt.Item2.beerKbn
                        row("てんさい栽培区分") = vt.Item2.tensaiKbn
                        row("農政局") = vt.Item2.kyoku
                        row("都道府県") = vt.Item2.jimusho
                        row("実査設置拠点") = vt.Item2.kyoten
                        row("地域コード") = vt.Item2.chiikiCd
                        row("部門コード") = vt.Item2.bumonCd
                        row("集計名称") = vt.Item2.syukeiName
                        row("集計条件") = vt.Item2.syukeiJouken
                    End If

                    For Each col As String In colArr
                        ' REV_004↓
                        'If Not vt.Item3(col).値 Is Nothing Then
                        '    row(col) = vt.Item3(col).値
                        'Else
                        '    row(col) = DBNull.Value
                        'End If
                        If vt.Item3.ContainsKey(col) Then
                            If Not vt.Item3(col).値 Is Nothing Then
                                row(col) = vt.Item3(col).値
                            Else
                                row(col) = DBNull.Value
                            End If
                        End If
                        ' REV_004↑
                    Next
                    row("更新日付") = Now()
                    row("更新者ID") = CommonInfo.UserId
                    row.EndEdit()
                    dt.Rows.Add(row)

                    If dt.Rows.Count = 1000 Then
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
            Next
        Catch ex As Exception
            Throw
        End Try

        Return ret

    End Function

    ''' <summary>
    ''' 集計結果表データ削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kkey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteTable(db As DBAccess, chosakubun As String, pKey As DAOSyukeiKekkahyo.PrimaryKey, kkey As DAOSyukeiKekkahyo.KomokuKey) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.集計結果表.テーブル名称(chosakubun).Reverse().ToArray()
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.集計結果表.テーブル名称(chosakubun)(0)) Then
                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    集計番号       = @集計番号 ")
                        If Not pKey.groupKey Is Nothing Then
                            .AppendLine("AND    連番       = @連番 ")
                        End If
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
                        If Not pKey.groupKey Is Nothing Then
                            .AppendLine("AND    連番       = @連番 ")
                        End If
                        .AppendLine("AND    集計番号       IN (SELECT 集計番号 ")
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        .AppendLine("                          AND    集計番号       = @集計番号 ")
                        If Not pKey.groupKey Is Nothing Then
                            .AppendLine("                          AND    連番       = @連番 ")
                        End If
                        .AppendLine("                          AND    農政局         = @農政局 ")
                        .AppendLine("                          AND    都道府県       = @都道府県 ")
                        .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        .AppendLine("                         ) ")
                    End With
                End If

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, pKey.syukeiNo))
                If Not pKey.groupKey Is Nothing Then
                    para.Add(db.CreateParameter("@連番", SqlDbType.Int, pKey.groupKey))
                End If
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kkey.kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kkey.jimusho))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kkey.kyoten))
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
    ''' 集計名称更新
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="syukeiName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function UpdateSyukeiName(db As DBAccess, chosakubun As String, pKey As DAOSyukeiKekkahyo.PrimaryKey, kKey As DAOSyukeiKekkahyo.KomokuKey, syukeiName As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine(String.Format("UPDATE   ""{0}""", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
                .AppendLine("SET    集計名称       = @集計名称 ")
                If Not (CommonInfo.Kubun2 = ComConst.区分２.農産物生産費) Then
                    .AppendLine(String.Format("     , ""{0}"" = @集計名称 ", ComConst.集計結果表.集計名称(chosakubun)))
                End If
                .AppendLine("      ,更新日付       = GETDATE() ")
                .AppendLine("      ,更新者ID       = @更新者ID")
                .AppendLine("WHERE  調査年         = @調査年 ")
                .AppendLine("AND    集計番号       = @集計番号 ")
                .AppendLine("AND    農政局         = @農政局 ")
                .AppendLine("AND    都道府県       = @都道府県 ")
                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
            para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, pKey.syukeiNo))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))
            para.Add(db.CreateParameter("@集計名称", SqlDbType.VarChar, syukeiName))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                ret = True
            Else
                Throw New Exception(String.Format("{0}更新失敗", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
            End If
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

        Dim comCol As String() = {"調査年", "集計番号", "連番", "平均種類", "規模階層", "生産費平均値種類", "田畑区分", "ビール麦販売区分", "てんさい栽培区分", "農政局", "都道府県", "実査設置拠点", "地域コード", "部門コード", "集計名称", "集計条件", "更新日付", "更新者ID"}

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
    ''' 集計番号を条件にその集計番号が存在するかチェックを行う
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="syukeiNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getSyukeiNo(db As DBAccess, chosakubun As String, chosaNen As String, syukeiNo As String) As String
        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT DISTINCT 集計番号 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
                .AppendLine("WHERE  調査年         = @調査年 ")
                .AppendLine("AND    集計番号       = @集計番号 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, syukeiNo))

            dr = db.ExecuteReader(sb.ToString, para)
            If dr.Read Then
                strReturn = dr("集計番号").ToString
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
    ''' 集計結果表データ追加（還元）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kkey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertTableKangen(db As DBAccess, chosaKoutei As String, chosakubun As String, pKey As DAOSyukeiKekkahyo.PrimaryKey, kkey As DAOSyukeiKekkahyo.KomokuKey) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.集計結果表.テーブル名称(chosakubun)
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.集計結果表.テーブル名称(chosakubun)(0)) Then
                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("INSERT INTO [{0}].[dbo].[{1}] ", chosaKoutei, tableName))
                        .AppendLine("SELECT * ")
                        .AppendLine(String.Format("FROM   {0}", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    集計番号       = @集計番号 ")
                        .AppendLine("AND    農政局         = @農政局 ")
                        .AppendLine("AND    都道府県       = @都道府県 ")
                        .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                    End With
                Else
                    ' SQL文の設定
                    With sb
                        .AppendLine(String.Format("INSERT INTO [{0}].[dbo].[{1}] ", chosaKoutei, tableName))
                        .AppendLine("SELECT * ")
                        .AppendLine(String.Format("FROM   {0}", tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    集計番号       IN (SELECT 集計番号 ")
                        .AppendLine(String.Format("                          FROM   {0}", ComConst.集計結果表.テーブル名称(chosakubun)(0)))
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        .AppendLine("                          AND    集計番号       = @集計番号 ")
                        .AppendLine("                          AND    農政局         = @農政局 ")
                        .AppendLine("                          AND    都道府県       = @都道府県 ")
                        .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        .AppendLine("                         ) ")
                    End With
                End If

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, pKey.syukeiNo))
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kkey.kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kkey.jimusho))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kkey.kyoten))

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
    ''' 集計結果表データ削除（還元）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKoutei"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kkey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteTableKangen(db As DBAccess, chosaKoutei As String, chosakubun As String, pKey As DAOSyukeiKekkahyo.PrimaryKey, kkey As DAOSyukeiKekkahyo.KomokuKey) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.集計結果表.テーブル名称(chosakubun).Reverse().ToArray()
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.集計結果表.テーブル名称(chosakubun)(0)) Then
                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        .AppendLine(String.Format("FROM   [{0}].[dbo].[{1}]", chosaKoutei, tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    集計番号       = @集計番号 ")
                        .AppendLine("AND    農政局         = @農政局 ")
                        .AppendLine("AND    都道府県       = @都道府県 ")
                        .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                    End With
                Else
                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        .AppendLine(String.Format("FROM   [{0}].[dbo].[{1}]", chosaKoutei, tableName))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    集計番号       IN (SELECT 集計番号 ")
                        .AppendLine(String.Format("                          FROM   [{0}].[dbo].[{1}]", chosaKoutei, ComConst.集計結果表.テーブル名称(chosakubun)(0)))
                        .AppendLine("                          WHERE  調査年         = @調査年 ")
                        .AppendLine("                          AND    集計番号       = @集計番号 ")
                        .AppendLine("                          AND    農政局         = @農政局 ")
                        .AppendLine("                          AND    都道府県       = @都道府県 ")
                        .AppendLine("                          AND    実査設置拠点   = @実査設置拠点 ")
                        .AppendLine("                         ) ")
                    End With
                End If

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, pKey.syukeiNo))
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kkey.kyoku))
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kkey.jimusho))
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kkey.kyoten))

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
End Class
