Imports System.Data.SqlClient

''' <summary>
''' 個別結果表テーブル操作
''' </summary>
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.07.16 |TSP                 | フェーズ1 要件No.4修正　
'//  REV_002   | 2020.11.10 |TSP                 | フェーズ3 要件No.6修正　,不具合対応No15 対応
'//  REV_003   | 2022.01.15 |日本ｺﾝﾋﾟｭｰﾀｼｽﾃﾑ     | 要件No.2対応
'//  REV_004   | 2022.10.11 |Daiko               | 要件No.1
'//  REV_005   | 2022.12.19 |Daiko               | 要件No.4
'//  REV_006   | 2023.01.13 |Daiko               | 要件No.3
'//  REV_006   | 2023.01.13 |Daiko               | 要件No.8
'//  REV_007   | 2025.09.11 |GCU                 | 要件No.8
'//  REV_008   | 2025.08.24 |GCU                 | 要件No.2 継続区分追加
'//            |            |                    |
'//*************************************************************************************************
''' <remarks> </remarks>
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
        Public 再計算区分 As String                '再計算区分
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
        ''' <summary>任意集計条件（指標部用）</summary>
        Public niniSyukeiShihyou As String
        ''' <summary>生産費平均値種類（営農は0）</summary>
        Public seisanhiHeikin As String
        'REV_006↓
        ''' <summary>基本・詳細項目集計</summary>
        Public kihonSyosaiSyukei As String
        'REV_006↑
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
        ''' <summary>営農類型</summary>
        Public Property einouruikei As String

        Sub New(syukeiCd As String, syukei1 As String, syukei2 As String, syukei3 As String, syukei4 As String, jouken As String, kaisouCd As String, bumonCd As String, einouruikei As String)
            'REV_006↓
            _syukeiCd = syukeiCd
            'REV_006↑
            _syukei1 = syukei1
            _syukei2 = syukei2
            _syukei3 = syukei3
            _syukei4 = syukei4
            _jouken = jouken
            _kaisouCd = kaisouCd
            _bumonCd = bumonCd
            _einouruikei = einouruikei
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
        'REV_006↓
        ''' <summary>集計コード</summary>
        Public Property syukeiCd As String
        'REV_006↑

        Sub New(kaisouNo As String, jouken As String, max As String, min As String)
            Me.New(kaisouNo, jouken, max, min, Nothing)
        End Sub

        Sub New(kaisouNo As String, jouken As String, max As String, min As String, syukeiCd As String)
            _kaisouNo = kaisouNo
            _jouken = jouken
            _max = max
            _min = min
            _syukeiCd = syukeiCd
        End Sub
    End Class

    ''' <summary>
    ''' 個別結果表調査年取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="kessokuchiHokana"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosaNen(db As DBAccess, kyoku As String, jimusho As String, kyoten As String, Optional kessokuchiHokana As String = ComConst.欠測値補完.無) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT DISTINCT 調査年 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0) & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokana)))
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
    ''' <param name="kessokuchiHokana"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosaNen(db As DBAccess, chosakubun As String, kyoku As String, jimusho As String, kyoten As String, kessokuchiHokana As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT DISTINCT 調査年 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosakubun)(0) & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(chosakubun, kessokuchiHokana)))
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
    ''' <param name="taishakuTaishohyo"></param>
    ''' <param name="kessokuchiHokan"></param>
    ''' <param name="chosahyo">Added by REV_006</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetList(db As DBAccess, chosaNen As String, kyoku As String, jimusho As String, kyoten As String, einouRuikei As String, Optional taishakuTaishohyo As String = Nothing, Optional kessokuchiHokan As String = ComConst.欠測値補完.無, Optional chosahyo As String = "") As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            Dim tblK As String = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0) &
                     ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan)
            Dim tblC As String = ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)

            With sb
                .AppendLine("SELECT K.農政局 ")
                .AppendLine("     , K.都道府県 ")
                .AppendLine("     , K.実査設置拠点 ")
                .AppendLine("     , K.センサス番号 ")
                .AppendLine("     , K.更新日付 ")
                'REV_007
                .AppendLine("     , C.Q00000701 ")
                .AppendLine("     , C.Q00000801 ")
                .AppendLine(String.Format("FROM   ""{0}"" K ", tblK))
                .AppendLine(String.Format("LEFT JOIN ""{0}"" C ", tblC))
                .AppendLine("  ON  C.調査年       = K.調査年 ")
                .AppendLine("  AND C.農政局       = K.農政局 ")
                .AppendLine("  AND C.都道府県     = K.都道府県 ")
                .AppendLine("  AND C.実査設置拠点 = K.実査設置拠点 ")
                .AppendLine("  AND C.センサス番号 = K.センサス番号 ")
                .AppendLine("WHERE  K.調査年      = @調査年 ")
                'REV_007

                If Not kyoku Is Nothing Then
                    .AppendLine("AND    K.農政局       = @農政局 ")
                End If
                If Not jimusho Is Nothing Then
                    .AppendLine("AND    K.都道府県     = @都道府県 ")
                End If
                If Not kyoten Is Nothing Then
                    .AppendLine("AND    K.実査設置拠点 = @実査設置拠点 ")
                End If
                If Not einouRuikei Is Nothing Then
                    .AppendLine(String.Format("AND    K.""{0}"" = @営農類型 ", ComConst.個別結果表.営農類型(CommonInfo.Chosakubun)))
                End If
                If Not taishakuTaishohyo Is Nothing Then
                    .AppendLine(String.Format("AND    K.""{0}"" = @貸借対照表 ", ComConst.個別結果表.貸借対照表(CommonInfo.Chosakubun)))
                End If
                If CommonInfo.SenmonChosain Then
                    .AppendLine("AND    K.センサス番号 IN (")
                    .AppendLine("          SELECT センサス番号 ")
                    .AppendLine("          FROM   専門調査員担当調査客体 ")
                    .AppendLine("          WHERE  ユーザーID = @ユーザーID ")
                    .AppendLine("       ) ")
                End If
                'REV_006↓
                If Not String.IsNullOrEmpty(chosahyo) Then
                    .AppendLine(" AND   K.K000026 = @調査票区分 ")
                End If
                'REV_006↑
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
            If Not taishakuTaishohyo Is Nothing Then
                para.Add(db.CreateParameter("@貸借対照表", SqlDbType.Decimal, taishakuTaishohyo))
            End If
            If CommonInfo.SenmonChosain Then
                para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, CommonInfo.UserId))
            End If
            'REV_006↓
            If Not String.IsNullOrEmpty(chosahyo) Then
                para.Add(db.CreateParameter("@調査票区分", SqlDbType.Decimal, chosahyo))
            End If
            'REV_006↑

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    '---REV001_ADD START
    ''' <summary>
    ''' 個別結果表リスト(還元資料作成画面)表示取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="einouRuikei"></param>
    ''' <param name="taishakuTaishohyo"></param>
    ''' <param name="kessokuchiHokan"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetKangenList(db As DBAccess, chosaNen As String, kyoku As String, jimusho As String, kyoten As String, einouRuikei As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT DISTINCT K.農政局 ")
                .AppendLine("     , K.都道府県 ")
                .AppendLine("     , K.実査設置拠点 ")
                .AppendLine("     , K.センサス番号 ")
                .AppendLine("     , JoinSK.氏名 AS 専門調査員名 ")
                .AppendLine("     , KS.還元資料作成日時 ")
                .AppendLine(String.Format("FROM   ""{0}"" AS K", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine("LEFT JOIN ""還元資料作成管理"" AS KS ")
                .AppendLine("ON K.調査年 = KS.調査年")
                .AppendLine("AND K.センサス番号 = KS.センサス番号")
                .AppendLine("AND KS.調査区分 = @調査区分 ")
                .AppendLine("LEFT JOIN ( ")
                .AppendLine("           SELECT SK.ユーザーID ")
                .AppendLine("                 ,SK.氏名")
                .AppendLine("                 ,STK.センサス番号")
                .AppendLine("           FROM ""専門調査員管理"" AS SK ")
                .AppendLine("           INNER JOIN ""専門調査員担当調査客体"" AS STK ")
                .AppendLine("               ON SK.都道府県 = STK.都道府県 ")
                .AppendLine("               AND SK.実査設置拠点 =　STK.実査設置拠点 ")
                .AppendLine("               AND SK.ユーザーID = STK.ユーザーID ")
                .AppendLine("           WHERE SK.都道府県 = @都道府県 ")
                .AppendLine("               AND SK.実査設置拠点 = @実査設置拠点 ")
                .AppendLine(") JoinSK ")
                .AppendLine("ON K.センサス番号 = JoinSK.センサス番号  ")
                .AppendLine("WHERE K.調査年         = @調査年 ")
                .AppendLine("AND    K.農政局       = @農政局 ")
                .AppendLine("AND    K.都道府県       = @都道府県 ")
                .AppendLine("AND    K.実査設置拠点   = @実査設置拠点 ")

                If Not einouRuikei Is Nothing Then
                    .AppendLine(String.Format("AND    ""{0}""   = @営農類型 ", ComConst.個別結果表.営農類型(CommonInfo.Chosakubun)))
                End If
                If CommonInfo.SenmonChosain Then
                    .AppendLine("AND    JoinSK.ユーザーID = @ユーザーID")
                End If
                .AppendLine("ORDER BY K.センサス番号 ASC")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
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
    ''' 個別結果表データ存在チェック(専門調査員含む)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CheckExistSenmon(db As DBAccess, pKey As DAOKobetsuKekkahyo.PrimaryKey, kKey As DAOKobetsuKekkahyo.KyotenKey) As Boolean
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
                If CommonInfo.SenmonChosain Then
                    .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                    .AppendLine("                          FROM   専門調査員担当調査客体 ")
                    .AppendLine("                          WHERE  ユーザーID = @ユーザーID ")
                    .AppendLine("                            AND    都道府県       = @都道府県 ")
                    .AppendLine("                            AND    実査設置拠点   = @実査設置拠点 ")
                    .AppendLine("                         ) ")
                End If
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))
            If CommonInfo.SenmonChosain Then
                para.Add(db.CreateParameter("@ユーザーID", SqlDbType.VarChar, CommonInfo.UserId))
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

    '---REV001_ADD END

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
    Public Shared Function GetCensusNo(db As DBAccess, chosakubun As String, chosaNen As String, kyoku As String, jimusho As String, kyoten As String, Optional kessokuchiHokan As String = ComConst.欠測値補完.有) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT センサス番号 ")
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosakubun)(0) & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(chosakubun, kessokuchiHokan)))
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
    ''' 個別結果表データ存在チェック（労働時間整理ファイル用）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="chosakbn"></param>
    ''' <returns></returns>
    ''' <remarks>REV_003 ADD</remarks>
    Public Shared Function CheckExistRoudou(db As DBAccess, pKey As DAOKobetsuKekkahyo.PrimaryKey, kKey As DAOKobetsuKekkahyo.KyotenKey, chosakbn As String) As Boolean
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
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosakbn)(0)))
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
    ''' <param name="kessokuchiHokan"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTable(db As DBAccess, pKey As DAOKobetsuKekkahyo.PrimaryKey, Optional kessokuchiHokan As String = ComConst.欠測値補完.無) As Dictionary(Of String, DataTable)
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
                    .AppendLine(String.Format("FROM   ""{0}""", tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan)))
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
    ''' 個別結果表テーブル取得（該当のみ）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="itemName"></param>
    ''' <param name="kessokuchiHokan"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTable(db As DBAccess, pKey As DAOKobetsuKekkahyo.PrimaryKey, itemName As String, Optional kessokuchiHokan As String = ComConst.欠測値補完.無) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            Dim idx As Integer = CInt(itemName.Substring(2, 1))
            Dim tableName As String = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(If(idx > 0, idx - 1, idx))

            ' SQL文の設定
            With sb
                .AppendLine((String.Format("SELECT {0}", itemName)))
                .AppendLine(String.Format("FROM   ""{0}""", tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan)))
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

    ''' <summary>
    ''' 個別結果表データ削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="kessokuchiHokan"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteTable(db As DBAccess, pKey As DAOKobetsuKekkahyo.PrimaryKey, kKey As DAOKobetsuKekkahyo.KyotenKey, Optional kessokuchiHokan As String = ComConst.欠測値補完.無) As Boolean
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
                        .AppendLine(String.Format("FROM   ""{0}""", tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan)))
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
                        .AppendLine(String.Format("FROM   ""{0}""", tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan)))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0) & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan)))
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
                    Throw New Exception(String.Format("{0}削除失敗", tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan)))
                End If
            Next
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 個別結果表データ削除（労働時間整理ファイル用）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="chosakbn"></param>
    ''' <param name="kessokuchiHokan"></param>
    ''' <returns></returns>
    ''' <remarks>REV_003 ADD</remarks>
    Public Shared Function DeleteTable_Roudou(db As DBAccess, pKey As DAOKobetsuKekkahyo.PrimaryKey, kKey As DAOKobetsuKekkahyo.KyotenKey, chosakbn As String, Optional kessokuchiHokan As String = ComConst.欠測値補完.無) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(chosakbn).Reverse().ToArray()
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                If tableName.Equals(ComConst.個別結果表.テーブル名称(chosakbn)(0)) Then
                    ' SQL文の設定
                    With sb
                        .AppendLine("DELETE ")
                        .AppendLine(String.Format("FROM   ""{0}""", tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(chosakbn, kessokuchiHokan)))
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
                        .AppendLine(String.Format("FROM   ""{0}""", tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(chosakbn, kessokuchiHokan)))
                        .AppendLine("WHERE  調査年         = @調査年 ")
                        .AppendLine("AND    センサス番号   IN (SELECT センサス番号 ")
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosakbn)(0) & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(chosakbn, kessokuchiHokan)))
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
                    Throw New Exception(String.Format("{0}削除失敗", tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(chosakbn, kessokuchiHokan)))
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
    ''' <param name="kessokuchiHokan"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertTable(db As DBAccess, pKey As DAOKobetsuKekkahyo.PrimaryKey, kKey As DAOKobetsuKekkahyo.KyotenKey, dc As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目), Optional kessokuchiHokan As String = ComConst.欠測値補完.無) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)
                '列名称取得
                Dim colArr As List(Of String) = GetColumns(db, tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan))

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine(String.Format("INSERT INTO ""{0}"" ", tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan)))
                    .AppendLine("( ")
                    .AppendLine("   調査年 ")
                    .AppendLine("  ,センサス番号 ")
                    If tableName = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0) Then
                        .AppendLine("  ,農政局 ")
                        .AppendLine("  ,都道府県 ")
                        .AppendLine("  ,実査設置拠点 ")
                    End If
                    For Each col As String In colArr
                        ' REV_004↓
                        '.AppendLine(String.Format("  ,{0} ", col))
                        If dc.ContainsKey(col) Then
                            .AppendLine(String.Format("  ,{0} ", col))
                        End If
                        ' REV_004↑
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
                        ' REV_004↓
                        '.AppendLine(String.Format("  ,@{0} ", col))
                        If dc.ContainsKey(col) Then
                            .AppendLine(String.Format("  ,@{0} ", col))
                        End If
                        ' REV_004↑
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
                    ' REV_004↓
                    'para.Add(db.CreateParameter(String.Format("@{0}", col), If(dc(col).型区分 = ComConst.型区分.数値型, SqlDbType.Decimal, SqlDbType.VarChar), dc(col).値))
                    If dc.ContainsKey(col) Then
                        para.Add(db.CreateParameter(String.Format("@{0}", col), If(dc(col).型区分 = ComConst.型区分.数値型, SqlDbType.Decimal, SqlDbType.VarChar), dc(col).値))
                    End If
                    ' REV_004↑
                Next
                para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

                If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                    ret = True
                Else
                    Throw New Exception(String.Format("{0}追加失敗", tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan)))
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
    Public Shared Function GetKobetsuTable(db As DBAccess,
                                           chosakubun As String,
                                           pKey As DAOKobetsuKekkahyo.PrimaryKey,
                                           kKey As DAOKobetsuKekkahyo.KyotenKey,
                                           pSyukeiInfo As SyukeiInfo,
                                           pKibokaisou As Kibokaisou,
                                           pEinouKeieitaiChusyutsu As EinouKeieitaiChusyutsu,
                                           pSeisanhiChusyutsu As SeisanhiChusyutsu,
                                           pSyukeiCensusNoList As List(Of String),
                                           Optional pNougyouKeieitaiJouken As String = Nothing,
                                           Optional kessokuchiHokana As String = ComConst.欠測値補完.無,
                                           Optional cboKeizokuKubun As String = Nothing, 'REV_008
                                           Optional chkZenkaiCensus As Boolean = False) As Dictionary(Of String, DataTable) 'REV_008

        Dim ret As New Dictionary(Of String, DataTable)
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim tableAddName As String = ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(chosakubun, ComConst.欠測値補完.有)

        Try
            Dim censusNoList As New List(Of String)
            Dim censusNoTable As DataTable = getCensusNoUseSyukei(db, chosakubun, pKey, kKey, pSyukeiInfo, pKibokaisou, pEinouKeieitaiChusyutsu, pSeisanhiChusyutsu, pSyukeiCensusNoList, pNougyouKeieitaiJouken, kessokuchiHokana)
            Dim prevCensusCol As String = ComConst.個別結果表.前回センサス番号(chosakubun)
            For Each dr As DataRow In censusNoTable.Rows
                If chkZenkaiCensus Then
                    ' 前回センサス番号を使用の場合、ComConst.個別結果表.前回センサス番号(chosakubun) の列を比較（K000005 条件も考慮）
                    Dim kPrev As String = If(String.IsNullOrEmpty(prevCensusCol), String.Empty, dr.Item(prevCensusCol).ToString)
                    Dim k005Val As Decimal = 0
                    If Not IsDBNull(dr.Item("K000005")) Then
                        Decimal.TryParse(dr.Item("K000005").ToString(), k005Val)
                    End If
                    Dim k006Val As Decimal = 0
                    If Not IsDBNull(dr.Item("K000006")) Then
                        Decimal.TryParse(dr.Item("K000006").ToString(), k006Val)
                    End If

                    Dim passK005 As Boolean = True
                    If CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費 Then
                        passK005 = (k005Val <> 0)
                    ElseIf CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then
                        If pSyukeiInfo.syukei1 = ComConst.集計区分.集計対象 Then
                            If chosakubun = ComConst.調査区分.米生産費統計_個別 Or chosakubun = ComConst.調査区分.米生産費統計_組織法人 Then
                                passK005 = (k005Val = 5)
                            Else
                                passK005 = (k006Val = 6)
                            End If
                        ElseIf pSyukeiInfo.syukei1 = ComConst.集計区分.全調査対象 Then
                            If chosakubun = ComConst.調査区分.米生産費統計_個別 Or chosakubun = ComConst.調査区分.米生産費統計_組織法人 Then
                                passK005 = (k005Val = 2 OrElse k005Val = 3 OrElse k005Val = 4 OrElse k005Val = 5)
                            Else
                                passK005 = (k006Val = 3 OrElse k006Val = 4 OrElse k006Val = 5 OrElse k006Val = 6)
                            End If

                        End If
                    End If

                    If kPrev <> "0000000000000000" AndAlso kPrev <> String.Empty AndAlso passK005 Then
                        censusNoList.Add(kPrev)
                    End If
                Else
                    censusNoList.Add(dr.Item("センサス番号").ToString)
                End If
            Next

            'REV_008↓
            ' 継続区分によるセンサス番号フィルタリング
            If Not String.IsNullOrEmpty(cboKeizokuKubun) AndAlso cboKeizokuKubun = ComConst.継続区分.継続のみ AndAlso censusNoList.Count > 0 Then
                ' 継続のみの場合、調査年のセンサス番号と前年のセンサス番号を比較
                Dim currentYear As Integer = Integer.Parse(pKey.chosaNen)
                Dim previousYear As String = (currentYear - 1).ToString()

                ' 単一のSQLで前年に存在するセンサス番号を取得
                Dim prevYearSb As New System.Text.StringBuilder()
                Dim prevYearPara As New List(Of DBAccess.Parameter)()

                ' 米生産費統計SQL CONDITION STRING
                Dim kubunCondition1 As String = "AND K000006 = 6"
                Dim kubunCondition2 As String = "AND K000006 IN (3,4,5,6)"
                ' 米生産費統計の場合、追加条件あり
                If chosakubun = ComConst.調査区分.米生産費統計_個別 Or chosakubun = ComConst.調査区分.米生産費統計_組織法人 Then
                    kubunCondition1 = "AND K000005 = 5"
                    kubunCondition2 = "AND K000005 IN (2,3,4,5)"
                End If

                With prevYearSb
                    .AppendLine("SELECT DISTINCT センサス番号")
                    .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosakubun)(0) & tableAddName))
                    .AppendLine("WHERE  調査年 = @前年")

                    ' 区分２による追加条件
                    If CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費 Then
                        .AppendLine("AND    K000005 <> 0")
                    ElseIf CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then
                        If pSyukeiInfo.syukei1 = ComConst.集計区分.集計対象 Then
                            .AppendLine(kubunCondition1)
                        ElseIf pSyukeiInfo.syukei1 = ComConst.集計区分.全調査対象 Then
                            .AppendLine(kubunCondition2)
                        End If
                    End If

                    .AppendLine("AND    センサス番号 IN (")

                    ' センサス番号のIN句を構築
                    For i As Integer = 0 To censusNoList.Count - 1
                        If i = 0 Then
                            .Append(String.Format("'{0}'", censusNoList(i)))
                        Else
                            .Append(String.Format(", '{0}'", censusNoList(i)))
                        End If
                    Next
                    .AppendLine(")")
                End With

                prevYearPara.Add(db.CreateParameter("@前年", SqlDbType.VarChar, previousYear))

                Dim prevYearResult As DataTable = db.GetDataTable(prevYearSb.ToString(), prevYearPara)

                ' 前年に存在するセンサス番号のみでリストを更新
                Dim filteredCensusNoList As New List(Of String)
                For Each row As DataRow In prevYearResult.Rows
                    filteredCensusNoList.Add(row("センサス番号").ToString())
                Next

                censusNoList = filteredCensusNoList
            End If

            Dim censusFilterCondition = "AND    センサス番号  IN("
            If chkZenkaiCensus Then
                If String.IsNullOrEmpty(prevCensusCol) Then
                    censusFilterCondition = "AND    センサス番号  IN("
                Else
                    censusFilterCondition = String.Format("AND    CAST({0} AS VARCHAR(16)) IN (", prevCensusCol)
                End If
            End If

            If censusNoList.Count > 0 Then
                Dim chosaNen As String = censusNoTable.Rows(0).Item("調査年").ToString
                Dim targetCensusNos As New List(Of String)
                Dim firstTableName As String = ComConst.個別結果表.テーブル名称(chosakubun)(0)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                With sb
                    .AppendLine("SELECT DISTINCT センサス番号")
                    .AppendLine(String.Format("FROM   ""{0}""", firstTableName & tableAddName))
                    .AppendLine("WHERE  調査年         = @調査年 ")
                    .AppendLine(censusFilterCondition)
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

                For Each row As DataRow In dt.Rows
                    targetCensusNos.Add(row("センサス番号").ToString())
                Next

                If targetCensusNos.Count > 0 Then
                    For Each tableName As String In ComConst.個別結果表.テーブル名称(chosakubun)

                        sb = New System.Text.StringBuilder
                        para = New List(Of DBAccess.Parameter)

                        ' SQL文の設定
                        With sb
                            .AppendLine("SELECT *")
                            .AppendLine(String.Format("FROM   ""{0}""", tableName & tableAddName))
                            .AppendLine("WHERE  調査年         = @調査年 ")
                            .AppendLine("AND    センサス番号  IN(")
                            For i As Integer = 0 To targetCensusNos.Count - 1
                                If i = 0 Then
                                    .Append(String.Format("'{0}'", targetCensusNos(i)))
                                Else
                                    .Append(String.Format(", '{0}'", targetCensusNos(i)))
                                End If
                            Next
                            .AppendLine(" )")
                            .AppendLine("ORDER BY センサス番号")
                        End With

                        para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
                        dt = db.GetDataTable(sb.ToString, para)

                        ret.Add(tableName & tableAddName, dt)
                    Next
                End If
            End If
            'REV_008↑
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_005↓
    ''' <summary>
    ''' 個別結果表取得SQLが実行できるか検証する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="param"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function isExecutableSQL(db As DBAccess, param As String, chosakubun As String) As Boolean
    Public Shared Function isExecutableSQL(db As DBAccess, param As String, chosakubun As String, versionKbn As String) As Boolean
        ' REV_005↑
        Dim ret As Boolean = True
        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing

        Try

            Try
                sb = New System.Text.StringBuilder

                If chosakubun.Equals(ComConst.営農経営体区分.農業経営体) Then
                    '農業経営体の場合は 営農個人と営農法人の両方でSQLを検証する
                    For Each einouChosaKbn As String In ComConst.営農経営体区分.リスト(chosakubun).調査区分
                        Dim tableAddName As String = ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(einouChosaKbn, ComConst.欠測値補完.有)

                        '固定文字列論理の置換
                        Dim executeParam As String = param
                        ' REV_005↓
                        'executeParam = ComUtil.ReplaceConstJouken(executeParam, einouChosaKbn)
                        executeParam = ComUtil.ReplaceConstJouken(executeParam, einouChosaKbn, versionKbn)
                        ' REV_005↑

                        ' SQL文の設定
                        With sb
                            .AppendLine("SELECT 調査年, センサス番号")
                            .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(einouChosaKbn)(0) & tableAddName))
                            .AppendLine(String.Format("WHERE    センサス番号   IN (SELECT ""{0}"".センサス番号 ", ComConst.個別結果表.テーブル名称(einouChosaKbn)(0) & tableAddName))
                            .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.個別結果表.テーブル名称(einouChosaKbn)(0) & tableAddName))
                            For i As Integer = 1 To ComConst.個別結果表.テーブル名称(einouChosaKbn).Count - 1
                                .AppendLine(String.Format("                          INNER JOIN   ""{0}""", ComConst.個別結果表.テーブル名称(einouChosaKbn)(i) & tableAddName))
                                .AppendLine(String.Format("                          ON   ""{0}"".センサス番号 = ""{1}"".センサス番号 ", ComConst.個別結果表.テーブル名称(einouChosaKbn)(i - 1) & tableAddName, ComConst.個別結果表.テーブル名称(einouChosaKbn)(i) & tableAddName))
                            Next
                            .AppendLine(String.Format("                          WHERE  {0}", executeParam))
                            .AppendLine(String.Format(" )"))
                        End With

                        dt = db.GetDataTable(sb.ToString)
                        sb.Clear()
                    Next
                Else
                    Dim tableAddName As String = ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(chosakubun, ComConst.欠測値補完.有)

                    ' SQL文の設定
                    With sb
                        .AppendLine("SELECT 調査年, センサス番号")
                        .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosakubun)(0) & tableAddName))
                        .AppendLine(String.Format("WHERE    センサス番号   IN (SELECT ""{0}"".センサス番号 ", ComConst.個別結果表.テーブル名称(chosakubun)(0) & tableAddName))
                        .AppendLine(String.Format("                          FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosakubun)(0) & tableAddName))
                        For i As Integer = 1 To ComConst.個別結果表.テーブル名称(chosakubun).Count - 1
                            .AppendLine(String.Format("                          INNER JOIN   ""{0}""", ComConst.個別結果表.テーブル名称(chosakubun)(i) & tableAddName))
                            .AppendLine(String.Format("                          ON   ""{0}"".センサス番号 = ""{1}"".センサス番号 ", ComConst.個別結果表.テーブル名称(chosakubun)(i - 1) & tableAddName, ComConst.個別結果表.テーブル名称(chosakubun)(i) & tableAddName))
                        Next
                        .AppendLine(String.Format("                          WHERE  {0}", param))
                        .AppendLine(String.Format(" )"))
                    End With

                    dt = db.GetDataTable(sb.ToString)
                End If

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
    Public Shared Function getCensusNoUseSyukei(db As DBAccess,
                                           chosakubun As String,
                                           pKey As DAOKobetsuKekkahyo.PrimaryKey,
                                           kKey As DAOKobetsuKekkahyo.KyotenKey,
                                           pSyukeiInfo As SyukeiInfo,
                                           pKibokaisou As Kibokaisou,
                                           pEinouKeieitaiChusyutsu As EinouKeieitaiChusyutsu,
                                           pSeisanhiChusyutsu As SeisanhiChusyutsu,
                                           pSyukeiCensusNoList As List(Of String),
                                           pNougyouKeieitaiJouken As String,
                                           kessokuchiHokana As String) As DataTable


        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim censusNoList As New List(Of String)
        Dim tableAddName As String = ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(chosakubun, ComConst.欠測値補完.有)

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine(String.Format("SELECT  ""{0}"".調査年, ""{0}"".センサス番号, CAST(""{0}"".{1} AS VARCHAR(16)) AS {1}, ""{0}"".K000005 AS K000005, ""{0}"".K000006 AS K000006 ", ComConst.個別結果表.テーブル名称(chosakubun)(0) & tableAddName, ComConst.個別結果表.前回センサス番号(chosakubun)))
                .AppendLine(String.Format("FROM   ""{0}""", ComConst.個別結果表.テーブル名称(chosakubun)(0) & tableAddName))
                For i As Integer = 1 To ComConst.個別結果表.テーブル名称(chosakubun).Count - 1
                    .AppendLine(String.Format("                          INNER JOIN   ""{0}""", ComConst.個別結果表.テーブル名称(chosakubun)(i) & tableAddName))
                    .AppendLine(String.Format("                          ON   ""{0}"".調査年 = ""{1}"".調査年 ", ComConst.個別結果表.テーブル名称(chosakubun)(i - 1) & tableAddName, ComConst.個別結果表.テーブル名称(chosakubun)(i) & tableAddName))
                    .AppendLine(String.Format("                          AND  ""{0}"".センサス番号 = ""{1}"".センサス番号 ", ComConst.個別結果表.テーブル名称(chosakubun)(i - 1) & tableAddName, ComConst.個別結果表.テーブル名称(chosakubun)(i) & tableAddName))
                Next
                .AppendLine(String.Format("                          WHERE  ""{0}"".調査年         = @調査年 ", ComConst.個別結果表.テーブル名称(chosakubun)(0) & tableAddName))
                If Not kKey Is Nothing Then
                    .AppendLine(String.Format("                          AND    ""{0}"".農政局         = @農政局 ", ComConst.個別結果表.テーブル名称(chosakubun)(0) & tableAddName))
                End If

                '個別結果表抽出条件：センサス番号
                If Not pSyukeiCensusNoList Is Nothing Then
                    .Append(String.Format("                          AND    ""{0}"".センサス番号         IN( ", ComConst.個別結果表.テーブル名称(chosakubun)(0) & tableAddName))
                    For i As Integer = 0 To pSyukeiCensusNoList.Count - 1
                        If i = 0 Then
                            .Append(String.Format("'{0}'", pSyukeiCensusNoList(i)))
                        Else
                            .Append(String.Format(", '{0}'", pSyukeiCensusNoList(i)))
                        End If
                    Next
                    .AppendLine(")")
                End If

                '個別結果表抽出条件：地域
                If pSyukeiInfo.chiiki.Length > 0 Then
                    .Append("                          AND    ")
                    .Append(ComConst.個別結果表.都道府県)
                    .Append("    IN( ")
                    For i As Integer = 0 To pSyukeiInfo.chiiki.Length - 1
                        If i = 0 Then
                            .Append(String.Format("'{0}'", pSyukeiInfo.chiiki(i)))
                        Else
                            .Append(String.Format(", '{0}'", pSyukeiInfo.chiiki(i)))
                        End If
                    Next
                    .AppendLine(")")
                End If

                '個別結果表抽出条件：集計倍率＞0
                .Append("                          AND ")
                If chosakubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) AndAlso (Not pEinouKeieitaiChusyutsu Is Nothing) AndAlso (Not pEinouKeieitaiChusyutsu.bumonCd Is Nothing) Then
                    .Append(ComConst.個別結果表.部門集計倍率(chosakubun))
                Else
                    .Append(ComConst.個別結果表.集計倍率(chosakubun))
                End If
                .AppendLine(" > 0")

                '個別結果表抽出条件《集計パターンが任意の場合のみ》：集計対象区分
                If chosakubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) AndAlso pSyukeiInfo.syukeipattern.Equals("任意") Then
                    .Append("                          AND ")
                    .Append(ComConst.個別結果表.集計対象区分(chosakubun))
                    .AppendLine(" = 1")
                End If

                If (Not String.IsNullOrEmpty(pSyukeiInfo.niniSyukei)) Or (Not pKibokaisou Is Nothing) Or (Not pEinouKeieitaiChusyutsu Is Nothing) Or (Not pSeisanhiChusyutsu Is Nothing) Then

                    '個別結果表抽出条件：任意集計
                    If Not String.IsNullOrEmpty(pSyukeiInfo.niniSyukei) Then
                        Dim niniSyukei As String = pSyukeiInfo.niniSyukei
                        '固定文字列を置換
                        ' REV_005↓
                        'niniSyukei = ComUtil.ReplaceConstJouken(niniSyukei, chosakubun)
                        niniSyukei = ComUtil.ReplaceConstJouken(niniSyukei, chosakubun, ComUtil.getVersionKubunTaikei(pKey.chosaNen, chosakubun))
                        ' REV_005↑
                        .Append("                          AND ")
                        .AppendLine(niniSyukei)
                    End If

                    '個別結果表抽出条件：営農経営体抽出条件
                    If (Not pEinouKeieitaiChusyutsu Is Nothing) AndAlso (Not pEinouKeieitaiChusyutsu.jouken Is Nothing) Then
                        .Append("                          AND ")
                        .AppendLine(If(pNougyouKeieitaiJouken Is Nothing, pEinouKeieitaiChusyutsu.jouken, pNougyouKeieitaiJouken))
                    End If

                    '個別結果表抽出条件：規模階層
                    If (Not pKibokaisou Is Nothing) AndAlso (Not pKibokaisou.jouken Is Nothing) Then
                        If (Not pKibokaisou.min Is Nothing) And (Not pKibokaisou.max Is Nothing) Then
                            Dim jouken As String = pKibokaisou.jouken
                            '固定文字列を置換
                            ' REV_005↓
                            'jouken = ComUtil.ReplaceConstJouken(jouken, chosakubun)
                            jouken = ComUtil.ReplaceConstJouken(jouken, chosakubun, ComUtil.getVersionKubunTaikei(pKey.chosaNen, chosakubun))
                            ' REV_005↑

                            '項番にISNULLを付与
                            Dim repJouken As New System.Text.StringBuilder
                            repJouken.Append(jouken)
                            Dim matchList As System.Text.RegularExpressions.MatchCollection = Nothing
                            '項目番号(K010101)を検索
                            matchList = System.Text.RegularExpressions.Regex.Matches(jouken, "K[0-9]{6}")
                            For Each mat As System.Text.RegularExpressions.Match In matchList
                                repJouken.Replace(mat.Value, "ISNULL(" & mat.Value & ", 0)")
                            Next
                            '上限下限が設定されている通常の規模階層
                            .Append("                          AND (")
                            .Append(repJouken.ToString)
                            .Append(" >= ")
                            .Append(pKibokaisou.min)
                            .Append(" AND ")
                            .Append(pKibokaisou.max)
                            .Append(" >= ")
                            .Append(repJouken.ToString)
                            .AppendLine(")")
                        Else
                            '階層に集計コードが指定されている規模階層
                            If pKibokaisou.jouken.Contains("@"c) Then
                                Dim joukenList As String()
                                If pKibokaisou.jouken IsNot Nothing Then
                                    joukenList = pKibokaisou.jouken.Split("@"c)
                                    .Append("                          AND ")
                                    .AppendLine(joukenList(CInt(chosakubun) - 1))
                                End If
                            Else
                                .Append("                          AND ")
                                .AppendLine(pKibokaisou.jouken)
                            End If
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

    ''' <summary>
    ''' 個別結果表データ追加
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="tableNo"></param>
    ''' <remarks></remarks>
    Public Shared Sub CopyTable(ByVal db As DBAccess, ByVal chosaNen As String, ByVal tableNo As Integer)

        If db Is Nothing Then
            Return
        End If

        If Not ComUtil.TryParseToInteger(chosaNen) Then
            Return
        End If

        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine(String.Format("INSERT INTO {0} ", ComConst.個別結果表.テーブル名称(ComConst.調査区分.営農類型別経営統計_個人)(tableNo) & ComConst.個別結果表.集計用テーブル付加名称))
                .AppendLine(" SELECT * ")
                .AppendLine(String.Format(" FROM {0} ", ComConst.個別結果表.テーブル名称(ComConst.調査区分.営農類型別経営統計_個人)(tableNo)))
                .AppendLine(" WHERE  調査年         = @調査年 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))

            db.ExecuteNonQuery(sb.ToString, para)

        Catch ex As Exception
            Throw
        End Try

    End Sub

    ''' <summary>
    ''' 貸借対照表区分が3のレコードを検索する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <remarks></remarks>
    Public Shared Function SelectKobetsuKekkaHyoSyuukeiTableWhereTaisyakuKubun3(ByVal db As DBAccess, ByVal chosaNen As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * ")
                .AppendLine(String.Format("FROM {0} ", ComConst.欠測値.個別結果表＿農業経営＿営農類型＿個人１＿集計用))
                .AppendLine("WHERE  調査年         = @調査年 ")
                .AppendLine("AND    K000018 = 3 ")
            End With

            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw
        End Try

        Return ret

    End Function

    ''' <summary>
    ''' 「個別結果表＿農業経営＿営農類型＿個人１＿集計用」テーブルを更新する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="row"></param>
    ''' <param name="censusNo"></param>
    ''' <remarks></remarks>
    Public Shared Sub UpdateKobetsuKekkaHyo1SyuukeiTable(ByVal db As DBAccess, ByVal row As ComUtil.営農欠測値適用平均値データ, ByVal censusNo As String)

        If db Is Nothing Then
            Return
        End If

        If row Is Nothing Then
            Return
        End If

        If String.IsNullOrWhiteSpace(censusNo) Then
            Return
        End If

        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine(String.Format("UPDATE {0} ", ComConst.欠測値.個別結果表＿農業経営＿営農類型＿個人１＿集計用))
                .AppendLine("SET   K010712      = @現金, ")
                .AppendLine("      K010713      = @預金等, ")
                .AppendLine("      K010720      = @売掛金＿未収金, ")
                .AppendLine("      K010732      = @自動車＿農機具, ")
                .AppendLine("      K010733      = @建物＿構築物, ")
                .AppendLine("      K010734      = @土地, ")
                .AppendLine("      K010737      = @果樹＿牛馬等, ")
                .AppendLine("      K010909      = @流動負債, ")
                .AppendLine("      K010911      = @買掛金, ")
                .AppendLine("      K010913      = @短期借入金, ")
                .AppendLine("      K010919      = @長期借入金, ")
                .AppendLine("      更新日付     = GETDATE(), ")
                .AppendLine("      更新者ID     = @更新者ID ")
                .AppendLine("WHERE 調査年       = @調査年 ")
                .AppendLine("AND   センサス番号 = @センサス番号 ")
            End With

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
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, row.調査年))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, censusNo))

            db.ExecuteNonQuery(sb.ToString, para)

        Catch ex As Exception
            Throw
        End Try

    End Sub

    '---REV002_ADD START
    ''' <summary>
    ''' 個別結果表テーブル取得(差分比較用)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kessokuchiHokan"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetKobetuTableHikaku(db As DBAccess, chosaNen As String, kyoku As String, jimusho As String, kyoten As String, einouRuikei As String, upLow As String,
                                                  ChosaKouteiTo As String, ChosaKouteiFrom As String, Optional jushin As Boolean = False) As Dictionary(Of String, DataTable)
        Dim ret As Dictionary(Of String, DataTable) = New Dictionary(Of String, DataTable)
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim kobetsuTableName As String
        Dim otherTableName As String
        Dim mainTableName As String
        Dim dt As DataTable
        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)
                If jushin Then
                    kobetsuTableName = String.Format("[{0}].[dbo].[受信＿{1}]", ChosaKouteiTo, ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0))
                    mainTableName = String.Format("[{0}].[dbo].[受信＿{1}]", ChosaKouteiTo, tableName)
                    otherTableName = String.Format("[{0}].[dbo].[{1}]", ChosaKouteiFrom, ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0))
                Else
                    kobetsuTableName = String.Format("[{0}].[dbo].[{1}]", ChosaKouteiFrom, ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0))
                    mainTableName = String.Format("[{0}].[dbo].[{1}]", ChosaKouteiFrom, tableName)
                    otherTableName = String.Format("[{0}].[dbo].[受信＿{1}]", ChosaKouteiTo, ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0))
                End If

                With sb
                    If tableName = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0) Then
                        .AppendLine("SELECT A.* ")
                        .AppendLine(String.Format("FROM {0} As A", kobetsuTableName))
                    Else
                        .AppendLine("SELECT B.* ")
                        .AppendLine("  ,A.農政局 ")
                        .AppendLine("  ,A.都道府県 ")
                        .AppendLine("  ,A.実査設置拠点 ")
                        .AppendLine(String.Format("FROM {0} As A", kobetsuTableName))

                        .AppendLine(String.Format("INNER JOIN {0} As B", mainTableName))
                        .AppendLine(" ON A.調査年 = B.調査年 ")
                        .AppendLine(" AND A.センサス番号 = B.センサス番号 ")
                        If jushin Then
                            .AppendLine(" AND A.上り下り区分 = B.上り下り区分 ")
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
                        .AppendLine(String.Format("AND    ""{0}""   = @営農類型 ", ComConst.個別結果表.営農類型(CommonInfo.Chosakubun)))
                    End If
                    If jushin Then
                        .AppendLine("                          AND    A.上り下り区分   = @上り下り区分 ")
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
    ''' 個別結果表データ削除（CSV一括登録）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteCSVTable(db As DBAccess, chosaKubun As String, chosaNen As String, dbName As String) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(chosaKubun).Reverse().ToArray()
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)


                ' SQL文の設定
                With sb
                    .AppendLine("DELETE FROM ")
                    .AppendLine(String.Format("[{0}].[dbo].[{1}] ", dbName, tableName))
                    .AppendLine("WHERE  調査年         = @調査年 ")
                End With

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))

                If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                    ret = True
                Else
                    Throw New Exception(String.Format("{0}削除失敗", dbName & ".dbo." & tableName))
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 個別結果表データ登録(CSV一括登録)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="csvList"></param>
    ''' <param name="dbName"></param>
    ''' <param name="progressDialog"></param>
    ''' <returns></returns>
    Public Shared Function InsertCSVTable(db As DBAccess, chosaKubun As String, chosaNen As String, csvList As Dictionary(Of String, List(Of String())),
                                          dbName As String, Optional ByRef progressDialog As ProgressDialog = Nothing) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim sb2 As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(chosaKubun)

                sb = New System.Text.StringBuilder
                sb2 = New System.Text.StringBuilder
                'カラム名,データ形式を取得
                Dim colArr As Dictionary(Of String, Type) = GetColumnsType(db, tableName)
                'CSVデータ取得
                Dim csvData As List(Of String()) = csvList(tableName)

                ' SQL文の設定
                With sb
                    .AppendLine(String.Format("INSERT INTO [{0}].[dbo].[{1}] ", dbName, tableName))
                    For i As Integer = 0 To colArr.Count - 1
                        If i = 0 Then
                            .AppendLine(String.Format("  ({0} ", colArr.Keys(i)))
                        Else
                            .AppendLine(String.Format("  ,{0} ", colArr.Keys(i)))
                        End If
                    Next
                    .AppendLine(") ")
                    .AppendLine("VALUES ")
                    With sb2
                        For Each csvLine As String() In csvData

                            For i As Integer = 0 To csvLine.Length - 1
                                Select Case colArr.Values(i).ToString
                                    Case Type.GetType("System.Int32").ToString, Type.GetType("System.Decimal").ToString
                                        If i = 0 Then
                                            .AppendLine(String.Format("({0}", If(String.IsNullOrEmpty(csvLine(i)), "NULL", csvLine(i))))
                                        Else
                                            .AppendLine(String.Format(",{0}", If(String.IsNullOrEmpty(csvLine(i)), "NULL", csvLine(i))))
                                        End If
                                    Case Type.GetType("System.String").ToString, Type.GetType("System.DateTime").ToString
                                        If i = 0 Then
                                            .AppendLine(String.Format("({0}", If(String.IsNullOrEmpty(csvLine(i)), "NULL", "'" & csvLine(i) & "'")))
                                        Else
                                            .AppendLine(String.Format(",{0}", If(String.IsNullOrEmpty(csvLine(i)), "NULL", "'" & csvLine(i) & "'")))
                                        End If
                                End Select

                                If i = csvLine.Length - 1 Then
                                    .AppendLine(");")
                                End If
                            Next

                            If db.ExecuteNonQuery(sb.ToString & sb2.ToString, para) >= 0 Then
                                ret = True
                                If Not IsNothing(progressDialog) Then
                                    progressDialog.AddValue = 1
                                End If
                            Else
                                Throw New Exception(String.Format("{0}登録失敗", dbName & ".dbo." & tableName))
                            End If

                            .Clear()
                        Next
                    End With

                End With

            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    Public Shared Function InsertCSVTableCopy(db As DBAccess, chosaKubun As String, chosaNen As String, dbName As String, Optional ByRef progressDialog As ProgressDialog = Nothing) As Boolean
        Dim ret As Boolean = False
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            For Each tableName As String In ComConst.個別結果表.テーブル名称(chosaKubun)
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine(String.Format("INSERT INTO [{0}].[dbo].[{1}] ", dbName, tableName))
                    .AppendLine(" SELECT * ")
                    .AppendLine(String.Format(" FROM  [BRAH].[dbo].[{0}] ", tableName))
                    .AppendLine("WHERE 調査年 = @調査年")
                End With

                para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))

                If db.ExecuteNonQuery(sb.ToString, para) >= 0 Then
                    ret = True
                    If Not IsNothing(progressDialog) Then
                        progressDialog.AddValue = 1
                    End If
                Else
                    Throw New Exception(String.Format("{0}登録失敗", dbName & ".dbo." & tableName))
                End If

            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret

    End Function

    ''' <summary>
    ''' 列名称・形式取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetColumnsType(db As DBAccess, tableName As String) As Dictionary(Of String, Type)
        Dim ret As New Dictionary(Of String, Type)

        Dim dt As DataTable
        Dim sb As System.Text.StringBuilder = Nothing

        Try
            sb = New System.Text.StringBuilder

            ' SQL文の設定
            With sb
                .AppendLine("SELECT TOP 0 *")
                .AppendLine(String.Format("FROM   ""{0}""", tableName))
            End With

            dt = db.GetDataTable(sb.ToString)

            For Each col As DataColumn In dt.Columns
                ret.Add(col.ColumnName, col.DataType)
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function
    '---REV002_ADD END

End Class
