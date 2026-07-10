Imports Microsoft.Office.Interop

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_002   | 2023.01.04 |大興電子通信        | 要件No.19
'//  REV_003   | 2023.12.22 |大興電子通信        | 既存不具合修正(連絡票No.163)
'//  REV_004   | 2024.05.31 |大興電子通信        | 追加要件No.1,3
'//  REV_005   | 2025.09.11 |GCU                 | 要件No.12
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 牛トレサデータプレプリント（異動データ）調査票出力
''' </summary>
''' <remarks></remarks>
Public Class BRA2110R
    Inherits ExcelOutputMultipleBaseClass

    Private Class 調査票項目
        Public Shared 識別番号 As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, "Q11020101"},
            {ComConst.牛トレサデータ.調査区分分類.子牛, "Q02020101"},
            {ComConst.牛トレサデータ.調査区分分類.乳用, "Q02020101"},
            {ComConst.牛トレサデータ.調査区分分類.交雑, "Q02020101"},
            {ComConst.牛トレサデータ.調査区分分類.去勢, "Q02020101"}
        }
        Public Shared 異動月 As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, "Q11020301"},
            {ComConst.牛トレサデータ.調査区分分類.子牛, "Q02020301"},
            {ComConst.牛トレサデータ.調査区分分類.乳用, "Q02020301"},
            {ComConst.牛トレサデータ.調査区分分類.交雑, "Q02020301"},
            {ComConst.牛トレサデータ.調査区分分類.去勢, "Q02020301"}
        }
        Public Shared 異動状況 As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, "Q11020401"},
            {ComConst.牛トレサデータ.調査区分分類.子牛, "Q02020401"},
            {ComConst.牛トレサデータ.調査区分分類.乳用, "Q02020401"},
            {ComConst.牛トレサデータ.調査区分分類.交雑, "Q02020401"},
            {ComConst.牛トレサデータ.調査区分分類.去勢, "Q02020401"}
        }
        Public Shared 種類コード As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, "Q11020901"},
            {ComConst.牛トレサデータ.調査区分分類.子牛, "Q02020701"},
            {ComConst.牛トレサデータ.調査区分分類.乳用, "Q02020701"},
            {ComConst.牛トレサデータ.調査区分分類.交雑, "Q02020701"},
            {ComConst.牛トレサデータ.調査区分分類.去勢, "Q02020701"}
        }
        Public Shared 品種コード As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, "Q11021001"},
            {ComConst.牛トレサデータ.調査区分分類.子牛, "Q02020801"},
            {ComConst.牛トレサデータ.調査区分分類.乳用, "Q02020801"},
            {ComConst.牛トレサデータ.調査区分分類.交雑, "Q02020801"},
            {ComConst.牛トレサデータ.調査区分分類.去勢, "Q02020801"}
        }
        Public Shared 性別コード As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, "Q11021101"},
            {ComConst.牛トレサデータ.調査区分分類.子牛, "Q02021001"},
            {ComConst.牛トレサデータ.調査区分分類.乳用, "Q02020901"},
            {ComConst.牛トレサデータ.調査区分分類.交雑, "Q02020901"},
            {ComConst.牛トレサデータ.調査区分分類.去勢, "Q02020901"}
        }
        Public Shared 生産年月 As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11021201", "Q11021301"}},
            {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02021101", "Q02021201"}},
            {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02021001", "Q02021101"}},
            {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02021001", "Q02021101"}},
            {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02021001", "Q02021101"}}
        }
    End Class

    Private _fromKikan As String
    Private _toKikan As String
    Private _hoseiList As List(Of BRA2120F.HoseiIdouInfo)
    Private _filenamechange As Long

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="outDir"></param>
    ''' <remarks></remarks>
'REV-001 START-----------------
    Public Sub New(ByVal outDir As String, ByVal fromKikan As String, ByVal toKikan As String, ByVal hoseiList As List(Of BRA2120F.HoseiIdouInfo), _chosaNen As String, filenamechange As Long)

        MyBase.New(ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun))).tempFileName, True, False, ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun))).reportName, outDir, False)

        'REV-001 END-----------------
        _filenamechange = filenamechange
        _fromKikan = fromKikan
        _toKikan = toKikan
        _hoseiList = hoseiList
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="xlSheets"></param>
    ''' <param name="unit"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(Of T As Class)(xlSheets As Microsoft.Office.Interop.Excel.Sheets, unit As T)
        Dim pkey As DAOChosahyo.PrimaryKey = CType(DirectCast(unit, Object), DAOChosahyo.PrimaryKey)
        'REV-002 START-----------------
        'REV-001 START-----------------
        If _filenamechange = 0 Then
            '作成ボタン押下時
            Dim fileName As String = ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(pkey.chosaNen, CommonInfo.Chosakubun))).reportName _
                                 & "（牛トレサ" & _fromKikan & "-" & _toKikan & "）" & "_" & pkey.chosaNen & "_" & pkey.censusNo & ".xlsm"
            'REV-001 END-----------------
            Me.OutPath = IO.Path.Combine(Me.OutDir, fileName)
        ElseIf _filenamechange = 1 Then
            '搾乳牛の成畜ボタン押下時
            Dim fileName As String = ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(pkey.chosaNen, CommonInfo.Chosakubun))).reportName _
                                 & "（牛トレサ成畜" & _fromKikan & "-" & _toKikan & "）" & "_" & pkey.chosaNen & "_" & pkey.censusNo & ".xlsm"

            Me.OutPath = IO.Path.Combine(Me.OutDir, fileName)
        ElseIf _filenamechange = 2 Then
            '全所有牛情報出力ボタン押下時
            Dim fileName As String = ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(pkey.chosaNen, CommonInfo.Chosakubun))).reportName _
                                 & "（牛トレサ全所有牛異動" & _fromKikan & "-" & _toKikan & "）" & "_" & pkey.chosaNen & "_" & pkey.censusNo & ".xlsm"
            'REV-001 END-----------------
            Me.OutPath = IO.Path.Combine(Me.OutDir, fileName)
        End If
        'REV-002 END-----------------
        Dim dtItem As DataTable
        Dim dtTresa As DataTable
        Dim dtChosahyo As Dictionary(Of String, DataTable)
        Dim dtChosahyoTresa As DataTable
        Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '調査票項目マスタ取得
            dtItem = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, pkey.chosaNen, ComConst.数式区分.数式ではない)

            '表紙シート、決算月、農家団体コードのデータを取得
            Dim dcKouban As Dictionary(Of String, String()) = GetChosahyoKoban(dtItem)

            '調査票テーブル取得
            dtChosahyo = DAOChosahyo.GetChosahyoTable(db, pkey, dcKouban)

            '調査票項目取得
            dcChosahyo = ComUtil.Chosahyo.GetItem(dtItem, dtChosahyo)

            '農家団体コードをすべて取得
            Dim farmCodeList As List(Of String) = ComUtil.Tresa.GetFarmCodeList(dcChosahyo)

            '異動データの範囲を設定
            Dim idoDateFrom As Integer = Integer.Parse(_fromKikan) * 100 + 1
            Dim idoDateFrom_y As String
            Dim idoDateFrom_m As String
            If CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 Then
                idoDateFrom_y = pkey.chosaNen
                idoDateFrom_m = "1"
            Else
                idoDateFrom_y = _fromKikan.Substring(0, 4)
                idoDateFrom_m = _fromKikan.Substring(4, 2)
            End If
            Dim idoDateTo As Integer = Integer.Parse(_toKikan) * 100 + 1
            Dim tmpDate As New Date(idoDateTo \ 10000, (idoDateTo \ 100) Mod 100, idoDateTo Mod 100)
            tmpDate = tmpDate.AddMonths(1)
            tmpDate = tmpDate.AddDays(-1)
            idoDateTo = (tmpDate.Year * 10000) + (tmpDate.Month * 100) + tmpDate.Day

            'REV-002 START-----------------
            If _filenamechange = 0 OrElse _filenamechange = 2 Then
                'REV-002 END-----------------
                '牛トレサ異動情報取得
                dtTresa = DAOOther.GetTresa(db, farmCodeList, idoDateFrom, idoDateTo, True)
            Else
                dtTresa = DAOOther.GetSeitikuTresa(db, farmCodeList, idoDateFrom, idoDateTo, idoDateFrom_y, idoDateFrom_m, True)
            End If
            '進捗加増
            Me.ProgressAddValue = 1

            '牛トレサデータ⇒調査票のデータ変換
            dtChosahyoTresa = GetChosahyoData(db, pkey, dtTresa, dcChosahyo, idoDateFrom_y, idoDateFrom_m)
        End Using

        '牛トレサデータから変換したデータをマージ
        dtChosahyo(ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(2)).Merge(dtChosahyoTresa)

        '調査票項目取得（牛トレサデータから変換したデータにExcel挿入位置情報を付加）
        dcChosahyo = ComUtil.Chosahyo.GetItem(dtItem, dtChosahyo)

        '進捗加増
        Me.ProgressAddValue = 1

        '調査票シートデータ設定
        ComUtil.Tresa.SetSheetData(dcChosahyo, xlSheets, CType(Me, ComObjectProcess), CommonInfo.Chosakubun)

        '進捗加増
        Me.ProgressAddValue = 1
    End Sub

    ''' <summary>
    ''' 表紙シート、決算月、農家団体コードの項番を取得
    ''' </summary>
    ''' <param name="dtItem">調査票項目マスタテーブル</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChosahyoKoban(dtItem As DataTable) As Dictionary(Of String, String())
        Dim targetRow() As DataRow = dtItem.Select("シート名 = '表紙'")
        Dim dcKouban As New Dictionary(Of String, String())
        Dim koubanList As New List(Of String)
        Dim variableList As New List(Of String)

        For Each row As DataRow In targetRow
            koubanList.Add(row("項目番号").ToString)
        Next
        If Not String.IsNullOrEmpty(ComConst.牛トレサデータ.決算月_項目番号(CommonInfo.Chosakubun)) Then
            koubanList.Add(ComConst.牛トレサデータ.決算月_項目番号(CommonInfo.Chosakubun))
        End If
        dcKouban.Add(ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0), koubanList.ToArray)
        dcKouban.Add(ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(2), {ComConst.調査票.項目番号.牛農家団体コード(CommonInfo.Chosakubun)})

        Return dcKouban
    End Function

    ''' <summary>
    ''' 画面で指定された調査年の調査票を取得する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetNowChosahyo(ByVal db As DBAccess, pkey As DAOChosahyo.PrimaryKey) As Dictionary(Of String, DAOChosahyo.調査票項目)



        '調査票項目マスタ取得
        Dim chosahyoItemMaster = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, pkey.chosaNen)

        '調査票テーブル取得
        Dim chosahyoTable = DAOChosahyo.GetChosahyoTable(db, New DAOChosahyo.PrimaryKey(pkey.chosaNen, pkey.censusNo))

        '調査票項目取得
        Return ComUtil.Chosahyo.GetItem(chosahyoItemMaster, chosahyoTable)

    End Function

    ''' <summary>
    ''' 経営体で所有する乳用種一覧を作成
    ''' </summary>
    ''' <param name="dcChosahyo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetNyuyouList(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As List(Of String)

        Dim nyuyouList As New List(Of String)

        '異動データより乳用種を取得
        Dim idoFrom As Integer = 0
        Dim idoTo As Integer = Integer.Parse(_toKikan) * 100 + 1

        Dim farmCodeList As List(Of String) = ComUtil.Tresa.GetFarmCodeList(dcChosahyo)
        Dim dtTresa As DataTable
        '牛トレサデータ取得
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dtTresa = DAOOther.GetTresa(db, farmCodeList, idoFrom, idoTo)
        End Using
        '取得した牛トレサデータのうち、乳用種(性別コード＝2、品種コード＝1、2、12)を抽出
        Dim whereHinsyu As String = "牛の識別CD IN (" & ComConst.牛トレサデータ.牛の識別CD.ホルスタイン種 & "," &
            ComConst.牛トレサデータ.牛の識別CD.ジャージー種 & "," &
            ComConst.牛トレサデータ.牛の識別CD.乳用種 & ")"
        Dim rows As DataRow() = dtTresa.Select("性別コード = " & ComConst.牛トレサデータ.性別CD.めす & " AND " & whereHinsyu)

        '抽出したデータの個体識別番号をリストに追加
        For Each row As DataRow In rows
            Dim kotaiNo As String = row("個体識別番号").ToString
            If Not nyuyouList.Contains(kotaiNo) Then
                nyuyouList.Add(kotaiNo)
            End If
        Next

        Return nyuyouList
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pkey"></param>
    ''' <param name="dtTresaHensei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChosahyoData(db As DBAccess, pkey As DAOChosahyo.PrimaryKey, dtTresaHensei As DataTable, chosahyo As Dictionary(Of String, DAOChosahyo.調査票項目), idoDateFrom_y As String, idoDateFrom_m As String) As DataTable

        Dim choIdouCode As String '異動コード
        Dim choShuruiCode As String '種類コード
        Dim choHinsyuCode As String '品種コード
        Dim choSexCode As String '性別コード
        Dim chosaBunrui As ComConst.牛トレサデータ.調査区分分類

        Dim meisaiNo As Integer = 1

        Dim chosahyoInsert As New Dictionary(Of String, DAOChosahyo.調査票項目)

        Dim dtChosahyo As New DataTable()
        dtChosahyo.Columns.Add("調査年", GetType(Integer))
        dtChosahyo.Columns.Add("センサス番号", GetType(String))
        dtChosahyo.Columns.Add("項目番号", GetType(String))
        dtChosahyo.Columns.Add("明細番号", GetType(Integer))
        dtChosahyo.Columns.Add("値", GetType(String))
        dtChosahyo.Columns.Add("更新日付", GetType(DateTime))
        dtChosahyo.Columns.Add("更新者ID", GetType(String))

        '母牛データ確認用
        Dim dtShikibetsu As New DataTable()
        dtShikibetsu.Columns.Add("母牛個体識別番号", GetType(String))

        '自家の乳用種一覧を作成(牛乳生産費で編成有無の判定で使用)
        Dim nyuyouList As List(Of String) = GetNyuyouList(chosahyo)

        '編成対象となる牛トレサデータ分ループ
        For Each row As DataRow In dtTresaHensei.Rows

            '補正対象レコードを除去
            If IsHoseiTarget(row) Then
                Continue For
            End If

            Dim kindCode As Integer = Integer.Parse(row("牛の識別CD").ToString)
            Dim sexCode As Integer = Integer.Parse(row("性別コード").ToString)
            Dim idouFlag As Integer = Integer.Parse(row("異動フラグ").ToString)

            If ComConst.牛トレサデータ.調査区分分類変換テーブル.ContainsKey(CommonInfo.Chosakubun) Then
                chosaBunrui = ComConst.牛トレサデータ.調査区分分類変換テーブル(CommonInfo.Chosakubun)
            Else
                '対象外の調査区分
                Exit For
            End If

            'REV-002 START-----------------
            '編成有無をチェック
            If _filenamechange <> 2 AndAlso ComUtil.Tresa.IsOrganizeData(chosaBunrui, kindCode, sexCode, idouFlag, row("母牛個体識別番号").ToString(), nyuyouList) = False Then
                Continue For
            End If
            'REV-002 END-----------------

            '購入・売却／異動の状況
            choIdouCode = ComUtil.Tresa.GetIdouCode(chosaBunrui, idouFlag)
            '種類コードを変換
            choShuruiCode = ComUtil.Tresa.GetShuruiCode(chosaBunrui, kindCode, sexCode, _filenamechange = 2) 'REV_004 CNG DAIKO 'optional引数追加
            '品種コードを変換
            choHinsyuCode = ComUtil.Tresa.GetHinshuCode(chosaBunrui, kindCode, sexCode, _filenamechange = 2) 'REV_004 CNG DAIKO 'optional引数追加
            '性別コードを変換
            choSexCode = ComConst.牛トレサデータ.性別コード変換テーブル(sexCode)

            '［以下調査票への挿入データ格納］

            Dim chosaNen As Integer = Integer.Parse(pkey.chosaNen)
            Dim censusNo As String = pkey.censusNo

            '成畜年月を取得
            Dim adultDate As String = ""
            If chosaBunrui = ComConst.牛トレサデータ.調査区分分類.牛乳 Then
                If ComUtil.Tresa.IsAdultTarget(choShuruiCode, choHinsyuCode) Then
                    'REV_005 START------
                    If _filenamechange = 0 Then
                        '作成ボタン押下時
                        '成畜年月を取得ロジック改修
                        adultDate = ComUtil.Tresa.GetAdultCattleDate(db, row("個体識別番号").ToString, row("生年月日").ToString, row("異動年月日").ToString, chosaNen.ToString)
                    ElseIf _filenamechange = 1 Then
                        '搾乳牛の成畜ボタン押下時
                        adultDate = ComUtil.Tresa.GetAdultDate(db, row("個体識別番号").ToString, row("生年月日").ToString, row("異動年月日").ToString)
                    Else
                        '全所有牛情報出力ボタン押下時
                        '全所有牛情報出力の場合、成畜年月を出力しない
                        adultDate = ""
                    End If
                    'REV_005 END------
                End If
            End If
            '母体の情報を取得（1件だけ抜ける）
            Dim momData As DataTable = GetTresaKotai(db, row("母牛個体識別番号").ToString)
            'REV_003↓成畜で母牛情報がない場合、システムエラーが発生するため出力対象外とする
            If _filenamechange = 1 AndAlso momData.Rows.Count <= 0 Then
                Continue For
            End If
            'REV_003↑

            'REV-002 START-----------------
            If _filenamechange = 0 OrElse _filenamechange = 2 Then
                'REV-002 END-----------------
                '識別番号
                dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.識別番号(chosaBunrui), meisaiNo, row("個体識別番号").ToString)
                '異動月
                dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.異動月(chosaBunrui), meisaiNo, row("異動年月日").ToString.Substring(4, 2))
                '購入・売却
                dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.異動状況(chosaBunrui), meisaiNo, choIdouCode)
            Else
                '成畜処理の場合、処理対象外かをここで確認
                '対象の母牛か確認
                If ("1" = momData(0)("牛の識別CD").ToString And "2" = momData(0)("性別コード").ToString) OrElse ("2" = momData(0)("牛の識別CD").ToString And "2" = momData(0)("性別コード").ToString) OrElse ("12" = momData(0)("牛の識別CD").ToString And "2" = momData(0)("性別コード").ToString) Then
                    '問題なし
                Else
                    '母牛としての条件を満たさないためスキップ
                    Continue For
                End If
                If (Integer.Parse(idoDateFrom_y) - Integer.Parse(momData(0)("生年月日").ToString.Substring(0, 4))) * 12 + (Integer.Parse(idoDateFrom_m) - Integer.Parse(momData(0)("生年月日").ToString.Substring(4, 2))) >= 30 Then
                    '生産年月から調査開始年月までの月齢をチェック、30以上は対象外（月齢の考え方は異動と同じ）、母牛の生年月日で確認
                    Continue For
                End If
                If Integer.Parse(row("生年月日").ToString.Substring(0, 4)) * 100 + Integer.Parse(row("生年月日").ToString.Substring(4, 2)) < Integer.Parse(_fromKikan.Substring(0, 4)) * 100 + Integer.Parse(_fromKikan.Substring(4, 2)) Then
                    '調査期間（牛トレサデータ期間）以前に成畜の牛は対象外（期間内の成畜のみ）※子畜の生年月日が調査開始年月以降
                    Continue For
                End If

                Dim drShikibetsu As DataRow() = dtShikibetsu.Select("母牛個体識別番号 = " & row("母牛個体識別番号").ToString)
                If 0 < drShikibetsu.Count Then
                    '同じ識別番号の処理の場合スキップ
                    Continue For
                Else
                    '同じ母牛個体識別番号を処理しないようにスキップ
                    dtShikibetsu.Rows.Add(row("母牛個体識別番号").ToString)
                End If
                '成畜処理の場合、識別番号を出力
                dtChosahyo.Rows.Add(chosaNen, censusNo, "Q11020701", meisaiNo, row("母牛個体識別番号").ToString)
            End If

            If chosaBunrui = ComConst.牛トレサデータ.調査区分分類.牛乳 Then
                '「識別番号（成畜異動の場合のみ入力）」は牛トレサデータの異動情報に対応したレコードでは出力しない
                'If adultDate.Length = 8 Then
                '    '成畜年月をプレプリントする場合のみ
                '    '識別番号
                '    dtChosahyo.Rows.Add(chosaNen, censusNo, "Q11020701", meisaiNo, row("個体識別番号").ToString)
                'End If
                '異動年月日
                Dim idouDate As String = row("異動年月日").ToString()
                idouDate = String.Format("{0}/{1}/{2}", idouDate.Substring(0, 4), idouDate.Substring(4, 2), idouDate.Substring(6, 2))
                dtChosahyo.Rows.Add(chosaNen, censusNo, "Q11020801", meisaiNo, idouDate)
            End If

            If chosaBunrui = ComConst.牛トレサデータ.調査区分分類.牛乳 Then
                '牛乳の場合
                'REV-002 START-----------------
                If _filenamechange = 0 OrElse _filenamechange = 2 Then
                    'REV-002 END-----------------
                    '通常
                    '種類 コード
                    dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.種類コード(chosaBunrui), meisaiNo, choShuruiCode)
                    '品種 コード
                    dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.品種コード(chosaBunrui), meisaiNo, choHinsyuCode)
                    '母畜の識別番号は子牛の場合のみなのでセットしない
                    '性別区分 コード
                    dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.性別コード(chosaBunrui), meisaiNo, choSexCode)
                    '生産年月
                    dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.生産年月(chosaBunrui)(0), meisaiNo, row("生年月日").ToString.Substring(0, 4))
                    dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.生産年月(chosaBunrui)(1), meisaiNo, row("生年月日").ToString.Substring(4, 2))
                    '成畜年月
                    If adultDate.Length = 8 Then
                        dtChosahyo.Rows.Add(chosaNen, censusNo, "Q11021401", meisaiNo, adultDate.Substring(0, 4))
                        dtChosahyo.Rows.Add(chosaNen, censusNo, "Q11021501", meisaiNo, adultDate.Substring(4, 2))
                    End If
                Else
                    '成畜
                    '種類 コード
                    dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.種類コード(chosaBunrui), meisaiNo, ComUtil.Tresa.GetShuruiCode(chosaBunrui, CInt(momData(0)("牛の識別CD")), 2)) ' REV_004 CNG DAIKO
                    '品種 コード
                    dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.品種コード(chosaBunrui), meisaiNo, ComUtil.Tresa.GetHinshuCode(chosaBunrui, CInt(momData(0)("牛の識別CD")), 2)) ' REV_004 CNG DAIKO
                    '母畜の識別番号は子牛の場合のみなのでセットしない
                    '性別区分 コード
                    dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.性別コード(chosaBunrui), meisaiNo, "2")
                    '生産年月（母牛分出力）
                    dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.生産年月(chosaBunrui)(0), meisaiNo, momData(0)("生年月日").ToString.Substring(0, 4))
                    dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.生産年月(chosaBunrui)(1), meisaiNo, momData(0)("生年月日").ToString.Substring(4, 2))
                    '成畜年月（子畜の生年月日をセット）
                    dtChosahyo.Rows.Add(chosaNen, censusNo, "Q11021401", meisaiNo, row("生年月日").ToString.Substring(0, 4))
                    dtChosahyo.Rows.Add(chosaNen, censusNo, "Q11021501", meisaiNo, row("生年月日").ToString.Substring(4, 2))
                End If

                '取得年月
                If row("異動フラグ").ToString.Equals(ComConst.牛トレサデータ.異動フラグ.転入搬入.ToString) Then
                    '異動フラグが 4(転入/搬入)の場合
                    dtChosahyo.Rows.Add(chosaNen, censusNo, "Q11021601", meisaiNo, row("異動年月日").ToString.Substring(0, 4))
                    dtChosahyo.Rows.Add(chosaNen, censusNo, "Q11021701", meisaiNo, row("異動年月日").ToString.Substring(4, 2))
                End If
            Else
                '牛乳以外
                '種類 コード
                dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.種類コード(chosaBunrui), meisaiNo, choShuruiCode)
                '品種 コード
                dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.品種コード(chosaBunrui), meisaiNo, choHinsyuCode)
                '母畜の識別番号
                If chosaBunrui = ComConst.牛トレサデータ.調査区分分類.子牛 Then
                    dtChosahyo.Rows.Add(chosaNen, censusNo, "Q02020901", meisaiNo, row("母牛個体識別番号").ToString)
                End If
                '性別区分 コード
                dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.性別コード(chosaBunrui), meisaiNo, choSexCode)
                '生産年月
                dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.生産年月(chosaBunrui)(0), meisaiNo, row("生年月日").ToString.Substring(0, 4))
                dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.生産年月(chosaBunrui)(1), meisaiNo, row("生年月日").ToString.Substring(4, 2))
            End If

            meisaiNo += 1
        Next

        Return dtChosahyo
    End Function

    ''' <summary>
    ''' 補正（削除）対象レコードかどうかを判定する
    ''' </summary>
    ''' <param name="row"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsHoseiTarget(row As DataRow) As Boolean
        For Each hosei In _hoseiList
            If Not hosei.IsDelete Then
                Continue For
            End If

            If hosei.OutFarmCode.Equals(row("農家団体コード").ToString) And
                hosei.KotaiNo.Equals(row("個体識別番号").ToString) And
                hosei.OutIdoDate.Equals(row("異動年月日").ToString) And
                hosei.OutIdoFlag.Equals(row("異動フラグ").ToString) Then
                Return True
            End If

            If hosei.InFarmCode.Equals(row("農家団体コード").ToString) And
                hosei.KotaiNo.Equals(row("個体識別番号").ToString) And
                hosei.InIdoDate.Equals(row("異動年月日").ToString) And
                hosei.InIdoFlag.Equals(row("異動フラグ").ToString) Then
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' 個体識別番号で個体情報を取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="cowID">個体識別番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetTresaKotai(db As DBAccess, cowID As String) As DataTable

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
                .AppendLine("WHERE  個体識別番号         = @個体識別番号 ")
            End With

            para.Add(db.CreateParameter("@個体識別番号", SqlDbType.VarChar, cowID))

            ret = db.GetDataTable(sb.ToString, para)
        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

End Class
