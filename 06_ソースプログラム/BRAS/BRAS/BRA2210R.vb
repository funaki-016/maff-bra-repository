Imports Microsoft.Office.Interop

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2023.01.04 |大興電子通信        | 要件No.19
'//  REV_002   | 2025.09.11 |GCU                 | 要件No.12
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 牛トレサプレプリント（総括データ）調査票出力
''' </summary>
''' <remarks></remarks>
Public Class BRA2210R
    Inherits ExcelOutputMultipleBaseClass

    Private Class 調査票項目
        Public Shared 識別番号 As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, "Q11021801"},
            {ComConst.牛トレサデータ.調査区分分類.子牛, "Q02030301"},
            {ComConst.牛トレサデータ.調査区分分類.乳用, "Q02021401"},
            {ComConst.牛トレサデータ.調査区分分類.交雑, "Q02021401"},
            {ComConst.牛トレサデータ.調査区分分類.去勢, "Q02021401"}
        }

        '一回目に牛乳生産費個別、アプリ終了せずに二回目牛乳生産費（法人）に変えた場合、牛乳生産費個別の値となり、１回目(CommonInfo.Chosakubun)の判断結果を保持し続けるため、方式変更
        'すべてにおいて、修正前後を残さないが、以下の方針で修正し、配列のどこを指すかはプログラムで判断
        '修正前：{ComConst.牛トレサデータ.調査区分分類.牛乳, If(CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別, "Q11022001", "Q11021901")}
        '修正後：{ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11022001", "Q11021901"}}
        Public Shared 種類コード As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11022001", "Dummy", "Q11021901", "Dummy"}},
            {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02030501"}},
            {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02021501"}},
            {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02021501"}},
            {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02021501"}}
        }

        Public Shared 品種コード As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11022101", "Dummy", "Q11022001", "Dummy"}},
            {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02030601"}},
            {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02021601"}},
            {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02021601"}},
            {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02021601"}}
        }
        Public Shared 性別コード As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11022201", "Dummy", "Q11022101", "Dummy"}},
            {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02030801"}},
            {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02021701"}},
            {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02021701"}},
            {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02021701"}}
        }
        Public Shared 生産年月 As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11022301", "Q11022401", "Q11022201", "Q11022301"}},
            {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02031001", "Q02031101"}},
            {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02021901", "Q02022001"}},
            {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02021901", "Q02022001"}},
            {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02021901", "Q02022001"}}
        }
        Public Shared 取得年月 As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11022701", "Q11022801", "Q11022601", "Q11022701"}},
            {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02031201", "Q02031301"}},
            {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02022101", "Q02022201"}},
            {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02022101", "Q02022201"}},
            {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02022101", "Q02022201"}}
        }
        '}
    End Class

    Private _fromKikan As String
    Private _toKikan As String

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
    Public Sub New(ByVal outDir As String, _chosaNen As String)

        MyBase.New(ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun))).tempFileName, True, False, ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun))).reportName, outDir, False)

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
        Dim fileName As String = ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(pkey.chosaNen, CommonInfo.Chosakubun))).reportName _
                                 & "（牛トレサ全所有牛総括）" & pkey.chosaNen & "_" & pkey.censusNo & ".xlsm"
        Me.OutPath = IO.Path.Combine(Me.OutDir, fileName)
        Dim dtItem As DataTable
        'Dim dtTresa As DataTable
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

            '【処理詳細仕様 6-1 ②】牛トレサデータ取得処理
            '調査票データ取得
            Dim chosahyo = GetNowChosahyo(db, pkey)
            '異動年月日を設定
            Dim idoDate As Integer = ComUtil.Tresa.GetIdoDateEnd(pkey.chosaNen, chosahyo)
            If idoDate = -1 Then
                ''決算月不明
                'resultList.Add(ValueTuple.Create(False, key.Item1.censusNo))
                Return
            End If

            '農家団体コードをすべて取得
            Dim farmCodeList As List(Of String) = ComUtil.Tresa.GetFarmCodeList(chosahyo)

            '牛トレサ異動情報取得
            Dim dtTresa As DataTable = DAOOther.GetTresa(db, farmCodeList, 0, idoDate)

            '編成対象とする牛トレサデータを絞り込む
            Dim dtTresaHensei As DataTable = GetHenseiTresaData(dtTresa)

            '進捗加増
            Me.ProgressAddValue = 1

            '牛トレサデータ⇒調査票のデータ変換
            'dtChosahyoTresa = GetChosahyoData(db, pkey, dtTresa, dcChosahyo)

            dtChosahyoTresa = GetChosahyoData(db, pkey, dtTresaHensei, ComUtil.Tresa.GetChosaBeginDate(idoDate))
        End Using

        '牛トレサデータから変換したデータをマージ
        dtChosahyo(ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(2)).Merge(dtChosahyoTresa)

        '調査票項目取得（牛トレサデータから変換したデータにExcel挿入位置情報を付加）
        dcChosahyo = ComUtil.Chosahyo.GetItem(dtItem, dtChosahyo)

        '進捗加増
        Me.ProgressAddValue = 1

        '調査票シートデータ設定
        ComUtil.Tresa.SetSheetDataSoukatsu(dcChosahyo, xlSheets, CType(Me, ComObjectProcess), CommonInfo.Chosakubun)

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
    ''' 調査票へ挿入するデータの取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="dtTresaHensei">調査票への編成対象の牛トレサデータ</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChosahyoData(db As DBAccess, pkey As DAOChosahyo.PrimaryKey, dtTresaHensei As DataTable, beginDate As Date) As DataTable

        Dim choShuruiCode As String '種類コード
        Dim choHinsyuCode As String '品種コード
        Dim choSexCode As String '性別コード
        Dim chosaBunrui As ComConst.牛トレサデータ.調査区分分類

        Dim meisaiNo As Integer = 1

        'Dim chosahyoInsert As New Dictionary(Of String, DAOChosahyo.調査票項目)
        Dim dtChosahyo As New DataTable()
        dtChosahyo.Columns.Add("調査年", GetType(Integer))
        dtChosahyo.Columns.Add("センサス番号", GetType(String))
        dtChosahyo.Columns.Add("項目番号", GetType(String))
        dtChosahyo.Columns.Add("明細番号", GetType(Integer))
        dtChosahyo.Columns.Add("値", GetType(String))
        dtChosahyo.Columns.Add("更新日付", GetType(DateTime))
        dtChosahyo.Columns.Add("更新者ID", GetType(String))

        'REV_003 Add START
        '経営分析調査_牛乳生産費の場合のみ、参照する配列番号を＋２する
        Dim gyunyuKeieiHosei As Integer = 0
        If CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
            gyunyuKeieiHosei = 2
        End If
        'REV_003 Add END

        '編成対象となる牛トレサデータ分ループ
        For Each row As DataRow In dtTresaHensei.Rows

            Dim kindCode As Integer = Integer.Parse(row("牛の識別CD").ToString)
            Dim sexCode As Integer = Integer.Parse(row("性別コード").ToString)
            Dim idouFlag As Integer = Integer.Parse(row("異動フラグ").ToString)

            If ComConst.牛トレサデータ.調査区分分類変換テーブル.ContainsKey(CommonInfo.Chosakubun) Then
                chosaBunrui = ComConst.牛トレサデータ.調査区分分類変換テーブル(CommonInfo.Chosakubun)
            Else
                '対象外の調査区分
                Exit For
            End If

            '種類コードを変換
            choShuruiCode = ComUtil.Tresa.GetShuruiCode(chosaBunrui, kindCode, sexCode)

            '品種コードを変換
            choHinsyuCode = ComUtil.Tresa.GetHinshuCode(chosaBunrui, kindCode, sexCode)

            '性別コードを変換
            choSexCode = ComConst.牛トレサデータ.性別コード変換テーブル(sexCode)

            '［以下調査票への挿入データ格納］

            Dim chosaNen As Integer = Integer.Parse(pkey.chosaNen)
            Dim censusNo As String = pkey.censusNo

            '識別番号
            dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.識別番号(chosaBunrui), meisaiNo, row("個体識別番号").ToString)

            '種類　コード
            'REV_003 MOD START
            'AddChosahyoData(chosahyoInsert, 調査票項目.種類コード(chosaBunrui), meisaiNo, choShuruiCode)
            dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.種類コード(chosaBunrui)(0 + gyunyuKeieiHosei), meisaiNo, choShuruiCode)

            '品種　コード
            'REV_003 MOD START
            'AddChosahyoData(chosahyoInsert, 調査票項目.品種コード(chosaBunrui), meisaiNo, choHinsyuCode)
            dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.品種コード(chosaBunrui)(0 + gyunyuKeieiHosei), meisaiNo, choHinsyuCode)
            'REV_003 MOD END

            '母畜の識別番号
            If chosaBunrui = ComConst.牛トレサデータ.調査区分分類.子牛 Then
                dtChosahyo.Rows.Add(chosaNen, censusNo, "Q02030701", meisaiNo, row("母牛個体識別番号").ToString)
            End If

            '性別区分　コード
            'REV_003 MOD START
            'AddChosahyoData(chosahyoInsert, 調査票項目.性別コード(chosaBunrui), meisaiNo, choSexCode)
            dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.性別コード(chosaBunrui)(0 + gyunyuKeieiHosei), meisaiNo, choSexCode)
            'REV_003 MOD END

            '生産年月
            'REV_003 MOD START
            'AddChosahyoData(chosahyoInsert, 調査票項目.生産年月(chosaBunrui)(0), meisaiNo, row("生年月日").ToString.Substring(0, 4))
            'AddChosahyoData(chosahyoInsert, 調査票項目.生産年月(chosaBunrui)(1), meisaiNo, row("生年月日").ToString.Substring(4, 2))
            dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.生産年月(chosaBunrui)(0 + gyunyuKeieiHosei), meisaiNo, row("生年月日").ToString.Substring(0, 4))
            dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.生産年月(chosaBunrui)(1 + gyunyuKeieiHosei), meisaiNo, row("生年月日").ToString.Substring(4, 2))
            'REV_003 MOD END

            'REV_002 SATRT------------ 
            '「全所有牛情報出力」により出力する帳票について、成畜年月を出力しないように修正する
            '成畜年月
            'If chosaBunrui = ComConst.牛トレサデータ.調査区分分類.牛乳 Then
            '    If ComUtil.Tresa.IsAdultTarget(choShuruiCode, choHinsyuCode) Then
            '        'REV_002 成畜年月を取得ロジック改修
            '        'Dim adultDate As String = ComUtil.Tresa.GetAdultDate(db, row("個体識別番号").ToString, row("生年月日").ToString, beginDate.ToString("yyyyMMdd"))

            '        If adultDate.Length = 8 Then
            '            Dim kobanYear As String = "Q11022501"
            '            Dim kobanMonth As String = "Q11022601"
            '            If CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
            '                kobanYear = "Q11022401"
            '                kobanMonth = "Q11022501"
            '            Else

            '            End If
            '            dtChosahyo.Rows.Add(chosaNen, censusNo, kobanYear, meisaiNo, adultDate.Substring(0, 4))
            '            dtChosahyo.Rows.Add(chosaNen, censusNo, kobanMonth, meisaiNo, adultDate.Substring(4, 2))
            '        End If
            '    End If
            'End If
            'REV_002 END------------ 

            '取得年月(牛乳)、転入(購入)　年月(子牛)、導入時　購入　年月(育成・肥育)
            If row("異動フラグ").ToString.Equals(ComConst.牛トレサデータ.異動フラグ.転入搬入.ToString) Then
                '異動フラグが 4(転入/搬入)の場合
                'REV_003 MOD START
                'AddChosahyoData(chosahyoInsert, 調査票項目.取得年月(chosaBunrui)(0), meisaiNo, row("異動年月日").ToString.Substring(0, 4))
                'AddChosahyoData(chosahyoInsert, 調査票項目.取得年月(chosaBunrui)(1), meisaiNo, row("異動年月日").ToString.Substring(4, 2))
                dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.取得年月(chosaBunrui)(0 + gyunyuKeieiHosei), meisaiNo, row("異動年月日").ToString.Substring(0, 4))
                dtChosahyo.Rows.Add(chosaNen, censusNo, 調査票項目.取得年月(chosaBunrui)(1 + gyunyuKeieiHosei), meisaiNo, row("異動年月日").ToString.Substring(4, 2))
                'REV_003 MOD END
            End If
            meisaiNo += 1
        Next

        Return dtChosahyo
    End Function

    ''' <summary>
    ''' 調査票データ（可変）に追加する
    ''' </summary>
    ''' <param name="dic"></param>
    ''' <param name="itemNo"></param>
    ''' <param name="meisaiNo"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Private Sub AddChosahyoData(ByRef dic As Dictionary(Of String, DAOChosahyo.調査票項目), itemNo As String, meisaiNo As Integer, value As String)
        Dim insertKey As String
        Dim insertData As New DAOChosahyo.調査票項目

        insertKey = GetChosahyoKey(itemNo, meisaiNo)
        insertData.値 = value

        'REV_002 Mod Start
        'dic.Add(insertKey, insertData)

        '重複キーがない場合にディクショナリに追加する。
        If Not dic.ContainsKey(insertKey) Then
            dic.Add(insertKey, insertData)
        End If
        'REV_002 Mod End

    End Sub

    ''' <summary>
    ''' 調査票データ（可変）のキーを取得
    ''' </summary>
    ''' <param name="itemNo"></param>
    ''' <param name="meisaiNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChosahyoKey(itemNo As String, meisaiNo As Integer) As String
        Dim ret As String = itemNo

        ret = ret & ComConst.ITEM_NO_DELIMITER & meisaiNo.ToString

        Return ret
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

    ''' <summary>
    ''' 編成対象の牛トレサデータ取得
    ''' </summary>
    ''' <param name="dtTresa">牛トレサデータ</param>
    ''' <returns>編成対象の牛トレサデータ</returns>
    ''' <remarks></remarks>
    Private Function GetHenseiTresaData(dtTresa As DataTable) As DataTable
        Dim dtTresaHensei As DataTable = dtTresa.Clone
        Dim tresaHenseiIndex As Integer = 0

        '個体識別番号をすべて抽出
        Dim dtKotai As DataTable = dtTresa.DefaultView.ToTable(True, "個体識別番号")
        For Each row As DataRow In dtKotai.Rows

            '個体識別番号ごとに牛の存在を確認
            Dim tresaRows As DataRow() = dtTresa.Select("個体識別番号 = '" & row("個体識別番号").ToString & "'")

            If tresaRows.Count <= 0 Then
                '編成対象外
                Continue For
            End If

            If ComUtil.OrganiezCheck.IsCowOrganize(tresaRows, tresaHenseiIndex) = False Then
                Continue For
            End If

            '編成対象に追加
            dtTresaHensei.ImportRow(tresaRows(tresaHenseiIndex))
        Next

        Return dtTresaHensei
    End Function

End Class
