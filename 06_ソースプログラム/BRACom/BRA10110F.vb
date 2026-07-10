'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_000   | 2023.01.13 |大興電子通信        | 要件No.5 新規作成
'//            |            |                    |
'//*************************************************************************************************

Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
''' <summary>
''' 標本リスト取込画面
''' </summary>
''' <remarks></remarks>
Public Class BRA10110F

    ''' <summary>
    ''' INSERT１回あたり行数
    ''' </summary>
    Private Const ROW_PER_INS As Integer = 1000

    ''' <summary>
    ''' 初期表示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BRA10110F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'センサス実施年設定
            HyohonUtil.SetCombobox(CboCensusNen, HyohonConst.センサス実施年リスト)
            CboCensusNen.SelectedIndex = 0
            '経営形態区分プルダウン設定
            HyohonUtil.SetCombobox(CboKeieiKeitai, HyohonUtil.Get経営形態区分リスト())
            '営農類型区分プルダウン設定
            HyohonUtil.SetCombobox(CboEinoRuikei, HyohonConst.営農類型区分リスト)
            '生産費区分プルダウン設定
            HyohonUtil.SetCombobox(CboSeisanhi, HyohonUtil.Get生産費区分リスト(HyohonConst.経営形態区分.識別対象外))
            '田畑区分プルダウン設定
            HyohonUtil.SetCombobox(CboTahata, HyohonConst.田畑区分リスト)
            '経営形態区分の空を選択
            CboKeieiKeitai.SelectedValue = HyohonConst.経営形態区分.識別対象外

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 経営形態区分プルダウン選択
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CboKeieiKeitai_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboKeieiKeitai.SelectedIndexChanged
        Try
            Dim keieiKeitai = DirectCast(HyohonUtil.GetComboValue(CboKeieiKeitai), HyohonConst.経営形態区分)
            Select Case keieiKeitai
                Case HyohonConst.経営形態区分.識別対象外
                    '空
                    CboEinoRuikei.SelectedValue = HyohonConst.営農類型区分.識別対象外
                    CboEinoRuikei.Enabled = False
                    CboSeisanhi.SelectedValue = HyohonConst.生産費区分.識別対象外
                    CboSeisanhi.Enabled = False
                Case HyohonConst.経営形態区分.個人経営体, HyohonConst.経営形態区分.法人経営体
                    '個人経営体、法人経営体
                    CboEinoRuikei.Enabled = True
                    CboSeisanhi.SelectedValue = HyohonConst.生産費区分.識別対象外
                    CboSeisanhi.Enabled = False
                Case Else
                    '個別経営体、組織法人経営体
                    CboEinoRuikei.SelectedValue = HyohonConst.営農類型区分.識別対象外
                    CboEinoRuikei.Enabled = False
                    HyohonUtil.SetCombobox(CboSeisanhi, HyohonUtil.Get生産費区分リスト(DirectCast(HyohonUtil.GetComboValue(CboKeieiKeitai), HyohonConst.経営形態区分)))
                    CboSeisanhi.Enabled = True
                    CboSeisanhi.SelectedValue = HyohonConst.生産費区分.識別対象外
            End Select

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 生産費区分プルダウン選択
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CboSeisanhi_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboSeisanhi.SelectedIndexChanged
        Try
            Dim keieiKeitai = DirectCast(HyohonUtil.GetComboValue(CboKeieiKeitai), HyohonConst.経営形態区分)
            Dim seisanhi = DirectCast(HyohonUtil.GetComboValue(CboSeisanhi), HyohonConst.生産費区分)
            If keieiKeitai = HyohonConst.経営形態区分.個別経営体 _
                AndAlso (seisanhi = HyohonConst.生産費区分.小麦 OrElse seisanhi = HyohonConst.生産費区分.大豆) Then
                '小麦、大豆
                CboTahata.Enabled = True
            Else
                '小麦、大豆以外
                CboTahata.SelectedValue = HyohonConst.生産費区分.識別対象外
                CboTahata.Enabled = False
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 取込ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnTorikomi_Click(sender As Object, e As EventArgs) Handles BtnTorikomi.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            '入力チェック
            '経営形態区分（必須）
            Dim cbo = CboKeieiKeitai
            If HyohonUtil.IsEmpty(cbo) Then
                cbo.Focus()
                Message.ShowMsgBox(MessageID.MSG_E_119, MsgBoxStyle.OkOnly)
                Return
            End If
            '営農類型区分（活性の場合、必須）
            cbo = CboEinoRuikei
            If cbo.Enabled AndAlso HyohonUtil.IsEmpty(cbo) Then
                cbo.Focus()
                Message.ShowMsgBox(MessageID.MSG_E_120, MsgBoxStyle.OkOnly)
                Return
            End If
            '生産費区分（活性の場合、必須）
            cbo = CboSeisanhi
            If cbo.Enabled AndAlso HyohonUtil.IsEmpty(cbo) Then
                cbo.Focus()
                Message.ShowMsgBox(MessageID.MSG_E_121, MsgBoxStyle.OkOnly)
                Return
            End If
            '田畑区分（活性の場合、必須）
            cbo = CboTahata
            If cbo.Enabled AndAlso HyohonUtil.IsEmpty(cbo) Then
                cbo.Focus()
                Message.ShowMsgBox(MessageID.MSG_E_140, MsgBoxStyle.OkOnly)
                Return
            End If

            '標本リストCSV選択
            Dim filePath = ComUtil.GetFilePath(Of OpenFileDialog)(Me, IniFileInfo.ExcelInPath, filter:="CSVファイル|*.csv")
            If String.IsNullOrEmpty(filePath) Then
                Return
            End If

            '進捗ダイアログ
            Dim progressDialog As New ProgressDialog()
            Dim ret = False
            Try
                Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    '進捗ダイアログを表示する
                    progressDialog.Maximum = 100 + 40
                    progressDialog.Show(Me)
                    'センサス実施年
                    Dim censusNen = CInt(CboCensusNen.SelectedValue)
                    '区分情報
                    Dim kubunClearInfo = HyohonConst.標本リスト区分クリア情報(censusNen)
                    Dim kubunCheckInfo = HyohonConst.標本リスト区分チェック情報(censusNen)
                    '標本リスト表頭
                    Dim hyohonHeaders = HyohonConst.標本リスト表頭名(censusNen)
                    Dim hyohonCount = UBound(hyohonHeaders)
                    'テーブルの列情報取得
                    Dim colInfo = HyohonUtil.GetColInfo(censusNen)

                    '取込用一時テーブル（固定長）作成
                    db.ExecuteNonQuery("CREATE TABLE #標本リストFIX(D NVARCHAR(4000))")

                    '標本リストCSV読込
                    Using fs = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                        Dim length = fs.Length
                        If length = 0 Then
                            progressDialog.ShowMsgForm(MessageID.MSG_E_122, {"1行目:0", hyohonCount.ToString()})
                            Return
                        End If
                        Dim progressValue = 0
                        Using sr = New StreamReader(fs, Encoding.GetEncoding("shift-jis"))
                            Dim countAll = 1
                            '表頭読み飛ばし
                            Dim headerCount = sr.ReadLine().Split(","c).Select(Function(y) y.Trim()).ToList().Count
                            If headerCount <> hyohonCount Then
                                progressDialog.ShowMsgForm(MessageID.MSG_E_122, {"1行目:" & headerCount, hyohonCount.ToString()})
                                Return
                            End If
                            Do While Not sr.EndOfStream
                                Dim csvList = New List(Of List(Of String))
                                Dim count = 0
                                '最大10000行ずつ一時テーブルへ格納
                                Do While Not sr.EndOfStream AndAlso count < 10000
                                    csvList.Add(sr.ReadLine().Split(","c).Select(Function(y) y.Trim()).ToList())
                                    count += 1
                                Loop
                                countAll += count
                                '進捗を進める
                                Dim progressAll = CInt(Math.Truncate(fs.Position / length * 100))
                                Dim progress = CInt(Math.Truncate((progressAll - progressValue) / 8))
                                If progress > 0 Then
                                    progressValue = progressAll
                                End If
                                progressDialog.AddValue = progress

                                '標本リスト表頭名配列と列数が一致しない場合、エラー
                                Dim unmatchCount = csvList.Select(Function(x, index) (val:=x, index)) _
                                    .Where(Function(x) x.val.Count <> hyohonCount).FirstOrDefault()
                                If countAll = 1 Then
                                    unmatchCount = (New List(Of String)(), 0)
                                End If
                                If unmatchCount.val IsNot Nothing Then
                                    progressDialog.ShowMsgForm(MessageID.MSG_E_122, {unmatchCount.index + 1 + countAll - count & "行目:" & unmatchCount.val.Count.ToString(), hyohonCount.ToString()})
                                    Return
                                End If
                                '進捗を進める
                                progressDialog.AddValue = progress

                                '列情報に基づくチェック
                                For Each info In colInfo
                                    '必須
                                    If Not info.IS_NULLABLE Then
                                        Dim err = csvList.Select(Function(x, index) (key:=x(0) ,val:=x((info.COLUMN_ID - 1)), index)).Where(Function(x) String.IsNullOrEmpty(x.val)).Take(10)
                                        If err.Count > 0 Then
                                            If info.COLUMN_ID = 1 Then
                                                progressDialog.ShowMsgForm(MessageID.MSG_E_127, {"※10件まで表示" & vbCrLf & String.Join(vbCrLf, err.Select(Function(x) x.index + 1 + countAll - count & "行目").ToArray()) & "の" & hyohonHeaders(1)})
                                            Else
                                                progressDialog.ShowMsgForm(MessageID.MSG_E_124, {info.COLUMN_ID & "．" & hyohonHeaders(info.COLUMN_ID) & "が入力されていません。", String.Join(vbCrLf, err.Select(Function(x) x.key).ToArray())})
                                            End If
                                            Return
                                        End If
                                    End If
                                    '型
                                    If info.TYPE = "NUMERIC" Then
                                        Dim decVal As Decimal
                                        Dim kiboNo = HyohonUtil.Get標本リスト列番号(censusNen, HyohonConst.標本リスト列番号キー.営農類型区分規模)
                                        Dim err = csvList.Where(Function(x) Not String.IsNullOrEmpty(x(info.COLUMN_ID - 1)) _
                                                AndAlso (Not Decimal.TryParse(x(info.COLUMN_ID - 1), decVal) OrElse info.COLUMN_ID <> kiboNo AndAlso decVal < 0)).Take(10)
                                        If err.Count > 0 Then
                                            progressDialog.ShowMsgForm(MessageID.MSG_E_124, {info.COLUMN_ID & "．" & hyohonHeaders(info.COLUMN_ID) & String.Format("は{0}数字で入力してください。", If(info.COLUMN_ID = kiboNo, "", "正の")), String.Join(vbCrLf, err.Select(Function(x) x(0) & String.Format(" 入力値[{0}]", x(info.COLUMN_ID - 1))))})
                                            Return
                                        End If
                                    End If
                                    '桁数
                                    Select Case info.TYPE
                                        Case "CHAR"
                                            Dim err = csvList.Where(Function(x) x(info.COLUMN_ID - 1).Length <> info.MAX_LENGTH).Take(10)
                                            If err.Count > 0 Then
                                                progressDialog.ShowMsgForm(MessageID.MSG_E_124, {info.COLUMN_ID & "．" & hyohonHeaders(info.COLUMN_ID) & String.Format("は{0}文字で入力してください。", info.MAX_LENGTH), String.Join(vbCrLf, err.Select(Function(x) x(0) & String.Format(" 入力値[{0}]", x(info.COLUMN_ID - 1))))})
                                                Return
                                            End If
                                        Case "VARCHAR", "NVARCHAR"
                                            Dim err = csvList.Where(Function(x) x(info.COLUMN_ID - 1).Length > info.MAX_LENGTH).Take(10)
                                            If err.Count > 0 Then
                                                progressDialog.ShowMsgForm(MessageID.MSG_E_124, {info.COLUMN_ID & "．" & hyohonHeaders(info.COLUMN_ID) & String.Format("は{0}文字以内で入力してください。", info.MAX_LENGTH), String.Join(vbCrLf, err.Select(Function(x) x(0) & String.Format(" 入力値[{0}]", x(info.COLUMN_ID - 1))))})
                                                Return
                                            End If
                                        Case "NUMERIC"
                                            If info.SCALE > 0 Then
                                                Dim err = csvList.Where(Function(x) Not String.IsNullOrEmpty(x(info.COLUMN_ID - 1)) AndAlso (Math.Truncate(CDec(x(info.COLUMN_ID - 1))).ToString().Replace("-", "").Length > info.PRECISION OrElse (Decimal.GetBits(CDec(x(info.COLUMN_ID - 1)))(3) >> 16 And &HFF) > info.SCALE)).Take(10)
                                                If err.Count > 0 Then
                                                    progressDialog.ShowMsgForm(MessageID.MSG_E_124, {info.COLUMN_ID & "．" & hyohonHeaders(info.COLUMN_ID) & String.Format("は整数:{0}桁、小数:{1}桁以内で入力してください。", info.PRECISION, info.SCALE), String.Join(vbCrLf, err.Select(Function(x) x(0) & String.Format(" 入力値[{0}]", x(info.COLUMN_ID - 1))))})
                                                    Return
                                                End If
                                            Else
                                                Dim err = csvList.Where(Function(x) Not String.IsNullOrEmpty(x(info.COLUMN_ID - 1)) AndAlso (x(info.COLUMN_ID - 1).Contains(".") OrElse CDec(x(info.COLUMN_ID - 1)).ToString().Replace("-", "").Length > info.PRECISION)).Take(10)
                                                If err.Count > 0 Then
                                                    progressDialog.ShowMsgForm(MessageID.MSG_E_124, {info.COLUMN_ID & "．" & hyohonHeaders(info.COLUMN_ID) & String.Format("は整数:{0}桁以内で入力してください。", info.PRECISION), String.Join(vbCrLf, err.Select(Function(x) x(0) & String.Format(" 入力値[{0}]", x(info.COLUMN_ID - 1))))})
                                                    Return
                                                End If
                                            End If
                                        Case Else
                                                Throw New Exception("CHAR,VARCHAR,NVARCHAR,NUMERIC以外の型は未実装です。")
                                    End Select
                                Next
                                '進捗を進める
                                progressDialog.AddValue = progress

                                '一連番号の重複をチェック
                                Dim duplicate = csvList.GroupBy(Function(x) x(0)).Where(Function(x) x.Count() > 1).Select(Function(x) x.Key).Take(10)
                                If duplicate.Count > 0 Then
                                    progressDialog.ShowMsgForm(MessageID.MSG_E_124, {"CSV内で重複しています。", String.Join(vbCrLf, duplicate)})
                                    Return
                                End If
                                '進捗を進める
                                progressDialog.AddValue = progress

                                '区分情報に基づき無効な値をクリア
                                For Each info In kubunClearInfo
                                    If info.RANGE.Contains("～") Then
                                        Dim vals = Split(info.RANGE, "～")
                                        csvList.Where(Function(x) Not String.IsNullOrEmpty(x(info.NO - 1)) AndAlso CDec(x(info.NO - 1)) < CDec(vals(0)) OrElse CDec(x(info.NO - 1)) > CDec(vals(1))).ToList().ForEach(Sub(x) x(info.NO - 1) = "")
                                    Else
                                        csvList.Where(Function(x) Not String.IsNullOrEmpty(x(info.NO - 1)) AndAlso Not info.RANGE.Split(","c).Contains(x(info.NO - 1))).ToList().ForEach(Sub(x) x(info.NO - 1) = "")
                                    End If
                                Next
                                '進捗を進める
                                progressDialog.AddValue = progress

                                '区分情報に基づくチェック
                                For Each info In kubunCheckInfo
                                    Dim err As IEnumerable(Of List(Of String))
                                    If info.RANGE.Contains("～") Then
                                        Dim vals = Split(info.RANGE, "～")
                                        err = csvList.Where(Function(x) Not String.IsNullOrEmpty(x(info.NO - 1)) AndAlso CDec(x(info.NO - 1)) < CDec(vals(0)) OrElse CDec(x(info.NO - 1)) > CDec(vals(1))).Take(10)
                                    Else
                                        err = csvList.Where(Function(x) Not String.IsNullOrEmpty(x(info.NO - 1)) AndAlso Not info.RANGE.Split(","c).Contains(x(info.NO - 1))).Take(10)
                                    End If
                                    If err.Count > 0 Then
                                        progressDialog.ShowMsgForm(MessageID.MSG_E_124, {info.NO & "．" & hyohonHeaders(info.NO) & String.Format("は[{0}]で入力してください。", info.RANGE), String.Join(vbCrLf, err.Select(Function(x) x(0) & String.Format(" 入力値[{0}]", x(info.NO - 1))))})
                                        Return
                                    End If
                                Next
                                '進捗を進める
                                progressDialog.AddValue = progress

                                '営農類型区分に基づくチェック
                                If CboEinoRuikei.Enabled Then
                                    Dim einoRuikei = HyohonUtil.GetComboValue(CboEinoRuikei).ToString()
                                    Dim einoNo = HyohonUtil.Get標本リスト列番号(censusNen, HyohonConst.標本リスト列番号キー.営農類型区分)
                                    Dim err = csvList.Where(Function(x) x(einoNo - 1) <> einoRuikei).Take(10)
                                    If err.Count > 0 Then
                                        progressDialog.ShowMsgForm(MessageID.MSG_E_124, {einoNo & "．" & hyohonHeaders(einoNo) & String.Format("は[{0}]を入力してください。", einoRuikei), String.Join(vbCrLf, err.Select(Function(x) x(0) & String.Format(" 入力値[{0}]", x(einoNo - 1))))})
                                        Return
                                    End If
                                End If
                                '進捗を進める
                                progressDialog.AddValue = progress

                                '取込用一時テーブル（固定長）に登録
                                Dim flg = True
                                Dim sbSql = New StringBuilder()
                                For i = 0 To csvList.Count - 1
                                    If flg Then
                                        sbSql.AppendLine("INSERT INTO #標本リストFIX VALUES")
                                        flg = False
                                    Else
                                        sbSql.AppendLine(",")
                                    End If
                                    sbSql.Append("('")
                                    ' 符号分でMAX_LENGTH+1
                                    sbSql.Append(String.Join("", csvList(i).Select(Function(x, index) If(Not String.IsNullOrEmpty(x) AndAlso colInfo(index).TYPE = "NUMERIC", CDec(x).ToString(), x).PadRight(colInfo(index).MAX_LENGTH + 1, " "c)).ToArray()))
                                    sbSql.Append("')")
                                    If (i + 1) Mod ROW_PER_INS = 0 OrElse i >= csvList.Count - 1 Then
                                        db.ExecuteNonQuery(sbSql.ToString(), logOutput:=False)
                                        sbSql = New StringBuilder()
                                        flg = True
                                    End If
                                Next
                                '進捗を進める
                                progressDialog.Value = progressValue
                            Loop
                        End Using
                    End Using
                    '進捗を進める
                    progressDialog.Value = 100
                    '取込実行
                    ret = ExecuteTorikomi(db, censusNen, colInfo, progressDialog)
                End Using
            Finally
                If progressDialog IsNot Nothing Then
                    '進捗ダイアログを閉じる
                    progressDialog.endDispose()
                    progressDialog = Nothing
                End If
            End Try

            '完了メッセージ
            If ret Then
                Message.ShowMsgBox(MessageID.MSG_I_006, MsgBoxStyle.OkOnly)
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 農政局、都道府県、実査設置拠点更新SQL
    ''' </summary>
    Private Const SQL_UPDATE_KANKATSU =
        "UPDATE #標本リストCSV SET" _
        & " 農政局 = (SELECT JM.局番号 FROM COMM.dbo.事務所名マスタ JM WHERE JM.事務所番号 = IIF(NO2 = 1, 51, NO2))," _
        & " 都道府県 = NO2," _
        & " 実査設置拠点 = (SELECT KM.実査設置拠点 FROM 標本リスト＿拠点マスタ KM" _
        & " WHERE KM.都道府県 = NO2" _
        & " AND (KM.市区町村 = NO4 OR KM.市区町村 = 0))"

    ''' <summary>
    ''' 管轄チェックSQL
    ''' </summary>
    Private Const SQL_SELECT_KANKATSU_CHECK =
        "SELECT TOP(10) NO1 FROM #標本リストCSV" _
        & " WHERE ISNULL(農政局, -1) <> @農政局" _
        & " OR ISNULL(都道府県, -1) <> IIF(@都道府県 = 0, 都道府県, @都道府県)" _
        & " OR ISNULL(実査設置拠点, -1) <> IIF(@実査設置拠点 = 0, 実査設置拠点, @実査設置拠点)"

    ''' <summary>
    ''' CSV取込SQL
    ''' </summary>
    Private Const SQL_MERGE_1 =
        "MERGE INTO 標本リスト{0} HL" _
        & " USING #標本リストCSV CSV" _
        & " ON HL.NO1 = CSV.NO1" _
        & " AND HL.農政局 =  CSV.農政局" _
        & " AND HL.都道府県 = CSV.都道府県" _
        & " AND HL.実査設置拠点 = CSV.実査設置拠点" _
        & " AND HL.経営形態区分 = @経営形態区分" _
        & " AND HL.営農類型区分 = @営農類型区分" _
        & " AND HL.生産費区分 = @生産費区分" _
        & " AND HL.田畑区分 = @田畑区分" _
        & " WHEN MATCHED THEN" _
        & " UPDATE SET"
    Private Const SQL_MERGE_2 =
        " HL.更新日付 = SYSDATETIME()," _
        & " HL.更新者ID = @更新者ID" _
        & " WHEN NOT MATCHED THEN" _
        & " INSERT VALUES("
    Private Const SQL_MERGE_3 =
        " CSV.農政局," _
        & " CSV.都道府県," _
        & " CSV.実査設置拠点," _
        & " @経営形態区分," _
        & " @営農類型区分," _
        & " @生産費区分," _
        & " @田畑区分," _
        & " SYSDATETIME()," _
        & " @更新者ID);"
    ''' <summary>
    ''' 取込実行
    ''' </summary>
    ''' <param name="censusNen"></param>
    ''' <param name="colInfo"></param>
    ''' <param name="progressDialog"></param>
    Private Function ExecuteTorikomi(db As DBAccess, censusNen As Integer, colInfo As (COLUMN_ID As Integer, NAME As String, IS_NULLABLE As Boolean, TYPE As String, MAX_LENGTH As Integer, PRECISION As Integer, SCALE As Integer)(), progressDialog As ProgressDialog) As Boolean

        '取込用一時テーブル（CSV）作成
        Dim sbSql = New StringBuilder()
        sbSql.Append("CREATE TABLE #標本リストCSV(")
        For Each info In colInfo
            Select Case info.TYPE
                Case "CHAR", "VARCHAR", "NVARCHAR"
                    sbSql.Append(String.Format(" {0} {1}({2}),", info.NAME, info.TYPE, info.MAX_LENGTH))
                Case "NUMERIC"
                    sbSql.Append(String.Format(" {0} {1}({2},{3}),", info.NAME, info.TYPE, info.PRECISION + info.SCALE, info.SCALE))
            End Select
        Next
        sbSql.Append("農政局 int,")
        sbSql.Append("都道府県 int,")
        sbSql.Append("実査設置拠点 int)")
        db.ExecuteNonQuery(sbSql.ToString())

        '固定長を切り取って登録
        sbSql = New StringBuilder
        sbSql.Append("INSERT INTO #標本リストCSV SELECT")
        Dim startPosition = 1
        For Each info In colInfo
            If info.TYPE = "NUMERIC" Then
                sbSql.Append(String.Format(" NULLIF(TRIM(SUBSTRING(D,{0},{1})),''),", startPosition, info.MAX_LENGTH + 1))
            Else
                sbSql.Append(String.Format(" TRIM(SUBSTRING(D,{0},{1})),", startPosition, info.MAX_LENGTH + 1))
            End If
            startPosition += info.MAX_LENGTH + 1
        Next
        sbSql.Append(" null, null, null")
        sbSql.Append(" FROM #標本リストFIX")
        db.ExecuteNonQuery(sbSql.ToString())
        '進捗を進める
        progressDialog.AddValue = 10

        '農政局,都道府県,実査設置拠点を更新
        db.ExecuteNonQuery(SQL_UPDATE_KANKATSU)
        '進捗を進める
        progressDialog.AddValue = 10

        '農政局工程、実査設置拠点工程の場合、管轄と一致するかを確認
        If CommonInfo.Koutei <> CommonInfo.KouteiKubun.Code.Honsyo Then
            Dim p = New List(Of DBAccess.Parameter)
            p.Add(db.CreateParameter("@農政局", SqlDbType.Int, CommonInfo.Kyoku))
            p.Add(db.CreateParameter("@都道府県", SqlDbType.Int, If(CommonInfo.Jimusyo = "51", "1", CommonInfo.Jimusyo)))
            p.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            Dim dt = db.GetDataTable(SQL_SELECT_KANKATSU_CHECK, p)
            If dt.Rows.Count > 0 Then
                progressDialog.ShowMsgForm(MessageID.MSG_E_124, {"都道府県、市区町村が管轄と異なります。", String.Join(vbCrLf, dt.Rows.Cast(Of DataRow).Select(Function(x) x("NO1").ToString()).ToArray())})
                Return False
            End If
        End If
        '進捗を進める
        progressDialog.AddValue = 10

        Try
            db.BeginTrans()

            '標本リストyyyyテーブルにMERGE
            sbSql = New StringBuilder()
            sbSql.Append(String.Format(SQL_MERGE_1, censusNen))
            For i = 2 To colInfo.Count
                sbSql.Append(String.Format(" HL.NO{0} = CSV.NO{0},", i))
            Next
            sbSql.Append(SQL_MERGE_2)
            For i = 1 To colInfo.Count
                sbSql.Append(String.Format(" CSV.NO{0},", i))
            Next
            sbSql.Append(SQL_MERGE_3)
            Dim para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@経営形態区分", SqlDbType.Int, HyohonUtil.GetComboValue(CboKeieiKeitai)))
            para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, HyohonUtil.GetComboValue(CboEinoRuikei)))
            para.Add(db.CreateParameter("@生産費区分", SqlDbType.Int, HyohonUtil.GetComboValue(CboSeisanhi)))
            para.Add(db.CreateParameter("@田畑区分", SqlDbType.Int, HyohonUtil.GetComboValue(CboTahata)))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))
            db.ExecuteNonQuery(sbSql.ToString(), para)

            '進捗を進める
            progressDialog.AddValue = 10

            db.CommitTrans()

            Return True
        Catch se As SqlException
            db.RollBackTrans()
            If se.Number = 2627 Then
                '一意制約エラーの場合はメッセージ表示
                Dim key = New RegularExpressions.Regex("a[0-9]{16}").Match(se.Message).Value
                progressDialog.ShowMsgForm(MessageID.MSG_E_125, {key})
                Return False
            End If
            Throw se
        Catch ex As Exception
            db.RollBackTrans()
            Throw ex
        End Try

    End Function

End Class
