'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2023.01.13 |大興電子通信        | フェーズ2 要件No.3
'//            |            |                    |
'//*************************************************************************************************

Imports System.IO
Imports System.Text

''' <summary>
''' 個別結果表CSV出力画面
''' </summary>
''' <remarks></remarks>
Public Class BRA3310F

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
    ''' <summary>
    ''' 調査票区分(REV_001)
    ''' </summary>
    Private _chosahyo As String

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA3310F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'REV_001↓
            '出力形式プルダウン設定
            If CInt(CommonInfo.Chosakubun) >= CInt(ComConst.調査区分.経営分析調査_二条大麦生産費) Then
                lblOutType.Visible = False
                CboOutType.Visible = False
            End If
            '調査票区分プルダウン設定
            '※調査年コンボボックス設定より前に設定
            CboChosahyo.ValueMember = "Key"
            CboChosahyo.DisplayMember = "Value"
            CboChosahyo.DataSource = New List(Of KeyValuePair(Of String, String)) From {
                    New KeyValuePair(Of String, String)("", ""),
                    New KeyValuePair(Of String, String)("1", "基本調査票"),
                    New KeyValuePair(Of String, String)("2", "詳細調査票")
                }
            CboChosahyo.SelectedIndex = 0
            If CommonInfo.Chosakubun <> ComConst.調査区分.営農類型別経営統計_個人 Then
                lblChosahyo.Visible = False
                CboChosahyo.Visible = False
            End If
            'REV_001↑

            '調査年コンボボックス設定
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.KobetsuKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
            End Using

            '局コンボボックス設定
            ComUtil.SetKyokuComboBox(cboKyoku)

            '営農類型コンボボックス設定
            ComUtil.SetEinouRuikeiComboBox(lblEinouRuikei, cboEinouRuikei)

            '欠測値補完コンボボックス設定
            ComUtil.SetKessokuchiHokanComboBox(lblKessokuchiHokan, cboKessokuchiHokan)

            '貸借対照表区分コンボボックス設定
            ComUtil.SetTaishakuTaishohyoComboBox(lblTaishakuTaishohyo, cboTaishakuTaishohyo)

            'DataGridView設定
            ComUtil.ConfigDgv(Me.dgvList)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 表示ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        Try
            '調査年選択チェック
            If cboChosaNen.SelectedValue Is Nothing Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            _chosaNen = cboChosaNen.SelectedValue.ToString
            _kyoku = If(IsDBNull(cboKyoku.SelectedValue), Nothing, cboKyoku.SelectedValue.ToString)
            _jimusho = If(IsDBNull(cboKyoten.SelectedValue), Nothing, CStr(CType(cboKyoten.SelectedItem, DataRowView)("事務所番号")))
            _kyoten = If(IsDBNull(cboKyoten.SelectedValue), Nothing, cboKyoten.SelectedValue.ToString)
            _einouRuikei = If(cboEinouRuikei.SelectedValue Is Nothing, Nothing, cboEinouRuikei.SelectedValue.ToString)
            _kessokuchiHokan = ComUtil.GetkessokuchiHokana(cboKessokuchiHokan)
            _taishakuTaishohyo = If(cboTaishakuTaishohyo.SelectedValue Is Nothing, Nothing, cboTaishakuTaishohyo.SelectedValue.ToString)
            'REV_001↓
            _chosahyo = If(CboChosahyo.SelectedValue?.ToString(), "")
            'REV_001↑

            '一覧表示
            Me.ShowList(_chosaNen, _kyoku, _jimusho, _kyoten, _einouRuikei, _taishakuTaishohyo, _kessokuchiHokan)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 出力ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '個別結果表主キー取得
            Dim keys As List(Of ValueTuple(Of DAOKobetsuKekkahyo.PrimaryKey, DAOKobetsuKekkahyo.KyotenKey))
            keys = Me.GetKobetsuKekkahyoKey(_chosaNen)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(keys, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            'フォルダパス取得
            Dim folderPath As String = ComUtil.GetFolderPath(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath))

            If folderPath.Equals(String.Empty) Then
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_028, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                'REV_001↓
                Dim outType = DirectCast(CboOutType.SelectedValue, 出力形式)
                Dim sqlInfoList As List(Of (String, String)) = Nothing
                If outType <> 出力形式.分割 Then
                    Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                        sqlInfoList = GetSelect統合個別結果表SQL(db, _chosaNen, outType, _kessokuchiHokan)
                    End Using
                End If
                'REV_001↑

                '個別結果表データ取得
                Dim dcs As Dictionary(Of String, DataTable) = Nothing
                For Each key In keys
                    Dim pKey As DAOKobetsuKekkahyo.PrimaryKey = key.Item1

                    Dim dc As Dictionary(Of String, DataTable)
                    Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                        'REV_001↓
                        'dc = DAOKobetsuKekkahyo.GetTable(db, pKey, _kessokuchiHokan)
                        If outType = 出力形式.分割 Then
                            dc = DAOKobetsuKekkahyo.GetTable(db, pKey, _kessokuchiHokan)
                        Else
                            dc = Select統合個別結果表(db, sqlInfoList, pKey)

                        End If
                        'REV_001↑
                    End Using

                    If dcs Is Nothing Then
                        dcs = dc
                    Else
                        For Each _dc As KeyValuePair(Of String, DataTable) In dc
                            Dim tableName As String = _dc.Key
                            Dim dt As DataTable = _dc.Value
                            dcs(tableName).ImportRow(dt(0))
                        Next
                    End If
                Next

                'ファイル名取得
                Dim fileNamePattern As String = Me.GetFileNamePattern(_chosaNen)

                'CSV出力
                Dim ret As ComConst.CSVファイル.enmOutputReturn = Me.PutCSV(folderPath, fileNamePattern, dcs)

                Select Case ret
                    Case ComConst.CSVファイル.enmOutputReturn.OK
                        '完了メッセージ
                        Message.ShowMsgBox(MessageID.MSG_I_002, MsgBoxStyle.OkOnly)
                    Case ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS
                        'エラーメッセージ
                        Message.ShowMsgBox(MessageID.MSG_E_058, MsgBoxStyle.OkOnly)
                End Select
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 全選択ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAllSelect_Click(sender As Object, e As EventArgs) Handles btnAllSelect.Click
        Try
            ComUtil.SetDataGridViewAllCheck(dgvList, True)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 全解除ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAllCancel_Click(sender As Object, e As EventArgs) Handles btnAllCancel.Click
        Try
            ComUtil.SetDataGridViewAllCheck(dgvList, False)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 農政局コンボボックス選択時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboKyoku_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKyoku.SelectedIndexChanged
        Try
            '拠点コンボボックス設定
            ComUtil.SetKyotenComboBox(cboKyoku, cboKyoten)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 欠測値補完コンボボックス選択時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboKessokuchiHokan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKessokuchiHokan.SelectedIndexChanged
        Try
            dgvList.Rows.Clear()

            '欠測値補完取得
            Dim kessokuchiHokan As String = ComUtil.GetkessokuchiHokana(cboKessokuchiHokan)

            '調査年コンボボックス設定
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.KobetsuKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center, kessokuchiHokan)
            End Using
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 個別結果表キー取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetKobetsuKekkahyoKey(chosaNen As String) As List(Of ValueTuple(Of DAOKobetsuKekkahyo.PrimaryKey, DAOKobetsuKekkahyo.KyotenKey))
        Dim ret As New List(Of ValueTuple(Of DAOKobetsuKekkahyo.PrimaryKey, DAOKobetsuKekkahyo.KyotenKey))

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                Dim pkey As DAOKobetsuKekkahyo.PrimaryKey = New DAOKobetsuKekkahyo.PrimaryKey(chosaNen, dgvList.Rows(i).Cells(7).Value.ToString)
                Dim kkey As DAOKobetsuKekkahyo.KyotenKey = New DAOKobetsuKekkahyo.KyotenKey(dgvList.Rows(i).Cells(8).Value.ToString, dgvList.Rows(i).Cells(9).Value.ToString, dgvList.Rows(i).Cells(10).Value.ToString)
                ret.Add(ValueTuple.Create(pkey, kkey))
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="einouRuikei"></param>
    ''' <param name="taishakuTaishohyo"></param>
    ''' <param name="kessokuchiHokan"></param>
    ''' <remarks></remarks>
    Private Sub ShowList(chosaNen As String, kyoku As String, jimusho As String, kyoten As String, einouRuikei As String, taishakuTaishohyo As String, kessokuchiHokan As String)
        Dim dt As DataTable

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            'REV_001↓
            'dt = DAOKobetsuKekkahyo.GetList(db, chosaNen, kyoku, jimusho, kyoten, einouRuikei, taishakuTaishohyo, kessokuchiHokan)
            dt = DAOKobetsuKekkahyo.GetList(db, chosaNen, kyoku, jimusho, kyoten, einouRuikei, taishakuTaishohyo, kessokuchiHokan, _chosahyo)
            'REV_001↑
        End Using

        dgvList.Rows.Clear()

        For Each row As DataRow In dt.Rows
            Dim censusNo As String = row("センサス番号").ToString

            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(1).Value = ComUtil.GetTodofuken(censusNo)
            dgvList.Rows(i).Cells(2).Value = ComUtil.GetShikuchoson(censusNo)
            dgvList.Rows(i).Cells(3).Value = ComUtil.GetKyuShikuchoson(censusNo)
            dgvList.Rows(i).Cells(4).Value = ComUtil.GetNogyoShuraku(censusNo)
            dgvList.Rows(i).Cells(5).Value = ComUtil.GetChosaku(censusNo)
            dgvList.Rows(i).Cells(6).Value = ComUtil.GetKyakutaiNo(censusNo)
            dgvList.Rows(i).Cells(7).Value = censusNo
            dgvList.Rows(i).Cells(8).Value = row("農政局").ToString
            dgvList.Rows(i).Cells(9).Value = row("都道府県").ToString
            dgvList.Rows(i).Cells(10).Value = row("実査設置拠点").ToString
        Next
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="keys"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(keys As List(Of ValueTuple(Of DAOKobetsuKekkahyo.PrimaryKey, DAOKobetsuKekkahyo.KyotenKey)), ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'センサス番号選択チェック
        If keys.Count = 0 Then
            msgId = MessageID.MSG_E_003
            Return ret
        End If

        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' CSV出力
    ''' </summary>
    ''' <param name="outDir"></param>
    ''' <param name="fileName"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PutCSV(outDir As String, fileName As String, dc As Dictionary(Of String, DataTable)) As ComConst.CSVファイル.enmOutputReturn
        Dim ret As ComConst.CSVファイル.enmOutputReturn = ComConst.CSVファイル.enmOutputReturn.CANCEL

        Dim sjisEnc As Encoding = Encoding.GetEncoding(ComConst.CSVファイル.CODEPAGE_SHIFT_JIS)

        Dim filePathTemp As String = System.IO.Path.Combine(outDir, fileName)

        If Not System.IO.Directory.Exists(outDir) Then
            Directory.CreateDirectory(outDir)
        End If

        'REV_001↓
        Dim outType = DirectCast(CboOutType.SelectedValue, 出力形式)
        Dim itemNames As List(Of (項目番号 As String, 項目名 As String)) = Nothing
        If outType <> 出力形式.分割 Then
            itemNames = GetItemNames(_chosaNen)
        End If
        'REV_001↑

        For Each kv As KeyValuePair(Of String, DataTable) In dc
            Dim tableName As String = kv.Key
            Dim dt As DataTable = kv.Value

            Try
                Using sw As New System.IO.StreamWriter(String.Format(filePathTemp, tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, _kessokuchiHokan)), False, sjisEnc)
                    Dim colArr(dt.Columns.Count - 1) As Object
                    dt.Columns.CopyTo(colArr, 0)
                    sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, colArr) & ComConst.CSVファイル.END_ADDITION)

                    'REV_001↓
                    '統合版は２行目に項目名出力
                    If outType <> 出力形式.分割 Then
                        Dim nameArr = dt.Columns.Cast(Of DataColumn).Select(Function(x) If(itemNames.Where(Function(y) y.項目番号 = x.ColumnName).Select(Function(y) y.項目名).FirstOrDefault(), x.ColumnName)).ToArray()
                        sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, nameArr) & ComConst.CSVファイル.END_ADDITION)
                    End If
                    'REV_001↑

                    For Each row As DataRow In dt.Rows
                        Dim arr As Object() = row.ItemArray().ToArray()
                        sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, arr) & ComConst.CSVファイル.END_ADDITION)
                    Next
                End Using
            Catch ex As Exception
                ret = ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS
                Return ret
            End Try
        Next

        ret = ComConst.CSVファイル.enmOutputReturn.OK

        Return ret
    End Function

    ''' <summary>
    ''' ファイル名パターン取得
    ''' </summary>
    ''' <param name="year"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetFileNamePattern(year As String) As String
        Return "{0}_" & year & ".csv"
    End Function

    ''' <summary>
    ''' 出力形式(REV_001)
    ''' </summary>
    Private Enum 出力形式
        分割
        統合
        個人_1_3
        農経個人_4_6
        農経法人_4_6
        法人_1_6_7
        法人_1_3_7
    End Enum
    ''' <summary>
    ''' 出力形式リスト（営農個人）(REV_001)
    ''' </summary>
    Private Shared ReadOnly 出力形式リスト＿営農個人 As New List(Of KeyValuePair(Of Integer, String)) From {
        New KeyValuePair(Of Integer, String)(出力形式.分割, "個人経営体No.1,2,3,4,5,6"),
        New KeyValuePair(Of Integer, String)(出力形式.統合, "個人経営体No.1～6"),
        New KeyValuePair(Of Integer, String)(出力形式.個人_1_3, "個人経営体No.1～3"),
        New KeyValuePair(Of Integer, String)(出力形式.農経個人_4_6, "農業経営体No.1～3")
    }
    ''' <summary>
    ''' 出力形式リスト（営農法人）(REV_001)
    ''' </summary>
    Private Shared ReadOnly 出力形式リスト＿営農法人 As New List(Of KeyValuePair(Of Integer, String)) From {
        New KeyValuePair(Of Integer, String)(出力形式.分割, "法人経営体No.1,2,3,4,5,6,7"),
        New KeyValuePair(Of Integer, String)(出力形式.法人_1_6_7, "法人経営体No.1～6,7"),
        New KeyValuePair(Of Integer, String)(出力形式.法人_1_3_7, "法人経営体No.1～3,7"),
        New KeyValuePair(Of Integer, String)(出力形式.農経法人_4_6, "農業経営体No.1～3")
    }
    ''' <summary>
    ''' 出力形式リスト（個別テーブル３つ）(REV_001)
    ''' </summary>
    Private Shared ReadOnly 出力形式リスト＿個別テーブル３ As New List(Of KeyValuePair(Of Integer, String)) From {
        New KeyValuePair(Of Integer, String)(出力形式.分割, "個別No.1,2,3"),
        New KeyValuePair(Of Integer, String)(出力形式.統合, "個別No.1～3")
    }
    ''' <summary>
    ''' 出力形式リスト（個別テーブル２つ）(REV_001)
    ''' </summary>
    Private Shared ReadOnly 出力形式リスト＿個別テーブル２ As New List(Of KeyValuePair(Of Integer, String)) From {
        New KeyValuePair(Of Integer, String)(出力形式.分割, "個別No.1,2"),
        New KeyValuePair(Of Integer, String)(出力形式.統合, "個別No.1～2")
    }
    ''' <summary>
    ''' 出力形式リスト（組織テーブル３つ）(REV_001)
    ''' </summary>
    Private Shared ReadOnly 出力形式リスト＿組織テーブル３ As New List(Of KeyValuePair(Of Integer, String)) From {
        New KeyValuePair(Of Integer, String)(出力形式.分割, "組織No.1,2,3"),
        New KeyValuePair(Of Integer, String)(出力形式.統合, "組織No.1～3")
    }
    ''' <summary>
    ''' 出力形式リスト（組織テーブル２つ）(REV_001)
    ''' </summary>
    Private Shared ReadOnly 出力形式リスト＿組織テーブル２ As New List(Of KeyValuePair(Of Integer, String)) From {
        New KeyValuePair(Of Integer, String)(出力形式.分割, "組織No.1,2"),
        New KeyValuePair(Of Integer, String)(出力形式.統合, "組織No.1～2")
    }
    ''' <summary>
    ''' 出力形式リスト取得(REV_001)
    ''' </summary>
    ''' <param name="version"></param>
    ''' <returns></returns>
    Private Shared Function Get出力形式リスト(version As String) As List(Of KeyValuePair(Of Integer, String))
        Dim list As List(Of KeyValuePair(Of Integer, String))
        Select Case CInt(CommonInfo.Chosakubun)
            Case CInt(ComConst.調査区分.営農類型別経営統計_個人)
                list = 出力形式リスト＿営農個人
            Case CInt(ComConst.調査区分.営農類型別経営統計_法人)
                list = 出力形式リスト＿営農法人
            Case CInt(ComConst.調査区分.米生産費統計_個別) To CInt(ComConst.調査区分.さとうきび生産費統計_個別)
                list = 出力形式リスト＿個別テーブル３
            Case CInt(ComConst.調査区分.米生産費統計_組織法人)
                list = 出力形式リスト＿組織テーブル３
            Case CInt(ComConst.調査区分.小麦生産費統計_組織法人) To CInt(ComConst.調査区分.大豆生産費統計_組織法人)
                list = 出力形式リスト＿組織テーブル２
            Case CInt(ComConst.調査区分.牛乳生産費統計_個別) To CInt(ComConst.調査区分.肥育豚生産費統計_個別)
                list = 出力形式リスト＿個別テーブル２
            Case Else
                Return New List(Of KeyValuePair(Of Integer, String)) From {New KeyValuePair(Of Integer, String)(出力形式.分割, "")}
        End Select
        Return list.Where(Function(x) x.Key = 出力形式.分割 OrElse version = ComConst.バージョン区分.結果表等項目2022).ToList()
    End Function

    ''' <summary>
    ''' 調査年選択(REV_001)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cboChosaNen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboChosaNen.SelectedIndexChanged
        Try
            Dim nen = If(cboChosaNen.SelectedValue?.ToString(), "0")
            If _chosaNen = nen Then
                Return
            End If
            _chosaNen = nen
            '一覧クリア
            dgvList.Rows.Clear()
            '出力形式プルダウン
            Dim version = ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun)
            CboOutType.ValueMember = "Key"
            CboOutType.DisplayMember = "Value"
            CboOutType.DataSource = Get出力形式リスト(version)
            CboOutType.SelectedIndex = 0
            '調査票プルダウン
            If version = ComConst.バージョン区分.結果表等項目2022 Then
                CboChosahyo.Enabled = True
            Else
                CboChosahyo.SelectedIndex = 0
                CboChosahyo.Enabled = False
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 統合個別結果表検索SQL取得(REV_001)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="outType"></param>
    ''' <param name="kessokuchiHokan"></param>
    ''' <returns></returns>
    Private Shared Function GetSelect統合個別結果表SQL(db As DBAccess, chosaNen As String, outType As 出力形式, kessokuchiHokan As String) As List(Of (tableName As String, sql As String))
        Dim sqlList = New List(Of (String, String))
        Dim colInfoLists = Get統合個別結果表列リスト(db, chosaNen, outType, kessokuchiHokan)
        Dim sbSql As StringBuilder
        Dim isFirst As Boolean
        Dim firstIndex = 0
        For Each colInfo In colInfoLists
            sbSql = New StringBuilder()
            sbSql.Append("SELECT")
            isFirst = True
            For Each colData In colInfo.colDataList
                If isFirst Then
                    isFirst = False
                Else
                    sbSql.Append(",")
                End If
                If outType = 出力形式.農経個人_4_6 AndAlso colData.colName = "実査設置拠点" Then
                    '実査設置拠点の後に固定値「経営種類」を追加
                    sbSql.Append(String.Format(" NO{0}.{1}", colData.index, colData.colName))
                    sbSql.Append(", 1 AS 経営種類")
                ElseIf outType = 出力形式.農経個人_4_6 AndAlso colData.colName = "K000024" Then
                    '空白(numericのため、null)を0に置換
                    sbSql.Append(String.Format(" ISNULL(NO{0}.{1}, 0) AS {1}", colData.index, colData.colName))
                ElseIf outType = 出力形式.農経法人_4_6 AndAlso colData.colName = "K000015" Then
                    'K000015の前に「集計対象区分」を追加
                    sbSql.Append(" IIF(ISNULL(K000015, 0) = 0, 0, 1) AS 集計対象区分,")
                    sbSql.Append(String.Format(" NO{0}.{1}", colData.index, colData.colName))
                ElseIf colData.colName = "K990001" Then
                    '通信欄の改行を半角スペースに置換
                    sbSql.Append(" REPLACE(K990001, CHAR(10), ' ') AS K990001")
                Else
                    sbSql.Append(String.Format(" NO{0}.{1}", colData.index, colData.colName))
                End If
            Next
            isFirst = True
            For Each index In colInfo.colDataList.Select(Function(x) x.index).Distinct().OrderBy(Function(x) x)
                If isFirst Then
                    sbSql.Append(String.Format(" FROM {0} NO{1}", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(index) & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan), index))
                    firstIndex = index
                    isFirst = False
                Else
                    sbSql.Append(String.Format(" JOIN {0} NO{1}", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(index) & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan), index))
                    sbSql.Append(String.Format(" ON NO{0}.調査年 = NO{1}.調査年", index, firstIndex))
                    sbSql.Append(String.Format(" AND NO{0}.センサス番号 = NO{1}.センサス番号", index, firstIndex))
                End If
            Next
            sbSql.Append(String.Format(" WHERE NO{0}.調査年 = @調査年", firstIndex))
            sbSql.Append(String.Format(" AND NO{0}.センサス番号 = @センサス番号", firstIndex))
            sqlList.Add((colInfo.tableName, sbSql.ToString()))
        Next
        Return sqlList
    End Function

    ''' <summary>
    ''' 統合個別結果表検索(REV_001)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="sqlInfoList"></param>
    ''' <param name="pkey"></param>
    ''' <returns></returns>
    Private Shared Function Select統合個別結果表(db As DBAccess, sqlInfoList As List(Of (tableName As String, sql As String)), pkey As DAOKobetsuKekkahyo.PrimaryKey) As Dictionary(Of String, DataTable)
        Dim dc = New Dictionary(Of String, DataTable)
        Dim para As List(Of DBAccess.Parameter)
        For Each sqlInfo In sqlInfoList
            para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pkey.chosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pkey.censusNo))
            dc.Add(sqlInfo.tableName, db.GetDataTable(sqlInfo.sql, para))
        Next
        Return dc
    End Function

    ''' <summary>
    ''' 統合個別結果表列検索SQL(REV_001)
    ''' </summary>
    Private Const SQL_SELECT_統合個別結果表列 =
        "SELECT" _
        & " {0} AS SORT_NO," _
        & " C.COLUMN_ID," _
        & " {1} AS IDX," _
        & " C.NAME AS COL_NAME" _
        & " FROM SYS.OBJECTS O" _
        & " JOIN SYS.COLUMNS C" _
        & " ON C.OBJECT_ID = O.OBJECT_ID" _
        & " WHERE O.NAME = '{2}'" _
        & " AND C.NAME {3} IN ({4})" _
        & " AND (C.NAME NOT LIKE 'K%'" _
        & "  OR EXISTS(SELECT 1 FROM 個別結果表項目マスタ M" _
        & "  WHERE M.バージョン区分 = {5}" _
        & "  AND M.調査区分 = {6}" _
        & "  AND M.項目番号 = C.NAME" _
        & "  AND M.裏項番区分 = 0))"
    ''' <summary>
    ''' 統合個別結果表列検索ORDER(REV_001)
    ''' </summary>
    Private Const SQL_SELECT_統合個別結果表列_ORDER =
        " ORDER BY" _
        & " CASE WHEN {0} = {1} AND COL_NAME IN ('K050747','K050847','K050947') THEN 99 ELSE SORT_NO END," _
        & " CASE WHEN COL_NAME LIKE 'K%' THEN COL_NAME ELSE FORMAT(COLUMN_ID,'0000000') END"
    ''' <summary>
    ''' キー情報列(REV_001)
    ''' </summary>
    Private Const キー情報列 = "'調査年','センサス番号','農政局','都道府県','実査設置拠点'"
    ''' <summary>
    ''' 農業経営体個人専用列(REV_001)
    ''' </summary>
    Private Const 農業経営体個人専用列 = ",'K000005','K000009','K000010','K000011','K000012','K000015','K000016','K000019','K000020','K000024'"
    ''' <summary>
    ''' 農業経営体法人専用列(REV_001)
    ''' </summary>
    Private Const 農業経営体法人専用列 = ",'K000005','K000006','K000013','K000014','K000015','K000017','K000018','K000019','K000020','K000024'"
    ''' <summary>
    ''' 更新情報列(REV_001)
    ''' </summary>
    Private Const 更新情報列 = "'更新日付','更新者ID'"
    ''' <summary>
    ''' 統合個別結果表列リスト取得(REV_001)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="outType"></param>
    ''' <param name="kessokuchiHokan"></param>
    ''' <returns></returns>
    Private Shared Function Get統合個別結果表列リスト(db As DBAccess, chosaNen As String, outType As 出力形式, kessokuchiHokan As String) As List(Of (tableName As String, colDataList As List(Of (index As Integer, colName As String))))
        Dim colLists = New List(Of (String, List(Of (Integer, String))))
        Dim indexes = GetIndexes(outType)
        '既存と同様にキーとなるtableNameには"＿集計用"を付けない
        Dim tableName = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(indexes.first)
        If indexes.first <> indexes.last Then
            tableName &= "＿" & Strings.Right(ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(indexes.last), (indexes.last + 1).ToString().Length)
        End If
        '農業経営体は専用の名称とする
        Select Case outType
            Case 出力形式.農経個人_4_6
                tableName = "個別結果表＿農業経営＿営農類型＿農業経営体＿個人"
            Case 出力形式.農経法人_4_6
                tableName = "個別結果表＿農業経営＿営農類型＿農業経営体＿法人"
        End Select
        Dim version = ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun)
        Dim sbSql = New StringBuilder()
        sbSql.Append("SELECT * FROM (")
        For i = indexes.first To indexes.last
            If i = indexes.first Then
                'キー情報、農業経営体専用列は先頭のテーブルから出力
                Dim includes = キー情報列
                If outType = 出力形式.農経個人_4_6 Then
                    includes &= 農業経営体個人専用列
                ElseIf outType = 出力形式.農経法人_4_6 Then
                    includes &= 農業経営体法人専用列
                End If
                sbSql.Append(String.Format(SQL_SELECT_統合個別結果表列, -1, 0, ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0) & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan), "", includes, version, CommonInfo.Chosakubun))
            End If
            'キー情報、更新情報を除外
            sbSql.Append(" UNION ALL ")
            sbSql.Append(String.Format(SQL_SELECT_統合個別結果表列, i, i, ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(i) & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan), "NOT", キー情報列 & "," & 更新情報列, version, CommonInfo.Chosakubun))
        Next
        sbSql.Append(" UNION ALL ")
        '更新情報は先頭のテーブルから出力
        sbSql.Append(String.Format(SQL_SELECT_統合個別結果表列, indexes.last + 1, 0, ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0) & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, kessokuchiHokan), "", 更新情報列, version, CommonInfo.Chosakubun))
        sbSql.Append(") WRK")
        sbSql.Append(String.Format(SQL_SELECT_統合個別結果表列_ORDER, outType.ToString("d"), 出力形式.農経個人_4_6.ToString("d")))
        colLists.Add((tableName, db.GetDataTable(sbSql.ToString()).Rows.Cast(Of DataRow).Select(Function(x) (CInt(x("IDX")), x("COL_NAME").ToString())).ToList()))

        '法人７は分けて出力
        Select Case outType
            Case 出力形式.法人_1_3_7, 出力形式.法人_1_6_7
                tableName = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(6)
                sbSql = New StringBuilder()
                sbSql.Append("SELECT * FROM (")
                sbSql.Append(String.Format(SQL_SELECT_統合個別結果表列, 0, 6, tableName, "", キー情報列, version, CommonInfo.Chosakubun))
                sbSql.Append(" UNION ALL ")
                sbSql.Append(String.Format(SQL_SELECT_統合個別結果表列, 1, 6, tableName, "NOT", キー情報列 & "," & 更新情報列, version, CommonInfo.Chosakubun))
                sbSql.Append(" UNION ALL ")
                sbSql.Append(String.Format(SQL_SELECT_統合個別結果表列, 2, 6, tableName, "", 更新情報列, version, CommonInfo.Chosakubun))
                sbSql.Append(") WRK")
                sbSql.Append(String.Format(SQL_SELECT_統合個別結果表列_ORDER, outType.ToString("d"), 出力形式.農経個人_4_6.ToString("d")))
                colLists.Add((tableName, db.GetDataTable(sbSql.ToString()).Rows.Cast(Of DataRow).Select(Function(x) (CInt(x("IDX")), x("COL_NAME").ToString())).ToList()))
        End Select

        Return colLists
    End Function

    ''' <summary>
    ''' 統合対象テーブル名称のIndexを取得(REV_001)
    ''' </summary>
    ''' <param name="outType"></param>
    ''' <returns></returns>
    Private Shared Function GetIndexes(outType As 出力形式) As (first As Integer, last As Integer)
        Dim first = 0
        Dim last = 0
        Select Case outType
            Case 出力形式.統合
                last = UBound(ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun))
            Case 出力形式.個人_1_3, 出力形式.法人_1_3_7
                last = 2
            Case 出力形式.法人_1_6_7
                last = 5
            Case 出力形式.農経個人_4_6, 出力形式.農経法人_4_6
                first = 3
                last = 5
        End Select
        Return (first, last)
    End Function

    ''' <summary>
    ''' 個別結果表項目名検索SQL(REV_001)
    ''' </summary>
    Private Const SQL_SELECT_ITEM_NAMES =
        "SELECT DISTINCT" _
        & " 項目番号," _
        & " REPLACE(項目名, CHAR(10), ' ') AS 項目名" _
        & " FROM 個別結果表作成論理{0}" _
        & " WHERE バージョン区分 = @バージョン区分" _
        & " AND 調査区分 = @調査区分" _
        & " ORDER BY 1"
    ''' <summary>
    ''' 個別結果表項目名取得(REV_001)
    ''' </summary>
    ''' <returns></returns>
    Private Shared Function GetItemNames(chosaNen As String) As List(Of (項目番号 As String, 項目名 As String))
        Dim einoKojin = ""
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            einoKojin = "＿営農個人"
        End If
        Dim version = ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun)
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Dim para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, version))
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            Dim dt = db.GetDataTable(String.Format(SQL_SELECT_ITEM_NAMES, einoKojin), para)
            Return dt.Rows.Cast(Of DataRow).Select(Function(x) (x("項目番号").ToString(), x("項目名").ToString())).ToList()
        End Using
    End Function

End Class
