Imports System.IO
Imports System.Text

''' <summary>
''' 牛トレサデータ管理画面
''' </summary>
''' <remarks></remarks>
Public Class BRA2010F

#Region "変数一覧"

    ''' <summary>調査年</summary>
    Private _yearFrom As String
    ''' <summary>局</summary>
    Private _monthFrom As String
    ''' <summary>事務所</summary>
    Private _yearTo As String
    ''' <summary>拠点</summary>
    Private _monthTo As String
    ''' <summary>営農類型</summary>
    Private _keika As String

    Private Const YEAR_FROM As String = "0"
    Private Const YEAR_TO As String = "1"
    Private Const MONTH_FROM As String = "0"
    Private Const MONTH_TO As String = "1"
    Private Const CHK_SCREEN As String = "0"
    Private Const CHK_LIST As String = "1"
    Private Const LIST_CHK_BOX As String = "0"
    Private Const LIST_CHK_KOME As String = "1"

    ''' <summary>進捗ダイアログ</summary>
    Private ProgressDialog As ProgressDialog

#End Region

#Region "【処理詳細仕様 1】初期表示"

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA2010F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'DataGridView設定
            ComUtil.ConfigDgv(Me.dgvList)
            dgvList.Columns(2).ReadOnly = False
            dgvList.Columns(3).ReadOnly = False
            dgvList.Columns(4).ReadOnly = False
            dgvList.Columns(5).ReadOnly = False
            dgvList.Columns(6).ReadOnly = False

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

#End Region

#Region "【処理詳細仕様 2】「表示」ボタンクリック"

    ''' <summary>
    ''' 表示ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        Try

            '画面入力項目単体チェック
            '開始年のチェック
            If YearCheck(txtYearFrom.Text, YEAR_FROM, CHK_SCREEN) <> 0 Then
                txtYearFrom.BackColor = Color.Red
                Exit Sub
            Else
                txtYearFrom.BackColor = Color.White
            End If

            '開始月のチェック
            If MonthCheck(txtMonthFrom.Text, MONTH_FROM, CHK_SCREEN) <> 0 Then
                txtMonthFrom.BackColor = Color.Red
                Exit Sub
            Else
                txtMonthFrom.BackColor = Color.White
            End If

            '終了年のチェック
            If YearCheck(txtYearTo.Text, YEAR_TO, CHK_SCREEN) <> 0 Then
                txtYearTo.BackColor = Color.Red
                Exit Sub
            Else
                txtYearTo.BackColor = Color.White
            End If

            '終了月のチェック
            If MonthCheck(txtMonthTo.Text, MONTH_TO, CHK_SCREEN) <> 0 Then
                txtMonthTo.BackColor = Color.Red
                Exit Sub
            Else
                txtMonthTo.BackColor = Color.White
            End If

            '画面入力項目間チェック
            '開始年月と終了年月の大小チェック
            Dim ret As Integer = DateConsistentCheck(txtYearFrom.Text, txtMonthFrom.Text, txtYearTo.Text, txtMonthTo.Text)

            Select Case ret
                Case -1
                    txtYearFrom.BackColor = Color.Red
                    txtMonthFrom.BackColor = Color.Red
                    Exit Sub
                Case -2
                    txtYearTo.BackColor = Color.Red
                    txtMonthTo.BackColor = Color.Red
                    Exit Sub
                Case -3
                    txtYearFrom.BackColor = Color.Red
                    txtMonthFrom.BackColor = Color.Red
                    txtYearTo.BackColor = Color.Red
                    txtMonthTo.BackColor = Color.Red
                    Exit Sub
                Case Else

            End Select

            _yearFrom = txtYearFrom.Text
            _monthFrom = txtMonthFrom.Text.PadLeft(2, "0"c)
            _yearTo = txtYearTo.Text
            _monthTo = txtMonthTo.Text.PadLeft(2, "0"c)
            If cboKeika.SelectedIndex = -1 Or cboKeika.SelectedIndex = 0 Then
                _keika = Nothing
            ElseIf cboKeika.SelectedIndex = 1 Then
                _keika = "2年以上経過"
            ElseIf cboKeika.SelectedIndex = 2 Then
                _keika = "－"
            End If

            '一覧表示
            Me.ShowList(_yearFrom, _monthFrom, _yearTo, _monthTo, _keika)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

#End Region

#Region "【処理詳細仕様 3】データ期間変更、【処理詳細仕様 4】取込データ名称変更、"
    ''' <summary>
    ''' データグリッドビュー値変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvList_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvList.CellEndEdit
        If e.ColumnIndex = 2 Or e.ColumnIndex = 3 Or e.ColumnIndex = 4 Or e.ColumnIndex = 5 Or e.ColumnIndex = 6 Then
            If dgvList(e.ColumnIndex, e.RowIndex).Value Is Nothing Then
                dgvList(0, e.RowIndex).Value = "*"
                btnDelete.Enabled = False
            Else
                If dgvList(e.ColumnIndex, e.RowIndex).Tag.ToString <> dgvList(e.ColumnIndex, e.RowIndex).Value.ToString Then
                    dgvList(0, e.RowIndex).Value = "*"
                    btnDelete.Enabled = False
                End If

                If e.ColumnIndex = 3 Or e.ColumnIndex = 5 Then
                    dgvList(e.ColumnIndex, e.RowIndex).Value = dgvList(e.ColumnIndex, e.RowIndex).Value.ToString.PadLeft(2, "0"c)
                End If
            End If
        End If

    End Sub
#End Region

#Region "【処理詳細仕様 5】「全選択」ボタンクリック"
    ''' <summary>
    ''' 全選択ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAllSelect_Click(sender As Object, e As EventArgs) Handles btnAllSelect.Click
        Try
            For i As Integer = 0 To dgvList.Rows.Count - 1
                If dgvList(1, i).ReadOnly = False Then
                    dgvList(1, i).Value = True
                End If
            Next
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region

#Region "【処理詳細仕様 6】「全解除」ボタンクリック"
    ''' <summary>
    ''' 全解除ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAllCancel_Click(sender As Object, e As EventArgs) Handles btnAllCancel.Click
        Try
            For i As Integer = 0 To dgvList.Rows.Count - 1
                If dgvList(1, i).ReadOnly = False Then
                    dgvList(1, i).Value = False
                End If
            Next
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region

#Region "【処理詳細仕様 7】「出力」ボタンクリック"

    ''' <summary>
    ''' 出力ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnMake_Click(sender As Object, e As EventArgs) Handles btnOutput.Click

        Dim msgRet As MsgBoxResult
        Dim dt As DataTable

        Try
            'チェックボックス選択有無チェック
            If Not CheckedInf(LIST_CHK_BOX) Then
                Message.ShowMsgBox(MessageID.MSG_E_095, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '出力処理開始確認メッセージ表示
            msgRet = Message.ShowMsgBox(MessageID.MSG_Q_039, MsgBoxStyle.YesNo)
            If msgRet = MsgBoxResult.No Then
                Exit Sub
            End If

            'フォルダパス取得
            Dim folderPath As String = ComUtil.GetFolderPath(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath))

            If folderPath.Equals(String.Empty) Then
                Exit Sub
            End If

            'ファイル出力
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                For i As Integer = 0 To dgvList.Rows.Count - 1
                    If dgvList(1, i).Value IsNot Nothing _
                        And CBool(dgvList(1, i).Value) = True Then

                        Dim txtFrom As String = dgvList(2, i).Value.ToString & dgvList(3, i).Value.ToString
                        Dim txtTo As String = dgvList(4, i).Value.ToString & dgvList(5, i).Value.ToString

                        'ファイル名作成(個体情報)
                        Dim fileName As String = Me.GetFileName("個体情報", txtFrom, txtTo, dgvList(6, i).Value.ToString)

                        '牛トレサ個体情報取込履歴取得
                        dt = DAOOther.GetTresaKotaiHistory(db, CInt(dgvList(11, i).Value))

                        'CSVファイル出力
                        Dim ret As ComConst.CSVファイル.enmOutputReturn = Me.PutCSV(folderPath, fileName, dt)

                        If ret = ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS Then
                            'エラーメッセージ表示
                            Message.ShowMsgBox(MessageID.MSG_E_098, MsgBoxStyle.OkOnly)
                            Exit Sub
                        End If

                        'ファイル名作成(異動情報)
                        fileName = Me.GetFileName("異動情報", txtFrom, txtTo, dgvList(6, i).Value.ToString)

                        '牛トレサ異動情報取込履歴取得
                        dt = DAOOther.GetTresaIdoHistory(db, CInt(dgvList(11, i).Value))

                        'CSVファイル出力
                        ret = Me.PutCSV(folderPath, fileName, dt)
                        If ret = ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS Then
                            'エラーメッセージ表示
                            Message.ShowMsgBox(MessageID.MSG_E_098, MsgBoxStyle.OkOnly)
                            Exit Sub
                        End If

                    End If
                Next
            End Using

            '完了メッセージ表示
            Message.ShowMsgBox(MessageID.MSG_I_038, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' ファイル名作成処理
    ''' </summary>
    ''' <param name="fileType"></param>
    ''' <param name="txtFrom"></param>
    ''' <param name="txtTo"></param>
    ''' <param name="txtImportName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetFileName(fileType As String, txtFrom As String, txtTo As String, txtImportName As String) As String
        Return fileType & "_" & txtFrom & "_" & txtTo & "_" & txtImportName & ".csv"
    End Function

    ''' <summary>
    ''' CSV出力処理
    ''' </summary>
    ''' <param name="outDir"></param>
    ''' <param name="fileName"></param>
    ''' <param name="dt"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PutCSV(outDir As String, fileName As String, dt As DataTable) As ComConst.CSVファイル.enmOutputReturn

        Dim ret As ComConst.CSVファイル.enmOutputReturn = ComConst.CSVファイル.enmOutputReturn.CANCEL

        Dim sjisEnc As Encoding = Encoding.GetEncoding(ComConst.CSVファイル.CODEPAGE_SHIFT_JIS)

        Dim filePathTemp As String = System.IO.Path.Combine(outDir, fileName)

        Dim strWriteText As String

        If Not System.IO.Directory.Exists(outDir) Then
            Directory.CreateDirectory(outDir)
        End If

        Try
            Using sw As New System.IO.StreamWriter(filePathTemp, False, sjisEnc)
                For Each row As DataRow In dt.Rows
                    strWriteText = ""
                    For Each item As Object In row.ItemArray
                        strWriteText = strWriteText & item.ToString & ","
                    Next
                    strWriteText = strWriteText.TrimEnd(CType(",", Char))
                    sw.WriteLine(strWriteText)
                Next
            End Using
        Catch ex As Exception
            ret = ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS
            Return ret
        End Try
        ret = ComConst.CSVファイル.enmOutputReturn.OK

        Return ret
    End Function

#End Region

#Region "【処理詳細仕様 8】「修正」ボタンクリック"
    ''' <summary>
    ''' 修正ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Dim msgRet As MsgBoxResult
        Dim dt As DataTable
        Dim ret As Boolean

        Try
            '「*」表示有無チェック
            If Not CheckedInf(LIST_CHK_KOME) Then
                Message.ShowMsgBox(MessageID.MSG_E_096, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '更新処理開始確認メッセージ表示
            msgRet = Message.ShowMsgBox(MessageID.MSG_Q_040, MsgBoxStyle.YesNo)
            If msgRet = MsgBoxResult.No Then
                Exit Sub
            End If

            'リストで、「*」の表示されている行分繰り返す
            For i As Integer = 0 To dgvList.Rows.Count - 1
                If dgvList(0, i).Value IsNot Nothing _
                    And CStr(dgvList(0, i).Value) = "*" Then

                    '入力内容チェック(開始年)
                    If YearCheck(CStr(dgvList(2, i).Value), YEAR_FROM, CHK_LIST) <> 0 Then
                        dgvList(2, i).Style.BackColor = Color.Red
                        dgvList.CurrentCell = dgvList(0, i)
                        Exit Sub
                    Else
                        dgvList(2, i).Style.BackColor = Color.White
                    End If

                    '入力内容チェック(開始月)
                    If MonthCheck(CStr(dgvList(3, i).Value), MONTH_FROM, CHK_LIST) <> 0 Then
                        dgvList(3, i).Style.BackColor = Color.Red
                        dgvList.CurrentCell = dgvList(0, i)
                        Exit Sub
                    Else
                        dgvList(3, i).Style.BackColor = Color.White
                    End If

                    '入力内容チェック(終了年)
                    If YearCheck(CStr(dgvList(4, i).Value), YEAR_TO, CHK_LIST) <> 0 Then
                        dgvList(4, i).Style.BackColor = Color.Red
                        dgvList.CurrentCell = dgvList(0, i)
                        Exit Sub
                    Else
                        dgvList(4, i).Style.BackColor = Color.White
                    End If

                    '入力内容チェック(終了月)
                    If MonthCheck(CStr(dgvList(5, i).Value), MONTH_TO, CHK_LIST) <> 0 Then
                        dgvList(5, i).Style.BackColor = Color.Red
                        dgvList.CurrentCell = dgvList(0, i)
                        Exit Sub
                    Else
                        dgvList(5, i).Style.BackColor = Color.White
                    End If

                    '入力内容チェック(取込データ名称)
                    If ImportDataNameCheck(CStr(dgvList(6, i).Value)) <> 0 Then
                        dgvList(6, i).Style.BackColor = Color.Red
                        dgvList.CurrentCell = dgvList(0, i)
                        Exit Sub
                    Else
                        dgvList(6, i).Style.BackColor = Color.White
                    End If

                    '入力項目間チェック
                    '開始年月と終了年月の大小チェック
                    Dim iRet As Integer = DateConsistentCheck(CStr(dgvList(2, i).Value), CStr(dgvList(3, i).Value), CStr(dgvList(4, i).Value), CStr(dgvList(5, i).Value))

                    Select Case iRet
                        Case -1
                            dgvList(2, i).Style.BackColor = Color.Red
                            dgvList(3, i).Style.BackColor = Color.Red
                            dgvList.CurrentCell = dgvList(0, i)
                            Exit Sub
                        Case -2
                            dgvList(4, i).Style.BackColor = Color.Red
                            dgvList(5, i).Style.BackColor = Color.Red
                            dgvList.CurrentCell = dgvList(0, i)
                            Exit Sub
                        Case -3
                            dgvList(2, i).Style.BackColor = Color.Red
                            dgvList(3, i).Style.BackColor = Color.Red
                            dgvList(4, i).Style.BackColor = Color.Red
                            dgvList(5, i).Style.BackColor = Color.Red
                            dgvList.CurrentCell = dgvList(0, i)
                            Exit Sub
                        Case Else

                    End Select
                End If
            Next

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try
                    'トランザクション開始
                    db.BeginTrans()

                    For i As Integer = 0 To dgvList.Rows.Count - 1
                        If dgvList(0, i).Value IsNot Nothing _
                            And CStr(dgvList(0, i).Value) = "*" Then
                            Dim txtFrom As String = dgvList(2, i).Value.ToString & dgvList(3, i).Value.ToString
                            Dim txtTo As String = dgvList(4, i).Value.ToString & dgvList(5, i).Value.ToString

                            '同一期間及び同一取込データ名称データの有無をチェック
                            dt = DAOOther.GetSamePeriodTresaData(db, txtFrom, txtTo, dgvList(6, i).Value.ToString)

                            If dt.Rows.Count <> 0 Then
                                Message.ShowMsgBox(MessageID.MSG_E_086, MsgBoxStyle.OkOnly)
                                Exit Sub
                            End If

                            '更新処理実施
                            ret = DAOOther.UpdateTresaData(db, txtFrom, txtTo, dgvList(6, i).Value.ToString, CInt(dgvList(11, i).Value))

                        End If
                    Next

                    'コミット
                    db.CommitTrans()

                    'リスト再表示
                    _yearFrom = txtYearFrom.Text
                    _monthFrom = txtMonthFrom.Text.PadLeft(2, "0"c)
                    _yearTo = txtYearTo.Text
                    _monthTo = txtMonthTo.Text.PadLeft(2, "0"c)
                    If cboKeika.SelectedIndex = -1 Or cboKeika.SelectedIndex = 0 Then
                        _keika = Nothing
                    ElseIf cboKeika.SelectedIndex = 1 Then
                        _keika = "2年以上経過"
                    ElseIf cboKeika.SelectedIndex = 2 Then
                        _keika = "－"
                    End If

                    Me.ShowList(_yearFrom, _monthFrom, _yearTo, _monthTo, _keika)

                    '完了メッセージ表示
                    Message.ShowMsgBox(MessageID.MSG_I_039, MsgBoxStyle.OkOnly)

                Catch ex As Exception
                    'エラーの場合は、ロールバック
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using

        Catch ex As Exception
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region
#Region "【処理詳細仕様 9】「削除」ボタンクリック"
    ''' <summary>
    ''' 削除ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim msgRet As MsgBoxResult
        Dim dt As DataTable

        Try
            'チェックボックス選択有無チェック
            If Not CheckedInf(LIST_CHK_BOX) Then
                Message.ShowMsgBox(MessageID.MSG_E_097, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '削除処理開始確認メッセージ表示
            msgRet = Message.ShowMsgBox(MessageID.MSG_Q_041, MsgBoxStyle.YesNo)
            If msgRet = MsgBoxResult.No Then
                Exit Sub
            End If

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try
                    'トランザクション開始
                    db.BeginTrans()
                    For i As Integer = 0 To dgvList.Rows.Count - 1
                        If dgvList(1, i).Value IsNot Nothing _
                            And CBool(dgvList(1, i).Value) = True Then

                            '牛トレサ取込履歴テーブルのデータ削除
                            DAOOther.DeleteTresaData(db, CInt(dgvList(11, i).Value), "牛トレサ取込履歴")
                            '牛トレサ個体情報取込履歴テーブルのデータ削除
                            DAOOther.DeleteTresaData(db, CInt(dgvList(11, i).Value), "牛トレサ個体情報取込履歴")
                            '牛トレサ異動情報取込履歴テーブルのデータ削除
                            DAOOther.DeleteTresaData(db, CInt(dgvList(11, i).Value), "牛トレサ異動情報取込履歴")

                        End If
                    Next

                    '牛トレサ個体情報テーブルのデータを削除
                    DAOOther.DeleteTresaKotaiInfo(db)

                    '牛トレサ異動情報テーブルのデータを削除
                    DAOOther.DeleteTresaIdoInfo(db)

                    'データがすべて削除された場合は、更新処理は実施しない
                    dt = DAOOther.GetTresaHistory(db)
                    If dt.Rows.Count <> 0 Then
                        '牛トレサ個体情報の更新
                        DAOOther.InsertTresaKotaiInfo(db)

                        '牛トレサ異動情報の更新
                        DAOOther.InsertTresaIdoInfo(db)
                    End If

                    'コミット
                    db.CommitTrans()
                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using

            'リスト再表示
            _yearFrom = txtYearFrom.Text
            _monthFrom = txtMonthFrom.Text.PadLeft(2, "0"c)
            _yearTo = txtYearTo.Text
            _monthTo = txtMonthTo.Text.PadLeft(2, "0"c)
            If cboKeika.SelectedIndex = -1 Or cboKeika.SelectedIndex = 0 Then
                _keika = Nothing
            ElseIf cboKeika.SelectedIndex = 1 Then
                _keika = "2年以上経過"
            ElseIf cboKeika.SelectedIndex = 2 Then
                _keika = "－"
            End If

            Me.ShowList(_yearFrom, _monthFrom, _yearTo, _monthTo, _keika)

            '完了メッセージ表示
            Message.ShowMsgBox(MessageID.MSG_I_040, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region

#Region "処理全般"


    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="einouRuikei"></param>
    ''' <remarks></remarks>
    Private Sub ShowList(strYearFrom As String, strMonthFrom As String, strYearTo As String, strMonthTo As String, strKeika As String)
        Dim dt As DataTable

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dt = DAOOther.GetTresaData(db, strYearFrom & strMonthFrom, strYearTo & strMonthTo, strKeika)
        End Using

        dgvList.Rows.Clear()

        For Each row As DataRow In dt.Rows
            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(2).Value = row("期間開始年").ToString
            dgvList.Rows(i).Cells(2).Tag = row("期間開始年").ToString
            dgvList.Rows(i).Cells(3).Value = row("期間開始月").ToString
            dgvList.Rows(i).Cells(3).Tag = row("期間開始月").ToString
            dgvList.Rows(i).Cells(4).Value = row("期間終了年").ToString
            dgvList.Rows(i).Cells(4).Tag = row("期間終了年").ToString
            dgvList.Rows(i).Cells(5).Value = row("期間終了月").ToString
            dgvList.Rows(i).Cells(5).Tag = row("期間終了月").ToString
            dgvList.Rows(i).Cells(6).Value = row("取込データ名称").ToString
            dgvList.Rows(i).Cells(6).Tag = row("取込データ名称").ToString
            dgvList.Rows(i).Cells(7).Value = row("データ登録日").ToString
            dgvList.Rows(i).Cells(8).Value = row("データ登録以降2年経過情報").ToString
            dgvList.Rows(i).Cells(9).Value = row("個体情報レコード数").ToString
            dgvList.Rows(i).Cells(10).Value = row("異動情報レコード数").ToString
            dgvList.Rows(i).Cells(11).Value = row("履歴番号").ToString
        Next

        btnDelete.Enabled = True
    End Sub

    ''' <summary>
    ''' 開始年/終了年 個別チェック
    ''' </summary>
    ''' <param name="txtYear">チェックする年の値</param>
    ''' <param name="iType">0:開始年、1:終了年</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function YearCheck(strYear As String, strType As String, strChkType As String) As Integer
        Dim iSystemYear As Integer
        Dim iYear As Integer
        Dim iRet As Integer = -1

        Try
            If strYear = "" Then
                If strChkType = CHK_SCREEN Then
                    Return 0
                ElseIf strChkType = CHK_LIST Then
                    If strType = YEAR_FROM Then
                        Message.ShowMsgBox(MessageID.MSG_E_069, MsgBoxStyle.OkOnly)
                    ElseIf strType = YEAR_TO Then
                        Message.ShowMsgBox(MessageID.MSG_E_073, MsgBoxStyle.OkOnly)
                    End If
                    Return -1
                End If
            End If

            'システム年を数値に変換
            iSystemYear = CInt(DateTime.Now.ToString("yyyy"))

            If Integer.TryParse(strYear, iYear) Then
                '入力年を数値に変換
                iYear = CInt(strYear)

                '入力年がシステム年-3～システム年+1の範囲かをチェック
                If iYear >= iSystemYear - 3 And iYear <= iSystemYear + 1 Then
                    iRet = 0
                Else
                    If strType = YEAR_FROM Then
                        Message.ShowMsgBox(MessageID.MSG_E_070, {(iSystemYear - 3).ToString, (iSystemYear + 1).ToString}, MsgBoxStyle.OkOnly)
                    ElseIf strType = YEAR_TO Then
                        Message.ShowMsgBox(MessageID.MSG_E_074, {(iSystemYear - 3).ToString, (iSystemYear + 1).ToString}, MsgBoxStyle.OkOnly)
                    End If
                    iRet = -1
                End If
            Else
                '入力内容が数値でない場合もエラー
                If strType = YEAR_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_070, {(iSystemYear - 3).ToString, (iSystemYear + 1).ToString}, MsgBoxStyle.OkOnly)
                ElseIf strType = YEAR_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_074, {(iSystemYear - 3).ToString, (iSystemYear + 1).ToString}, MsgBoxStyle.OkOnly)
                End If
                iRet = -1
            End If

        Catch ex As Exception
            Throw ex
        End Try
        Return iRet
    End Function

    ''' <summary>
    ''' 開始月/終了月 個別チェック
    ''' </summary>
    ''' <param name="txtMonth">チェックする月の値</param>
    ''' <param name="iType">0:開始月、1:終了月</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function MonthCheck(strMonth As String, strType As String, strChkType As String) As Integer
        Dim iMonth As Integer
        Dim iRet As Integer = -1

        Try
            If strMonth = "" Then
                If strChkType = CHK_SCREEN Then
                    Return 0
                ElseIf strChkType = CHK_LIST Then
                    If strType = MONTH_FROM Then
                        Message.ShowMsgBox(MessageID.MSG_E_071, MsgBoxStyle.OkOnly)
                    ElseIf strType = MONTH_TO Then
                        Message.ShowMsgBox(MessageID.MSG_E_075, MsgBoxStyle.OkOnly)
                    End If
                    Return -1
                End If
            End If

            If strChkType = CHK_SCREEN Then
                If strMonth.Length = 2 AndAlso strMonth.Substring(0, 1) = "0" Then
                    If strType = MONTH_FROM Then
                        Message.ShowMsgBox(MessageID.MSG_E_072, MsgBoxStyle.OkOnly)
                    ElseIf strType = MONTH_TO Then
                        Message.ShowMsgBox(MessageID.MSG_E_076, MsgBoxStyle.OkOnly)
                    End If
                    Return -1
                End If
            End If

            If Integer.TryParse(strMonth, iMonth) Then
                If iMonth >= 1 And iMonth <= 12 Then
                    iRet = 0
                Else
                    If strType = MONTH_FROM Then
                        Message.ShowMsgBox(MessageID.MSG_E_072, MsgBoxStyle.OkOnly)
                    ElseIf strType = MONTH_TO Then
                        Message.ShowMsgBox(MessageID.MSG_E_076, MsgBoxStyle.OkOnly)
                    End If
                    iRet = -1
                End If
            Else
                '入力内容が数値でない場合もエラー
                If strType = MONTH_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_072, MsgBoxStyle.OkOnly)
                ElseIf strType = MONTH_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_076, MsgBoxStyle.OkOnly)
                End If
                iRet = -1
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return iRet
    End Function

    ''' <summary>
    ''' 開始/終了年月間整合性チェック
    ''' </summary>
    ''' <param name="txtYearFrom">開始年の値</param>
    ''' <param name="txtMonthFrom">開始月の値</param>
    ''' <param name="txtYearTo">終了年の値</param>
    ''' <param name="txtMonthTo">終了月の値</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DateConsistentCheck(strYearFrom As String, strMonthFrom As String, strYearTo As String, strMonthTo As String) As Integer
        Dim iRet As Integer = -1
        Dim txtFrom As String
        Dim txtTo As String

        Try
            If Not ((strYearFrom = "" And strMonthFrom = "") Or (strYearFrom <> "" And strMonthFrom <> "")) Then
                Message.ShowMsgBox(MessageID.MSG_E_084, MsgBoxStyle.OkOnly)
                Return -1
            End If

            If Not ((strYearTo = "" And strMonthTo = "") Or (strYearTo <> "" And strMonthTo <> "")) Then
                Message.ShowMsgBox(MessageID.MSG_E_085, MsgBoxStyle.OkOnly)
                Return -2
            End If

            txtFrom = strYearFrom & strMonthFrom.PadLeft(2, "0"c)
            txtTo = strYearTo & strMonthTo.PadLeft(2, "0"c)

            If (txtFrom = "00" Or txtTo = "00") Then
                iRet = 0
            Else
                If txtFrom > txtTo Then
                    Message.ShowMsgBox(MessageID.MSG_E_083, MsgBoxStyle.OkOnly)
                    iRet = -3
                Else
                    iRet = 0
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return iRet
    End Function

    Private Function ImportDataNameCheck(strImportDataName As String) As Integer
        Dim iRet As Integer = -1

        Try
            If strImportDataName = "" Then
                Message.ShowMsgBox(MessageID.MSG_E_077, MsgBoxStyle.OkOnly)
                Return -1
            End If

            If System.Text.Encoding.GetEncoding(932).GetByteCount(strImportDataName) <= 100 Then
                iRet = 0
            Else
                Message.ShowMsgBox(MessageID.MSG_E_078, MsgBoxStyle.OkOnly)
                iRet = -1
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return iRet
    End Function


    ''' <summary>
    ''' 一覧選択確認
    ''' </summary>
    ''' <param name="iType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckedInf(strType As String) As Boolean
        Dim checked As Boolean = False

        If strType = LIST_CHK_BOX Then
            For i As Integer = 0 To dgvList.Rows.Count - 1
                If dgvList(1, i).Value IsNot Nothing _
                    And CBool(dgvList(1, i).Value) = True Then
                    'チェックされていた場合
                    checked = True
                    Exit For
                End If
            Next
        ElseIf strType = LIST_CHK_KOME Then
            For i As Integer = 0 To dgvList.Rows.Count - 1
                If dgvList(0, i).Value IsNot Nothing _
                    And CStr(dgvList(0, i).Value) = "*" Then
                    'チェックされていた場合
                    checked = True
                    Exit For
                End If
            Next
        End If
        Return checked
    End Function
#End Region
End Class
