Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO

''' <summary>
''' DB管理画面
''' </summary>
''' <remarks></remarks>
Public Class BRA9110F

    ''' <summary>開始付加文字</summary>
    Private Const START_ADDITION As String = """"
    ''' <summary>終了付加文字</summary>
    Private Const END_ADDITION As String = START_ADDITION
    ''' <summary>CSV区切文字</summary>
    Private Const CSV_DELIMITER As String = ","
    ''' <summary>コードページ_Shift_JIS</summary>
    Private Const CODEPAGE_SHIFT_JIS As String = "Shift_JIS"

    ''' <summary>進捗ダイアログ</summary>
    Private ProgressDialog As ProgressDialog

    Private Sub BRA9110F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '区分１コンボボックス設定
            ComUtil.SetKubun1ComboBox(cboKubun1)

            Select Case CommonInfo.Koutei
                Case CommonInfo.KouteiKubun.Code.Honsyo
                    chkZenchosakubun.Visible = False
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    chkZenchosakubun.Visible = False
                Case CommonInfo.KouteiKubun.Code.Center
                    chkZenchosakubun.Checked = True
            End Select
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub chkZenchosakubun_CheckStateChanged(sender As Object, e As EventArgs) Handles chkZenchosakubun.CheckStateChanged
        Select Case DirectCast(sender, CheckBox).CheckState
            Case CheckState.Checked
                'チェック
                cboKubun1.Enabled = False
                cboKubun2.Enabled = False
                cboChosakubun.Enabled = False
            Case CheckState.Unchecked
                '未チェック
                cboKubun1.Enabled = True
                cboKubun2.Enabled = True
                cboChosakubun.Enabled = True
        End Select
    End Sub

    Private Sub cboKubun1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKubun1.SelectedIndexChanged
        Try
            '区分２コンボボックス設定
            ComUtil.SetKubun2ComboBox(cboKubun1, cboKubun2)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub cboKubun2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKubun2.SelectedIndexChanged
        Try
            '調査区分コンボボックス設定
            ComUtil.SetChosakubunComboBox(cboKubun1, cboKubun2, cboChosakubun)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnBackUp_Click(sender As Object, e As EventArgs) Handles btnBackUp.Click
        Dim bln As Boolean = False

        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Dim year As String = txtYear.Text

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(year, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_018, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                Dim chosaKubunArr As String()
                Dim tblCnt As Integer

                If chkZenchosakubun.Checked Then
                    chosaKubunArr = New String(ComConst.調査区分.リスト.Keys.Count - 1) {}
                    ComConst.調査区分.リスト.Keys.CopyTo(chosaKubunArr, 0)
                    tblCnt = (From dc In ComConst.調査票.テーブル名称 Select dc.Value.Length).Sum + (From dc In ComConst.個別結果表.テーブル名称 Select dc.Value.Length).Sum
                Else
                    Dim chosaKubun As String = cboChosakubun.SelectedValue.ToString()
                    chosaKubunArr = {chosaKubun}
                    tblCnt = ComConst.調査票.テーブル名称(chosaKubun).Length + ComConst.個別結果表.テーブル名称(chosaKubun).Length
                End If

                Try
                    ProgressDialog = New ProgressDialog

                    '進捗ダイアログを表示する
                    ProgressDialog.Maximum = tblCnt
                    ProgressDialog.Show(Me)

                    Dim backUpPath As String = IniFileInfo.BackUpPath
                    Dim fileNamePattern As String = Me.GetFileNamePattern(year)

                    For Each chosaKubun As String In chosaKubunArr
                        'テーブル取得（バックアップ）
                        Dim dtChosahyo As Dictionary(Of String, DataTable) = Nothing
                        Dim dtKobetsu As Dictionary(Of String, DataTable) = Nothing

                        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                            Select Case CommonInfo.Koutei
                                Case CommonInfo.KouteiKubun.Code.Center
                                    dtChosahyo = DAOChosahyo.GetBackUpChosahyoTable(db, chosaKubun, year, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
                                    dtKobetsu = DAOKobetsuKekkahyo.GetBackUpTable(db, chosaKubun, year, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
                                Case CommonInfo.KouteiKubun.Code.Kyoku
                                    dtChosahyo = DAOChosahyo.GetBackUpChosahyoTable(db, chosaKubun, year, CommonInfo.Kyoku, Nothing, Nothing)
                                    dtKobetsu = DAOKobetsuKekkahyo.GetBackUpTable(db, chosaKubun, year, CommonInfo.Kyoku, Nothing, Nothing)
                                Case CommonInfo.KouteiKubun.Code.Honsyo
                                    dtChosahyo = DAOChosahyo.GetBackUpChosahyoTable(db, chosaKubun, year, Nothing, Nothing, Nothing)
                                    dtKobetsu = DAOKobetsuKekkahyo.GetBackUpTable(db, chosaKubun, year, Nothing, Nothing, Nothing)
                            End Select
                        End Using

                        If dtChosahyo(ComConst.調査票.テーブル名称(chosaKubun)(0)).Rows.Count > 0 Then
                            'テーブルバックアップ
                            Me.BackUpTable(backUpPath, fileNamePattern, dtChosahyo)
                            bln = True
                        Else
                            '進捗を進める
                            ProgressDialog.AddValue = dtChosahyo.Count
                        End If

                        If dtKobetsu(ComConst.個別結果表.テーブル名称(chosaKubun)(0)).Rows.Count > 0 Then
                            'テーブルバックアップ
                            Me.BackUpTable(backUpPath, fileNamePattern, dtKobetsu)
                            bln = True
                        Else
                            '進捗を進める
                            ProgressDialog.AddValue = dtKobetsu.Count
                        End If
                    Next
                Finally
                    If Not ProgressDialog Is Nothing Then
                        '進捗ダイアログを閉じる
                        ProgressDialog.endDispose()
                        ProgressDialog = Nothing
                    End If
                End Try

                If bln Then
                    '完了メッセージ
                    Message.ShowMsgBox(MessageID.MSG_I_010, MsgBoxStyle.OkOnly)
                Else
                    'エラーメッセージ
                    Message.ShowMsgBox(MessageID.MSG_E_034, MsgBoxStyle.OkOnly)
                End If
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnRestore_Click(sender As Object, e As EventArgs) Handles btnRestore.Click
        Dim bln As Boolean = False

        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Dim year As String = txtYear.Text

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(year, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_019, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                Dim chosaKubunArr As String()
                Dim tblCnt As Integer

                If chkZenchosakubun.Checked Then
                    chosaKubunArr = New String(ComConst.調査区分.リスト.Keys.Count - 1) {}
                    ComConst.調査区分.リスト.Keys.CopyTo(chosaKubunArr, 0)
                    tblCnt = (From dc In ComConst.調査票.テーブル名称 Select dc.Value.Length).Sum + (From dc In ComConst.個別結果表.テーブル名称 Select dc.Value.Length).Sum
                Else
                    Dim chosaKubun As String = cboChosakubun.SelectedValue.ToString()
                    chosaKubunArr = {chosaKubun}
                    tblCnt = ComConst.調査票.テーブル名称(chosaKubun).Length + ComConst.個別結果表.テーブル名称(chosaKubun).Length
                End If

                Try
                    ProgressDialog = New ProgressDialog

                    '進捗ダイアログを表示する
                    ProgressDialog.Maximum = tblCnt
                    ProgressDialog.Show(Me)

                    Dim backUpPath As String = IniFileInfo.BackUpPath
                    Dim fileNamePattern As String = Me.GetFileNamePattern(year)

                    For Each chosaKubun As String In chosaKubunArr
                        'テーブル名称
                        Dim parentTable As String() = {ComConst.調査票.テーブル名称(chosaKubun)(0), ComConst.個別結果表.テーブル名称(chosaKubun)(0)}
                        Dim tblChosahyo As String() = ComConst.調査票.テーブル名称(chosaKubun)
                        Dim tblKobetsu As String() = ComConst.個別結果表.テーブル名称(chosaKubun)

                        If Me.CheckRestoreFileExists(backUpPath, fileNamePattern, parentTable) Then
                            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                                Try
                                    db.BeginTrans()

                                    '調査票データ削除（レストア）
                                    Select Case CommonInfo.Koutei
                                        Case CommonInfo.KouteiKubun.Code.Center
                                            DAOChosahyo.DeleteRestoreChosahyoTable(db, chosaKubun, year, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
                                            DAOKobetsuKekkahyo.DeleteRestoreTable(db, chosaKubun, year, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
                                        Case CommonInfo.KouteiKubun.Code.Kyoku
                                            DAOChosahyo.DeleteRestoreChosahyoTable(db, chosaKubun, year, CommonInfo.Kyoku, Nothing, Nothing)
                                            DAOKobetsuKekkahyo.DeleteRestoreTable(db, chosaKubun, year, CommonInfo.Kyoku, Nothing, Nothing)
                                        Case CommonInfo.KouteiKubun.Code.Honsyo
                                            DAOChosahyo.DeleteRestoreChosahyoTable(db, chosaKubun, year, Nothing, Nothing, Nothing)
                                            DAOKobetsuKekkahyo.DeleteRestoreTable(db, chosaKubun, year, Nothing, Nothing, Nothing)
                                    End Select

                                    'テーブルレストア
                                    Me.RestoreTable(db, backUpPath, fileNamePattern, tblChosahyo)
                                    Me.RestoreTable(db, backUpPath, fileNamePattern, tblKobetsu)

                                    db.CommitTrans()

                                    bln = True
                                Catch ex As Exception
                                    db.RollBackTrans()
                                    Throw ex
                                End Try
                            End Using
                        Else
                            '進捗を進める
                            ProgressDialog.AddValue = tblChosahyo.Length + tblKobetsu.Length
                        End If
                    Next
                Finally
                    If Not ProgressDialog Is Nothing Then
                        '進捗ダイアログを閉じる
                        ProgressDialog.endDispose()
                        ProgressDialog = Nothing
                    End If
                End Try

                If bln Then
                    '完了メッセージ
                    Message.ShowMsgBox(MessageID.MSG_I_011, MsgBoxStyle.OkOnly)
                Else
                    'エラーメッセージ
                    Message.ShowMsgBox(MessageID.MSG_E_035, MsgBoxStyle.OkOnly)
                End If
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
    ''' エラーチェック
    ''' </summary>
    ''' <param name="year"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(year As String, ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        '未入力チェック
        If year.Equals(String.Empty) Then
            msgId = MessageID.MSG_E_004
            Return ret
        End If

        '半角数字チェック
        If Not Regex.IsMatch(year, "^[0-9]+$") Then
            msgId = MessageID.MSG_E_022
            Return ret
        End If

        '調査区分選択チェック
        If Not chkZenchosakubun.Checked And cboChosakubun.SelectedValue Is Nothing Then
            msgId = MessageID.MSG_E_001
            Return ret
        End If

        ret = True

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
    ''' テーブルバックアップ
    ''' </summary>
    ''' <param name="outDir"></param>
    ''' <param name="fileName"></param>
    ''' <param name="dc"></param>
    ''' <remarks></remarks>
    Private Sub BackUpTable(outDir As String, fileName As String, dc As Dictionary(Of String, DataTable))
        Dim sjisEnc As Encoding = Encoding.GetEncoding(CODEPAGE_SHIFT_JIS)

        Dim filePathTemp As String = System.IO.Path.Combine(outDir, fileName)

        If Not System.IO.Directory.Exists(outDir) Then
            System.IO.Directory.CreateDirectory(outDir)
        End If

        For Each kv As KeyValuePair(Of String, DataTable) In dc
            Dim tableName As String = kv.Key
            Dim dt As DataTable = kv.Value

            If dt.Rows.Count > 0 Then
                Using sw As New System.IO.StreamWriter(String.Format(filePathTemp, tableName), False, sjisEnc)

                    For Each row As DataRow In dt.Rows
                        Dim arr As Object() = row.ItemArray().ToArray()
                        Dim repArr As New List(Of Object)
                        For Each elm As Object In arr
                            If Not IsDBNull(elm) Then
                                elm = CStr(elm).Replace(START_ADDITION, START_ADDITION & START_ADDITION)
                            End If
                            repArr.Add(elm)
                        Next
                        sw.WriteLine(START_ADDITION & String.Join(START_ADDITION & CSV_DELIMITER & END_ADDITION, repArr.ToArray) & END_ADDITION)
                    Next
                End Using
            End If

            '進捗を進める
            ProgressDialog.AddValue = 1
        Next
    End Sub

    ''' <summary>
    ''' テーブルレストア
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="inDir"></param>
    ''' <param name="fileName"></param>
    ''' <param name="tableNameArr"></param>
    ''' <remarks></remarks>
    Private Sub RestoreTable(db As DBAccess, inDir As String, fileName As String, tableNameArr As String())
        Dim sjisEnc As Encoding = Encoding.GetEncoding(CODEPAGE_SHIFT_JIS)

        Dim filePathTemp As String = System.IO.Path.Combine(inDir, fileName)

        For Each tableName As String In tableNameArr
            Dim path As String = String.Format(filePathTemp, tableName)

            If System.IO.File.Exists(path) Then
                Dim lstLine As New List(Of String())

                Using parser As New TextFieldParser(path, sjisEnc)

                    parser.TextFieldType = FieldType.Delimited
                    parser.SetDelimiters(CSV_DELIMITER)

                    While Not parser.EndOfData
                        Dim arr As String() = parser.ReadFields()
                        lstLine.Add(arr)
                    End While
                End Using

                'CSV一括読み込み
                DAOOther.BulkCopyCSV(db, tableName, lstLine)
            End If

            '進捗を進める
            ProgressDialog.AddValue = 1
        Next
    End Sub

    ''' <summary>
    ''' レストアファイル存在チェック
    ''' </summary>
    ''' <param name="inDir"></param>
    ''' <param name="fileName"></param>
    ''' <param name="tableNameArr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckRestoreFileExists(inDir As String, fileName As String, tableNameArr As String()) As Boolean
        Dim ret As Boolean = False

        Dim filePathTemp As String = System.IO.Path.Combine(inDir, fileName)

        For Each tableName As String In tableNameArr
            Dim path As String = String.Format(filePathTemp, tableName)

            If System.IO.File.Exists(path) Then
                ret = True
            End If
        Next

        Return ret
    End Function
End Class
