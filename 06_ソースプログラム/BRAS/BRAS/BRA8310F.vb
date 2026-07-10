''' <summary>
''' 専門調査員一覧・帳票出力画面
''' </summary>
''' <remarks></remarks>
Public Class BRA8310F
    ''' <summary>
    ''' 画面起動
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA8310F_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Try

            'DataGridView設定
            ComUtil.ConfigDgv(Me.dgvList)

            '一覧表示
            Me.ShowList()

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click
        Try
            Dim frm As New BRA8410F(BRA8410F.編集モード種別.新規)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Try
            'ユーザーID取得
            Dim userID As String = Me.GetUserID()

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(userID, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim frm As New BRA8410F(BRA8410F.編集モード種別.修正, userID)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            'ユーザーID取得
            Dim userID As String = Me.GetUserID()

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(userID, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_007, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    Try
                        db.BeginTrans()

                        '専門調査員管理削除
                        DAOOther.DeleteSenmonChosain(db, userID)

                        '専門調査員担当調査客体削除
                        DAOOther.DeleteSenmonChosainKyakutai(db, userID)

                        db.CommitTrans()
                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try
                End Using

                '完了メッセージ
                Message.ShowMsgBox(MessageID.MSG_I_005, MsgBoxStyle.OkOnly)

                '一覧表示
                Me.ShowList()
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnEditKyakutai_Click(sender As Object, e As EventArgs) Handles btnEditKyakutai.Click
        Try
            'ユーザーID取得
            Dim userID As String = Me.GetUserID()

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(userID, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim frm As New BRA8510F(userID)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub dgvList_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvList.CellValueChanged
        If e.ColumnIndex = 0 Then
            If CType(dgvList(e.ColumnIndex, e.RowIndex).Value, Boolean) = True Then
                For rowIndex = 0 To dgvList.Rows.Count - 1
                    If rowIndex <> e.RowIndex Then
                        dgvList(0, rowIndex).Value = False
                        dgvList(0, rowIndex).ReadOnly = False
                    End If
                Next
                dgvList(e.ColumnIndex, e.RowIndex).ReadOnly = True
            End If
        End If
    End Sub

    Private Sub dgvList_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvList.CurrentCellDirtyStateChanged
        If dgvList.CurrentCellAddress.X = 0 AndAlso dgvList.IsCurrentCellDirty Then
            dgvList.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ShowList()
        Dim dt As DataTable

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dt = DAOOther.GetSenmonChosainList(db)
        End Using

        dgvList.Rows.Clear()

        For Each row As DataRow In dt.Rows
            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(1).Value = row("ユーザーID").ToString
            dgvList.Rows(i).Cells(2).Value = row("氏名").ToString
            dgvList.Rows(i).Cells(3).Value = DateTime.Parse(row("更新日付").ToString).ToString(ComConst.DATETIME_FORMAT)
            dgvList.Rows(i).Cells(4).Value = row("担当客体数").ToString
        Next
    End Sub

    ''' <summary>
    ''' ユーザーID取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetUserID() As String
        Dim ret As String = Nothing

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                ret = dgvList.Rows(i).Cells(1).Value.ToString
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="userID"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(userID As String, ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'ユーザーIDチェック
        If userID Is Nothing Then
            msgId = MessageID.MSG_E_011
            Return ret
        End If

        ret = True

        Return ret
    End Function

    Private Sub btnOutput_Click(sender As Object, e As EventArgs) Handles btnOutput.Click
        Try

            Dim fileName As String = ComConst.専門調査員及び担当調査客体一覧.出力用ファイル名称.reportName & "_" _
                                    & CommonInfo.Jimusyo & "_" _
                                    & CommonInfo.Center & ".xlsx"

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            
            '専門調査員管理存在チェック
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                If Not DAOOther.CheckSenmonChosainExist(db) Then
                    'エラーメッセージ
                    Message.ShowMsgBox(MessageID.MSG_E_023, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
            End Using

            '専門調査員及び担当調査客体一覧出力
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New BRA8310R(filePath)
                    ret = ExcelOutput.Execute(MessageID.MSG_Q_004)
                End Using

                If ret = ExcelOutputBaseClass.enmOutputReturn.OK Then
                    '完了メッセージ
                    Message.ShowMsgBox(MessageID.MSG_I_002, MsgBoxStyle.OkOnly)
                End If
            Catch ex As ExcelOutputBaseClass.SaveAsException
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_006, MsgBoxStyle.OkOnly)
            Catch ex As Exception
                Throw
            End Try

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub
End Class
