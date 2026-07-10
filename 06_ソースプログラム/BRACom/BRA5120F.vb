Imports Microsoft.Office.Interop

''' <summary>
''' 地域選択画面
''' </summary>
''' <remarks></remarks>
Public Class BRA5120F

    Private Sub BRA5120F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            If CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Kyoku Then
                For Each kv As KeyValuePair(Of String, ComConst.地域.詳細) In ComConst.地域.リスト
                    If Not ComConst.地域.リスト(kv.Key).局.Equals(Integer.Parse(CommonInfo.Kyoku).ToString) Then
                        DirectCast(Me.Controls(ComConst.PRE_CHECKBOX & Integer.Parse(kv.Key).ToString(ComConst.DIGIT_2_FORMAT)), CheckBox).Enabled = False
                    End If
                Next
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnDecision_Click(sender As Object, e As EventArgs) Handles btnDecision.Click
        Try
            Dim dc As New Dictionary(Of String, String)
            For Each kv As KeyValuePair(Of String, ComConst.地域.詳細) In ComConst.地域.リスト
                If DirectCast(Me.Controls(ComConst.PRE_CHECKBOX & Integer.Parse(kv.Key).ToString(ComConst.DIGIT_2_FORMAT)), CheckBox).Checked Then
                    dc.Add(kv.Key, ComConst.地域.リスト(kv.Key).名称)
                End If
            Next

            '地域設定
            DirectCast(Me.Owner, BRA5110F).SetChiiki(dc)

            Me.Close()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnAllSelect_Click(sender As Object, e As EventArgs) Handles btnAllSelect.Click
        Try
            Me.SetCheckBoxAllCheck(True)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnAllCancel_Click(sender As Object, e As EventArgs) Handles btnAllCancel.Click
        Try
            Me.SetCheckBoxAllCheck(False)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnGroup1AllSelect_Click(sender As Object, e As EventArgs) Handles btnGroup1AllSelect.Click
        Try
            Me.SetCheckBoxAllCheck(True, ComConst.地域区分.全国)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnGroup2AllSelect_Click(sender As Object, e As EventArgs) Handles btnGroup2AllSelect.Click
        Try
            Me.SetCheckBoxAllCheck(True, ComConst.地域区分.大地域)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnGroup3AllSelect_Click(sender As Object, e As EventArgs) Handles btnGroup3AllSelect.Click
        Try
            Me.SetCheckBoxAllCheck(True, ComConst.地域区分.小地域)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnGroup4AllSelect_Click(sender As Object, e As EventArgs) Handles btnGroup4AllSelect.Click
        Try
            Me.SetCheckBoxAllCheck(True, ComConst.地域区分.農政局)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnGroup5AllSelect_Click(sender As Object, e As EventArgs) Handles btnGroup5AllSelect.Click
        Try
            Me.SetCheckBoxAllCheck(True, ComConst.地域区分.都府県)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' CheckBoxチェック全設定（活性時のみ）
    ''' </summary>
    ''' <param name="val"></param>
    ''' <param name="grp"></param>
    ''' <remarks></remarks>
    Private Sub SetCheckBoxAllCheck(val As Boolean, Optional grp As String = Nothing)
        For Each kv As KeyValuePair(Of String, ComConst.地域.詳細) In ComConst.地域.リスト
            Dim chk As CheckBox = DirectCast(Me.Controls(ComConst.PRE_CHECKBOX & Integer.Parse(kv.Key).ToString(ComConst.DIGIT_2_FORMAT)), CheckBox)

            If grp Is Nothing Then
                If chk.Enabled Then
                    chk.Checked = val
                End If
            Else
                If kv.Value.地域区分.Equals(grp) Then
                    If chk.Enabled Then
                        chk.Checked = val
                    End If
                End If
            End If
        Next
    End Sub
End Class
