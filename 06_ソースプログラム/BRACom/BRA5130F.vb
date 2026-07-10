Imports Microsoft.Office.Interop

''' <summary>
''' 部門選択画面
''' </summary>
''' <remarks></remarks>
Public Class BRA5130F

    Private Sub BRA5130F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnDecision_Click(sender As Object, e As EventArgs) Handles btnDecision.Click
        Try
            Dim dc As New Dictionary(Of String, String)
            For Each kv As KeyValuePair(Of String, ComConst.部門.詳細) In ComConst.部門.リスト
                If DirectCast(Me.Controls(ComConst.PRE_CHECKBOX & Integer.Parse(kv.Key).ToString(ComConst.DIGIT_2_FORMAT)), CheckBox).Checked Then
                    dc.Add(kv.Key, ComConst.部門.リスト(kv.Key).名称)
                End If
            Next

            '部門設定
            DirectCast(Me.Owner, BRA5110F).SetBumon(dc)

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
            Me.SetCheckBoxAllCheck(True, ComConst.営農類型区分.畑作)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Me.SetCheckBoxAllCheck(True, ComConst.営農類型区分.露地野菜作)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Me.SetCheckBoxAllCheck(True, ComConst.営農類型区分.施設野菜作)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Me.SetCheckBoxAllCheck(True, ComConst.営農類型区分.果樹作)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Me.SetCheckBoxAllCheck(True, ComConst.営農類型区分.施設花き作)
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
        For Each kv As KeyValuePair(Of String, ComConst.部門.詳細) In ComConst.部門.リスト
            Dim chk As CheckBox = DirectCast(Me.Controls(ComConst.PRE_CHECKBOX & Integer.Parse(kv.Key).ToString(ComConst.DIGIT_2_FORMAT)), CheckBox)

            If grp Is Nothing Then
                If chk.Enabled Then
                    chk.Checked = val
                End If
            Else
                If kv.Value.営農類型.Equals(grp) Then
                    If chk.Enabled Then
                        chk.Checked = val
                    End If
                End If
            End If
        Next
    End Sub
End Class
