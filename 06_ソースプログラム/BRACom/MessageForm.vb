''' <summary>
''' メッセージフォーム
''' </summary>
''' <remarks></remarks>
Public Class MessageForm

    ''' <summary>
    ''' OKボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' 画面のタイトルを変更する
    ''' </summary>
    ''' <param name="title"></param>
    ''' <remarks></remarks>
    Public Sub ChangeFormTitle(ByVal title As String)
        Me.Text = title
    End Sub

    ''' <summary>
    ''' メッセージ表示
    ''' </summary>
    ''' <param name="title"></param>
    ''' <param name="label"></param>
    ''' <param name="msgId"></param>
    ''' <param name="msgPara"></param>
    ''' <param name="frm"></param>
    ''' <remarks></remarks>
    Public Sub ShowMessage(title As String, label As String, msgId As String, msgPara() As String, frm As Form)
        Me.Text = title
        Me.lblMessage.Text = label

        Dim msg As String
        msg = Message.GetMessage(msgId).ToString
        msg = Message.MyStringFormat(msg, msgPara)

        Me.txtMessage.Text = msg
        Me.ShowDialog(frm)
    End Sub
End Class