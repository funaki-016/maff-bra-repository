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
    ''' <param name="frm"></param>
    ''' <param name="caption"></param>
    ''' <param name="label"></param>
    ''' <param name="message"></param>
    ''' <param name="icon"></param>
    ''' <remarks></remarks>
    Public Sub ShowMessage(frm As System.Windows.Forms.IWin32Window, caption As String, label As String, message As String, icon As MessageBoxIcon)
        Try
            Me.Text = caption
            Me.lblMessage.Text = label
            Me.pbxIcon.Image = GetIconImage(icon)
            Me.txtMessage.Text = message
            Me.ShowDialog(frm)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' アイコンイメージ取得
    ''' </summary>
    ''' <param name="msgBoxIcon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIconImage(msgBoxIcon As MessageBoxIcon) As Image
        Dim icon As Icon

        Select Case msgBoxIcon
            Case MessageBoxIcon.Question
                icon = SystemIcons.Question
            Case MessageBoxIcon.Information
                icon = SystemIcons.Information
            Case MessageBoxIcon.Warning
                icon = SystemIcons.Warning
            Case Else
                icon = SystemIcons.Error
        End Select

        Dim bmp As New Bitmap(32, 32)
        Dim g As Graphics = Graphics.FromImage(bmp)
        g.DrawIcon(icon, 0, 0)
        g.Dispose()
        Return bmp
    End Function
End Class