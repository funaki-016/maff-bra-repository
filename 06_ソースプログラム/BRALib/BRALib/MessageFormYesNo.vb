''' <summary>
''' メッセージフォーム(はい/いいえ)
''' </summary>
''' <remarks></remarks>
Public Class MessageFormYesNo

    Private e As DialogResult
    Private ret As DialogResult

    ''' <summary>
    ''' 画面表示
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function ShowDialog() As Windows.Forms.DialogResult
        MyBase.ShowDialog()
        Return ret
    End Function

    ''' <summary>
    ''' はいボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnYes_Click(sender As Object, e As EventArgs) Handles btnYes.Click
        ret = Windows.Forms.DialogResult.Yes
        Me.Close()
    End Sub

    ''' <summary>
    ''' いいえボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnNo_Click(sender As Object, e As EventArgs) Handles btnNo.Click
        ret = Windows.Forms.DialogResult.No
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
    Public Function ShowMessage(frm As System.Windows.Forms.IWin32Window, caption As String, label As String, message As String, icon As MessageBoxIcon) As Windows.Forms.DialogResult
        Try
            Me.Text = caption
            Me.lblMessage.Text = label
            Me.pbxIcon.Image = GetIconImage(icon)
            Me.txtMessage.Text = message
            Me.ShowDialog(frm)
            Return ret
        Catch ex As Exception
            Throw
        End Try
    End Function

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