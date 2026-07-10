Public Class BRA01XXF

    ''' <summary>
    ''' タイトル設定
    ''' </summary>
    ''' <remarks></remarks>
    Friend Overrides Sub SetLabelTitle()
        '調査区分名を設定
        'Me.lblTitle.Text = CommonInfo.ChosakubunName
    End Sub

    Private Sub BRA01XXF_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        CommonInfo.ChosakubunName = Nothing
    End Sub
End Class
