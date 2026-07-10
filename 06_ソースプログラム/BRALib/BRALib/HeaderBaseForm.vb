Public Class HeaderBaseForm

    'メインメニューフラグ
    Public mainMenu As Boolean = False

    ''' <summary>
    ''' 画面起動
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BaseForm_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Dim kouteiKubunName As String = String.Empty

        '工程別設定
        Select Case CommonInfo.Koutei
            Case CommonInfo.KouteiKubun.Code.Center
                kouteiKubunName = CommonInfo.KouteiKubun.Name.Center

                Me.lblInformation2.Text = CommonInfo.JimusyoName

                Me.lblInformation3.Text = CommonInfo.CenterName
                Me.lblInformation3.Visible = True

            Case CommonInfo.KouteiKubun.Code.Kyoku
                kouteiKubunName = CommonInfo.KouteiKubun.Name.Kyoku

                Me.lblInformation2.Text = CommonInfo.KyokuName

            Case Else
                kouteiKubunName = CommonInfo.KouteiKubun.Name.Honsyo

                Me.lblInformation2.Text = CommonInfo.HONSYONAME

        End Select

        'タイトル設定
        Me.SetLabelTitle()

        'キャプション設定
        Me.Text = Me.lblTitle.Text & " - " & kouteiKubunName & " - " & Me.lblSyori.Text

        '画面右側の工程ラベル設定
        Me.lblKoutei.Text = kouteiKubunName

    End Sub

    ''' <summary>
    ''' フォームを閉じる前
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub HeaderBaseForm_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not mainMenu Then
            If Me.Owner IsNot Nothing Then
                Me.Owner.Show()
                Me.Owner.Refresh()
            End If
        End If
    End Sub

    ''' <summary>
    ''' 終了ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Overridable Sub BtnReturn_Click(sender As Object, e As EventArgs) Handles btnReturn.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' 処理名取得
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Function getSyoriName() As String
        Return Me.lblSyori.Text
    End Function

    ''' <summary>
    ''' タイトル設定
    ''' </summary>
    ''' <remarks></remarks>
    Friend Overridable Sub SetLabelTitle()
        If Not CommonInfo.ChosakubunName Is Nothing Then
            '調査区分名を設定
            Me.lblTitle.Text = CommonInfo.ChosakubunName
        End If
    End Sub
End Class