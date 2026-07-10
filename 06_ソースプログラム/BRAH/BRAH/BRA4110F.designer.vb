<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA4110F
    Inherits Maff.BRA.ExcelInputBaseForm

    'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnOutPut = New System.Windows.Forms.Button()
        Me.btnSubstitute = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtParameters = New System.Windows.Forms.TextBox()
        Me.txtChosanen = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'lblKoutei
        '
        Me.lblKoutei.TabIndex = 4
        Me.lblKoutei.Text = "本省工程"
        '
        'lblInformation2
        '
        Me.lblInformation2.TabIndex = 0
        Me.lblInformation2.Text = "本省"
        '
        'lblInformation3
        '
        Me.lblInformation3.TabIndex = 1
        '
        'btnReturn
        '
        Me.btnReturn.Location = New System.Drawing.Point(697, 162)
        Me.btnReturn.TabIndex = 11
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "個別結果表一括登録"
        '
        'btnOutPut
        '
        Me.btnOutPut.Location = New System.Drawing.Point(601, 162)
        Me.btnOutPut.Name = "btnOutPut"
        Me.btnOutPut.Size = New System.Drawing.Size(90, 30)
        Me.btnOutPut.TabIndex = 10
        Me.btnOutPut.Text = "取込"
        Me.btnOutPut.UseVisualStyleBackColor = True
        '
        'btnSubstitute
        '
        Me.btnSubstitute.Location = New System.Drawing.Point(606, 125)
        Me.btnSubstitute.Name = "btnSubstitute"
        Me.btnSubstitute.Size = New System.Drawing.Size(77, 21)
        Me.btnSubstitute.TabIndex = 9
        Me.btnSubstitute.Text = "参照"
        Me.btnSubstitute.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(158, 129)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 12)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "個別結果表CSV"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(158, 92)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 12)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "調査年（産）"
        '
        'txtParameters
        '
        Me.txtParameters.Location = New System.Drawing.Point(252, 126)
        Me.txtParameters.Name = "txtParameters"
        Me.txtParameters.Size = New System.Drawing.Size(348, 19)
        Me.txtParameters.TabIndex = 12
        '
        'txtChosanen
        '
        Me.txtChosanen.Location = New System.Drawing.Point(252, 89)
        Me.txtChosanen.Name = "txtChosanen"
        Me.txtChosanen.Size = New System.Drawing.Size(72, 19)
        Me.txtChosanen.TabIndex = 13
        '
        'BRA4110F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 210)
        Me.Controls.Add(Me.txtChosanen)
        Me.Controls.Add(Me.txtParameters)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnSubstitute)
        Me.Controls.Add(Me.btnOutPut)
        Me.Name = "BRA4110F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.btnOutPut, 0)
        Me.Controls.SetChildIndex(Me.btnSubstitute, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.Label10, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.txtParameters, 0)
        Me.Controls.SetChildIndex(Me.txtChosanen, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOutPut As System.Windows.Forms.Button
    Friend WithEvents btnSubstitute As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtParameters As System.Windows.Forms.TextBox
    Friend WithEvents txtChosanen As System.Windows.Forms.TextBox

End Class
