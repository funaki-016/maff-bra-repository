<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA9110F
    Inherits Maff.BRA.HeaderBaseForm

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
        Me.btnRestore = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cboChosakubun = New System.Windows.Forms.ComboBox()
        Me.cboKubun2 = New System.Windows.Forms.ComboBox()
        Me.cboKubun1 = New System.Windows.Forms.ComboBox()
        Me.btnBackUp = New System.Windows.Forms.Button()
        Me.txtYear = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.chkZenchosakubun = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout
        '
        'lblKoutei
        '
        Me.lblKoutei.Location = New System.Drawing.Point(450, 11)
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
        Me.btnReturn.Location = New System.Drawing.Point(489, 276)
        Me.btnReturn.TabIndex = 16
        '
        'lblTitle
        '
        Me.lblTitle.Size = New System.Drawing.Size(425, 19)
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.Size = New System.Drawing.Size(425, 19)
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "ＤＢ管理"
        '
        'btnRestore
        '
        Me.btnRestore.Location = New System.Drawing.Point(308, 276)
        Me.btnRestore.Name = "btnRestore"
        Me.btnRestore.Size = New System.Drawing.Size(90, 30)
        Me.btnRestore.TabIndex = 15
        Me.btnRestore.Text = "レストア"
        Me.btnRestore.UseVisualStyleBackColor = true
        '
        'Label2
        '
        Me.Label2.AutoSize = true
        Me.Label2.Location = New System.Drawing.Point(95, 224)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "調査区分"
        '
        'Label1
        '
        Me.Label1.AutoSize = true
        Me.Label1.Location = New System.Drawing.Point(95, 178)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 12)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "区分２"
        '
        'Label10
        '
        Me.Label10.AutoSize = true
        Me.Label10.Location = New System.Drawing.Point(95, 134)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(37, 12)
        Me.Label10.TabIndex = 8
        Me.Label10.Text = "区分１"
        '
        'cboChosakubun
        '
        Me.cboChosakubun.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChosakubun.FormattingEnabled = true
        Me.cboChosakubun.Location = New System.Drawing.Point(171, 221)
        Me.cboChosakubun.Name = "cboChosakubun"
        Me.cboChosakubun.Size = New System.Drawing.Size(277, 20)
        Me.cboChosakubun.TabIndex = 13
        '
        'cboKubun2
        '
        Me.cboKubun2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKubun2.FormattingEnabled = true
        Me.cboKubun2.Location = New System.Drawing.Point(171, 175)
        Me.cboKubun2.Name = "cboKubun2"
        Me.cboKubun2.Size = New System.Drawing.Size(277, 20)
        Me.cboKubun2.TabIndex = 11
        '
        'cboKubun1
        '
        Me.cboKubun1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKubun1.FormattingEnabled = true
        Me.cboKubun1.Location = New System.Drawing.Point(171, 131)
        Me.cboKubun1.Name = "cboKubun1"
        Me.cboKubun1.Size = New System.Drawing.Size(277, 20)
        Me.cboKubun1.TabIndex = 9
        '
        'btnBackUp
        '
        Me.btnBackUp.Location = New System.Drawing.Point(186, 276)
        Me.btnBackUp.Name = "btnBackUp"
        Me.btnBackUp.Size = New System.Drawing.Size(90, 30)
        Me.btnBackUp.TabIndex = 14
        Me.btnBackUp.Text = "バックアップ"
        Me.btnBackUp.UseVisualStyleBackColor = true
        '
        'txtYear
        '
        Me.txtYear.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtYear.Location = New System.Drawing.Point(171, 86)
        Me.txtYear.MaxLength = 4
        Me.txtYear.Name = "txtYear"
        Me.txtYear.Size = New System.Drawing.Size(50, 19)
        Me.txtYear.TabIndex = 6
        Me.txtYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = true
        Me.Label3.Location = New System.Drawing.Point(95, 89)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 12)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "調査年（産）"
        '
        'chkZenchosakubun
        '
        Me.chkZenchosakubun.AutoSize = true
        Me.chkZenchosakubun.Location = New System.Drawing.Point(364, 86)
        Me.chkZenchosakubun.Name = "chkZenchosakubun"
        Me.chkZenchosakubun.Size = New System.Drawing.Size(84, 16)
        Me.chkZenchosakubun.TabIndex = 7
        Me.chkZenchosakubun.Text = "全調査区分"
        Me.chkZenchosakubun.UseVisualStyleBackColor = true
        '
        'BRA9110F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 12!)
        Me.ClientSize = New System.Drawing.Size(606, 337)
        Me.Controls.Add(Me.chkZenchosakubun)
        Me.Controls.Add(Me.txtYear)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnBackUp)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.cboChosakubun)
        Me.Controls.Add(Me.cboKubun2)
        Me.Controls.Add(Me.cboKubun1)
        Me.Controls.Add(Me.btnRestore)
        Me.Name = "BRA9110F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.btnRestore, 0)
        Me.Controls.SetChildIndex(Me.cboKubun1, 0)
        Me.Controls.SetChildIndex(Me.cboKubun2, 0)
        Me.Controls.SetChildIndex(Me.cboChosakubun, 0)
        Me.Controls.SetChildIndex(Me.Label10, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.btnBackUp, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.txtYear, 0)
        Me.Controls.SetChildIndex(Me.chkZenchosakubun, 0)
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents btnRestore As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Protected WithEvents cboChosakubun As System.Windows.Forms.ComboBox
    Protected WithEvents cboKubun2 As System.Windows.Forms.ComboBox
    Protected WithEvents cboKubun1 As System.Windows.Forms.ComboBox
    Friend WithEvents btnBackUp As System.Windows.Forms.Button
    Friend WithEvents txtYear As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkZenchosakubun As System.Windows.Forms.CheckBox

End Class
