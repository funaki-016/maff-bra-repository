<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA3510F
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
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtUpdateDate = New System.Windows.Forms.TextBox()
        Me.cboVerKubun = New System.Windows.Forms.ComboBox()
        Me.lblVerKubun = New System.Windows.Forms.Label()
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
        Me.btnReturn.Location = New System.Drawing.Point(697, 152)
        Me.btnReturn.TabIndex = 10
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "個別結果検討表作成論理管理"
        '
        'btnOutPut
        '
        Me.btnOutPut.Location = New System.Drawing.Point(601, 152)
        Me.btnOutPut.Name = "btnOutPut"
        Me.btnOutPut.Size = New System.Drawing.Size(90, 30)
        Me.btnOutPut.TabIndex = 9
        Me.btnOutPut.Text = "出力"
        Me.btnOutPut.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(409, 152)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(90, 30)
        Me.btnEdit.TabIndex = 7
        Me.btnEdit.Text = "修正"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(505, 152)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(90, 30)
        Me.btnImport.TabIndex = 8
        Me.btnImport.Text = "取込"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(280, 99)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 12)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "更新日時"
        '
        'txtUpdateDate
        '
        Me.txtUpdateDate.Location = New System.Drawing.Point(350, 96)
        Me.txtUpdateDate.MaxLength = 50
        Me.txtUpdateDate.Name = "txtUpdateDate"
        Me.txtUpdateDate.ReadOnly = True
        Me.txtUpdateDate.Size = New System.Drawing.Size(100, 19)
        Me.txtUpdateDate.TabIndex = 6
        Me.txtUpdateDate.TabStop = False
        Me.txtUpdateDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'cboVerKubun
        '
        Me.cboVerKubun.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboVerKubun.FormattingEnabled = True
        Me.cboVerKubun.Location = New System.Drawing.Point(159, 96)
        Me.cboVerKubun.Name = "cboVerKubun"
        Me.cboVerKubun.Size = New System.Drawing.Size(100, 20)
        Me.cboVerKubun.TabIndex = 14
        '
        'lblVerKubun
        '
        Me.lblVerKubun.AutoSize = True
        Me.lblVerKubun.Location = New System.Drawing.Point(75, 99)
        Me.lblVerKubun.Name = "lblVerKubun"
        Me.lblVerKubun.Size = New System.Drawing.Size(74, 12)
        Me.lblVerKubun.TabIndex = 13
        Me.lblVerKubun.Text = "バージョン区分"
        '
        'BRA3510F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 210)
        Me.Controls.Add(Me.cboVerKubun)
        Me.Controls.Add(Me.lblVerKubun)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtUpdateDate)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnOutPut)
        Me.Name = "BRA3510F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.btnOutPut, 0)
        Me.Controls.SetChildIndex(Me.btnEdit, 0)
        Me.Controls.SetChildIndex(Me.btnImport, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.txtUpdateDate, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.lblVerKubun, 0)
        Me.Controls.SetChildIndex(Me.cboVerKubun, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOutPut As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtUpdateDate As System.Windows.Forms.TextBox
    Friend WithEvents cboVerKubun As ComboBox
    Friend WithEvents lblVerKubun As Label
End Class
