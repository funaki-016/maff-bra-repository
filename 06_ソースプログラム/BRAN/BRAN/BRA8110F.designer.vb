<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA8110F
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
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboTodofuken = New System.Windows.Forms.ComboBox()
        Me.btnOutPut = New System.Windows.Forms.Button()
        Me.SuspendLayout()
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
        Me.btnReturn.Location = New System.Drawing.Point(489, 138)
        Me.btnReturn.TabIndex = 8
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
        Me.lblSyori.Text = "毎勤データ都道府県選択"
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(378, 138)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(90, 30)
        Me.btnSelect.TabIndex = 7
        Me.btnSelect.Text = "選択"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(215, 94)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "都道府県"
        '
        'cboTodofuken
        '
        Me.cboTodofuken.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTodofuken.FormattingEnabled = True
        Me.cboTodofuken.Location = New System.Drawing.Point(289, 91)
        Me.cboTodofuken.Name = "cboTodofuken"
        Me.cboTodofuken.Size = New System.Drawing.Size(72, 20)
        Me.cboTodofuken.TabIndex = 6
        '
        'btnOutPut
        '
        Me.btnOutPut.Location = New System.Drawing.Point(271, 138)
        Me.btnOutPut.Name = "btnOutPut"
        Me.btnOutPut.Size = New System.Drawing.Size(90, 30)
        Me.btnOutPut.TabIndex = 9
        Me.btnOutPut.Text = "出力"
        Me.btnOutPut.UseVisualStyleBackColor = True
        '
        'BRA8110F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(606, 194)
        Me.Controls.Add(Me.btnOutPut)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboTodofuken)
        Me.Controls.Add(Me.btnSelect)
        Me.Name = "BRA8110F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.btnSelect, 0)
        Me.Controls.SetChildIndex(Me.cboTodofuken, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.btnOutPut, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Protected WithEvents cboTodofuken As System.Windows.Forms.ComboBox
    Friend WithEvents btnOutPut As System.Windows.Forms.Button

End Class
