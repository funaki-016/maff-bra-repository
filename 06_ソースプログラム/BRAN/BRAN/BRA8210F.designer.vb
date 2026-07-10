<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA8210F
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
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboTodofuken = New System.Windows.Forms.ComboBox()
        Me.btnOutPut = New System.Windows.Forms.Button()
        Me.cboSeisanhi = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtYear = New System.Windows.Forms.TextBox()
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
        Me.btnReturn.Location = New System.Drawing.Point(490, 209)
        Me.btnReturn.TabIndex = 13
        '
        'lblTitle
        '
        Me.lblTitle.Location = New System.Drawing.Point(91, 9)
        Me.lblTitle.Size = New System.Drawing.Size(425, 19)
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.Location = New System.Drawing.Point(91, 43)
        Me.lblSyori.Size = New System.Drawing.Size(425, 19)
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "労賃単価都道府県選択"
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(381, 209)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(90, 30)
        Me.btnSelect.TabIndex = 12
        Me.btnSelect.Text = "選択"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(195, 168)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "都道府県"
        '
        'cboTodofuken
        '
        Me.cboTodofuken.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTodofuken.FormattingEnabled = True
        Me.cboTodofuken.Location = New System.Drawing.Point(272, 165)
        Me.cboTodofuken.Name = "cboTodofuken"
        Me.cboTodofuken.Size = New System.Drawing.Size(72, 20)
        Me.cboTodofuken.TabIndex = 10
        '
        'btnOutPut
        '
        Me.btnOutPut.Location = New System.Drawing.Point(272, 209)
        Me.btnOutPut.Name = "btnOutPut"
        Me.btnOutPut.Size = New System.Drawing.Size(90, 30)
        Me.btnOutPut.TabIndex = 11
        Me.btnOutPut.Text = "出力"
        Me.btnOutPut.UseVisualStyleBackColor = True
        '
        'cboSeisanhi
        '
        Me.cboSeisanhi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSeisanhi.FormattingEnabled = True
        Me.cboSeisanhi.Location = New System.Drawing.Point(272, 87)
        Me.cboSeisanhi.Name = "cboSeisanhi"
        Me.cboSeisanhi.Size = New System.Drawing.Size(213, 20)
        Me.cboSeisanhi.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(195, 90)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 12)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "生産費"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(195, 129)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 12)
        Me.Label10.TabIndex = 7
        Me.Label10.Text = "調査年（産）"
        '
        'txtYear
        '
        Me.txtYear.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtYear.Location = New System.Drawing.Point(272, 126)
        Me.txtYear.MaxLength = 4
        Me.txtYear.Name = "txtYear"
        Me.txtYear.Size = New System.Drawing.Size(50, 19)
        Me.txtYear.TabIndex = 8
        Me.txtYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'BRA8210F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(606, 253)
        Me.Controls.Add(Me.txtYear)
        Me.Controls.Add(Me.cboSeisanhi)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.btnOutPut)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboTodofuken)
        Me.Controls.Add(Me.btnSelect)
        Me.Name = "BRA8210F"
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
        Me.Controls.SetChildIndex(Me.Label10, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.cboSeisanhi, 0)
        Me.Controls.SetChildIndex(Me.txtYear, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Protected WithEvents cboTodofuken As System.Windows.Forms.ComboBox
    Friend WithEvents btnOutPut As System.Windows.Forms.Button
    Protected WithEvents cboSeisanhi As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtYear As System.Windows.Forms.TextBox

End Class
