<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BRA9610F
    Inherits Maff.BRA.HeaderBaseForm

    'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnOutPut = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.nogyoChikiMastaYear = New System.Windows.Forms.ComboBox()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.txtUpdateDate = New System.Windows.Forms.TextBox()
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
        Me.btnReturn.Location = New System.Drawing.Point(502, 229)
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
        Me.lblSyori.Text = "農業地域類型マスタ管理"
        '
        'btnOutPut
        '
        Me.btnOutPut.Location = New System.Drawing.Point(406, 229)
        Me.btnOutPut.Name = "btnOutPut"
        Me.btnOutPut.Size = New System.Drawing.Size(90, 30)
        Me.btnOutPut.TabIndex = 15
        Me.btnOutPut.Text = "出力"
        Me.btnOutPut.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(120, 145)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 12)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "更新日時"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(70, 103)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(103, 12)
        Me.Label10.TabIndex = 8
        Me.Label10.Text = "農業地域類型マスタ"
        '
        'nogyoChikiMastaYear
        '
        Me.nogyoChikiMastaYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.nogyoChikiMastaYear.FormattingEnabled = True
        Me.nogyoChikiMastaYear.Location = New System.Drawing.Point(179, 99)
        Me.nogyoChikiMastaYear.Name = "nogyoChikiMastaYear"
        Me.nogyoChikiMastaYear.Size = New System.Drawing.Size(210, 20)
        Me.nogyoChikiMastaYear.TabIndex = 9
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(310, 229)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(90, 30)
        Me.btnImport.TabIndex = 14
        Me.btnImport.Text = "取込"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'txtUpdateDate
        '
        Me.txtUpdateDate.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtUpdateDate.Location = New System.Drawing.Point(179, 142)
        Me.txtUpdateDate.MaxLength = 50
        Me.txtUpdateDate.Name = "txtUpdateDate"
        Me.txtUpdateDate.ReadOnly = True
        Me.txtUpdateDate.Size = New System.Drawing.Size(100, 19)
        Me.txtUpdateDate.TabIndex = 17
        Me.txtUpdateDate.TabStop = False
        Me.txtUpdateDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'BRA9610F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(606, 273)
        Me.Controls.Add(Me.txtUpdateDate)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.nogyoChikiMastaYear)
        Me.Controls.Add(Me.btnOutPut)
        Me.Name = "BRA9610F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.btnOutPut, 0)
        Me.Controls.SetChildIndex(Me.nogyoChikiMastaYear, 0)
        Me.Controls.SetChildIndex(Me.Label10, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.btnImport, 0)
        Me.Controls.SetChildIndex(Me.txtUpdateDate, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOutPut As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Protected WithEvents nogyoChikiMastaYear As System.Windows.Forms.ComboBox
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents txtUpdateDate As TextBox
End Class
