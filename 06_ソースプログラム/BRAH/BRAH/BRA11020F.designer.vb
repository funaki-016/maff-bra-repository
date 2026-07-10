<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BRA11020F
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
        Me.BtnInput = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblKoutei
        '
        Me.lblKoutei.Location = New System.Drawing.Point(404, 11)
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
        Me.btnReturn.Location = New System.Drawing.Point(444, 93)
        Me.btnReturn.Margin = New System.Windows.Forms.Padding(4)
        Me.btnReturn.TabIndex = 6
        '
        'lblTitle
        '
        Me.lblTitle.Size = New System.Drawing.Size(400, 19)
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.Size = New System.Drawing.Size(400, 19)
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "米在庫調査票データCSV取込"
        '
        'BtnInput
        '
        Me.BtnInput.Location = New System.Drawing.Point(325, 93)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(90, 30)
        Me.BtnInput.TabIndex = 5
        Me.BtnInput.Text = "取込"
        Me.BtnInput.UseVisualStyleBackColor = True
        '
        'BRA11020F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(549, 141)
        Me.Controls.Add(Me.BtnInput)
        Me.Margin = New System.Windows.Forms.Padding(5)
        Me.Name = "BRA11020F"
        Me.Padding = New System.Windows.Forms.Padding(15, 14, 15, 14)
        Me.Text = "農業経営統計調査 - 本省工程 - 米在庫調査票データCSV取込"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.BtnInput, 0)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnInput As System.Windows.Forms.Button
End Class
