<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HeaderBaseForm
    Inherits BaseForm

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblKoutei = New System.Windows.Forms.Label()
        Me.lblInformation2 = New System.Windows.Forms.Label()
        Me.lblInformation3 = New System.Windows.Forms.Label()
        Me.btnReturn = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblSyori
        '
        Me.lblSyori.AutoSize = False
        Me.lblSyori.Location = New System.Drawing.Point(77, 43)
        Me.lblSyori.Margin = New System.Windows.Forms.Padding(0)
        Me.lblSyori.Size = New System.Drawing.Size(606, 19)
        Me.lblSyori.TabIndex = 4
        Me.lblSyori.Text = "調査情報入力"
        Me.lblSyori.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblTitle
        '
        Me.lblTitle.Font = New System.Drawing.Font("ＭＳ ゴシック", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(77, 9)
        Me.lblTitle.Margin = New System.Windows.Forms.Padding(0)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(606, 19)
        Me.lblTitle.TabIndex = 3
        Me.lblTitle.Text = "農業経営統計調査"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblKoutei
        '
        Me.lblKoutei.BackColor = System.Drawing.Color.White
        Me.lblKoutei.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblKoutei.Location = New System.Drawing.Point(667, 11)
        Me.lblKoutei.Margin = New System.Windows.Forms.Padding(0)
        Me.lblKoutei.Name = "lblKoutei"
        Me.lblKoutei.Size = New System.Drawing.Size(130, 20)
        Me.lblKoutei.TabIndex = 5
        Me.lblKoutei.Text = "○○工程"
        Me.lblKoutei.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblInformation2
        '
        Me.lblInformation2.BackColor = System.Drawing.Color.White
        Me.lblInformation2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblInformation2.Location = New System.Drawing.Point(11, 11)
        Me.lblInformation2.Margin = New System.Windows.Forms.Padding(0)
        Me.lblInformation2.Name = "lblInformation2"
        Me.lblInformation2.Size = New System.Drawing.Size(130, 20)
        Me.lblInformation2.TabIndex = 1
        Me.lblInformation2.Text = "局名、本省"
        Me.lblInformation2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblInformation3
        '
        Me.lblInformation3.BackColor = System.Drawing.Color.White
        Me.lblInformation3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblInformation3.Location = New System.Drawing.Point(11, 31)
        Me.lblInformation3.Margin = New System.Windows.Forms.Padding(0)
        Me.lblInformation3.Name = "lblInformation3"
        Me.lblInformation3.Size = New System.Drawing.Size(130, 20)
        Me.lblInformation3.TabIndex = 2
        Me.lblInformation3.Text = "実査設置拠点名"
        Me.lblInformation3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblInformation3.Visible = False
        '
        'btnReturn
        '
        Me.btnReturn.Location = New System.Drawing.Point(704, 512)
        Me.btnReturn.Name = "btnReturn"
        Me.btnReturn.Size = New System.Drawing.Size(90, 30)
        Me.btnReturn.TabIndex = 12
        Me.btnReturn.Text = "戻る"
        Me.btnReturn.UseVisualStyleBackColor = True
        '
        'HeaderBaseForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(808, 556)
        Me.Controls.Add(Me.btnReturn)
        Me.Controls.Add(Me.lblInformation3)
        Me.Controls.Add(Me.lblInformation2)
        Me.Controls.Add(Me.lblKoutei)
        Me.Controls.Add(Me.lblTitle)
        Me.Name = "HeaderBaseForm"
        Me.Padding = New System.Windows.Forms.Padding(11)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.ResumeLayout(False)

End Sub
    Protected WithEvents lblKoutei As System.Windows.Forms.Label
    Protected WithEvents lblInformation2 As System.Windows.Forms.Label
    Protected WithEvents lblInformation3 As System.Windows.Forms.Label
    Protected Friend WithEvents btnReturn As System.Windows.Forms.Button
    Protected WithEvents lblTitle As System.Windows.Forms.Label
End Class
