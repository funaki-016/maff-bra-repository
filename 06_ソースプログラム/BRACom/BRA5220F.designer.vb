<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA5220F
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtAfterName = New System.Windows.Forms.TextBox()
        Me.txtBeforeName = New System.Windows.Forms.TextBox()
        Me.btnModify = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblKoutei
        '
        Me.lblKoutei.Location = New System.Drawing.Point(627, 11)
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
        Me.btnReturn.Location = New System.Drawing.Point(670, 222)
        Me.btnReturn.TabIndex = 10
        '
        'lblTitle
        '
        Me.lblTitle.Location = New System.Drawing.Point(84, 9)
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.Location = New System.Drawing.Point(84, 43)
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "集計結果表名称変更画面"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(170, 160)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 12)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "変更後名称"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(170, 108)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 12)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "変更前名称"
        '
        'txtAfterName
        '
        Me.txtAfterName.Location = New System.Drawing.Point(244, 157)
        Me.txtAfterName.MaxLength = 70
        Me.txtAfterName.Name = "txtAfterName"
        Me.txtAfterName.Size = New System.Drawing.Size(330, 19)
        Me.txtAfterName.TabIndex = 8
        '
        'txtBeforeName
        '
        Me.txtBeforeName.Location = New System.Drawing.Point(244, 105)
        Me.txtBeforeName.MaxLength = 70
        Me.txtBeforeName.Name = "txtBeforeName"
        Me.txtBeforeName.ReadOnly = True
        Me.txtBeforeName.Size = New System.Drawing.Size(330, 19)
        Me.txtBeforeName.TabIndex = 6
        Me.txtBeforeName.TabStop = False
        '
        'btnModify
        '
        Me.btnModify.Location = New System.Drawing.Point(574, 222)
        Me.btnModify.Name = "btnModify"
        Me.btnModify.Size = New System.Drawing.Size(90, 30)
        Me.btnModify.TabIndex = 9
        Me.btnModify.Text = "修正"
        Me.btnModify.UseVisualStyleBackColor = True
        '
        'BRA5220F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(774, 272)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtAfterName)
        Me.Controls.Add(Me.txtBeforeName)
        Me.Controls.Add(Me.btnModify)
        Me.Name = "BRA5220F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.btnModify, 0)
        Me.Controls.SetChildIndex(Me.txtBeforeName, 0)
        Me.Controls.SetChildIndex(Me.txtAfterName, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtAfterName As System.Windows.Forms.TextBox
    Friend WithEvents txtBeforeName As System.Windows.Forms.TextBox
    Friend WithEvents btnModify As System.Windows.Forms.Button

End Class
