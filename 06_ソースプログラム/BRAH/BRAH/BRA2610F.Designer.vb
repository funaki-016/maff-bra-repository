<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BRA2610F
    Inherits Maff.BRA.ExcelInputBaseForm

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
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtUpdateDate = New System.Windows.Forms.TextBox()
        Me.cboChosanen = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'lblKoutei
        '
        Me.lblKoutei.Location = New System.Drawing.Point(878, 14)
        Me.lblKoutei.Size = New System.Drawing.Size(181, 30)
        Me.lblKoutei.TabIndex = 4
        Me.lblKoutei.Text = "本省工程"
        '
        'lblInformation2
        '
        Me.lblInformation2.Size = New System.Drawing.Size(230, 30)
        Me.lblInformation2.TabIndex = 0
        Me.lblInformation2.Text = "本省"
        '
        'lblInformation3
        '
        Me.lblInformation3.TabIndex = 1
        '
        'btnReturn
        '
        Me.btnReturn.Location = New System.Drawing.Point(929, 190)
        Me.btnReturn.Margin = New System.Windows.Forms.Padding(5)
        Me.btnReturn.TabIndex = 10
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.Location = New System.Drawing.Point(-46, 44)
        Me.lblSyori.Size = New System.Drawing.Size(1077, 30)
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "制度受取金・積立金等項目"
        '
        'btnOutPut
        '
        Me.btnOutPut.Location = New System.Drawing.Point(801, 190)
        Me.btnOutPut.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOutPut.Name = "btnOutPut"
        Me.btnOutPut.Size = New System.Drawing.Size(120, 38)
        Me.btnOutPut.TabIndex = 9
        Me.btnOutPut.Text = "出力"
        Me.btnOutPut.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(545, 190)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(120, 38)
        Me.btnEdit.TabIndex = 7
        Me.btnEdit.Text = "修正"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(673, 190)
        Me.btnImport.Margin = New System.Windows.Forms.Padding(4)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(120, 38)
        Me.btnImport.TabIndex = 8
        Me.btnImport.Text = "取込"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(527, 124)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 15)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "更新日時"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(265, 124)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(83, 15)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "調査年（産）"
        '
        'txtUpdateDate
        '
        Me.txtUpdateDate.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtUpdateDate.Location = New System.Drawing.Point(602, 121)
        Me.txtUpdateDate.Margin = New System.Windows.Forms.Padding(4)
        Me.txtUpdateDate.MaxLength = 50
        Me.txtUpdateDate.Name = "txtUpdateDate"
        Me.txtUpdateDate.ReadOnly = True
        Me.txtUpdateDate.Size = New System.Drawing.Size(132, 22)
        Me.txtUpdateDate.TabIndex = 6
        Me.txtUpdateDate.TabStop = False
        Me.txtUpdateDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'cboChosanen
        '
        Me.cboChosanen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChosanen.FormattingEnabled = True
        Me.cboChosanen.Location = New System.Drawing.Point(356, 121)
        Me.cboChosanen.Margin = New System.Windows.Forms.Padding(4)
        Me.cboChosanen.Name = "cboChosanen"
        Me.cboChosanen.Size = New System.Drawing.Size(132, 23)
        Me.cboChosanen.TabIndex = 6
        '
        'BRA2610F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.ClientSize = New System.Drawing.Size(1077, 262)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtUpdateDate)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnOutPut)
        Me.Controls.Add(Me.cboChosanen)
        Me.Margin = New System.Windows.Forms.Padding(7, 6, 7, 6)
        Me.Name = "BRA2610F"
        Me.Padding = New System.Windows.Forms.Padding(20, 18, 20, 18)
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.cboChosanen, 0)
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
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOutPut As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtUpdateDate As System.Windows.Forms.TextBox
    Friend WithEvents cboChosanen As System.Windows.Forms.ComboBox

End Class


