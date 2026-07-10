<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA1910F
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
        Me.btnImport = New System.Windows.Forms.Button()
        Me.txtMonthTo = New System.Windows.Forms.TextBox()
        Me.txtYearTo = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtMonthFrom = New System.Windows.Forms.TextBox()
        Me.txtYearFrom = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtImportName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtKotaiPath = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtIdoPath = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.btnRefKotai = New System.Windows.Forms.Button()
        Me.btnRefIdo = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblKoutei
        '
        Me.lblKoutei.Location = New System.Drawing.Point(635, 11)
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
        Me.btnReturn.Location = New System.Drawing.Point(668, 237)
        Me.btnReturn.TabIndex = 10
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "牛トレサデータ取込"
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(560, 237)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(90, 30)
        Me.btnImport.TabIndex = 8
        Me.btnImport.Text = "取込"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'txtMonthTo
        '
        Me.txtMonthTo.Location = New System.Drawing.Point(388, 90)
        Me.txtMonthTo.MaxLength = 2
        Me.txtMonthTo.Name = "txtMonthTo"
        Me.txtMonthTo.Size = New System.Drawing.Size(39, 19)
        Me.txtMonthTo.TabIndex = 29
        '
        'txtYearTo
        '
        Me.txtYearTo.Location = New System.Drawing.Point(316, 90)
        Me.txtYearTo.MaxLength = 4
        Me.txtYearTo.Name = "txtYearTo"
        Me.txtYearTo.Size = New System.Drawing.Size(39, 19)
        Me.txtYearTo.TabIndex = 28
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(361, 93)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(17, 12)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "年"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(433, 93)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(17, 12)
        Me.Label7.TabIndex = 26
        Me.Label7.Text = "月"
        '
        'txtMonthFrom
        '
        Me.txtMonthFrom.Location = New System.Drawing.Point(218, 90)
        Me.txtMonthFrom.MaxLength = 2
        Me.txtMonthFrom.Name = "txtMonthFrom"
        Me.txtMonthFrom.Size = New System.Drawing.Size(39, 19)
        Me.txtMonthFrom.TabIndex = 25
        '
        'txtYearFrom
        '
        Me.txtYearFrom.Location = New System.Drawing.Point(146, 90)
        Me.txtYearFrom.MaxLength = 4
        Me.txtYearFrom.Name = "txtYearFrom"
        Me.txtYearFrom.Size = New System.Drawing.Size(39, 19)
        Me.txtYearFrom.TabIndex = 24
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(191, 93)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(17, 12)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "年"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(263, 93)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(37, 12)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "月　～"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(98, 93)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(29, 12)
        Me.Label3.TabIndex = 21
        Me.Label3.Text = "期間"
        '
        'txtImportName
        '
        Me.txtImportName.Location = New System.Drawing.Point(185, 121)
        Me.txtImportName.MaxLength = 100
        Me.txtImportName.Name = "txtImportName"
        Me.txtImportName.Size = New System.Drawing.Size(477, 19)
        Me.txtImportName.TabIndex = 31
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(98, 124)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 12)
        Me.Label1.TabIndex = 30
        Me.Label1.Text = "取込データ名称"
        '
        'txtKotaiPath
        '
        Me.txtKotaiPath.Location = New System.Drawing.Point(167, 156)
        Me.txtKotaiPath.MaxLength = 2
        Me.txtKotaiPath.Name = "txtKotaiPath"
        Me.txtKotaiPath.Size = New System.Drawing.Size(399, 19)
        Me.txtKotaiPath.TabIndex = 33
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(98, 159)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 32
        Me.Label2.Text = "個体情報"
        '
        'txtIdoPath
        '
        Me.txtIdoPath.Location = New System.Drawing.Point(167, 192)
        Me.txtIdoPath.MaxLength = 2
        Me.txtIdoPath.Name = "txtIdoPath"
        Me.txtIdoPath.Size = New System.Drawing.Size(399, 19)
        Me.txtIdoPath.TabIndex = 35
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(98, 195)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(53, 12)
        Me.Label8.TabIndex = 34
        Me.Label8.Text = "異動履歴"
        '
        'btnRefKotai
        '
        Me.btnRefKotai.Location = New System.Drawing.Point(572, 150)
        Me.btnRefKotai.Name = "btnRefKotai"
        Me.btnRefKotai.Size = New System.Drawing.Size(90, 30)
        Me.btnRefKotai.TabIndex = 36
        Me.btnRefKotai.Text = "参照"
        Me.btnRefKotai.UseVisualStyleBackColor = True
        '
        'btnRefIdo
        '
        Me.btnRefIdo.Location = New System.Drawing.Point(572, 186)
        Me.btnRefIdo.Name = "btnRefIdo"
        Me.btnRefIdo.Size = New System.Drawing.Size(90, 30)
        Me.btnRefIdo.TabIndex = 37
        Me.btnRefIdo.Text = "参照"
        Me.btnRefIdo.UseVisualStyleBackColor = True
        '
        'BRA1910F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(778, 281)
        Me.Controls.Add(Me.btnRefIdo)
        Me.Controls.Add(Me.btnRefKotai)
        Me.Controls.Add(Me.txtIdoPath)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtKotaiPath)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtImportName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtMonthTo)
        Me.Controls.Add(Me.txtYearTo)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtMonthFrom)
        Me.Controls.Add(Me.txtYearFrom)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnImport)
        Me.Name = "BRA1910F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.btnImport, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.txtYearFrom, 0)
        Me.Controls.SetChildIndex(Me.txtMonthFrom, 0)
        Me.Controls.SetChildIndex(Me.Label7, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.txtYearTo, 0)
        Me.Controls.SetChildIndex(Me.txtMonthTo, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.txtImportName, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.txtKotaiPath, 0)
        Me.Controls.SetChildIndex(Me.Label8, 0)
        Me.Controls.SetChildIndex(Me.txtIdoPath, 0)
        Me.Controls.SetChildIndex(Me.btnRefKotai, 0)
        Me.Controls.SetChildIndex(Me.btnRefIdo, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents txtMonthTo As System.Windows.Forms.TextBox
    Friend WithEvents txtYearTo As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtMonthFrom As System.Windows.Forms.TextBox
    Friend WithEvents txtYearFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtImportName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtKotaiPath As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtIdoPath As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnRefKotai As System.Windows.Forms.Button
    Friend WithEvents btnRefIdo As System.Windows.Forms.Button

End Class
