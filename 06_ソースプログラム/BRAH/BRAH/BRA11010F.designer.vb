<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BRA11010F
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.BtnOutput = New System.Windows.Forms.Button()
        Me.BtnInput = New System.Windows.Forms.Button()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.TxtFilename = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TxtChosanen = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.BtnCsvOutput = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblKoutei
        '
        Me.lblKoutei.Location = New System.Drawing.Point(664, 11)
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
        Me.btnReturn.Location = New System.Drawing.Point(704, 295)
        Me.btnReturn.Margin = New System.Windows.Forms.Padding(4)
        Me.btnReturn.TabIndex = 14
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "米在庫リンケージCSV出力"
        '
        'BtnOutput
        '
        Me.BtnOutput.Enabled = False
        Me.BtnOutput.Location = New System.Drawing.Point(114, 292)
        Me.BtnOutput.Name = "BtnOutput"
        Me.BtnOutput.Size = New System.Drawing.Size(90, 30)
        Me.BtnOutput.TabIndex = 12
        Me.BtnOutput.Text = "設定出力"
        Me.BtnOutput.UseVisualStyleBackColor = True
        '
        'BtnInput
        '
        Me.BtnInput.Location = New System.Drawing.Point(18, 292)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(90, 30)
        Me.BtnInput.TabIndex = 11
        Me.BtnInput.Text = "設定入力"
        Me.BtnInput.UseVisualStyleBackColor = True
        '
        'dgvList
        '
        Me.dgvList.AllowUserToAddRows = False
        Me.dgvList.AllowUserToDeleteRows = False
        Me.dgvList.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvList.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvList.ColumnHeadersHeight = 25
        Me.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvList.ColumnHeadersVisible = False
        Me.dgvList.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgvList.Location = New System.Drawing.Point(18, 136)
        Me.dgvList.Name = "dgvList"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvList.RowHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvList.RowHeadersVisible = False
        Me.dgvList.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.dgvList.RowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvList.RowTemplate.Height = 21
        Me.dgvList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvList.Size = New System.Drawing.Size(776, 140)
        Me.dgvList.TabIndex = 10
        '
        'TxtFilename
        '
        Me.TxtFilename.Enabled = False
        Me.TxtFilename.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.TxtFilename.Location = New System.Drawing.Point(114, 86)
        Me.TxtFilename.MaxLength = 100
        Me.TxtFilename.Name = "TxtFilename"
        Me.TxtFilename.Size = New System.Drawing.Size(680, 19)
        Me.TxtFilename.TabIndex = 6
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(16, 89)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(87, 12)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "ファイル識別名称"
        '
        'TxtChosanen
        '
        Me.TxtChosanen.Enabled = False
        Me.TxtChosanen.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TxtChosanen.Location = New System.Drawing.Point(114, 111)
        Me.TxtChosanen.MaxLength = 4
        Me.TxtChosanen.Name = "TxtChosanen"
        Me.TxtChosanen.Size = New System.Drawing.Size(42, 19)
        Me.TxtChosanen.TabIndex = 8
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(62, 114)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(41, 12)
        Me.Label15.TabIndex = 7
        Me.Label15.Text = "調査年"
        '
        'BtnCsvOutput
        '
        Me.BtnCsvOutput.Enabled = False
        Me.BtnCsvOutput.Location = New System.Drawing.Point(606, 295)
        Me.BtnCsvOutput.Name = "BtnCsvOutput"
        Me.BtnCsvOutput.Size = New System.Drawing.Size(90, 30)
        Me.BtnCsvOutput.TabIndex = 13
        Me.BtnCsvOutput.Text = "CSV出力"
        Me.BtnCsvOutput.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(162, 114)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(17, 12)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "年"
        '
        'BRA11010F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 341)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BtnCsvOutput)
        Me.Controls.Add(Me.TxtFilename)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.TxtChosanen)
        Me.Controls.Add(Me.dgvList)
        Me.Controls.Add(Me.BtnOutput)
        Me.Controls.Add(Me.BtnInput)
        Me.Controls.Add(Me.Label15)
        Me.Margin = New System.Windows.Forms.Padding(5)
        Me.Name = "BRA11010F"
        Me.Padding = New System.Windows.Forms.Padding(15, 14, 15, 14)
        Me.Text = "農業経営統計調査 - 本省工程 - 米在庫リンケージCSV出力"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.Label15, 0)
        Me.Controls.SetChildIndex(Me.BtnInput, 0)
        Me.Controls.SetChildIndex(Me.BtnOutput, 0)
        Me.Controls.SetChildIndex(Me.dgvList, 0)
        Me.Controls.SetChildIndex(Me.TxtChosanen, 0)
        Me.Controls.SetChildIndex(Me.Label7, 0)
        Me.Controls.SetChildIndex(Me.TxtFilename, 0)
        Me.Controls.SetChildIndex(Me.BtnCsvOutput, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnOutput As System.Windows.Forms.Button
    Friend WithEvents BtnInput As System.Windows.Forms.Button
    Protected WithEvents dgvList As DataGridView
    Friend WithEvents TxtFilename As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents TxtChosanen As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents BtnCsvOutput As Button
    Friend WithEvents Label1 As Label
End Class
