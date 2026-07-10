<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA2120F
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.dgcSelect = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.dgcCensus = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.番号 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnAllCancel = New System.Windows.Forms.Button()
        Me.btnAllSelect = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnReturnhosei = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblKoutei
        '
        Me.lblKoutei.Location = New System.Drawing.Point(841, 11)
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
        Me.btnReturn.Location = New System.Drawing.Point(856, 444)
        Me.btnReturn.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnReturn.TabIndex = 7
        '
        'lblTitle
        '
        Me.lblTitle.Location = New System.Drawing.Point(178, 9)
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.Size = New System.Drawing.Size(808, 24)
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "補正異動情報一覧"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dgvList)
        Me.GroupBox1.Controls.Add(Me.btnAllCancel)
        Me.GroupBox1.Controls.Add(Me.btnAllSelect)
        Me.GroupBox1.Location = New System.Drawing.Point(27, 88)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(931, 340)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "対象客体"
        '
        'dgvList
        '
        Me.dgvList.AllowUserToAddRows = False
        Me.dgvList.AllowUserToDeleteRows = False
        Me.dgvList.AllowUserToResizeColumns = False
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
        Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgcSelect, Me.dgcCensus, Me.Column1, Me.Column2, Me.Column3, Me.Column4, Me.Column5, Me.Column6, Me.Column7, Me.番号})
        Me.dgvList.Location = New System.Drawing.Point(21, 24)
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
        Me.dgvList.Size = New System.Drawing.Size(888, 258)
        Me.dgvList.TabIndex = 9
        '
        'dgcSelect
        '
        Me.dgcSelect.DataPropertyName = "チェックボックス"
        Me.dgcSelect.HeaderText = ""
        Me.dgcSelect.Name = "dgcSelect"
        Me.dgcSelect.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgcSelect.Width = 30
        '
        'dgcCensus
        '
        Me.dgcCensus.HeaderText = "センサス番号"
        Me.dgcCensus.MaxInputLength = 16
        Me.dgcCensus.Name = "dgcCensus"
        Me.dgcCensus.ReadOnly = True
        Me.dgcCensus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcCensus.Width = 150
        '
        'Column1
        '
        Me.Column1.HeaderText = "個体識別番号"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column1.Width = 120
        '
        'Column2
        '
        Me.Column2.HeaderText = "転出異動年月日"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column2.Width = 130
        '
        'Column3
        '
        Me.Column3.HeaderText = "転出異動フラグ"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        Me.Column3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column3.Width = 110
        '
        'Column4
        '
        Me.Column4.HeaderText = "転出農家団体コード"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        Me.Column4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column4.Width = 140
        '
        'Column5
        '
        Me.Column5.HeaderText = "転入異動年月日"
        Me.Column5.Name = "Column5"
        Me.Column5.ReadOnly = True
        Me.Column5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column5.Width = 130
        '
        'Column6
        '
        Me.Column6.HeaderText = "転入異動フラグ"
        Me.Column6.Name = "Column6"
        Me.Column6.ReadOnly = True
        Me.Column6.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column6.Width = 110
        '
        'Column7
        '
        Me.Column7.HeaderText = "転入農家団体コード"
        Me.Column7.Name = "Column7"
        Me.Column7.ReadOnly = True
        Me.Column7.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column7.Width = 140
        '
        '番号
        '
        Me.番号.HeaderText = "番号"
        Me.番号.Name = "番号"
        Me.番号.ReadOnly = True
        Me.番号.Visible = False
        '
        'btnAllCancel
        '
        Me.btnAllCancel.Location = New System.Drawing.Point(117, 299)
        Me.btnAllCancel.Name = "btnAllCancel"
        Me.btnAllCancel.Size = New System.Drawing.Size(90, 30)
        Me.btnAllCancel.TabIndex = 11
        Me.btnAllCancel.Text = "全解除"
        Me.btnAllCancel.UseVisualStyleBackColor = True
        '
        'btnAllSelect
        '
        Me.btnAllSelect.Location = New System.Drawing.Point(21, 299)
        Me.btnAllSelect.Name = "btnAllSelect"
        Me.btnAllSelect.Size = New System.Drawing.Size(90, 30)
        Me.btnAllSelect.TabIndex = 10
        Me.btnAllSelect.Text = "全選択"
        Me.btnAllSelect.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(760, 444)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(90, 30)
        Me.btnDelete.TabIndex = 6
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnReturnhosei
        '
        Me.btnReturnhosei.Location = New System.Drawing.Point(856, 444)
        Me.btnReturnhosei.Name = "btnReturnhosei"
        Me.btnReturnhosei.Size = New System.Drawing.Size(90, 30)
        Me.btnReturnhosei.TabIndex = 8
        Me.btnReturnhosei.Text = "戻る"
        Me.btnReturnhosei.UseVisualStyleBackColor = True
        '
        'BRA2120F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(984, 491)
        Me.Controls.Add(Me.btnReturnhosei)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnDelete)
        Me.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5)
        Me.Name = "BRA2120F"
        Me.Padding = New System.Windows.Forms.Padding(15, 14, 15, 14)
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.btnDelete, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.btnReturnhosei, 0)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnAllCancel As System.Windows.Forms.Button
    Friend WithEvents btnAllSelect As System.Windows.Forms.Button
    Protected WithEvents dgvList As System.Windows.Forms.DataGridView
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnReturnhosei As System.Windows.Forms.Button
    Friend WithEvents dgcSelect As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents dgcCensus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 番号 As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
