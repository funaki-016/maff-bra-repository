<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA2010F
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
        Me.btnOutput = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.txtMonthTo = New System.Windows.Forms.TextBox()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.dgcUpdate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcSelect = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.dgcYearFrom = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcMonthFrom = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcYearTo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcMonthTo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcImportName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcRegDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcKeika = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcKotaiCnt = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcIdoCnt = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcRirekiNum = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtYearTo = New System.Windows.Forms.TextBox()
        Me.btnShow = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cboKeika = New System.Windows.Forms.ComboBox()
        Me.txtMonthFrom = New System.Windows.Forms.TextBox()
        Me.btnAllCancel = New System.Windows.Forms.Button()
        Me.txtYearFrom = New System.Windows.Forms.TextBox()
        Me.btnAllSelect = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblKoutei
        '
        Me.lblKoutei.Location = New System.Drawing.Point(788, 11)
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
        Me.btnReturn.Location = New System.Drawing.Point(826, 541)
        Me.btnReturn.TabIndex = 7
        '
        'lblTitle
        '
        Me.lblTitle.Location = New System.Drawing.Point(151, 9)
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.Location = New System.Drawing.Point(151, 43)
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "牛トレサデータ管理"
        '
        'btnOutput
        '
        Me.btnOutput.Location = New System.Drawing.Point(576, 396)
        Me.btnOutput.Name = "btnOutput"
        Me.btnOutput.Size = New System.Drawing.Size(90, 30)
        Me.btnOutput.TabIndex = 6
        Me.btnOutput.Text = "出力"
        Me.btnOutput.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnDelete)
        Me.GroupBox1.Controls.Add(Me.btnUpdate)
        Me.GroupBox1.Controls.Add(Me.txtMonthTo)
        Me.GroupBox1.Controls.Add(Me.dgvList)
        Me.GroupBox1.Controls.Add(Me.btnOutput)
        Me.GroupBox1.Controls.Add(Me.txtYearTo)
        Me.GroupBox1.Controls.Add(Me.btnShow)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.cboKeika)
        Me.GroupBox1.Controls.Add(Me.txtMonthFrom)
        Me.GroupBox1.Controls.Add(Me.btnAllCancel)
        Me.GroupBox1.Controls.Add(Me.txtYearFrom)
        Me.GroupBox1.Controls.Add(Me.btnAllSelect)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Location = New System.Drawing.Point(27, 88)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(880, 441)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "対象客体"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(768, 396)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(90, 30)
        Me.btnDelete.TabIndex = 22
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(672, 396)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(90, 30)
        Me.btnUpdate.TabIndex = 21
        Me.btnUpdate.Text = "修正"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'txtMonthTo
        '
        Me.txtMonthTo.Location = New System.Drawing.Point(299, 21)
        Me.txtMonthTo.MaxLength = 2
        Me.txtMonthTo.Name = "txtMonthTo"
        Me.txtMonthTo.Size = New System.Drawing.Size(39, 19)
        Me.txtMonthTo.TabIndex = 20
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
        Me.dgvList.ColumnHeadersHeight = 35
        Me.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgcUpdate, Me.dgcSelect, Me.dgcYearFrom, Me.dgcMonthFrom, Me.dgcYearTo, Me.dgcMonthTo, Me.dgcImportName, Me.dgcRegDate, Me.dgcKeika, Me.dgcKotaiCnt, Me.dgcIdoCnt, Me.dgcRirekiNum})
        Me.dgvList.Location = New System.Drawing.Point(21, 125)
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
        Me.dgvList.Size = New System.Drawing.Size(837, 258)
        Me.dgvList.TabIndex = 9
        '
        'dgcUpdate
        '
        Me.dgcUpdate.HeaderText = ""
        Me.dgcUpdate.Name = "dgcUpdate"
        Me.dgcUpdate.Width = 20
        '
        'dgcSelect
        '
        Me.dgcSelect.DataPropertyName = "チェックボックス"
        Me.dgcSelect.HeaderText = ""
        Me.dgcSelect.Name = "dgcSelect"
        Me.dgcSelect.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgcSelect.Width = 30
        '
        'dgcYearFrom
        '
        Me.dgcYearFrom.HeaderText = " 期間   開始年"
        Me.dgcYearFrom.MaxInputLength = 4
        Me.dgcYearFrom.Name = "dgcYearFrom"
        Me.dgcYearFrom.Width = 60
        '
        'dgcMonthFrom
        '
        Me.dgcMonthFrom.HeaderText = " 期間   開始月"
        Me.dgcMonthFrom.Name = "dgcMonthFrom"
        Me.dgcMonthFrom.Width = 60
        '
        'dgcYearTo
        '
        Me.dgcYearTo.HeaderText = " 期間   終了年"
        Me.dgcYearTo.Name = "dgcYearTo"
        Me.dgcYearTo.Width = 60
        '
        'dgcMonthTo
        '
        Me.dgcMonthTo.HeaderText = " 期間   終了月"
        Me.dgcMonthTo.Name = "dgcMonthTo"
        Me.dgcMonthTo.Width = 60
        '
        'dgcImportName
        '
        Me.dgcImportName.HeaderText = "取込データ名称"
        Me.dgcImportName.Name = "dgcImportName"
        Me.dgcImportName.Width = 200
        '
        'dgcRegDate
        '
        Me.dgcRegDate.HeaderText = "データ登録日"
        Me.dgcRegDate.Name = "dgcRegDate"
        '
        'dgcKeika
        '
        Me.dgcKeika.HeaderText = "データ登録以降2年経過情報"
        Me.dgcKeika.Name = "dgcKeika"
        Me.dgcKeika.Width = 90
        '
        'dgcKotaiCnt
        '
        Me.dgcKotaiCnt.HeaderText = "個体情報レコード数"
        Me.dgcKotaiCnt.Name = "dgcKotaiCnt"
        Me.dgcKotaiCnt.Width = 60
        '
        'dgcIdoCnt
        '
        Me.dgcIdoCnt.HeaderText = "異動履歴レコード数"
        Me.dgcIdoCnt.Name = "dgcIdoCnt"
        Me.dgcIdoCnt.Width = 60
        '
        'dgcRirekiNum
        '
        Me.dgcRirekiNum.HeaderText = "履歴番号"
        Me.dgcRirekiNum.Name = "dgcRirekiNum"
        Me.dgcRirekiNum.Visible = False
        '
        'txtYearTo
        '
        Me.txtYearTo.Location = New System.Drawing.Point(227, 21)
        Me.txtYearTo.MaxLength = 4
        Me.txtYearTo.Name = "txtYearTo"
        Me.txtYearTo.Size = New System.Drawing.Size(39, 19)
        Me.txtYearTo.TabIndex = 19
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(21, 79)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(90, 30)
        Me.btnShow.TabIndex = 8
        Me.btnShow.Text = "表示"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(272, 24)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(17, 12)
        Me.Label6.TabIndex = 18
        Me.Label6.Text = "年"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(22, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(147, 12)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "データ登録以降2年経過情報"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(344, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(17, 12)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "月"
        '
        'cboKeika
        '
        Me.cboKeika.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKeika.FormattingEnabled = True
        Me.cboKeika.Items.AddRange(New Object() {"", "2年以上経過", "－"})
        Me.cboKeika.Location = New System.Drawing.Point(182, 52)
        Me.cboKeika.Name = "cboKeika"
        Me.cboKeika.Size = New System.Drawing.Size(107, 20)
        Me.cboKeika.TabIndex = 5
        '
        'txtMonthFrom
        '
        Me.txtMonthFrom.Location = New System.Drawing.Point(129, 21)
        Me.txtMonthFrom.MaxLength = 2
        Me.txtMonthFrom.Name = "txtMonthFrom"
        Me.txtMonthFrom.Size = New System.Drawing.Size(39, 19)
        Me.txtMonthFrom.TabIndex = 16
        '
        'btnAllCancel
        '
        Me.btnAllCancel.Location = New System.Drawing.Point(117, 396)
        Me.btnAllCancel.Name = "btnAllCancel"
        Me.btnAllCancel.Size = New System.Drawing.Size(90, 30)
        Me.btnAllCancel.TabIndex = 11
        Me.btnAllCancel.Text = "全解除"
        Me.btnAllCancel.UseVisualStyleBackColor = True
        '
        'txtYearFrom
        '
        Me.txtYearFrom.Location = New System.Drawing.Point(57, 21)
        Me.txtYearFrom.MaxLength = 4
        Me.txtYearFrom.Name = "txtYearFrom"
        Me.txtYearFrom.Size = New System.Drawing.Size(39, 19)
        Me.txtYearFrom.TabIndex = 15
        '
        'btnAllSelect
        '
        Me.btnAllSelect.Location = New System.Drawing.Point(21, 396)
        Me.btnAllSelect.Name = "btnAllSelect"
        Me.btnAllSelect.Size = New System.Drawing.Size(90, 30)
        Me.btnAllSelect.TabIndex = 10
        Me.btnAllSelect.Text = "全選択"
        Me.btnAllSelect.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(102, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(17, 12)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "年"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(22, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(29, 12)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "期間"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(174, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(37, 12)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "月　～"
        '
        'BRA2010F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(932, 580)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "BRA2010F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnOutput As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnAllCancel As System.Windows.Forms.Button
    Friend WithEvents btnAllSelect As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Protected WithEvents cboKeika As System.Windows.Forms.ComboBox
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Protected WithEvents dgvList As System.Windows.Forms.DataGridView
    Friend WithEvents txtMonthTo As System.Windows.Forms.TextBox
    Friend WithEvents txtYearTo As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtMonthFrom As System.Windows.Forms.TextBox
    Friend WithEvents txtYearFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents dgcUpdate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcSelect As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents dgcYearFrom As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcMonthFrom As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcYearTo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcMonthTo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcImportName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcRegDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcKeika As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcKotaiCnt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcIdoCnt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcRirekiNum As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
