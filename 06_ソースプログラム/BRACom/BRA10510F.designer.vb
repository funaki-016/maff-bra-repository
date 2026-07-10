<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA10510F
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
        Me.BtnSendRecv = New System.Windows.Forms.Button()
        Me.BtnShow = New System.Windows.Forms.Button()
        Me.CboKyoku = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CboKyoten = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.CboCensusNen = New System.Windows.Forms.ComboBox()
        Me.BtnAllRelease = New System.Windows.Forms.Button()
        Me.BtnAllSelect = New System.Windows.Forms.Button()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.COL_0 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.COL_1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblKoutei
        '
        Me.lblKoutei.Location = New System.Drawing.Point(703, 11)
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
        Me.btnReturn.Location = New System.Drawing.Point(740, 411)
        Me.btnReturn.TabIndex = 16
        '
        'lblTitle
        '
        Me.lblTitle.Location = New System.Drawing.Point(119, 9)
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.Location = New System.Drawing.Point(119, 43)
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "標本リスト送受信（送受信先）"
        '
        'BtnSendRecv
        '
        Me.BtnSendRecv.Location = New System.Drawing.Point(635, 411)
        Me.BtnSendRecv.Name = "BtnSendRecv"
        Me.BtnSendRecv.Size = New System.Drawing.Size(90, 30)
        Me.BtnSendRecv.TabIndex = 15
        Me.BtnSendRecv.Text = "送受信"
        Me.BtnSendRecv.UseVisualStyleBackColor = True
        '
        'BtnShow
        '
        Me.BtnShow.Location = New System.Drawing.Point(705, 94)
        Me.BtnShow.Name = "BtnShow"
        Me.BtnShow.Size = New System.Drawing.Size(90, 30)
        Me.BtnShow.TabIndex = 11
        Me.BtnShow.Text = "表示"
        Me.BtnShow.UseVisualStyleBackColor = True
        '
        'CboKyoku
        '
        Me.CboKyoku.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboKyoku.FormattingEnabled = True
        Me.CboKyoku.Location = New System.Drawing.Point(276, 100)
        Me.CboKyoku.Name = "CboKyoku"
        Me.CboKyoku.Size = New System.Drawing.Size(110, 20)
        Me.CboKyoku.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(403, 103)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 12)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "実査設置拠点"
        '
        'CboKyoten
        '
        Me.CboKyoten.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboKyoten.FormattingEnabled = True
        Me.CboKyoten.Location = New System.Drawing.Point(488, 100)
        Me.CboKyoten.Name = "CboKyoten"
        Me.CboKyoten.Size = New System.Drawing.Size(72, 20)
        Me.CboKyoten.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(229, 103)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 12)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "農政局"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(51, 103)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(79, 12)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "センサス実施年"
        '
        'CboCensusNen
        '
        Me.CboCensusNen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboCensusNen.FormattingEnabled = True
        Me.CboCensusNen.Items.AddRange(New Object() {"2020"})
        Me.CboCensusNen.Location = New System.Drawing.Point(136, 100)
        Me.CboCensusNen.Name = "CboCensusNen"
        Me.CboCensusNen.Size = New System.Drawing.Size(72, 20)
        Me.CboCensusNen.TabIndex = 6
        '
        'BtnAllRelease
        '
        Me.BtnAllRelease.Location = New System.Drawing.Point(146, 411)
        Me.BtnAllRelease.Name = "BtnAllRelease"
        Me.BtnAllRelease.Size = New System.Drawing.Size(90, 30)
        Me.BtnAllRelease.TabIndex = 14
        Me.BtnAllRelease.Text = "全解除"
        Me.BtnAllRelease.UseVisualStyleBackColor = True
        '
        'BtnAllSelect
        '
        Me.BtnAllSelect.Location = New System.Drawing.Point(50, 411)
        Me.BtnAllSelect.Name = "BtnAllSelect"
        Me.BtnAllSelect.Size = New System.Drawing.Size(90, 30)
        Me.BtnAllSelect.TabIndex = 13
        Me.BtnAllSelect.Text = "全選択"
        Me.BtnAllSelect.UseVisualStyleBackColor = True
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
        Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.COL_0, Me.COL_1, Me.COL_2, Me.COL_3, Me.COL_4, Me.COL_5, Me.COL_6, Me.COL_7, Me.COL_8, Me.COL_9, Me.COL_10, Me.COL_11, Me.COL_12, Me.COL_13, Me.COL_14, Me.COL_15})
        Me.dgvList.Location = New System.Drawing.Point(50, 135)
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
        Me.dgvList.Size = New System.Drawing.Size(745, 258)
        Me.dgvList.TabIndex = 12
        '
        'COL_0
        '
        Me.COL_0.DataPropertyName = "チェックボックス"
        Me.COL_0.HeaderText = ""
        Me.COL_0.Name = "COL_0"
        Me.COL_0.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.COL_0.Width = 30
        '
        'COL_1
        '
        Me.COL_1.HeaderText = "農政局"
        Me.COL_1.MaxInputLength = 2
        Me.COL_1.Name = "COL_1"
        Me.COL_1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COL_1.Width = 95
        '
        'COL_2
        '
        Me.COL_2.HeaderText = "実査設定拠点"
        Me.COL_2.MaxInputLength = 3
        Me.COL_2.Name = "COL_2"
        Me.COL_2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COL_2.Width = 90
        '
        'COL_3
        '
        Me.COL_3.HeaderText = "経営形態区分"
        Me.COL_3.MaxInputLength = 2
        Me.COL_3.Name = "COL_3"
        Me.COL_3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COL_3.Width = 95
        '
        'COL_4
        '
        Me.COL_4.HeaderText = "営農類型・生産費区分（田畑区分）"
        Me.COL_4.MaxInputLength = 3
        Me.COL_4.Name = "COL_4"
        Me.COL_4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COL_4.Width = 190
        '
        'COL_5
        '
        Me.COL_5.HeaderText = "更新／送信日時"
        Me.COL_5.MaxInputLength = 16
        Me.COL_5.Name = "COL_5"
        Me.COL_5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COL_5.Width = 110
        '
        'COL_6
        '
        Me.COL_6.HeaderText = "送受信日時"
        Me.COL_6.MaxInputLength = 16
        Me.COL_6.Name = "COL_6"
        Me.COL_6.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COL_6.Width = 110
        '
        'COL_7
        '
        Me.COL_7.HeaderText = "農政局（非表示）"
        Me.COL_7.Name = "COL_7"
        Me.COL_7.Visible = False
        '
        'COL_8
        '
        Me.COL_8.HeaderText = "都道府県（非表示）"
        Me.COL_8.Name = "COL_8"
        Me.COL_8.Visible = False
        '
        'COL_9
        '
        Me.COL_9.HeaderText = "実査設置拠点（非表示）"
        Me.COL_9.Name = "COL_9"
        Me.COL_9.Visible = False
        '
        'COL_10
        '
        Me.COL_10.HeaderText = "経営形態区分（非表示）"
        Me.COL_10.Name = "COL_10"
        Me.COL_10.Visible = False
        '
        'COL_11
        '
        Me.COL_11.HeaderText = "営農類型区分（非表示）"
        Me.COL_11.Name = "COL_11"
        Me.COL_11.Visible = False
        '
        'COL_12
        '
        Me.COL_12.HeaderText = "生産費区分（非表示）"
        Me.COL_12.Name = "COL_12"
        Me.COL_12.Visible = False
        '
        'COL_13
        '
        Me.COL_13.HeaderText = "田畑区分（非表示）"
        Me.COL_13.Name = "COL_13"
        Me.COL_13.Visible = False
        '
        'COL_14
        '
        Me.COL_14.HeaderText = "送受信可否（非表示）"
        Me.COL_14.Name = "COL_14"
        Me.COL_14.Visible = False
        '
        'COL_15
        '
        Me.COL_15.HeaderText = "送受信日時チェック用（非表示）"
        Me.COL_15.Name = "COL_15"
        Me.COL_15.Visible = False
        '
        'BRA10510F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(844, 461)
        Me.Controls.Add(Me.BtnShow)
        Me.Controls.Add(Me.CboKyoku)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CboKyoten)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.CboCensusNen)
        Me.Controls.Add(Me.BtnAllRelease)
        Me.Controls.Add(Me.BtnAllSelect)
        Me.Controls.Add(Me.dgvList)
        Me.Controls.Add(Me.BtnSendRecv)
        Me.Name = "BRA10510F"
        Me.Text = "農業経営統計調査 - 実査設置拠点工程 - 標本リスト送受信（送受信先）"
        Me.Controls.SetChildIndex(Me.BtnSendRecv, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.dgvList, 0)
        Me.Controls.SetChildIndex(Me.BtnAllSelect, 0)
        Me.Controls.SetChildIndex(Me.BtnAllRelease, 0)
        Me.Controls.SetChildIndex(Me.CboCensusNen, 0)
        Me.Controls.SetChildIndex(Me.Label10, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.CboKyoten, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.CboKyoku, 0)
        Me.Controls.SetChildIndex(Me.BtnShow, 0)
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnSendRecv As System.Windows.Forms.Button
    Friend WithEvents BtnShow As Button
    Protected WithEvents CboKyoku As ComboBox
    Friend WithEvents Label2 As Label
    Protected WithEvents CboKyoten As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label10 As Label
    Protected WithEvents CboCensusNen As ComboBox
    Friend WithEvents BtnAllRelease As Button
    Friend WithEvents BtnAllSelect As Button
    Protected WithEvents dgvList As DataGridView
    Friend WithEvents COL_0 As DataGridViewCheckBoxColumn
    Friend WithEvents COL_1 As DataGridViewTextBoxColumn
    Friend WithEvents COL_2 As DataGridViewTextBoxColumn
    Friend WithEvents COL_3 As DataGridViewTextBoxColumn
    Friend WithEvents COL_4 As DataGridViewTextBoxColumn
    Friend WithEvents COL_5 As DataGridViewTextBoxColumn
    Friend WithEvents COL_6 As DataGridViewTextBoxColumn
    Friend WithEvents COL_7 As DataGridViewTextBoxColumn
    Friend WithEvents COL_8 As DataGridViewTextBoxColumn
    Friend WithEvents COL_9 As DataGridViewTextBoxColumn
    Friend WithEvents COL_10 As DataGridViewTextBoxColumn
    Friend WithEvents COL_11 As DataGridViewTextBoxColumn
    Friend WithEvents COL_12 As DataGridViewTextBoxColumn
    Friend WithEvents COL_13 As DataGridViewTextBoxColumn
    Friend WithEvents COL_14 As DataGridViewTextBoxColumn
    Friend WithEvents COL_15 As DataGridViewTextBoxColumn
End Class
