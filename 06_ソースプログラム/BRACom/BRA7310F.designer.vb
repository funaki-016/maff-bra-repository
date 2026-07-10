<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA7310F
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.btnReduce = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.DataGridViewCheckBoxColumn1 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.dgcShukeiNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcShukeiName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dcgUpdateDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcSendDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcShukeiConditions = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.局 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.事務所 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.拠点 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cboChosaNen = New System.Windows.Forms.ComboBox()
        Me.btnShow = New System.Windows.Forms.Button()
        Me.cboEinouKeieitai = New System.Windows.Forms.ComboBox()
        Me.lblEinouKeieitai = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblKoutei
        '
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
        Me.btnReturn.Location = New System.Drawing.Point(704, 530)
        Me.btnReturn.TabIndex = 7
        '
        'lblTitle
        '
        Me.lblTitle.Location = New System.Drawing.Point(101, 9)
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.Location = New System.Drawing.Point(101, 43)
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "集計結果表還元"
        '
        'btnReduce
        '
        Me.btnReduce.Location = New System.Drawing.Point(608, 530)
        Me.btnReduce.Name = "btnReduce"
        Me.btnReduce.Size = New System.Drawing.Size(90, 30)
        Me.btnReduce.TabIndex = 6
        Me.btnReduce.Text = "還元"
        Me.btnReduce.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dgvList)
        Me.GroupBox1.Controls.Add(Me.cboChosaNen)
        Me.GroupBox1.Controls.Add(Me.btnShow)
        Me.GroupBox1.Controls.Add(Me.cboEinouKeieitai)
        Me.GroupBox1.Controls.Add(Me.lblEinouKeieitai)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Location = New System.Drawing.Point(27, 88)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(754, 426)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        '
        'dgvList
        '
        Me.dgvList.AllowUserToAddRows = False
        Me.dgvList.AllowUserToDeleteRows = False
        Me.dgvList.AllowUserToResizeColumns = False
        Me.dgvList.AllowUserToResizeRows = False
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvList.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvList.ColumnHeadersHeight = 25
        Me.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewCheckBoxColumn1, Me.dgcShukeiNo, Me.dgcShukeiName, Me.dcgUpdateDate, Me.dgcSendDate, Me.dgcShukeiConditions, Me.局, Me.事務所, Me.拠点})
        Me.dgvList.Location = New System.Drawing.Point(21, 49)
        Me.dgvList.Name = "dgvList"
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvList.RowHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.dgvList.RowHeadersVisible = False
        Me.dgvList.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.dgvList.RowsDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvList.RowTemplate.Height = 21
        Me.dgvList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvList.Size = New System.Drawing.Size(712, 359)
        Me.dgvList.TabIndex = 5
        '
        'DataGridViewCheckBoxColumn1
        '
        Me.DataGridViewCheckBoxColumn1.DataPropertyName = "チェックボックス"
        Me.DataGridViewCheckBoxColumn1.HeaderText = ""
        Me.DataGridViewCheckBoxColumn1.Name = "DataGridViewCheckBoxColumn1"
        Me.DataGridViewCheckBoxColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewCheckBoxColumn1.Width = 30
        '
        'dgcShukeiNo
        '
        Me.dgcShukeiNo.HeaderText = "集計番号"
        Me.dgcShukeiNo.MaxInputLength = 6
        Me.dgcShukeiNo.Name = "dgcShukeiNo"
        Me.dgcShukeiNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'dgcShukeiName
        '
        Me.dgcShukeiName.HeaderText = "集計名称"
        Me.dgcShukeiName.MaxInputLength = 0
        Me.dgcShukeiName.Name = "dgcShukeiName"
        Me.dgcShukeiName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcShukeiName.Width = 300
        '
        'dcgUpdateDate
        '
        Me.dcgUpdateDate.HeaderText = "更新日時"
        Me.dcgUpdateDate.MaxInputLength = 8
        Me.dcgUpdateDate.Name = "dcgUpdateDate"
        Me.dcgUpdateDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'dgcSendDate
        '
        Me.dgcSendDate.HeaderText = "還元日時"
        Me.dgcSendDate.MaxInputLength = 8
        Me.dgcSendDate.Name = "dgcSendDate"
        Me.dgcSendDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'dgcShukeiConditions
        '
        Me.dgcShukeiConditions.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.dgcShukeiConditions.HeaderText = "集計条件"
        Me.dgcShukeiConditions.Name = "dgcShukeiConditions"
        Me.dgcShukeiConditions.Width = 78
        '
        '局
        '
        Me.局.HeaderText = "局"
        Me.局.Name = "局"
        Me.局.Visible = False
        '
        '事務所
        '
        Me.事務所.HeaderText = "事務所"
        Me.事務所.Name = "事務所"
        Me.事務所.Visible = False
        '
        '拠点
        '
        Me.拠点.HeaderText = "拠点"
        Me.拠点.Name = "拠点"
        Me.拠点.Visible = False
        '
        'cboChosaNen
        '
        Me.cboChosaNen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChosaNen.FormattingEnabled = True
        Me.cboChosaNen.Location = New System.Drawing.Point(107, 18)
        Me.cboChosaNen.Name = "cboChosaNen"
        Me.cboChosaNen.Size = New System.Drawing.Size(72, 20)
        Me.cboChosaNen.TabIndex = 1
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(643, 12)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(90, 30)
        Me.btnShow.TabIndex = 4
        Me.btnShow.Text = "表示"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'cboEinouKeieitai
        '
        Me.cboEinouKeieitai.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEinouKeieitai.FormattingEnabled = True
        Me.cboEinouKeieitai.Location = New System.Drawing.Point(332, 18)
        Me.cboEinouKeieitai.Name = "cboEinouKeieitai"
        Me.cboEinouKeieitai.Size = New System.Drawing.Size(90, 20)
        Me.cboEinouKeieitai.TabIndex = 3
        '
        'lblEinouKeieitai
        '
        Me.lblEinouKeieitai.AutoSize = True
        Me.lblEinouKeieitai.Location = New System.Drawing.Point(225, 21)
        Me.lblEinouKeieitai.Name = "lblEinouKeieitai"
        Me.lblEinouKeieitai.Size = New System.Drawing.Size(89, 12)
        Me.lblEinouKeieitai.TabIndex = 2
        Me.lblEinouKeieitai.Text = "営農経営体区分"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(22, 21)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 12)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "調査年（産）"
        '
        'BRA7310F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 576)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnReduce)
        Me.Name = "BRA7310F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.btnReduce, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnReduce As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblEinouKeieitai As System.Windows.Forms.Label
    Protected WithEvents cboEinouKeieitai As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Protected WithEvents cboChosaNen As System.Windows.Forms.ComboBox
    Protected WithEvents dgvList As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewCheckBoxColumn1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents dgcShukeiNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcShukeiName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dcgUpdateDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcSendDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcShukeiConditions As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 局 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 事務所 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 拠点 As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
