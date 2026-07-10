<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA8310F
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.btnInsert = New System.Windows.Forms.Button()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnEditKyakutai = New System.Windows.Forms.Button()
        Me.btnOutput = New System.Windows.Forms.Button()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.chkSelect = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.txtUserID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtShimei = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtTourokubi = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtKyakutaiSu = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
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
        Me.btnReturn.Location = New System.Drawing.Point(691, 424)
        Me.btnReturn.TabIndex = 11
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "専門調査員一覧・帳票出力"
        '
        'btnInsert
        '
        Me.btnInsert.Location = New System.Drawing.Point(211, 424)
        Me.btnInsert.Name = "btnInsert"
        Me.btnInsert.Size = New System.Drawing.Size(90, 30)
        Me.btnInsert.TabIndex = 6
        Me.btnInsert.Text = "新規登録"
        Me.btnInsert.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(307, 424)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(90, 30)
        Me.btnEdit.TabIndex = 7
        Me.btnEdit.Text = "修正"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(403, 424)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(90, 30)
        Me.btnDelete.TabIndex = 8
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEditKyakutai
        '
        Me.btnEditKyakutai.Location = New System.Drawing.Point(499, 424)
        Me.btnEditKyakutai.Name = "btnEditKyakutai"
        Me.btnEditKyakutai.Size = New System.Drawing.Size(90, 30)
        Me.btnEditKyakutai.TabIndex = 9
        Me.btnEditKyakutai.Text = "調査客体編集"
        Me.btnEditKyakutai.UseVisualStyleBackColor = True
        '
        'btnOutput
        '
        Me.btnOutput.Location = New System.Drawing.Point(595, 424)
        Me.btnOutput.Name = "btnOutput"
        Me.btnOutput.Size = New System.Drawing.Size(90, 30)
        Me.btnOutput.TabIndex = 10
        Me.btnOutput.Text = "帳票全件出力"
        Me.btnOutput.UseVisualStyleBackColor = True
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
        Me.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.chkSelect, Me.txtUserID, Me.txtShimei, Me.txtTourokubi, Me.txtKyakutaiSu})
        Me.dgvList.Location = New System.Drawing.Point(11, 21)
        Me.dgvList.Name = "dgvList"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvList.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvList.RowHeadersVisible = False
        Me.dgvList.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvList.RowTemplate.Height = 21
        Me.dgvList.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.dgvList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvList.Size = New System.Drawing.Size(753, 295)
        Me.dgvList.TabIndex = 0
        '
        'chkSelect
        '
        Me.chkSelect.DataPropertyName = "チェックボックス"
        Me.chkSelect.HeaderText = ""
        Me.chkSelect.Name = "chkSelect"
        Me.chkSelect.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.chkSelect.Width = 30
        '
        'txtUserID
        '
        Me.txtUserID.HeaderText = "ユーザID"
        Me.txtUserID.MaxInputLength = 30
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.ReadOnly = True
        Me.txtUserID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.txtUserID.Width = 210
        '
        'txtShimei
        '
        Me.txtShimei.HeaderText = "氏名"
        Me.txtShimei.MaxInputLength = 30
        Me.txtShimei.Name = "txtShimei"
        Me.txtShimei.ReadOnly = True
        Me.txtShimei.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.txtShimei.Width = 293
        '
        'txtTourokubi
        '
        Me.txtTourokubi.HeaderText = "システム更新日時"
        Me.txtTourokubi.MaxInputLength = 20
        Me.txtTourokubi.Name = "txtTourokubi"
        Me.txtTourokubi.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'txtKyakutaiSu
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.txtKyakutaiSu.DefaultCellStyle = DataGridViewCellStyle2
        Me.txtKyakutaiSu.HeaderText = "担当客体数"
        Me.txtKyakutaiSu.MaxInputLength = 5
        Me.txtKyakutaiSu.Name = "txtKyakutaiSu"
        Me.txtKyakutaiSu.ReadOnly = True
        Me.txtKyakutaiSu.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dgvList)
        Me.GroupBox1.Location = New System.Drawing.Point(17, 79)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(779, 333)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "専門調査員"
        '
        'BRA8310F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 474)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnInsert)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEditKyakutai)
        Me.Controls.Add(Me.btnOutput)
        Me.Name = "BRA8310F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.btnOutput, 0)
        Me.Controls.SetChildIndex(Me.btnEditKyakutai, 0)
        Me.Controls.SetChildIndex(Me.btnDelete, 0)
        Me.Controls.SetChildIndex(Me.btnEdit, 0)
        Me.Controls.SetChildIndex(Me.btnInsert, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnInsert As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEditKyakutai As System.Windows.Forms.Button
    Friend WithEvents btnOutput As System.Windows.Forms.Button
    Protected WithEvents dgvList As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkSelect As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents txtUserID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txtShimei As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txtTourokubi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txtKyakutaiSu As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
