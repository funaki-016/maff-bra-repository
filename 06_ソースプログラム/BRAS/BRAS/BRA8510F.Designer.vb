<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA8510F
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
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.dgcSelect = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.dgcTodofuken = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcShityoson = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcKyuShityoson = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcNogyoSyuraku = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcTyosaKubun = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcKyakutaiNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcCensus = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtChosainShimei = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblKyakutaiNo = New System.Windows.Forms.Label()
        Me.lblTyosaKubun = New System.Windows.Forms.Label()
        Me.lblNogyoSyuraku = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblShityoson = New System.Windows.Forms.Label()
        Me.lblTodofuken = New System.Windows.Forms.Label()
        Me.lblCensus = New System.Windows.Forms.Label()
        Me.txtTodofuken = New System.Windows.Forms.TextBox()
        Me.btnInsert = New System.Windows.Forms.Button()
        Me.txtShikuchoson = New System.Windows.Forms.TextBox()
        Me.txtChosaku = New System.Windows.Forms.TextBox()
        Me.txtKyakutaiNo = New System.Windows.Forms.TextBox()
        Me.txtNogyoShuraku = New System.Windows.Forms.TextBox()
        Me.txtKyuShikuchoson = New System.Windows.Forms.TextBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
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
        Me.btnReturn.Location = New System.Drawing.Point(691, 471)
        Me.btnReturn.TabIndex = 10
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "専門調査員担当調査客体一覧"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(573, 471)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(90, 30)
        Me.btnDelete.TabIndex = 9
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dgvList)
        Me.GroupBox1.Location = New System.Drawing.Point(27, 106)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(754, 300)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "担当調査客体"
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
        Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgcSelect, Me.dgcTodofuken, Me.dgcShityoson, Me.dgcKyuShityoson, Me.dgcNogyoSyuraku, Me.dgcTyosaKubun, Me.dgcKyakutaiNo, Me.dgcCensus})
        Me.dgvList.Location = New System.Drawing.Point(17, 23)
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
        Me.dgvList.Size = New System.Drawing.Size(712, 258)
        Me.dgvList.TabIndex = 0
        '
        'dgcSelect
        '
        Me.dgcSelect.DataPropertyName = "チェックボックス"
        Me.dgcSelect.HeaderText = ""
        Me.dgcSelect.Name = "dgcSelect"
        Me.dgcSelect.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgcSelect.Width = 30
        '
        'dgcTodofuken
        '
        Me.dgcTodofuken.HeaderText = "都道府県"
        Me.dgcTodofuken.MaxInputLength = 2
        Me.dgcTodofuken.Name = "dgcTodofuken"
        Me.dgcTodofuken.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcTodofuken.Width = 60
        '
        'dgcShityoson
        '
        Me.dgcShityoson.HeaderText = "市区町村"
        Me.dgcShityoson.MaxInputLength = 3
        Me.dgcShityoson.Name = "dgcShityoson"
        Me.dgcShityoson.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcShityoson.Width = 60
        '
        'dgcKyuShityoson
        '
        Me.dgcKyuShityoson.HeaderText = "旧市区町村"
        Me.dgcKyuShityoson.MaxInputLength = 2
        Me.dgcKyuShityoson.Name = "dgcKyuShityoson"
        Me.dgcKyuShityoson.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcKyuShityoson.Width = 75
        '
        'dgcNogyoSyuraku
        '
        Me.dgcNogyoSyuraku.HeaderText = "農業集落"
        Me.dgcNogyoSyuraku.MaxInputLength = 3
        Me.dgcNogyoSyuraku.Name = "dgcNogyoSyuraku"
        Me.dgcNogyoSyuraku.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcNogyoSyuraku.Width = 60
        '
        'dgcTyosaKubun
        '
        Me.dgcTyosaKubun.HeaderText = "調査区"
        Me.dgcTyosaKubun.MaxInputLength = 3
        Me.dgcTyosaKubun.Name = "dgcTyosaKubun"
        Me.dgcTyosaKubun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcTyosaKubun.Width = 60
        '
        'dgcKyakutaiNo
        '
        Me.dgcKyakutaiNo.HeaderText = "客体番号"
        Me.dgcKyakutaiNo.MaxInputLength = 3
        Me.dgcKyakutaiNo.Name = "dgcKyakutaiNo"
        Me.dgcKyakutaiNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcKyakutaiNo.Width = 60
        '
        'dgcCensus
        '
        Me.dgcCensus.HeaderText = "センサス番号"
        Me.dgcCensus.MaxInputLength = 16
        Me.dgcCensus.Name = "dgcCensus"
        Me.dgcCensus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcCensus.Width = 120
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(34, 82)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 12)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "専門調査員"
        '
        'txtChosainShimei
        '
        Me.txtChosainShimei.Location = New System.Drawing.Point(108, 79)
        Me.txtChosainShimei.MaxLength = 50
        Me.txtChosainShimei.Name = "txtChosainShimei"
        Me.txtChosainShimei.ReadOnly = True
        Me.txtChosainShimei.Size = New System.Drawing.Size(328, 19)
        Me.txtChosainShimei.TabIndex = 6
        Me.txtChosainShimei.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lblKyakutaiNo)
        Me.GroupBox2.Controls.Add(Me.lblTyosaKubun)
        Me.GroupBox2.Controls.Add(Me.lblNogyoSyuraku)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.lblShityoson)
        Me.GroupBox2.Controls.Add(Me.lblTodofuken)
        Me.GroupBox2.Controls.Add(Me.lblCensus)
        Me.GroupBox2.Controls.Add(Me.txtTodofuken)
        Me.GroupBox2.Controls.Add(Me.btnInsert)
        Me.GroupBox2.Controls.Add(Me.txtShikuchoson)
        Me.GroupBox2.Controls.Add(Me.txtChosaku)
        Me.GroupBox2.Controls.Add(Me.txtKyakutaiNo)
        Me.GroupBox2.Controls.Add(Me.txtNogyoShuraku)
        Me.GroupBox2.Controls.Add(Me.txtKyuShikuchoson)
        Me.GroupBox2.Location = New System.Drawing.Point(27, 426)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(508, 89)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "新規客体追加"
        '
        'lblKyakutaiNo
        '
        Me.lblKyakutaiNo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKyakutaiNo.Location = New System.Drawing.Point(322, 37)
        Me.lblKyakutaiNo.Name = "lblKyakutaiNo"
        Me.lblKyakutaiNo.Size = New System.Drawing.Size(60, 18)
        Me.lblKyakutaiNo.TabIndex = 11
        Me.lblKyakutaiNo.Text = "客体番号"
        Me.lblKyakutaiNo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblTyosaKubun
        '
        Me.lblTyosaKubun.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTyosaKubun.Location = New System.Drawing.Point(263, 37)
        Me.lblTyosaKubun.Name = "lblTyosaKubun"
        Me.lblTyosaKubun.Size = New System.Drawing.Size(60, 18)
        Me.lblTyosaKubun.TabIndex = 9
        Me.lblTyosaKubun.Text = "調査区"
        Me.lblTyosaKubun.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblNogyoSyuraku
        '
        Me.lblNogyoSyuraku.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNogyoSyuraku.Location = New System.Drawing.Point(204, 37)
        Me.lblNogyoSyuraku.Name = "lblNogyoSyuraku"
        Me.lblNogyoSyuraku.Size = New System.Drawing.Size(60, 18)
        Me.lblNogyoSyuraku.TabIndex = 7
        Me.lblNogyoSyuraku.Text = "農業集落"
        Me.lblNogyoSyuraku.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Location = New System.Drawing.Point(135, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 18)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "旧市区町村"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblShityoson
        '
        Me.lblShityoson.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblShityoson.Location = New System.Drawing.Point(76, 37)
        Me.lblShityoson.Name = "lblShityoson"
        Me.lblShityoson.Size = New System.Drawing.Size(60, 18)
        Me.lblShityoson.TabIndex = 3
        Me.lblShityoson.Text = "市区町村"
        Me.lblShityoson.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblTodofuken
        '
        Me.lblTodofuken.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTodofuken.Location = New System.Drawing.Point(17, 37)
        Me.lblTodofuken.Name = "lblTodofuken"
        Me.lblTodofuken.Size = New System.Drawing.Size(60, 18)
        Me.lblTodofuken.TabIndex = 1
        Me.lblTodofuken.Text = "都道府県"
        Me.lblTodofuken.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblCensus
        '
        Me.lblCensus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCensus.Location = New System.Drawing.Point(17, 19)
        Me.lblCensus.Name = "lblCensus"
        Me.lblCensus.Size = New System.Drawing.Size(365, 18)
        Me.lblCensus.TabIndex = 0
        Me.lblCensus.Text = "センサス番号"
        Me.lblCensus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtTodofuken
        '
        Me.txtTodofuken.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtTodofuken.Location = New System.Drawing.Point(17, 55)
        Me.txtTodofuken.MaxLength = 2
        Me.txtTodofuken.Name = "txtTodofuken"
        Me.txtTodofuken.Size = New System.Drawing.Size(60, 19)
        Me.txtTodofuken.TabIndex = 2
        Me.txtTodofuken.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnInsert
        '
        Me.btnInsert.Location = New System.Drawing.Point(402, 45)
        Me.btnInsert.Name = "btnInsert"
        Me.btnInsert.Size = New System.Drawing.Size(90, 30)
        Me.btnInsert.TabIndex = 13
        Me.btnInsert.Text = "登録"
        Me.btnInsert.UseVisualStyleBackColor = True
        '
        'txtShikuchoson
        '
        Me.txtShikuchoson.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtShikuchoson.Location = New System.Drawing.Point(76, 55)
        Me.txtShikuchoson.MaxLength = 3
        Me.txtShikuchoson.Name = "txtShikuchoson"
        Me.txtShikuchoson.Size = New System.Drawing.Size(60, 19)
        Me.txtShikuchoson.TabIndex = 4
        Me.txtShikuchoson.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtChosaku
        '
        Me.txtChosaku.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtChosaku.Location = New System.Drawing.Point(263, 55)
        Me.txtChosaku.MaxLength = 3
        Me.txtChosaku.Name = "txtChosaku"
        Me.txtChosaku.Size = New System.Drawing.Size(60, 19)
        Me.txtChosaku.TabIndex = 10
        Me.txtChosaku.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtKyakutaiNo
        '
        Me.txtKyakutaiNo.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtKyakutaiNo.Location = New System.Drawing.Point(322, 55)
        Me.txtKyakutaiNo.MaxLength = 3
        Me.txtKyakutaiNo.Name = "txtKyakutaiNo"
        Me.txtKyakutaiNo.Size = New System.Drawing.Size(60, 19)
        Me.txtKyakutaiNo.TabIndex = 12
        Me.txtKyakutaiNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtNogyoShuraku
        '
        Me.txtNogyoShuraku.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtNogyoShuraku.Location = New System.Drawing.Point(204, 55)
        Me.txtNogyoShuraku.MaxLength = 3
        Me.txtNogyoShuraku.Name = "txtNogyoShuraku"
        Me.txtNogyoShuraku.Size = New System.Drawing.Size(60, 19)
        Me.txtNogyoShuraku.TabIndex = 8
        Me.txtNogyoShuraku.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtKyuShikuchoson
        '
        Me.txtKyuShikuchoson.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtKyuShikuchoson.Location = New System.Drawing.Point(135, 55)
        Me.txtKyuShikuchoson.MaxLength = 2
        Me.txtKyuShikuchoson.Name = "txtKyuShikuchoson"
        Me.txtKyuShikuchoson.Size = New System.Drawing.Size(70, 19)
        Me.txtKyuShikuchoson.TabIndex = 6
        Me.txtKyuShikuchoson.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'BRA8510F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 527)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtChosainShimei)
        Me.Controls.Add(Me.btnDelete)
        Me.Name = "BRA8510F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.btnDelete, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.txtChosainShimei, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.GroupBox2, 0)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Protected WithEvents dgvList As System.Windows.Forms.DataGridView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtChosainShimei As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnInsert As System.Windows.Forms.Button
    Friend WithEvents txtShikuchoson As System.Windows.Forms.TextBox
    Friend WithEvents txtKyuShikuchoson As System.Windows.Forms.TextBox
    Friend WithEvents txtNogyoShuraku As System.Windows.Forms.TextBox
    Friend WithEvents txtChosaku As System.Windows.Forms.TextBox
    Friend WithEvents txtTodofuken As System.Windows.Forms.TextBox
    Friend WithEvents txtKyakutaiNo As System.Windows.Forms.TextBox
    Friend WithEvents lblTyosaKubun As System.Windows.Forms.Label
    Friend WithEvents lblNogyoSyuraku As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblShityoson As System.Windows.Forms.Label
    Friend WithEvents lblTodofuken As System.Windows.Forms.Label
    Friend WithEvents lblCensus As System.Windows.Forms.Label
    Friend WithEvents lblKyakutaiNo As System.Windows.Forms.Label
    Friend WithEvents dgcSelect As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents dgcTodofuken As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcShityoson As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcKyuShityoson As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcNogyoSyuraku As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcTyosaKubun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcKyakutaiNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcCensus As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
