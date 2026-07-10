<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA4010F
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.btnModify = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cboTaishakuTaishohyo = New System.Windows.Forms.ComboBox()
        Me.lblTaishakuTaishohyo = New System.Windows.Forms.Label()
        Me.cboKessokuchiHokan = New System.Windows.Forms.ComboBox()
        Me.lblKessokuchiHokan = New System.Windows.Forms.Label()
        Me.cboEinouRuikei = New System.Windows.Forms.ComboBox()
        Me.lblEinouRuikei = New System.Windows.Forms.Label()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.dgcSelect = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.dgcTodofuken = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcShityoson = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcKyuShityoson = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcNogyoSyuraku = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcTyosaKubun = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcKyakutaiNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcCensus = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcUpdateDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.局 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.事務所 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.拠点 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnShow = New System.Windows.Forms.Button()
        Me.cboKyoku = New System.Windows.Forms.ComboBox()
        Me.lblKyoten = New System.Windows.Forms.Label()
        Me.cboKyoten = New System.Windows.Forms.ComboBox()
        Me.lblKyoku = New System.Windows.Forms.Label()
        Me.lblChosaNen = New System.Windows.Forms.Label()
        Me.cboChosaNen = New System.Windows.Forms.ComboBox()
        Me.btnAllCancel = New System.Windows.Forms.Button()
        Me.btnAllSelect = New System.Windows.Forms.Button()
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
        Me.lblSyori.Text = "個別結果表項目指定修正"
        '
        'btnModify
        '
        Me.btnModify.Location = New System.Drawing.Point(608, 530)
        Me.btnModify.Name = "btnModify"
        Me.btnModify.Size = New System.Drawing.Size(90, 30)
        Me.btnModify.TabIndex = 6
        Me.btnModify.Text = "修正"
        Me.btnModify.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnAllCancel)
        Me.GroupBox1.Controls.Add(Me.btnAllSelect)
        Me.GroupBox1.Controls.Add(Me.cboTaishakuTaishohyo)
        Me.GroupBox1.Controls.Add(Me.lblTaishakuTaishohyo)
        Me.GroupBox1.Controls.Add(Me.cboKessokuchiHokan)
        Me.GroupBox1.Controls.Add(Me.lblKessokuchiHokan)
        Me.GroupBox1.Controls.Add(Me.cboEinouRuikei)
        Me.GroupBox1.Controls.Add(Me.lblEinouRuikei)
        Me.GroupBox1.Controls.Add(Me.dgvList)
        Me.GroupBox1.Controls.Add(Me.btnShow)
        Me.GroupBox1.Controls.Add(Me.cboKyoku)
        Me.GroupBox1.Controls.Add(Me.lblKyoten)
        Me.GroupBox1.Controls.Add(Me.cboKyoten)
        Me.GroupBox1.Controls.Add(Me.lblKyoku)
        Me.GroupBox1.Controls.Add(Me.lblChosaNen)
        Me.GroupBox1.Controls.Add(Me.cboChosaNen)
        Me.GroupBox1.Location = New System.Drawing.Point(27, 88)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(754, 426)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "対象客体"
        '
        'cboTaishakuTaishohyo
        '
        Me.cboTaishakuTaishohyo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTaishakuTaishohyo.FormattingEnabled = True
        Me.cboTaishakuTaishohyo.Location = New System.Drawing.Point(451, 49)
        Me.cboTaishakuTaishohyo.Name = "cboTaishakuTaishohyo"
        Me.cboTaishakuTaishohyo.Size = New System.Drawing.Size(90, 20)
        Me.cboTaishakuTaishohyo.TabIndex = 11
        '
        'lblTaishakuTaishohyo
        '
        Me.lblTaishakuTaishohyo.AutoSize = True
        Me.lblTaishakuTaishohyo.Location = New System.Drawing.Point(380, 52)
        Me.lblTaishakuTaishohyo.Name = "lblTaishakuTaishohyo"
        Me.lblTaishakuTaishohyo.Size = New System.Drawing.Size(65, 12)
        Me.lblTaishakuTaishohyo.TabIndex = 10
        Me.lblTaishakuTaishohyo.Text = "貸借対照表"
        '
        'cboKessokuchiHokan
        '
        Me.cboKessokuchiHokan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKessokuchiHokan.FormattingEnabled = True
        Me.cboKessokuchiHokan.Location = New System.Drawing.Point(274, 49)
        Me.cboKessokuchiHokan.Name = "cboKessokuchiHokan"
        Me.cboKessokuchiHokan.Size = New System.Drawing.Size(90, 20)
        Me.cboKessokuchiHokan.TabIndex = 9
        '
        'lblKessokuchiHokan
        '
        Me.lblKessokuchiHokan.AutoSize = True
        Me.lblKessokuchiHokan.Location = New System.Drawing.Point(196, 52)
        Me.lblKessokuchiHokan.Name = "lblKessokuchiHokan"
        Me.lblKessokuchiHokan.Size = New System.Drawing.Size(65, 12)
        Me.lblKessokuchiHokan.TabIndex = 8
        Me.lblKessokuchiHokan.Text = "欠測値補完"
        '
        'cboEinouRuikei
        '
        Me.cboEinouRuikei.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEinouRuikei.FormattingEnabled = True
        Me.cboEinouRuikei.Location = New System.Drawing.Point(254, 18)
        Me.cboEinouRuikei.Name = "cboEinouRuikei"
        Me.cboEinouRuikei.Size = New System.Drawing.Size(110, 20)
        Me.cboEinouRuikei.TabIndex = 7
        '
        'lblEinouRuikei
        '
        Me.lblEinouRuikei.AutoSize = True
        Me.lblEinouRuikei.Location = New System.Drawing.Point(196, 21)
        Me.lblEinouRuikei.Name = "lblEinouRuikei"
        Me.lblEinouRuikei.Size = New System.Drawing.Size(53, 12)
        Me.lblEinouRuikei.TabIndex = 6
        Me.lblEinouRuikei.Text = "営農類型"
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
        Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgcSelect, Me.dgcTodofuken, Me.dgcShityoson, Me.dgcKyuShityoson, Me.dgcNogyoSyuraku, Me.dgcTyosaKubun, Me.dgcKyakutaiNo, Me.dgcCensus, Me.dgcUpdateDate, Me.局, Me.事務所, Me.拠点})
        Me.dgvList.Location = New System.Drawing.Point(21, 107)
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
        Me.dgvList.Size = New System.Drawing.Size(712, 258)
        Me.dgvList.TabIndex = 13
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
        'dgcUpdateDate
        '
        Me.dgcUpdateDate.HeaderText = "更新日時"
        Me.dgcUpdateDate.Name = "dgcUpdateDate"
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
        Me.拠点.ReadOnly = True
        Me.拠点.Visible = False
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(274, 74)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(90, 30)
        Me.btnShow.TabIndex = 12
        Me.btnShow.Text = "表示"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'cboKyoku
        '
        Me.cboKyoku.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKyoku.FormattingEnabled = True
        Me.cboKyoku.Location = New System.Drawing.Point(69, 49)
        Me.cboKyoku.Name = "cboKyoku"
        Me.cboKyoku.Size = New System.Drawing.Size(110, 20)
        Me.cboKyoku.TabIndex = 3
        '
        'lblKyoten
        '
        Me.lblKyoten.AutoSize = True
        Me.lblKyoten.Location = New System.Drawing.Point(22, 83)
        Me.lblKyoten.Name = "lblKyoten"
        Me.lblKyoten.Size = New System.Drawing.Size(77, 12)
        Me.lblKyoten.TabIndex = 4
        Me.lblKyoten.Text = "実査設置拠点"
        '
        'cboKyoten
        '
        Me.cboKyoten.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKyoten.FormattingEnabled = True
        Me.cboKyoten.Location = New System.Drawing.Point(107, 80)
        Me.cboKyoten.Name = "cboKyoten"
        Me.cboKyoten.Size = New System.Drawing.Size(72, 20)
        Me.cboKyoten.TabIndex = 5
        '
        'lblKyoku
        '
        Me.lblKyoku.AutoSize = True
        Me.lblKyoku.Location = New System.Drawing.Point(22, 52)
        Me.lblKyoku.Name = "lblKyoku"
        Me.lblKyoku.Size = New System.Drawing.Size(41, 12)
        Me.lblKyoku.TabIndex = 2
        Me.lblKyoku.Text = "農政局"
        '
        'lblChosaNen
        '
        Me.lblChosaNen.AutoSize = True
        Me.lblChosaNen.Location = New System.Drawing.Point(22, 21)
        Me.lblChosaNen.Name = "lblChosaNen"
        Me.lblChosaNen.Size = New System.Drawing.Size(65, 12)
        Me.lblChosaNen.TabIndex = 0
        Me.lblChosaNen.Text = "調査年（産）"
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
        'btnAllCancel
        '
        Me.btnAllCancel.Location = New System.Drawing.Point(117, 381)
        Me.btnAllCancel.Name = "btnAllCancel"
        Me.btnAllCancel.Size = New System.Drawing.Size(90, 30)
        Me.btnAllCancel.TabIndex = 15
        Me.btnAllCancel.Text = "全解除"
        Me.btnAllCancel.UseVisualStyleBackColor = True
        '
        'btnAllSelect
        '
        Me.btnAllSelect.Location = New System.Drawing.Point(21, 381)
        Me.btnAllSelect.Name = "btnAllSelect"
        Me.btnAllSelect.Size = New System.Drawing.Size(90, 30)
        Me.btnAllSelect.TabIndex = 14
        Me.btnAllSelect.Text = "全選択"
        Me.btnAllSelect.UseVisualStyleBackColor = True
        '
        'BRA4010F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 576)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnModify)
        Me.Name = "BRA4010F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.btnModify, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnModify As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblKyoten As System.Windows.Forms.Label
    Protected WithEvents cboKyoten As System.Windows.Forms.ComboBox
    Friend WithEvents lblKyoku As System.Windows.Forms.Label
    Protected WithEvents cboKyoku As System.Windows.Forms.ComboBox
    Friend WithEvents lblChosaNen As System.Windows.Forms.Label
    Protected WithEvents cboChosaNen As System.Windows.Forms.ComboBox
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Protected WithEvents dgvList As System.Windows.Forms.DataGridView
    Protected WithEvents cboEinouRuikei As System.Windows.Forms.ComboBox
    Friend WithEvents lblEinouRuikei As System.Windows.Forms.Label
    Protected WithEvents cboKessokuchiHokan As System.Windows.Forms.ComboBox
    Friend WithEvents lblKessokuchiHokan As System.Windows.Forms.Label
    Protected WithEvents cboTaishakuTaishohyo As System.Windows.Forms.ComboBox
    Friend WithEvents lblTaishakuTaishohyo As System.Windows.Forms.Label
    Friend WithEvents dgcSelect As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents dgcTodofuken As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcShityoson As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcKyuShityoson As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcNogyoSyuraku As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcTyosaKubun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcKyakutaiNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcCensus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcUpdateDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 局 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 事務所 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 拠点 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnAllCancel As System.Windows.Forms.Button
    Friend WithEvents btnAllSelect As System.Windows.Forms.Button

End Class
