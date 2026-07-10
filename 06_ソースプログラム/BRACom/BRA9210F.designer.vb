<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA9210F
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
        Me.btnOutPut = New System.Windows.Forms.Button()
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
        Me.btnShow = New System.Windows.Forms.Button()
        Me.cboKyoku = New System.Windows.Forms.ComboBox()
        Me.lblKyoten = New System.Windows.Forms.Label()
        Me.cboKyoten = New System.Windows.Forms.ComboBox()
        Me.lblKyoku = New System.Windows.Forms.Label()
        Me.lblChosaNen = New System.Windows.Forms.Label()
        Me.cboChosaNen = New System.Windows.Forms.ComboBox()
        Me.btnAllCancel = New System.Windows.Forms.Button()
        Me.btnAllSelect = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtChohyoMei = New System.Windows.Forms.TextBox()
        Me.btnCSVoutput = New System.Windows.Forms.Button()
        Me.chkShokuin = New System.Windows.Forms.CheckBox()
        Me.GroupBoxYoshi = New System.Windows.Forms.GroupBox()
        Me.rbYoshiA3 = New System.Windows.Forms.RadioButton()
        Me.rbYoshiA4 = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxYoshi.SuspendLayout()
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
        Me.btnReturn.Location = New System.Drawing.Point(704, 543)
        Me.btnReturn.Margin = New System.Windows.Forms.Padding(4)
        Me.btnReturn.TabIndex = 11
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "帳票出力"
        '
        'btnOutPut
        '
        Me.btnOutPut.Location = New System.Drawing.Point(608, 543)
        Me.btnOutPut.Name = "btnOutPut"
        Me.btnOutPut.Size = New System.Drawing.Size(90, 30)
        Me.btnOutPut.TabIndex = 10
        Me.btnOutPut.Text = "出力"
        Me.btnOutPut.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
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
        Me.GroupBox1.Controls.Add(Me.btnAllCancel)
        Me.GroupBox1.Controls.Add(Me.btnAllSelect)
        Me.GroupBox1.Location = New System.Drawing.Point(27, 106)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(754, 426)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "対象客体"
        '
        'cboTaishakuTaishohyo
        '
        Me.cboTaishakuTaishohyo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTaishakuTaishohyo.FormattingEnabled = True
        Me.cboTaishakuTaishohyo.Location = New System.Drawing.Point(457, 49)
        Me.cboTaishakuTaishohyo.Name = "cboTaishakuTaishohyo"
        Me.cboTaishakuTaishohyo.Size = New System.Drawing.Size(90, 20)
        Me.cboTaishakuTaishohyo.TabIndex = 11
        '
        'lblTaishakuTaishohyo
        '
        Me.lblTaishakuTaishohyo.AutoSize = True
        Me.lblTaishakuTaishohyo.Location = New System.Drawing.Point(386, 52)
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
        Me.dgvList.Location = New System.Drawing.Point(21, 107)
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
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(270, 84)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 12)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "帳票名"
        '
        'txtChohyoMei
        '
        Me.txtChohyoMei.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtChohyoMei.Location = New System.Drawing.Point(318, 81)
        Me.txtChohyoMei.MaxLength = 50
        Me.txtChohyoMei.Name = "txtChohyoMei"
        Me.txtChohyoMei.ReadOnly = True
        Me.txtChohyoMei.Size = New System.Drawing.Size(134, 19)
        Me.txtChohyoMei.TabIndex = 6
        Me.txtChohyoMei.TabStop = False
        Me.txtChohyoMei.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnCSVoutput
        '
        Me.btnCSVoutput.Location = New System.Drawing.Point(512, 543)
        Me.btnCSVoutput.Name = "btnCSVoutput"
        Me.btnCSVoutput.Size = New System.Drawing.Size(90, 30)
        Me.btnCSVoutput.TabIndex = 9
        Me.btnCSVoutput.Text = "CSV出力"
        Me.btnCSVoutput.UseVisualStyleBackColor = True
        '
        'chkShokuin
        '
        Me.chkShokuin.AutoSize = True
        Me.chkShokuin.Location = New System.Drawing.Point(365, 551)
        Me.chkShokuin.Name = "chkShokuin"
        Me.chkShokuin.Size = New System.Drawing.Size(141, 16)
        Me.chkShokuin.TabIndex = 8
        Me.chkShokuin.Text = "職員記入欄を折りたたむ"
        Me.chkShokuin.UseVisualStyleBackColor = True
        '
        'GroupBoxYoshi
        '
        Me.GroupBoxYoshi.Controls.Add(Me.rbYoshiA3)
        Me.GroupBoxYoshi.Controls.Add(Me.rbYoshiA4)
        Me.GroupBoxYoshi.Location = New System.Drawing.Point(600, 55)
        Me.GroupBoxYoshi.Name = "GroupBoxYoshi"
        Me.GroupBoxYoshi.Size = New System.Drawing.Size(151, 45)
        Me.GroupBoxYoshi.TabIndex = 12
        Me.GroupBoxYoshi.TabStop = False
        Me.GroupBoxYoshi.Text = "用紙サイズ"
        '
        'rbYoshiA3
        '
        Me.rbYoshiA3.AutoSize = True
        Me.rbYoshiA3.Location = New System.Drawing.Point(90, 18)
        Me.rbYoshiA3.Name = "rbYoshiA3"
        Me.rbYoshiA3.Size = New System.Drawing.Size(37, 16)
        Me.rbYoshiA3.TabIndex = 1
        Me.rbYoshiA3.Text = "A3"
        Me.rbYoshiA3.UseVisualStyleBackColor = True
        '
        'rbYoshiA4
        '
        Me.rbYoshiA4.AutoSize = True
        Me.rbYoshiA4.Checked = True
        Me.rbYoshiA4.Location = New System.Drawing.Point(27, 18)
        Me.rbYoshiA4.Name = "rbYoshiA4"
        Me.rbYoshiA4.Size = New System.Drawing.Size(37, 16)
        Me.rbYoshiA4.TabIndex = 0
        Me.rbYoshiA4.TabStop = True
        Me.rbYoshiA4.Text = "A4"
        Me.rbYoshiA4.UseVisualStyleBackColor = True
        '
        'BRA9210F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 587)
        Me.Controls.Add(Me.GroupBoxYoshi)
        Me.Controls.Add(Me.chkShokuin)
        Me.Controls.Add(Me.btnCSVoutput)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtChohyoMei)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnOutPut)
        Me.Margin = New System.Windows.Forms.Padding(5)
        Me.Name = "BRA9210F"
        Me.Padding = New System.Windows.Forms.Padding(15, 14, 15, 14)
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.btnOutPut, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.txtChohyoMei, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.btnCSVoutput, 0)
        Me.Controls.SetChildIndex(Me.chkShokuin, 0)
        Me.Controls.SetChildIndex(Me.GroupBoxYoshi, 0)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxYoshi.ResumeLayout(False)
        Me.GroupBoxYoshi.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOutPut As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnAllCancel As System.Windows.Forms.Button
    Friend WithEvents btnAllSelect As System.Windows.Forms.Button
    Friend WithEvents lblKyoten As System.Windows.Forms.Label
    Protected WithEvents cboKyoten As System.Windows.Forms.ComboBox
    Friend WithEvents lblKyoku As System.Windows.Forms.Label
    Protected WithEvents cboKyoku As System.Windows.Forms.ComboBox
    Friend WithEvents lblChosaNen As System.Windows.Forms.Label
    Protected WithEvents cboChosaNen As System.Windows.Forms.ComboBox
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtChohyoMei As System.Windows.Forms.TextBox
    Protected WithEvents dgvList As System.Windows.Forms.DataGridView
    Protected WithEvents cboEinouRuikei As System.Windows.Forms.ComboBox
    Friend WithEvents lblEinouRuikei As System.Windows.Forms.Label
    Friend WithEvents dgcSelect As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents dgcTodofuken As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcShityoson As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcKyuShityoson As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcNogyoSyuraku As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcTyosaKubun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcKyakutaiNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcCensus As System.Windows.Forms.DataGridViewTextBoxColumn
    Protected WithEvents cboTaishakuTaishohyo As System.Windows.Forms.ComboBox
    Friend WithEvents lblTaishakuTaishohyo As System.Windows.Forms.Label
    Protected WithEvents cboKessokuchiHokan As System.Windows.Forms.ComboBox
    Friend WithEvents lblKessokuchiHokan As System.Windows.Forms.Label
    Friend WithEvents btnCSVoutput As System.Windows.Forms.Button
    Friend WithEvents chkShokuin As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBoxYoshi As GroupBox
    Friend WithEvents rbYoshiA3 As RadioButton
    Friend WithEvents rbYoshiA4 As RadioButton
End Class
