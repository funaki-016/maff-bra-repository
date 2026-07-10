<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA9330F
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
        Me.txtEinouKeieitai = New System.Windows.Forms.TextBox()
        Me.txtChosanen = New System.Windows.Forms.TextBox()
        Me.lblSyukeiNoZennen = New System.Windows.Forms.Label()
        Me.lblSyukeiNoHonnen = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cboKiboKaisou = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboChiiki = New System.Windows.Forms.ComboBox()
        Me.txtSyukeiNameZennen = New System.Windows.Forms.TextBox()
        Me.txtSyukeiNameHonnen = New System.Windows.Forms.TextBox()
        Me.lblSyukeiNameZennen = New System.Windows.Forms.Label()
        Me.lblSyukeiNameHonnen = New System.Windows.Forms.Label()
        Me.txtSyukeiNoZennen = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtSyukeiNoHonnen = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.dgcSelect = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.dgcHeikinSyurui = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcKiboKaisou = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcChiiki = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcShukeiKosuuHonnen = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcShukeiKosuuZennen = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.本年連番 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.前年連番 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.田畑区分 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ビール麦販売区分 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.てんさい栽培区分 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnShow = New System.Windows.Forms.Button()
        Me.lblEinouKeieitai = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtChohyoMei = New System.Windows.Forms.TextBox()
        Me.btnAllSelect = New System.Windows.Forms.Button()
        Me.btnAllCancel = New System.Windows.Forms.Button()
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
        Me.btnReturn.Location = New System.Drawing.Point(704, 617)
        Me.btnReturn.TabIndex = 11
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "集計結果帳票出力"
        '
        'btnOutPut
        '
        Me.btnOutPut.Location = New System.Drawing.Point(608, 617)
        Me.btnOutPut.Name = "btnOutPut"
        Me.btnOutPut.Size = New System.Drawing.Size(90, 30)
        Me.btnOutPut.TabIndex = 10
        Me.btnOutPut.Text = "出力"
        Me.btnOutPut.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtEinouKeieitai)
        Me.GroupBox1.Controls.Add(Me.txtChosanen)
        Me.GroupBox1.Controls.Add(Me.lblSyukeiNoZennen)
        Me.GroupBox1.Controls.Add(Me.lblSyukeiNoHonnen)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.cboKiboKaisou)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.cboChiiki)
        Me.GroupBox1.Controls.Add(Me.txtSyukeiNameZennen)
        Me.GroupBox1.Controls.Add(Me.txtSyukeiNameHonnen)
        Me.GroupBox1.Controls.Add(Me.lblSyukeiNameZennen)
        Me.GroupBox1.Controls.Add(Me.lblSyukeiNameHonnen)
        Me.GroupBox1.Controls.Add(Me.txtSyukeiNoZennen)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtSyukeiNoHonnen)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.dgvList)
        Me.GroupBox1.Controls.Add(Me.btnShow)
        Me.GroupBox1.Controls.Add(Me.lblEinouKeieitai)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Location = New System.Drawing.Point(27, 107)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(754, 504)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        '
        'txtEinouKeieitai
        '
        Me.txtEinouKeieitai.Location = New System.Drawing.Point(322, 22)
        Me.txtEinouKeieitai.Name = "txtEinouKeieitai"
        Me.txtEinouKeieitai.ReadOnly = True
        Me.txtEinouKeieitai.Size = New System.Drawing.Size(72, 19)
        Me.txtEinouKeieitai.TabIndex = 3
        Me.txtEinouKeieitai.TabStop = False
        '
        'txtChosanen
        '
        Me.txtChosanen.Location = New System.Drawing.Point(90, 22)
        Me.txtChosanen.Name = "txtChosanen"
        Me.txtChosanen.ReadOnly = True
        Me.txtChosanen.Size = New System.Drawing.Size(72, 19)
        Me.txtChosanen.TabIndex = 1
        Me.txtChosanen.TabStop = False
        '
        'lblSyukeiNoZennen
        '
        Me.lblSyukeiNoZennen.AutoSize = True
        Me.lblSyukeiNoZennen.Location = New System.Drawing.Point(351, 83)
        Me.lblSyukeiNoZennen.Name = "lblSyukeiNoZennen"
        Me.lblSyukeiNoZennen.Size = New System.Drawing.Size(53, 12)
        Me.lblSyukeiNoZennen.TabIndex = 10
        Me.lblSyukeiNoZennen.Text = "集計番号"
        '
        'lblSyukeiNoHonnen
        '
        Me.lblSyukeiNoHonnen.AutoSize = True
        Me.lblSyukeiNoHonnen.Location = New System.Drawing.Point(351, 51)
        Me.lblSyukeiNoHonnen.Name = "lblSyukeiNoHonnen"
        Me.lblSyukeiNoHonnen.Size = New System.Drawing.Size(53, 12)
        Me.lblSyukeiNoHonnen.TabIndex = 5
        Me.lblSyukeiNoHonnen.Text = "集計番号"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(31, 120)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(53, 12)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "規模階層"
        '
        'cboKiboKaisou
        '
        Me.cboKiboKaisou.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKiboKaisou.FormattingEnabled = True
        Me.cboKiboKaisou.Location = New System.Drawing.Point(90, 117)
        Me.cboKiboKaisou.Name = "cboKiboKaisou"
        Me.cboKiboKaisou.Size = New System.Drawing.Size(90, 20)
        Me.cboKiboKaisou.TabIndex = 15
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(287, 120)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(29, 12)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "地域"
        '
        'cboChiiki
        '
        Me.cboChiiki.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChiiki.FormattingEnabled = True
        Me.cboChiiki.Location = New System.Drawing.Point(322, 117)
        Me.cboChiiki.Name = "cboChiiki"
        Me.cboChiiki.Size = New System.Drawing.Size(126, 20)
        Me.cboChiiki.TabIndex = 17
        '
        'txtSyukeiNameZennen
        '
        Me.txtSyukeiNameZennen.Location = New System.Drawing.Point(580, 80)
        Me.txtSyukeiNameZennen.Name = "txtSyukeiNameZennen"
        Me.txtSyukeiNameZennen.ReadOnly = True
        Me.txtSyukeiNameZennen.Size = New System.Drawing.Size(153, 19)
        Me.txtSyukeiNameZennen.TabIndex = 13
        Me.txtSyukeiNameZennen.TabStop = False
        '
        'txtSyukeiNameHonnen
        '
        Me.txtSyukeiNameHonnen.Location = New System.Drawing.Point(580, 48)
        Me.txtSyukeiNameHonnen.Name = "txtSyukeiNameHonnen"
        Me.txtSyukeiNameHonnen.ReadOnly = True
        Me.txtSyukeiNameHonnen.Size = New System.Drawing.Size(153, 19)
        Me.txtSyukeiNameHonnen.TabIndex = 8
        Me.txtSyukeiNameHonnen.TabStop = False
        '
        'lblSyukeiNameZennen
        '
        Me.lblSyukeiNameZennen.AutoSize = True
        Me.lblSyukeiNameZennen.Location = New System.Drawing.Point(521, 83)
        Me.lblSyukeiNameZennen.Name = "lblSyukeiNameZennen"
        Me.lblSyukeiNameZennen.Size = New System.Drawing.Size(53, 12)
        Me.lblSyukeiNameZennen.TabIndex = 12
        Me.lblSyukeiNameZennen.Text = "集計名称"
        '
        'lblSyukeiNameHonnen
        '
        Me.lblSyukeiNameHonnen.AutoSize = True
        Me.lblSyukeiNameHonnen.Location = New System.Drawing.Point(521, 51)
        Me.lblSyukeiNameHonnen.Name = "lblSyukeiNameHonnen"
        Me.lblSyukeiNameHonnen.Size = New System.Drawing.Size(53, 12)
        Me.lblSyukeiNameHonnen.TabIndex = 7
        Me.lblSyukeiNameHonnen.Text = "集計名称"
        '
        'txtSyukeiNoZennen
        '
        Me.txtSyukeiNoZennen.Location = New System.Drawing.Point(410, 80)
        Me.txtSyukeiNoZennen.Name = "txtSyukeiNoZennen"
        Me.txtSyukeiNoZennen.ReadOnly = True
        Me.txtSyukeiNoZennen.Size = New System.Drawing.Size(90, 19)
        Me.txtSyukeiNoZennen.TabIndex = 11
        Me.txtSyukeiNoZennen.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(287, 83)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 12)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "前年"
        '
        'txtSyukeiNoHonnen
        '
        Me.txtSyukeiNoHonnen.Location = New System.Drawing.Point(410, 48)
        Me.txtSyukeiNoHonnen.Name = "txtSyukeiNoHonnen"
        Me.txtSyukeiNoHonnen.ReadOnly = True
        Me.txtSyukeiNoHonnen.Size = New System.Drawing.Size(90, 19)
        Me.txtSyukeiNoHonnen.TabIndex = 6
        Me.txtSyukeiNoHonnen.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(287, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 12)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "本年"
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
        Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgcSelect, Me.dgcHeikinSyurui, Me.dgcKiboKaisou, Me.dgcChiiki, Me.dgcShukeiKosuuHonnen, Me.dgcShukeiKosuuZennen, Me.本年連番, Me.前年連番, Me.田畑区分, Me.ビール麦販売区分, Me.てんさい栽培区分})
        Me.dgvList.Location = New System.Drawing.Point(21, 163)
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
        Me.dgvList.Size = New System.Drawing.Size(712, 317)
        Me.dgvList.TabIndex = 21
        '
        'dgcSelect
        '
        Me.dgcSelect.DataPropertyName = "チェックボックス"
        Me.dgcSelect.HeaderText = ""
        Me.dgcSelect.Name = "dgcSelect"
        Me.dgcSelect.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgcSelect.Width = 30
        '
        'dgcHeikinSyurui
        '
        Me.dgcHeikinSyurui.HeaderText = "生産費平均値種類"
        Me.dgcHeikinSyurui.MaxInputLength = 1
        Me.dgcHeikinSyurui.Name = "dgcHeikinSyurui"
        Me.dgcHeikinSyurui.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcHeikinSyurui.Width = 110
        '
        'dgcKiboKaisou
        '
        Me.dgcKiboKaisou.HeaderText = "規模階層"
        Me.dgcKiboKaisou.MaxInputLength = 1
        Me.dgcKiboKaisou.Name = "dgcKiboKaisou"
        Me.dgcKiboKaisou.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'dgcChiiki
        '
        Me.dgcChiiki.HeaderText = "地域"
        Me.dgcChiiki.MaxInputLength = 8
        Me.dgcChiiki.Name = "dgcChiiki"
        Me.dgcChiiki.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgcChiiki.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcChiiki.Width = 190
        '
        'dgcShukeiKosuuHonnen
        '
        Me.dgcShukeiKosuuHonnen.HeaderText = "本年　集計経営体数"
        Me.dgcShukeiKosuuHonnen.Name = "dgcShukeiKosuuHonnen"
        Me.dgcShukeiKosuuHonnen.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgcShukeiKosuuHonnen.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcShukeiKosuuHonnen.Width = 130
        '
        'dgcShukeiKosuuZennen
        '
        Me.dgcShukeiKosuuZennen.HeaderText = "前年　集計経営体数"
        Me.dgcShukeiKosuuZennen.Name = "dgcShukeiKosuuZennen"
        Me.dgcShukeiKosuuZennen.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgcShukeiKosuuZennen.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcShukeiKosuuZennen.Width = 130
        '
        '本年連番
        '
        Me.本年連番.HeaderText = "本年連番"
        Me.本年連番.Name = "本年連番"
        Me.本年連番.Visible = False
        '
        '前年連番
        '
        Me.前年連番.HeaderText = "前年連番"
        Me.前年連番.Name = "前年連番"
        Me.前年連番.Visible = False
        '
        '田畑区分
        '
        Me.田畑区分.HeaderText = "田畑区分"
        Me.田畑区分.Name = "田畑区分"
        '
        'ビール麦販売区分
        '
        Me.ビール麦販売区分.HeaderText = "ビール麦販売区分"
        Me.ビール麦販売区分.Name = "ビール麦販売区分"
        '
        'てんさい栽培区分
        '
        Me.てんさい栽培区分.HeaderText = "てんさい栽培区分"
        Me.てんさい栽培区分.Name = "てんさい栽培区分"
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(643, 111)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(90, 30)
        Me.btnShow.TabIndex = 18
        Me.btnShow.Text = "表示"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'lblEinouKeieitai
        '
        Me.lblEinouKeieitai.AutoSize = True
        Me.lblEinouKeieitai.Location = New System.Drawing.Point(227, 25)
        Me.lblEinouKeieitai.Name = "lblEinouKeieitai"
        Me.lblEinouKeieitai.Size = New System.Drawing.Size(89, 12)
        Me.lblEinouKeieitai.TabIndex = 2
        Me.lblEinouKeieitai.Text = "営農経営体区分"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(19, 25)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 12)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "調査年（産）"
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
        Me.txtChohyoMei.Size = New System.Drawing.Size(138, 19)
        Me.txtChohyoMei.TabIndex = 6
        Me.txtChohyoMei.TabStop = False
        Me.txtChohyoMei.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnAllSelect
        '
        Me.btnAllSelect.Location = New System.Drawing.Point(60, 617)
        Me.btnAllSelect.Name = "btnAllSelect"
        Me.btnAllSelect.Size = New System.Drawing.Size(90, 30)
        Me.btnAllSelect.TabIndex = 8
        Me.btnAllSelect.Text = "全選択"
        Me.btnAllSelect.UseVisualStyleBackColor = True
        '
        'btnAllCancel
        '
        Me.btnAllCancel.Location = New System.Drawing.Point(156, 617)
        Me.btnAllCancel.Name = "btnAllCancel"
        Me.btnAllCancel.Size = New System.Drawing.Size(90, 30)
        Me.btnAllCancel.TabIndex = 9
        Me.btnAllCancel.Text = "全解除"
        Me.btnAllCancel.UseVisualStyleBackColor = True
        '
        'BRA9330F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 661)
        Me.Controls.Add(Me.btnAllCancel)
        Me.Controls.Add(Me.btnAllSelect)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtChohyoMei)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnOutPut)
        Me.Name = "BRA9330F"
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
        Me.Controls.SetChildIndex(Me.btnAllSelect, 0)
        Me.Controls.SetChildIndex(Me.btnAllCancel, 0)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOutPut As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblEinouKeieitai As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Protected WithEvents dgvList As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtChohyoMei As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Protected WithEvents cboKiboKaisou As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Protected WithEvents cboChiiki As System.Windows.Forms.ComboBox
    Friend WithEvents txtSyukeiNameZennen As System.Windows.Forms.TextBox
    Friend WithEvents txtSyukeiNameHonnen As System.Windows.Forms.TextBox
    Friend WithEvents lblSyukeiNameZennen As System.Windows.Forms.Label
    Friend WithEvents lblSyukeiNameHonnen As System.Windows.Forms.Label
    Friend WithEvents txtSyukeiNoZennen As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSyukeiNoHonnen As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnAllSelect As System.Windows.Forms.Button
    Friend WithEvents btnAllCancel As System.Windows.Forms.Button
    Friend WithEvents lblSyukeiNoZennen As System.Windows.Forms.Label
    Friend WithEvents lblSyukeiNoHonnen As System.Windows.Forms.Label
    Friend WithEvents txtEinouKeieitai As System.Windows.Forms.TextBox
    Friend WithEvents txtChosanen As System.Windows.Forms.TextBox
    Friend WithEvents dgcSelect As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents dgcHeikinSyurui As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcKiboKaisou As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcChiiki As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcShukeiKosuuHonnen As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcShukeiKosuuZennen As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 本年連番 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 前年連番 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 田畑区分 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ビール麦販売区分 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents てんさい栽培区分 As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
