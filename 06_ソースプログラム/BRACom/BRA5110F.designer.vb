<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA5110F
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
        Me.btnAggregate = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblCensusShitei = New System.Windows.Forms.Label()
        Me.btnInputCensus = New System.Windows.Forms.Button()
        Me.cboEinouKeieitai = New System.Windows.Forms.ComboBox()
        Me.lblEinouKeieitai = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboNiniKaisou = New System.Windows.Forms.ComboBox()
        Me.cboSyukei4 = New System.Windows.Forms.ComboBox()
        Me.cboSyukei3 = New System.Windows.Forms.ComboBox()
        Me.cboSyukei2 = New System.Windows.Forms.ComboBox()
        Me.cboKiboKaisou = New System.Windows.Forms.ComboBox()
        Me.cboKeizokuKubun = New System.Windows.Forms.ComboBox()
        Me.lblKeizokuKubun = New System.Windows.Forms.Label()
        Me.chkZenkaiCensus = New System.Windows.Forms.CheckBox()
        Me.cboSyukei1 = New System.Windows.Forms.ComboBox()
        Me.btnBumon = New System.Windows.Forms.Button()
        Me.btnChiiki = New System.Windows.Forms.Button()
        Me.lblSyukeiName = New System.Windows.Forms.Label()
        Me.lblSyukeiNo = New System.Windows.Forms.Label()
        Me.txtSyukeiName = New System.Windows.Forms.TextBox()
        Me.txtSyukeiNo = New System.Windows.Forms.TextBox()
        Me.txtBumon = New System.Windows.Forms.RichTextBox()
        Me.txtChiiki = New System.Windows.Forms.RichTextBox()
        Me.lblSyukei4 = New System.Windows.Forms.Label()
        Me.lblSyukei3 = New System.Windows.Forms.Label()
        Me.lblSyukei2 = New System.Windows.Forms.Label()
        Me.lblRead = New System.Windows.Forms.Label()
        Me.lblKiboKaisou = New System.Windows.Forms.Label()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.dgcJouken1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcJouken2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcJouken3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcJouken4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnRead = New System.Windows.Forms.Button()
        Me.cboHeikinSyurui = New System.Windows.Forms.ComboBox()
        Me.lblSyukei1 = New System.Windows.Forms.Label()
        Me.lbHeikinSyurui = New System.Windows.Forms.Label()
        Me.lblChosaNen = New System.Windows.Forms.Label()
        Me.cboChosaNen = New System.Windows.Forms.ComboBox()
        Me.btnSyukeiJouken = New System.Windows.Forms.Button()
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
        Me.btnReturn.Location = New System.Drawing.Point(712, 615)
        Me.btnReturn.TabIndex = 0
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
        Me.lblSyori.Text = "集計結果表作成"
        '
        'btnAggregate
        '
        Me.btnAggregate.Location = New System.Drawing.Point(616, 615)
        Me.btnAggregate.Name = "btnAggregate"
        Me.btnAggregate.Size = New System.Drawing.Size(90, 30)
        Me.btnAggregate.TabIndex = 8
        Me.btnAggregate.Text = "集計"
        Me.btnAggregate.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblCensusShitei)
        Me.GroupBox1.Controls.Add(Me.btnInputCensus)
        Me.GroupBox1.Controls.Add(Me.cboEinouKeieitai)
        Me.GroupBox1.Controls.Add(Me.lblEinouKeieitai)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cboNiniKaisou)
        Me.GroupBox1.Controls.Add(Me.cboSyukei4)
        Me.GroupBox1.Controls.Add(Me.cboSyukei3)
        Me.GroupBox1.Controls.Add(Me.cboSyukei2)
        Me.GroupBox1.Controls.Add(Me.cboKiboKaisou)
        Me.GroupBox1.Controls.Add(Me.cboKeizokuKubun)
        Me.GroupBox1.Controls.Add(Me.cboSyukei1)
        Me.GroupBox1.Controls.Add(Me.btnBumon)
        Me.GroupBox1.Controls.Add(Me.btnChiiki)
        Me.GroupBox1.Controls.Add(Me.lblSyukeiName)
        Me.GroupBox1.Controls.Add(Me.lblSyukeiNo)
        Me.GroupBox1.Controls.Add(Me.txtSyukeiName)
        Me.GroupBox1.Controls.Add(Me.txtSyukeiNo)
        Me.GroupBox1.Controls.Add(Me.txtBumon)
        Me.GroupBox1.Controls.Add(Me.txtChiiki)
        Me.GroupBox1.Controls.Add(Me.lblSyukei4)
        Me.GroupBox1.Controls.Add(Me.lblSyukei3)
        Me.GroupBox1.Controls.Add(Me.lblSyukei2)
        Me.GroupBox1.Controls.Add(Me.lblRead)
        Me.GroupBox1.Controls.Add(Me.lblKiboKaisou)
        Me.GroupBox1.Controls.Add(Me.lblKeizokuKubun)
        Me.GroupBox1.Controls.Add(Me.chkZenkaiCensus)
        Me.GroupBox1.Controls.Add(Me.dgvList)
        Me.GroupBox1.Controls.Add(Me.btnRead)
        Me.GroupBox1.Controls.Add(Me.cboHeikinSyurui)
        Me.GroupBox1.Controls.Add(Me.lblSyukei1)
        Me.GroupBox1.Controls.Add(Me.lbHeikinSyurui)
        Me.GroupBox1.Controls.Add(Me.lblChosaNen)
        Me.GroupBox1.Controls.Add(Me.cboChosaNen)
        Me.GroupBox1.Location = New System.Drawing.Point(27, 68)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(775, 541)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "集計条件設定"
        '
        'lblCensusShitei
        '
        Me.lblCensusShitei.AutoSize = True
        Me.lblCensusShitei.Location = New System.Drawing.Point(120, 506)
        Me.lblCensusShitei.Name = "lblCensusShitei"
        Me.lblCensusShitei.Size = New System.Drawing.Size(0, 12)
        Me.lblCensusShitei.TabIndex = 28
        '
        'btnInputCensus
        '
        Me.btnInputCensus.Location = New System.Drawing.Point(8, 497)
        Me.btnInputCensus.Name = "btnInputCensus"
        Me.btnInputCensus.Size = New System.Drawing.Size(106, 30)
        Me.btnInputCensus.TabIndex = 27
        Me.btnInputCensus.Text = "センサス番号取込"
        Me.btnInputCensus.UseVisualStyleBackColor = True
        '
        'cboEinouKeieitai
        '
        Me.cboEinouKeieitai.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEinouKeieitai.FormattingEnabled = True
        Me.cboEinouKeieitai.Location = New System.Drawing.Point(263, 37)
        Me.cboEinouKeieitai.Name = "cboEinouKeieitai"
        Me.cboEinouKeieitai.Size = New System.Drawing.Size(97, 20)
        Me.cboEinouKeieitai.TabIndex = 5
        '
        'lblEinouKeieitai
        '
        Me.lblEinouKeieitai.AutoSize = True
        Me.lblEinouKeieitai.Location = New System.Drawing.Point(168, 40)
        Me.lblEinouKeieitai.Name = "lblEinouKeieitai"
        Me.lblEinouKeieitai.Size = New System.Drawing.Size(89, 12)
        Me.lblEinouKeieitai.TabIndex = 4
        Me.lblEinouKeieitai.Text = "営農経営体区分"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(180, 93)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 12)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "任意階層利用"
        '
        'cboNiniKaisou
        '
        Me.cboNiniKaisou.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboNiniKaisou.FormattingEnabled = True
        Me.cboNiniKaisou.Location = New System.Drawing.Point(263, 90)
        Me.cboNiniKaisou.Name = "cboNiniKaisou"
        Me.cboNiniKaisou.Size = New System.Drawing.Size(97, 20)
        Me.cboNiniKaisou.TabIndex = 9
        '
        'cboSyukei4
        '
        Me.cboSyukei4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSyukei4.FormattingEnabled = True
        Me.cboSyukei4.Location = New System.Drawing.Point(665, 147)
        Me.cboSyukei4.Name = "cboSyukei4"
        Me.cboSyukei4.Size = New System.Drawing.Size(97, 20)
        Me.cboSyukei4.TabIndex = 21
        '
        'cboSyukei3
        '
        Me.cboSyukei3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSyukei3.FormattingEnabled = True
        Me.cboSyukei3.Location = New System.Drawing.Point(466, 147)
        Me.cboSyukei3.Name = "cboSyukei3"
        Me.cboSyukei3.Size = New System.Drawing.Size(97, 20)
        Me.cboSyukei3.TabIndex = 19
        '
        'cboSyukei2
        '
        Me.cboSyukei2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSyukei2.FormattingEnabled = True
        Me.cboSyukei2.Location = New System.Drawing.Point(263, 147)
        Me.cboSyukei2.Name = "cboSyukei2"
        Me.cboSyukei2.Size = New System.Drawing.Size(97, 20)
        Me.cboSyukei2.TabIndex = 17
        '
        'cboKiboKaisou
        '
        Me.cboKiboKaisou.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKiboKaisou.FormattingEnabled = True
        Me.cboKiboKaisou.Location = New System.Drawing.Point(466, 90)
        Me.cboKiboKaisou.Name = "cboKiboKaisou"
        Me.cboKiboKaisou.Size = New System.Drawing.Size(97, 20)
        Me.cboKiboKaisou.TabIndex = 11
        '
        'cboKeizokuKubun
        '
        Me.cboKeizokuKubun.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKeizokuKubun.FormattingEnabled = True
        Me.cboKeizokuKubun.Location = New System.Drawing.Point(665, 90)
        Me.cboKeizokuKubun.Name = "cboKeizokuKubun"
        Me.cboKeizokuKubun.Size = New System.Drawing.Size(97, 20)
        Me.cboKeizokuKubun.TabIndex = 12
        '
        'cboSyukei1
        '
        Me.cboSyukei1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSyukei1.FormattingEnabled = True
        Me.cboSyukei1.Location = New System.Drawing.Point(65, 147)
        Me.cboSyukei1.Name = "cboSyukei1"
        Me.cboSyukei1.Size = New System.Drawing.Size(97, 20)
        Me.cboSyukei1.TabIndex = 15
        '
        'btnBumon
        '
        Me.btnBumon.Location = New System.Drawing.Point(409, 192)
        Me.btnBumon.Name = "btnBumon"
        Me.btnBumon.Size = New System.Drawing.Size(86, 30)
        Me.btnBumon.TabIndex = 24
        Me.btnBumon.Text = "部門選択"
        Me.btnBumon.UseVisualStyleBackColor = True
        '
        'btnChiiki
        '
        Me.btnChiiki.Location = New System.Drawing.Point(20, 192)
        Me.btnChiiki.Name = "btnChiiki"
        Me.btnChiiki.Size = New System.Drawing.Size(86, 30)
        Me.btnChiiki.TabIndex = 22
        Me.btnChiiki.Text = "地域選択"
        Me.btnChiiki.UseVisualStyleBackColor = True
        '
        'lblSyukeiName
        '
        Me.lblSyukeiName.AutoSize = True
        Me.lblSyukeiName.Location = New System.Drawing.Point(592, 488)
        Me.lblSyukeiName.Name = "lblSyukeiName"
        Me.lblSyukeiName.Size = New System.Drawing.Size(53, 12)
        Me.lblSyukeiName.TabIndex = 30
        Me.lblSyukeiName.Text = "集計名称"
        '
        'lblSyukeiNo
        '
        Me.lblSyukeiNo.AutoSize = True
        Me.lblSyukeiNo.Location = New System.Drawing.Point(425, 488)
        Me.lblSyukeiNo.Name = "lblSyukeiNo"
        Me.lblSyukeiNo.Size = New System.Drawing.Size(53, 12)
        Me.lblSyukeiNo.TabIndex = 29
        Me.lblSyukeiNo.Text = "集計番号"
        '
        'txtSyukeiName
        '
        Me.txtSyukeiName.Location = New System.Drawing.Point(498, 503)
        Me.txtSyukeiName.MaxLength = 40
        Me.txtSyukeiName.Name = "txtSyukeiName"
        Me.txtSyukeiName.Size = New System.Drawing.Size(245, 19)
        Me.txtSyukeiName.TabIndex = 32
        '
        'txtSyukeiNo
        '
        Me.txtSyukeiNo.Location = New System.Drawing.Point(398, 503)
        Me.txtSyukeiNo.MaxLength = 6
        Me.txtSyukeiNo.Name = "txtSyukeiNo"
        Me.txtSyukeiNo.Size = New System.Drawing.Size(100, 19)
        Me.txtSyukeiNo.TabIndex = 31
        '
        'txtBumon
        '
        Me.txtBumon.Location = New System.Drawing.Point(495, 192)
        Me.txtBumon.Name = "txtBumon"
        Me.txtBumon.ReadOnly = True
        Me.txtBumon.Size = New System.Drawing.Size(246, 48)
        Me.txtBumon.TabIndex = 25
        Me.txtBumon.Text = ""
        '
        'txtChiiki
        '
        Me.txtChiiki.AutoWordSelection = True
        Me.txtChiiki.Location = New System.Drawing.Point(106, 192)
        Me.txtChiiki.Name = "txtChiiki"
        Me.txtChiiki.ReadOnly = True
        Me.txtChiiki.Size = New System.Drawing.Size(246, 48)
        Me.txtChiiki.TabIndex = 23
        Me.txtChiiki.Text = ""
        '
        'lblSyukei4
        '
        Me.lblSyukei4.Location = New System.Drawing.Point(569, 150)
        Me.lblSyukei4.Name = "lblSyukei4"
        Me.lblSyukei4.Size = New System.Drawing.Size(90, 12)
        Me.lblSyukei4.TabIndex = 20
        Me.lblSyukei4.Text = "てんさい栽培区分"
        Me.lblSyukei4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblSyukei3
        '
        Me.lblSyukei3.Location = New System.Drawing.Point(367, 150)
        Me.lblSyukei3.Name = "lblSyukei3"
        Me.lblSyukei3.Size = New System.Drawing.Size(93, 12)
        Me.lblSyukei3.TabIndex = 18
        Me.lblSyukei3.Text = "ビール麦販売区分"
        Me.lblSyukei3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblSyukei2
        '
        Me.lblSyukei2.Location = New System.Drawing.Point(204, 150)
        Me.lblSyukei2.Name = "lblSyukei2"
        Me.lblSyukei2.Size = New System.Drawing.Size(53, 12)
        Me.lblSyukei2.TabIndex = 16
        Me.lblSyukei2.Text = "田畑区分"
        Me.lblSyukei2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblRead
        '
        Me.lblRead.AutoSize = True
        Me.lblRead.Location = New System.Drawing.Point(589, 35)
        Me.lblRead.Name = "lblRead"
        Me.lblRead.Size = New System.Drawing.Size(77, 12)
        Me.lblRead.TabIndex = 0
        Me.lblRead.Text = "集計条件読込"
        '
        'lblKiboKaisou
        '
        Me.lblKiboKaisou.AutoSize = True
        Me.lblKiboKaisou.Location = New System.Drawing.Point(407, 93)
        Me.lblKiboKaisou.Name = "lblKiboKaisou"
        Me.lblKiboKaisou.Size = New System.Drawing.Size(53, 12)
        Me.lblKiboKaisou.TabIndex = 10
        Me.lblKiboKaisou.Text = "規模階層"
        '
        'lblKeizokuKubun
        '
        Me.lblKeizokuKubun.AutoSize = True
        Me.lblKeizokuKubun.Location = New System.Drawing.Point(606, 93)
        Me.lblKeizokuKubun.Name = "lblKeizokuKubun"
        Me.lblKeizokuKubun.Size = New System.Drawing.Size(53, 12)
        Me.lblKeizokuKubun.TabIndex = 11
        Me.lblKeizokuKubun.Text = "継続区分"
        '
        'chkZenkaiCensus
        '
        Me.chkZenkaiCensus.AutoSize = True
        Me.chkZenkaiCensus.Location = New System.Drawing.Point(606, 116)
        Me.chkZenkaiCensus.Name = "chkZenkaiCensus"
        Me.chkZenkaiCensus.Size = New System.Drawing.Size(156, 16)
        Me.chkZenkaiCensus.TabIndex = 13
        Me.chkZenkaiCensus.Text = "前回センサス番号を使用"
        Me.chkZenkaiCensus.UseVisualStyleBackColor = True
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
        Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgcJouken1, Me.dgcJouken2, Me.dgcJouken3, Me.dgcJouken4})
        Me.dgvList.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnF2
        Me.dgvList.Location = New System.Drawing.Point(29, 246)
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
        Me.dgvList.Size = New System.Drawing.Size(712, 239)
        Me.dgvList.TabIndex = 26
        '
        'dgcJouken1
        '
        Me.dgcJouken1.HeaderText = "集計条件１"
        Me.dgcJouken1.Name = "dgcJouken1"
        Me.dgcJouken1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcJouken1.Width = 173
        '
        'dgcJouken2
        '
        Me.dgcJouken2.HeaderText = "集計条件２"
        Me.dgcJouken2.Name = "dgcJouken2"
        Me.dgcJouken2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcJouken2.Width = 173
        '
        'dgcJouken3
        '
        Me.dgcJouken3.HeaderText = "集計条件３"
        Me.dgcJouken3.Name = "dgcJouken3"
        Me.dgcJouken3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcJouken3.Width = 173
        '
        'dgcJouken4
        '
        Me.dgcJouken4.HeaderText = "集計条件４"
        Me.dgcJouken4.Name = "dgcJouken4"
        Me.dgcJouken4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcJouken4.Width = 173
        '
        'btnRead
        '
        Me.btnRead.Location = New System.Drawing.Point(672, 27)
        Me.btnRead.Name = "btnRead"
        Me.btnRead.Size = New System.Drawing.Size(90, 30)
        Me.btnRead.TabIndex = 1
        Me.btnRead.Text = "読込"
        Me.btnRead.UseVisualStyleBackColor = True
        '
        'cboHeikinSyurui
        '
        Me.cboHeikinSyurui.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboHeikinSyurui.FormattingEnabled = True
        Me.cboHeikinSyurui.Location = New System.Drawing.Point(65, 90)
        Me.cboHeikinSyurui.Name = "cboHeikinSyurui"
        Me.cboHeikinSyurui.Size = New System.Drawing.Size(97, 20)
        Me.cboHeikinSyurui.TabIndex = 7
        '
        'lblSyukei1
        '
        Me.lblSyukei1.Location = New System.Drawing.Point(6, 150)
        Me.lblSyukei1.Name = "lblSyukei1"
        Me.lblSyukei1.Size = New System.Drawing.Size(53, 12)
        Me.lblSyukei1.TabIndex = 14
        Me.lblSyukei1.Text = "集計区分"
        Me.lblSyukei1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbHeikinSyurui
        '
        Me.lbHeikinSyurui.AutoSize = True
        Me.lbHeikinSyurui.Location = New System.Drawing.Point(6, 93)
        Me.lbHeikinSyurui.Name = "lbHeikinSyurui"
        Me.lbHeikinSyurui.Size = New System.Drawing.Size(53, 12)
        Me.lbHeikinSyurui.TabIndex = 6
        Me.lbHeikinSyurui.Text = "平均種類"
        '
        'lblChosaNen
        '
        Me.lblChosaNen.AutoSize = True
        Me.lblChosaNen.Location = New System.Drawing.Point(6, 40)
        Me.lblChosaNen.Name = "lblChosaNen"
        Me.lblChosaNen.Size = New System.Drawing.Size(65, 12)
        Me.lblChosaNen.TabIndex = 2
        Me.lblChosaNen.Text = "調査年（産）"
        '
        'cboChosaNen
        '
        Me.cboChosaNen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChosaNen.FormattingEnabled = True
        Me.cboChosaNen.Location = New System.Drawing.Point(90, 37)
        Me.cboChosaNen.Name = "cboChosaNen"
        Me.cboChosaNen.Size = New System.Drawing.Size(72, 20)
        Me.cboChosaNen.TabIndex = 3
        '
        'btnSyukeiJouken
        '
        Me.btnSyukeiJouken.Location = New System.Drawing.Point(27, 615)
        Me.btnSyukeiJouken.Name = "btnSyukeiJouken"
        Me.btnSyukeiJouken.Size = New System.Drawing.Size(90, 30)
        Me.btnSyukeiJouken.TabIndex = 7
        Me.btnSyukeiJouken.Text = "集計条件出力"
        Me.btnSyukeiJouken.UseVisualStyleBackColor = True
        '
        'BRA5110F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(816, 661)
        Me.Controls.Add(Me.btnSyukeiJouken)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnAggregate)
        Me.Name = "BRA5110F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.btnAggregate, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.btnSyukeiJouken, 0)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnAggregate As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblSyukei1 As System.Windows.Forms.Label
    Friend WithEvents lbHeikinSyurui As System.Windows.Forms.Label
    Protected WithEvents cboHeikinSyurui As System.Windows.Forms.ComboBox
    Friend WithEvents lblChosaNen As System.Windows.Forms.Label
    Protected WithEvents cboChosaNen As System.Windows.Forms.ComboBox
    Friend WithEvents btnRead As System.Windows.Forms.Button
    Protected WithEvents dgvList As System.Windows.Forms.DataGridView
    Friend WithEvents lblKiboKaisou As System.Windows.Forms.Label
    Friend WithEvents dgcJouken1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcJouken2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcJouken3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcJouken4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcTyosaKubun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcKyakutaiNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcCensus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcUpdateDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 局 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 事務所 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 拠点 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents lblRead As System.Windows.Forms.Label
    Friend WithEvents lblSyukei2 As System.Windows.Forms.Label
    Friend WithEvents lblSyukei4 As System.Windows.Forms.Label
    Friend WithEvents lblSyukei3 As System.Windows.Forms.Label
    Friend WithEvents btnSyukeiJouken As System.Windows.Forms.Button
    Friend WithEvents txtBumon As System.Windows.Forms.RichTextBox
    Friend WithEvents txtChiiki As System.Windows.Forms.RichTextBox
    Friend WithEvents lblSyukeiName As System.Windows.Forms.Label
    Friend WithEvents lblSyukeiNo As System.Windows.Forms.Label
    Friend WithEvents txtSyukeiName As System.Windows.Forms.TextBox
    Friend WithEvents txtSyukeiNo As System.Windows.Forms.TextBox
    Friend WithEvents btnBumon As System.Windows.Forms.Button
    Friend WithEvents btnChiiki As System.Windows.Forms.Button
    Protected WithEvents cboSyukei4 As System.Windows.Forms.ComboBox
    Protected WithEvents cboSyukei3 As System.Windows.Forms.ComboBox
    Protected WithEvents cboSyukei2 As System.Windows.Forms.ComboBox
    Protected WithEvents cboKiboKaisou As System.Windows.Forms.ComboBox
    Protected WithEvents cboSyukei1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Protected WithEvents cboNiniKaisou As System.Windows.Forms.ComboBox
    Protected WithEvents cboEinouKeieitai As System.Windows.Forms.ComboBox
    Friend WithEvents lblEinouKeieitai As System.Windows.Forms.Label
    Friend WithEvents btnInputCensus As System.Windows.Forms.Button
    Friend WithEvents lblCensusShitei As System.Windows.Forms.Label
    Protected WithEvents cboKeizokuKubun As System.Windows.Forms.ComboBox
    Friend WithEvents lblKeizokuKubun As System.Windows.Forms.Label
    Friend WithEvents chkZenkaiCensus As System.Windows.Forms.CheckBox

End Class
