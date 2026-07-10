<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BRA10210F
    Inherits Maff.BRA.HeaderBaseForm

    'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.CboEinoRuikei = New System.Windows.Forms.ComboBox()
        Me.lblEinouRuikei = New System.Windows.Forms.Label()
        Me.BtnShow = New System.Windows.Forms.Button()
        Me.CboKyoku = New System.Windows.Forms.ComboBox()
        Me.lblKyoten = New System.Windows.Forms.Label()
        Me.CboKyoten = New System.Windows.Forms.ComboBox()
        Me.lblKyoku = New System.Windows.Forms.Label()
        Me.BtnAllRelease = New System.Windows.Forms.Button()
        Me.BtnAllSelect = New System.Windows.Forms.Button()
        Me.CboCensusNen = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CboKeieiKeitai = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.CboTahata = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.CboSeisanhi = New System.Windows.Forms.ComboBox()
        Me.BtnCsvOutput = New System.Windows.Forms.Button()
        Me.TxtTodofuken = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TxtShikuchoson = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.BtnDel = New System.Windows.Forms.Button()
        Me.BtnAdd = New System.Windows.Forms.Button()
        Me.BtnSave = New System.Windows.Forms.Button()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.TxtIchirenNo = New System.Windows.Forms.TextBox()
        Me.CboKiboKaiso = New System.Windows.Forms.ComboBox()
        Me.CboHyohon = New System.Windows.Forms.ComboBox()
        Me.CboHyohonKoho = New System.Windows.Forms.ComboBox()
        Me.BtnClear = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
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
        Me.btnReturn.TabIndex = 28
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "標本リスト管理"
        '
        'CboEinoRuikei
        '
        Me.CboEinoRuikei.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboEinoRuikei.FormattingEnabled = True
        Me.CboEinoRuikei.Location = New System.Drawing.Point(114, 196)
        Me.CboEinoRuikei.Name = "CboEinoRuikei"
        Me.CboEinoRuikei.Size = New System.Drawing.Size(110, 20)
        Me.CboEinoRuikei.TabIndex = 14
        '
        'lblEinouRuikei
        '
        Me.lblEinouRuikei.AutoSize = True
        Me.lblEinouRuikei.Location = New System.Drawing.Point(29, 199)
        Me.lblEinouRuikei.Name = "lblEinouRuikei"
        Me.lblEinouRuikei.Size = New System.Drawing.Size(77, 12)
        Me.lblEinouRuikei.TabIndex = 13
        Me.lblEinouRuikei.Text = "営農類型区分"
        '
        'BtnShow
        '
        Me.BtnShow.Location = New System.Drawing.Point(704, 250)
        Me.BtnShow.Name = "BtnShow"
        Me.BtnShow.Size = New System.Drawing.Size(90, 30)
        Me.BtnShow.TabIndex = 21
        Me.BtnShow.Text = "表示"
        Me.BtnShow.UseVisualStyleBackColor = True
        '
        'CboKyoku
        '
        Me.CboKyoku.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboKyoku.FormattingEnabled = True
        Me.CboKyoku.Location = New System.Drawing.Point(114, 118)
        Me.CboKyoku.Name = "CboKyoku"
        Me.CboKyoku.Size = New System.Drawing.Size(110, 20)
        Me.CboKyoku.TabIndex = 8
        '
        'lblKyoten
        '
        Me.lblKyoten.AutoSize = True
        Me.lblKyoten.Location = New System.Drawing.Point(29, 147)
        Me.lblKyoten.Name = "lblKyoten"
        Me.lblKyoten.Size = New System.Drawing.Size(77, 12)
        Me.lblKyoten.TabIndex = 9
        Me.lblKyoten.Text = "実査設置拠点"
        '
        'CboKyoten
        '
        Me.CboKyoten.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboKyoten.FormattingEnabled = True
        Me.CboKyoten.Location = New System.Drawing.Point(114, 144)
        Me.CboKyoten.Name = "CboKyoten"
        Me.CboKyoten.Size = New System.Drawing.Size(72, 20)
        Me.CboKyoten.TabIndex = 10
        '
        'lblKyoku
        '
        Me.lblKyoku.AutoSize = True
        Me.lblKyoku.Location = New System.Drawing.Point(29, 121)
        Me.lblKyoku.Name = "lblKyoku"
        Me.lblKyoku.Size = New System.Drawing.Size(41, 12)
        Me.lblKyoku.TabIndex = 7
        Me.lblKyoku.Text = "農政局"
        '
        'BtnAllRelease
        '
        Me.BtnAllRelease.Location = New System.Drawing.Point(114, 540)
        Me.BtnAllRelease.Name = "BtnAllRelease"
        Me.BtnAllRelease.Size = New System.Drawing.Size(90, 30)
        Me.BtnAllRelease.TabIndex = 24
        Me.BtnAllRelease.Text = "全解除"
        Me.BtnAllRelease.UseVisualStyleBackColor = True
        '
        'BtnAllSelect
        '
        Me.BtnAllSelect.Location = New System.Drawing.Point(18, 540)
        Me.BtnAllSelect.Name = "BtnAllSelect"
        Me.BtnAllSelect.Size = New System.Drawing.Size(90, 30)
        Me.BtnAllSelect.TabIndex = 23
        Me.BtnAllSelect.Text = "全選択"
        Me.BtnAllSelect.UseVisualStyleBackColor = True
        '
        'CboCensusNen
        '
        Me.CboCensusNen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboCensusNen.FormattingEnabled = True
        Me.CboCensusNen.Items.AddRange(New Object() {"2020"})
        Me.CboCensusNen.Location = New System.Drawing.Point(114, 92)
        Me.CboCensusNen.Name = "CboCensusNen"
        Me.CboCensusNen.Size = New System.Drawing.Size(72, 20)
        Me.CboCensusNen.TabIndex = 6
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(29, 95)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(79, 12)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "センサス実施年"
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
        Me.dgvList.Location = New System.Drawing.Point(18, 286)
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
        Me.dgvList.Size = New System.Drawing.Size(776, 226)
        Me.dgvList.TabIndex = 22
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 173)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 12)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "経営形態区分"
        '
        'CboKeieiKeitai
        '
        Me.CboKeieiKeitai.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboKeieiKeitai.FormattingEnabled = True
        Me.CboKeieiKeitai.Location = New System.Drawing.Point(114, 170)
        Me.CboKeieiKeitai.Name = "CboKeieiKeitai"
        Me.CboKeieiKeitai.Size = New System.Drawing.Size(150, 20)
        Me.CboKeieiKeitai.TabIndex = 12
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(29, 251)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 12)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "田畑区分"
        '
        'CboTahata
        '
        Me.CboTahata.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboTahata.Enabled = False
        Me.CboTahata.FormattingEnabled = True
        Me.CboTahata.Location = New System.Drawing.Point(114, 248)
        Me.CboTahata.Name = "CboTahata"
        Me.CboTahata.Size = New System.Drawing.Size(150, 20)
        Me.CboTahata.TabIndex = 18
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(29, 225)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 12)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "生産費区分"
        '
        'CboSeisanhi
        '
        Me.CboSeisanhi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboSeisanhi.Enabled = False
        Me.CboSeisanhi.FormattingEnabled = True
        Me.CboSeisanhi.Location = New System.Drawing.Point(114, 222)
        Me.CboSeisanhi.Name = "CboSeisanhi"
        Me.CboSeisanhi.Size = New System.Drawing.Size(150, 20)
        Me.CboSeisanhi.TabIndex = 16
        '
        'BtnCsvOutput
        '
        Me.BtnCsvOutput.Location = New System.Drawing.Point(606, 250)
        Me.BtnCsvOutput.Name = "BtnCsvOutput"
        Me.BtnCsvOutput.Size = New System.Drawing.Size(90, 30)
        Me.BtnCsvOutput.TabIndex = 20
        Me.BtnCsvOutput.Text = "CSV出力"
        Me.BtnCsvOutput.UseVisualStyleBackColor = True
        '
        'TxtTodofuken
        '
        Me.TxtTodofuken.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TxtTodofuken.Location = New System.Drawing.Point(93, 46)
        Me.TxtTodofuken.MaxLength = 50
        Me.TxtTodofuken.Name = "TxtTodofuken"
        Me.TxtTodofuken.Size = New System.Drawing.Size(150, 19)
        Me.TxtTodofuken.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(10, 49)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(69, 12)
        Me.Label7.TabIndex = 1
        Me.Label7.Text = "都道府県CD"
        '
        'TxtShikuchoson
        '
        Me.TxtShikuchoson.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TxtShikuchoson.Location = New System.Drawing.Point(93, 71)
        Me.TxtShikuchoson.MaxLength = 50
        Me.TxtShikuchoson.Name = "TxtShikuchoson"
        Me.TxtShikuchoson.Size = New System.Drawing.Size(150, 19)
        Me.TxtShikuchoson.TabIndex = 4
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(10, 99)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(53, 12)
        Me.Label14.TabIndex = 5
        Me.Label14.Text = "規模階層"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(10, 74)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(69, 12)
        Me.Label15.TabIndex = 3
        Me.Label15.Text = "市区町村CD"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(10, 125)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(53, 12)
        Me.Label16.TabIndex = 7
        Me.Label16.Text = "標本区分"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(10, 153)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(77, 12)
        Me.Label17.TabIndex = 9
        Me.Label17.Text = "標本候補区分"
        '
        'BtnDel
        '
        Me.BtnDel.Location = New System.Drawing.Point(389, 540)
        Me.BtnDel.Name = "BtnDel"
        Me.BtnDel.Size = New System.Drawing.Size(90, 30)
        Me.BtnDel.TabIndex = 26
        Me.BtnDel.Text = "削除"
        Me.BtnDel.UseVisualStyleBackColor = True
        '
        'BtnAdd
        '
        Me.BtnAdd.Location = New System.Drawing.Point(293, 540)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.Size = New System.Drawing.Size(90, 30)
        Me.BtnAdd.TabIndex = 25
        Me.BtnAdd.Text = "追加"
        Me.BtnAdd.UseVisualStyleBackColor = True
        '
        'BtnSave
        '
        Me.BtnSave.Location = New System.Drawing.Point(606, 543)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(90, 30)
        Me.BtnSave.TabIndex = 27
        Me.BtnSave.Text = "保存"
        Me.BtnSave.UseVisualStyleBackColor = True
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(273, 31)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(53, 12)
        Me.Label19.TabIndex = 11
        Me.Label19.Text = "一連番号"
        '
        'TxtIchirenNo
        '
        Me.TxtIchirenNo.AcceptsReturn = True
        Me.TxtIchirenNo.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TxtIchirenNo.Location = New System.Drawing.Point(275, 46)
        Me.TxtIchirenNo.MaxLength = 170
        Me.TxtIchirenNo.Multiline = True
        Me.TxtIchirenNo.Name = "TxtIchirenNo"
        Me.TxtIchirenNo.Size = New System.Drawing.Size(129, 122)
        Me.TxtIchirenNo.TabIndex = 12
        '
        'CboKiboKaiso
        '
        Me.CboKiboKaiso.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboKiboKaiso.FormattingEnabled = True
        Me.CboKiboKaiso.Location = New System.Drawing.Point(93, 96)
        Me.CboKiboKaiso.Name = "CboKiboKaiso"
        Me.CboKiboKaiso.Size = New System.Drawing.Size(78, 20)
        Me.CboKiboKaiso.TabIndex = 6
        '
        'CboHyohon
        '
        Me.CboHyohon.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboHyohon.FormattingEnabled = True
        Me.CboHyohon.Location = New System.Drawing.Point(93, 122)
        Me.CboHyohon.Name = "CboHyohon"
        Me.CboHyohon.Size = New System.Drawing.Size(150, 20)
        Me.CboHyohon.TabIndex = 8
        '
        'CboHyohonKoho
        '
        Me.CboHyohonKoho.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboHyohonKoho.FormattingEnabled = True
        Me.CboHyohonKoho.Location = New System.Drawing.Point(93, 148)
        Me.CboHyohonKoho.Name = "CboHyohonKoho"
        Me.CboHyohonKoho.Size = New System.Drawing.Size(150, 20)
        Me.CboHyohonKoho.TabIndex = 10
        '
        'BtnClear
        '
        Me.BtnClear.Location = New System.Drawing.Point(12, 18)
        Me.BtnClear.Name = "BtnClear"
        Me.BtnClear.Size = New System.Drawing.Size(90, 25)
        Me.BtnClear.TabIndex = 0
        Me.BtnClear.Text = "クリア"
        Me.BtnClear.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.BtnClear)
        Me.GroupBox1.Controls.Add(Me.TxtTodofuken)
        Me.GroupBox1.Controls.Add(Me.CboHyohonKoho)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.CboHyohon)
        Me.GroupBox1.Controls.Add(Me.TxtShikuchoson)
        Me.GroupBox1.Controls.Add(Me.CboKiboKaiso)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.TxtIchirenNo)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Location = New System.Drawing.Point(346, 69)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(444, 175)
        Me.GroupBox1.TabIndex = 19
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "抽出条件"
        '
        'BRA10210F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 587)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BtnSave)
        Me.Controls.Add(Me.BtnDel)
        Me.Controls.Add(Me.BtnAdd)
        Me.Controls.Add(Me.BtnCsvOutput)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.CboTahata)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.CboSeisanhi)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CboKeieiKeitai)
        Me.Controls.Add(Me.dgvList)
        Me.Controls.Add(Me.CboCensusNen)
        Me.Controls.Add(Me.BtnAllRelease)
        Me.Controls.Add(Me.BtnShow)
        Me.Controls.Add(Me.BtnAllSelect)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.lblKyoten)
        Me.Controls.Add(Me.CboKyoten)
        Me.Controls.Add(Me.CboKyoku)
        Me.Controls.Add(Me.lblKyoku)
        Me.Controls.Add(Me.CboEinoRuikei)
        Me.Controls.Add(Me.lblEinouRuikei)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
        Me.Margin = New System.Windows.Forms.Padding(5)
        Me.MinimumSize = New System.Drawing.Size(824, 626)
        Me.Name = "BRA10210F"
        Me.Padding = New System.Windows.Forms.Padding(15, 14, 15, 14)
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.Text = "農業経営統計調査 - 本省工程 - 標本リスト管理"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.lblEinouRuikei, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.CboEinoRuikei, 0)
        Me.Controls.SetChildIndex(Me.lblKyoku, 0)
        Me.Controls.SetChildIndex(Me.CboKyoku, 0)
        Me.Controls.SetChildIndex(Me.CboKyoten, 0)
        Me.Controls.SetChildIndex(Me.lblKyoten, 0)
        Me.Controls.SetChildIndex(Me.Label10, 0)
        Me.Controls.SetChildIndex(Me.BtnAllSelect, 0)
        Me.Controls.SetChildIndex(Me.BtnShow, 0)
        Me.Controls.SetChildIndex(Me.BtnAllRelease, 0)
        Me.Controls.SetChildIndex(Me.CboCensusNen, 0)
        Me.Controls.SetChildIndex(Me.dgvList, 0)
        Me.Controls.SetChildIndex(Me.CboKeieiKeitai, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.CboSeisanhi, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.CboTahata, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.BtnCsvOutput, 0)
        Me.Controls.SetChildIndex(Me.BtnAdd, 0)
        Me.Controls.SetChildIndex(Me.BtnDel, 0)
        Me.Controls.SetChildIndex(Me.BtnSave, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnAllRelease As System.Windows.Forms.Button
    Friend WithEvents BtnAllSelect As System.Windows.Forms.Button
    Friend WithEvents lblKyoten As System.Windows.Forms.Label
    Protected WithEvents CboKyoten As System.Windows.Forms.ComboBox
    Friend WithEvents lblKyoku As System.Windows.Forms.Label
    Protected WithEvents CboKyoku As System.Windows.Forms.ComboBox
    Friend WithEvents BtnShow As System.Windows.Forms.Button
    Protected WithEvents CboEinoRuikei As System.Windows.Forms.ComboBox
    Friend WithEvents lblEinouRuikei As System.Windows.Forms.Label
    Protected WithEvents CboCensusNen As ComboBox
    Friend WithEvents Label10 As Label
    Protected WithEvents dgvList As DataGridView
    Friend WithEvents Label1 As Label
    Protected WithEvents CboKeieiKeitai As ComboBox
    Friend WithEvents Label6 As Label
    Protected WithEvents CboTahata As ComboBox
    Friend WithEvents Label5 As Label
    Protected WithEvents CboSeisanhi As ComboBox
    Friend WithEvents BtnCsvOutput As Button
    Friend WithEvents TxtTodofuken As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents TxtShikuchoson As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents Label17 As Label
    Friend WithEvents BtnDel As Button
    Friend WithEvents BtnAdd As Button
    Friend WithEvents BtnSave As Button
    Friend WithEvents Label19 As Label
    Friend WithEvents TxtIchirenNo As TextBox
    Protected WithEvents CboKiboKaiso As ComboBox
    Protected WithEvents CboHyohon As ComboBox
    Protected WithEvents CboHyohonKoho As ComboBox
    Friend WithEvents BtnClear As Button
    Friend WithEvents GroupBox1 As GroupBox
End Class
