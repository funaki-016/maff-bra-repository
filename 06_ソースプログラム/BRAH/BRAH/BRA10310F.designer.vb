<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BRA10310F
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.BtnShow = New System.Windows.Forms.Button()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.BtnSave = New System.Windows.Forms.Button()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.CboAoiro = New System.Windows.Forms.ComboBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.CboShurakuEino = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.CboShufukugyo = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.CboTahata = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CboSeisanhi = New System.Windows.Forms.ComboBox()
        Me.CboEinoruikei = New System.Windows.Forms.ComboBox()
        Me.CboKeieiKeitai = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LblTarget = New System.Windows.Forms.Label()
        Me.BtnClear = New System.Windows.Forms.Button()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.btnReturn.Location = New System.Drawing.Point(704, 593)
        Me.btnReturn.Margin = New System.Windows.Forms.Padding(4)
        Me.btnReturn.TabIndex = 32
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "抽出条件設定"
        '
        'BtnShow
        '
        Me.BtnShow.Location = New System.Drawing.Point(704, 267)
        Me.BtnShow.Name = "BtnShow"
        Me.BtnShow.Size = New System.Drawing.Size(90, 30)
        Me.BtnShow.TabIndex = 26
        Me.BtnShow.Text = "表示"
        Me.BtnShow.UseVisualStyleBackColor = True
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
        Me.dgvList.ColumnHeadersVisible = False
        Me.dgvList.Location = New System.Drawing.Point(18, 303)
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
        Me.dgvList.Size = New System.Drawing.Size(776, 283)
        Me.dgvList.TabIndex = 29
        '
        'BtnSave
        '
        Me.BtnSave.Location = New System.Drawing.Point(607, 594)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(90, 30)
        Me.BtnSave.TabIndex = 31
        Me.BtnSave.Text = "保存"
        Me.BtnSave.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(429, 212)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(116, 12)
        Me.Label17.TabIndex = 21
        Me.Label17.Text = "（営農個人・副業のみ）"
        '
        'CboAoiro
        '
        Me.CboAoiro.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboAoiro.Enabled = False
        Me.CboAoiro.FormattingEnabled = True
        Me.CboAoiro.Location = New System.Drawing.Point(515, 189)
        Me.CboAoiro.Name = "CboAoiro"
        Me.CboAoiro.Size = New System.Drawing.Size(150, 20)
        Me.CboAoiro.TabIndex = 22
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(429, 192)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(77, 12)
        Me.Label18.TabIndex = 20
        Me.Label18.Text = "青色申告区分"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(429, 166)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 12)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "（小麦、大豆のみ）"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(429, 123)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(190, 12)
        Me.Label16.TabIndex = 9
        Me.Label16.Text = "（個別経営体、組織法人経営体のみ）"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(138, 252)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(122, 12)
        Me.Label15.TabIndex = 24
        Me.Label15.Text = "（営農法人水田作のみ）"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(138, 209)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(86, 12)
        Me.Label14.TabIndex = 18
        Me.Label14.Text = "（営農個人のみ）"
        '
        'CboShurakuEino
        '
        Me.CboShurakuEino.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboShurakuEino.Enabled = False
        Me.CboShurakuEino.FormattingEnabled = True
        Me.CboShurakuEino.Location = New System.Drawing.Point(223, 229)
        Me.CboShurakuEino.Name = "CboShurakuEino"
        Me.CboShurakuEino.Size = New System.Drawing.Size(150, 20)
        Me.CboShurakuEino.TabIndex = 25
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(136, 232)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(77, 12)
        Me.Label13.TabIndex = 23
        Me.Label13.Text = "集落営農区分"
        '
        'CboShufukugyo
        '
        Me.CboShufukugyo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboShufukugyo.Enabled = False
        Me.CboShufukugyo.FormattingEnabled = True
        Me.CboShufukugyo.Location = New System.Drawing.Point(223, 186)
        Me.CboShufukugyo.Name = "CboShufukugyo"
        Me.CboShufukugyo.Size = New System.Drawing.Size(150, 20)
        Me.CboShufukugyo.TabIndex = 19
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(138, 189)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(65, 12)
        Me.Label12.TabIndex = 17
        Me.Label12.Text = "主副業区分"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(138, 123)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(41, 12)
        Me.Label11.TabIndex = 6
        Me.Label11.Text = "（必須）"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(138, 166)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(166, 12)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "（個人経営体、法人経営体のみ）"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(429, 146)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 12)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "田畑区分"
        '
        'CboTahata
        '
        Me.CboTahata.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboTahata.Enabled = False
        Me.CboTahata.FormattingEnabled = True
        Me.CboTahata.Location = New System.Drawing.Point(515, 143)
        Me.CboTahata.Name = "CboTahata"
        Me.CboTahata.Size = New System.Drawing.Size(150, 20)
        Me.CboTahata.TabIndex = 16
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(429, 103)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 12)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "生産費区分"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(136, 146)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 12)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "営農類型区分"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(136, 103)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 12)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "経営形態区分" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'CboSeisanhi
        '
        Me.CboSeisanhi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboSeisanhi.FormattingEnabled = True
        Me.CboSeisanhi.Items.AddRange(New Object() {"米生産費"})
        Me.CboSeisanhi.Location = New System.Drawing.Point(515, 100)
        Me.CboSeisanhi.Name = "CboSeisanhi"
        Me.CboSeisanhi.Size = New System.Drawing.Size(150, 20)
        Me.CboSeisanhi.TabIndex = 10
        '
        'CboEinoruikei
        '
        Me.CboEinoruikei.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboEinoruikei.Enabled = False
        Me.CboEinoruikei.FormattingEnabled = True
        Me.CboEinoruikei.Location = New System.Drawing.Point(223, 143)
        Me.CboEinoruikei.Name = "CboEinoruikei"
        Me.CboEinoruikei.Size = New System.Drawing.Size(150, 20)
        Me.CboEinoruikei.TabIndex = 13
        '
        'CboKeieiKeitai
        '
        Me.CboKeieiKeitai.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboKeieiKeitai.FormattingEnabled = True
        Me.CboKeieiKeitai.Items.AddRange(New Object() {"個別経営体"})
        Me.CboKeieiKeitai.Location = New System.Drawing.Point(223, 100)
        Me.CboKeieiKeitai.Name = "CboKeieiKeitai"
        Me.CboKeieiKeitai.Size = New System.Drawing.Size(150, 20)
        Me.CboKeieiKeitai.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(18, 288)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(59, 12)
        Me.Label3.TabIndex = 27
        Me.Label3.Text = "設定対象："
        '
        'LblTarget
        '
        Me.LblTarget.AutoSize = True
        Me.LblTarget.Location = New System.Drawing.Point(83, 288)
        Me.LblTarget.Name = "LblTarget"
        Me.LblTarget.Size = New System.Drawing.Size(117, 12)
        Me.LblTarget.TabIndex = 28
        Me.LblTarget.Text = "個別経営体 米生産費"
        '
        'BtnClear
        '
        Me.BtnClear.Location = New System.Drawing.Point(18, 594)
        Me.BtnClear.Name = "BtnClear"
        Me.BtnClear.Size = New System.Drawing.Size(90, 30)
        Me.BtnClear.TabIndex = 30
        Me.BtnClear.Text = "クリア"
        Me.BtnClear.UseVisualStyleBackColor = True
        '
        'BRA10310F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 641)
        Me.Controls.Add(Me.BtnClear)
        Me.Controls.Add(Me.LblTarget)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.CboAoiro)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.CboShurakuEino)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.CboShufukugyo)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.CboTahata)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CboSeisanhi)
        Me.Controls.Add(Me.CboEinoruikei)
        Me.Controls.Add(Me.CboKeieiKeitai)
        Me.Controls.Add(Me.BtnSave)
        Me.Controls.Add(Me.dgvList)
        Me.Controls.Add(Me.BtnShow)
        Me.Margin = New System.Windows.Forms.Padding(5)
        Me.Name = "BRA10310F"
        Me.Padding = New System.Windows.Forms.Padding(15, 14, 15, 14)
        Me.Text = "農業経営統計調査 - 本省工程 - 抽出条件設定"
        Me.Controls.SetChildIndex(Me.BtnShow, 0)
        Me.Controls.SetChildIndex(Me.dgvList, 0)
        Me.Controls.SetChildIndex(Me.BtnSave, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.CboKeieiKeitai, 0)
        Me.Controls.SetChildIndex(Me.CboEinoruikei, 0)
        Me.Controls.SetChildIndex(Me.CboSeisanhi, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.CboTahata, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.Label7, 0)
        Me.Controls.SetChildIndex(Me.Label11, 0)
        Me.Controls.SetChildIndex(Me.Label12, 0)
        Me.Controls.SetChildIndex(Me.CboShufukugyo, 0)
        Me.Controls.SetChildIndex(Me.Label13, 0)
        Me.Controls.SetChildIndex(Me.CboShurakuEino, 0)
        Me.Controls.SetChildIndex(Me.Label14, 0)
        Me.Controls.SetChildIndex(Me.Label15, 0)
        Me.Controls.SetChildIndex(Me.Label16, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.Label18, 0)
        Me.Controls.SetChildIndex(Me.CboAoiro, 0)
        Me.Controls.SetChildIndex(Me.Label17, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.LblTarget, 0)
        Me.Controls.SetChildIndex(Me.BtnClear, 0)
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnShow As System.Windows.Forms.Button
    Protected WithEvents dgvList As DataGridView
    Friend WithEvents BtnSave As Button
    Friend WithEvents Label17 As Label
    Protected WithEvents CboAoiro As ComboBox
    Friend WithEvents Label18 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label14 As Label
    Protected WithEvents CboShurakuEino As ComboBox
    Friend WithEvents Label13 As Label
    Protected WithEvents CboShufukugyo As ComboBox
    Friend WithEvents Label12 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Protected WithEvents CboTahata As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Protected WithEvents CboSeisanhi As ComboBox
    Protected WithEvents CboEinoruikei As ComboBox
    Protected WithEvents CboKeieiKeitai As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents LblTarget As Label
    Friend WithEvents BtnClear As Button
End Class
