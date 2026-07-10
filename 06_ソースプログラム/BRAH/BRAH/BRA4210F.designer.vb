<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BRA4210F
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
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.BtnApply = New System.Windows.Forms.Button()
        Me.BtnCalc = New System.Windows.Forms.Button()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.BtnSave = New System.Windows.Forms.Button()
        Me.lblCalcDateTime = New System.Windows.Forms.Label()
        Me.CboChosaNen = New System.Windows.Forms.ComboBox()
        Me.CboEinoruikei = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CboChiiki = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.COL_EINORUIKEI = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_CHIIKI = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_KIBO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_BOSHUDAN = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_SHUKEITAISHO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_BAIRITSU = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_SHURAKUEINO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_CHIIKI_CODE = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblKoutei
        '
        Me.lblKoutei.Location = New System.Drawing.Point(424, 11)
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
        Me.btnReturn.Location = New System.Drawing.Point(458, 613)
        Me.btnReturn.Margin = New System.Windows.Forms.Padding(4)
        Me.btnReturn.TabIndex = 17
        '
        'lblTitle
        '
        Me.lblTitle.Size = New System.Drawing.Size(410, 19)
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.Size = New System.Drawing.Size(410, 19)
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "集計倍率管理"
        '
        'BtnApply
        '
        Me.BtnApply.Enabled = False
        Me.BtnApply.Location = New System.Drawing.Point(163, 613)
        Me.BtnApply.Name = "BtnApply"
        Me.BtnApply.Size = New System.Drawing.Size(120, 30)
        Me.BtnApply.TabIndex = 15
        Me.BtnApply.Text = "営農集計倍率適用"
        Me.BtnApply.UseVisualStyleBackColor = True
        '
        'BtnCalc
        '
        Me.BtnCalc.Enabled = False
        Me.BtnCalc.Location = New System.Drawing.Point(18, 613)
        Me.BtnCalc.Name = "BtnCalc"
        Me.BtnCalc.Size = New System.Drawing.Size(120, 30)
        Me.BtnCalc.TabIndex = 14
        Me.BtnCalc.Text = "営農集計倍率計算"
        Me.BtnCalc.UseVisualStyleBackColor = True
        '
        'dgvList
        '
        Me.dgvList.AllowUserToAddRows = False
        Me.dgvList.AllowUserToDeleteRows = False
        Me.dgvList.AllowUserToResizeColumns = False
        Me.dgvList.AllowUserToResizeRows = False
        Me.dgvList.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvList.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvList.ColumnHeadersHeight = 20
        Me.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.COL_EINORUIKEI, Me.COL_CHIIKI, Me.COL_KIBO, Me.COL_BOSHUDAN, Me.COL_SHUKEITAISHO, Me.COL_BAIRITSU, Me.COL_SHURAKUEINO, Me.COL_CHIIKI_CODE})
        Me.dgvList.Location = New System.Drawing.Point(18, 151)
        Me.dgvList.Name = "dgvList"
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle8.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvList.RowHeadersDefaultCellStyle = DataGridViewCellStyle8
        Me.dgvList.RowHeadersVisible = False
        Me.dgvList.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.dgvList.RowsDefaultCellStyle = DataGridViewCellStyle9
        Me.dgvList.RowTemplate.Height = 21
        Me.dgvList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvList.Size = New System.Drawing.Size(530, 442)
        Me.dgvList.TabIndex = 13
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(255, 80)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(53, 12)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "計算日時"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(18, 80)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(41, 12)
        Me.Label15.TabIndex = 5
        Me.Label15.Text = "調査年"
        '
        'BtnSave
        '
        Me.BtnSave.Enabled = False
        Me.BtnSave.Location = New System.Drawing.Point(311, 613)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(90, 30)
        Me.BtnSave.TabIndex = 16
        Me.BtnSave.Text = "保存"
        Me.BtnSave.UseVisualStyleBackColor = True
        '
        'lblCalcDateTime
        '
        Me.lblCalcDateTime.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCalcDateTime.Location = New System.Drawing.Point(314, 77)
        Me.lblCalcDateTime.Name = "lblCalcDateTime"
        Me.lblCalcDateTime.Size = New System.Drawing.Size(120, 20)
        Me.lblCalcDateTime.TabIndex = 8
        Me.lblCalcDateTime.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CboChosaNen
        '
        Me.CboChosaNen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboChosaNen.FormattingEnabled = True
        Me.CboChosaNen.Location = New System.Drawing.Point(75, 77)
        Me.CboChosaNen.Name = "CboChosaNen"
        Me.CboChosaNen.Size = New System.Drawing.Size(72, 20)
        Me.CboChosaNen.TabIndex = 6
        '
        'CboEinoruikei
        '
        Me.CboEinoruikei.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboEinoruikei.Enabled = False
        Me.CboEinoruikei.FormattingEnabled = True
        Me.CboEinoruikei.Location = New System.Drawing.Point(75, 113)
        Me.CboEinoruikei.Name = "CboEinoruikei"
        Me.CboEinoruikei.Size = New System.Drawing.Size(180, 20)
        Me.CboEinoruikei.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 116)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 12)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "営農類型"
        '
        'CboChiiki
        '
        Me.CboChiiki.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboChiiki.Enabled = False
        Me.CboChiiki.FormattingEnabled = True
        Me.CboChiiki.Location = New System.Drawing.Point(314, 113)
        Me.CboChiiki.Name = "CboChiiki"
        Me.CboChiiki.Size = New System.Drawing.Size(140, 20)
        Me.CboChiiki.TabIndex = 12
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(279, 116)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 12)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "地域"
        '
        'COL_EINORUIKEI
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.ControlLight
        Me.COL_EINORUIKEI.DefaultCellStyle = DataGridViewCellStyle2
        Me.COL_EINORUIKEI.HeaderText = "営農類型"
        Me.COL_EINORUIKEI.Name = "COL_EINORUIKEI"
        Me.COL_EINORUIKEI.ReadOnly = True
        Me.COL_EINORUIKEI.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COL_EINORUIKEI.Width = 70
        '
        'COL_CHIIKI
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.ControlLight
        Me.COL_CHIIKI.DefaultCellStyle = DataGridViewCellStyle3
        Me.COL_CHIIKI.HeaderText = "地域"
        Me.COL_CHIIKI.Name = "COL_CHIIKI"
        Me.COL_CHIIKI.ReadOnly = True
        Me.COL_CHIIKI.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COL_CHIIKI.Width = 90
        '
        'COL_KIBO
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.ControlLight
        Me.COL_KIBO.DefaultCellStyle = DataGridViewCellStyle4
        Me.COL_KIBO.HeaderText = "規模区分"
        Me.COL_KIBO.Name = "COL_KIBO"
        Me.COL_KIBO.ReadOnly = True
        Me.COL_KIBO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COL_KIBO.Width = 70
        '
        'COL_BOSHUDAN
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.NullValue = Nothing
        Me.COL_BOSHUDAN.DefaultCellStyle = DataGridViewCellStyle5
        Me.COL_BOSHUDAN.HeaderText = "母集団"
        Me.COL_BOSHUDAN.MaxInputLength = 9
        Me.COL_BOSHUDAN.Name = "COL_BOSHUDAN"
        Me.COL_BOSHUDAN.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'COL_SHUKEITAISHO
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.ControlLight
        DataGridViewCellStyle6.NullValue = Nothing
        Me.COL_SHUKEITAISHO.DefaultCellStyle = DataGridViewCellStyle6
        Me.COL_SHUKEITAISHO.HeaderText = "集計対象数"
        Me.COL_SHUKEITAISHO.Name = "COL_SHUKEITAISHO"
        Me.COL_SHUKEITAISHO.ReadOnly = True
        Me.COL_SHUKEITAISHO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COL_SHUKEITAISHO.Width = 90
        '
        'COL_BAIRITSU
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.ControlLight
        DataGridViewCellStyle7.NullValue = Nothing
        Me.COL_BAIRITSU.DefaultCellStyle = DataGridViewCellStyle7
        Me.COL_BAIRITSU.HeaderText = "営農集計倍率"
        Me.COL_BAIRITSU.Name = "COL_BAIRITSU"
        Me.COL_BAIRITSU.ReadOnly = True
        Me.COL_BAIRITSU.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COL_BAIRITSU.Width = 90
        '
        'COL_SHURAKUEINO
        '
        Me.COL_SHURAKUEINO.HeaderText = "非表示＿集落営農区分"
        Me.COL_SHURAKUEINO.Name = "COL_SHURAKUEINO"
        Me.COL_SHURAKUEINO.Visible = False
        '
        'COL_CHIIKI_CODE
        '
        Me.COL_CHIIKI_CODE.HeaderText = "非表示＿地域区分"
        Me.COL_CHIIKI_CODE.Name = "COL_CHIIKI_CODE"
        Me.COL_CHIIKI_CODE.Visible = False
        '
        'BRA4210F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(569, 661)
        Me.Controls.Add(Me.CboChiiki)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CboEinoruikei)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CboChosaNen)
        Me.Controls.Add(Me.lblCalcDateTime)
        Me.Controls.Add(Me.BtnSave)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.dgvList)
        Me.Controls.Add(Me.BtnApply)
        Me.Controls.Add(Me.BtnCalc)
        Me.Controls.Add(Me.Label15)
        Me.Margin = New System.Windows.Forms.Padding(5)
        Me.Name = "BRA4210F"
        Me.Padding = New System.Windows.Forms.Padding(15, 14, 15, 14)
        Me.Text = "農業経営統計調査 - 本省工程 - 米在庫リンケージCSV出力"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.Label15, 0)
        Me.Controls.SetChildIndex(Me.BtnCalc, 0)
        Me.Controls.SetChildIndex(Me.BtnApply, 0)
        Me.Controls.SetChildIndex(Me.dgvList, 0)
        Me.Controls.SetChildIndex(Me.Label7, 0)
        Me.Controls.SetChildIndex(Me.BtnSave, 0)
        Me.Controls.SetChildIndex(Me.lblCalcDateTime, 0)
        Me.Controls.SetChildIndex(Me.CboChosaNen, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.CboEinoruikei, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.CboChiiki, 0)
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnApply As System.Windows.Forms.Button
    Friend WithEvents BtnCalc As System.Windows.Forms.Button
    Protected WithEvents dgvList As DataGridView
    Friend WithEvents Label7 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents BtnSave As Button
    Friend WithEvents lblCalcDateTime As Label
    Protected WithEvents CboChosaNen As ComboBox
    Protected WithEvents CboEinoruikei As ComboBox
    Friend WithEvents Label1 As Label
    Protected WithEvents CboChiiki As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents COL_EINORUIKEI As DataGridViewTextBoxColumn
    Friend WithEvents COL_CHIIKI As DataGridViewTextBoxColumn
    Friend WithEvents COL_KIBO As DataGridViewTextBoxColumn
    Friend WithEvents COL_BOSHUDAN As DataGridViewTextBoxColumn
    Friend WithEvents COL_SHUKEITAISHO As DataGridViewTextBoxColumn
    Friend WithEvents COL_BAIRITSU As DataGridViewTextBoxColumn
    Friend WithEvents COL_SHURAKUEINO As DataGridViewTextBoxColumn
    Friend WithEvents COL_CHIIKI_CODE As DataGridViewTextBoxColumn
End Class
