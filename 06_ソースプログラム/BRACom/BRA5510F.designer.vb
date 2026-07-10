<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA5510F
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
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.dgcSelect = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.ShukeiNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcShukeiName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcUpdateDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgcShukeiConditions = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.局 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.事務所 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.拠点 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnShow = New System.Windows.Forms.Button()
        Me.cboEinouKeieitai = New System.Windows.Forms.ComboBox()
        Me.lblEinouKeieitai = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cboChosaNen = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtParameters = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtLayout = New System.Windows.Forms.TextBox()
        Me.btnParameters = New System.Windows.Forms.Button()
        Me.btnLayout = New System.Windows.Forms.Button()
        Me.chkReplace = New System.Windows.Forms.CheckBox()
        Me.lblSyukeiName = New System.Windows.Forms.Label()
        Me.lblSyukeiNo = New System.Windows.Forms.Label()
        Me.txtSyukeiNameComp = New System.Windows.Forms.TextBox()
        Me.txtSyukeiNoComp = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtSyukeiNameThis = New System.Windows.Forms.TextBox()
        Me.txtSyukeiNoThis = New System.Windows.Forms.TextBox()
        Me.btnSetThis = New System.Windows.Forms.Button()
        Me.btnSetComp = New System.Windows.Forms.Button()
        Me.txtYearThis = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtYearComp = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
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
        Me.btnReturn.Location = New System.Drawing.Point(704, 547)
        Me.btnReturn.TabIndex = 30
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
        Me.lblSyori.Text = "任意帳票出力"
        '
        'btnOutPut
        '
        Me.btnOutPut.Location = New System.Drawing.Point(608, 547)
        Me.btnOutPut.Name = "btnOutPut"
        Me.btnOutPut.Size = New System.Drawing.Size(90, 30)
        Me.btnOutPut.TabIndex = 29
        Me.btnOutPut.Text = "出力"
        Me.btnOutPut.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dgvList)
        Me.GroupBox1.Controls.Add(Me.btnShow)
        Me.GroupBox1.Controls.Add(Me.cboEinouKeieitai)
        Me.GroupBox1.Controls.Add(Me.lblEinouKeieitai)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.cboChosaNen)
        Me.GroupBox1.Location = New System.Drawing.Point(27, 88)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(754, 313)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
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
        Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgcSelect, Me.ShukeiNo, Me.dgcShukeiName, Me.dgcUpdateDate, Me.dgcShukeiConditions, Me.局, Me.事務所, Me.拠点})
        Me.dgvList.Location = New System.Drawing.Point(21, 49)
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
        Me.dgvList.Size = New System.Drawing.Size(712, 254)
        Me.dgvList.TabIndex = 5
        '
        'dgcSelect
        '
        Me.dgcSelect.DataPropertyName = "チェックボックス"
        Me.dgcSelect.HeaderText = ""
        Me.dgcSelect.Name = "dgcSelect"
        Me.dgcSelect.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgcSelect.Width = 30
        '
        'ShukeiNo
        '
        Me.ShukeiNo.HeaderText = "集計番号"
        Me.ShukeiNo.MaxInputLength = 6
        Me.ShukeiNo.Name = "ShukeiNo"
        Me.ShukeiNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'dgcShukeiName
        '
        Me.dgcShukeiName.HeaderText = "集計名称"
        Me.dgcShukeiName.MaxInputLength = 0
        Me.dgcShukeiName.Name = "dgcShukeiName"
        Me.dgcShukeiName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.dgcShukeiName.Width = 400
        '
        'dgcUpdateDate
        '
        Me.dgcUpdateDate.HeaderText = "更新日時"
        Me.dgcUpdateDate.MaxInputLength = 8
        Me.dgcUpdateDate.Name = "dgcUpdateDate"
        Me.dgcUpdateDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'dgcShukeiConditions
        '
        Me.dgcShukeiConditions.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
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
        'cboChosaNen
        '
        Me.cboChosaNen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChosaNen.FormattingEnabled = True
        Me.cboChosaNen.Location = New System.Drawing.Point(107, 18)
        Me.cboChosaNen.Name = "cboChosaNen"
        Me.cboChosaNen.Size = New System.Drawing.Size(72, 20)
        Me.cboChosaNen.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(28, 482)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(85, 12)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "再編成パラメータ"
        '
        'txtParameters
        '
        Me.txtParameters.Location = New System.Drawing.Point(119, 479)
        Me.txtParameters.MaxLength = 50
        Me.txtParameters.Name = "txtParameters"
        Me.txtParameters.ReadOnly = True
        Me.txtParameters.Size = New System.Drawing.Size(484, 19)
        Me.txtParameters.TabIndex = 23
        Me.txtParameters.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(40, 516)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 12)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "帳票レイアウト"
        '
        'txtLayout
        '
        Me.txtLayout.Location = New System.Drawing.Point(119, 513)
        Me.txtLayout.MaxLength = 50
        Me.txtLayout.Name = "txtLayout"
        Me.txtLayout.ReadOnly = True
        Me.txtLayout.Size = New System.Drawing.Size(484, 19)
        Me.txtLayout.TabIndex = 26
        Me.txtLayout.TabStop = False
        '
        'btnParameters
        '
        Me.btnParameters.Location = New System.Drawing.Point(608, 473)
        Me.btnParameters.Name = "btnParameters"
        Me.btnParameters.Size = New System.Drawing.Size(90, 30)
        Me.btnParameters.TabIndex = 24
        Me.btnParameters.Text = "参照"
        Me.btnParameters.UseVisualStyleBackColor = True
        '
        'btnLayout
        '
        Me.btnLayout.Location = New System.Drawing.Point(608, 507)
        Me.btnLayout.Name = "btnLayout"
        Me.btnLayout.Size = New System.Drawing.Size(90, 30)
        Me.btnLayout.TabIndex = 27
        Me.btnLayout.Text = "参照"
        Me.btnLayout.UseVisualStyleBackColor = True
        '
        'chkReplace
        '
        Me.chkReplace.AutoSize = True
        Me.chkReplace.Location = New System.Drawing.Point(30, 555)
        Me.chkReplace.Name = "chkReplace"
        Me.chkReplace.Size = New System.Drawing.Size(134, 16)
        Me.chkReplace.TabIndex = 28
        Me.chkReplace.Text = "３経営体未満""x""置換"
        Me.chkReplace.UseVisualStyleBackColor = True
        '
        'lblSyukeiName
        '
        Me.lblSyukeiName.AutoSize = True
        Me.lblSyukeiName.Location = New System.Drawing.Point(301, 449)
        Me.lblSyukeiName.Name = "lblSyukeiName"
        Me.lblSyukeiName.Size = New System.Drawing.Size(53, 12)
        Me.lblSyukeiName.TabIndex = 19
        Me.lblSyukeiName.Text = "集計名称"
        '
        'lblSyukeiNo
        '
        Me.lblSyukeiNo.AutoSize = True
        Me.lblSyukeiNo.Location = New System.Drawing.Point(180, 449)
        Me.lblSyukeiNo.Name = "lblSyukeiNo"
        Me.lblSyukeiNo.Size = New System.Drawing.Size(53, 12)
        Me.lblSyukeiNo.TabIndex = 17
        Me.lblSyukeiNo.Text = "集計番号"
        '
        'txtSyukeiNameComp
        '
        Me.txtSyukeiNameComp.Location = New System.Drawing.Point(358, 446)
        Me.txtSyukeiNameComp.Name = "txtSyukeiNameComp"
        Me.txtSyukeiNameComp.ReadOnly = True
        Me.txtSyukeiNameComp.Size = New System.Drawing.Size(245, 19)
        Me.txtSyukeiNameComp.TabIndex = 20
        Me.txtSyukeiNameComp.TabStop = False
        '
        'txtSyukeiNoComp
        '
        Me.txtSyukeiNoComp.Location = New System.Drawing.Point(239, 446)
        Me.txtSyukeiNoComp.Name = "txtSyukeiNoComp"
        Me.txtSyukeiNoComp.ReadOnly = True
        Me.txtSyukeiNoComp.Size = New System.Drawing.Size(56, 19)
        Me.txtSyukeiNoComp.TabIndex = 18
        Me.txtSyukeiNoComp.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(28, 449)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 12)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "比較"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(28, 416)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(29, 12)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "本年"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(301, 416)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 12)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "集計名称"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(180, 416)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 12)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "集計番号"
        '
        'txtSyukeiNameThis
        '
        Me.txtSyukeiNameThis.Location = New System.Drawing.Point(358, 413)
        Me.txtSyukeiNameThis.Name = "txtSyukeiNameThis"
        Me.txtSyukeiNameThis.ReadOnly = True
        Me.txtSyukeiNameThis.Size = New System.Drawing.Size(245, 19)
        Me.txtSyukeiNameThis.TabIndex = 12
        Me.txtSyukeiNameThis.TabStop = False
        '
        'txtSyukeiNoThis
        '
        Me.txtSyukeiNoThis.Location = New System.Drawing.Point(239, 413)
        Me.txtSyukeiNoThis.Name = "txtSyukeiNoThis"
        Me.txtSyukeiNoThis.ReadOnly = True
        Me.txtSyukeiNoThis.Size = New System.Drawing.Size(56, 19)
        Me.txtSyukeiNoThis.TabIndex = 10
        Me.txtSyukeiNoThis.TabStop = False
        '
        'btnSetThis
        '
        Me.btnSetThis.Location = New System.Drawing.Point(608, 407)
        Me.btnSetThis.Name = "btnSetThis"
        Me.btnSetThis.Size = New System.Drawing.Size(90, 30)
        Me.btnSetThis.TabIndex = 13
        Me.btnSetThis.Text = "設定"
        Me.btnSetThis.UseVisualStyleBackColor = True
        '
        'btnSetComp
        '
        Me.btnSetComp.Location = New System.Drawing.Point(608, 440)
        Me.btnSetComp.Name = "btnSetComp"
        Me.btnSetComp.Size = New System.Drawing.Size(90, 30)
        Me.btnSetComp.TabIndex = 21
        Me.btnSetComp.Text = "設定"
        Me.btnSetComp.UseVisualStyleBackColor = True
        '
        'txtYearThis
        '
        Me.txtYearThis.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtYearThis.Location = New System.Drawing.Point(134, 413)
        Me.txtYearThis.MaxLength = 4
        Me.txtYearThis.Name = "txtYearThis"
        Me.txtYearThis.ReadOnly = True
        Me.txtYearThis.Size = New System.Drawing.Size(35, 19)
        Me.txtYearThis.TabIndex = 8
        Me.txtYearThis.TabStop = False
        Me.txtYearThis.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(63, 416)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(65, 12)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "調査年（産）"
        '
        'txtYearComp
        '
        Me.txtYearComp.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtYearComp.Location = New System.Drawing.Point(134, 446)
        Me.txtYearComp.MaxLength = 4
        Me.txtYearComp.Name = "txtYearComp"
        Me.txtYearComp.ReadOnly = True
        Me.txtYearComp.Size = New System.Drawing.Size(35, 19)
        Me.txtYearComp.TabIndex = 16
        Me.txtYearComp.TabStop = False
        Me.txtYearComp.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(63, 449)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(65, 12)
        Me.Label8.TabIndex = 15
        Me.Label8.Text = "調査年（産）"
        '
        'BRA5510F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 594)
        Me.Controls.Add(Me.txtYearComp)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtYearThis)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btnSetComp)
        Me.Controls.Add(Me.btnSetThis)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtSyukeiNameThis)
        Me.Controls.Add(Me.txtSyukeiNoThis)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblSyukeiName)
        Me.Controls.Add(Me.lblSyukeiNo)
        Me.Controls.Add(Me.txtSyukeiNameComp)
        Me.Controls.Add(Me.txtSyukeiNoComp)
        Me.Controls.Add(Me.chkReplace)
        Me.Controls.Add(Me.btnLayout)
        Me.Controls.Add(Me.btnParameters)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtLayout)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtParameters)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnOutPut)
        Me.Name = "BRA5510F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.btnOutPut, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.txtParameters, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.txtLayout, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.btnParameters, 0)
        Me.Controls.SetChildIndex(Me.btnLayout, 0)
        Me.Controls.SetChildIndex(Me.chkReplace, 0)
        Me.Controls.SetChildIndex(Me.txtSyukeiNoComp, 0)
        Me.Controls.SetChildIndex(Me.txtSyukeiNameComp, 0)
        Me.Controls.SetChildIndex(Me.lblSyukeiNo, 0)
        Me.Controls.SetChildIndex(Me.lblSyukeiName, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.txtSyukeiNoThis, 0)
        Me.Controls.SetChildIndex(Me.txtSyukeiNameThis, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.btnSetThis, 0)
        Me.Controls.SetChildIndex(Me.btnSetComp, 0)
        Me.Controls.SetChildIndex(Me.Label7, 0)
        Me.Controls.SetChildIndex(Me.txtYearThis, 0)
        Me.Controls.SetChildIndex(Me.Label8, 0)
        Me.Controls.SetChildIndex(Me.txtYearComp, 0)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOutPut As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblEinouKeieitai As System.Windows.Forms.Label
    Protected WithEvents cboEinouKeieitai As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Protected WithEvents cboChosaNen As System.Windows.Forms.ComboBox
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Protected WithEvents dgvList As System.Windows.Forms.DataGridView
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtParameters As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtLayout As System.Windows.Forms.TextBox
    Friend WithEvents btnParameters As System.Windows.Forms.Button
    Friend WithEvents btnLayout As System.Windows.Forms.Button
    Friend WithEvents chkReplace As System.Windows.Forms.CheckBox
    Friend WithEvents lblSyukeiName As System.Windows.Forms.Label
    Friend WithEvents lblSyukeiNo As System.Windows.Forms.Label
    Friend WithEvents txtSyukeiNameComp As System.Windows.Forms.TextBox
    Friend WithEvents txtSyukeiNoComp As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtSyukeiNameThis As System.Windows.Forms.TextBox
    Friend WithEvents txtSyukeiNoThis As System.Windows.Forms.TextBox
    Friend WithEvents btnSetThis As System.Windows.Forms.Button
    Friend WithEvents btnSetComp As System.Windows.Forms.Button
    Friend WithEvents txtYearThis As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtYearComp As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dgcSelect As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ShukeiNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcShukeiName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcUpdateDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcShukeiConditions As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 局 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 事務所 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 拠点 As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
