<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA01XXF
    Inherits BRA.HeaderBaseForm

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
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.btnMainteDB = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cboChosakubun = New System.Windows.Forms.ComboBox()
        Me.cboKubun2 = New System.Windows.Forms.ComboBox()
        Me.cboKubun1 = New System.Windows.Forms.ComboBox()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblKoutei
        '
        Me.lblKoutei.Location = New System.Drawing.Point(493, 8)
        Me.lblKoutei.TabIndex = 3
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
        Me.btnReturn.Location = New System.Drawing.Point(530, 448)
        Me.btnReturn.TabIndex = 6
        '
        'lblTitle
        '
        Me.lblTitle.Location = New System.Drawing.Point(14, 9)
        '
        'lblSyori
        '
        Me.lblSyori.Location = New System.Drawing.Point(14, 43)
        Me.lblSyori.TabIndex = 2
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.btnMainteDB)
        Me.GroupBox3.Location = New System.Drawing.Point(90, 339)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(454, 86)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "メンテナンス"
        '
        'btnMainteDB
        '
        Me.btnMainteDB.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnMainteDB.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnMainteDB.Location = New System.Drawing.Point(28, 33)
        Me.btnMainteDB.Name = "btnMainteDB"
        Me.btnMainteDB.Size = New System.Drawing.Size(172, 28)
        Me.btnMainteDB.TabIndex = 0
        Me.btnMainteDB.Text = "DB管理"
        Me.btnMainteDB.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnSelect)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.cboChosakubun)
        Me.GroupBox1.Controls.Add(Me.cboKubun2)
        Me.GroupBox1.Controls.Add(Me.cboKubun1)
        Me.GroupBox1.Location = New System.Drawing.Point(90, 125)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(454, 199)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "実施業務"
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(346, 156)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(90, 30)
        Me.btnSelect.TabIndex = 6
        Me.btnSelect.Text = "選択"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(47, 124)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "調査区分"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(47, 78)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 12)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "区分２"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(47, 34)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(37, 12)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "区分１"
        '
        'cboChosakubun
        '
        Me.cboChosakubun.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChosakubun.FormattingEnabled = True
        Me.cboChosakubun.Location = New System.Drawing.Point(123, 121)
        Me.cboChosakubun.Name = "cboChosakubun"
        Me.cboChosakubun.Size = New System.Drawing.Size(277, 20)
        Me.cboChosakubun.TabIndex = 5
        '
        'cboKubun2
        '
        Me.cboKubun2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKubun2.FormattingEnabled = True
        Me.cboKubun2.Location = New System.Drawing.Point(123, 75)
        Me.cboKubun2.Name = "cboKubun2"
        Me.cboKubun2.Size = New System.Drawing.Size(277, 20)
        Me.cboKubun2.TabIndex = 3
        '
        'cboKubun1
        '
        Me.cboKubun1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKubun1.FormattingEnabled = True
        Me.cboKubun1.Location = New System.Drawing.Point(123, 31)
        Me.cboKubun1.Name = "cboKubun1"
        Me.cboKubun1.Size = New System.Drawing.Size(277, 20)
        Me.cboKubun1.TabIndex = 1
        '
        'BRA01XXF
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(634, 492)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "BRA01XXF"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox3, 0)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Protected WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Protected Friend WithEvents btnMainteDB As System.Windows.Forms.Button
    Protected Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Protected WithEvents cboChosakubun As System.Windows.Forms.ComboBox
    Protected WithEvents cboKubun2 As System.Windows.Forms.ComboBox
    Protected WithEvents cboKubun1 As System.Windows.Forms.ComboBox
    Protected WithEvents GroupBox1 As System.Windows.Forms.GroupBox

End Class
