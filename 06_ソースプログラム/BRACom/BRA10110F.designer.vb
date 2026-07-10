<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BRA10110F
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
        Me.BtnTorikomi = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.CboCensusNen = New System.Windows.Forms.ComboBox()
        Me.CboKeieiKeitai = New System.Windows.Forms.ComboBox()
        Me.CboEinoRuikei = New System.Windows.Forms.ComboBox()
        Me.CboSeisanhi = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.CboTahata = New System.Windows.Forms.ComboBox()
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
        Me.btnReturn.Location = New System.Drawing.Point(704, 237)
        Me.btnReturn.TabIndex = 16
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "標本リスト取込"
        '
        'BtnTorikomi
        '
        Me.BtnTorikomi.Location = New System.Drawing.Point(593, 237)
        Me.BtnTorikomi.Name = "BtnTorikomi"
        Me.BtnTorikomi.Size = New System.Drawing.Size(90, 30)
        Me.BtnTorikomi.TabIndex = 15
        Me.BtnTorikomi.Text = "取込"
        Me.BtnTorikomi.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(138, 97)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(79, 12)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "センサス実施年"
        '
        'CboCensusNen
        '
        Me.CboCensusNen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboCensusNen.FormattingEnabled = True
        Me.CboCensusNen.Location = New System.Drawing.Point(223, 94)
        Me.CboCensusNen.Name = "CboCensusNen"
        Me.CboCensusNen.Size = New System.Drawing.Size(72, 20)
        Me.CboCensusNen.TabIndex = 6
        '
        'CboKeieiKeitai
        '
        Me.CboKeieiKeitai.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboKeieiKeitai.FormattingEnabled = True
        Me.CboKeieiKeitai.Location = New System.Drawing.Point(223, 120)
        Me.CboKeieiKeitai.Name = "CboKeieiKeitai"
        Me.CboKeieiKeitai.Size = New System.Drawing.Size(150, 20)
        Me.CboKeieiKeitai.TabIndex = 8
        '
        'CboEinoRuikei
        '
        Me.CboEinoRuikei.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboEinoRuikei.Enabled = False
        Me.CboEinoRuikei.FormattingEnabled = True
        Me.CboEinoRuikei.Location = New System.Drawing.Point(223, 146)
        Me.CboEinoRuikei.Name = "CboEinoRuikei"
        Me.CboEinoRuikei.Size = New System.Drawing.Size(150, 20)
        Me.CboEinoRuikei.TabIndex = 10
        '
        'CboSeisanhi
        '
        Me.CboSeisanhi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboSeisanhi.Enabled = False
        Me.CboSeisanhi.FormattingEnabled = True
        Me.CboSeisanhi.Location = New System.Drawing.Point(223, 172)
        Me.CboSeisanhi.Name = "CboSeisanhi"
        Me.CboSeisanhi.Size = New System.Drawing.Size(150, 20)
        Me.CboSeisanhi.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(138, 123)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 12)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "経営形態区分"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(138, 149)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 12)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "営農類型区分"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(138, 175)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 12)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "生産費区分"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(138, 201)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 12)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "田畑区分"
        '
        'CboTahata
        '
        Me.CboTahata.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboTahata.Enabled = False
        Me.CboTahata.FormattingEnabled = True
        Me.CboTahata.Location = New System.Drawing.Point(223, 198)
        Me.CboTahata.Name = "CboTahata"
        Me.CboTahata.Size = New System.Drawing.Size(150, 20)
        Me.CboTahata.TabIndex = 14
        '
        'BRA10110F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 281)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.CboTahata)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CboSeisanhi)
        Me.Controls.Add(Me.CboEinoRuikei)
        Me.Controls.Add(Me.CboKeieiKeitai)
        Me.Controls.Add(Me.CboCensusNen)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.BtnTorikomi)
        Me.Name = "BRA10110F"
        Me.Text = "農業経営統計調査 - 本省工程 - 標本リスト取込"
        Me.Controls.SetChildIndex(Me.BtnTorikomi, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.Label10, 0)
        Me.Controls.SetChildIndex(Me.CboCensusNen, 0)
        Me.Controls.SetChildIndex(Me.CboKeieiKeitai, 0)
        Me.Controls.SetChildIndex(Me.CboEinoRuikei, 0)
        Me.Controls.SetChildIndex(Me.CboSeisanhi, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.CboTahata, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnTorikomi As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Protected WithEvents CboCensusNen As ComboBox
    Protected WithEvents CboKeieiKeitai As ComboBox
    Protected WithEvents CboEinoRuikei As ComboBox
    Protected WithEvents CboSeisanhi As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Protected WithEvents CboTahata As ComboBox
End Class
