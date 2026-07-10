<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BRA10410F
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
        Me.BtnPrint = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.CboCensusNen = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CboKyokuTo = New System.Windows.Forms.ComboBox()
        Me.CboKeieiKeitai = New System.Windows.Forms.ComboBox()
        Me.CboEinoruikei = New System.Windows.Forms.ComboBox()
        Me.CboSeisanhi = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.CboTahata = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.CboShufukugyo = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.CboShurakuEino = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.CboAoiro = New System.Windows.Forms.ComboBox()
        Me.Label18 = New System.Windows.Forms.Label()
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
        Me.btnReturn.Location = New System.Drawing.Point(704, 329)
        Me.btnReturn.TabIndex = 33
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "規模区分別期待標本及び予定経営体数出力"
        '
        'BtnPrint
        '
        Me.BtnPrint.Location = New System.Drawing.Point(608, 329)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(90, 30)
        Me.BtnPrint.TabIndex = 32
        Me.BtnPrint.Text = "様式出力"
        Me.BtnPrint.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(136, 100)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(79, 12)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "センサス実施年"
        '
        'CboCensusNen
        '
        Me.CboCensusNen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboCensusNen.FormattingEnabled = True
        Me.CboCensusNen.Items.AddRange(New Object() {"2020"})
        Me.CboCensusNen.Location = New System.Drawing.Point(223, 97)
        Me.CboCensusNen.Name = "CboCensusNen"
        Me.CboCensusNen.Size = New System.Drawing.Size(72, 20)
        Me.CboCensusNen.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(136, 143)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(77, 12)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "地方農政局等"
        '
        'CboChiho
        '
        Me.CboKyokuTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboKyokuTo.FormattingEnabled = True
        Me.CboKyokuTo.Location = New System.Drawing.Point(223, 140)
        Me.CboKyokuTo.Name = "CboChiho"
        Me.CboKyokuTo.Size = New System.Drawing.Size(150, 20)
        Me.CboKyokuTo.TabIndex = 10
        '
        'CboKeieiKeitai
        '
        Me.CboKeieiKeitai.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboKeieiKeitai.FormattingEnabled = True
        Me.CboKeieiKeitai.Location = New System.Drawing.Point(223, 183)
        Me.CboKeieiKeitai.Name = "CboKeieiKeitai"
        Me.CboKeieiKeitai.Size = New System.Drawing.Size(150, 20)
        Me.CboKeieiKeitai.TabIndex = 13
        '
        'CboEinoruikei
        '
        Me.CboEinoruikei.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboEinoruikei.Enabled = False
        Me.CboEinoruikei.FormattingEnabled = True
        Me.CboEinoruikei.Location = New System.Drawing.Point(223, 226)
        Me.CboEinoruikei.Name = "CboEinoruikei"
        Me.CboEinoruikei.Size = New System.Drawing.Size(150, 20)
        Me.CboEinoruikei.TabIndex = 19
        '
        'CboSeisanhi
        '
        Me.CboSeisanhi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboSeisanhi.Enabled = False
        Me.CboSeisanhi.FormattingEnabled = True
        Me.CboSeisanhi.Location = New System.Drawing.Point(515, 183)
        Me.CboSeisanhi.Name = "CboSeisanhi"
        Me.CboSeisanhi.Size = New System.Drawing.Size(150, 20)
        Me.CboSeisanhi.TabIndex = 16
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(136, 186)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 12)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "経営形態区分" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(136, 229)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 12)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "営農類型区分"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(429, 186)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 12)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "生産費区分"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(429, 229)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 12)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = "田畑区分"
        '
        'CboTahata
        '
        Me.CboTahata.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboTahata.Enabled = False
        Me.CboTahata.FormattingEnabled = True
        Me.CboTahata.Location = New System.Drawing.Point(515, 226)
        Me.CboTahata.Name = "CboTahata"
        Me.CboTahata.Size = New System.Drawing.Size(150, 20)
        Me.CboTahata.TabIndex = 22
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(138, 249)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(166, 12)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "（個人経営体、法人経営体のみ）"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(138, 120)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(41, 12)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "（必須）"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(138, 163)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(41, 12)
        Me.Label9.TabIndex = 9
        Me.Label9.Text = "（必須）"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(138, 206)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(41, 12)
        Me.Label11.TabIndex = 12
        Me.Label11.Text = "（必須）"
        '
        'CboShufukugyo
        '
        Me.CboShufukugyo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboShufukugyo.Enabled = False
        Me.CboShufukugyo.FormattingEnabled = True
        Me.CboShufukugyo.Location = New System.Drawing.Point(223, 269)
        Me.CboShufukugyo.Name = "CboShufukugyo"
        Me.CboShufukugyo.Size = New System.Drawing.Size(150, 20)
        Me.CboShufukugyo.TabIndex = 25
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(138, 272)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(65, 12)
        Me.Label12.TabIndex = 23
        Me.Label12.Text = "主副業区分"
        '
        'CboShurakuEino
        '
        Me.CboShurakuEino.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboShurakuEino.Enabled = False
        Me.CboShurakuEino.FormattingEnabled = True
        Me.CboShurakuEino.Location = New System.Drawing.Point(223, 312)
        Me.CboShurakuEino.Name = "CboShurakuEino"
        Me.CboShurakuEino.Size = New System.Drawing.Size(150, 20)
        Me.CboShurakuEino.TabIndex = 31
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(136, 315)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(77, 12)
        Me.Label13.TabIndex = 29
        Me.Label13.Text = "集落営農区分"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(138, 292)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(86, 12)
        Me.Label14.TabIndex = 24
        Me.Label14.Text = "（営農個人のみ）"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(138, 335)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(122, 12)
        Me.Label15.TabIndex = 30
        Me.Label15.Text = "（営農法人水田作のみ）"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(429, 206)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(190, 12)
        Me.Label16.TabIndex = 15
        Me.Label16.Text = "（個別経営体、組織法人経営体のみ）"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(429, 249)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 12)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "（小麦、大豆のみ）"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(429, 295)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(116, 12)
        Me.Label17.TabIndex = 27
        Me.Label17.Text = "（営農個人・副業のみ）"
        '
        'CboAoiro
        '
        Me.CboAoiro.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboAoiro.Enabled = False
        Me.CboAoiro.FormattingEnabled = True
        Me.CboAoiro.Location = New System.Drawing.Point(515, 272)
        Me.CboAoiro.Name = "CboAoiro"
        Me.CboAoiro.Size = New System.Drawing.Size(150, 20)
        Me.CboAoiro.TabIndex = 28
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(429, 275)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(77, 12)
        Me.Label18.TabIndex = 26
        Me.Label18.Text = "青色申告区分"
        '
        'BRA10410F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 373)
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
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.CboTahata)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CboSeisanhi)
        Me.Controls.Add(Me.CboEinoruikei)
        Me.Controls.Add(Me.CboKeieiKeitai)
        Me.Controls.Add(Me.CboKyokuTo)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.CboCensusNen)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.BtnPrint)
        Me.Name = "BRA10410F"
        Me.Text = "農業経営統計調査 - 本省工程 - 規模区分別期待標本及び予定経営体数出力"
        Me.Controls.SetChildIndex(Me.BtnPrint, 0)
        Me.Controls.SetChildIndex(Me.Label10, 0)
        Me.Controls.SetChildIndex(Me.CboCensusNen, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.CboKyokuTo, 0)
        Me.Controls.SetChildIndex(Me.CboKeieiKeitai, 0)
        Me.Controls.SetChildIndex(Me.CboEinoruikei, 0)
        Me.Controls.SetChildIndex(Me.CboSeisanhi, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.CboTahata, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.Label7, 0)
        Me.Controls.SetChildIndex(Me.Label8, 0)
        Me.Controls.SetChildIndex(Me.Label9, 0)
        Me.Controls.SetChildIndex(Me.Label11, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
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
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Protected WithEvents CboCensusNen As ComboBox
    Friend WithEvents Label3 As Label
    Protected WithEvents CboKyokuTo As ComboBox
    Protected WithEvents CboKeieiKeitai As ComboBox
    Protected WithEvents CboEinoruikei As ComboBox
    Protected WithEvents CboSeisanhi As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Protected WithEvents CboTahata As ComboBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label11 As Label
    Protected WithEvents CboShufukugyo As ComboBox
    Friend WithEvents Label12 As Label
    Protected WithEvents CboShurakuEino As ComboBox
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label17 As Label
    Protected WithEvents CboAoiro As ComboBox
    Friend WithEvents Label18 As Label
End Class
