<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA8220F
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
        Me.btnSave = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtUpperLimit = New System.Windows.Forms.TextBox()
        Me.txtAvg_5_29_Hon = New System.Windows.Forms.TextBox()
        Me.txtAvg_5_29_Hi = New System.Windows.Forms.TextBox()
        Me.txtAvg_5_29_Zen = New System.Windows.Forms.TextBox()
        Me.txtAvg_Saiyo_Zen = New System.Windows.Forms.TextBox()
        Me.txtAvg_Saiyo_Hon = New System.Windows.Forms.TextBox()
        Me.txtAvg_Saiyo_Hi = New System.Windows.Forms.TextBox()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.TextBox8 = New System.Windows.Forms.TextBox()
        Me.TextBox9 = New System.Windows.Forms.TextBox()
        Me.txtW_5_29_Hi = New System.Windows.Forms.TextBox()
        Me.txtW_5_29_Zen = New System.Windows.Forms.TextBox()
        Me.txtW_5_29_Hon = New System.Windows.Forms.TextBox()
        Me.TextBox13 = New System.Windows.Forms.TextBox()
        Me.TextBox14 = New System.Windows.Forms.TextBox()
        Me.TextBox15 = New System.Windows.Forms.TextBox()
        Me.TextBox16 = New System.Windows.Forms.TextBox()
        Me.txtM_5_29_Hi = New System.Windows.Forms.TextBox()
        Me.txtM_5_29_Zen = New System.Windows.Forms.TextBox()
        Me.txtM_5_29_Hon = New System.Windows.Forms.TextBox()
        Me.TextBox20 = New System.Windows.Forms.TextBox()
        Me.TextBox21 = New System.Windows.Forms.TextBox()
        Me.TextBox22 = New System.Windows.Forms.TextBox()
        Me.TextBox23 = New System.Windows.Forms.TextBox()
        Me.TextBox24 = New System.Windows.Forms.TextBox()
        Me.TextBox17 = New System.Windows.Forms.TextBox()
        Me.TextBox26 = New System.Windows.Forms.TextBox()
        Me.TextBox27 = New System.Windows.Forms.TextBox()
        Me.TextBox28 = New System.Windows.Forms.TextBox()
        Me.TextBox29 = New System.Windows.Forms.TextBox()
        Me.TextBox30 = New System.Windows.Forms.TextBox()
        Me.TextBox31 = New System.Windows.Forms.TextBox()
        Me.TextBox32 = New System.Windows.Forms.TextBox()
        Me.TextBox33 = New System.Windows.Forms.TextBox()
        Me.txtAvg_Hyouka = New System.Windows.Forms.TextBox()
        Me.txtAvg_Tsukin = New System.Windows.Forms.TextBox()
        Me.TextBox38 = New System.Windows.Forms.TextBox()
        Me.TextBox41 = New System.Windows.Forms.TextBox()
        Me.TextBox42 = New System.Windows.Forms.TextBox()
        Me.TextBox36 = New System.Windows.Forms.TextBox()
        Me.TextBox37 = New System.Windows.Forms.TextBox()
        Me.TextBox39 = New System.Windows.Forms.TextBox()
        Me.TextBox40 = New System.Windows.Forms.TextBox()
        Me.TextBox43 = New System.Windows.Forms.TextBox()
        Me.TextBox44 = New System.Windows.Forms.TextBox()
        Me.TextBox45 = New System.Windows.Forms.TextBox()
        Me.txtAvg_30_Taihi = New System.Windows.Forms.TextBox()
        Me.txtAvg_30_Sangyo = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtSeisanhi = New System.Windows.Forms.TextBox()
        Me.txtChosaNen = New System.Windows.Forms.TextBox()
        Me.txtTodofuken = New System.Windows.Forms.TextBox()
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
        Me.btnReturn.Location = New System.Drawing.Point(704, 530)
        Me.btnReturn.TabIndex = 63
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "労賃単価算出"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(608, 530)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(90, 30)
        Me.btnSave.TabIndex = 62
        Me.btnSave.Text = "保存"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(201, 99)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 12)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "生産費"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(201, 128)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 12)
        Me.Label10.TabIndex = 7
        Me.Label10.Text = "調査年（産）"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(201, 156)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "都道府県"
        '
        'txtUpperLimit
        '
        Me.txtUpperLimit.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtUpperLimit.Location = New System.Drawing.Point(531, 196)
        Me.txtUpperLimit.MaxLength = 50
        Me.txtUpperLimit.Name = "txtUpperLimit"
        Me.txtUpperLimit.ReadOnly = True
        Me.txtUpperLimit.Size = New System.Drawing.Size(83, 19)
        Me.txtUpperLimit.TabIndex = 15
        Me.txtUpperLimit.TabStop = False
        Me.txtUpperLimit.Text = "男女平均"
        '
        'txtAvg_5_29_Hon
        '
        Me.txtAvg_5_29_Hon.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtAvg_5_29_Hon.Location = New System.Drawing.Point(531, 214)
        Me.txtAvg_5_29_Hon.MaxLength = 50
        Me.txtAvg_5_29_Hon.Name = "txtAvg_5_29_Hon"
        Me.txtAvg_5_29_Hon.ReadOnly = True
        Me.txtAvg_5_29_Hon.Size = New System.Drawing.Size(83, 19)
        Me.txtAvg_5_29_Hon.TabIndex = 20
        Me.txtAvg_5_29_Hon.TabStop = False
        Me.txtAvg_5_29_Hon.Text = "9999999999"
        Me.txtAvg_5_29_Hon.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAvg_5_29_Hi
        '
        Me.txtAvg_5_29_Hi.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtAvg_5_29_Hi.Location = New System.Drawing.Point(531, 250)
        Me.txtAvg_5_29_Hi.MaxLength = 50
        Me.txtAvg_5_29_Hi.Name = "txtAvg_5_29_Hi"
        Me.txtAvg_5_29_Hi.ReadOnly = True
        Me.txtAvg_5_29_Hi.Size = New System.Drawing.Size(83, 19)
        Me.txtAvg_5_29_Hi.TabIndex = 28
        Me.txtAvg_5_29_Hi.TabStop = False
        Me.txtAvg_5_29_Hi.Text = "999.9"
        Me.txtAvg_5_29_Hi.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAvg_5_29_Zen
        '
        Me.txtAvg_5_29_Zen.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtAvg_5_29_Zen.Location = New System.Drawing.Point(531, 232)
        Me.txtAvg_5_29_Zen.MaxLength = 50
        Me.txtAvg_5_29_Zen.Name = "txtAvg_5_29_Zen"
        Me.txtAvg_5_29_Zen.ReadOnly = True
        Me.txtAvg_5_29_Zen.Size = New System.Drawing.Size(83, 19)
        Me.txtAvg_5_29_Zen.TabIndex = 24
        Me.txtAvg_5_29_Zen.TabStop = False
        Me.txtAvg_5_29_Zen.Text = "9999999999"
        Me.txtAvg_5_29_Zen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAvg_Saiyo_Zen
        '
        Me.txtAvg_Saiyo_Zen.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtAvg_Saiyo_Zen.Location = New System.Drawing.Point(531, 286)
        Me.txtAvg_Saiyo_Zen.MaxLength = 50
        Me.txtAvg_Saiyo_Zen.Name = "txtAvg_Saiyo_Zen"
        Me.txtAvg_Saiyo_Zen.Size = New System.Drawing.Size(83, 19)
        Me.txtAvg_Saiyo_Zen.TabIndex = 37
        Me.txtAvg_Saiyo_Zen.Text = "9999999999"
        Me.txtAvg_Saiyo_Zen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAvg_Saiyo_Hon
        '
        Me.txtAvg_Saiyo_Hon.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtAvg_Saiyo_Hon.Location = New System.Drawing.Point(531, 268)
        Me.txtAvg_Saiyo_Hon.MaxLength = 50
        Me.txtAvg_Saiyo_Hon.Name = "txtAvg_Saiyo_Hon"
        Me.txtAvg_Saiyo_Hon.Size = New System.Drawing.Size(83, 19)
        Me.txtAvg_Saiyo_Hon.TabIndex = 33
        Me.txtAvg_Saiyo_Hon.Text = "9999999999"
        Me.txtAvg_Saiyo_Hon.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAvg_Saiyo_Hi
        '
        Me.txtAvg_Saiyo_Hi.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtAvg_Saiyo_Hi.Location = New System.Drawing.Point(531, 304)
        Me.txtAvg_Saiyo_Hi.MaxLength = 50
        Me.txtAvg_Saiyo_Hi.Name = "txtAvg_Saiyo_Hi"
        Me.txtAvg_Saiyo_Hi.ReadOnly = True
        Me.txtAvg_Saiyo_Hi.Size = New System.Drawing.Size(83, 19)
        Me.txtAvg_Saiyo_Hi.TabIndex = 41
        Me.txtAvg_Saiyo_Hi.TabStop = False
        Me.txtAvg_Saiyo_Hi.Text = "999.9"
        Me.txtAvg_Saiyo_Hi.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox7
        '
        Me.TextBox7.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox7.Location = New System.Drawing.Point(449, 304)
        Me.TextBox7.MaxLength = 50
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.ReadOnly = True
        Me.TextBox7.Size = New System.Drawing.Size(83, 19)
        Me.TextBox7.TabIndex = 40
        Me.TextBox7.TabStop = False
        Me.TextBox7.Text = "-"
        Me.TextBox7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox8
        '
        Me.TextBox8.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox8.Location = New System.Drawing.Point(449, 286)
        Me.TextBox8.MaxLength = 50
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.ReadOnly = True
        Me.TextBox8.Size = New System.Drawing.Size(83, 19)
        Me.TextBox8.TabIndex = 36
        Me.TextBox8.TabStop = False
        Me.TextBox8.Text = "-"
        Me.TextBox8.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox9
        '
        Me.TextBox9.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox9.Location = New System.Drawing.Point(449, 268)
        Me.TextBox9.MaxLength = 50
        Me.TextBox9.Name = "TextBox9"
        Me.TextBox9.ReadOnly = True
        Me.TextBox9.Size = New System.Drawing.Size(83, 19)
        Me.TextBox9.TabIndex = 32
        Me.TextBox9.TabStop = False
        Me.TextBox9.Text = "-"
        Me.TextBox9.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtW_5_29_Hi
        '
        Me.txtW_5_29_Hi.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtW_5_29_Hi.Location = New System.Drawing.Point(449, 250)
        Me.txtW_5_29_Hi.MaxLength = 50
        Me.txtW_5_29_Hi.Name = "txtW_5_29_Hi"
        Me.txtW_5_29_Hi.ReadOnly = True
        Me.txtW_5_29_Hi.Size = New System.Drawing.Size(83, 19)
        Me.txtW_5_29_Hi.TabIndex = 27
        Me.txtW_5_29_Hi.TabStop = False
        Me.txtW_5_29_Hi.Text = "999.9"
        Me.txtW_5_29_Hi.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtW_5_29_Zen
        '
        Me.txtW_5_29_Zen.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtW_5_29_Zen.Location = New System.Drawing.Point(449, 232)
        Me.txtW_5_29_Zen.MaxLength = 50
        Me.txtW_5_29_Zen.Name = "txtW_5_29_Zen"
        Me.txtW_5_29_Zen.ReadOnly = True
        Me.txtW_5_29_Zen.Size = New System.Drawing.Size(83, 19)
        Me.txtW_5_29_Zen.TabIndex = 23
        Me.txtW_5_29_Zen.TabStop = False
        Me.txtW_5_29_Zen.Text = "9999999999"
        Me.txtW_5_29_Zen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtW_5_29_Hon
        '
        Me.txtW_5_29_Hon.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtW_5_29_Hon.Location = New System.Drawing.Point(449, 214)
        Me.txtW_5_29_Hon.MaxLength = 50
        Me.txtW_5_29_Hon.Name = "txtW_5_29_Hon"
        Me.txtW_5_29_Hon.ReadOnly = True
        Me.txtW_5_29_Hon.Size = New System.Drawing.Size(83, 19)
        Me.txtW_5_29_Hon.TabIndex = 19
        Me.txtW_5_29_Hon.TabStop = False
        Me.txtW_5_29_Hon.Text = "9999999999"
        Me.txtW_5_29_Hon.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox13
        '
        Me.TextBox13.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox13.Location = New System.Drawing.Point(449, 196)
        Me.TextBox13.MaxLength = 50
        Me.TextBox13.Name = "TextBox13"
        Me.TextBox13.ReadOnly = True
        Me.TextBox13.Size = New System.Drawing.Size(83, 19)
        Me.TextBox13.TabIndex = 14
        Me.TextBox13.TabStop = False
        Me.TextBox13.Text = "女"
        '
        'TextBox14
        '
        Me.TextBox14.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox14.Location = New System.Drawing.Point(367, 304)
        Me.TextBox14.MaxLength = 50
        Me.TextBox14.Name = "TextBox14"
        Me.TextBox14.ReadOnly = True
        Me.TextBox14.Size = New System.Drawing.Size(83, 19)
        Me.TextBox14.TabIndex = 39
        Me.TextBox14.TabStop = False
        Me.TextBox14.Text = "-"
        Me.TextBox14.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox15
        '
        Me.TextBox15.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox15.Location = New System.Drawing.Point(367, 286)
        Me.TextBox15.MaxLength = 50
        Me.TextBox15.Name = "TextBox15"
        Me.TextBox15.ReadOnly = True
        Me.TextBox15.Size = New System.Drawing.Size(83, 19)
        Me.TextBox15.TabIndex = 35
        Me.TextBox15.TabStop = False
        Me.TextBox15.Text = "-"
        Me.TextBox15.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox16
        '
        Me.TextBox16.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox16.Location = New System.Drawing.Point(367, 268)
        Me.TextBox16.MaxLength = 50
        Me.TextBox16.Name = "TextBox16"
        Me.TextBox16.ReadOnly = True
        Me.TextBox16.Size = New System.Drawing.Size(83, 19)
        Me.TextBox16.TabIndex = 31
        Me.TextBox16.TabStop = False
        Me.TextBox16.Text = "-"
        Me.TextBox16.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtM_5_29_Hi
        '
        Me.txtM_5_29_Hi.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtM_5_29_Hi.Location = New System.Drawing.Point(367, 250)
        Me.txtM_5_29_Hi.MaxLength = 50
        Me.txtM_5_29_Hi.Name = "txtM_5_29_Hi"
        Me.txtM_5_29_Hi.ReadOnly = True
        Me.txtM_5_29_Hi.Size = New System.Drawing.Size(83, 19)
        Me.txtM_5_29_Hi.TabIndex = 26
        Me.txtM_5_29_Hi.TabStop = False
        Me.txtM_5_29_Hi.Text = "999.9"
        Me.txtM_5_29_Hi.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtM_5_29_Zen
        '
        Me.txtM_5_29_Zen.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtM_5_29_Zen.Location = New System.Drawing.Point(367, 232)
        Me.txtM_5_29_Zen.MaxLength = 50
        Me.txtM_5_29_Zen.Name = "txtM_5_29_Zen"
        Me.txtM_5_29_Zen.ReadOnly = True
        Me.txtM_5_29_Zen.Size = New System.Drawing.Size(83, 19)
        Me.txtM_5_29_Zen.TabIndex = 22
        Me.txtM_5_29_Zen.TabStop = False
        Me.txtM_5_29_Zen.Text = "9999999999"
        Me.txtM_5_29_Zen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtM_5_29_Hon
        '
        Me.txtM_5_29_Hon.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtM_5_29_Hon.Location = New System.Drawing.Point(367, 214)
        Me.txtM_5_29_Hon.MaxLength = 50
        Me.txtM_5_29_Hon.Name = "txtM_5_29_Hon"
        Me.txtM_5_29_Hon.ReadOnly = True
        Me.txtM_5_29_Hon.Size = New System.Drawing.Size(83, 19)
        Me.txtM_5_29_Hon.TabIndex = 18
        Me.txtM_5_29_Hon.TabStop = False
        Me.txtM_5_29_Hon.Text = "9999999999"
        Me.txtM_5_29_Hon.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox20
        '
        Me.TextBox20.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox20.Location = New System.Drawing.Point(367, 196)
        Me.TextBox20.MaxLength = 50
        Me.TextBox20.Name = "TextBox20"
        Me.TextBox20.ReadOnly = True
        Me.TextBox20.Size = New System.Drawing.Size(83, 19)
        Me.TextBox20.TabIndex = 13
        Me.TextBox20.TabStop = False
        Me.TextBox20.Text = "男"
        '
        'TextBox21
        '
        Me.TextBox21.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox21.Location = New System.Drawing.Point(285, 304)
        Me.TextBox21.MaxLength = 50
        Me.TextBox21.Name = "TextBox21"
        Me.TextBox21.ReadOnly = True
        Me.TextBox21.Size = New System.Drawing.Size(83, 19)
        Me.TextBox21.TabIndex = 38
        Me.TextBox21.TabStop = False
        Me.TextBox21.Text = "対前年比"
        '
        'TextBox22
        '
        Me.TextBox22.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox22.Location = New System.Drawing.Point(285, 286)
        Me.TextBox22.MaxLength = 50
        Me.TextBox22.Name = "TextBox22"
        Me.TextBox22.ReadOnly = True
        Me.TextBox22.Size = New System.Drawing.Size(83, 19)
        Me.TextBox22.TabIndex = 34
        Me.TextBox22.TabStop = False
        Me.TextBox22.Text = "前年値"
        '
        'TextBox23
        '
        Me.TextBox23.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox23.Location = New System.Drawing.Point(285, 268)
        Me.TextBox23.MaxLength = 50
        Me.TextBox23.Name = "TextBox23"
        Me.TextBox23.ReadOnly = True
        Me.TextBox23.Size = New System.Drawing.Size(83, 19)
        Me.TextBox23.TabIndex = 30
        Me.TextBox23.TabStop = False
        Me.TextBox23.Text = "本年値"
        '
        'TextBox24
        '
        Me.TextBox24.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox24.Location = New System.Drawing.Point(285, 250)
        Me.TextBox24.MaxLength = 50
        Me.TextBox24.Name = "TextBox24"
        Me.TextBox24.ReadOnly = True
        Me.TextBox24.Size = New System.Drawing.Size(83, 19)
        Me.TextBox24.TabIndex = 25
        Me.TextBox24.TabStop = False
        Me.TextBox24.Text = "対前年比"
        '
        'TextBox17
        '
        Me.TextBox17.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox17.Location = New System.Drawing.Point(285, 232)
        Me.TextBox17.MaxLength = 50
        Me.TextBox17.Name = "TextBox17"
        Me.TextBox17.ReadOnly = True
        Me.TextBox17.Size = New System.Drawing.Size(83, 19)
        Me.TextBox17.TabIndex = 21
        Me.TextBox17.TabStop = False
        Me.TextBox17.Text = "前年値"
        '
        'TextBox26
        '
        Me.TextBox26.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox26.Location = New System.Drawing.Point(285, 214)
        Me.TextBox26.MaxLength = 50
        Me.TextBox26.Name = "TextBox26"
        Me.TextBox26.ReadOnly = True
        Me.TextBox26.Size = New System.Drawing.Size(83, 19)
        Me.TextBox26.TabIndex = 17
        Me.TextBox26.TabStop = False
        Me.TextBox26.Text = "本年値"
        '
        'TextBox27
        '
        Me.TextBox27.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox27.Location = New System.Drawing.Point(285, 196)
        Me.TextBox27.MaxLength = 50
        Me.TextBox27.Name = "TextBox27"
        Me.TextBox27.ReadOnly = True
        Me.TextBox27.Size = New System.Drawing.Size(83, 19)
        Me.TextBox27.TabIndex = 12
        Me.TextBox27.TabStop = False
        '
        'TextBox28
        '
        Me.TextBox28.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox28.Location = New System.Drawing.Point(203, 340)
        Me.TextBox28.MaxLength = 50
        Me.TextBox28.Name = "TextBox28"
        Me.TextBox28.ReadOnly = True
        Me.TextBox28.Size = New System.Drawing.Size(165, 19)
        Me.TextBox28.TabIndex = 46
        Me.TextBox28.TabStop = False
        Me.TextBox28.Text = "評価単価（65際未満）"
        '
        'TextBox29
        '
        Me.TextBox29.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox29.Location = New System.Drawing.Point(203, 322)
        Me.TextBox29.MaxLength = 50
        Me.TextBox29.Name = "TextBox29"
        Me.TextBox29.ReadOnly = True
        Me.TextBox29.Size = New System.Drawing.Size(165, 19)
        Me.TextBox29.TabIndex = 42
        Me.TextBox29.TabStop = False
        Me.TextBox29.Text = "通勤手当割合"
        '
        'TextBox30
        '
        Me.TextBox30.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox30.Location = New System.Drawing.Point(367, 340)
        Me.TextBox30.MaxLength = 50
        Me.TextBox30.Name = "TextBox30"
        Me.TextBox30.ReadOnly = True
        Me.TextBox30.Size = New System.Drawing.Size(83, 19)
        Me.TextBox30.TabIndex = 47
        Me.TextBox30.TabStop = False
        Me.TextBox30.Text = "-"
        Me.TextBox30.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox31
        '
        Me.TextBox31.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox31.Location = New System.Drawing.Point(367, 322)
        Me.TextBox31.MaxLength = 50
        Me.TextBox31.Name = "TextBox31"
        Me.TextBox31.ReadOnly = True
        Me.TextBox31.Size = New System.Drawing.Size(83, 19)
        Me.TextBox31.TabIndex = 43
        Me.TextBox31.TabStop = False
        Me.TextBox31.Text = "-"
        Me.TextBox31.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox32
        '
        Me.TextBox32.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox32.Location = New System.Drawing.Point(449, 340)
        Me.TextBox32.MaxLength = 50
        Me.TextBox32.Name = "TextBox32"
        Me.TextBox32.ReadOnly = True
        Me.TextBox32.Size = New System.Drawing.Size(83, 19)
        Me.TextBox32.TabIndex = 48
        Me.TextBox32.TabStop = False
        Me.TextBox32.Text = "-"
        Me.TextBox32.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox33
        '
        Me.TextBox33.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox33.Location = New System.Drawing.Point(449, 322)
        Me.TextBox33.MaxLength = 50
        Me.TextBox33.Name = "TextBox33"
        Me.TextBox33.ReadOnly = True
        Me.TextBox33.Size = New System.Drawing.Size(83, 19)
        Me.TextBox33.TabIndex = 44
        Me.TextBox33.TabStop = False
        Me.TextBox33.Text = "-"
        Me.TextBox33.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtAvg_Hyouka
        '
        Me.txtAvg_Hyouka.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtAvg_Hyouka.Location = New System.Drawing.Point(531, 340)
        Me.txtAvg_Hyouka.MaxLength = 50
        Me.txtAvg_Hyouka.Name = "txtAvg_Hyouka"
        Me.txtAvg_Hyouka.ReadOnly = True
        Me.txtAvg_Hyouka.Size = New System.Drawing.Size(83, 19)
        Me.txtAvg_Hyouka.TabIndex = 49
        Me.txtAvg_Hyouka.TabStop = False
        Me.txtAvg_Hyouka.Text = "9999999999"
        Me.txtAvg_Hyouka.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAvg_Tsukin
        '
        Me.txtAvg_Tsukin.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtAvg_Tsukin.Location = New System.Drawing.Point(531, 322)
        Me.txtAvg_Tsukin.MaxLength = 50
        Me.txtAvg_Tsukin.Name = "txtAvg_Tsukin"
        Me.txtAvg_Tsukin.Size = New System.Drawing.Size(83, 19)
        Me.txtAvg_Tsukin.TabIndex = 45
        Me.txtAvg_Tsukin.Text = "9999999999"
        Me.txtAvg_Tsukin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox38
        '
        Me.TextBox38.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox38.Location = New System.Drawing.Point(203, 268)
        Me.TextBox38.MaxLength = 50
        Me.TextBox38.Multiline = True
        Me.TextBox38.Name = "TextBox38"
        Me.TextBox38.ReadOnly = True
        Me.TextBox38.Size = New System.Drawing.Size(83, 55)
        Me.TextBox38.TabIndex = 29
        Me.TextBox38.TabStop = False
        Me.TextBox38.Text = "採用単価"
        '
        'TextBox41
        '
        Me.TextBox41.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox41.Location = New System.Drawing.Point(203, 214)
        Me.TextBox41.MaxLength = 50
        Me.TextBox41.Multiline = True
        Me.TextBox41.Name = "TextBox41"
        Me.TextBox41.ReadOnly = True
        Me.TextBox41.Size = New System.Drawing.Size(83, 55)
        Me.TextBox41.TabIndex = 16
        Me.TextBox41.TabStop = False
        Me.TextBox41.Text = "5～29人"
        '
        'TextBox42
        '
        Me.TextBox42.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox42.Location = New System.Drawing.Point(203, 196)
        Me.TextBox42.MaxLength = 50
        Me.TextBox42.Name = "TextBox42"
        Me.TextBox42.ReadOnly = True
        Me.TextBox42.Size = New System.Drawing.Size(83, 19)
        Me.TextBox42.TabIndex = 11
        Me.TextBox42.TabStop = False
        '
        'TextBox36
        '
        Me.TextBox36.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox36.Location = New System.Drawing.Point(203, 398)
        Me.TextBox36.MaxLength = 50
        Me.TextBox36.Multiline = True
        Me.TextBox36.Name = "TextBox36"
        Me.TextBox36.ReadOnly = True
        Me.TextBox36.Size = New System.Drawing.Size(83, 37)
        Me.TextBox36.TabIndex = 51
        Me.TextBox36.TabStop = False
        '
        'TextBox37
        '
        Me.TextBox37.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox37.Location = New System.Drawing.Point(285, 416)
        Me.TextBox37.MaxLength = 50
        Me.TextBox37.Name = "TextBox37"
        Me.TextBox37.ReadOnly = True
        Me.TextBox37.Size = New System.Drawing.Size(83, 19)
        Me.TextBox37.TabIndex = 57
        Me.TextBox37.TabStop = False
        Me.TextBox37.Text = "対比単価"
        '
        'TextBox39
        '
        Me.TextBox39.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox39.Location = New System.Drawing.Point(285, 398)
        Me.TextBox39.MaxLength = 50
        Me.TextBox39.Name = "TextBox39"
        Me.TextBox39.ReadOnly = True
        Me.TextBox39.Size = New System.Drawing.Size(83, 19)
        Me.TextBox39.TabIndex = 53
        Me.TextBox39.TabStop = False
        Me.TextBox39.Text = "産業計対比"
        '
        'TextBox40
        '
        Me.TextBox40.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox40.Location = New System.Drawing.Point(367, 416)
        Me.TextBox40.MaxLength = 50
        Me.TextBox40.Name = "TextBox40"
        Me.TextBox40.ReadOnly = True
        Me.TextBox40.Size = New System.Drawing.Size(83, 19)
        Me.TextBox40.TabIndex = 58
        Me.TextBox40.TabStop = False
        Me.TextBox40.Text = "-"
        Me.TextBox40.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox43
        '
        Me.TextBox43.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox43.Location = New System.Drawing.Point(367, 398)
        Me.TextBox43.MaxLength = 50
        Me.TextBox43.Name = "TextBox43"
        Me.TextBox43.ReadOnly = True
        Me.TextBox43.Size = New System.Drawing.Size(83, 19)
        Me.TextBox43.TabIndex = 54
        Me.TextBox43.TabStop = False
        Me.TextBox43.Text = "-"
        Me.TextBox43.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox44
        '
        Me.TextBox44.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox44.Location = New System.Drawing.Point(449, 416)
        Me.TextBox44.MaxLength = 50
        Me.TextBox44.Name = "TextBox44"
        Me.TextBox44.ReadOnly = True
        Me.TextBox44.Size = New System.Drawing.Size(83, 19)
        Me.TextBox44.TabIndex = 59
        Me.TextBox44.TabStop = False
        Me.TextBox44.Text = "-"
        Me.TextBox44.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox45
        '
        Me.TextBox45.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TextBox45.Location = New System.Drawing.Point(449, 398)
        Me.TextBox45.MaxLength = 50
        Me.TextBox45.Name = "TextBox45"
        Me.TextBox45.ReadOnly = True
        Me.TextBox45.Size = New System.Drawing.Size(83, 19)
        Me.TextBox45.TabIndex = 55
        Me.TextBox45.TabStop = False
        Me.TextBox45.Text = "-"
        Me.TextBox45.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtAvg_30_Taihi
        '
        Me.txtAvg_30_Taihi.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtAvg_30_Taihi.Location = New System.Drawing.Point(531, 416)
        Me.txtAvg_30_Taihi.MaxLength = 50
        Me.txtAvg_30_Taihi.Name = "txtAvg_30_Taihi"
        Me.txtAvg_30_Taihi.ReadOnly = True
        Me.txtAvg_30_Taihi.Size = New System.Drawing.Size(83, 19)
        Me.txtAvg_30_Taihi.TabIndex = 60
        Me.txtAvg_30_Taihi.TabStop = False
        Me.txtAvg_30_Taihi.Text = "9999999999"
        Me.txtAvg_30_Taihi.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAvg_30_Sangyo
        '
        Me.txtAvg_30_Sangyo.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtAvg_30_Sangyo.Location = New System.Drawing.Point(531, 398)
        Me.txtAvg_30_Sangyo.MaxLength = 50
        Me.txtAvg_30_Sangyo.Name = "txtAvg_30_Sangyo"
        Me.txtAvg_30_Sangyo.ReadOnly = True
        Me.txtAvg_30_Sangyo.Size = New System.Drawing.Size(83, 19)
        Me.txtAvg_30_Sangyo.TabIndex = 56
        Me.txtAvg_30_Sangyo.TabStop = False
        Me.txtAvg_30_Sangyo.Text = "999.9"
        Me.txtAvg_30_Sangyo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(214, 405)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 24)
        Me.Label3.TabIndex = 52
        Me.Label3.Text = "30人以上" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "規模の動向"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(208, 383)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(41, 12)
        Me.Label4.TabIndex = 50
        Me.Label4.Text = "（参考）"
        '
        'txtSeisanhi
        '
        Me.txtSeisanhi.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtSeisanhi.Location = New System.Drawing.Point(278, 96)
        Me.txtSeisanhi.MaxLength = 50
        Me.txtSeisanhi.Name = "txtSeisanhi"
        Me.txtSeisanhi.ReadOnly = True
        Me.txtSeisanhi.Size = New System.Drawing.Size(213, 19)
        Me.txtSeisanhi.TabIndex = 64
        Me.txtSeisanhi.TabStop = False
        '
        'txtChosaNen
        '
        Me.txtChosaNen.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtChosaNen.Location = New System.Drawing.Point(278, 125)
        Me.txtChosaNen.MaxLength = 50
        Me.txtChosaNen.Name = "txtChosaNen"
        Me.txtChosaNen.ReadOnly = True
        Me.txtChosaNen.Size = New System.Drawing.Size(72, 19)
        Me.txtChosaNen.TabIndex = 65
        Me.txtChosaNen.TabStop = False
        '
        'txtTodofuken
        '
        Me.txtTodofuken.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtTodofuken.Location = New System.Drawing.Point(278, 153)
        Me.txtTodofuken.MaxLength = 50
        Me.txtTodofuken.Name = "txtTodofuken"
        Me.txtTodofuken.ReadOnly = True
        Me.txtTodofuken.Size = New System.Drawing.Size(72, 19)
        Me.txtTodofuken.TabIndex = 66
        Me.txtTodofuken.TabStop = False
        '
        'BRA8220F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 576)
        Me.Controls.Add(Me.txtTodofuken)
        Me.Controls.Add(Me.txtChosaNen)
        Me.Controls.Add(Me.txtSeisanhi)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox36)
        Me.Controls.Add(Me.TextBox37)
        Me.Controls.Add(Me.TextBox39)
        Me.Controls.Add(Me.TextBox40)
        Me.Controls.Add(Me.TextBox43)
        Me.Controls.Add(Me.TextBox44)
        Me.Controls.Add(Me.TextBox45)
        Me.Controls.Add(Me.txtAvg_30_Taihi)
        Me.Controls.Add(Me.txtAvg_30_Sangyo)
        Me.Controls.Add(Me.TextBox38)
        Me.Controls.Add(Me.TextBox41)
        Me.Controls.Add(Me.TextBox42)
        Me.Controls.Add(Me.TextBox28)
        Me.Controls.Add(Me.TextBox29)
        Me.Controls.Add(Me.TextBox30)
        Me.Controls.Add(Me.TextBox31)
        Me.Controls.Add(Me.TextBox32)
        Me.Controls.Add(Me.TextBox33)
        Me.Controls.Add(Me.txtAvg_Hyouka)
        Me.Controls.Add(Me.txtAvg_Tsukin)
        Me.Controls.Add(Me.TextBox21)
        Me.Controls.Add(Me.TextBox22)
        Me.Controls.Add(Me.TextBox23)
        Me.Controls.Add(Me.TextBox24)
        Me.Controls.Add(Me.TextBox17)
        Me.Controls.Add(Me.TextBox26)
        Me.Controls.Add(Me.TextBox27)
        Me.Controls.Add(Me.TextBox14)
        Me.Controls.Add(Me.TextBox15)
        Me.Controls.Add(Me.TextBox16)
        Me.Controls.Add(Me.txtM_5_29_Hi)
        Me.Controls.Add(Me.txtM_5_29_Zen)
        Me.Controls.Add(Me.txtM_5_29_Hon)
        Me.Controls.Add(Me.TextBox20)
        Me.Controls.Add(Me.TextBox7)
        Me.Controls.Add(Me.TextBox8)
        Me.Controls.Add(Me.TextBox9)
        Me.Controls.Add(Me.txtW_5_29_Hi)
        Me.Controls.Add(Me.txtW_5_29_Zen)
        Me.Controls.Add(Me.txtW_5_29_Hon)
        Me.Controls.Add(Me.TextBox13)
        Me.Controls.Add(Me.txtAvg_Saiyo_Hi)
        Me.Controls.Add(Me.txtAvg_Saiyo_Zen)
        Me.Controls.Add(Me.txtAvg_Saiyo_Hon)
        Me.Controls.Add(Me.txtAvg_5_29_Hi)
        Me.Controls.Add(Me.txtAvg_5_29_Zen)
        Me.Controls.Add(Me.txtAvg_5_29_Hon)
        Me.Controls.Add(Me.txtUpperLimit)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.btnSave)
        Me.Name = "BRA8220F"
        Me.Text = "農業経営統計調査 - 本省工程 - 調査情報入力"
        Me.Controls.SetChildIndex(Me.btnSave, 0)
        Me.Controls.SetChildIndex(Me.Label10, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.txtUpperLimit, 0)
        Me.Controls.SetChildIndex(Me.txtAvg_5_29_Hon, 0)
        Me.Controls.SetChildIndex(Me.txtAvg_5_29_Zen, 0)
        Me.Controls.SetChildIndex(Me.txtAvg_5_29_Hi, 0)
        Me.Controls.SetChildIndex(Me.txtAvg_Saiyo_Hon, 0)
        Me.Controls.SetChildIndex(Me.txtAvg_Saiyo_Zen, 0)
        Me.Controls.SetChildIndex(Me.txtAvg_Saiyo_Hi, 0)
        Me.Controls.SetChildIndex(Me.TextBox13, 0)
        Me.Controls.SetChildIndex(Me.txtW_5_29_Hon, 0)
        Me.Controls.SetChildIndex(Me.txtW_5_29_Zen, 0)
        Me.Controls.SetChildIndex(Me.txtW_5_29_Hi, 0)
        Me.Controls.SetChildIndex(Me.TextBox9, 0)
        Me.Controls.SetChildIndex(Me.TextBox8, 0)
        Me.Controls.SetChildIndex(Me.TextBox7, 0)
        Me.Controls.SetChildIndex(Me.TextBox20, 0)
        Me.Controls.SetChildIndex(Me.txtM_5_29_Hon, 0)
        Me.Controls.SetChildIndex(Me.txtM_5_29_Zen, 0)
        Me.Controls.SetChildIndex(Me.txtM_5_29_Hi, 0)
        Me.Controls.SetChildIndex(Me.TextBox16, 0)
        Me.Controls.SetChildIndex(Me.TextBox15, 0)
        Me.Controls.SetChildIndex(Me.TextBox14, 0)
        Me.Controls.SetChildIndex(Me.TextBox27, 0)
        Me.Controls.SetChildIndex(Me.TextBox26, 0)
        Me.Controls.SetChildIndex(Me.TextBox17, 0)
        Me.Controls.SetChildIndex(Me.TextBox24, 0)
        Me.Controls.SetChildIndex(Me.TextBox23, 0)
        Me.Controls.SetChildIndex(Me.TextBox22, 0)
        Me.Controls.SetChildIndex(Me.TextBox21, 0)
        Me.Controls.SetChildIndex(Me.txtAvg_Tsukin, 0)
        Me.Controls.SetChildIndex(Me.txtAvg_Hyouka, 0)
        Me.Controls.SetChildIndex(Me.TextBox33, 0)
        Me.Controls.SetChildIndex(Me.TextBox32, 0)
        Me.Controls.SetChildIndex(Me.TextBox31, 0)
        Me.Controls.SetChildIndex(Me.TextBox30, 0)
        Me.Controls.SetChildIndex(Me.TextBox29, 0)
        Me.Controls.SetChildIndex(Me.TextBox28, 0)
        Me.Controls.SetChildIndex(Me.TextBox42, 0)
        Me.Controls.SetChildIndex(Me.TextBox41, 0)
        Me.Controls.SetChildIndex(Me.TextBox38, 0)
        Me.Controls.SetChildIndex(Me.txtAvg_30_Sangyo, 0)
        Me.Controls.SetChildIndex(Me.txtAvg_30_Taihi, 0)
        Me.Controls.SetChildIndex(Me.TextBox45, 0)
        Me.Controls.SetChildIndex(Me.TextBox44, 0)
        Me.Controls.SetChildIndex(Me.TextBox43, 0)
        Me.Controls.SetChildIndex(Me.TextBox40, 0)
        Me.Controls.SetChildIndex(Me.TextBox39, 0)
        Me.Controls.SetChildIndex(Me.TextBox37, 0)
        Me.Controls.SetChildIndex(Me.TextBox36, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.txtSeisanhi, 0)
        Me.Controls.SetChildIndex(Me.txtChosaNen, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.txtTodofuken, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtUpperLimit As System.Windows.Forms.TextBox
    Friend WithEvents txtAvg_5_29_Hon As System.Windows.Forms.TextBox
    Friend WithEvents txtAvg_5_29_Hi As System.Windows.Forms.TextBox
    Friend WithEvents txtAvg_5_29_Zen As System.Windows.Forms.TextBox
    Friend WithEvents txtAvg_Saiyo_Zen As System.Windows.Forms.TextBox
    Friend WithEvents txtAvg_Saiyo_Hon As System.Windows.Forms.TextBox
    Friend WithEvents txtAvg_Saiyo_Hi As System.Windows.Forms.TextBox
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox9 As System.Windows.Forms.TextBox
    Friend WithEvents txtW_5_29_Hi As System.Windows.Forms.TextBox
    Friend WithEvents txtW_5_29_Zen As System.Windows.Forms.TextBox
    Friend WithEvents txtW_5_29_Hon As System.Windows.Forms.TextBox
    Friend WithEvents TextBox13 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox14 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox15 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox16 As System.Windows.Forms.TextBox
    Friend WithEvents txtM_5_29_Hi As System.Windows.Forms.TextBox
    Friend WithEvents txtM_5_29_Zen As System.Windows.Forms.TextBox
    Friend WithEvents txtM_5_29_Hon As System.Windows.Forms.TextBox
    Friend WithEvents TextBox20 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox21 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox22 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox23 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox24 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox17 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox26 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox27 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox28 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox29 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox30 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox31 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox32 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox33 As System.Windows.Forms.TextBox
    Friend WithEvents txtAvg_Hyouka As System.Windows.Forms.TextBox
    Friend WithEvents txtAvg_Tsukin As System.Windows.Forms.TextBox
    Friend WithEvents TextBox38 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox41 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox42 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox36 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox37 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox39 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox40 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox43 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox44 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox45 As System.Windows.Forms.TextBox
    Friend WithEvents txtAvg_30_Taihi As System.Windows.Forms.TextBox
    Friend WithEvents txtAvg_30_Sangyo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSeisanhi As System.Windows.Forms.TextBox
    Friend WithEvents txtChosaNen As System.Windows.Forms.TextBox
    Friend WithEvents txtTodofuken As System.Windows.Forms.TextBox

End Class
