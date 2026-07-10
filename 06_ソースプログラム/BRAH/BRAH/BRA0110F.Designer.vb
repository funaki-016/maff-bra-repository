<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA0110F
    Inherits Maff.BRA.BRA01XXF

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
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnTraceMainte = New System.Windows.Forms.Button()
        Me.btnTraceImport = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.BtnHyohonJushin = New System.Windows.Forms.Button()
        Me.BtnHyohonPrint = New System.Windows.Forms.Button()
        Me.BtnChushutsuJoken = New System.Windows.Forms.Button()
        Me.BtnHyohonKanri = New System.Windows.Forms.Button()
        Me.BtnHyohonTorikomi = New System.Windows.Forms.Button()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.BtnLinkageCsv = New System.Windows.Forms.Button()
        Me.BtnMainteNogyotTiikiMaster = New System.Windows.Forms.Button()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.BtnMainteNogyotTiikiMaster)
        Me.GroupBox3.Location = New System.Drawing.Point(90, 516)
        Me.GroupBox3.Size = New System.Drawing.Size(294, 86)
        Me.GroupBox3.TabIndex = 7
        Me.GroupBox3.Controls.SetChildIndex(Me.BtnMainteNogyotTiikiMaster, 0)
        Me.GroupBox3.Controls.SetChildIndex(Me.btnMainteDB, 0)
        '
        'btnMainteDB
        '
        Me.btnMainteDB.Location = New System.Drawing.Point(6, 33)
        Me.btnMainteDB.Size = New System.Drawing.Size(124, 28)
        '
        'btnSelect
        '
        '
        'cboKubun2
        '
        '
        'cboKubun1
        '
        '
        'btnReturn
        '
        Me.btnReturn.Location = New System.Drawing.Point(530, 617)
        Me.btnReturn.TabIndex = 9
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnTraceMainte)
        Me.GroupBox2.Controls.Add(Me.btnTraceImport)
        Me.GroupBox2.Location = New System.Drawing.Point(90, 330)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(454, 79)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "牛トレサデータ"
        '
        'btnTraceMainte
        '
        Me.btnTraceMainte.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnTraceMainte.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnTraceMainte.Location = New System.Drawing.Point(210, 31)
        Me.btnTraceMainte.Name = "btnTraceMainte"
        Me.btnTraceMainte.Size = New System.Drawing.Size(172, 28)
        Me.btnTraceMainte.TabIndex = 1
        Me.btnTraceMainte.Text = "牛トレサデータ管理"
        Me.btnTraceMainte.UseVisualStyleBackColor = True
        '
        'btnTraceImport
        '
        Me.btnTraceImport.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnTraceImport.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnTraceImport.Location = New System.Drawing.Point(28, 31)
        Me.btnTraceImport.Name = "btnTraceImport"
        Me.btnTraceImport.Size = New System.Drawing.Size(172, 28)
        Me.btnTraceImport.TabIndex = 0
        Me.btnTraceImport.Text = "牛トレサデータ取込"
        Me.btnTraceImport.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.BtnHyohonJushin)
        Me.GroupBox4.Controls.Add(Me.BtnHyohonPrint)
        Me.GroupBox4.Controls.Add(Me.BtnChushutsuJoken)
        Me.GroupBox4.Controls.Add(Me.BtnHyohonKanri)
        Me.GroupBox4.Controls.Add(Me.BtnHyohonTorikomi)
        Me.GroupBox4.Location = New System.Drawing.Point(90, 415)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(454, 95)
        Me.GroupBox4.TabIndex = 6
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "標本管理"
        '
        'BtnHyohonJushin
        '
        Me.BtnHyohonJushin.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHyohonJushin.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHyohonJushin.Location = New System.Drawing.Point(173, 52)
        Me.BtnHyohonJushin.Name = "BtnHyohonJushin"
        Me.BtnHyohonJushin.Size = New System.Drawing.Size(132, 28)
        Me.BtnHyohonJushin.TabIndex = 4
        Me.BtnHyohonJushin.Text = "標本リスト受信"
        Me.BtnHyohonJushin.UseVisualStyleBackColor = True
        '
        'BtnHyohonPrint
        '
        Me.BtnHyohonPrint.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHyohonPrint.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHyohonPrint.Location = New System.Drawing.Point(35, 52)
        Me.BtnHyohonPrint.Name = "BtnHyohonPrint"
        Me.BtnHyohonPrint.Size = New System.Drawing.Size(132, 28)
        Me.BtnHyohonPrint.TabIndex = 3
        Me.BtnHyohonPrint.Text = "予定経営体数一覧表"
        Me.BtnHyohonPrint.UseVisualStyleBackColor = True
        '
        'BtnChushutsuJoken
        '
        Me.BtnChushutsuJoken.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnChushutsuJoken.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnChushutsuJoken.Location = New System.Drawing.Point(311, 18)
        Me.BtnChushutsuJoken.Name = "BtnChushutsuJoken"
        Me.BtnChushutsuJoken.Size = New System.Drawing.Size(132, 28)
        Me.BtnChushutsuJoken.TabIndex = 2
        Me.BtnChushutsuJoken.Text = "抽出条件設定"
        Me.BtnChushutsuJoken.UseVisualStyleBackColor = True
        '
        'BtnHyohonKanri
        '
        Me.BtnHyohonKanri.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHyohonKanri.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHyohonKanri.Location = New System.Drawing.Point(173, 18)
        Me.BtnHyohonKanri.Name = "BtnHyohonKanri"
        Me.BtnHyohonKanri.Size = New System.Drawing.Size(132, 28)
        Me.BtnHyohonKanri.TabIndex = 1
        Me.BtnHyohonKanri.Text = "標本リスト管理"
        Me.BtnHyohonKanri.UseVisualStyleBackColor = True
        '
        'BtnHyohonTorikomi
        '
        Me.BtnHyohonTorikomi.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHyohonTorikomi.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHyohonTorikomi.Location = New System.Drawing.Point(35, 18)
        Me.BtnHyohonTorikomi.Name = "BtnHyohonTorikomi"
        Me.BtnHyohonTorikomi.Size = New System.Drawing.Size(132, 28)
        Me.BtnHyohonTorikomi.TabIndex = 0
        Me.BtnHyohonTorikomi.Text = "標本リスト取込"
        Me.BtnHyohonTorikomi.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.BtnLinkageCsv)
        Me.GroupBox5.Location = New System.Drawing.Point(390, 516)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(154, 86)
        Me.GroupBox5.TabIndex = 8
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "米在庫リンケージ"
        '
        'BtnLinkageCsv
        '
        Me.BtnLinkageCsv.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnLinkageCsv.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnLinkageCsv.Location = New System.Drawing.Point(19, 33)
        Me.BtnLinkageCsv.Name = "BtnLinkageCsv"
        Me.BtnLinkageCsv.Size = New System.Drawing.Size(124, 28)
        Me.BtnLinkageCsv.TabIndex = 0
        Me.BtnLinkageCsv.Text = "CSV出力"
        Me.BtnLinkageCsv.UseVisualStyleBackColor = True
        '
        'BtnMainteNogyotTiikiMaster
        '
        Me.BtnMainteNogyotTiikiMaster.Location = New System.Drawing.Point(136, 33)
        Me.BtnMainteNogyotTiikiMaster.Name = "BtnMainteNogyotTiikiMaster"
        Me.BtnMainteNogyotTiikiMaster.Size = New System.Drawing.Size(139, 28)
        Me.BtnMainteNogyotTiikiMaster.TabIndex = 1
        Me.BtnMainteNogyotTiikiMaster.Text = "農業地域類型マスタ管理"
        Me.BtnMainteNogyotTiikiMaster.UseVisualStyleBackColor = True
        '
        'BRA0110F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(634, 661)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "BRA0110F"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox3, 0)
        Me.Controls.SetChildIndex(Me.GroupBox2, 0)
        Me.Controls.SetChildIndex(Me.GroupBox4, 0)
        Me.Controls.SetChildIndex(Me.GroupBox5, 0)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnTraceMainte As System.Windows.Forms.Button
    Friend WithEvents btnTraceImport As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents BtnHyohonJushin As Button
    Friend WithEvents BtnHyohonPrint As Button
    Friend WithEvents BtnChushutsuJoken As Button
    Friend WithEvents BtnHyohonKanri As Button
    Friend WithEvents BtnHyohonTorikomi As Button
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents BtnLinkageCsv As Button
    Friend WithEvents BtnMainteNogyotTiikiMaster As Button
End Class
