<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA0130F
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
        Me.btnMainteChosain = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.BtnHyohonSoshin = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.BtnHyohonJushin = New System.Windows.Forms.Button()
        Me.BtnHyohonPrint = New System.Windows.Forms.Button()
        Me.BtnHyohonKanri = New System.Windows.Forms.Button()
        Me.BtnHyohonTorikomi = New System.Windows.Forms.Button()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.btnMainteChosain)
        Me.GroupBox3.Location = New System.Drawing.Point(90, 455)
        Me.GroupBox3.TabIndex = 7
        Me.GroupBox3.Controls.SetChildIndex(Me.btnMainteDB, 0)
        Me.GroupBox3.Controls.SetChildIndex(Me.btnMainteChosain, 0)
        '
        'btnMainteDB
        '
        Me.btnMainteDB.Location = New System.Drawing.Point(228, 34)
        Me.btnMainteDB.TabIndex = 1
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(349, 164)
        '
        'cboKubun2
        '
        '
        'cboKubun1
        '
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(90, 118)
        Me.GroupBox1.Size = New System.Drawing.Size(454, 206)
        Me.GroupBox1.TabIndex = 5
        '
        'lblKoutei
        '
        Me.lblKoutei.TabIndex = 4
        '
        'lblInformation3
        '
        Me.lblInformation3.TabIndex = 2
        '
        'btnReturn
        '
        Me.btnReturn.Location = New System.Drawing.Point(530, 547)
        Me.btnReturn.TabIndex = 8
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 1
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        '
        'btnMainteChosain
        '
        Me.btnMainteChosain.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnMainteChosain.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnMainteChosain.Location = New System.Drawing.Point(35, 34)
        Me.btnMainteChosain.Name = "btnMainteChosain"
        Me.btnMainteChosain.Size = New System.Drawing.Size(172, 28)
        Me.btnMainteChosain.TabIndex = 0
        Me.btnMainteChosain.Text = "専門調査員管理"
        Me.btnMainteChosain.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.GroupBox5)
        Me.GroupBox2.Controls.Add(Me.GroupBox4)
        Me.GroupBox2.Controls.Add(Me.BtnHyohonPrint)
        Me.GroupBox2.Controls.Add(Me.BtnHyohonKanri)
        Me.GroupBox2.Controls.Add(Me.BtnHyohonTorikomi)
        Me.GroupBox2.Location = New System.Drawing.Point(90, 330)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(454, 119)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "標本管理"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.BtnHyohonSoshin)
        Me.GroupBox5.Location = New System.Drawing.Point(305, 65)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(137, 45)
        Me.GroupBox5.TabIndex = 4
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "局へ"
        '
        'BtnHyohonSoshin
        '
        Me.BtnHyohonSoshin.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHyohonSoshin.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHyohonSoshin.Location = New System.Drawing.Point(6, 11)
        Me.BtnHyohonSoshin.Name = "BtnHyohonSoshin"
        Me.BtnHyohonSoshin.Size = New System.Drawing.Size(132, 28)
        Me.BtnHyohonSoshin.TabIndex = 0
        Me.BtnHyohonSoshin.Text = "標本リスト送信"
        Me.BtnHyohonSoshin.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.BtnHyohonJushin)
        Me.GroupBox4.Location = New System.Drawing.Point(305, 12)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(137, 45)
        Me.GroupBox4.TabIndex = 2
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "局から"
        '
        'BtnHyohonJushin
        '
        Me.BtnHyohonJushin.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHyohonJushin.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHyohonJushin.Location = New System.Drawing.Point(5, 11)
        Me.BtnHyohonJushin.Name = "BtnHyohonJushin"
        Me.BtnHyohonJushin.Size = New System.Drawing.Size(132, 28)
        Me.BtnHyohonJushin.TabIndex = 0
        Me.BtnHyohonJushin.Text = "標本リスト受信"
        Me.BtnHyohonJushin.UseVisualStyleBackColor = True
        '
        'BtnHyohonPrint
        '
        Me.BtnHyohonPrint.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHyohonPrint.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHyohonPrint.Location = New System.Drawing.Point(29, 76)
        Me.BtnHyohonPrint.Name = "BtnHyohonPrint"
        Me.BtnHyohonPrint.Size = New System.Drawing.Size(132, 28)
        Me.BtnHyohonPrint.TabIndex = 3
        Me.BtnHyohonPrint.Text = "予定経営体数一覧表"
        Me.BtnHyohonPrint.UseVisualStyleBackColor = True
        '
        'BtnHyohonKanri
        '
        Me.BtnHyohonKanri.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHyohonKanri.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHyohonKanri.Location = New System.Drawing.Point(167, 23)
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
        Me.BtnHyohonTorikomi.Location = New System.Drawing.Point(29, 23)
        Me.BtnHyohonTorikomi.Name = "BtnHyohonTorikomi"
        Me.BtnHyohonTorikomi.Size = New System.Drawing.Size(132, 28)
        Me.BtnHyohonTorikomi.TabIndex = 0
        Me.BtnHyohonTorikomi.Text = "標本リスト取込"
        Me.BtnHyohonTorikomi.UseVisualStyleBackColor = True
        '
        'BRA0130F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(634, 591)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "BRA0130F"
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox3, 0)
        Me.Controls.SetChildIndex(Me.GroupBox2, 0)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnMainteChosain As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents BtnHyohonKanri As Button
    Friend WithEvents BtnHyohonTorikomi As Button
    Friend WithEvents BtnHyohonSoshin As Button
    Friend WithEvents BtnHyohonPrint As Button
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents BtnHyohonJushin As Button
    Friend WithEvents GroupBox5 As GroupBox
End Class
