<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BRA0120F
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
        Me.btnRouchinTanka = New System.Windows.Forms.Button()
        Me.btnMaikinData = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.BtnHyohonJushin = New System.Windows.Forms.Button()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.BtnHyohonSoshinUp = New System.Windows.Forms.Button()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.BtnHyohonSoshinDown = New System.Windows.Forms.Button()
        Me.BtnHyohonPrint = New System.Windows.Forms.Button()
        Me.BtnHyohonKanri = New System.Windows.Forms.Button()
        Me.BtnHyohonTorikomi = New System.Windows.Forms.Button()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox3
        '
        Me.GroupBox3.Location = New System.Drawing.Point(90, 555)
        Me.GroupBox3.TabIndex = 7
        '
        'btnMainteDB
        '
        '
        'btnSelect
        '
        '
        'cboKubun2
        '
        Me.cboKubun2.TabIndex = 4
        '
        'cboKubun1
        '
        Me.cboKubun1.TabIndex = 3
        '
        'btnReturn
        '
        Me.btnReturn.Location = New System.Drawing.Point(530, 647)
        Me.btnReturn.TabIndex = 8
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnRouchinTanka)
        Me.GroupBox2.Controls.Add(Me.btnMaikinData)
        Me.GroupBox2.Location = New System.Drawing.Point(90, 334)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(454, 86)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "労賃単価作成"
        '
        'btnRouchinTanka
        '
        Me.btnRouchinTanka.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnRouchinTanka.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnRouchinTanka.Location = New System.Drawing.Point(252, 33)
        Me.btnRouchinTanka.Name = "btnRouchinTanka"
        Me.btnRouchinTanka.Size = New System.Drawing.Size(172, 28)
        Me.btnRouchinTanka.TabIndex = 1
        Me.btnRouchinTanka.Text = "労賃単価算出"
        Me.btnRouchinTanka.UseVisualStyleBackColor = True
        '
        'btnMaikinData
        '
        Me.btnMaikinData.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnMaikinData.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnMaikinData.Location = New System.Drawing.Point(28, 33)
        Me.btnMaikinData.Name = "btnMaikinData"
        Me.btnMaikinData.Size = New System.Drawing.Size(172, 28)
        Me.btnMaikinData.TabIndex = 0
        Me.btnMaikinData.Text = "毎勤データ入力・修正"
        Me.btnMaikinData.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.GroupBox7)
        Me.GroupBox4.Controls.Add(Me.GroupBox6)
        Me.GroupBox4.Controls.Add(Me.GroupBox5)
        Me.GroupBox4.Controls.Add(Me.BtnHyohonPrint)
        Me.GroupBox4.Controls.Add(Me.BtnHyohonKanri)
        Me.GroupBox4.Controls.Add(Me.BtnHyohonTorikomi)
        Me.GroupBox4.Location = New System.Drawing.Point(90, 426)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(454, 123)
        Me.GroupBox4.TabIndex = 6
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "標本管理"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.BtnHyohonJushin)
        Me.GroupBox7.Location = New System.Drawing.Point(156, 65)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(143, 49)
        Me.GroupBox7.TabIndex = 4
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "実査設置拠点から"
        '
        'BtnHyohonJushin
        '
        Me.BtnHyohonJushin.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHyohonJushin.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHyohonJushin.Location = New System.Drawing.Point(6, 15)
        Me.BtnHyohonJushin.Name = "BtnHyohonJushin"
        Me.BtnHyohonJushin.Size = New System.Drawing.Size(132, 28)
        Me.BtnHyohonJushin.TabIndex = 0
        Me.BtnHyohonJushin.Text = "標本リスト受信"
        Me.BtnHyohonJushin.UseVisualStyleBackColor = True
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.BtnHyohonSoshinUp)
        Me.GroupBox6.Location = New System.Drawing.Point(305, 65)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(143, 49)
        Me.GroupBox6.TabIndex = 5
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "本省へ"
        '
        'BtnHyohonSoshinUp
        '
        Me.BtnHyohonSoshinUp.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHyohonSoshinUp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHyohonSoshinUp.Location = New System.Drawing.Point(6, 15)
        Me.BtnHyohonSoshinUp.Name = "BtnHyohonSoshinUp"
        Me.BtnHyohonSoshinUp.Size = New System.Drawing.Size(132, 28)
        Me.BtnHyohonSoshinUp.TabIndex = 0
        Me.BtnHyohonSoshinUp.Text = "標本リスト送信"
        Me.BtnHyohonSoshinUp.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.BtnHyohonSoshinDown)
        Me.GroupBox5.Location = New System.Drawing.Point(305, 12)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(143, 49)
        Me.GroupBox5.TabIndex = 2
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "実査設置拠点へ"
        '
        'BtnHyohonSoshinDown
        '
        Me.BtnHyohonSoshinDown.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHyohonSoshinDown.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHyohonSoshinDown.Location = New System.Drawing.Point(6, 15)
        Me.BtnHyohonSoshinDown.Name = "BtnHyohonSoshinDown"
        Me.BtnHyohonSoshinDown.Size = New System.Drawing.Size(132, 28)
        Me.BtnHyohonSoshinDown.TabIndex = 0
        Me.BtnHyohonSoshinDown.Text = "標本リスト送信"
        Me.BtnHyohonSoshinDown.UseVisualStyleBackColor = True
        '
        'BtnHyohonPrint
        '
        Me.BtnHyohonPrint.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHyohonPrint.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHyohonPrint.Location = New System.Drawing.Point(15, 80)
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
        Me.BtnHyohonKanri.Location = New System.Drawing.Point(162, 27)
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
        Me.BtnHyohonTorikomi.Location = New System.Drawing.Point(15, 27)
        Me.BtnHyohonTorikomi.Name = "BtnHyohonTorikomi"
        Me.BtnHyohonTorikomi.Size = New System.Drawing.Size(132, 28)
        Me.BtnHyohonTorikomi.TabIndex = 0
        Me.BtnHyohonTorikomi.Text = "標本リスト取込"
        Me.BtnHyohonTorikomi.UseVisualStyleBackColor = True
        '
        'BRA0120F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(634, 691)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "BRA0120F"
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
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Protected WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Protected Friend WithEvents btnRouchinTanka As System.Windows.Forms.Button
    Protected Friend WithEvents btnMaikinData As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents BtnHyohonJushin As Button
    Friend WithEvents BtnHyohonSoshinUp As Button
    Friend WithEvents BtnHyohonPrint As Button
    Friend WithEvents BtnHyohonSoshinDown As Button
    Friend WithEvents BtnHyohonKanri As Button
    Friend WithEvents BtnHyohonTorikomi As Button
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents GroupBox7 As GroupBox
    Friend WithEvents GroupBox6 As GroupBox
End Class
