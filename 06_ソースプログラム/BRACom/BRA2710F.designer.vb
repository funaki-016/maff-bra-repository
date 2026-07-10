<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BRA2710F
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
        Me.btnImport = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboRoudoFileShurui = New System.Windows.Forms.ComboBox()
        Me.cboChosaNen = New System.Windows.Forms.ComboBox()
        Me.txtRoudoFile = New System.Windows.Forms.TextBox()
        Me.btnTmpOut = New System.Windows.Forms.Button()
        Me.btnReference = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboKyoku = New System.Windows.Forms.ComboBox()
        Me.cboKyoten = New System.Windows.Forms.ComboBox()
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
        Me.btnReturn.Location = New System.Drawing.Point(704, 219)
        Me.btnReturn.TabIndex = 18
        '
        'lblTitle
        '
        Me.lblTitle.TabIndex = 2
        '
        'lblSyori
        '
        Me.lblSyori.TabIndex = 3
        Me.lblSyori.Text = "労働時間整理ファイル取込"
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(588, 219)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(90, 30)
        Me.btnImport.TabIndex = 17
        Me.btnImport.Text = "取込"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(32, 97)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 12)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "調査年（産）"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(32, 132)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(135, 12)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "労働時間整理ファイル種類"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(32, 169)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(111, 12)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "労働時間整理ファイル"
        '
        'cboRoudoFileShurui
        '
        Me.cboRoudoFileShurui.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoudoFileShurui.FormattingEnabled = True
        Me.cboRoudoFileShurui.Items.AddRange(New Object() {"労働時間整理ファイル（営個、月別）", "労働時間整理ファイル（営個、日別）", "労働時間整理ファイル（営法、月別）", "労働時間整理ファイル（営法、日別）", "労働時間整理ファイル（畜生、月別）", "労働時間整理ファイル（畜生、日別）", "労働時間整理ファイル（農生、月別）", "労働時間整理ファイル（農生、日別）"})
        Me.cboRoudoFileShurui.Location = New System.Drawing.Point(185, 129)
        Me.cboRoudoFileShurui.Name = "cboRoudoFileShurui"
        Me.cboRoudoFileShurui.Size = New System.Drawing.Size(498, 20)
        Me.cboRoudoFileShurui.TabIndex = 12
        '
        'cboChosaNen
        '
        Me.cboChosaNen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChosaNen.FormattingEnabled = True
        Me.cboChosaNen.Location = New System.Drawing.Point(185, 94)
        Me.cboChosaNen.Name = "cboChosaNen"
        Me.cboChosaNen.Size = New System.Drawing.Size(72, 20)
        Me.cboChosaNen.TabIndex = 6
        '
        'txtRoudoFile
        '
        Me.txtRoudoFile.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtRoudoFile.Location = New System.Drawing.Point(185, 166)
        Me.txtRoudoFile.MaxLength = 4
        Me.txtRoudoFile.Name = "txtRoudoFile"
        Me.txtRoudoFile.Size = New System.Drawing.Size(498, 19)
        Me.txtRoudoFile.TabIndex = 15
        '
        'btnTmpOut
        '
        Me.btnTmpOut.Location = New System.Drawing.Point(704, 123)
        Me.btnTmpOut.Name = "btnTmpOut"
        Me.btnTmpOut.Size = New System.Drawing.Size(90, 30)
        Me.btnTmpOut.TabIndex = 13
        Me.btnTmpOut.Text = "ﾃﾝﾌﾟﾚｰﾄ出力"
        Me.btnTmpOut.UseVisualStyleBackColor = True
        '
        'btnReference
        '
        Me.btnReference.Location = New System.Drawing.Point(704, 160)
        Me.btnReference.Name = "btnReference"
        Me.btnReference.Size = New System.Drawing.Size(90, 30)
        Me.btnReference.TabIndex = 16
        Me.btnReference.Text = "参照"
        Me.btnReference.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(290, 97)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 12)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "農政局"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(466, 97)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 12)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "実査設置拠点"
        '
        'cboKyoku
        '
        Me.cboKyoku.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKyoku.FormattingEnabled = True
        Me.cboKyoku.Location = New System.Drawing.Point(337, 94)
        Me.cboKyoku.Name = "cboKyoku"
        Me.cboKyoku.Size = New System.Drawing.Size(110, 20)
        Me.cboKyoku.TabIndex = 8
        '
        'cboKyoten
        '
        Me.cboKyoten.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKyoten.FormattingEnabled = True
        Me.cboKyoten.Location = New System.Drawing.Point(549, 94)
        Me.cboKyoten.Name = "cboKyoten"
        Me.cboKyoten.Size = New System.Drawing.Size(72, 20)
        Me.cboKyoten.TabIndex = 10
        '
        'BRA2710F
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(808, 263)
        Me.Controls.Add(Me.cboKyoku)
        Me.Controls.Add(Me.cboKyoten)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnReference)
        Me.Controls.Add(Me.btnTmpOut)
        Me.Controls.Add(Me.txtRoudoFile)
        Me.Controls.Add(Me.cboChosaNen)
        Me.Controls.Add(Me.cboRoudoFileShurui)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.btnImport)
        Me.Name = "BRA2710F"
        Me.Text = "労働時間整理表取込"
        Me.Controls.SetChildIndex(Me.btnImport, 0)
        Me.Controls.SetChildIndex(Me.lblSyori, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblKoutei, 0)
        Me.Controls.SetChildIndex(Me.lblInformation2, 0)
        Me.Controls.SetChildIndex(Me.lblInformation3, 0)
        Me.Controls.SetChildIndex(Me.btnReturn, 0)
        Me.Controls.SetChildIndex(Me.Label10, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.cboRoudoFileShurui, 0)
        Me.Controls.SetChildIndex(Me.cboChosaNen, 0)
        Me.Controls.SetChildIndex(Me.txtRoudoFile, 0)
        Me.Controls.SetChildIndex(Me.btnTmpOut, 0)
        Me.Controls.SetChildIndex(Me.btnReference, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.cboKyoten, 0)
        Me.Controls.SetChildIndex(Me.cboKyoku, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Protected WithEvents cboRoudoFileShurui As ComboBox
    Protected WithEvents cboChosaNen As ComboBox
    Friend WithEvents txtRoudoFile As TextBox
    Friend WithEvents btnTmpOut As Button
    Friend WithEvents btnReference As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Protected WithEvents cboKyoku As ComboBox
    Protected WithEvents cboKyoten As ComboBox
End Class
