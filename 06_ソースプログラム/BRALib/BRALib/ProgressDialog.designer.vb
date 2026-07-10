<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ProgressDialog
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.MainProgressBar = New System.Windows.Forms.ProgressBar()
        Me.ProcessTargetLabel = New System.Windows.Forms.Label()
        Me.ProcessCountLabel = New System.Windows.Forms.Label()
        Me.ProgressBackgroundWorker = New System.ComponentModel.BackgroundWorker()
        Me.ProcessLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'MainProgressBar
        '
        Me.MainProgressBar.Location = New System.Drawing.Point(14, 33)
        Me.MainProgressBar.Name = "MainProgressBar"
        Me.MainProgressBar.Size = New System.Drawing.Size(400, 19)
        Me.MainProgressBar.TabIndex = 1
        '
        'ProcessTargetLabel
        '
        Me.ProcessTargetLabel.AutoSize = True
        Me.ProcessTargetLabel.Location = New System.Drawing.Point(14, 11)
        Me.ProcessTargetLabel.Margin = New System.Windows.Forms.Padding(0, 0, 0, 7)
        Me.ProcessTargetLabel.Name = "ProcessTargetLabel"
        Me.ProcessTargetLabel.Size = New System.Drawing.Size(59, 12)
        Me.ProcessTargetLabel.TabIndex = 0
        Me.ProcessTargetLabel.Text = "処理中・・・"
        '
        'ProcessCountLabel
        '
        Me.ProcessCountLabel.AutoSize = True
        Me.ProcessCountLabel.Location = New System.Drawing.Point(270, 77)
        Me.ProcessCountLabel.Margin = New System.Windows.Forms.Padding(0, 0, 7, 7)
        Me.ProcessCountLabel.Name = "ProcessCountLabel"
        Me.ProcessCountLabel.Size = New System.Drawing.Size(53, 12)
        Me.ProcessCountLabel.TabIndex = 2
        Me.ProcessCountLabel.Text = "処理件数"
        Me.ProcessCountLabel.Visible = False
        '
        'ProcessLabel
        '
        Me.ProcessLabel.Location = New System.Drawing.Point(330, 77)
        Me.ProcessLabel.Margin = New System.Windows.Forms.Padding(0, 0, 0, 7)
        Me.ProcessLabel.Name = "ProcessLabel"
        Me.ProcessLabel.Size = New System.Drawing.Size(87, 12)
        Me.ProcessLabel.TabIndex = 3
        Me.ProcessLabel.Text = "計算中・・・"
        Me.ProcessLabel.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ProcessLabel.Visible = False
        '
        'ProgressDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(428, 107)
        Me.ControlBox = False
        Me.Controls.Add(Me.ProcessLabel)
        Me.Controls.Add(Me.ProcessCountLabel)
        Me.Controls.Add(Me.ProcessTargetLabel)
        Me.Controls.Add(Me.MainProgressBar)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ProgressDialog"
        Me.Padding = New System.Windows.Forms.Padding(11)
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "進捗状況表示"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ProcessTargetLabel As System.Windows.Forms.Label
    Friend WithEvents ProcessCountLabel As System.Windows.Forms.Label
    Friend WithEvents ProgressBackgroundWorker As System.ComponentModel.BackgroundWorker
    Friend WithEvents ProcessLabel As System.Windows.Forms.Label
    Friend WithEvents MainProgressBar As System.Windows.Forms.ProgressBar
End Class
