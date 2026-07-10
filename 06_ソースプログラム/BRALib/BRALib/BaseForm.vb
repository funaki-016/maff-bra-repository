Imports System.Windows.Forms

Public Class BaseForm

    ''' <summary>ストップウォッチ(UserLog出力用)</summary>
    Private sw As New System.Diagnostics.Stopwatch()

    ''' <summary>処理名(UserLog出力用)</summary>
    ''' <remarks>UserLogに出力する処理名を変更したい場合のみ設定する</remarks>
    Protected _processName As String = String.Empty

    ''' <summary>
    ''' キー押下時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Overridable Sub BaseForm_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode And Keys.KeyCode) = Keys.Enter Then
            Do
                Me.ProcessTabKey(True)
                If Not (TypeOf Me.ActiveControl Is Button) Then
                    Exit Do
                End If
            Loop
        End If
    End Sub

    ''' <summary>
    ''' 画面表示後
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BaseForm_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown

        If _processName = String.Empty Then
            _processName = getSyoriName()
        End If

        'ストップウォッチを開始する
        sw.Start()

        'ユーザログ出力
        OutputLog.WriteUserLog(OutputLog.LogLevel.Info, _processName, OutputLog.LogType.SyoriStart)

    End Sub

    ''' <summary>
    ''' 画面終了時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BaseForm_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        'ストップウォッチを止める
        sw.Stop()

        'ユーザログ出力
        OutputLog.WriteUserLog(OutputLog.LogLevel.Info, _processName, OutputLog.LogType.SyoriEnd, sw.Elapsed.TotalSeconds.ToString("#0"))

    End Sub


    ''' <summary>
    ''' 処理名取得
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Function getSyoriName() As String
        Return Me.lblSyori.Text
    End Function

End Class