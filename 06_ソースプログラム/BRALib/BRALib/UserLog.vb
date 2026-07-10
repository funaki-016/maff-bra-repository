Public Class UserLog

    Private _sw As New System.Diagnostics.Stopwatch()
    Private _procName As String

    Public Sub New(procName As String)
        _procName = procName
    End Sub

    Public Sub LogStart()
        'ストップウォッチを開始する
        _sw.Start()

        'ユーザログ出力
        OutputLog.WriteUserLog(OutputLog.LogLevel.Info, _procName, OutputLog.LogType.SyoriStart)
    End Sub

    Public Sub LogStop()

        'ストップウォッチを止める
        _sw.Stop()

        'ユーザログ出力
        OutputLog.WriteUserLog(OutputLog.LogLevel.Info, _procName, OutputLog.LogType.SyoriEnd, _sw.Elapsed.TotalSeconds.ToString("#0"))
    End Sub
End Class
