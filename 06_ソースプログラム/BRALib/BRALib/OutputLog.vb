Imports ToukeiCommon.LogManager

Public Class OutputLog

    ''' <summary>ログレベル</summary>
    Public Enum LogLevel
        ''' <summary>デバッグ</summary>
        Debug = ToukeiCommon.LogManager.LogLevel.Debug

        ''' <summary>情報</summary>
        Info = ToukeiCommon.LogManager.LogLevel.Info

        ''' <summary>警告</summary>
        Warn = ToukeiCommon.LogManager.LogLevel.Warn

        ''' <summary>エラー</summary>
        Err = ToukeiCommon.LogManager.LogLevel.Err

        ''' <summary>致命的</summary>
        Fatal = ToukeiCommon.LogManager.LogLevel.Fatal
    End Enum

    ''' <summary>ログ種別</summary>
    Public Enum LogType
        ''' <summary>処理開始</summary>
        SyoriStart

        ''' <summary>処理終了</summary>
        SyoriEnd
    End Enum

    ''' <summary>メッセージID：処理開始</summary>
    Public Const MSGID_SYORISTART As String = "I-00001"
    ''' <summary>メッセージID：処理終了</summary>
    Public Const MSGID_SYORIEND As String = "I-00002"
    ''' <summary>メッセージID：SQL</summary>
    Public Const MSGID_SQL As String = "I-00003"
    ''' <summary>メッセージID：エラー</summary>
    Public Const MSGID_ERROR As String = "E-00001"

    ''' <summary>メッセージ：処理開始</summary>
    Public Const MSG_SYORISTART As String = "処理を開始しました"
    ''' <summary>メッセージ：処理終了</summary>
    Public Const MSG_SYORIEND As String = "処理を終了しました"
    ''' <summary>メッセージ：SQL</summary>
    Public Const MSG_SQL As String = "SQLログ出力"
    ''' <summary>メッセージ：エラー</summary>
    Public Const MSG_ERROR As String = "処理が異常終了しました"

    ''' <summary>
    ''' 初期化処理
    ''' </summary>
    ''' <param name="userLogFilePath">ユーザログファイルパス</param>
    ''' <param name="sysLogFilePath">システムログファイルパス</param>
    ''' <param name="historyCount">履歴件数</param>
    ''' <param name="maxFileSize">出力制限サイズ</param>
    ''' <param name="userDefLogLevel">ユーザログデフォルトログレベル</param>
    ''' <param name="sysDefLogLevel">システムログデフォルトログレベル</param>
    ''' <param name="userId">ユーザID</param>
    ''' <param name="baseCode">拠点コード</param>
    ''' <param name="searchCode">調査ID</param>
    ''' <param name="opeCode">工程コード</param>
    ''' <param name="addInfo">ユーザ付加情報</param>
    ''' <remarks></remarks>
    Public Shared Sub Initialize(ByVal userLogFilePath As String, _
                                 ByVal sysLogFilePath As String, _
                                 ByVal historyCount As Integer, _
                                 ByVal maxFileSize As Integer, _
                                 ByVal userDefLogLevel As LogLevel, _
                                 ByVal sysDefLogLevel As LogLevel, _
                                 ByVal userId As String, _
                                 ByVal baseCode As String, _
                                 ByVal searchCode As String, _
                                 ByVal opeCode As String, _
                                 ByVal addInfo As String)


        Try

            'ユーザーログファイルパスのフォルダ作成
            If Not System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(userLogFilePath)) Then
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(userLogFilePath))
            End If

            'システムログファイルパスのフォルダ作成
            If Not System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(sysLogFilePath)) Then
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(sysLogFilePath))
            End If

            'ログ管理クラス初期化
            ToukeiCommon.LogManager.Initialize(userLogFilePath, _
                                               sysLogFilePath, _
                                               historyCount, _
                                               maxFileSize, _
                                               DirectCast(userDefLogLevel, ToukeiCommon.LogManager.LogLevel), _
                                               DirectCast(sysDefLogLevel, ToukeiCommon.LogManager.LogLevel), _
                                               userId, _
                                               baseCode, _
                                               searchCode, _
                                               opeCode, _
                                               addInfo)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' ユーザログ出力
    ''' </summary>
    ''' <param name="logLevel">ログレベル</param>
    ''' <param name="processName">処理名</param>
    ''' <param name="logType">ログ種別</param>
    ''' <param name="execSec">処理時間</param>
    ''' <param name="msgId">メッセージID</param>
    ''' <param name="msgText">メッセージ</param>
    ''' <remarks></remarks>
    Public Shared Sub WriteUserLog(ByVal logLevel As LogLevel, _
                                   ByVal processName As String, _
                                   ByVal logType As LogType, _
                                   Optional ByVal execSec As String = "", _
                                   Optional ByVal msgId As String = "", _
                                   Optional ByVal msgText As String = "")

        Try

            If msgId = "" Or msgText = "" Then
                Select Case (logType)
                    Case OutputLog.LogType.SyoriStart
                        msgId = MSGID_SYORISTART
                        msgText = MSG_SYORISTART
                    Case OutputLog.LogType.SyoriEnd
                        msgId = MSGID_SYORIEND
                        msgText = MSG_SYORIEND
                End Select
            End If

            If execSec <> "" Then
                '処理時間を追記
                msgText = String.Format("{0},{1}", msgText, execSec)
            End If

            'ユーザログ出力
            ToukeiCommon.LogManager.WriteUserLog(DirectCast(logLevel, ToukeiCommon.LogManager.LogLevel), _
                                                 processName, _
                                                 msgId, _
                                                 msgText, _
                                                 False)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' システムログ出力（例外情報無し）
    ''' </summary>
    ''' <param name="logLevel">ログ区分</param>
    ''' <param name="msgId">メッセージID</param>
    ''' <param name="msgText">メッセージ</param>
    ''' <remarks></remarks>
    Public Shared Sub WriteSystemLog(ByVal logLevel As LogLevel, _
                                     Optional ByVal msgId As String = "", _
                                     Optional ByVal msgText As String = "")

        Try

            If msgId = "" Or msgText = "" Then
                If logLevel = OutputLog.LogLevel.Err Then
                    msgId = MSGID_ERROR
                    msgText = MSG_ERROR
                End If
            End If

            'ユーザログ出力
            ToukeiCommon.LogManager.WriteSystemLog(DirectCast(logLevel, ToukeiCommon.LogManager.LogLevel), _
                                                   GetClassName(), _
                                                   msgId, _
                                                   msgText, _
                                                   False)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' システムログ出力（例外情報有り）
    ''' </summary>
    ''' <param name="logLevel">ログ区分</param>
    ''' <param name="exep">例外情報</param>
    ''' <param name="msgId">メッセージID</param>
    ''' <param name="msgText">メッセージ</param>
    ''' <remarks></remarks>
    Public Shared Sub WriteSystemLog(ByVal logLevel As LogLevel, _
                                     ByVal exep As Exception, _
                                     Optional ByVal msgId As String = "", _
                                     Optional ByVal msgText As String = "")

        Try

            If msgId = "" Or msgText = "" Then
                If logLevel = OutputLog.LogLevel.Err Then
                    msgId = MSGID_ERROR
                    msgText = MSG_ERROR
                End If
            End If

            'ユーザログ出力
            ToukeiCommon.LogManager.WriteSystemLog(DirectCast(logLevel, ToukeiCommon.LogManager.LogLevel), _
                                                   GetClassName(), _
                                                   msgId, _
                                                   msgText, _
                                                   exep, _
                                                   False)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' 実行中のクラス名を取得する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetClassName() As String
        Dim sf As New StackFrame(2)
        Return sf.GetMethod().DeclaringType.FullName
    End Function

End Class
