Namespace My

    ' 次のイベントは MyApplication に対して利用できます:
    ' 
    ' Startup: アプリケーションが開始されたとき、スタートアップ フォームが作成される前に発生します。
    ' Shutdown: アプリケーション フォームがすべて閉じられた後に発生します。このイベントは、通常の終了以外の方法でアプリケーションが終了されたときには発生しません。
    ' UnhandledException: ハンドルされていない例外がアプリケーションで発生したときに発生するイベントです。
    ' StartupNextInstance: 単一インスタンス アプリケーションが起動され、それが既にアクティブであるときに発生します。
    ' NetworkAvailabilityChanged: ネットワーク接続が接続されたとき、または切断されたときに発生します。
    Partial Friend Class MyApplication
        ''' <summary>
        ''' アプリケーション開始時
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub MyApplication_Startup(sender As Object, e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup

            Try
                'ユーザ情報の取得
                Dim toukeilib As ToukeiLib.ToukeiClass = New ToukeiLib.ToukeiClass
                CommonInfo.Connect = My.Settings.COMMConnectionString
                CommonInfo.UserId = toukeilib.GetUserID()
                CommonInfo.UserOffice = toukeilib.GetUserOffice(CommonInfo.UserId)

                '局名設定
                CommonInfo.KyokuName = MasterDao.GetKyokuName(CommonInfo.Kyoku).Trim()
                '事務所名設定
                CommonInfo.JimusyoName = MasterDao.GetJimusyoName(CommonInfo.Jimusyo).Trim()
                'センター名設定
                CommonInfo.CenterName = MasterDao.GetCenterName(CommonInfo.Jimusyo, CommonInfo.Center).Trim()

                '設定ファイル名の取得
                IniFileInfo.IniFile = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & System.IO.Path.DirectorySeparatorChar & My.Settings.IniFile

                'ログ管理クラス初期化
                OutputLog.Initialize(IniFileInfo.LogUserLogFilePath(), _
                                     IniFileInfo.LogSystemLogFilePath(), _
                                     CType(IniFileInfo.LogHistoryCount(), Integer), _
                                     CType(IniFileInfo.LogMaxSize(), Integer), _
                                     CType(IniFileInfo.LogUserLogLevel(), OutputLog.LogLevel), _
                                     CType(IniFileInfo.LogSystemLogLevel(), OutputLog.LogLevel), _
                                     CommonInfo.UserId, _
                                     CommonInfo.UserOffice, _
                                     CommonInfo.ChosaID, _
                                     CommonInfo.Koutei, _
                                     "")

                'メッセージクラス初期化
                Message.Initialize(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & System.IO.Path.DirectorySeparatorChar & "Message.txt")

                '専門調査員設定
                If CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Center Then
                    Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                        CommonInfo.SenmonChosain = ComUtil.IsSenmonChosain(db, CommonInfo.UserId)
                    End Using
                End If
            Catch ex As Exception
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
                End
            End Try
        End Sub

    End Class


End Namespace

