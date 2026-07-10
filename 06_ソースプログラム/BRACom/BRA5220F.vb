''' <summary>
''' 集計結果表名称変更画面
''' </summary>
''' <remarks></remarks>
Public Class BRA5220F

    ''' <summary>調査年</summary>
    Private _chosaNen As String
    ''' <summary>主キー</summary>
    Private _pkey As DAOSyukeiKekkahyo.PrimaryKey
    ''' <summary>項目キー</summary>
    Private _kkey As DAOSyukeiKekkahyo.KomokuKey
    ''' <summary>営農経営体区分</summary>
    Private _einouKeieitai As String
    ''' <summary>調査区分</summary>
    Private _chosakubun As String

    Public Sub New(chosaNen As String, sKey As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey), einouKeieitai As String)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        _chosaNen = chosaNen
        _pkey = sKey.Item1
        _kkey = sKey.Item2
        _einouKeieitai = einouKeieitai
        _chosakubun = ComUtil.GetChosakubun(_einouKeieitai)
    End Sub

    Private Sub BRA5220F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Dim syukeiName As String
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '集計名称取得
                syukeiName = DAOSyukeiKekkahyo.GetSyukeiName(db, _chosakubun, _pkey, _kkey)
            End Using

            txtBeforeName.Text = syukeiName
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnModify_Click(sender As Object, e As EventArgs) Handles btnModify.Click
        Try
            Dim afterName As String = txtAfterName.Text

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(afterName, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_001, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    DAOSyukeiKekkahyo.UpdateSyukeiName(db, _chosakubun, _pkey, _kkey, afterName)
                End Using

                '完了メッセージ
                Message.ShowMsgBox(MessageID.MSG_I_001, MsgBoxStyle.OkOnly)

                '一覧表示
                DirectCast(Me.Owner, BRA5210F).ShowList(_chosaNen, _einouKeieitai, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)

                Me.Close()
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="userID"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(afterName As String, ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'ユーザーIDチェック
        If afterName.Equals(String.Empty) Then
            msgId = MessageID.MSG_E_041
            Return ret
        End If

        ret = True

        Return ret
    End Function
End Class
