Imports System.Text.RegularExpressions

''' <summary>
''' 専門調査員情報入力・修正画面
''' </summary>
''' <remarks></remarks>
Public Class BRA8410F

    ''' <summary>編集モード種別</summary>
    Public Enum 編集モード種別
        新規 = 1
        修正
    End Enum

    ''' <summary>編集モード</summary>
    Private _editMode As 編集モード種別

    ''' <summary>ユーザID</summary>
    Private _userID As String

    '''
    Public Sub New(editMode As 編集モード種別, Optional userID As String = Nothing)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        _editMode = editMode
        _userID = userID

    End Sub

    Private Sub BRA8410F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            If _editMode = 編集モード種別.修正 Then

                Dim dt As DataTable

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    '専門調査員管理取得
                    dt = DAOOther.GetSenmonChosain(db, _userID)
                End Using

                For Each row As DataRow In dt.Rows
                    txtUserID.Text = row("ユーザーID").ToString
                    txtShimei.Text = row("氏名").ToString
                Next

                txtUserID.ReadOnly = True
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnRegistr_Click(sender As Object, e As EventArgs) Handles btnRegistr.Click
        Try
            Dim userID As String = txtUserID.Text
            Dim shimei As String = txtShimei.Text

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(userID, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_001, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    If _editMode = 編集モード種別.新規 Then
                        '専門調査員管理存在チェック
                        If DAOOther.CheckSenmonChosainExist(db, userID) Then
                            'エラーメッセージ
                            Message.ShowMsgBox(MessageID.MSG_E_005, MsgBoxStyle.OkOnly)
                            Exit Sub
                        Else
                            Dim inPath As String = IniFileInfo.SenmonchosainInPath(userID)
                            Dim outPath As String = IniFileInfo.SenmonchosainOutPath(userID)
                            Try
                                System.IO.Directory.CreateDirectory(inPath)
                                System.IO.Directory.CreateDirectory(outPath)
                            Catch ex As Exception
                                'エラーメッセージ
                                Message.ShowMsgBox(MessageID.MSG_E_020, {inPath & vbCrLf & outPath}, MsgBoxStyle.OkOnly)
                                Exit Sub
                            End Try
                            Try
                                db.BeginTrans()

                                '専門調査員管理追加
                                DAOOther.InsertSenmonChosain(db, userID, shimei)

                                db.CommitTrans()
                            Catch ex As Exception
                                db.RollBackTrans()
                                Throw ex
                            End Try
                        End If
                    Else
                        Try
                            db.BeginTrans()

                            '専門調査員管理更新
                            DAOOther.UpdateSenmonChosain(db, userID, shimei)

                            db.CommitTrans()
                        Catch ex As Exception
                            db.RollBackTrans()
                            Throw ex
                        End Try
                    End If
                End Using

                '完了メッセージ
                Message.ShowMsgBox(MessageID.MSG_I_001, MsgBoxStyle.OkOnly)

                '一覧表示
                DirectCast(Me.Owner, BRA8310F).ShowList()

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
    Private Function CheckError(userID As String, ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'ユーザーIDチェック
        If Not Regex.IsMatch(userID, "^[0-9a-zA-Z_-]+$") Then
            msgId = MessageID.MSG_E_012
            Return ret
        End If

        ret = True

        Return ret
    End Function
End Class
