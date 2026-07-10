Imports System.Text.RegularExpressions

''' <summary>
''' 専門調査員担当調査客体一覧画面
''' </summary>
''' <remarks></remarks>
Public Class BRA8510F

    ''' <summary>ユーザID</summary>
    Private _userID As String

    Public Sub New(userID As String)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        _userID = userID
    End Sub

    ''' <summary>
    ''' 画面起動
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA8510F_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Try
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '専門調査員管理取得
                txtChosainShimei.Text = DAOOther.GetSenmonChosainShimei(db, _userID)
            End Using

            'DataGridView設定
            ComUtil.ConfigDgv(Me.dgvList)

            '一覧表示
            Me.ShowList(_userID)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click
        Try
            Dim todofuken As String = txtTodofuken.Text
            Dim shikuchoson As String = txtShikuchoson.Text
            Dim kyuShikuchoson As String = txtKyuShikuchoson.Text
            Dim nogyoShuraku As String = txtNogyoShuraku.Text
            Dim chosaku As String = txtChosaku.Text
            Dim kyakutaiNo As String = txtKyakutaiNo.Text

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(todofuken, shikuchoson, kyuShikuchoson, nogyoShuraku, chosaku, kyakutaiNo, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            'センサス番号取得
            Dim censusNo As String = ComUtil.GetCensusNo(todofuken, shikuchoson, kyuShikuchoson, nogyoShuraku, chosaku, kyakutaiNo)

            '専門調査員担当調査客体ユーザーID取得
            Dim userID As List(Of String)
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                userID = DAOOther.GetSenmonChosainKyakutaiUserID(db, censusNo)
            End Using

            'ユーザID存在チェック
            If userID.Count > 0 Then
                If userID.Contains(_userID) Then
                    'エラーメッセージ
                    Message.ShowMsgBox(MessageID.MSG_E_015, MsgBoxStyle.OkOnly)
                    Exit Sub
                Else
                    '確認メッセージ
                    If Message.ShowMsgBox(MessageID.MSG_Q_008, {String.Join(",", userID)}, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
            End If

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try
                    db.BeginTrans()

                    '専門調査員担当調査客体追加
                    DAOOther.InsertSenmonChosainKyakutai(db, _userID, censusNo)

                    db.CommitTrans()
                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using

            '完了メッセージ
            Message.ShowMsgBox(MessageID.MSG_I_001, MsgBoxStyle.OkOnly)

            '一覧表示
            Me.ShowList(_userID)

            '一覧表示
            DirectCast(Me.Owner, BRA8310F).ShowList()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            'センサス番号取得
            Dim censusNo As String = Me.GetCensusNo()

            '削除エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(censusNo, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try
                    db.BeginTrans()

                    '専門調査員担当調査客体削除
                    DAOOther.DeleteSenmonChosainKyakutai(db, _userID, censusNo)

                    db.CommitTrans()
                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using

            '完了メッセージ
            Message.ShowMsgBox(MessageID.MSG_I_005, MsgBoxStyle.OkOnly)

            '一覧表示
            Me.ShowList(_userID)

            '一覧表示
            DirectCast(Me.Owner, BRA8310F).ShowList()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub dgvList_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvList.CellValueChanged
        If e.ColumnIndex = 0 Then
            If CType(dgvList(e.ColumnIndex, e.RowIndex).Value, Boolean) = True Then
                For rowIndex = 0 To dgvList.Rows.Count - 1
                    If rowIndex <> e.RowIndex Then
                        dgvList(0, rowIndex).Value = False
                        dgvList(0, rowIndex).ReadOnly = False
                    End If
                Next
                dgvList(e.ColumnIndex, e.RowIndex).ReadOnly = True
            End If
        End If
    End Sub

    Private Sub dgvList_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvList.CurrentCellDirtyStateChanged
        If dgvList.CurrentCellAddress.X = 0 AndAlso dgvList.IsCurrentCellDirty Then
            dgvList.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <param name="userID"></param>
    ''' <remarks></remarks>
    Public Sub ShowList(userID As String)
        Dim dt As DataTable

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dt = DAOOther.GetSenmonChosainKyakutai(db, userID)
        End Using

        dgvList.Rows.Clear()

        For Each row As DataRow In dt.Rows
            Dim censusNo As String = row("センサス番号").ToString

            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(1).Value = ComUtil.GetTodofuken(censusNo)
            dgvList.Rows(i).Cells(2).Value = ComUtil.GetShikuchoson(censusNo)
            dgvList.Rows(i).Cells(3).Value = ComUtil.GetKyuShikuchoson(censusNo)
            dgvList.Rows(i).Cells(4).Value = ComUtil.GetNogyoShuraku(censusNo)
            dgvList.Rows(i).Cells(5).Value = ComUtil.GetChosaku(censusNo)
            dgvList.Rows(i).Cells(6).Value = ComUtil.GetKyakutaiNo(censusNo)
            dgvList.Rows(i).Cells(7).Value = censusNo
        Next
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="todofuken"></param>
    ''' <param name="shikuchoson"></param>
    ''' <param name="kyuShikuchoson"></param>
    ''' <param name="nogyoShuraku"></param>
    ''' <param name="chosaku"></param>
    ''' <param name="kyakutaiNo"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(todofuken As String, shikuchoson As String, kyuShikuchoson As String, nogyoShuraku As String, chosaku As String, kyakutaiNo As String, ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        Dim arr As String() = {todofuken, shikuchoson, kyuShikuchoson, nogyoShuraku, chosaku, kyakutaiNo}

        '未入力チェック
        If arr.Contains(String.Empty) Then
            msgId = MessageID.MSG_E_013
            Return ret
        End If

        '半角数字チェック
        For Each val As String In arr
            If Not Regex.IsMatch(val, "^[0-9]+$") Then
                msgId = MessageID.MSG_E_021
                Return ret
            End If
        Next

        '都道府県チェック
        If Not Integer.Parse(ComUtil.ConvJimusyoNo(todofuken)).ToString("00").Equals(CommonInfo.Jimusyo) Then
            msgId = MessageID.MSG_E_014
            Return ret
        End If

        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' センサス番号取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCensusNo() As String
        Dim ret As String = Nothing

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                ret = dgvList.Rows(i).Cells(7).Value.ToString
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="censusNo"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(censusNo As String, ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'センサス番号チェック
        If censusNo Is Nothing Then
            msgId = MessageID.MSG_E_003
            Return ret
        End If

        ret = True

        Return ret
    End Function
End Class
