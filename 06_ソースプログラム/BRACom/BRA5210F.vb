''' <summary>
''' 集計結果表データ管理画面
''' </summary>
''' <remarks></remarks>
Public Class BRA5210F

    ''' <summary>調査年</summary>
    Private _chosaNen As String
    ''' <summary>営農経営体区分</summary>
    Private _einouKeieitai As String

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA5210F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '営農経営体区分コンボボックス設定
            ComUtil.SetEinouKeieitaiComboBox(lblEinouKeieitai, cboEinouKeieitai)

            '調査区分取得
            Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

            If Not (CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計) Then
                '調査年コンボボックス設定
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    ComUtil.SyukeiKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, chosakubun, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
                End Using
            End If

            'DataGridView設定
            ComUtil.ConfigDgv(Me.dgvList)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        Try
            '調査年選択チェック
            If cboChosaNen.SelectedValue Is Nothing Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            _chosaNen = cboChosaNen.SelectedValue.ToString
            _einouKeieitai = If(cboEinouKeieitai.SelectedValue Is Nothing, Nothing, cboEinouKeieitai.SelectedValue.ToString)

            '一覧表示
            Me.ShowList(_chosaNen, _einouKeieitai, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnModify_Click(sender As Object, e As EventArgs) Handles btnModify.Click
        Try
            '集計結果表主キー取得
            Dim sKey As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey) = Me.GetSyukeiKekkahyoSelectKey(_chosaNen)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(sKey, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim frm As New BRA5220F(_chosaNen, sKey, _einouKeieitai)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 削除ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            '集計結果表主キー取得
            Dim sKey As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey) = Me.GetSyukeiKekkahyoSelectKey(_chosaNen)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(sKey, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_003, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                '調査区分取得
                Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

                Dim pKey As DAOSyukeiKekkahyo.PrimaryKey = sKey.Item1
                Dim kkey As DAOSyukeiKekkahyo.KomokuKey = sKey.Item2

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    Try
                        db.BeginTrans()

                        '集計結果表データ削除
                        DAOSyukeiKekkahyo.DeleteTable(db, chosakubun, pKey, kkey)

                        db.CommitTrans()

                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try
                End Using

                '完了メッセージ
                Message.ShowMsgBox(MessageID.MSG_I_005, MsgBoxStyle.OkOnly)

                '一覧表示
                Me.ShowList(_chosaNen, _einouKeieitai, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
            End If
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

    Private Sub cboEinouKeieitai_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboEinouKeieitai.SelectedIndexChanged
        dgvList.Rows.Clear()

        '調査区分取得
        Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

        '調査年コンボボックス設定
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            ComUtil.SyukeiKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, chosakubun, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
        End Using
    End Sub

    ''' <summary>
    ''' 集計結果表主キー取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSyukeiKekkahyoSelectKey(chosaNen As String) As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey)
        Dim ret As New ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey)

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                Dim pkey As DAOSyukeiKekkahyo.PrimaryKey = New DAOSyukeiKekkahyo.PrimaryKey(chosaNen, dgvList.Rows(i).Cells(1).Value.ToString)
                Dim kkey As DAOSyukeiKekkahyo.KomokuKey = New DAOSyukeiKekkahyo.KomokuKey(dgvList.Rows(i).Cells(5).Value.ToString, dgvList.Rows(i).Cells(6).Value.ToString, dgvList.Rows(i).Cells(7).Value.ToString)
                ret = ValueTuple.Create(pkey, kkey)

                Exit For
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <param name="chosakubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <remarks></remarks>
    Public Sub ShowList(chosaNen As String, einouKeieitai As String, kyoku As String, jimusho As String, kyoten As String)
        Dim dt As DataTable = Nothing

        Dim chosakubun As String = ComUtil.GetChosakubun(einouKeieitai)

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Select Case CommonInfo.Koutei
                Case CommonInfo.KouteiKubun.Code.Center
                    dt = DAOSyukeiKekkahyo.GetList(db, chosakubun, chosaNen, kyoku, jimusho, kyoten)
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    dt = DAOSyukeiKekkahyo.GetList(db, chosakubun, chosaNen, kyoku, Nothing, Nothing)
                Case CommonInfo.KouteiKubun.Code.Honsyo
                    dt = DAOSyukeiKekkahyo.GetList(db, chosakubun, chosaNen, Nothing, Nothing, Nothing)
            End Select
        End Using

        dgvList.Rows.Clear()

        For Each row As DataRow In dt.Rows
            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(1).Value = row("集計番号").ToString
            dgvList.Rows(i).Cells(2).Value = row("集計名称").ToString
            dgvList.Rows(i).Cells(3).Value = DateTime.Parse(row("更新日付").ToString).ToString(ComConst.DATETIME_FORMAT)
            dgvList.Rows(i).Cells(4).Value = row("集計条件").ToString
            dgvList.Rows(i).Cells(5).Value = row("農政局").ToString
            dgvList.Rows(i).Cells(6).Value = row("都道府県").ToString
            dgvList.Rows(i).Cells(7).Value = row("実査設置拠点").ToString

            dgvList.Rows(i).Cells(4).Style.Alignment = DataGridViewContentAlignment.MiddleLeft
        Next
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="keys"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Function CheckError(sKey As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey), ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        '集計番号選択チェック
        If sKey.Item1 Is Nothing Then
            msgId = MessageID.MSG_E_030
            Return ret
        End If

        ret = True

        Return ret
    End Function
End Class
