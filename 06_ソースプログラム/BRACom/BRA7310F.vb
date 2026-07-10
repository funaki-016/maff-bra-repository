''' <summary>
''' 集計結果表還元画面
''' </summary>
''' <remarks></remarks>
Public Class BRA7310F

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
    Private Sub BRA7310F_Load(sender As Object, e As EventArgs) Handles Me.Load
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

    Private Sub btnReduce_Click(sender As Object, e As EventArgs) Handles btnReduce.Click
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

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
            If Message.ShowMsgBox(MessageID.MSG_Q_031, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                '調査区分取得
                Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

                Dim pKey As DAOSyukeiKekkahyo.PrimaryKey = sKey.Item1
                Dim kkey As DAOSyukeiKekkahyo.KomokuKey = sKey.Item2

                'キー取得
                Dim dc As Dictionary(Of String, String) = Me.GetKey(pKey, kkey)

                Dim chosaKouteiArr As String() = {}
                Select Case CommonInfo.Koutei
                    Case CommonInfo.KouteiKubun.Code.Honsyo
                        chosaKouteiArr = {CommonInfo.ChosaID & CommonInfo.KouteiKubun.Code.Kyoku, CommonInfo.ChosaID & CommonInfo.KouteiKubun.Code.Center}
                    Case CommonInfo.KouteiKubun.Code.Kyoku
                        chosaKouteiArr = {CommonInfo.ChosaID & CommonInfo.KouteiKubun.Code.Center}
                End Select

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    Try

                        db.BeginTrans()

                        For Each chosaKoutei In chosaKouteiArr
                            '集計結果表データ削除（還元）
                            DAOSyukeiKekkahyo.DeleteTableKangen(db, chosaKoutei, chosakubun, pKey, kkey)

                            '集計結果表データ追加（還元）
                            DAOSyukeiKekkahyo.InsertTableKangen(db, chosaKoutei, chosakubun, pKey, kkey)
                        Next

                        '還元管理データ存在チェック
                        If DAOOther.CheckReduceManagementExist(db, chosakubun, dc) Then
                            '還元管理還元日時更新
                            DAOOther.UpdateReduceManagementSendDate(db, chosakubun, dc)
                        Else
                            '還元管理追加
                            DAOOther.InsertReduceManagement(db, chosakubun, dc)
                        End If

                        db.CommitTrans()

                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try
                End Using

                '完了メッセージ
                Message.ShowMsgBox(MessageID.MSG_I_027, MsgBoxStyle.OkOnly)

                '一覧表示
                Me.ShowList(_chosaNen, _einouKeieitai, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
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
                Dim kkey As DAOSyukeiKekkahyo.KomokuKey = New DAOSyukeiKekkahyo.KomokuKey(dgvList.Rows(i).Cells(6).Value.ToString, dgvList.Rows(i).Cells(7).Value.ToString, dgvList.Rows(i).Cells(8).Value.ToString)
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
        Dim dtSyukei As DataTable = Nothing
        Dim dtManage As DataTable

        Dim chosakubun As String = ComUtil.GetChosakubun(einouKeieitai)

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Select Case CommonInfo.Koutei
                Case CommonInfo.KouteiKubun.Code.Center
                    dtSyukei = DAOSyukeiKekkahyo.GetList(db, chosakubun, chosaNen, kyoku, jimusho, kyoten)
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    dtSyukei = DAOSyukeiKekkahyo.GetList(db, chosakubun, chosaNen, kyoku, Nothing, Nothing)
                Case CommonInfo.KouteiKubun.Code.Honsyo
                    dtSyukei = DAOSyukeiKekkahyo.GetList(db, chosakubun, chosaNen, Nothing, Nothing, Nothing)
            End Select

            dtManage = DAOOther.GetReduceManagement(db, chosakubun, chosaNen, kyoku, jimusho, kyoten)
        End Using

        dgvList.Rows.Clear()

        For Each row As DataRow In dtSyukei.Rows
            Dim syukeiNo As String = row("集計番号").ToString
            Dim reduceDate As String = Nothing

            Dim query = (From dr In dtManage Where dr("集計番号").ToString = syukeiNo).Take(1).ToArray
            If query.Any() Then
                reduceDate = query(0)("還元日時").ToString
            End If

            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(1).Value = syukeiNo
            dgvList.Rows(i).Cells(2).Value = row("集計名称").ToString
            dgvList.Rows(i).Cells(3).Value = DateTime.Parse(row("更新日付").ToString).ToString(ComConst.DATETIME_FORMAT)
            dgvList.Rows(i).Cells(4).Value = If(String.IsNullOrEmpty(reduceDate), reduceDate, DateTime.Parse(reduceDate).ToString(ComConst.DATETIME_FORMAT))
            dgvList.Rows(i).Cells(5).Value = row("集計条件").ToString
            dgvList.Rows(i).Cells(6).Value = row("農政局").ToString
            dgvList.Rows(i).Cells(7).Value = row("都道府県").ToString
            dgvList.Rows(i).Cells(8).Value = row("実査設置拠点").ToString

            dgvList.Rows(i).Cells(5).Style.Alignment = DataGridViewContentAlignment.MiddleLeft
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

    ''' <summary>
    ''' キー取得
    ''' </summary>
    ''' <param name="pKey"></param>
    ''' <param name="kkey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetKey(pKey As DAOSyukeiKekkahyo.PrimaryKey, kkey As DAOSyukeiKekkahyo.KomokuKey) As Dictionary(Of String, String)
        Dim ret As New Dictionary(Of String, String)

        ret("調査年") = pKey.chosaNen
        ret("集計番号") = pKey.syukeiNo
        ret("農政局") = kkey.kyoku
        ret("都道府県") = kkey.jimusho
        ret("実査設置拠点") = kkey.kyoten

        Return ret
    End Function
End Class
