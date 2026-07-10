'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2025.09.11 |GCU                 | 要件No8 一覧に調査対象経営体と担当名称追加
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 個別結果表入力・修正画面
''' </summary>
''' <remarks></remarks>
Public Class BRA3210F

    ''' <summary>調査年</summary>
    Private _chosaNen As String
    ''' <summary>局</summary>
    Private _kyoku As String
    ''' <summary>事務所</summary>
    Private _jimusho As String
    ''' <summary>拠点</summary>
    Private _kyoten As String
    ''' <summary>営農類型</summary>
    Private _einouRuikei As String
    ''' <summary>欠測値補完</summary>
    Private _kessokuchiHokan As String
    ''' <summary>貸借対照表</summary>
    Private _taishakuTaishohyo As String

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA3210F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '局コンボボックス設定
            ComUtil.SetKyokuComboBox(cboKyoku)

            '営農類型コンボボックス設定
            ComUtil.SetEinouRuikeiComboBox(lblEinouRuikei, cboEinouRuikei)

            '欠測値補完コンボボックス設定
            ComUtil.SetKessokuchiHokanComboBox(lblKessokuchiHokan, cboKessokuchiHokan)

            '貸借対照表区分コンボボックス設定
            ComUtil.SetTaishakuTaishohyoComboBox(lblTaishakuTaishohyo, cboTaishakuTaishohyo)

            If Not (CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 _
                And CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Honsyo) Then
                '調査年コンボボックス設定
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    ComUtil.KobetsuKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
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

    ''' <summary>
    ''' 表示ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        Try
            '調査年選択チェック
            If cboChosaNen.SelectedValue Is Nothing Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            _chosaNen = cboChosaNen.SelectedValue.ToString
            _kyoku = If(IsDBNull(cboKyoku.SelectedValue), Nothing, cboKyoku.SelectedValue.ToString)
            _jimusho = If(IsDBNull(cboKyoten.SelectedValue), Nothing, CStr(CType(cboKyoten.SelectedItem, DataRowView)("事務所番号")))
            _kyoten = If(IsDBNull(cboKyoten.SelectedValue), Nothing, cboKyoten.SelectedValue.ToString)
            _einouRuikei = If(cboEinouRuikei.SelectedValue Is Nothing, Nothing, cboEinouRuikei.SelectedValue.ToString)
            _kessokuchiHokan = ComUtil.GetkessokuchiHokana(cboKessokuchiHokan)
            _taishakuTaishohyo = If(cboTaishakuTaishohyo.SelectedValue Is Nothing, Nothing, cboTaishakuTaishohyo.SelectedValue.ToString)

            '一覧表示
            Me.ShowList(_chosaNen, _kyoku, _jimusho, _kyoten, _einouRuikei, _taishakuTaishohyo, _kessokuchiHokan)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 修正ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnModify_Click(sender As Object, e As EventArgs) Handles btnModify.Click
        Dim xlForm As BRA3220X = Nothing

        Try
            Dim pKey As DAOKobetsuKekkahyo.PrimaryKey
            Dim kKey As New DAOKobetsuKekkahyo.KyotenKey

            pKey = Me.GetKobetsuKekkahyoPrimaryKey(_chosaNen, kKey)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(pKey, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            xlForm = New BRA3220X(Me, pKey, kKey, _kessokuchiHokan)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 農政局コンボボックス選択時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboKyoku_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKyoku.SelectedIndexChanged
        Try
            '拠点コンボボックス設定
            ComUtil.SetKyotenComboBox(cboKyoku, cboKyoten)
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
    ''' 欠測値補完コンボボックス選択時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboKessokuchiHokan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKessokuchiHokan.SelectedIndexChanged
        Try
            dgvList.Rows.Clear()

            '欠測値補完取得
            Dim kessokuchiHokan As String = ComUtil.GetkessokuchiHokana(cboKessokuchiHokan)

            '調査年コンボボックス設定
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.KobetsuKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center, kessokuchiHokan)
            End Using
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 個別結果表主キー取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="kKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetKobetsuKekkahyoPrimaryKey(chosaNen As String, Optional ByRef kKey As DAOKobetsuKekkahyo.KyotenKey = Nothing) As DAOKobetsuKekkahyo.PrimaryKey
        Dim ret As DAOKobetsuKekkahyo.PrimaryKey = Nothing

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                ret = New DAOKobetsuKekkahyo.PrimaryKey(chosaNen, dgvList.Rows(i).Cells(7).Value.ToString)
                If Not kKey Is Nothing Then
                    kKey.kyoku = dgvList.Rows(i).Cells(9).Value.ToString
                    kKey.jimusho = dgvList.Rows(i).Cells(10).Value.ToString
                    kKey.kyoten = dgvList.Rows(i).Cells(11).Value.ToString
                End If
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="einouRuikei"></param>
    ''' <param name="taishakuTaishohyo"></param>
    ''' <param name="kessokuchiHokan"></param>
    ''' <remarks></remarks>
    Private Sub ShowList(chosaNen As String, kyoku As String, jimusho As String, kyoten As String, einouRuikei As String, taishakuTaishohyo As String, kessokuchiHokan As String)
        Dim dt As DataTable
        Dim senmonName As String = ""

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dt = DAOKobetsuKekkahyo.GetList(db, chosaNen, kyoku, jimusho, kyoten, einouRuikei, taishakuTaishohyo, kessokuchiHokan)
            senmonName = DAOChosahyo.GetSenmonChosainName(db)
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
            dgvList.Rows(i).Cells(8).Value = DateTime.Parse(row("更新日付").ToString).ToString(ComConst.DATETIME_FORMAT)
            dgvList.Rows(i).Cells(9).Value = row("農政局").ToString
            dgvList.Rows(i).Cells(10).Value = row("都道府県").ToString
            dgvList.Rows(i).Cells(11).Value = row("実査設置拠点").ToString
            'REV_001
            dgvList.Rows(i).Cells(12).Value = row("Q00000701").ToString
            If CommonInfo.SenmonChosain Then
                dgvList.Rows(i).Cells(13).Value = senmonName
            Else
                dgvList.Rows(i).Cells(13).Value = row("Q00000801").ToString
            End If

        Next
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="pKey"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(pKey As DAOKobetsuKekkahyo.PrimaryKey, ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'センサス番号選択チェック
        If pKey Is Nothing Then
            msgId = MessageID.MSG_E_003
            Return ret
        End If

        ret = True

        Return ret
    End Function
End Class
