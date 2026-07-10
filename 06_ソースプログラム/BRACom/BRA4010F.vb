''' <summary>
''' 個別結果表項目指定修正画面
''' </summary>
''' <remarks></remarks>
Public Class BRA4010F

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
    Private Sub BRA4010F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '調査年コンボボックス設定
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.KobetsuKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
            End Using

            '局コンボボックス設定
            ComUtil.SetKyokuComboBox(cboKyoku)

            '営農類型コンボボックス設定
            ComUtil.SetEinouRuikeiComboBox(lblEinouRuikei, cboEinouRuikei)

            '欠測値補完コンボボックス設定
            ComUtil.SetKessokuchiHokanComboBox(lblKessokuchiHokan, cboKessokuchiHokan)

            '貸借対照表区分コンボボックス設定
            ComUtil.SetTaishakuTaishohyoComboBox(lblTaishakuTaishohyo, cboTaishakuTaishohyo)

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
    ''' 全選択ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAllSelect_Click(sender As Object, e As EventArgs) Handles btnAllSelect.Click
        Try
            ComUtil.SetDataGridViewAllCheck(dgvList, True)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 全解除ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAllCancel_Click(sender As Object, e As EventArgs) Handles btnAllCancel.Click
        Try
            ComUtil.SetDataGridViewAllCheck(dgvList, False)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 修正ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnModify_Click(sender As Object, e As EventArgs) Handles btnModify.Click
        Dim xlForm As BRA4020X = Nothing
        Try
            Dim censusNo As New List(Of String)

            For i As Integer = 0 To dgvList.Rows.Count - 1
                If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                    censusNo.Add(dgvList.Rows(i).Cells(7).Value.ToString)
                End If
            Next

            If _kessokuchiHokan Is Nothing Then
                _kessokuchiHokan = ComUtil.GetkessokuchiHokana(cboKessokuchiHokan)
            End If

            xlForm = New BRA4020X(Me, _chosaNen, censusNo, _kessokuchiHokan)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
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
            _kessokuchiHokan = ComUtil.GetkessokuchiHokana(cboKessokuchiHokan)

            '調査年コンボボックス設定
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.KobetsuKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center, _kessokuchiHokan)
            End Using
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

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dt = DAOKobetsuKekkahyo.GetList(db, chosaNen, kyoku, jimusho, kyoten, einouRuikei, taishakuTaishohyo, kessokuchiHokan)
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
        Next
    End Sub
End Class
