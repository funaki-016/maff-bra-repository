''' <summary>
''' 調査情報入力（本省）画面
''' </summary>
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.10.26 |TSP)                | フェーズ3 要件No.2修正
'//  REV_002   | 2023.01.13 |大興電子通信        | フェーズ2 要件No.5
'//  REV_003   | 2023.10.30 |大興電子通信        | フェーズ2 要件No.1
'//  REV_004   | 2024.05.31 |大興電子通信        | 要件No.1
'//            |            |                    |
'//*************************************************************************************************
''' <remarks></remarks>
Public Class BRA0110F

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        Try
            '調査区分選択チェック
            If cboChosakubun.SelectedValue Is Nothing Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_001, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '区分１設定
            CommonInfo.Kubun1 = cboKubun1.SelectedValue.ToString
            '区分１名設定
            CommonInfo.Kubun1Name = ComConst.区分１.リスト(CommonInfo.Kubun1)

            '区分２設定
            CommonInfo.Kubun2 = cboKubun2.SelectedValue.ToString
            '区分２名設定
            CommonInfo.Kubun2Name = ComConst.区分２.リスト(CommonInfo.Kubun2).名称

            '調査区分設定
            CommonInfo.Chosakubun = cboChosakubun.SelectedValue.ToString
            '調査区分名設定
            CommonInfo.ChosakubunName = ComConst.調査区分.リスト(CommonInfo.Chosakubun).名称

            Dim frm As New BRA0210F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub BRA0110F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Me.btnReturn.Text = "終了"

            '区分１コンボボックス設定
            ComUtil.SetKubun1ComboBox(cboKubun1)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMainteDB_Click(sender As Object, e As EventArgs) Handles btnMainteDB.Click
        Try
            Dim frm As New BRA9110F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub cboKubun1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKubun1.SelectedIndexChanged
        Try
            '区分２コンボボックス設定
            ComUtil.SetKubun2ComboBox(cboKubun1, cboKubun2)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub cboKubun2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKubun2.SelectedIndexChanged
        Try
            '調査区分コンボボックス設定
            ComUtil.SetChosakubunComboBox(cboKubun1, cboKubun2, cboChosakubun)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    '--- REV_001 ADD START
    Private Sub btnTraceImport_Click(sender As Object, e As EventArgs) Handles btnTraceImport.Click
        Try
            Dim frm As New BRA1910F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnTraceMainte_Click(sender As Object, e As EventArgs) Handles btnTraceMainte.Click
        Try
            Dim frm As New BRA2010F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
    '--- REV_001 ADD END

    ''' <summary>
    ''' 標本リスト取込ボタンクリック(REV_002)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnHyohonTorikomi_Click(sender As Object, e As EventArgs) Handles BtnHyohonTorikomi.Click
        Try
            Dim frm As New BRA10110F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 標本リスト管理ボタンクリック(REV_002)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnHyohonKanri_Click(sender As Object, e As EventArgs) Handles BtnHyohonKanri.Click
        Try
            Dim frm As New BRA10210F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 抽出条件設定ボタンクリック(REV_002)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnChushutsuJoken_Click(sender As Object, e As EventArgs) Handles BtnChushutsuJoken.Click
        Try
            Dim frm As New BRA10310F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 予定経営体数一覧表ボタンクリック(REV_002)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnHyohonPrint_Click(sender As Object, e As EventArgs) Handles BtnHyohonPrint.Click
        Try
            Dim frm As New BRA10410F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 標本リスト受信ボタンクリック(REV_002)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnHyohonJushin_Click(sender As Object, e As EventArgs) Handles BtnHyohonJushin.Click
        Try
            Dim frm As New BRA10510F(HyohonConst.標本リスト送受信区分.受信, ComConst.上り下り区分.上位工程への送信)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' [米在庫リンケージ]CSV出力ボタンクリック(REV_003)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnLinkageCsv_Click(sender As Object, e As EventArgs) Handles BtnLinkageCsv.Click
        Try
            Dim frm As New BRA11010F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    'REV_004 ADD START ---------------
    ''' <summary>
    ''' [メンテナンス]農業地域類型マスタ管理ボタンクリック(REV_004)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnMainteMaster_Click(sender As Object, e As EventArgs) Handles BtnMainteNogyotTiikiMaster.Click
        Try
            Dim frm As New BRA9610F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
    'REV_004 ADD END ---------------
End Class
