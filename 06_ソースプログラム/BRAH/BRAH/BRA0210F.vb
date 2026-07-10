''' <summary>
''' 本省工程メニュー
''' </summary>
''' <remarks></remarks>
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.11.20 |TSP)                | フェーズ3 要件No.6修正
'//  REV_002   | 2022.10.06 |大興電子通信        | 要件No.20 労働時間整理ファイル取込ボタン追加
'//  REV_003   | 2022.12.20 |大興電子通信        | 要件No.15 結果検討表作成論理管理各工程共通化
'//  REV_004   | 2023.01.13 |大興電子通信        | 要件No.4－⑦ 営農の集計結果検討表
'//  REV_005   | 2023.10.30 |大興電子通信        | 要件No.9 集計倍率管理
'//            |            |                    |
'//*************************************************************************************************
Public Class BRA0210F

    Private Sub BRA0210F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'ボタン判定
            If Me.IsButton({ComConst.調査区分.営農類型別経営統計_個人}) Then
                btnMenu0403.Enabled = False
                btnMenu0404.Enabled = False
            End If

            'REV_004↓
            'If Not Me.IsButton({ComConst.調査区分.営農類型別経営統計_個人, ComConst.調査区分.営農類型別経営統計_法人}) Then
            '    btnMenu0402.Enabled = False
            '    btnMenu0410.Enabled = False
            'End If
            'REV_004↑

            If Me.IsButton({ComConst.調査区分.営農類型別経営統計_個人, ComConst.調査区分.営農類型別経営統計_法人, ComConst.調査区分.米生産費統計_個別, ComConst.調査区分.小麦生産費統計_個別, ComConst.調査区分.二条大麦生産費統計_個別, ComConst.調査区分.六条大麦生産費統計_個別, ComConst.調査区分.はだか麦生産費統計_個別, ComConst.調査区分.そば生産費統計_個別, ComConst.調査区分.大豆生産費統計_個別, ComConst.調査区分.原料用かんしょ生産費統計_個別, ComConst.調査区分.原料用ばれいしょ生産費統計_個別, ComConst.調査区分.なたね生産費統計_個別, ComConst.調査区分.てんさい生産費統計_個別, ComConst.調査区分.さとうきび生産費統計_個別, ComConst.調査区分.米生産費統計_組織法人, ComConst.調査区分.小麦生産費統計_組織法人, ComConst.調査区分.大豆生産費統計_組織法人, ComConst.調査区分.経営分析調査_二条大麦生産費, ComConst.調査区分.経営分析調査_六条大麦生産費, ComConst.調査区分.経営分析調査_はだか麦生産費, ComConst.調査区分.経営分析調査_そば生産費, ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, ComConst.調査区分.経営分析調査_なたね生産費, ComConst.調査区分.経営分析調査_てんさい生産費, ComConst.調査区分.経営分析調査_さとうきび生産費
}) Then
                btnMenu0206.Enabled = False

            End If

            'REV_005↓
            '営農個人、営農法人の場合
            If Me.IsButton({ComConst.調査区分.営農類型別経営統計_個人, ComConst.調査区分.営農類型別経営統計_法人}) Then
                btnMenu0314.Enabled = False
            End If
            'REV_005↑

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0201_Click(sender As Object, e As EventArgs) Handles btnMenu0201.Click
        Try
            Dim frm As New BRA1710F(ComConst.エラーチェック種別.enm.基本)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0203_Click(sender As Object, e As EventArgs) Handles btnMenu0203.Click
        Try
            Dim frm As New BRA1410F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0204_Click(sender As Object, e As EventArgs) Handles btnMenu0204.Click
        Try
            Dim frm As New BRA9210F(BRA9210F.帳票種別.電子調査票)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0205_Click(sender As Object, e As EventArgs) Handles btnMenu0205.Click
        Try
            Dim frm As New BRA2410F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub


    Private Sub btnMenu0311_Click(sender As Object, e As EventArgs) Handles btnMenu0311.Click
        Try
            Dim frm As New BRA9210F(BRA9210F.帳票種別.個別結果表)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0312_Click(sender As Object, e As EventArgs) Handles btnMenu0312.Click
        Try
            Dim frm As New BRA9210F(BRA9210F.帳票種別.個別結果検討表)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0202_Click(sender As Object, e As EventArgs) Handles btnMenu0202.Click
        Try
            Dim frm As New BRA1710F(ComConst.エラーチェック種別.enm.追加)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0306_Click(sender As Object, e As EventArgs) Handles btnMenu0306.Click
        Try
            Dim frm As New BRA3110F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0308_Click(sender As Object, e As EventArgs) Handles btnMenu0308.Click
        Try
            Dim frm As New BRA3210F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0309_Click(sender As Object, e As EventArgs) Handles btnMenu0309.Click
        Try
            Dim frm As New BRA4010F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0310_Click(sender As Object, e As EventArgs) Handles btnMenu0310.Click
        Try
            Dim frm As New BRA3310F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0301_Click(sender As Object, e As EventArgs) Handles btnMenu0301.Click
        Try
            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                Dim frm As New BRA3411F()
                frm.Show(Me)
                Me.Hide()
            Else
                Dim frm As New BRA3410F()
                frm.Show(Me)
                Me.Hide()
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0302_Click(sender As Object, e As EventArgs) Handles btnMenu0302.Click
        Try
            Dim frm As New BRA3510F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0307_Click(sender As Object, e As EventArgs) Handles btnMenu0307.Click
        Try
            Dim frm As New BRA3610F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0303_Click(sender As Object, e As EventArgs) Handles btnMenu0303.Click
        Try
            Dim frm As New BRA3710F(ComConst.エラーチェック種別.enm.基本)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0305_Click(sender As Object, e As EventArgs) Handles btnMenu0305.Click
        Try
            Dim frm As New BRA3710F(ComConst.エラーチェック種別.enm.追加)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0304_Click(sender As Object, e As EventArgs) Handles btnMenu0304.Click
        Try
            Dim frm As New BRA3810F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0101_Click(sender As Object, e As EventArgs) Handles btnMenu0101.Click
        Try
            Dim frm As New BRA7210F(ComConst.上り下り区分.上位工程への送信, ComConst.送受信データ種別.調査票_個別結果表)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0501_Click(sender As Object, e As EventArgs) Handles btnMenu0501.Click
        Try
            Dim frm As New BRA7110F(ComConst.上り下り区分.下位工程への送信, ComConst.送受信データ種別.調査票_個別結果表)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0405_Click(sender As Object, e As EventArgs) Handles btnMenu0405.Click
        Try
            Dim frm As New BRA5110F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0401_Click(sender As Object, e As EventArgs) Handles btnMenu0401.Click
        Try
            Dim frm As New BRA5310F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0402_Click(sender As Object, e As EventArgs) Handles btnMenu0402.Click
        Try
            ' REV_003↓
            'Dim frm As New BRA5410F()
            Dim frm As New BRA5410F(ComConst.集計結果検討表作成論理.論理種別.集計結果検討表)
            ' REV_003↑ 
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0411_Click(sender As Object, e As EventArgs) Handles btnMenu0411.Click
        Try
            Dim frm As New BRA5510F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0409_Click(sender As Object, e As EventArgs) Handles btnMenu0409.Click
        Try
            Dim frm As New BRA9310F(BRA9330F.帳票種別.集計結果表)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0410_Click(sender As Object, e As EventArgs) Handles btnMenu0410.Click
        Try
            Dim frm As New BRA9310F(BRA9330F.帳票種別.集計結果検討表)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0408_Click(sender As Object, e As EventArgs) Handles btnMenu0408.Click
        Try
            Dim frm As New BRA5610F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0406_Click(sender As Object, e As EventArgs) Handles btnMenu0406.Click
        Try
            Dim frm As New BRA5210F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0502_Click(sender As Object, e As EventArgs) Handles btnMenu0502.Click
        Try
            Dim frm As New BRA7310F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0403_Click(sender As Object, e As EventArgs) Handles btnMenu0403.Click
        Try
            Dim frm As New BRA5710F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0404_Click(sender As Object, e As EventArgs) Handles btnMenu0404.Click
        Try
            Dim frm As New BRA5810F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0407_Click(sender As Object, e As EventArgs) Handles btnMenu0407.Click
        Try
            Dim frm As New BRA5910F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' ボタン判定
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsButton(chosakubun As String()) As Boolean
        Dim ret As Boolean = False

        If Not chosakubun.Contains(CommonInfo.Chosakubun) Then
            ret = True
        End If

        Return ret
    End Function

    Private Sub btnMenu0503_Click(sender As Object, e As EventArgs) Handles btnMenu0503.Click
        Try
            Dim frm As New BRA7510F(ComConst.送受信区分.本省_局へ_送信)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    '---REV_002 ADD START
    Private Sub btnMenu0102_Click(sender As Object, e As EventArgs) Handles btnMenu0102.Click
        Try
            Dim frm As New BRA7510F(ComConst.送受信区分.本省_局から_受信)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0313_Click(sender As Object, e As EventArgs) Handles btnMenu0313.Click
        Try
            Dim frm As New BRA4110F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
    '---REV_002 ADD END

    Private Sub btnMenu0206_Click(sender As Object, e As EventArgs) Handles btnMenu0206.Click
        Try
            Dim frm As New BRA2610F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    '---REV_002 ADD END

    Private Sub btnMenu0207_Click(sender As Object, e As EventArgs) Handles btnMenu0207.Click
        Try
            Dim frm As New BRA9410F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 労働時間整理ファイル取込ボタンクリック(REV_002)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnMenu0208_Click(sender As Object, e As EventArgs) Handles btnMenu0208.Click
        Try
            Dim frm As New BRA2710F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 集計倍率管理ボタンクリック(REV_005)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnMenu0314_Click(sender As Object, e As EventArgs) Handles btnMenu0314.Click
        Try
            Dim frm As New BRA4210F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
End Class
