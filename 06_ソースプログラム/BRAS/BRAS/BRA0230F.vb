''' <summary>
''' 実査設置拠点工程メニュー
''' </summary>
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.07.16 |TSP)kawasuzaki      | フェーズ1 要件No.4修正　
'//  REV_002   | 2020.11.06 |TSP                 | フェーズ3 要件No.1、6修正　
'//  REV_003   | 2022.01.11 |日本ｺﾝﾋﾟｭｰﾀｼｽﾃﾑ     | 要件No1-②　
'//  REV_004   | 2022.01.18 |日本ｺﾝﾋﾟｭｰﾀｼｽﾃﾑ     | 労働時間整理ファイルメニュー　追加
'//  REV_005   | 2022.10.07 |大興電子通信        | 変更要件No.2 プレプリント（総括）ボタンにツールチップ表示
'//
'//*************************************************************************************************
''' <remarks>
''' </remarks>
Public Class BRA0230F

    ''' <summary>
    ''' ツールチップ(REV_005)
    ''' </summary>
    Dim ToolTip1 As ToolTip = New ToolTip()

    Private Sub BRA0230F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '専門調査員判定
            If CommonInfo.SenmonChosain Then
                btnMenu0101.Enabled = False
                btnMenu0201.Enabled = False
                btnMenu0202.Enabled = False
                btnMenu0301.Enabled = False
                btnMenu0302.Enabled = False
                '---REV_002 ADD START
                btnMenu0303.Enabled = False
                btnMenu0304.Enabled = False
                '---REV_002 ADD END
                '2022/1/29 ADD START 専門調査員なら「審査論理管理（範囲）」ボタンを非表示にする
                btnMenu0110.Enabled = False
                '2022/1/29 ADD END
            End If

            'ボタン非活性判定
            If Me.IsButtonDisable() Then
                'btnMenu0102.Enabled = False 　'REV-003 コメントアウト
                btnMenu0104.Enabled = False
            End If

            '---REV_002 ADD START
            If Not ComConst.牛トレサデータ.調査区分分類変換テーブル.ContainsKey(CommonInfo.Chosakubun) Then
                btnMenu0107.Enabled = False
                btnMenu0108.Enabled = False
                btnMenu0109.Enabled = False
            End If
            '---REV_002 ADD END

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0101_Click(sender As Object, e As EventArgs) Handles btnMenu0101.Click
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

    Private Sub btnMenu0102_Click(sender As Object, e As EventArgs) Handles btnMenu0102.Click
        Try
            Dim frm As New BRA1110F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0103_Click(sender As Object, e As EventArgs) Handles btnMenu0103.Click
        Try
            Dim frm As New BRA1210F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0104_Click(sender As Object, e As EventArgs) Handles btnMenu0104.Click
        Try
            Dim frm As New BRA1310F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0105_Click(sender As Object, e As EventArgs) Handles btnMenu0105.Click
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

    Private Sub btnMenu0106_Click(sender As Object, e As EventArgs) Handles btnMenu0106.Click
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

    Private Sub btnMenu0110_Click(sender As Object, e As EventArgs) Handles btnMenu0110.Click
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

    Private Sub btnMenu0201_Click(sender As Object, e As EventArgs) Handles btnMenu0201.Click
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

    Private Sub btnMenu0202_Click(sender As Object, e As EventArgs) Handles btnMenu0202.Click
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

    Private Sub btnMenu0203_Click(sender As Object, e As EventArgs) Handles btnMenu0203.Click
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

    Private Sub btnMenu0204_Click(sender As Object, e As EventArgs) Handles btnMenu0204.Click
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

    Private Sub btnMenu0205_Click(sender As Object, e As EventArgs) Handles btnMenu0205.Click
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

    Private Sub btnMenu0206_Click(sender As Object, e As EventArgs) Handles btnMenu0206.Click
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

    Private Sub btnMenu0207_Click(sender As Object, e As EventArgs) Handles btnMenu0207.Click
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

    Private Sub btnMenu0301_Click(sender As Object, e As EventArgs) Handles btnMenu0301.Click
        Try
            Dim frm As New BRA7110F(ComConst.上り下り区分.上位工程への送信, ComConst.送受信データ種別.調査票_個別結果表)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0302_Click(sender As Object, e As EventArgs) Handles btnMenu0302.Click
        Try
            Dim frm As New BRA7210F(ComConst.上り下り区分.下位工程への送信, ComConst.送受信データ種別.調査票_個別結果表)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    '---REV_01 ADD START
    Private Sub btnMenu0401_Click(sender As Object, e As EventArgs) Handles btnMenu0401.Click
        Try
            Dim frm As New BRA7410F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
    '---REV_01 ADD END

    ''' <summary>
    ''' ボタン非活性判定
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsButtonDisable() As Boolean
        Dim ret As Boolean = False

        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            ret = True
        End If

        Return ret
    End Function

    '---REV_002 ADD START
    Private Sub btnMenu0107_Click(sender As Object, e As EventArgs) Handles btnMenu0107.Click
        Try
            Dim frm As New BRA2110F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0108_Click(sender As Object, e As EventArgs) Handles btnMenu0108.Click
        Try
            Dim frm As New BRA2210F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnMenu0109_Click(sender As Object, e As EventArgs) Handles btnMenu0109.Click
        Try
            Dim frm As New BRA2310F()
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
            Dim frm As New BRA7510F(ComConst.送受信区分.実査設置拠点_局へ_送信)
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
            Dim frm As New BRA7510F(ComConst.送受信区分.実査設置拠点_局から_受信)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    '---REV_004 ADD START
    Private Sub btnMenu0111_Click(sender As Object, e As EventArgs) Handles btnMenu0111.Click
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
    '---REV_004 ADD END

    '---REV_002 ADD END

    Private Sub btnMenu0112_Click(sender As Object, e As EventArgs) Handles btnMenu0112.Click
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
    ''' プレプリント（総括）作成ボタンのツールチップ表示(REV_005)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>ツールチップを即時表示するため、イベントで処理</remarks>
    Private Sub btnMenu0108_ShowToolTip(sender As Object, e As EventArgs) Handles btnMenu0108.MouseEnter, btnMenu0108.GotFocus
        Dim msg = "プレプリント(総括)作成は、新規調査対象経営体向けの機能となります。" + vbCrLf + "前年からの既存調査対象経営体に対しては「プレプリント調査票作成」を行ってください。"
        Dim x = GroupBox1.Location.X + GroupBox6.Location.X + btnMenu0108.Location.X
        Dim y = GroupBox1.Location.Y + GroupBox6.Location.Y + btnMenu0108.Location.Y + btnMenu0108.Height + 35
        ToolTip1.Show(msg, Me, x, y)
    End Sub

    ''' <summary>
    ''' プレプリント（総括）作成ボタンのツールチップ非表示(REV_005)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>ツールチップを即時表示するため、イベントで処理</remarks>
    Private Sub btnMenu0108_HideToolTip(sender As Object, e As EventArgs) Handles btnMenu0108.MouseLeave, btnMenu0108.LostFocus
        ToolTip1.Hide(Me)
    End Sub
End Class
