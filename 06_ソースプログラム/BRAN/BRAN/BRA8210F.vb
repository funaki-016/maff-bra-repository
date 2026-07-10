Imports System.Text.RegularExpressions

''' <summary>
''' 労賃単価都道府県選択画面
''' </summary>
''' <remarks></remarks>
Public Class BRA8210F

    Private Sub BRA8210F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Dim lstCmbItems As New ArrayList()

            For Each Key In ComConst.生産費区分.リスト.Keys
                lstCmbItems.Add(New DictionaryEntry(Key, ComConst.生産費区分.リスト(Key)))
            Next

            cboSeisanhi.DisplayMember = "Value"
            cboSeisanhi.ValueMember = "Key"
            cboSeisanhi.DataSource = lstCmbItems
            
            cboTodofuken.DataSource = MasterDao.GetJimusyoMaster(CommonInfo.Kyoku)
            cboTodofuken.ValueMember = "事務所番号"
            cboTodofuken.DisplayMember = "事務所名"
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 選択ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        Try
            Dim seisanhi As String = cboSeisanhi.SelectedValue.ToString
            Dim chosaNen As String = txtYear.Text
            Dim jimusho As String = cboTodofuken.SelectedValue.ToString
            Dim judgmen As Boolean = False

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(chosaNen, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '毎勤データ存在チェック
            For i As Integer = 0 To 1
                Dim term As Dictionary(Of String, String) = ComUtil.RouchinTanka.GetTargetTerm(seisanhi, (Integer.Parse(chosaNen) - i).ToString)

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    If Not DAOOther.GetMaitsukiKinrouToukei(db, jimusho, term(ComConst.労賃単価.TERM_LOWER), term(ComConst.労賃単価.TERM_UPPER)).Rows.Count = 12 Then
                        judgmen = True
                    End If
                End Using
            Next
            If judgmen = True Then
                If Message.ShowMsgBox(MessageID.MSG_Q_020, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If

            Dim frmBRA0904F As New BRA8220F(seisanhi, chosaNen, jimusho)
            frmBRA0904F.Show(Me)
            Me.Hide()

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click
        Try

            Dim seisanhi As String = cboSeisanhi.SelectedValue.ToString
            Dim chosaNen As String = txtYear.Text
            Dim jimusho As String = cboTodofuken.SelectedValue.ToString

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(chosaNen, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '労賃単価存在チェック
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                If Not DAOOther.GetRouchinTanka(db, seisanhi, chosaNen, jimusho).Rows.Count > 0 Then
                    Message.ShowMsgBox(MessageID.MSG_E_023, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
            End Using

            Dim fileName As String = ComConst.労賃単価.出力用ファイル名称.reportName & "_" _
                                     & Integer.Parse(seisanhi).ToString("00") & "_" _
                                     & Integer.Parse(chosaNen).ToString("0000") & "_" _
                                     & Integer.Parse(jimusho).ToString("00") & ".xlsx"
            
            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '労賃単価結果表出力
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New BRA8210R(seisanhi, chosaNen, jimusho, filePath)
                    ret = ExcelOutput.Execute(MessageID.MSG_Q_004)
                End Using

                If ret = ExcelOutputBaseClass.enmOutputReturn.OK Then
                    '完了メッセージ
                    Message.ShowMsgBox(MessageID.MSG_I_002, MsgBoxStyle.OkOnly)
                End If
            Catch ex As ExcelOutputBaseClass.SaveAsException
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_006, MsgBoxStyle.OkOnly)
            Catch ex As Exception
                Throw
            End Try

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="year"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(year As String, ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        '未入力チェック
        If year.Equals(String.Empty) Then
            msgId = MessageID.MSG_E_004
            Return ret
        End If

        '半角数字チェック
        If Not Regex.IsMatch(year, "^[0-9]+$") Then
            msgId = MessageID.MSG_E_022
            Return ret
        End If

        ret = True

        Return ret
    End Function
End Class
