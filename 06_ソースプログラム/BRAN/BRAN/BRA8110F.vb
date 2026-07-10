''' <summary>
''' 毎勤データ都道府県選択画面
''' </summary>
''' <remarks></remarks>
Public Class BRA8110F

    Private Sub BRA8110F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            cboTodofuken.DataSource = MasterDao.GetJimusyoMaster(CommonInfo.Kyoku)
            cboTodofuken.ValueMember = "事務所番号"
            cboTodofuken.DisplayMember = "事務所名"
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        Dim xlForm As BRA8120X = Nothing

        Try

            xlForm = New BRA8120X(Me, cboTodofuken.SelectedValue.ToString)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally

        End Try
    End Sub

    Private Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click
        Try

            Dim jimusho As String = cboTodofuken.SelectedValue.ToString

            Dim fileName As String = ComConst.毎月勤労統計.出力用ファイル名称.reportName & "_" & Integer.Parse(jimusho).ToString("00") & ".xlsx"

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '毎月勤労統計の月別データ出力
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New BRA8110R(jimusho, filePath)
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
End Class
