Imports Microsoft.Office.Interop

''' <summary>
''' 営農欠測値平均値作成画面
''' </summary>
''' <remarks></remarks>
Public Class BRA5710F

#Region "【処理詳細仕様 1】初期表示"

    ''' <summary>
    ''' フォームのロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA5710F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '【処理詳細仕様 1-2】調査年（産）設定

            '調査年コンボボックス設定
            ComUtil.SetChosaNenComboBox(cboChosaNen, ComConst.個別結果表.テーブル名称(ComConst.調査区分.営農類型別経営統計_個人)(0))

            '更新日時設定
            ComUtil.UpdateDate(txtUpdateDate, cboChosaNen.Text, ComConst.欠測値.営農欠測値適用平均値データ)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

#End Region

#Region "【処理詳細仕様 2】「調査年（産）」コンボボックス選択"

    ''' <summary>
    ''' 「調査年（産）」コンボボックスの選択を変更した
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboChosaNen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboChosaNen.SelectedIndexChanged

        '更新日時設定
        ComUtil.UpdateDate(txtUpdateDate, cboChosaNen.Text, ComConst.欠測値.営農欠測値適用平均値データ)

    End Sub

#End Region

#Region "【処理詳細仕様 3】「作成」ボタンクリック"

    ''' <summary>
    ''' 作成ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnMake_Click(sender As Object, e As EventArgs) Handles btnMake.Click

        Try
            If String.IsNullOrWhiteSpace(cboChosaNen.Text) Then
                Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
                Return
            End If

            If Message.ShowMsgBox(MessageID.MSG_Q_022, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If

            Dim progressDialog = New ProgressDialog

            Try
                '進捗ダイアログを表示する(処理詳細仕様 3-1～3-3をカウントの対象とする)
                progressDialog.Maximum = 3
                progressDialog.Show(Me)

                Using db As New DBAccess(My.Settings.BRAHConnectionString)
                    Try
                        db.BeginTrans()

                        '【処理詳細仕様 3-1】平均値算出データ削除
                        DAOOther.DeleteTable(db, cboChosaNen.Text, ComConst.欠測値.営農欠測値適用平均値データ)

                        '進捗を進める
                        progressDialog.AddValue = 1

                        '【処理詳細仕様 3-2】平均値算出データ取得
                        Dim rowList = DAOOther.SelectKobetsuKekkaHyoRowListOfAverage(db, cboChosaNen.Text)

                        '進捗を進める
                        progressDialog.AddValue = 1

                        '【処理詳細仕様 3-3】平均値算出データ登録
                        DAOOther.InsertAverageTable(db, rowList)

                        '進捗を進める
                        progressDialog.AddValue = 1

                        db.CommitTrans()

                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try
                End Using
            Finally
                If Not ProgressDialog Is Nothing Then
                    '進捗ダイアログを閉じる
                    ProgressDialog.endDispose()
                    ProgressDialog = Nothing
                End If
            End Try

            '【処理詳細仕様 3-4】完了メッセージ表示
            Message.ShowMsgBox(MessageID.MSG_I_023, MsgBoxStyle.OkOnly)

            '更新日時設定
            ComUtil.UpdateDate(txtUpdateDate, cboChosaNen.Text, ComConst.欠測値.営農欠測値適用平均値データ)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try

    End Sub

#End Region

#Region "【処理詳細仕様 4】「出力」ボタンクリック"

    ''' <summary>
    ''' 出力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click

        If String.IsNullOrWhiteSpace(cboChosaNen.Text) Then
            Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
            Return
        End If

        '帳票出力データを検索する
        Dim rowList = ComUtil.GetAverageDataRowList(cboChosaNen.Text)

        If rowList Is Nothing OrElse Not rowList.Any() Then
            Message.ShowMsgBox(MessageID.MSG_E_023, MsgBoxStyle.OkOnly)
            Return
        End If

        Try
            '【処理詳細仕様 ①】ファイル保存ダイアログを表示する
            Dim fileName As String = ComConst.欠測値.適用データ一覧表出力用ファイル.Report.reportName & "_" & cboChosaNen.Text & ".xlsx"

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            '【処理詳細仕様 ③】帳票を出力する
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New BRA5710R(filePath, cboChosaNen.Text)
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
        End Try

    End Sub

#End Region

End Class
