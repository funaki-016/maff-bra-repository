Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO

''' <summary>
''' 個別結果票CSV一括登録
''' </summary>
''' <remarks></remarks>
Public Class BRA4110F
    '個別結果票CSVフォルダ名
    Private Const DIRECTORY_NAME As String = "個別結果表CSV"

    'CSV区切文字
    Private Const CSV_DELIMITER As String = ","
    'コードページ_Shift_JIS
    Private Const CODEPAGE_SHIFT_JIS As String = "Shift_JIS"

#Region "【処理詳細仕様 1】初期表示"

    ''' <summary>
    ''' フォームのロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA4110F_Load(sender As Object, e As EventArgs) Handles Me.Load

    End Sub

#End Region



#Region "【処理詳細仕様 3】参照ボタン押下"
    ''' <summary>
    ''' 参照ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSubstitute_Click(sender As Object, e As EventArgs) Handles btnSubstitute.Click
        Try
            'フォルダパス取得
            Dim folderPath As String = ComUtil.GetFolderPath(Me, IniFileInfo.ExcelInPath)

            If folderPath.Equals(String.Empty) Then
                Exit Sub
            End If

            txtParameters.Text = folderPath
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region

#Region "【処理詳細仕様 4】取込ボタン押下"
    Private Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click

        Dim chosanen As String = String.Empty
        Dim folderPath As String = String.Empty
        Dim errMassage As String = String.Empty
        Dim dbList As String() = {"BRAH", "BRAN", "BRAS"}

        Try
            chosanen = txtChosanen.Text
            folderPath = txtParameters.Text

            '入力値のエラーチェックを行う
            errMassage = txtCheckError(chosanen, folderPath)

            If Not errMassage = String.Empty Then
                Message.ShowMsgBox(errMassage, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '処理確認
            If Message.ShowMsgBox(MessageID.MSG_Q_046, {chosanen}, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If

            ' 対象ファイルを検索する
            Dim fileNameList As String() = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)

            'フォルダ内エラーチェック
            errMassage = CheckError(fileNameList, folderPath, chosanen)

            If Not errMassage = String.Empty Then
                Message.ShowMsgBox(errMassage, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim progressDialog = New ProgressDialog

            Try
                'CSVデータ取込
                Dim csvData As Dictionary(Of String, List(Of String()))
                csvData = getCSVData(fileNameList, folderPath, chosanen, errMassage)

                'データ取込時にエラーが発生した場合、処理を終了する。
                If errMassage <> String.Empty Then
                    Message.ShowMsgBox(errMassage, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                'ダイアログボックスの表示
                'ダイアログの最大値は、CSVファイルの全体の件数+BRAN,BRASのテーブル数
                progressDialog.Maximum = csvData.Values(0).Count * csvData.Keys.Count + csvData.Keys.Count * 2
                progressDialog.Show(Me)

                Using db As New DBAccess(My.Settings.BRAHConnectionString)
                    Try
                        db.BeginTrans()

                        For Each dbName As String In dbList

                            '個別結果票データ削除
                            DAOKobetsuKekkahyo.DeleteCSVTable(db, CommonInfo.Chosakubun, chosanen, dbName)

                            '個別結果票データ登録
                            If dbName = dbList(0) Then
                                DAOKobetsuKekkahyo.InsertCSVTable(db, CommonInfo.Chosakubun, chosanen, csvData, dbName, progressDialog)
                            Else
                                DAOKobetsuKekkahyo.InsertCSVTableCopy(db, CommonInfo.Chosakubun, chosanen, dbName, progressDialog)
                            End If
                        Next

                        db.CommitTrans()

                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try
                End Using
            Finally
                If Not progressDialog Is Nothing Then
                    '進捗ダイアログを閉じる
                    progressDialog.endDispose()
                    progressDialog = Nothing
                End If
            End Try

            '完了メッセージ表示
            Message.ShowMsgBox(MessageID.MSG_I_044, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try

    End Sub
#End Region

    ''' <summary>
    ''' CSVデータ取込
    ''' </summary>
    ''' <param name="fileNameList"></param>
    ''' <param name="folderPath"></param>
    ''' <param name="chosanen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getCSVData(fileNameList As String(), folderPath As String, chosanen As String, ByRef MassageId As String) As Dictionary(Of String, List(Of String()))
        Dim ret As New Dictionary(Of String, List(Of String()))
        Dim sjisEnc As Encoding = Encoding.GetEncoding(CODEPAGE_SHIFT_JIS)

        For Each fileName As String In fileNameList
            '前方一致でファイルを検索
            Dim fileList As String() = Directory.GetFileSystemEntries(folderPath, fileName & "_" & chosanen & "*.csv")

            Dim filePath As String = fileList(0)


            If System.IO.File.Exists(filePath) Then
                Dim lstLine As New List(Of String())

                Using parser As New TextFieldParser(filePath, sjisEnc)

                    parser.TextFieldType = FieldType.Delimited
                    parser.SetDelimiters(CSV_DELIMITER)
                    parser.HasFieldsEnclosedInQuotes = True

                    While Not parser.EndOfData
                        Dim arr As String() = parser.ReadFields()
                        lstLine.Add(arr)
                    End While
                End Using

                If lstLine(0)(0) <> chosanen Then
                    MassageId = MessageID.MSG_E_094
                    Exit For
                End If

                ret.Add(fileName, lstLine)
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 画面入力値のエラーチェック
    ''' </summary>
    ''' <param name="chosanen"></param>
    ''' <param name="folderPath"></param>
    ''' <returns></returns>
    Private Function txtCheckError(chosanen As String, folderPath As String) As String
        Dim ret As String = String.Empty

        '調査年が入力されていない場合、エラー表示
        If String.IsNullOrEmpty(chosanen) Then
            ret = MessageID.MSG_E_002
            Return ret
        End If

        '[個別結果票CSV]フォルダが選択されていない場合、エラー表示
        If Not Directory.Exists(folderPath) Then
            ret = MessageID.MSG_E_091
            Return ret
        End If

        Return ret
    End Function

    ''' <summary>
    ''' ファイル名のエラーチェック
    ''' </summary>
    ''' <param name="fileNameList"></param>
    ''' <param name="folderPath"></param>
    ''' <param name="chosanen"></param>
    ''' <returns></returns>
    Private Function CheckError(fileNameList As String(), folderPath As String, chosanen As String) As String
        Dim ret As String = String.Empty

        For Each fileName As String In fileNameList
            Dim fileList As String()
            fileList = Directory.GetFileSystemEntries(folderPath, fileName & "_*" & ".csv")

            If fileList.Length = 0 Then
                ret = MessageID.MSG_E_092
                Return ret
            End If

            fileList = Directory.GetFileSystemEntries(folderPath, fileName & "_" & chosanen & "*.csv")

            If fileList.Length = 0 Then
                'ファイル名が存在しない場合エラー表示
                ret = MessageID.MSG_E_093
                Return ret
            End If

        Next

        Return ret
    End Function

End Class
