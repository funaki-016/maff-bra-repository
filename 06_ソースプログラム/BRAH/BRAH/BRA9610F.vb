Imports Microsoft.Office.Interop


'//*************************************************************************************************
'//  修正履歴
'// ------------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前           |                修  正  内  容
'// -----------+------------+------------------------+----------------------------------------------
'//  REV_001   | 2024.05.31 |Daiko                   | 要件No.1 新規作成
'//            |            |                        |
'//*************************************************************************************************

''' <summary>
''' 農業地域類型マスタ管理画面
''' </summary>
''' <remarks></remarks>
Public Class BRA9610F

    Private Sub BRA9610F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '農業地域類型マスタコンボボックス設定
            ComUtil.NogyoChikiMstMainte.SetMainteMastComboBox(nogyoChikiMastaYear)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub NogyoChikiMastaYear_SelectedIndexChanged(sender As Object, e As EventArgs) Handles nogyoChikiMastaYear.SelectedIndexChanged
        Try
            '農業地域類型マスタ選択チェック
            If Not nogyoChikiMastaYear.SelectedIndex = 0 Then
                '更新日時設定
                Me.UpdateDate(txtUpdateDate, nogyoChikiMastaYear)
            Else
                txtUpdateDate.Text = String.Empty
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 更新日時設定
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <param name="nogyoChikiMastaYear"></param>
    Public Sub UpdateDate(txt As TextBox, cbo As ComboBox)
        Dim dtn As String

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '農業地域類型マスタ調査年 更新日時取得
            dtn = DAOOther.GetNogyoChikiMstYearUpdateDate(db, cbo.SelectedValue.ToString)
        End Using

        If Not dtn.Equals(String.Empty) Then
            '更新日時設定
            txt.Text = DateTime.Parse(dtn).ToString(ComConst.DATETIME_FORMAT)
        Else
            txt.Text = String.Empty
        End If
    End Sub

    Private Sub BtnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Try
            '多重起動防止のためボタンを非活性にする
            btnImport.Enabled = False
            btnOutPut.Enabled = False
            btnReturn.Enabled = False

            '農業地域類型マスタ選択チェック
            If nogyoChikiMastaYear.SelectedIndex = 0 Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_161, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of OpenFileDialog)(Me, IniFileInfo.ExcelInPath)
            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Dim ret As Boolean
            '農業地域類型マスタ一覧取込クラス生成
            Using ExcelImport = New ImportNogyoChikiMstYearList(nogyoChikiMastaYear.SelectedValue.ToString)
                ret = ExcelImport.Execute(filePath, Me)
            End Using

            If ret Then
                '更新日時設定
                Me.UpdateDate(txtUpdateDate, nogyoChikiMastaYear)
                '完了メッセージ
                Message.ShowMsgBox(MessageID.MSG_I_006, MsgBoxStyle.OkOnly)
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            'エラーメッセージ
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            btnImport.Enabled = True
            btnOutPut.Enabled = True
            btnReturn.Enabled = True
        End Try
    End Sub

    Private Sub BtnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click
        Dim dt As DataTable

        Try
            '農業地域類型マスタ選択チェック
            If nogyoChikiMastaYear.SelectedIndex = 0 Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_161, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Try
                '出力データ有無確認
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    dt = DAOOther.GetNogyoChikiMstYearList(db, nogyoChikiMastaYear.SelectedValue.ToString)
                End Using
                If dt.Rows.Count < 1 Then
                    'エラーメッセージ
                    Message.ShowMsgBox(MessageID.MSG_E_023, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
            Catch ex As Exception
                Throw
            End Try

            '出力ファイル名設定
            Dim fileName As String = ComConst.農業地域類型マスタ管理.ファイル情報.reportName & "_" _
                & nogyoChikiMastaYear.SelectedValue.ToString & ".xlsx"

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '農業地域類型マスタ一覧出力
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn

                Using ExcelOutput = New BRA9610R(filePath, nogyoChikiMastaYear.SelectedValue.ToString)
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
    ''' 農業地域類型マスタ調査年一覧取込クラス
    ''' </summary>
    ''' <remarks></remarks>
    Private Class ImportNogyoChikiMstYearList
        Inherits ExcelProcess

        ''' <summary>進捗ダイアログ</summary>
        Private ProgressDialog As New ProgressDialog()

        ''' <summary>農業地域類型マスタ調査年</summary>
        Private _year As String

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="year"></param>
        Public Sub New(year As String)
            MyBase.New(True)
            '農業地域類型マスタ調査年設定
            _year = year
        End Sub

        ''' <summary>
        ''' 処理実行
        ''' </summary>
        ''' <param name="filePath"></param>
        ''' <param name="myParent"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Execute(filePath As String, myParent As Form) As Boolean
            Dim ret As Boolean = False

            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Sheets = Nothing

            'Excelアプリ無効
            Me.DisableExcelApp()

            Try
                '進捗ダイアログを表示する
                ProgressDialog.Show(myParent)

                Try
                    'Workbookを開く
                    xlBook = xlBooks.Open(filePath)
                    Try
                        xlSheets = xlBook.Worksheets

                        '取込処理
                        ret = Me.Import(xlSheets, myParent)

                    Catch ex As Exception
                        Throw ex
                    Finally
                        'Sheetsの解放
                        ReleaseComObject(xlSheets)
                    End Try
                Catch ex As Exception
                    Throw ex
                Finally
                    'Workbookを閉じる
                    If xlBook IsNot Nothing Then
                        xlBook.Saved = True
                        xlBook.Close()
                    End If
                    'Workbookの解放
                    ReleaseComObject(xlBook)
                End Try

            Finally
                If Not ProgressDialog Is Nothing Then
                    '進捗ダイアログを閉じる
                    ProgressDialog.endDispose()
                    ProgressDialog = Nothing
                End If

            End Try

            'Excelアプリ有効
            Me.EnableExcelApp()

            Return ret
        End Function

        ''' <summary>
        ''' 取込処理
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="myParent"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Import(xlSheets As Excel.Sheets, myParent As Form) As Boolean
            Dim ret As Boolean = False


            'ファイル形式チェック
            If Not Me.CheckFileFormat(xlSheets) Then
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_019, MsgBoxStyle.OkOnly)
                Return ret
            End If

            'シートデータ取得
            Dim lstDc As List(Of Dictionary(Of String, String)) = ComUtil.NogyoChikiMstMainte.GetSheetData(xlSheets, CType(Me, ComObjectProcess))

            'プログレスバーの最大値 ※登録処理はbulkinsertのためプログレスバーはエラーチェックとデータ取込の処理で表現
            ProgressDialog.Maximum = 2

            'エラーチェック
            Dim details As New List(Of String)
            If Not ComUtil.NogyoChikiMstMainte.CheckError(lstDc, details) Then
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                'メッセージフォーム表示
                Message.ShowMsgForm(myParent, MessageID.MSG_E_024, {String.Join(vbCrLf, details)})
                Return ret
            End If

            ProgressDialog.AddValue = 1

            'データ取込
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try
                    '農業地域類型マスタ調査年データ存在チェック
                    Dim count = DAOOther.GetNogyoChikiMstYearList(db, _year).Rows.Count

                    Dim delFlg = False
                    If count > 0 Then
                        '確認メッセージ
                        If ProgressDialog.ShowMsgBox(MessageID.MSG_Q_069, {_year}, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                            delFlg = True
                        Else
                            '進捗ダイアログを閉じる
                            ProgressDialog.endDispose()

                            Return False
                        End If
                    End If

                    db.BeginTrans()

                    If delFlg Then
                        '農業地域類型マスタ調査年データ削除
                        DAOOther.DeleteNogyoChikiMstYear(db, _year)
                    End If

                    '農業地域類型マスタ調査年データ登録
                    DAOOther.InsertNogyoChikiMstYearList(db, ComConst.農業地域類型マスタ管理.テーブル名称, _year, lstDc)

                    db.CommitTrans()

                    ProgressDialog.AddValue = 1

                    ret = True
                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using

            Return ret
        End Function

        ''' <summary>
        ''' ファイル形式チェック
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function CheckFileFormat(xlSheets As Excel.Sheets) As Boolean
            Dim ret As Boolean = True

            'ファイル形式
            Dim FileFormat As New Dictionary(Of String, String) From {
            {"A5", ComConst.農業地域類型マスタ管理.ファイル情報.reportName},
            {"G5", _year},
            {"A6", "都道府県"},
            {"G6", "第２次分類"}
            }

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シート名チェック
                Dim bln As Boolean = False
                For Each wsh As Excel.Worksheet In xlSheets
                    Try
                        If wsh.Name.Equals(ComConst.農業地域類型マスタ管理.ファイル情報.SheetName) Then
                            bln = True
                            Exit For
                        End If
                    Finally
                        ReleaseComObject(wsh)
                    End Try
                Next
                If Not bln Then
                    ret = False
                    Return ret
                End If

                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.農業地域類型マスタ管理.ファイル情報.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                'タイトル、調査年、表頭チェック
                For Each kv As KeyValuePair(Of String, String) In FileFormat
                    Dim rng As Excel.Range = Nothing
                    Try
                        rng = xlSheet.Range(kv.Key)
                        Dim val As String = If(rng.Value Is Nothing, String.Empty, rng.Value.ToString)
                        If Not val = kv.Value Then
                            ret = False
                            Return ret
                        End If
                    Finally
                        ReleaseComObject(rng)
                    End Try
                Next

                '一行目入力チェック
                Dim rng1 As Excel.Range = Nothing
                Try
                    rng1 = xlSheet.Range("A7")
                    If rng1.Value Is Nothing Then
                        ret = False
                        Return ret
                    End If
                Finally
                    ReleaseComObject(rng1)
                End Try

                'シート保護設定
                If protect Then
                    xlSheet.Protect()
                End If

            Finally
                ReleaseComObject(xlSheet)
            End Try

            Return ret
        End Function
    End Class

End Class
