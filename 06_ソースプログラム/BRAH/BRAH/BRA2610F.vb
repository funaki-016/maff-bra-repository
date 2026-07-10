Imports Microsoft.Office.Interop

''' <summary>
''' 制度受取金・積立金等項目画面
''' </summary>
''' <remarks></remarks>
Public Class BRA2610F
    Private Sub BRA2610F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '更新日時設定
            Me.UpdateDate(txtUpdateDate)

            '調査年コンボボックス設定
            ComUtil.Seidouketorikin.SetSeidouketorinenComboBox(cboChosanen)


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
    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click

        Dim xlForm As BRA2620X = Nothing

        '調査年
        ComUtil.Seidouketorikin.Chosanen = cboChosanen.SelectedValue.ToString

        Try

            xlForm = New BRA2620X(Me)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally

        End Try
    End Sub

    ''' <summary>
    ''' 取込ボタン押下
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Try

            ComUtil.Seidouketorikin.Chosanen = cboChosanen.SelectedValue.ToString

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of OpenFileDialog)(Me, IniFileInfo.ExcelInPath)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Dim ret As Boolean
            '制度受取金・積立金等項目取込クラス生成
            Using ExcelImport = New ImportSeidouketorikin()
                '処理実行
                ret = ExcelImport.Execute(filePath, Me)
            End Using

            If ret Then
                '更新日時設定
                Me.UpdateDate(txtUpdateDate)

                '完了メッセージ
                Message.ShowMsgBox(MessageID.MSG_I_006, MsgBoxStyle.OkOnly)
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub


    ''' <summary>
    ''' 出力ボタン押下
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click

        Try

            '調査年
            ComUtil.Seidouketorikin.Chosanen = cboChosanen.SelectedValue.ToString

            'ファイル名取得
            Dim fileName As String = "制度受取金・積立金等項目" & "_" & "(" & CommonInfo.ChosakubunName & ")" & ".xlsx"

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            ' 制度受取金・積立金等項目出力
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New BRA2610R(filePath)
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
    ''' 更新日時設定
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub UpdateDate(txt As TextBox)
        Dim dtn As String
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dtn = DAOOther.GetSeidouketorikinDate(db)
        End Using

        If Not dtn.Equals(String.Empty) Then
            txt.Text = DateTime.Parse(dtn).ToString(ComConst.DATETIME_FORMAT)
        Else
            txt.Text = String.Empty
        End If
    End Sub

    Public Shared Function chosanen() As String
        Dim ret As String
        Dim chosanenselct As String
        chosanenselct = BRA2610F.cboChosanen.SelectedValue.ToString
        Return ret
    End Function


    ''' <summary>
    ''' 制度受取金・積立金等項目取込クラス
    ''' </summary>
    ''' <remarks></remarks>
    Private Class ImportSeidouketorikin
        Inherits ExcelProcess

        ''' <summary>進捗ダイアログ</summary>
        Private ProgressDialog As New ProgressDialog()

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            MyBase.New(True)
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
            Dim lstDc As List(Of Dictionary(Of String, String)) = ComUtil.Seidouketorikin.GetSheetData(xlSheets, CType(Me, ComObjectProcess))


            'プログレスバーの最大値
            ProgressDialog.Maximum = lstDc.Count

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try

                    db.BeginTrans()

                    '制度受取金・積立金等項目削除
                    DAOOther.DeleteSeidouketorikin(db)

                    'Dim Koubannum As Integer
                    'Dim Kouban As String
                    'Koubannum = 9010000

                    For Each dc As Dictionary(Of String, String) In lstDc
                        '制度受取金・積立金等項目追加
                        DAOOther.InsertSeidouketorikin(db, dc)
                        'Koubannum = Koubannum + 1
                        'Kouban = "Q0" & Koubannum
                        ''該当調査票テーブル区分名反映
                        'DAOOther.InsertOtherSeidouketorikin(db, dc, Kouban)

                        '進捗を進める
                        ProgressDialog.AddValue = 1
                    Next

                    '制度受取金・積立金項番削除対象確認
                    Dim adress(151) As String
                    Dim Seidokinkoubannum1 As Integer
                    Dim Seidokinkoubannum2 As Integer
                    Dim Seidokinkouban As String
                    Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)

                    Select Case CommonInfo.Kubun2
                        Case ComConst.区分２.営農類型別経営統計
                            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                                Seidokinkoubannum1 = 9010110
                                Seidokinkoubannum2 = 9010210
                            Else
                                Seidokinkoubannum1 = 1010110
                                Seidokinkoubannum2 = 1010210
                            End If
                        Case ComConst.区分２.農産物生産費
                            If CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_個別 Then
                                Seidokinkoubannum1 = 1110110
                                Seidokinkoubannum2 = 1110210
                            ElseIf CommonInfo.Chosakubun = ComConst.調査区分.原料用ばれいしょ生産費統計_個別 Or
                                           CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費 Then
                                Seidokinkoubannum1 = 1060110
                                Seidokinkoubannum2 = 1060210
                            ElseIf CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_組織法人 Then
                                Seidokinkoubannum1 = 1130110
                                Seidokinkoubannum2 = 1130210
                            Else
                                Seidokinkoubannum1 = 1070110
                                Seidokinkoubannum2 = 1070210
                            End If
                    End Select
                    Dim chosanenselect As String
                    If ComUtil.Seidouketorikin.Chosanen = "1" Then
                        chosanenselect = "2022"
                    ElseIf ComUtil.Seidouketorikin.Chosanen = "2" Then
                        chosanenselect = "2023"
                    ElseIf ComUtil.Seidouketorikin.Chosanen = "3" Then
                        chosanenselect = "2024"
                    ElseIf ComUtil.Seidouketorikin.Chosanen = "4" Then
                        chosanenselect = "2025"
                    ElseIf ComUtil.Seidouketorikin.Chosanen = "5" Then
                        chosanenselect = "2026"
                    ElseIf ComUtil.Seidouketorikin.Chosanen = "6" Then
                        chosanenselect = "2027"
                    ElseIf ComUtil.Seidouketorikin.Chosanen = "7" Then
                        chosanenselect = "2028"
                    ElseIf ComUtil.Seidouketorikin.Chosanen = "8" Then
                        chosanenselect = "2029"
                    ElseIf ComUtil.Seidouketorikin.Chosanen = "9" Then
                        chosanenselect = "2030"
                    ElseIf ComUtil.Seidouketorikin.Chosanen = "10" Then
                        chosanenselect = "2031"
                    End If

                    For i = 1 To 152
                        If i <= 76 Then
                            Seidokinkoubannum1 = Seidokinkoubannum1 + 1
                            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                                Seidokinkouban = "Q1" & Seidokinkoubannum1
                            Else
                                Seidokinkouban = "Q0" & Seidokinkoubannum1
                            End If
                            Dim Seidokinkoubancheck = DAOOther.Seidokincheck(db, CommonInfo.Chosakubun, chosanenselect, Seidokinkouban)
                            If Seidokinkoubancheck.Rows.Count = 0 Then
                                adress(i - 1) = Seidokinkouban
                            End If
                        Else
                            Seidokinkoubannum2 = Seidokinkoubannum2 + 1
                            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                                Seidokinkouban = "Q1" & Seidokinkoubannum2
                            Else
                                Seidokinkouban = "Q0" & Seidokinkoubannum2
                            End If
                            Dim Seidokinkoubancheck = DAOOther.Seidokincheck2(db, CommonInfo.Chosakubun, chosanenselect, Seidokinkouban)
                            If Seidokinkoubancheck.Rows.Count = 0 Then
                                adress(i - 1) = Seidokinkouban
                            End If

                        End If
                    Next

                    '制度受取金・積立金項番削除
                    DAOOther.DeleteSeidouketorikinNum(db, adress)

                    db.CommitTrans()

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
                      {"B1", CommonInfo.ChosakubunName} _
                    , {"B2", "制度受取金・積立金等項目"} _
                    , {"B4", "項番"} _
                    , {"C4", "出力項目名"} _
                    , {"D4", "制度受取金等項番"} _
                    , {"E4", "制度積立金等項番"}
                }

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シート名チェック
                Dim bln As Boolean = False
                For Each wsh As Excel.Worksheet In xlSheets
                    Try
                        If wsh.Name.Equals(ComConst.制度受取金積立金等項目.出力用ファイル名称.SheetName) Then
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
                xlSheet = DirectCast(xlSheets.Item(ComConst.制度受取金積立金等項目.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                '調査区分、タイトル、表頭チェック
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
