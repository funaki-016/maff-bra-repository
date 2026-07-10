Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop.Forms
Imports System.Runtime.InteropServices
Imports System.Reflection

''' <summary>
''' 制度受取金・積立金等項目入力・修正（EXCEL）
''' </summary>
''' <remarks></remarks>
Public Class BRA2620X
    Inherits ExcelInputBaseClass

    ''' <summary>保存ボタン</summary>
    Private WithEvents btnSaveClose As CommandButton
    ''' <summary>戻るボタン</summary>
    Private WithEvents btnNoSaveClose As CommandButton

    ''' <summary>ヘッダータイトル文字列</summary>
    Private Const HEADER_TITLE As String = "制度受取金・積立金等項目入力・修正画面"

    ''' <summary>Excelユーザーフォームハンドル</summary>
    Private _formHwnd As Win32WindowWrapper

    ''' <summary>更新日時設定デリゲート</summary>
    Private Delegate Sub UpdateDateDelegate(txt As System.Windows.Forms.TextBox)
    ''' <summary>制度受取金・積立金等項目入力・修正画面</summary>
    Private _Form As BRA2610F

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef frm As BRA2610F)
        MyBase.New(frm, System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), "制度受取金・積立金等項目.xlsm"), True)
        Try
            _Form = frm

            'ヘッダーを設定する
            Me.SetHeader()

            'データを設定する
            Me.SetData()

            'Excel画面を表示する
            Me.ShowExcel()

            '処理待ち画面を閉じる
            Me.CloseWaitForm()
        Catch ex As Exception
            'Excel画面を閉じる
            Me.CloseExcel()
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Excel画面を表示する
    ''' </summary>
    ''' <param name="pFilePath"></param>
    ''' <remarks></remarks>
    Private Sub ShowExcel()
        Dim uf As UserForm = Nothing
        Try
            'ボタンの設定
            uf = ComUtilStrictOff.GetExcelForm(xlBook)
            btnSaveClose = ComUtilStrictOff.GetExcelBtnSaveClose(uf)
            btnNoSaveClose = ComUtilStrictOff.GetExcelBtnNoSaveClose(uf)

            'ユーザーフォームを表示する
            ComUtilStrictOff.ShowExcelForm(uf)

            'Excelユーザーフォームハンドル設定
            _formHwnd = New Win32WindowWrapper(ComUtilStrictOff.GetFormHwnd(uf))

            '画面更新有効
            xlApp.ScreenUpdating = True

        Catch ex As Exception
            Throw
        Finally
            ReleaseComObject(uf)
        End Try
    End Sub

    ''' <summary>
    ''' Excel画面を閉じる
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub CloseExcel()

        Try
            'ボタンの解放
            ReleaseComObject(btnSaveClose)
            ReleaseComObject(btnNoSaveClose)

            'Excel画面を閉じる
            MyBase.CloseExcel()

            '更新日時設定
            Me.UpdateDate()

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 保存ボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnSaveClose_Click() Handles btnSaveClose.Click
        Dim ProgressDialog As New ProgressDialog()

        Try
            '確認メッセージ
            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_001, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                '進捗ダイアログを表示する
                ProgressDialog.Maximum = 3
                ProgressDialog.Show(_formHwnd)

                'Excel前処理
                Me.BeforeExcel()

                'シートデータ取得
                Dim lstDc As List(Of Dictionary(Of String, String)) = ComUtil.Seidouketorikin.GetSheetData(xlSheets, CType(Me, ComObjectProcess))

                '進捗を進める
                ProgressDialog.AddValue = 1


                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    Try

                        db.BeginTrans()

                        '制度受取金・積立金等項目削除
                        DAOOther.DeleteSeidouketorikin(db)

                        For Each dc As Dictionary(Of String, String) In lstDc
                            '制度受取金・積立金等項目追加                            
                            DAOOther.InsertSeidouketorikin(db, dc)


                            'Koubannum = Koubannum + 1
                            'Kouban = "Q0" & Koubannum
                            ''該当調査票テーブル区分名反映
                            'DAOOther.InsertOtherSeidouketorikin(db, dc, Kouban)
                        Next


                        'Dim Koubannum As Integer
                        'Dim Kouban As String
                        'Koubannum = 9010000


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

                        '>>>2022/02/02
                        Dim dbN As New DBAccess(My.Settings.BRANConnectionString)
                        Try
                            DAOOther.DeleteSeidouketorikinNum(dbN, adress)
                            dbN.CommitTrans()
                        Catch ex As Exception
                            dbN.RollBackTrans()
                            Throw ex
                        End Try

                        Dim dbS As New DBAccess(My.Settings.BRASConnectionString)
                        Try
                            DAOOther.DeleteSeidouketorikinNum(dbS, adress)
                            dbS.CommitTrans()
                        Catch ex As Exception
                            dbS.RollBackTrans()
                            Throw ex
                        End Try
                        '<<<2022/02/02

                    Catch ex As Exception
                            db.RollBackTrans()
                        Throw ex
                    End Try
                End Using


                '進捗を進める
                ProgressDialog.AddValue = 1

                    '進捗ダイアログを閉じる
                    ProgressDialog.endDispose()

                    '完了メッセージ
                    Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_001, MsgBoxStyle.OkOnly)
                End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            'Excel後処理
            Me.AfterExcel()

            If Not ProgressDialog Is Nothing Then
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                ProgressDialog = Nothing
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 戻るボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnNoSaveClose_Click() Handles btnNoSaveClose.Click
        Try

            'Excel画面を閉じる
            Me.CloseExcel()

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' ヘッダーを設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetHeader()
        Dim xlSheet As Excel.Worksheet = Nothing

        Try
            xlSheet = DirectCast(xlSheets.Item(1), Excel.Worksheet)

            xlSheet.Unprotect()

            SetExcelValue(xlSheet, 1, 2, CommonInfo.ChosakubunName)
            SetExcelValue(xlSheet, 2, 2, HEADER_TITLE)

            xlSheet.Protect()
        Catch ex As Exception
            Throw
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub

    ''' <summary>
    ''' データを設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetData()
        Try
            xlApp.Interactive = False

            'Excel前処理
            Me.BeforeExcel()

            Dim dt As DataTable

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '制度受取金・積立金等項目入力・修正
                dt = DAOOther.GetSeidouketorikin(db)
            End Using

            'シートデータ設定
            ComUtil.Seidouketorikin.SetSheetData(dt, xlSheets, CType(Me, ExcelProcess))
        Catch ex As Exception
            Throw
        Finally
            'Excel後処理
            Me.AfterExcel()

            xlApp.Interactive = True
        End Try
    End Sub

    ''' <summary>
    ''' 更新日時設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateDate()
        'デリゲートの実行
        _Form.Invoke(New UpdateDateDelegate(AddressOf _Form.UpdateDate), _Form.txtUpdateDate)
    End Sub
End Class
