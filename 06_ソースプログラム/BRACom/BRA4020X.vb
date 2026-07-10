Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop.Forms
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO
Imports System.Text
Imports System.IO

''' <summary>
''' 個別結果表項目指定修正（EXCEL）画面
''' </summary>
''' <remarks></remarks>
''' 
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2021.12.07 |日本コンピュータシステム | 要件No1-③対応
'//  REV_002   | 2022.10.11 |daiko               | 要件No1 バージョン区分追加
'//            |            |                    |
'//*************************************************************************************************

Public Class BRA4020X
    Inherits ExcelInputBaseClass

    ''' <summary>保存（再計算あり）ボタン</summary>
    Private WithEvents btnSaveCalculate As CommandButton
    ''' <summary>保存（再計算なし）ボタン</summary>
    Private WithEvents btnSaveNoCalculate As CommandButton
    ''' <summary>戻るボタン</summary>
    Private WithEvents btnNoSaveClose As CommandButton
    ''' <summary>修正前データ表示ボタン</summary>
    Private WithEvents btnBeforeDataShow As CommandButton
    ''' <summary>ファイル入力ボタン</summary>
    Private WithEvents btnFileInput As CommandButton
    ''' <summary>ファイル出力ボタン</summary>
    Private WithEvents btnFileOutput As CommandButton

    ''' <summary>ヘッダータイトル文字列</summary>
    Private Const HEADER_TITLE As String = "個別結果表項目指定修正画面"

    ''' <summary>ファイルタイトル文字列</summary>
    Private Const FILE_TITLE As String = "個別結果表項目指定修正"

    ''' <summary>Excelユーザーフォームハンドル</summary>
    Private _formHwnd As Win32WindowWrapper

    ''' <summary>個別結果表項目指定修正画面</summary>
    Private _Form As BRA4010F
    ''' <summary>調査年</summary>
    Private _chosaNen As String
    ''' <summary>センサスNo</summary>
    Private _censusNo As List(Of String)
    ''' <summary>欠測値補完</summary>
    Private _kessokuchiHokan As String

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef frm As BRA4010F, chosaNen As String, censusNo As List(Of String), kessokuchiHokan As String)
        MyBase.New(frm, System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.個別結果表項目指定修正.入力用ファイル名称), True)
        Try
            _Form = frm
            _chosaNen = chosaNen
            _censusNo = censusNo
            _kessokuchiHokan = kessokuchiHokan

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
            btnSaveCalculate = ComUtilStrictOff.GetExcelBtnSaveCalculate(uf)
            btnSaveNoCalculate = ComUtilStrictOff.GetExcelBtnSaveNoCalculate(uf)
            btnNoSaveClose = ComUtilStrictOff.GetExcelBtnNoSaveClose(uf)
            btnBeforeDataShow = ComUtilStrictOff.GetExcelBtnBeforeDataShow(uf)
            btnFileInput = ComUtilStrictOff.GetExcelBtnFileInput(uf)
            btnFileOutput = ComUtilStrictOff.GetExcelBtnFileOutput(uf)

            'ユーザーフォームを表示する
            ComUtilStrictOff.ShowExcelForm(uf)

            If Not CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Honsyo Then
                'ボタン非活性
                btnSaveNoCalculate.Enabled = False
            End If

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
            ReleaseComObject(btnSaveCalculate)
            ReleaseComObject(btnSaveNoCalculate)
            ReleaseComObject(btnNoSaveClose)
            ReleaseComObject(btnBeforeDataShow)
            ReleaseComObject(btnFileInput)
            ReleaseComObject(btnFileOutput)

            'Excel画面を閉じる
            MyBase.CloseExcel()

            ''更新日時設定
            'Me.UpdateDate()

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 保存（再計算あり）ボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnSaveCalculate_Click() Handles btnSaveCalculate.Click
        Dim ProgressDialog As New ProgressDialog()
        Try
            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_025, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                '進捗ダイアログを表示する
                ProgressDialog.Show(_formHwnd)

                'Excel前処理
                Me.BeforeExcel()

                Dim createRonriCnt As Integer
                Dim dtKobetsuItemMst As DataTable
                Dim dtSheet As DataTable
                Dim einoresult As Object() = Nothing
                Dim errorcensusNo As String

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                    '個別結果表項目マスタ取得
                    ' REV_002↓
                    'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun)
                    dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
                    ' REV_002↑

                    '修正データ取得
                    dtSheet = GetSheetData()

                    'エラーチェック
                    Dim details As New List(Of String)
                    If Not Me.CheckError(dtSheet, dtKobetsuItemMst, details) Then
                        '進捗ダイアログを閉じる
                        ProgressDialog.endDispose()
                        'エラーメッセージ
                        Message.ShowMsgForm(_formHwnd, MessageID.MSG_E_024, {String.Join(vbCrLf, details)})
                        Exit Sub
                    End If

                    '保存前確認
                    If (From dr In dtSheet Where String.IsNullOrEmpty(dr(3).ToString)).Any Then
                        If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_030, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                            '進捗ダイアログを閉じる
                            ProgressDialog.endDispose()
                            Exit Sub
                        End If
                    End If

                    '修正対象センサス番号の個別結果表作成論理(営農個人)の総レコード件数を取得する
                    If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                        einoresult = Me.GetCreateRonriRecCntEinoKojin((From dr In dtSheet Select CStr(dr("センサス番号"))).Distinct().ToList, dtSheet)
                        createRonriCnt = CType(einoresult(0), Integer)

                    Else
                        ' REV_002
                        'createRonriCnt = DAOOther.GetKobetsuKekkahyoSakuseiRonri(db, True).Rows.Count * (From dr In dtSheet Select CStr(dr("センサス番号"))).Distinct().Count↓
                        createRonriCnt = DAOOther.GetKobetsuKekkahyoSakuseiRonri(db, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun), True).Rows.Count * (From dr In dtSheet Select CStr(dr("センサス番号"))).Distinct().Count
                        ' REV_002↑
                    End If
                    '個別結果表作成論理のデータが１件も存在しない場合
                    If createRonriCnt = 0 Then
                        '進捗ダイアログを閉じる
                        ProgressDialog.endDispose()
                        'エラーメッセージ
                        Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_017, MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If
                    '貸借対照表区分がNULLの場合
                    If createRonriCnt = -1 Then
                        '進捗ダイアログを閉じる
                        ProgressDialog.endDispose()
                        errorcensusNo = CType(einoresult(1), String)
                        'エラーメッセージ
                        Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_061, {errorcensusNo}, MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If

                    '進捗最大値を設定
                    ProgressDialog.Maximum = createRonriCnt

                    '再計算処理
                    Recalculate(ProgressDialog, db, dtKobetsuItemMst, dtSheet)

                End Using

                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()

                '完了メッセージ
                Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_016, MsgBoxStyle.OkOnly)

            End If
        Catch ex As CreateKobetsuException
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_037, {ex.ItemNo}, MsgBoxStyle.OkOnly)

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
    ''' 保存（再計算なし）ボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnSaveNoCalculate_Click() Handles btnSaveNoCalculate.Click
        Dim ProgressDialog As New ProgressDialog()
        Try
            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_026, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                '進捗ダイアログを表示する
                ProgressDialog.Maximum = 3
                ProgressDialog.Show(_formHwnd)

                'Excel前処理
                Me.BeforeExcel()

                Dim dtKobetsuItemMst As DataTable
                Dim dtSheet As DataTable

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                    '個別結果表項目マスタ取得
                    ' REV_002↓
                    'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun)
                    dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
                    ' REV_002↑

                    '修正データ取得
                    dtSheet = GetSheetData()

                    '進捗最大値を設定
                    ProgressDialog.Maximum = 3 + (From dr In dtSheet Select CStr(dr("センサス番号"))).Distinct().Count

                    'エラーチェック
                    Dim details As New List(Of String)
                    If Not Me.CheckError(dtSheet, dtKobetsuItemMst, details) Then
                        '進捗ダイアログを閉じる
                        ProgressDialog.endDispose()
                        'エラーメッセージ
                        Message.ShowMsgForm(_formHwnd, MessageID.MSG_E_024, {String.Join(vbCrLf, details)})
                        Exit Sub
                    End If

                    '進捗を進める
                    ProgressDialog.AddValue = 1

                    '保存前確認
                    If (From dr In dtSheet Where String.IsNullOrEmpty(dr(3).ToString)).Any Then
                        If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_030, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                            '進捗ダイアログを閉じる
                            ProgressDialog.endDispose()
                            Exit Sub
                        End If
                    End If

                    '保存処理
                    UpdateDate(ProgressDialog, db, dtSheet, dtKobetsuItemMst)

                    '進捗を進める
                    ProgressDialog.AddValue = 1

                End Using

                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()

                '完了メッセージ
                Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_017, MsgBoxStyle.OkOnly)

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
    ''' 修正前データ表示ボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnBeforeDataShow_Click() Handles btnBeforeDataShow.Click
        Dim ProgressDialog As New ProgressDialog()
        Dim xlSheet As Excel.Worksheet = Nothing
        Dim xlRange As Excel.Range = Nothing

        Try

            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_027, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                Dim dtSheet As DataTable

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                '進捗ダイアログを表示する
                ProgressDialog.Maximum = 5
                ProgressDialog.Show(_formHwnd)

                'Excel前処理
                Me.BeforeExcel()

                '入力データ取得
                dtSheet = GetSheetData()

                '進捗ダイアログ最大値を設定
                ProgressDialog.Maximum = dtSheet.Rows.Count

                Dim dtItem As DataTable
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    ' REV_002↓
                    'dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun)
                    dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
                    ' REV_002↑

                    'エラーチェック
                    Dim details As New List(Of String)
                    If Not Me.CheckError(dtSheet, dtItem, details, True) Then
                        '進捗ダイアログを閉じる
                        ProgressDialog.endDispose()
                        'エラーメッセージ
                        Message.ShowMsgForm(_formHwnd, MessageID.MSG_E_024, {String.Join(vbCrLf, details)})
                        Exit Sub
                    End If

                    '進捗を進める
                    ProgressDialog.AddValue = 1

                    Dim beforeData As New List(Of String)

                    '修正前データ取得

                    For Each dr As DataRow In dtSheet.Rows
                        Dim pKey As New DAOKobetsuKekkahyo.PrimaryKey(_chosaNen, dr.Item(0).ToString)
                        Dim dtKobetsu As DataTable
                        Dim dcKobetsu As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)

                        dtKobetsu = DAOKobetsuKekkahyo.GetTable(db, pKey, dr(1).ToString, _kessokuchiHokan)
                        dcKobetsu = ComUtil.KobetsuKekkahyo.GetItem(dtItem, New Dictionary(Of String, DataTable) From {{String.Empty, dtKobetsu}})

                        beforeData.Add(ComUtil.KobetsuKekkahyo.GetformattedValue(dcKobetsu(dr("項目番号").ToString)))

                        '進捗を進める
                        ProgressDialog.AddValue = 1

                    Next

                    xlSheet = DirectCast(xlSheets.Item(1), Excel.Worksheet)
                    xlRange = xlSheet.Range("E5:E5004")
                    xlRange.ClearContents()

                    'データを設定
                    For i As Integer = 0 To beforeData.Count - 1
                        SetExcelValue(xlSheet, i + 5, 5, beforeData(i))
                    Next
                End Using

                '進捗を進める
                ProgressDialog.AddValue = 1

                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()

                '完了メッセージ
                Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_022, MsgBoxStyle.OkOnly)

            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally

            ReleaseComObject(xlSheet)
            ReleaseComObject(xlRange)

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
    ''' ファイル入力ボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnFileInput_Click() Handles btnFileInput.Click
        Dim ProgressDialog As New ProgressDialog()
        Try
            'Excel前処理
            Me.BeforeExcel()

            'ファイルパス取得
            Dim filePath As String
            With xlApp.Application.FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogFilePicker)
                .Title = "ファイルを選択して下さい。"
                'ファイル複数選択を不可にする
                .AllowMultiSelect = False
                'ファイルフィルタのクリア
                .Filters.Clear()
                'ファイルフィルタの追加
                .Filters.Add("CSVファイル", "*.csv")
                '初期表示フォルダの設定
                .InitialFileName = IniFileInfo.ExcelInPath & "\"
                If .Show() = 0 Then
                    Exit Sub
                End If
                filePath = .SelectedItems(0).ToString
            End With

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '進捗ダイアログを表示する
            ProgressDialog.Maximum = 3
            ProgressDialog.Show(_formHwnd)

            '入力ファイル取得
            Dim lstArr As List(Of String()) = Me.GetInputFile(filePath)

            '進捗最大値を入力行数分追加
            ProgressDialog.Maximum = 3 + lstArr.Count

            '進捗を進める
            ProgressDialog.AddValue = 1

            'ファイル形式チェック
            If Not lstArr(0)(0).Equals(FILE_TITLE) OrElse Not lstArr(0)(1).Equals(CommonInfo.ChosakubunName) Then
                Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_019, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '進捗を進める
            ProgressDialog.AddValue = 1

            'ファイルの内容を画面にセット
            setCsvData(ProgressDialog, lstArr)

            '進捗を進める
            ProgressDialog.AddValue = 1

            '進捗ダイアログを閉じる
            ProgressDialog.endDispose()

            '完了メッセージ
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_018, MsgBoxStyle.OkOnly)

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
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        End Try
    End Sub

    ''' <summary>
    ''' ファイル出力ボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnFileOutput_Click() Handles btnFileOutput.Click
        Dim ProgressDialog As New ProgressDialog()
        Try
            Dim dtSheet As DataTable

            'Excel前処理
            Me.BeforeExcel()

            '修正データ取得(全件)
            dtSheet = GetSheetData(False)

            For i As Integer = dtSheet.Rows.Count - 1 To 0 Step -1
                If String.IsNullOrEmpty(dtSheet.Rows(i).Item(0).ToString) AndAlso _
                   String.IsNullOrEmpty(dtSheet.Rows(i).Item(1).ToString) AndAlso _
                   String.IsNullOrEmpty(dtSheet.Rows(i).Item(2).ToString) AndAlso _
                   String.IsNullOrEmpty(dtSheet.Rows(i).Item(3).ToString) Then
                    dtSheet.Rows(i).Delete()
                Else
                    Exit For
                End If
            Next

            'フォルダパス取得
            Dim fileName As String = FILE_TITLE & "_" & _chosaNen & "_" & CommonInfo.ChosakubunName & "_" & CommonInfo.Kyoku & ".csv"
            Dim filePath As String

            filePath = CStr(xlApp.GetSaveAsFilename(IniFileInfo.ExcelOutPath & "\" & fileName, "CSVファイル (*.csv), *.csv"))
            If filePath.Equals("False") Then
                Exit Sub
            End If

            '進捗ダイアログを表示する
            ProgressDialog.Maximum = 3 + dtSheet.Rows.Count
            ProgressDialog.Show(_formHwnd)

            'CSVファイル出力
            Dim ret As ComConst.CSVファイル.enmOutputReturn = PutOutputFile(ProgressDialog, filePath, dtSheet)

            Select Case ret
                Case ComConst.CSVファイル.enmOutputReturn.OK
                    '進捗を進める
                    ProgressDialog.AddValue = 1

                    '進捗ダイアログを閉じる
                    ProgressDialog.endDispose()

                    '完了メッセージ
                    Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_019, MsgBoxStyle.OkOnly)
                Case ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS
                    '進捗ダイアログを閉じる
                    ProgressDialog.endDispose()

                    'エラーメッセージ
                    Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_058, MsgBoxStyle.OkOnly)
            End Select

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
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        End Try
    End Sub

    ''' <summary>
    ''' データを設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetData()
        Dim xlSheet As Excel.Worksheet = Nothing

        Try
            xlSheet = DirectCast(xlSheets.Item(1), Excel.Worksheet)

            'シート保護確認
            If xlSheet.ProtectContents Then
                xlSheet.Unprotect()
            End If

            'ヘッダーを設定
            SetExcelValue(xlSheet, 1, 2, CommonInfo.ChosakubunName)
            SetExcelValue(xlSheet, 2, 2, HEADER_TITLE)

            'データを設定
            If _censusNo.Count > 0 Then
                SetExcelValue(xlSheet, 2, 6, _chosaNen)
                For i As Integer = 0 To _censusNo.Count - 1
                    SetExcelValue(xlSheet, i + 5, 3, _censusNo(i))
                Next
            End If

            xlSheet.Protect()
        Catch ex As Exception
            Throw
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub

    ''' <summary>
    ''' 個別結果表項目指定修正シートデータ取得
    ''' </summary>
    ''' <param name="blNoCensusData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSheetData(Optional blNoCensusData As Boolean = True) As DataTable
        Dim ret As New DataTable
        Dim xlSheet As Excel.Worksheet = Nothing
        Dim rng As Excel.Range = Nothing
        Dim rngChosaNen As Excel.Range = Nothing

        Try
            xlSheet = DirectCast(xlSheets.Item(1), Excel.Worksheet)

            'シート保護確認
            If xlSheet.ProtectContents Then
                xlSheet.Unprotect()
            End If

            Dim arrData(,) As Object
            Dim page As Excel.PageSetup = xlSheet.PageSetup

            rng = xlSheet.Range("C4:F5004")

            arrData = DirectCast(rng.Value, Object(,))

            For i As Integer = 1 To arrData.GetLength(0)
                If i = 1 Then
                    For j As Integer = 1 To arrData.GetLength(1)
                        Dim dc As New DataColumn
                        dc.ColumnName = arrData(i, j).ToString
                        dc.DefaultValue = Nothing
                        dc.DataType = Type.GetType("System.String")
                        ret.Columns.Add(dc)
                    Next
                Else
                    Dim row As DataRow
                    row = ret.NewRow
                    For j As Integer = 1 To arrData.GetLength(1)
                        Dim value As Object = arrData(i, j)
                        row(arrData(1, j).ToString) = If(value Is Nothing OrElse String.IsNullOrEmpty(value.ToString), Nothing, value.ToString)
                    Next
                    ret.Rows.Add(row)
                End If
                '2行目以降で、かつ最終行またはセンサス番号が値なしの場合
                If i >= 2 And (i = arrData.GetLength(0) OrElse (blNoCensusData AndAlso String.IsNullOrEmpty(CStr(arrData(i + 1, 1))))) Then
                    Exit For
                End If
            Next

            rngChosaNen = xlSheet.Range("F2")
            _chosaNen = If(rngChosaNen.Value Is Nothing, Nothing, CStr(rngChosaNen.Value))

        Finally
            ReleaseComObject(xlSheet)
            ReleaseComObject(rng)
            ReleaseComObject(rngChosaNen)
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="dtSheet"></param>
    ''' <param name="dtKobetsuItemMst"></param>
    ''' <param name="details"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(dtSheet As DataTable, dtKobetsuItemMst As DataTable, ByRef details As List(Of String), Optional before As Boolean = False) As Boolean
        Dim ret As Boolean = True
        Dim dt As DataTable = Nothing
        Const max As Integer = ComConst.ERR_MESSAGE_MAX

        Dim msg As String() = {"" _
                             , "{0}件目：調査年（産）を入力してください。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）の「項目番号」を入力してください。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）の「センサス番号」に存在しないセンサス番号が入力されております。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）の「項目番号」に存在しない項目番号が入力されております。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）の「項目番号」は修正不可項目であるため修正できません。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）、項目番号（{3}）の「修正データ」の型が一致しません。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）、項目番号（{3}）の「修正データ」の桁数がデータベースの桁数を超えています。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）、項目番号（{3}）の「修正データ」の桁数が個別結果表作成論理の表示単位桁数を超えています。" _
        }

        Dim kyoku As String = Nothing
        Dim lstCensusNo As List(Of String) = Nothing
        Dim dicKobetsuItemMst As New Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)

        If Not CommonInfo.Koutei.Equals(CommonInfo.KouteiKubun.Code.Honsyo) Then
            kyoku = CommonInfo.Kyoku
        End If

        For Each dr As DataRow In dtKobetsuItemMst.Rows
            Dim item As New DAOKobetsuKekkahyo.個別結果表項目
            With item
                .シート名 = dr("シート名").ToString
                .行位置 = Integer.Parse(dr("行位置").ToString)
                .列位置 = Integer.Parse(dr("列位置").ToString)
                .値 = Nothing
                .型区分 = dr("型区分").ToString
                .有効桁数 = Integer.Parse(dr("有効桁数").ToString)
                .小数点以下桁数 = Integer.Parse(dr("小数点以下桁数").ToString)
                .表示単位 = dr("表示単位").ToString
                .再計算区分 = setEditable(dr)
            End With
            dicKobetsuItemMst.Add(dr("項目番号").ToString, item)
        Next

        Dim cnt As Integer = 0

        '1）調査年（産）は、入力されているか。
        If String.IsNullOrEmpty(_chosaNen) Then
            cnt = cnt + 1
            details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2)))
            ret = False
            If cnt = max Then Return ret
        Else
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                'センサス番号取得
                dt = DAOKobetsuKekkahyo.GetCensusNo(db, CommonInfo.Chosakubun, _chosaNen, kyoku, Nothing, Nothing, _kessokuchiHokan)
                lstCensusNo = (From dr In dt Select CStr(dr("センサス番号"))).ToList
            End Using
        End If

        Dim rowCount As Integer = 0
        For Each row As DataRow In dtSheet.Rows
            Dim censusNo As String = row.Item(0).ToString
            Dim koumokuNo As String = row.Item(1).ToString
            Dim shuuseiData As String = If(before, String.Empty, row.Item(3).ToString)

            rowCount += 1
            '2）項目番号は、入力されているか。
            If String.IsNullOrEmpty(koumokuNo) Then
                cnt = cnt + 1
                details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), rowCount, censusNo))
                ret = False
                If cnt = max Then Return ret
            End If
            '3）センサス番号が存在するか。
            If lstCensusNo Is Nothing OrElse Not lstCensusNo.Contains(censusNo) Then
                cnt = cnt + 1
                details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), rowCount, censusNo))
                ret = False
                If cnt = max Then Return ret
            End If

            '4）項目番号が存在するか。
            '5）項目番号が修正可能項目番号であるか。
            '6）修正データが個別結果表項目マスタの型と一致しているか。
            '7）修正データがデータベースの桁数に収まっているか。
            '8）修正データが個別結果表作成論理に表示単位が設定されている場合、表示単位の桁数に収まっているか。
            If Not dicKobetsuItemMst.ContainsKey(koumokuNo) Then
                If Not String.IsNullOrEmpty(koumokuNo) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), rowCount, censusNo))
                    ret = False
                    If cnt = max Then Return ret
                End If
            ElseIf dicKobetsuItemMst(koumokuNo).再計算区分.Equals("○") Then
                If Not String.IsNullOrEmpty(koumokuNo) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), rowCount, censusNo))
                    ret = False
                    If cnt = max Then Return ret
                End If
            ElseIf Not String.IsNullOrEmpty(shuuseiData) And dicKobetsuItemMst(koumokuNo).型区分 = ComConst.型区分.数値型 Then
                Dim val As Decimal
                If Not Decimal.TryParse(shuuseiData, val) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), rowCount, censusNo, koumokuNo))
                    ret = False
                    If cnt = max Then Return ret
                ElseIf dicKobetsuItemMst(koumokuNo).有効桁数 > 0 Then
                    Dim pattern As String
                    If dicKobetsuItemMst(koumokuNo).小数点以下桁数 > 0 Then
                        pattern = "^-?[0-9]{1," & dicKobetsuItemMst(koumokuNo).有効桁数 - dicKobetsuItemMst(koumokuNo).小数点以下桁数 & "}(\.[0-9]{1," & dicKobetsuItemMst(koumokuNo).小数点以下桁数 & "})?$"
                    Else
                        pattern = "^-?[0-9]{1," & dicKobetsuItemMst(koumokuNo).有効桁数 & "}$"
                    End If
                    If Not Regex.IsMatch(val.ToString, pattern) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), rowCount, censusNo, koumokuNo))
                        ret = False
                        If cnt = max Then Return ret
                    Else
                        If dicKobetsuItemMst(koumokuNo).小数点以下桁数 > 0 Then
                            If Not String.IsNullOrEmpty(dicKobetsuItemMst(koumokuNo).表示単位) Then
                                Dim digit As Integer
                                Dim unit As Decimal = Decimal.Parse(dicKobetsuItemMst(koumokuNo).表示単位)
                                If ComConst.個別結果表作成論理.表示単位.リスト.ContainsKey(unit) Then
                                    digit = dicKobetsuItemMst(koumokuNo).表示単位.TrimEnd("0"c).Substring(dicKobetsuItemMst(koumokuNo).表示単位.TrimEnd("0"c).IndexOf("."c) + 1).Length
                                End If
                                If digit > 0 Then
                                    pattern = "^-?[0-9]{1," & dicKobetsuItemMst(koumokuNo).有効桁数 - dicKobetsuItemMst(koumokuNo).小数点以下桁数 & "}(\.[0-9]{1," & digit & "})?$"
                                Else
                                    pattern = "^-?[0-9]{1," & dicKobetsuItemMst(koumokuNo).有効桁数 - dicKobetsuItemMst(koumokuNo).小数点以下桁数 & "}$"
                                End If
                                If Not Regex.IsMatch(val.ToString, pattern) Then
                                    cnt = cnt + 1
                                    details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2), rowCount, censusNo, koumokuNo))
                                    ret = False
                                    If cnt = max Then Return ret
                                End If
                            End If
                        End If
                    End If
                End If
            End If

        Next
        Return ret
    End Function


    ''' <summary>
    ''' 保存処理
    ''' </summary>
    ''' <param name="progressDialog"></param>
    ''' <param name="db"></param>
    ''' <param name="dtSheet"></param>
    ''' <param name="dtKobetsuItemMst"></param>
    ''' <remarks></remarks>
    Private Sub UpdateDate(ByVal progressDialog As ProgressDialog, db As DBAccess, dtSheet As DataTable, dtKobetsuItemMst As DataTable)

        Dim updateData As New Dictionary(Of String, Dictionary(Of String, String))

        Try

            For Each dr As DataRow In dtSheet.Select(Nothing, dtSheet.Columns.Item(0).ColumnName & " ASC, " & dtSheet.Columns.Item(1).ColumnName & " ASC")
                Dim itemAndValue As New Dictionary(Of String, String)
                Dim censusNo As String = dr(0).ToString
                Dim itemNo As String = dr(1).ToString
                Dim value As String = If(IsDBNull(dr(3)) OrElse String.IsNullOrEmpty(dr(3).ToString), Nothing, dr(3).ToString)

                If updateData.ContainsKey(censusNo) Then
                    If updateData(censusNo).ContainsKey(itemNo) Then
                        updateData(censusNo)(itemNo) = value
                    Else
                        updateData(censusNo).Add(itemNo, value)
                    End If
                Else
                    itemAndValue.Add(itemNo, value)
                    updateData.Add(censusNo, itemAndValue)
                End If
            Next

            '進捗を進める
            progressDialog.AddValue = 1

            For Each kvp As KeyValuePair(Of String, Dictionary(Of String, String)) In updateData
                Dim pKey As New DAOKobetsuKekkahyo.PrimaryKey(_chosaNen, kvp.Key)
                Dim dtKobetsu As Dictionary(Of String, DataTable)
                Dim dcKobetsu As New Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)

                dtKobetsu = DAOKobetsuKekkahyo.GetTable(db, pKey, _kessokuchiHokan)

                dcKobetsu = ComUtil.KobetsuKekkahyo.GetItem(dtKobetsuItemMst, dtKobetsu)

                '文字型補正（空文字→NULL）
                Dim query = (From dr In dtKobetsuItemMst.AsEnumerable().ToList() Where dr("型区分").ToString = ComConst.型区分.文字型 Select dr).ToList
                For Each dr In query
                    If dcKobetsu(dr("項目番号").ToString).値.Equals(String.Empty) Then
                        dcKobetsu(dr("項目番号").ToString).値 = Nothing
                    End If
                Next

                Dim tableName As String = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)
                Dim kKey As New DAOKobetsuKekkahyo.KyotenKey(dtKobetsu(tableName)(0)("農政局").ToString, _
                                                             dtKobetsu(tableName)(0)("都道府県").ToString, _
                                                             dtKobetsu(tableName)(0)("実査設置拠点").ToString)

                For Each _kvp As KeyValuePair(Of String, String) In kvp.Value
                    dcKobetsu(_kvp.Key).値 = _kvp.Value
                Next

                Try
                    db.BeginTrans()

                    '個別結果表データ削除
                    DAOKobetsuKekkahyo.DeleteTable(db, pKey, kKey, _kessokuchiHokan)

                    '個別結果表データ追加
                    DAOKobetsuKekkahyo.InsertTable(db, pKey, kKey, dcKobetsu, _kessokuchiHokan)

                    db.CommitTrans()

                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try

                '進捗を進める
                progressDialog.AddValue = 1

            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 再計算処理を行う
    ''' </summary>
    ''' <param name="progressDialog"></param>
    ''' <param name="db"></param>
    ''' <param name="dtKobetsuItemMst"></param>
    ''' <param name="dtSheet"></param>
    ''' <remarks></remarks>
    Private Sub Recalculate(ByVal progressDialog As ProgressDialog, db As DBAccess, dtKobetsuItemMst As DataTable, dtSheet As DataTable)
        Dim dtChoItemMst As DataTable
        Dim kobetsuList As Dictionary(Of String, Object)
        Dim itemInfoList As List(Of CreateKobetsu.ItemInfo)
        Dim ItemInfoListKobetsu As List(Of CreateKobetsu.ItemInfo)
        Dim dtKobetsuItemMstUra As DataTable
        Dim updateData As New Dictionary(Of String, Dictionary(Of String, String))
        Dim dtCreateRonri As DataTable = Nothing
        Dim taisyakuKubun As String = Nothing

        Try

            '調査票項目マスタ取得
            'REV-001 START----------
            dtChoItemMst = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)
            'REV-001 END----------
            '個別結果表項目マスタ取得(裏項番含める)
            ' REV_002↓
            'dtKobetsuItemMstUra = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, True)
            dtKobetsuItemMstUra = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun), True)
            ' REV_002↑

            For Each dr As DataRow In dtSheet.Select(Nothing, dtSheet.Columns.Item(0).ColumnName & " ASC, " & dtSheet.Columns.Item(1).ColumnName & " ASC")
                Dim itemAndValue As New Dictionary(Of String, String)
                Dim censusNo As String = dr(0).ToString
                Dim itemNo As String = dr(1).ToString
                Dim value As String = If(IsDBNull(dr(3)) OrElse String.IsNullOrEmpty(dr(3).ToString), Nothing, dr(3).ToString)

                If updateData.ContainsKey(censusNo) Then
                    If updateData(censusNo).ContainsKey(itemNo) Then
                        updateData(censusNo)(itemNo) = value
                    Else
                        updateData(censusNo).Add(itemNo, value)
                    End If
                Else
                    itemAndValue.Add(itemNo, value)
                    updateData.Add(censusNo, itemAndValue)
                End If
            Next

            For Each kvp As KeyValuePair(Of String, Dictionary(Of String, String)) In updateData
                Dim pKey As New DAOKobetsuKekkahyo.PrimaryKey(_chosaNen, kvp.Key)
                Dim dtKobetsu As Dictionary(Of String, DataTable)
                Dim dcKobetsu As New Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)

                dtKobetsu = DAOKobetsuKekkahyo.GetTable(db, pKey, _kessokuchiHokan)

                dcKobetsu = ComUtil.KobetsuKekkahyo.GetItem(dtKobetsuItemMst, dtKobetsu)

                Dim tableName As String = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)
                Dim kKey As New DAOKobetsuKekkahyo.KyotenKey(dtKobetsu(tableName)(0)("農政局").ToString, _
                                                             dtKobetsu(tableName)(0)("都道府県").ToString, _
                                                             dtKobetsu(tableName)(0)("実査設置拠点").ToString)

                For Each _kvp As KeyValuePair(Of String, String) In kvp.Value
                    dcKobetsu(_kvp.Key).値 = _kvp.Value
                Next

                If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                    taisyakuKubun = dcKobetsu(ComConst.個別結果表.貸借対照表(ComConst.調査区分.営農類型別経営統計_個人)).値
                    '個別結果表作成論理＿営農個人取得(再計算する論理のみ取得する)
                    ' REV_002↓
                    'dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db, taisyakuKubun, True)
                    dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db, taisyakuKubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun), True)
                    ' REV_002↑
                Else
                    '個別結果表作成論理取得(再計算する論理のみ取得する)
                    ' REV_002↓
                    'dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonri(db, True)
                    dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonri(db, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun), True)
                    ' REV_002↑
                End If

                kobetsuList = New Dictionary(Of String, Object)
                For Each kv As KeyValuePair(Of String, DAOKobetsuKekkahyo.個別結果表項目) In dcKobetsu
                    dcKobetsu(kv.Key).値 = ComUtil.KobetsuKekkahyo.GetData(ComUtil.KobetsuKekkahyo.GetformattedValue(kv.Value), kv.Value.型区分, True)
                    kobetsuList.Add(kv.Key, ComUtil.KobetsuKekkahyo.GetData(ComUtil.KobetsuKekkahyo.GetformattedValue(kv.Value), kv.Value.型区分, True))
                Next
                '個別結果表・個別結果検討表作成クラス
                Dim kobetsu As CreateKobetsu = New CreateKobetsu(db,
                                                                 CommonInfo.Chosakubun,
                                                                 _chosaNen,
                                                                 CreateKobetsu.enmCreateType.個別結果表再計算,
                                                                 dtChoItemMst,
                                                                 dtKobetsuItemMstUra,
                                                                 Nothing,
                                                                 dtCreateRonri,
                                                                 kobetsuList,
                                                                 progressDialog)
                '個別結果表再計算実行
                itemInfoList = kobetsu.Execute(kvp.Key)
                '個別結果表(当年データ、裏項番以外、再計算対象)で抽出
                ItemInfoListKobetsu = (From n In itemInfoList Where n.ItemType = CreateKobetsu.enmItemType.個別結果表 And Not n.ItemNo.Contains("前") And n.IsHidden = False And n.IsReCalc).ToList
                For Each info In ItemInfoListKobetsu
                    '個別結果表データを上書き
                    If dcKobetsu.ContainsKey(info.ItemNo) Then
                        dcKobetsu(info.ItemNo).値 = If(info.Value Is Nothing, Nothing, info.Value.ToString)
                    End If
                Next

                Try
                    db.BeginTrans()

                    '個別結果表データ削除
                    DAOKobetsuKekkahyo.DeleteTable(db, pKey, kKey, _kessokuchiHokan)

                    '個別結果表データ追加
                    DAOKobetsuKekkahyo.InsertTable(db, pKey, kKey, dcKobetsu, _kessokuchiHokan)

                    db.CommitTrans()

                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try

                dtCreateRonri.Clear()
                dtCreateRonri = Nothing
            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 入力ファイル取得
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetInputFile(filePath As String) As List(Of String())
        Dim ret As List(Of String()) = Nothing

        Dim sjisEnc As Encoding = Encoding.GetEncoding(ComConst.CSVファイル.CODEPAGE_SHIFT_JIS)

        'ファイル入力
        If System.IO.File.Exists(filePath) Then
            ret = New List(Of String())

            Using parser As New TextFieldParser(filePath, sjisEnc)

                parser.TextFieldType = FieldType.Delimited
                parser.SetDelimiters(ComConst.CSVファイル.CSV_DELIMITER)

                While Not parser.EndOfData
                    Dim arr As String() = parser.ReadFields()
                    ret.Add(arr)
                End While
            End Using
        End If

        Return ret
    End Function

    ''' <summary>
    ''' ファイルの内容を画面にセット
    ''' </summary>
    ''' <param name="ProgressDialog"></param>
    ''' <param name="lstArr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Sub setCsvData(ByRef ProgressDialog As ProgressDialog, lstArr As List(Of String()))
        Dim xlSheet As Excel.Worksheet = Nothing
        Dim xlRange As Excel.Range = Nothing

        Try
            xlSheet = DirectCast(xlSheets.Item(1), Excel.Worksheet)
            xlRange = xlSheet.Range("C5:F5004")
            xlRange.ClearContents()
            '調査年
            SetExcelValue(xlSheet, 2, 6, lstArr(0)(2))
            '進捗を進める
            ProgressDialog.AddValue = 1
            For i As Integer = 1 To lstArr.Count - 1
                'センサス番号
                SetExcelValue(xlSheet, i + 4, 3, lstArr(i)(0))
                '項目番号
                SetExcelValue(xlSheet, i + 4, 4, lstArr(i)(1))
                '修正前データ
                SetExcelValue(xlSheet, i + 4, 5, lstArr(i)(2))
                '修正データ
                SetExcelValue(xlSheet, i + 4, 6, lstArr(i)(3))
                '進捗を進める
                ProgressDialog.AddValue = 1
            Next

        Catch ex As Exception
            Throw
        Finally
            ReleaseComObject(xlSheet)
            ReleaseComObject(xlRange)
        End Try
    End Sub

    ''' <summary>
    ''' 出力ファイル作成
    ''' </summary>
    ''' <param name="ProgressDialog"></param>
    ''' <param name="filePath"></param>
    ''' <param name="dtSheet"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PutOutputFile(ProgressDialog As ProgressDialog, filePath As String, dtSheet As DataTable) As ComConst.CSVファイル.enmOutputReturn
        Dim ret As ComConst.CSVファイル.enmOutputReturn = ComConst.CSVファイル.enmOutputReturn.CANCEL
        Dim sjisEnc As Encoding = Encoding.GetEncoding(ComConst.CSVファイル.CODEPAGE_SHIFT_JIS)

        'ディレクトリ作成
        Dim outDir As String = Path.GetDirectoryName(filePath)
        If Not System.IO.Directory.Exists(outDir) Then
            Directory.CreateDirectory(outDir)
        End If

        '進捗を進める
        ProgressDialog.AddValue = 1

        'ファイル存在チェック
        If IO.File.Exists(filePath) Then
            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_004, {IO.Path.GetFileName(filePath)}, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                Return ret
            End If
        End If

        '進捗を進める
        ProgressDialog.AddValue = 1

        Try
            'ファイル出力
            Using sw As New System.IO.StreamWriter(filePath, False, sjisEnc)
                sw.WriteLine(ComConst.CSVファイル.START_ADDITION & FILE_TITLE & ComConst.CSVファイル.END_ADDITION & ComConst.CSVファイル.CSV_DELIMITER _
                             & ComConst.CSVファイル.START_ADDITION & CommonInfo.ChosakubunName & ComConst.CSVファイル.END_ADDITION & ComConst.CSVファイル.CSV_DELIMITER _
                             & ComConst.CSVファイル.START_ADDITION & _chosaNen & ComConst.CSVファイル.END_ADDITION)

                '進捗を進める
                ProgressDialog.AddValue = 1

                For Each dr As DataRow In dtSheet.Rows
                    Dim arr As Object() = {dr(0), dr(1), dr(2), dr(3)}
                    sw.WriteLine(ComConst.CSVファイル.START_ADDITION _
                                 & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, arr) _
                                 & ComConst.CSVファイル.END_ADDITION)

                    '進捗を進める
                    ProgressDialog.AddValue = 1

                Next
            End Using
            ret = ComConst.CSVファイル.enmOutputReturn.OK
        Catch ex As Exception
            ret = ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 修正対象センサス番号の個別結果表作成論理(営農個人)の総レコード件数を取得する
    ''' </summary>
    ''' <param name="censusNoList"></param>
    ''' <param name="dtSheet"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCreateRonriRecCntEinoKojin(ByVal censusNoList As List(Of String), ByVal dtSheet As DataTable) As Object()
        Dim countdata As Integer
        Dim ret As Object() = {0, ""}

        Dim dicKobetsu As Dictionary(Of String, DataTable) = Nothing
        Dim taisyakuKubun As String = Nothing

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            For Each censusNo As String In censusNoList
                Dim isExist As Boolean = False
                taisyakuKubun = Nothing

                '修正画面に「貸借対照表区分」が存在するかをチェックする
                '降順にループ(後に出てきた方を優先するため)
                For i As Integer = dtSheet.Rows.Count - 1 To 0 Step -1
                    If dtSheet.Rows(i).Item(0).ToString = censusNo AndAlso
                       dtSheet.Rows(i).Item(1).ToString = ComConst.個別結果表.貸借対照表(ComConst.調査区分.営農類型別経営統計_個人) Then
                        isExist = True
                        taisyakuKubun = dtSheet.Rows(i).Item(3).ToString
                        Exit For
                    End If
                Next

                '修正画面に「貸借対照表区分」が存在する場合
                If isExist Then
                    '個別結果表作成論理のレコード件数を加算する
                    ' REV_002↓
                    'countdata += DAOOther.GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db, taisyakuKubun, True).Rows.Count
                    countdata += DAOOther.GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db, taisyakuKubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun), True).Rows.Count
                    ' REV_002↑

                    If IsDBNull(taisyakuKubun) OrElse String.IsNullOrEmpty(CStr(taisyakuKubun)) Then
                        countdata = -1
                        ret(0) = countdata
                        ret(1) = censusNo
                        Return ret
                        Exit Function
                    End If
                Else
                    Dim pKey As New DAOKobetsuKekkahyo.PrimaryKey(_chosaNen, censusNo)
                    '個別結果表テーブル取得
                    dicKobetsu = DAOKobetsuKekkahyo.GetTable(db, pKey, _kessokuchiHokan)
                    For Each kv As KeyValuePair(Of String, DataTable) In dicKobetsu
                        Dim dtKobetsu As DataTable = kv.Value
                        If dtKobetsu.Rows.Count > 0 AndAlso dtKobetsu.Columns.Contains(ComConst.個別結果表.貸借対照表(ComConst.調査区分.営農類型別経営統計_個人)) Then
                            Dim value As Object = dtKobetsu.Rows(0).Item(ComConst.個別結果表.貸借対照表(ComConst.調査区分.営農類型別経営統計_個人))
                            If Not IsDBNull(value) AndAlso Not String.IsNullOrEmpty(CStr(value)) Then
                                taisyakuKubun = CStr(value)
                                '個別結果表作成論理のレコード件数を加算する
                                ' REV_002↓
                                'countdata += DAOOther.GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db, taisyakuKubun, True).Rows.Count
                                countdata += DAOOther.GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db, taisyakuKubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun), True).Rows.Count
                                ' REV_002↑
                                Exit For
                            Else
                                countdata = -1
                                ret(0) = countdata
                                ret(1) = censusNo
                                Return ret
                                Exit Function
                            End If
                        End If
                    Next
                End If
            Next
        End Using

        ret(0) = countdata
        Return ret
    End Function

    ''' <summary>
    ''' 再計算項目ではないが修正不可、或いは再計算項目だが修正可能である項目番号を判定する
    ''' </summary>
    ''' <param name="dr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function setEditable(dr As DataRow) As String
        Dim ret As String

        '再計算項目ではないが修正不可の項目か
        If ComConst.個別結果表.修正不可項目(CommonInfo.Chosakubun).Contains(dr("項目番号").ToString) Then
            ret = "○"
        Else
            ret = dr("再計算区分").ToString()
        End If

        '【営農個人のみ】再計算項目だが修正可能な項目か
        If CommonInfo.Chosakubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) AndAlso ComConst.個別結果表.営農個人修正可能項目.Contains(dr("項目番号").ToString) Then
            ret = String.Empty
        End If

        Return ret
    End Function
End Class
