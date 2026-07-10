'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2023.08.09 |大興電子通信        | 要件No.11 数式項目を変更禁止エラーにする
'//            |            |                    |
'//*************************************************************************************************

Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop.Forms
Imports Microsoft.VisualBasic.FileIO

''' <summary>
''' 調査票項目指定修正（EXCEL）画面
''' </summary>
''' <remarks></remarks>
''' 

Public Class BRA9420X
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
    ''' <summary>ロックファイル読取</summary>
    'Private _lockFileStream As System.IO.FileStream
    Private _lockFileStream As Dictionary(Of String, System.IO.FileStream)
    ''' <summary>調査票ブックコレクション</summary>
    Private xlBooksChosahyo As Excel.Workbooks
    ''' <summary>調査票シートコレクション</summary>
    Private xlSheetsChosahyo As Excel.Sheets
    ''' <summary>調査票ブック</summary>
    Private xlBookChosahyo As Excel.Workbook
    ''' <summary>調査票Excelアプリケーション</summary>
    Private xlAppChosahyo As Excel.Application

    ''' <summary>ヘッダータイトル文字列</summary>
    Private Const HEADER_TITLE As String = "調査票項目指定修正画面"

    ''' <summary>ファイルタイトル文字列</summary>
    Private Const FILE_TITLE As String = "調査票項目指定修正"

    ''' <summary>ファイルタイトル文字列(専門調査員)</summary>
    Private Const FILE_TITLE_SENMON As String = "_調査票項目指定修正"

    ''' <summary>Excelユーザーフォームハンドル</summary>
    Private _formHwnd As Win32WindowWrapper

    ''' <summary>調査票項目指定修正画面</summary>
    Private _Form As BRA9410F
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
    Public Sub New(ByRef frm As BRA9410F, chosaNen As String, censusNo As List(Of String))
        MyBase.New(frm, System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.調査票項目指定修正.入力用ファイル名称), True)
        Try
            _Form = frm
            _chosaNen = chosaNen
            _censusNo = censusNo
            _lockFileStream = New Dictionary(Of String, System.IO.FileStream)

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
            '↓2022/02/01 CSVボタン追加
            btnFileInput = ComUtilStrictOff.GetExcelBtnFileInput(uf)
            btnFileOutput = ComUtilStrictOff.GetExcelBtnFileOutput(uf)
            '↑2022/02/01 CSVボタン追加

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
            '↓2022/02/01 CSVボタン追加
            ReleaseComObject(btnFileInput)
            ReleaseComObject(btnFileOutput)
            '↑2022/02/01 CSVボタン追加

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
                ProgressDialog.Maximum = 7
                ProgressDialog.Show(_formHwnd)

                'Excel前処理
                Me.BeforeExcel()

                Dim dtChosahyoItemMst As DataTable
                Dim dtSheet As DataTable
                Dim einoresult As Object() = Nothing
                Dim lockFilePathList As Dictionary(Of String, String) = New Dictionary(Of String, String)
                For Each censusNo As String In _censusNo
                    lockFilePathList.Add(censusNo, ComUtil.GetLockFilePath(_chosaNen, censusNo))
                Next

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                    '調査票項目マスタ取得
                    dtChosahyoItemMst = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)

                    '修正データ取得
                    dtSheet = GetSheetData()

                    'エラーチェック
                    Dim details As New List(Of String)
                    If Not Me.CheckError(dtSheet, dtChosahyoItemMst, details,, True) Then
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

                    '-------------------------------------------
                    '調査票を表示して再計算を実施
                    '(前提として、調査票が使用可能であること)
                    '-------------------------------------------
                    '上位フォルダが存在しない場合、処理終了
                    If Not System.IO.Directory.Exists(IniFileInfo.LockParentPath) Then
                        Message.ShowMsgBox(MessageID.MSG_E_115, {CommonInfo.Koutei}, MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If

                    Dim isLocked = False
                    For Each censusNo As String In _censusNo
                        'ロックファイルが存在する場合、ループを抜ける  'REV_003 ADD
                        Dim lockFilePath As String = lockFilePathList(censusNo)
                        If Me.LockFileExist(_chosaNen, censusNo) Then
                            isLocked = True
                            Exit For
                        End If

                        ' ファイルをオープン状態にする  'REV_003 ADD
                        System.IO.File.Create(lockFilePath).Close()
                        If _lockFileStream.ContainsKey(censusNo) Then
                            _lockFileStream(censusNo) = New System.IO.FileStream(lockFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Write)
                        Else
                            _lockFileStream.Add(censusNo, New System.IO.FileStream(lockFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Write))
                        End If

                        '調査票テンプレートからワークへコピー
                        FileCopy(System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.調査票.入力用ファイル名称(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun)))),
                                 ComUtil.GetTemplateCopyFilePath(_chosaNen, censusNo, True))
                    Next

                    'ロックファイルが1件でも存在する場合、処理終了  'REV_003 ADD
                    If isLocked Then
                        Message.ShowMsgBox(MessageID.MSG_E_105, MsgBoxStyle.OkOnly)

                        Exit Sub
                    End If

                    '保存処理
                    UpdateDate(ProgressDialog, db, dtSheet, dtChosahyoItemMst)

                    '進捗を進める
                    ProgressDialog.AddValue = 1

                    '再計算処理
                    Recalculate(ProgressDialog, db, dtSheet)

                    '進捗を進める
                    ProgressDialog.AddValue = 1

                End Using

                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()

                '完了メッセージ
                Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_016, MsgBoxStyle.OkOnly)

            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            For Each censusNo As String In _censusNo
                If System.IO.File.Exists(ComUtil.GetTemplateCopyFilePath(_chosaNen, censusNo, True)) Then
                    'コピーしたテンプレートファイルの削除
                    Kill(ComUtil.GetTemplateCopyFilePath(_chosaNen, censusNo, True))
                End If

                'ロックファイル削除・ストリーム解放
                If _lockFileStream.ContainsKey(censusNo) AndAlso Not _lockFileStream(censusNo) Is Nothing Then
                    ComUtil.DeleteLockFile(_lockFileStream(censusNo))
                End If
            Next

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

                Dim dtChosahyoItemMst As DataTable
                Dim dtSheet As DataTable

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                    '調査票項目マスタ取得
                    dtChosahyoItemMst = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)

                    '修正データ取得
                    dtSheet = GetSheetData()

                    '進捗最大値を設定
                    ProgressDialog.Maximum = 3 + (From dr In dtSheet Select CStr(dr("センサス番号"))).Distinct().Count

                    'エラーチェック
                    Dim details As New List(Of String)
                    If Not Me.CheckError(dtSheet, dtChosahyoItemMst, details) Then
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
                    UpdateDate(ProgressDialog, db, dtSheet, dtChosahyoItemMst)

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
                    dtItem = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)

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
                        Dim pKey As New DAOChosahyo.PrimaryKey(_chosaNen, dr.Item(0).ToString)
                        Dim dtKobetsu As Dictionary(Of String, DataTable)
                        Dim dcKobetsu As Dictionary(Of String, DAOChosahyo.調査票項目)

                        dtKobetsu = DAOChosahyo.GetChosahyoTable(db, pKey)
                        ''dcKobetsu = ComUtil.KobetsuKekkahyo.GetItem(dtItem, New Dictionary(Of String, DataTable) From {{String.Empty, dtKobetsu}})
                        dcKobetsu = ComUtil.Chosahyo.GetItem(dtItem, dtKobetsu)

                        beforeData.Add(dcKobetsu(dr("項目番号").ToString).値)

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
                .InitialFileName = If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainInPath, IniFileInfo.ExcelInPath) & "\"
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
                If String.IsNullOrEmpty(dtSheet.Rows(i).Item(0).ToString) AndAlso
                   String.IsNullOrEmpty(dtSheet.Rows(i).Item(1).ToString) AndAlso
                   String.IsNullOrEmpty(dtSheet.Rows(i).Item(2).ToString) AndAlso
                   String.IsNullOrEmpty(dtSheet.Rows(i).Item(3).ToString) Then
                    dtSheet.Rows(i).Delete()
                Else
                    Exit For
                End If
            Next

            Dim fileName As String = FILE_TITLE & "_" & _chosaNen & "_" & CommonInfo.ChosakubunName & "_" & CommonInfo.Kyoku & ".csv"
            Dim filePath As String

            'フォルダパス取得
            If CommonInfo.SenmonChosain Then
                fileName = CommonInfo.Chosakubun.PadLeft(2, "0"c) & CommonInfo.ChosakubunName & FILE_TITLE_SENMON & ".csv"
            Else
                fileName = FILE_TITLE & "_" & _chosaNen & "_" & CommonInfo.ChosakubunName & "_" & CommonInfo.Kyoku & ".csv"
            End If

            filePath = CStr(xlApp.GetSaveAsFilename(If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath) & "\" & fileName, "CSVファイル (*.csv), *.csv"))
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
    ''' 調査票項目指定修正シートデータ取得
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
    ''' <param name="dtChosahyoItemMst"></param>
    ''' <param name="details"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(dtSheet As DataTable, dtChosahyoItemMst As DataTable, ByRef details As List(Of String), Optional before As Boolean = False, Optional recalculation As Boolean = False) As Boolean
        Dim ret As Boolean = True
        Dim dt As DataTable = Nothing
        Const max As Integer = ComConst.ERR_MESSAGE_MAX

        Dim msg As String() = {"" _
                             , "{0}件目：調査年（産）を入力してください。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）の「項目番号」を入力してください。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）の「センサス番号」に存在しないセンサス番号が入力されております。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）の「項目番号」に存在しない項目番号が入力されております。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）、項目番号（{3}）の「修正データ」の型が一致しません。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）、項目番号（{3}）の「修正データ」の桁数がデータベースの桁数を超えています。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）、は操作可能なセンサス番号ではありません。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）、の自管轄の都道府県と、取込む電子調査票の都道府県が異なっています。" _
                             , "{0}件目：入力データ内に指定しているセンサス番号が多すぎます。3つ以内の指定としてください。" _
                             , "{0}件目：No（{1}）　センサス番号（{2}）の項目番号に変更禁止となる項目番号が入力されています。"
        }

        Dim kyoku As String = Nothing
        Dim lstCensusNo As List(Of String) = Nothing
        'REV_001↓
        'Dim dicChosahyoItemMst As New Dictionary(Of String, DAOChosahyo.調査票項目)
        Dim dicChosahyoItemMst As New Dictionary(Of String, DAOChosahyo.調査票項目EX)
        'REV_001↑

        If Not CommonInfo.Koutei.Equals(CommonInfo.KouteiKubun.Code.Honsyo) Then
            kyoku = CommonInfo.Kyoku
        End If

        For Each dr As DataRow In dtChosahyoItemMst.Rows
            'REV_001↓
            'Dim item As New DAOChosahyo.調査票項目
            Dim item As New DAOChosahyo.調査票項目EX
            'REV_001↑
            With item
                .シート名 = dr("シート名").ToString
                .行位置 = Integer.Parse(dr("行位置").ToString)
                .列位置 = Integer.Parse(dr("列位置").ToString)
                .値 = Nothing
                .型区分 = dr("型区分").ToString
                .有効桁数 = If(String.IsNullOrEmpty(dr("有効桁数").ToString), 0, Integer.Parse(dr("有効桁数").ToString))
                .小数点以下桁数 = If(String.IsNullOrEmpty(dr("小数点以下桁数").ToString), 0, Integer.Parse(dr("小数点以下桁数").ToString))
                'REV_001↓
                .数式区分 = dr("数式区分").ToString
                'REV_001↑
            End With
            dicChosahyoItemMst.Add(dr("項目番号").ToString, item)
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
                dt = DAOChosahyo.GetCensusNo(db, CommonInfo.Chosakubun, _chosaNen, kyoku, Nothing, Nothing)
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

            '10)変更禁止の項目番号が入力されていないか。
            If {"Q00000101", "Q00001001", "Q00001002", "Q00001003", "Q00001004", "Q00001005", "Q00001006", "Q00002001", "Q00002002", "Q00002003", "Q00002004", "Q00002005", "Q00002006"}.Contains(koumokuNo) _
                Or (ComConst.区分２.畜産物生産費 = CommonInfo.Kubun2 And {"Q00000201"}.Contains(koumokuNo)) _
                Or (ComConst.区分２.農産物生産費 = CommonInfo.Kubun2 And {"Q00001101", "Q00001201"}.Contains(koumokuNo)) _
                Or (ComConst.区分２.営農類型別経営統計 = CommonInfo.Kubun2 And {"Q00000201"}.Contains(koumokuNo)) Then
                cnt = cnt + 1
                details.Add(String.Format(msg(10), cnt.ToString.PadLeft(2), rowCount, censusNo))
                ret = False
                If cnt = max Then Return ret
            End If

            'REV_001↓
            '10)変更禁止の項目番号（数式項目）が入力されていないか。
            If dicChosahyoItemMst.ContainsKey(koumokuNo) AndAlso dicChosahyoItemMst(koumokuNo).数式区分 = ComConst.数式区分.数式である Then
                cnt = cnt + 1
                details.Add(String.Format(msg(10), cnt.ToString.PadLeft(2), rowCount, censusNo))
                ret = False
                If cnt = max Then Return ret
            End If
            'REV_001↑

            '↓2022/02/01 エラーチェック(再計算)
            '再計算のみ
            If recalculation AndAlso Not String.IsNullOrEmpty(censusNo) Then
                'REV START　20220307　連絡票No295
                'Dim dr As DataRow() = dtSheet.Select("センサス番号 = '" + censusNo + "'")

                'REV START 2022/03/25　センサス数の制限を無制限に
                'Dim dv As DataView = Nothing
                'Dim dtDis As DataTable = Nothing
                'dv = New DataView(dtSheet)
                'dtDis = dv.ToTable(True, "センサス番号")

                'センサス番号が4つ以上指定されていないか。
                'If dtDis.Rows.Count >= 4 Then
                '    cnt = cnt + 1
                '    details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2), rowCount, censusNo))
                '    ret = False
                '    If cnt = max Then Return ret
                'End If
                'REVEND 2022/03/25
            End If
            '↑2022/02/01 エラーチェック(再計算)

            '4）項目番号が存在するか。
            '5）修正データが個別結果表項目マスタの型と一致しているか。
            '6）修正データがデータベースの桁数に収まっているか。
            If Not dicChosahyoItemMst.ContainsKey(koumokuNo) Then
                If Not String.IsNullOrEmpty(koumokuNo) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), rowCount, censusNo))
                    ret = False
                    If cnt = max Then Return ret
                End If
            ElseIf Not String.IsNullOrEmpty(shuuseiData) And dicChosahyoItemMst(koumokuNo).型区分 = ComConst.型区分.数値型 Then
                Dim val As Decimal
                If Not Decimal.TryParse(shuuseiData, val) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), rowCount, censusNo, koumokuNo))
                    ret = False
                    If cnt = max Then Return ret
                ElseIf dicChosahyoItemMst(koumokuNo).有効桁数 > 0 Then
                    Dim pattern As String
                    If dicChosahyoItemMst(koumokuNo).小数点以下桁数 > 0 Then
                        pattern = "^-?[0-9]{1," & dicChosahyoItemMst(koumokuNo).有効桁数 - dicChosahyoItemMst(koumokuNo).小数点以下桁数 & "}(\.[0-9]{1," & dicChosahyoItemMst(koumokuNo).小数点以下桁数 & "})?$"
                    Else
                        pattern = "^-?[0-9]{1," & dicChosahyoItemMst(koumokuNo).有効桁数 & "}$"
                    End If
                    If Not Regex.IsMatch(val.ToString, pattern) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), rowCount, censusNo, koumokuNo))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
            End If

            '修正データ前表示処理の場合は、次行へ
            If String.IsNullOrEmpty(shuuseiData) Then
                Continue For
            End If

            '7）操作しているユーザが専門調査員の場合、その専門調査員が扱えるセンサス番号かどうか。
            If CommonInfo.SenmonChosain Then
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    If Not DAOOther.CheckSenmonChosainKyakutaiExist(db, CommonInfo.UserId, censusNo) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), rowCount, censusNo))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End Using
            End If

            '8）農業経営統計調査システムにログインしている実査設置拠点を管轄している都道府県と、電子調査票の都道府県（センサス番号内）が一致しているか。
            If Not CommonInfo.Jimusyo = ComUtil.ConvJimusyoNoKoumokuShitei(censusNo).PadLeft(2, "0"c) And CommonInfo.Koutei = "S" Then
                cnt = cnt + 1
                details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2), rowCount, censusNo))
                ret = False
                If cnt = max Then Return ret
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
    ''' <param name="dtChosahyoItemMst"></param>
    ''' <remarks></remarks>
    Private Sub UpdateDate(ByVal progressDialog As ProgressDialog, db As DBAccess, dtSheet As DataTable, dtChosahyoItemMst As DataTable)

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
                Dim pKey As New DAOChosahyo.PrimaryKey(_chosaNen, kvp.Key)
                Dim dtKobetsu As Dictionary(Of String, DataTable)
                Dim dcKobetsu As New Dictionary(Of String, DAOChosahyo.調査票項目)

                dtKobetsu = DAOChosahyo.GetChosahyoTable(db, pKey)

                dcKobetsu = ComUtil.Chosahyo.GetItem(dtChosahyoItemMst, dtKobetsu)

                '文字型補正（空文字→NULL）
                Dim query = (From dr In dtChosahyoItemMst.AsEnumerable().ToList() Select dr).ToList
                For Each dr In query
                    Try
                        If dcKobetsu(dr("項目番号").ToString).値.Equals(String.Empty) Then
                            dcKobetsu(dr("項目番号").ToString).値 = Nothing
                        End If
                    Catch ex As Exception

                    End Try
                Next

                Dim tableName As String = ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)
                Dim kKey As New DAOChosahyo.KotenKey(dtKobetsu(tableName)(0)("農政局").ToString,
                                                             dtKobetsu(tableName)(0)("都道府県").ToString,
                                                             dtKobetsu(tableName)(0)("実査設置拠点").ToString)

                For Each _kvp As KeyValuePair(Of String, String) In kvp.Value
                    dcKobetsu(_kvp.Key).値 = _kvp.Value
                Next

                '↓ 2022/02/07 調査対象経営体・担当名称取得
                Dim dtChosahyo As DataTable             '調査票リスト
                Dim chosaSoshiki As String = Nothing    '調査対象経営体
                Dim tantoMeisho As String = Nothing     '担当名称

                dtChosahyo = DAOChosahyo.GetChosahyoList(db, pKey.chosaNen, kKey.kyoku, kKey.jimusho, kKey.kyoten, Nothing, Nothing, Nothing)
                For Each row As DataRow In dtChosahyo.Rows
                    If pKey.censusNo = row(3).ToString Then
                        chosaSoshiki = row(5).ToString
                        tantoMeisho = row(6).ToString
                    End If
                Next
                '↑ 2022/02/07 調査対象経営体・担当名称取得

                Try
                    db.BeginTrans()

                    '調査票データ削除
                    DAOChosahyo.DeleteChosahyoTableKahenIgai(db, pKey, kKey)

                    '調査票データ追加
                    '↓ 2022/02/07 引数変更
                    'DAOChosahyo.InsertChosahyoTableKahenIgai(db, pKey, kKey, dcKobetsu, "", "")
                    DAOChosahyo.InsertChosahyoTableKahenIgai(db, pKey, kKey, dcKobetsu, chosaSoshiki, tantoMeisho)
                    '↑ 2022/02/07 引数変更

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
    ''' <param name="dtSheet"></param>
    ''' <remarks></remarks>
    Private Sub Recalculate(ByVal progressDialog As ProgressDialog, db As DBAccess, dtSheet As DataTable)

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

            xlAppChosahyo = New Excel.Application
            xlBooksChosahyo = xlAppChosahyo.Workbooks

            For Each kvp As KeyValuePair(Of String, Dictionary(Of String, String)) In updateData
                Dim pKey As New DAOChosahyo.PrimaryKey(_chosaNen, kvp.Key)                                      '調査票用主キー
                Dim kKey As New DAOChosahyo.KotenKey(CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)   '拠点キー

                'Workbookを開く
                Me.ShowExcel(ComUtil.GetTemplateCopyFilePath(_chosaNen, pKey.censusNo, True))

                SetData(xlSheetsChosahyo, pKey, kKey)

                ' 計算処理実行
                If ModuleKeisanClick() = False Then
                    If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_051, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                        Return
                    End If
                End If

                Dim dtItem As DataTable
                Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)
                Dim dtChosahyo As DataTable
                Dim chosaSoshiki As String = Nothing
                Dim tantoMeisho As String = Nothing

                dtChosahyo = DAOChosahyo.GetChosahyoList(db, pKey.chosaNen, kKey.kyoku, kKey.jimusho, kKey.kyoten, Nothing, Nothing, Nothing)
                For Each row As DataRow In dtChosahyo.Rows
                    If pKey.censusNo = row(3).ToString Then
                        chosaSoshiki = row(5).ToString
                        tantoMeisho = row(6).ToString
                    End If
                Next

                '調査票項目マスタ取得
                dtItem = MyGetChosahyoItemMaster(db, CommonInfo.Chosakubun, pKey.chosaNen)

                '調査票シートデータ取得
                dcChosahyo = MyGetSheetData(dtItem, xlSheetsChosahyo, CType(Me, ComObjectProcess), pKey.chosaNen)

                Try
                    db.BeginTrans()

                    '調査票データ削除
                    DAOChosahyo.DeleteChosahyoTable(db, pKey, kKey)

                    '調査票データ追加
                    DAOChosahyo.InsertChosahyoTable(db, pKey, kKey, dcChosahyo, chosaSoshiki, tantoMeisho)

                    db.CommitTrans()
                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try

                'Sheetsの解放
                ReleaseComObject(xlSheetsChosahyo)
                'Workbookを閉じる
                If xlBookChosahyo IsNot Nothing Then
                    xlBookChosahyo.Saved = True
                    xlBookChosahyo.Close(False, Type.Missing, Type.Missing)
                End If
                'Workbookの解放
                ReleaseComObject(xlBookChosahyo)

            Next

            '進捗を進める
            progressDialog.AddValue = 1

        Catch ex As Exception
            Throw ex
        Finally
            'Excelを閉じる(調査票)
            CloseExcelChosahyo()
        End Try

    End Sub

    ''' <summary>
    ''' Excel画面を閉じる(調査票)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CloseExcelChosahyo()
        Try
            '変更無効化
            If xlBookChosahyo IsNot Nothing Then
                xlBookChosahyo.Saved = True
                xlBookChosahyo.Close(False, Type.Missing, Type.Missing)
            End If

            'リボン表示
            xlAppChosahyo.ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"",True)")

            '画面更新無効解除
            xlAppChosahyo.ScreenUpdating = True

            'SheetsChosahyoの解放
            ReleaseComObject(xlSheetsChosahyo)
            'WorkbookChosahyoの解放
            ReleaseComObject(xlBookChosahyo)
            'WorkbooksChosahyoの解放
            ReleaseComObject(xlBooksChosahyo)
            'ガベージコレクト強制
            GCCollect()
            'Excelの終了
            xlAppChosahyo.Quit()
            'ExcelApplicationの解放
            ReleaseComObject(xlAppChosahyo)
            'ガベージコレクト強制
            GCCollect()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Excel画面を開く(調査票)
    ''' </summary>
    ''' <param name="pFilePath"></param>
    ''' <remarks></remarks>
    Private Sub ShowExcel(ByVal pFilePath As String)
        Try
            'Excelアプリケーションの設定
            xlAppChosahyo.Visible = False
            xlAppChosahyo.UserControl = False

            'リボン非表示
            xlAppChosahyo.ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"",False)")
            '画面更新無効
            xlAppChosahyo.ScreenUpdating = False

            ''Excelウィンドウの表示サイズを変更(最大化)
            'xlAppChosahyo.WindowState = Excel.XlWindowState.xlMaximized

            'Excelオブジェクトの設定
            xlBooksChosahyo = xlAppChosahyo.Workbooks
            xlBookChosahyo = xlBooksChosahyo.Open(pFilePath)
            xlSheetsChosahyo = xlBookChosahyo.Worksheets
            Dim xlSheet As Excel.Worksheet = Nothing
            Try
                xlSheet = DirectCast(xlSheetsChosahyo.Item(1), Excel.Worksheet)
                xlSheet.Activate()
            Finally
                ReleaseComObject(xlSheet)
            End Try

        Catch ex As Exception
            If xlAppChosahyo IsNot Nothing Then
                CloseExcelChosahyo()
            End If
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' データを設定する(調査票)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetData(xlSheets As Excel.Sheets, pKey As DAOChosahyo.PrimaryKey, ByRef kKey As DAOChosahyo.KotenKey)
        Try

            Dim dtItem As DataTable
            Dim dtChosahyo As Dictionary(Of String, DataTable)
            Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)
            Dim tableName As String = ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0)

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '調査票項目マスタ取得
                dtItem = MyGetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen, ComConst.数式区分.数式ではない)

                '調査票テーブル取得
                dtChosahyo = DAOChosahyo.GetChosahyoTable(db, pKey)
            End Using

            '調査票項目取得
            dcChosahyo = ComUtil.Chosahyo.GetItem(dtItem, dtChosahyo)

            kKey = New DAOChosahyo.KotenKey(dtChosahyo(tableName)(0)("農政局").ToString,
                                            dtChosahyo(tableName)(0)("都道府県").ToString,
                                            dtChosahyo(tableName)(0)("実査設置拠点").ToString)

            xlAppChosahyo.Interactive = False

            'Excel前処理(調査票)
            BeforeExcelChosahyo()

            '調査票シートデータ設定
            ComUtil.Chosahyo.SetSheetData(dcChosahyo, xlSheets, CType(Me, ComObjectProcess))

        Catch ex As Exception
            Throw
        Finally
            'Excel後処理(調査票)
            AfterExcelChosahyo()
            xlAppChosahyo.Interactive = True
        End Try
    End Sub

    ''' <summary>
    ''' Excel前処理(調査票)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub BeforeExcelChosahyo()
        Try

            ''手動計算にする
            'xlAppChosahyo.Calculation = Excel.XlCalculation.xlCalculationManual
            '画面更新を無効にする
            xlAppChosahyo.ScreenUpdating = False
            'イベント発生を無効にする
            xlAppChosahyo.EnableEvents = False
            'アラート表示を無効にする
            xlAppChosahyo.DisplayAlerts = False

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Excel後処理(調査票)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AfterExcelChosahyo()
        Try

            'アラート表示を有効にする
            xlAppChosahyo.DisplayAlerts = True
            'イベント発生を有効にする
            xlAppChosahyo.EnableEvents = True
            '画面更新を有効にする
            xlAppChosahyo.ScreenUpdating = False
            ''自動計算にする
            'xlAppChosahyo.Calculation = Excel.XlCalculation.xlCalculationAutomatic

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 調査票項目マスタ取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="suushikiKubun"></param>
    ''' <returns></returns>
    ''' <remarks>'REV_005 ADD</remarks>
    Private Function MyGetChosahyoItemMaster(db As DBAccess, chosaKubun As String, chosaNen As String, Optional suushikiKubun As String = Nothing) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        'バージョン区分を調査年から指定する
        Dim ver_kubun = ComUtil.getVersionKubun(chosaNen, CommonInfo.Chosakubun)

        Try

            Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where Not val.Contains("＿可変")

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT     A.*")
                .AppendLine("         , B.precision AS 有効桁数 ")
                .AppendLine("         , B.scale AS 小数点以下桁数")
                .AppendLine("         , C.項目名 AS 出力項目名")
                .AppendLine("FROM      (SELECT * ")
                .AppendLine("           FROM   調査票項目マスタ ")
                .AppendLine("           WHERE  調査区分 = @調査区分 ")
                .AppendLine("           AND  バージョン区分 = @バージョン区分 ")
                If Not suushikiKubun Is Nothing Then
                    .AppendLine("           AND    数式区分 = @数式区分 ")
                End If
                .AppendLine("           AND 行位置 <> 0")
                .AppendLine("           AND 列位置 <> 0")
                .AppendLine("          ) A")
                .AppendLine("LEFT JOIN (SELECT * ")
                .AppendLine("           FROM   sys.columns")
                .AppendLine("           WHERE  object_id IN (SELECT object_id")
                .AppendLine("                                FROM   sys.tables")
                .AppendLine("                                WHERE  name IN ('" & String.Join("', '", query) & "')")
                .AppendLine("                               )")
                .AppendLine("          ) B")
                .AppendLine("ON A.項目番号 = B.name")
                .AppendLine("LEFT JOIN BRAH.dbo.調査票項目名マスタ C ")
                .AppendLine("ON A.項目番号 = C.項番")
                .AppendLine("AND A.調査区分 = C.調査区分")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosaKubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, ver_kubun))
            If Not suushikiKubun Is Nothing Then
                para.Add(db.CreateParameter("@数式区分", SqlDbType.Int, suushikiKubun))
            End If

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票シートデータ取得
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="xlSheets"></param>
    ''' <param name="comObject"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function MyGetSheetData(dt As DataTable, xlSheets As Excel.Sheets, comObject As ComObjectProcess, _chosaNen As String) As Dictionary(Of String, DAOChosahyo.調査票項目)
        Dim ret As New Dictionary(Of String, DAOChosahyo.調査票項目)


        Dim sheets = From dr In dt Group dr By dr!シート名 Into Group Select シート名
        Dim xlSheet As Excel.Worksheet = Nothing
        Dim rng As Excel.Range = Nothing

        For Each sheet In sheets
            Try
                xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

                'シート保護確認
                If xlSheet.ProtectContents Then
                    xlSheet.Unprotect()
                End If

                rng = Nothing
                Try
                    Dim arrData(,) As Object
                    rng = xlSheet.Range(ComConst.調査票.シートデータ範囲(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun)))(sheet.ToString))

                    arrData = DirectCast(rng.Value, Object(,))

                    Dim query = From dr In dt Where dr("シート名").ToString = sheet.ToString Select dr

                    Dim seidouketori_count As Integer = 9010001 'REV-010 ADD

                    For Each dr As DataRow In query
                        If dr("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                            Dim item As New DAOChosahyo.調査票項目
                            With item
                                '"Q090100"で始まる項目番号は処理しない
                                If dr("項目番号").ToString.StartsWith("Q090100") Then
                                    Continue For
                                End If

                                .シート名 = dr("シート名").ToString
                                .行位置 = Integer.Parse(dr("行位置").ToString)

                                .列位置 = Integer.Parse(dr("列位置").ToString)

                                If .行位置 = 0 Or .列位置 = 0 Then
                                    .値 = Nothing
                                Else
                                    .値 = If(arrData(.行位置, .列位置) Is Nothing, Nothing, If(String.IsNullOrEmpty(arrData(.行位置, .列位置).ToString), Nothing, arrData(.行位置, .列位置).ToString))
                                    If .値 IsNot Nothing Then
                                        ComUtil.Chosahyo.EscapeToiawasesaki(dr, .値)
                                    End If
                                End If

                                .型区分 = dr("型区分").ToString
                                .有効桁数 = Integer.Parse(dr("有効桁数").ToString)
                                .小数点以下桁数 = Integer.Parse(dr("小数点以下桁数").ToString)
                            End With
                            ret.Add(dr("項目番号").ToString, item)
                        Else
                            Dim increment As Integer = Integer.Parse(dr("可変増量").ToString)
                            For i As Integer = 1 To Integer.Parse(dr("可変最大数").ToString)
                                Dim item2 As New DAOChosahyo.調査票項目
                                With item2
                                    .シート名 = dr("シート名").ToString
                                    .行位置 = Integer.Parse(dr("行位置").ToString) + If(dr("可変方向").ToString = ComConst.可変方向.縦, (i - 1) * increment, 0)
                                    .列位置 = Integer.Parse(dr("列位置").ToString) + If(dr("可変方向").ToString = ComConst.可変方向.横, (i - 1) * increment, 0)
                                    .値 = If(arrData(.行位置, .列位置) Is Nothing, Nothing, If(String.IsNullOrEmpty(arrData(.行位置, .列位置).ToString), Nothing, arrData(.行位置, .列位置).ToString))
                                    .型区分 = dr("型区分").ToString
                                End With
                                If Not String.IsNullOrEmpty(item2.値) Then
                                    ret.Add(dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & i.ToString, item2)
                                End If
                            Next
                        End If
                    Next
                Finally
                    comObject.ReleaseComObject(rng)
                End Try

            Finally
                comObject.ReleaseComObject(xlSheet)
            End Try
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 計算処理実行
    ''' </summary>
    ''' <returns></returns>
    Public Function ModuleKeisanClick() As Boolean
        Dim ret_val As Boolean = True

        Try
            '自動計算にする
            xlAppChosahyo.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            ret_val = True

        Catch ex As Exception
            ret_val = False
        End Try

        Return ret_val

    End Function

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
    ''' ロックファイル存在チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>REV_003 ADD</remarks>
    Private Function LockFileExist(chosaNen As String, censusNo As String) As Boolean
        Dim ret As Boolean = False
        Dim lockDirectoryPath = IniFileInfo.LockPath

        'ロックファイル用フォルダが存在しない場合、作成する
        If Not System.IO.Directory.Exists(lockDirectoryPath) Then
            System.IO.Directory.CreateDirectory(lockDirectoryPath)
        End If

        'ロックファイルが存在する場合、メッセージ表示→処理終了
        If System.IO.File.Exists(ComUtil.GetLockFilePath(chosaNen, censusNo)) Then
            ret = True
        End If

        Return ret
    End Function

End Class
