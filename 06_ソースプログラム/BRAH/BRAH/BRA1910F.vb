'------------------------------------------------------------------------------------------
'| REV | 変更年月日 | 変更者             | 変更内容
'------------------------------------------------------------------------------------------
'| 001 | 2020/11/10 | TSP                | 新規作成
'| 002 | 2021/02/16 | TSP                | 牛トレサ取込時、余計なチェックをしないよう修正（連絡票No.644）
'------------------------------------------------------------------------------------------

Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO


''' <summary>
''' 牛トレサデータ取込画面
''' </summary>
''' <remarks></remarks>
Public Class BRA1910F

    Private Const YEAR_FROM As String = "0"
    Private Const YEAR_TO As String = "1"
    Private Const MONTH_FROM As String = "0"
    Private Const MONTH_TO As String = "1"
    Private Const KOTAI As String = "0"
    Private Const IDO As String = "1"


    ''' <summary>進捗ダイアログ</summary>
    'Private ProgressDialog As New ProgressDialog()

    ''' <summary>重複キーチェック用ディクショナリー</summary>
    Private mTresaDataDic As Dictionary(Of String, String) = Nothing

    ''' <summary>
    ''' フォームロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA1910F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' [個体情報]「参照」ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRefKotai_Click(sender As Object, e As EventArgs) Handles btnRefKotai.Click
        Try
            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of OpenFileDialog)(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainInPath, IniFileInfo.ExcelInPath), , "csvファイル(*.csv)|*.csv")

            If filePath.Equals(String.Empty) Then
                Exit Sub
            Else
                txtKotaiPath.Text = filePath
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' [異動情報]「参照」ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRefIdo_Click(sender As Object, e As EventArgs) Handles btnRefIdo.Click
        Try
            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of OpenFileDialog)(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainInPath, IniFileInfo.ExcelInPath), , "csvファイル(*.csv)|*.csv")

            If filePath.Equals(String.Empty) Then
                Exit Sub
            Else
                txtIdoPath.Text = filePath
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 取込ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click

        Dim dt As DataTable
        Dim msgRet As MsgBoxResult
        Dim progressDialog As New ProgressDialog()
        Dim blSameFlg As Boolean = False

        Dim iRirekiNum As Integer

        Try
            '画面入力項目単体チェック
            '開始年のチェック
            If YearCheck(txtYearFrom.Text, YEAR_FROM) <> 0 Then
                txtYearFrom.BackColor = Color.Red
                Exit Sub
            Else
                txtYearFrom.BackColor = Color.White
            End If

            '開始月のチェック
            If MonthCheck(txtMonthFrom.Text, MONTH_FROM) <> 0 Then
                txtMonthFrom.BackColor = Color.Red
                Exit Sub
            Else
                txtMonthFrom.BackColor = Color.White
            End If

            '終了年のチェック
            If YearCheck(txtYearTo.Text, YEAR_TO) <> 0 Then
                txtYearTo.BackColor = Color.Red
                Exit Sub
            Else
                txtYearTo.BackColor = Color.White
            End If

            '終了月のチェック
            If MonthCheck(txtMonthTo.Text, MONTH_TO) <> 0 Then
                txtMonthTo.BackColor = Color.Red
                Exit Sub
            Else
                txtMonthTo.BackColor = Color.White
            End If

            '取込データ名称のチェック
            If ImportDataNameCheck(txtImportName.Text) <> 0 Then
                txtImportName.BackColor = Color.Red
                Exit Sub
            Else
                txtImportName.BackColor = Color.White
            End If

            '個体情報ファイルパスのチェック
            If FileExistCheck(txtKotaiPath.Text, KOTAI) <> 0 Then
                txtKotaiPath.BackColor = Color.Red
                Exit Sub
            Else
                txtKotaiPath.BackColor = Color.White
            End If

            '個体情報ファイルパスのチェック
            If FileExistCheck(txtIdoPath.Text, IDO) <> 0 Then
                txtIdoPath.BackColor = Color.Red
                Exit Sub
            Else
                txtIdoPath.BackColor = Color.White
            End If

            '画面入力項目間チェック
            '開始年月と終了年月の大小チェック
            If DateConsistentCheck(txtYearFrom.Text, txtMonthFrom.Text, txtYearTo.Text, txtMonthTo.Text) <> 0 Then
                txtYearFrom.BackColor = Color.Red
                txtMonthFrom.BackColor = Color.Red
                txtYearTo.BackColor = Color.Red
                txtMonthTo.BackColor = Color.Red
                Exit Sub
            End If

            Dim txtFrom As String = txtYearFrom.Text & txtMonthFrom.Text.PadLeft(2, "0"c)
            Dim txtTo As String = txtYearTo.Text & txtMonthTo.Text.PadLeft(2, "0"c)

            '同一期間の登録済チェック
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                dt = DAOOther.GetSamePeriodTresaData(db, txtFrom, txtTo, txtImportName.Text)

                If dt.Rows.Count <> 0 Then
                    msgRet = Message.ShowMsgBox(MessageID.MSG_Q_038, MsgBoxStyle.YesNo)
                    If msgRet = MsgBoxResult.No Then
                        Exit Sub
                    Else
                        blSameFlg = True
                    End If
                Else
                    dt = DAOOther.GetDuplicatePeriodTresaData(db, txtYearFrom.Text & txtMonthFrom.Text.PadLeft(2, "0"c), txtYearTo.Text & txtMonthTo.Text.PadLeft(2, "0"c))
                    If dt.Rows.Count <> 0 Then
                        msgRet = Message.ShowMsgBox(MessageID.MSG_Q_045, MsgBoxStyle.YesNo)
                        If msgRet = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                '個体情報ファイルの読み込み
                Dim lstArrKotai As List(Of String()) = Me.GetInputFile(txtKotaiPath.Text)
                '異動情報ファイルの読み込み
                Dim lstArrIdo As List(Of String()) = Me.GetInputFile(txtIdoPath.Text)

                '個体情報ファイルの1行目のデータ数が、12かどうかを判定
                If lstArrKotai(0).Length <> 12 Then
                    Message.ShowMsgBox(MessageID.MSG_E_080, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                '異動情報ファイルの1行目のデータ数が、8かどうかを判定
                If lstArrIdo(0).Length <> 8 Then
                    Message.ShowMsgBox(MessageID.MSG_E_082, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                '進捗ダイアログの最大値に、個体情報ファイルの行数+異動情報ファイルの行数を設定
                progressDialog.Maximum = lstArrKotai.Count + lstArrIdo.Count

                '進捗ダイアログを表示する
                progressDialog.Show(Me)

                Try
                    'トランザクション開始
                    db.BeginTrans()

                    If blSameFlg = True Then
                        iRirekiNum = CInt(dt.Rows(0)("履歴番号"))
                        '牛トレサ取込履歴テーブルのデータ削除
                        DAOOther.DeleteTresaData(db, iRirekiNum, "牛トレサ取込履歴")
                        '牛トレサ個体情報取込履歴テーブルのデータ削除
                        DAOOther.DeleteTresaData(db, iRirekiNum, "牛トレサ個体情報取込履歴")
                        '牛トレサ異動情報取込履歴テーブルのデータ削除
                        DAOOther.DeleteTresaData(db, iRirekiNum, "牛トレサ異動情報取込履歴")
                    End If

                    '牛トレサ取込履歴テーブルにレコード追加
                    DAOOther.InsertTresaImportHistory(db, txtFrom, txtTo, txtImportName.Text)

                    '牛トレサ取込履歴テーブルに追加したレコードのID値を取得
                    dt = DAOOther.GetSamePeriodTresaData(db, txtFrom, txtTo, txtImportName.Text)
                    'If dt.Rows.Count = 0 Then
                    '    db.RollBackTrans()
                    '    Exit Sub
                    'End If

                    '履歴番号セット
                    iRirekiNum = CInt(dt.Rows(0)("履歴番号"))

                    'エラー格納用変数初期化
                    Dim details As New List(Of String)
                    Dim all_details As New List(Of String)
                    Dim msg As String = "{0}件目：{1}"
                    Dim msgPara As String() = {}

                    'キー重複チェック用変数初期化
                    Dim strDicKey As String
                    mTresaDataDic = New Dictionary(Of String, String)

                    '個体情報ファイルのレコード数分以下の処理を行う
                    For i As Integer = 0 To lstArrKotai.Count - 1

                        progressDialog.AddValue = 1

                        '個体情報ファイルのデータチェック
                        KotaiInfoCheck(lstArrKotai(i), i + 1, details)
                        'エラー数が0件でない場合
                        If details.Count <> 0 Then
                            For Each errData As String In details
                                If all_details.Count < 50 Then
                                    all_details.Add(String.Format(msg, all_details.Count + 1, errData))
                                Else
                                    Exit For
                                End If
                            Next
                            msgPara = {String.Join(vbCrLf, all_details)}
                            details.Clear()
                        Else
                            'エラー数が0件の場合
                            If DAOOther.InsertTresaKotaiHistory(db, lstArrKotai(i), iRirekiNum) Then
                                strDicKey = lstArrKotai(i)(0).PadLeft(16, "0"c) & lstArrKotai(i)(1).PadLeft(16, "0"c)
                                mTresaDataDic.Add(strDicKey, i.ToString)
                            End If
                        End If
                    Next

                    '1件でもエラーがあったら、エラーの内容を表示し処理続行確認メッセージを表示する
                    If msgPara.Length > 0 Then
                        Dim reterr As DialogResult
                        'msgRet = Message.ShowMsgBox(MessageID.MSG_Q_037, {"個体情報ファイル", msgPara(0)}, MsgBoxStyle.YesNo)
                        'If msgRet = MsgBoxResult.No Then
                        '    Exit Sub
                        'End If
                        reterr = Message.ShowMsgFormYesNo(Me, MessageID.MSG_Q_037, {"個体情報ファイル"}.Concat(msgPara).ToArray)
                        If reterr = Windows.Forms.DialogResult.No Then
                            db.RollBackTrans()
                            Exit Sub
                        End If
                    End If

                    details.Clear()
                    all_details.Clear()

                    '異動情報ファイルのレコード数分以下の処理を行う
                    For i As Integer = 0 To lstArrIdo.Count - 1

                        progressDialog.AddValue = 1

                        '異動情報ファイルのデータチェック
                        IdoInfoCheck(lstArrIdo(i), i + 1, details)
                        'エラー数が0件でない場合
                        If details.Count <> 0 Then
                            For Each errData As String In details
                                If all_details.Count < 50 Then
                                    all_details.Add(String.Format(msg, all_details.Count + 1, errData))
                                Else
                                    Exit For
                                End If
                            Next
                            msgPara = {String.Join(vbCrLf, all_details)}
                            details.Clear()
                        Else
                            'エラー数が0件の場合
                            If DAOOther.InsertTresaIdoHistory(db, lstArrIdo(i), iRirekiNum) Then
                                strDicKey = lstArrIdo(i)(0).PadLeft(16, "0"c) & lstArrIdo(i)(1).PadLeft(8, "0"c) & lstArrIdo(i)(2) & lstArrIdo(i)(3).PadLeft(16, "0"c)
                                mTresaDataDic.Add(strDicKey, i.ToString)
                            End If
                        End If
                    Next

                    '1件でもエラーがあったら、エラーの内容を表示し処理続行確認メッセージを表示する
                    If msgPara.Length > 0 Then
                        Dim reterr As DialogResult
                        'msgRet = Message.ShowMsgBox(MessageID.MSG_Q_037, {"異動情報ファイル", msgPara(0)}, MsgBoxStyle.YesNo)
                        'If msgRet = MsgBoxResult.No Then
                        '    Exit Sub
                        'End If
                        reterr = Message.ShowMsgFormYesNo(Me, MessageID.MSG_Q_037, {"異動情報ファイル"}.Concat(msgPara).ToArray)
                        If reterr = Windows.Forms.DialogResult.No Then
                            db.RollBackTrans()
                            Exit Sub
                        End If
                    End If


                    '牛トレサ個体情報テーブルのデータを削除
                    DAOOther.DeleteTresaKotaiInfo(db)

                    '牛トレサ個体情報の更新
                    DAOOther.InsertTresaKotaiInfo(db)

                    '牛トレサ異動情報テーブルのデータを削除
                    DAOOther.DeleteTresaIdoInfo(db)

                    '牛トレサ異動情報の更新
                    DAOOther.InsertTresaIdoInfo(db)

                    'コミット
                    db.CommitTrans()

                    ProgressDialog.endDispose()
                    ProgressDialog = Nothing

                    '完了メッセージ表示
                    Message.ShowMsgBox(MessageID.MSG_I_037, MsgBoxStyle.OkOnly)

                Catch ex As Exception
                    db.RollBackTrans()
                    progressDialog.endDispose()
                    progressDialog = Nothing
                    Throw ex
                End Try
            End Using

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub


    ''' <summary>
    ''' 開始年/終了年 個別チェック
    ''' </summary>
    ''' <param name="txtYear">チェックする年の値</param>
    ''' <param name="iType">0:開始年、1:終了年</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function YearCheck(strYear As String, strType As String) As Integer
        Dim iSystemYear As Integer
        Dim iYear As Integer
        Dim iRet As Integer = -1

        Try
            '年が空の場合は、未入力エラー
            If strYear.Length = 0 Then
                If strType = YEAR_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_069, MsgBoxStyle.OkOnly)
                ElseIf strType = YEAR_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_073, MsgBoxStyle.OkOnly)
                End If
                Return -1
            End If

            'システム年を数値に変換
            iSystemYear = CInt(DateTime.Now.ToString("yyyy"))

            If Integer.TryParse(strYear, iYear) Then
                '入力年を数値に変換
                iYear = CInt(strYear)

                '入力年がシステム年-3～システム年+1の範囲かをチェック
                If iYear >= iSystemYear - 3 And iYear <= iSystemYear + 1 Then
                    iRet = 0
                Else
                    If strType = YEAR_FROM Then
                        Message.ShowMsgBox(MessageID.MSG_E_070, {(iSystemYear - 3).ToString, (iSystemYear + 1).ToString}, MsgBoxStyle.OkOnly)
                    ElseIf strType = YEAR_TO Then
                        Message.ShowMsgBox(MessageID.MSG_E_074, {(iSystemYear - 3).ToString, (iSystemYear + 1).ToString}, MsgBoxStyle.OkOnly)
                    End If
                    iRet = -1
                End If
            Else
                '入力内容が数値でない場合もエラー
                If strType = YEAR_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_070, {(iSystemYear - 3).ToString, (iSystemYear + 1).ToString}, MsgBoxStyle.OkOnly)
                ElseIf strType = YEAR_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_074, {(iSystemYear - 3).ToString, (iSystemYear + 1).ToString}, MsgBoxStyle.OkOnly)
                End If
                iRet = -1
            End If

        Catch ex As Exception
            Throw ex
        End Try
        Return iRet
    End Function

    ''' <summary>
    ''' 開始月/終了月 個別チェック
    ''' </summary>
    ''' <param name="txtMonth">チェックする月の値</param>
    ''' <param name="iType">0:開始月、1:終了月</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function MonthCheck(strMonth As String, strType As String) As Integer
        Dim iMonth As Integer
        Dim iRet As Integer = -1

        Try
            '月が空の場合は、未入力エラー
            If strMonth.Length = 0 Then
                If strType = MONTH_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_071, MsgBoxStyle.OkOnly)
                ElseIf strType = MONTH_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_075, MsgBoxStyle.OkOnly)
                End If
                Return -1
            End If

            If strMonth.Length = 2 AndAlso strMonth.Substring(0, 1) = "0" Then
                If strType = MONTH_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_072, MsgBoxStyle.OkOnly)
                ElseIf strType = MONTH_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_076, MsgBoxStyle.OkOnly)
                End If
                Return -1
            End If

            If Integer.TryParse(strMonth, iMonth) Then
                If iMonth >= 1 And iMonth <= 12 Then
                    iRet = 0
                Else
                    If strType = MONTH_FROM Then
                        Message.ShowMsgBox(MessageID.MSG_E_072, MsgBoxStyle.OkOnly)
                    ElseIf strType = MONTH_TO Then
                        Message.ShowMsgBox(MessageID.MSG_E_076, MsgBoxStyle.OkOnly)
                    End If
                    iRet = -1
                End If
            Else
                '入力内容が数値でない場合もエラー
                If strType = MONTH_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_072, MsgBoxStyle.OkOnly)
                ElseIf strType = MONTH_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_076, MsgBoxStyle.OkOnly)
                End If
                iRet = -1
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return iRet
    End Function

    ''' <summary>
    ''' 取込データ名称チェック
    ''' </summary>
    ''' <param name="txtImportDataName">チェックする取込データ名称の値</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ImportDataNameCheck(strImportDataName As String) As Integer
        Dim iRet As Integer = -1

        Try
            If strImportDataName = "" Then
                Message.ShowMsgBox(MessageID.MSG_E_077, MsgBoxStyle.OkOnly)
                Return -1
            End If

            If System.Text.Encoding.GetEncoding(932).GetByteCount(strImportDataName) <= 100 Then
                iRet = 0
            Else
                Message.ShowMsgBox(MessageID.MSG_E_078, MsgBoxStyle.OkOnly)
                iRet = -1
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return iRet
    End Function

    ''' <summary>
    ''' ファイルパスチェック
    ''' </summary>
    ''' <param name="txtFilePath">チェックするファイルパス</param>
    ''' <param name="iType">0：個体情報ファイルパス、1：異動情報ファイルパス</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FileExistCheck(strFilePath As String, strType As String) As Integer

        Dim iRet As Integer = -1

        Try
            'ファイルパスが空の場合はエラー
            If strFilePath.Length = 0 Then
                If strType = KOTAI Then
                    Message.ShowMsgBox(MessageID.MSG_E_079, MsgBoxStyle.OkOnly)
                ElseIf strType = IDO Then
                    Message.ShowMsgBox(MessageID.MSG_E_081, MsgBoxStyle.OkOnly)
                End If
                Return -1
            End If

            If System.IO.File.Exists(strFilePath) Then
                iRet = 0
            Else
                If strType = KOTAI Then
                    Message.ShowMsgBox(MessageID.MSG_E_080, MsgBoxStyle.OkOnly)
                ElseIf strType = IDO Then
                    Message.ShowMsgBox(MessageID.MSG_E_082, MsgBoxStyle.OkOnly)
                End If
                iRet = -1
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return iRet
    End Function

    ''' <summary>
    ''' 開始/終了年月間整合性チェック
    ''' </summary>
    ''' <param name="txtYearFrom">開始年の値</param>
    ''' <param name="txtMonthFrom">開始月の値</param>
    ''' <param name="txtYearTo">終了年の値</param>
    ''' <param name="txtMonthTo">終了月の値</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DateConsistentCheck(strYearFrom As String, strMonthFrom As String, strYearTo As String, strMonthTo As String) As Integer
        Dim iRet As Integer = -1
        Dim strFrom As String
        Dim strTo As String

        Try
            strFrom = strYearFrom & strMonthFrom.PadLeft(2, "0"c)
            strTo = strYearTo & strMonthTo.PadLeft(2, "0"c)

            If strFrom > strTo Then
                Message.ShowMsgBox(MessageID.MSG_E_083, MsgBoxStyle.OkOnly)
                iRet = -1
            Else
                iRet = 0
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return iRet
    End Function

    ''' <summary>
    ''' 入力ファイル取得
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetInputFile(strFilePath As String) As List(Of String())
        Dim ret As List(Of String()) = Nothing

        Dim sjisEnc As Encoding = Encoding.GetEncoding(ComConst.CSVファイル.CODEPAGE_SHIFT_JIS)

        'ファイル入力
        If System.IO.File.Exists(strFilePath) Then
            ret = New List(Of String())

            Using parser As New TextFieldParser(strFilePath, sjisEnc)

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
    ''' 個体情報識別ファイルデータチェック
    ''' </summary>
    ''' <param name="lstKotai"></param>
    ''' <param name="iLine"></param>
    ''' <param name="details"></param>
    ''' <remarks></remarks>
    Private Sub KotaiInfoCheck(lstKotai As String(), iLine As Integer, ByRef details As List(Of String))
        Try
            '個体識別番号のチェック
            If lstKotai(0).ToString = "" Then
                SetErrorDetails(iLine, "個体識別番号", "未設定", details)
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstKotai(0)) >= 17 Then
                    SetErrorDetails(iLine, "個体識別番号", "桁数不正", details)
                End If
            End If
            '農家団体コードのチェック
            If lstKotai(1).ToString = "" Then
                SetErrorDetails(iLine, "農家団体コード", "未設定", details)
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstKotai(1)) >= 17 Then
                    SetErrorDetails(iLine, "農家団体コード", "桁数不正", details)
                End If
            End If

            'キー重複チェック
            If details.Count = 0 Then
                If mTresaDataDic.ContainsKey(lstKotai(0).PadLeft(16, "0"c) & lstKotai(1).PadLeft(16, "0"c)) Then
                    SetErrorDetails(iLine, "キー重複エラー", "", details)
                End If
            End If

            '農家団体区分CDのチェック
            If lstKotai(2).ToString = "" Then
                'REV_002 DEL START
                'SetErrorDetails(iLine, "農家団体区分CD", "未設定", details)
                'REV_002 DEL END
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstKotai(2)) >= 3 Then
                    SetErrorDetails(iLine, "農家団体区分CD", "桁数不正", details)
                Else
                    If IsNumeric(lstKotai(2)) = False Then
                        SetErrorDetails(iLine, "農家団体区分CD", "数値不正", details)
                    End If
                End If
            End If
            '牛の識別CDのチェック
            If lstKotai(3).ToString = "" Then
                SetErrorDetails(iLine, "牛の識別CD", "未設定", details)
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstKotai(3)) >= 3 Then
                    SetErrorDetails(iLine, "牛の識別CD", "桁数不正", details)
                Else
                    If IsNumeric(lstKotai(3)) = False Then
                        SetErrorDetails(iLine, "牛の識別CD", "数値不正", details)
                    End If
                End If
            End If
            '性別CDのチェック
            If lstKotai(4).ToString = "" Then
                SetErrorDetails(iLine, "性別CD", "未設定", details)
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstKotai(4)) >= 2 Then
                    SetErrorDetails(iLine, "性別CD", "桁数不正", details)
                Else
                    If IsNumeric(lstKotai(4)) = False Then
                        SetErrorDetails(iLine, "性別CD", "数値不正", details)
                    End If
                End If
            End If
            '生年月日のチェック
            If lstKotai(5).ToString = "" Then
                'REV_002 DEL START
                'SetErrorDetails(iLine, "生年月日", "未設定", details)
                'REV_002 DEL END
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstKotai(5)) >= 9 Then
                    SetErrorDetails(iLine, "生年月日", "桁数不正", details)
                Else
                    If IsNumeric(lstKotai(5)) = False Then
                        SetErrorDetails(iLine, "生年月日", "数値不正", details)
                    End If
                End If
            End If
            '母牛個体識別番号のチェック
            If lstKotai(6).ToString = "" Then
                'REV_002 DEL START
                'SetErrorDetails(iLine, "母牛個体識別番号", "未設定", details)
                'REV_002 DEL END
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstKotai(6)) >= 17 Then
                    SetErrorDetails(iLine, "母牛個体識別番号", "桁数不正", details)
                End If
            End If
            '輸入国CDのチェック
            If lstKotai(7).ToString <> "" Then
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstKotai(7)) >= 3 Then
                    SetErrorDetails(iLine, "輸入国CD", "桁数不正", details)
                End If
            End If
            '輸入年月日のチェック
            If lstKotai(8).ToString <> "" Then
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstKotai(8)) >= 9 Then
                    SetErrorDetails(iLine, "輸入年月日", "桁数不正", details)
                Else
                    If IsNumeric(lstKotai(8)) = False Then
                        SetErrorDetails(iLine, "輸入年月日", "数値不正", details)
                    End If
                End If
            End If
            '検疫所コードのチェック
            If lstKotai(9).ToString <> "" Then
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstKotai(9)) >= 17 Then
                    SetErrorDetails(iLine, "検疫所コード", "桁数不正", details)
                End If
            End If
            'アクティブ牛フラグのチェック
            If lstKotai(10).ToString = "" Then
                'REV_002 DEL START
                'SetErrorDetails(iLine, "アクティブ牛フラグ", "未設定", details)
                'REV_002 DEL END
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstKotai(10)) >= 3 Then
                    SetErrorDetails(iLine, "アクティブ牛フラグ", "桁数不正", details)
                Else
                    If IsNumeric(lstKotai(10)) = False Then
                        SetErrorDetails(iLine, "アクティブ牛フラグ", "数値不正", details)
                    End If
                End If
            End If
            '照会日のチェック
            If lstKotai(11).ToString = "" Then
                'REV_002 DEL START
                'SetErrorDetails(iLine, "照会日", "未設定", details)
                'REV_002 DEL END
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstKotai(11)) >= 9 Then
                    SetErrorDetails(iLine, "照会日", "桁数不正", details)
                Else
                    If IsNumeric(lstKotai(11)) = False Then
                        SetErrorDetails(iLine, "照会日", "数値不正", details)
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 異動情報ファイルデータチェック
    ''' </summary>
    ''' <param name="lstIdo"></param>
    ''' <param name="iLine"></param>
    ''' <param name="details"></param>
    ''' <remarks></remarks>
    Private Sub IdoInfoCheck(lstIdo As String(), iLine As Integer, ByRef details As List(Of String))
        Try
            '個体識別番号のチェック
            If lstIdo(0).ToString = "" Then
                SetErrorDetails(iLine, "個体識別番号", "未設定", details)
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstIdo(0)) >= 17 Then
                    SetErrorDetails(iLine, "個体識別番号", "桁数不正", details)
                End If
            End If
            '異動年月日のチェック
            If lstIdo(1).ToString = "" Then
                SetErrorDetails(iLine, "異動年月日", "未設定", details)
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstIdo(1)) >= 9 Then
                    SetErrorDetails(iLine, "異動年月日", "桁数不正", details)
                Else
                    If IsNumeric(lstIdo(1)) = False Then
                        SetErrorDetails(iLine, "異動年月日", "数値不正", details)
                    End If
                End If
            End If
            '異動フラグのチェック
            If lstIdo(2).ToString = "" Then
                SetErrorDetails(iLine, "異動フラグ", "未設定", details)
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstIdo(2)) >= 2 Then
                    SetErrorDetails(iLine, "異動フラグ", "桁数不正", details)
                Else
                    If IsNumeric(lstIdo(2)) = False Then
                        SetErrorDetails(iLine, "異動フラグ", "数値不正", details)
                    End If
                End If
            End If
            '農家団体コードのチェック
            If lstIdo(3).ToString = "" Then
                SetErrorDetails(iLine, "農家団体コード", "未設定", details)
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstIdo(3)) >= 17 Then
                    SetErrorDetails(iLine, "農家団体コード", "桁数不正", details)
                End If
            End If

            'キー重複チェック
            If details.Count = 0 Then
                Dim strKey As String
                strKey = lstIdo(0).PadLeft(16, "0"c) & lstIdo(1).PadLeft(8, "0"c) & lstIdo(2).PadLeft(1, "0"c) & lstIdo(3).PadLeft(16, "0"c)
                If mTresaDataDic.ContainsKey(strKey) Then
                    SetErrorDetails(iLine, "キー重複エラー", "", details)
                End If
            End If

            '農家団体区分CDのチェック
            If lstIdo(4).ToString = "" Then
                'REV_002 DEL START
                'SetErrorDetails(iLine, "農家団体区分CD", "未設定", details)
                'REV_002 DEL END
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstIdo(4)) >= 3 Then
                    SetErrorDetails(iLine, "農家団体区分CD", "桁数不正", details)
                Else
                    If IsNumeric(lstIdo(4)) = False Then
                        SetErrorDetails(iLine, "農家団体区分CD", "数値不正", details)
                    End If
                End If
            End If
            '農家団体名称のチェック
            If lstIdo(5).ToString = "" Then
                'REV_002 DEL START
                'SetErrorDetails(iLine, "農家団体名称", "未設定", details)
                'REV_002 DEL END
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstIdo(5)) >= 85 Then
                    SetErrorDetails(iLine, "農家団体名称", "桁数不正", details)
                End If
            End If
            '飼養地都道府県名のチェック
            If lstIdo(6).ToString = "" Then
                'REV_002 DEL START
                'SetErrorDetails(iLine, "飼養地都道府県名", "未設定", details)
                'REV_002 DEL END
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstIdo(6)) >= 9 Then
                    SetErrorDetails(iLine, "飼養地都道府県名", "桁数不正", details)
                End If
            End If
            '飼養所在地のチェック
            If lstIdo(7).ToString = "" Then
                'REV_002 DEL START
                'SetErrorDetails(iLine, "飼養所在地", "未設定", details)
                'REV_002 DEL END
            Else
                If System.Text.Encoding.GetEncoding(932).GetByteCount(lstIdo(7)) >= 85 Then
                    SetErrorDetails(iLine, "飼養所在地", "桁数不正", details)
                End If
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' エラー詳細セット
    ''' </summary>
    ''' <param name="iLine"></param>
    ''' <param name="strItemName"></param>
    ''' <param name="strErrMessage"></param>
    ''' <param name="details"></param>
    ''' <remarks></remarks>
    Private Sub SetErrorDetails(ByVal iLine As Integer, ByVal strItemName As String, ByVal strErrMessage As String, ByRef details As List(Of String))
        Dim msg As String

        msg = "{0}行目 {1} {2}"

        details.Add(String.Format(msg, iLine, strItemName, strErrMessage))
    End Sub

End Class
