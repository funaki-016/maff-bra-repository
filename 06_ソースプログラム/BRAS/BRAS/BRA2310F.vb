Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2023.04.28 |Daiko               | 変更要件No.3
'//*************************************************************************************************

''' <summary>
''' 牛資産異動情報取込画面
''' </summary>
''' <remarks></remarks>
Public Class BRA2310F

#Region "変数一覧"

    ''' <summary>調査年</summary>
    Private _chosaNen As String
    ''' <summary>局</summary>
    Private _kyoku As String
    ''' <summary>事務所</summary>
    Private _jimusho As String
    ''' <summary>拠点</summary>
    Private _kyoten As String
    ''' <summary>営農類型</summary>
    Private _einouRuikei As String

    Private Const YEAR_FROM As String = "0"
    Private Const YEAR_TO As String = "1"
    Private Const MONTH_FROM As String = "0"
    Private Const MONTH_TO As String = "1"

    Private Const C_OK As Integer = 0
    Private Const C_NG As Integer = -1
    Private Const C_Cancel As Integer = 1

    '要件No.22
    Private Shared filetypr As Long

    ''' <summary>進捗ダイアログ</summary>
    Private ProgressDialog As ProgressDialog

#End Region

#Region "【処理詳細仕様 1】初期表示"

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA2310F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try


        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region

#Region "【処理詳細仕様 2】「参照」ボタンクリック"
    ''' <summary>
    ''' 参照ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRefer_Click(sender As Object, e As EventArgs) Handles btnRefer.Click
        Try
            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of OpenFileDialog)(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainInPath, IniFileInfo.ExcelInPath), , "Excelファイル(*.xlsx;*.xlsm)|*.xlsx;*.xlsm")

            If filePath.Equals(String.Empty) Then
                Exit Sub
            Else
                txtFilePath.Text = filePath
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region

#Region "【処理詳細仕様 3】「取込」ボタンクリック"
    ''' <summary>
    ''' 取込ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

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


            '取込ファイルパスのチェック
            If FileExistCheck(txtFilePath.Text) <> 0 Then
                txtFilePath.BackColor = Color.Red
                Exit Sub
            Else
                txtFilePath.BackColor = Color.White
            End If

            'ファイル種類のチェック
            Dim txtchange As String = txtFilePath.Text
            Dim filetyprBoolean As Boolean = txtchange.Contains("成畜")
            If filetyprBoolean = True Then
                filetypr = 1
            Else
                filetypr = 0
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

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_048, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If

            Dim ret As Integer
            '調査票取込クラス生成
            Using ImportIdoTresa = New ImportIdoTresa(txtYearFrom.Text, txtMonthFrom.Text, txtYearTo.Text, txtMonthTo.Text)
                '処理実行
                ret = ImportIdoTresa.Execute(txtFilePath.Text, Me)
            End Using
            If ret = C_OK Then
                Message.ShowMsgBox(MessageID.MSG_I_043, MsgBoxStyle.OkOnly)
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

#End Region

#Region "処理全般"

    ''' <summary>
    ''' 開始年/終了年 個別チェック
    ''' </summary>
    ''' <param name="txtYear">チェックする年の値</param>
    ''' <param name="iType">0:開始年、1:終了年</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function YearCheck(strYear As String, strType As String) As Integer
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

            If Integer.TryParse(strYear, iYear) Then

                iRet = 0
            Else
                '入力値が数値でない場合もエラー
                If strType = YEAR_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_069, MsgBoxStyle.OkOnly)
                ElseIf strType = YEAR_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_073, MsgBoxStyle.OkOnly)
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
    ''' ファイルパスチェック
    ''' </summary>
    ''' <param name="txtFilePath">チェックするファイルパス</param>
    ''' <param name="iType">0：個体情報ファイルパス、1：異動情報ファイルパス</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FileExistCheck(strFilePath As String) As Integer

        Dim iRet As Integer = -1

        Try
            'ファイルパスが空の場合はエラー
            If strFilePath.Length = 0 Then
                Message.ShowMsgBox(MessageID.MSG_E_087, MsgBoxStyle.OkOnly)
                Return -1
            End If

            If System.IO.File.Exists(strFilePath) Then
                iRet = 0
            Else
                Message.ShowMsgBox(MessageID.MSG_E_088, MsgBoxStyle.OkOnly)
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

#End Region

    ''' <summary>
    ''' 牛資産異動情報取込クラス
    ''' </summary>
    ''' <remarks></remarks>
    Private Class ImportIdoTresa
        Inherits ExcelProcess

        ''' <summary>開始年</summary>
        Private _yearFrom As String
        ''' <summary>開始月</summary>
        Private _monthFrom As String
        ''' <summary>終了年</summary>
        Private _yearTo As String
        ''' <summary>終了月</summary>
        Private _monthTo As String



        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="chosaNen"></param>
        ''' <remarks></remarks>
        Public Sub New(yearFrom As String, monthFrom As String, yearTo As String, monthTo As String)
            MyBase.New(True)

            _yearFrom = yearFrom
            _monthFrom = monthFrom
            _yearTo = yearTo
            _monthTo = monthTo

        End Sub

        ''' <summary>
        ''' 処理実行
        ''' </summary>
        ''' <param name="files"></param>
        ''' <param name="myParent"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Execute(files As String, myParent As Form) As Integer
            Dim impRet As Integer
            Dim ok As New List(Of String)
            Dim ng As New List(Of String)

            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Sheets = Nothing

            '調査票項目マスタ取得
            Dim dtItem As DataTable
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                dtItem = DAOOther.GetChosahyoItemMasterIdoTresa(db, CommonInfo.Chosakubun)
            End Using

            'Excelアプリ無効
            Me.DisableExcelApp()

            Try
                'Workbookを開く
                xlBook = xlBooks.Open(files)

                Try
                    xlSheets = xlBook.Worksheets

                    Dim msgId As String = String.Empty
                    Dim msgPara As String() = {}

                    '取込処理
                    impRet = Me.Import(xlSheets, dtItem, msgId, msgPara)
                    If impRet = C_NG Then
                        Select Case msgId
                            Case MessageID.MSG_E_008
                                If msgPara.Length > 0 Then
                                    Message.ShowMsgForm(myParent, msgId, {IO.Path.GetFileName(files)}.Concat(msgPara).ToArray)
                                End If
                            Case MessageID.MSG_E_033
                                Message.ShowMsgBox(msgId, {files}, MsgBoxStyle.OkOnly)
                        End Select
                    End If

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

            'Excelアプリ有効
            Me.EnableExcelApp()
            Return impRet
        End Function

        ''' <summary>
        ''' 取込処理
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="dtItem"></param>
        ''' <param name="msgId"></param>
        ''' <param name="msgPara"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Import(xlSheets As Excel.Sheets, dtItem As DataTable, ByRef msgId As String, ByRef msgPara As String()) As Integer
            Dim ret As Integer = C_NG
            Dim msgRet As MsgBoxResult

            'Excelシート存在チェック
            If Not Me.CheckExcelSheetExist(dtItem, xlSheets) Then
                'メッセージ設定
                msgId = MessageID.MSG_E_033
                Return ret
            End If

            Dim dcChosahyoCheck As Dictionary(Of String, DAOChosahyo.調査票項目)

            'エラーチェック用の牛トレサ異動データ取得
            dcChosahyoCheck = ComUtil.Tresa.GetSheetDataIdoTresa(dtItem, xlSheets, CType(Me, ComObjectProcess))

            'エラーチェック
            Dim details As New List(Of String)
            If Not Me.CheckError(xlSheets, dcChosahyoCheck, details) Then
                'メッセージ設定
                msgId = MessageID.MSG_E_008
                msgPara = {String.Join(vbCrLf, details)}
                Return ret
            End If

            '調査年取得
            Dim ChosaNen = dcChosahyoCheck(ComConst.調査票.項目番号.調査年).値
            '異動月の下限,上限の設定
            Dim Kessan As String
            Dim IdoLow As String
            Dim IdoUp As String
            Dim KessanNum = ComConst.牛トレサデータ.決算月_項目番号(CommonInfo.Chosakubun)
            If KessanNum = "" Then
                Kessan = "12"
                IdoLow = ChosaNen & "01"
                IdoUp = ChosaNen & "12"
            Else
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    Kessan = DAOChosahyo.GettoresaKessan(db, ChosaNen, ComUtil.Chosahyo.GetCensusNo(dcChosahyoCheck), CommonInfo.Chosakubun, KessanNum)
                End Using
                Dim theday As New DateTime(2018, 1, 1)
                IdoLow = theday.AddMonths(CInt(Kessan)).ToString("yyyyMM")
                IdoUp = ChosaNen & Kessan.PadLeft(2, "0"c)
            End If

            Dim fromKikan = _yearFrom & _monthFrom.PadLeft(2, "0"c)
            Dim toKikan = _yearTo & _monthTo.PadLeft(2, "0"c)

            Dim idoFrom As String
            Dim idoTo As String

            '削除範囲の決定(開始)
            If (CInt(IdoLow) >= CInt(fromKikan)) Then
                If (CInt(IdoLow) <= CInt(toKikan)) Then
                    idoFrom = IdoLow
                Else
                    idoFrom = ""
                End If
            Else
                If (CInt(IdoLow) <= CInt(toKikan)) Then
                    idoFrom = fromKikan
                Else
                    idoFrom = ""
                End If
            End If

            '削除範囲の決定(終了)
            If (CInt(IdoUp) >= CInt(fromKikan)) Then
                If (CInt(IdoUp) <= CInt(toKikan)) Then
                    idoTo = IdoUp
                Else
                    If (CInt(IdoLow) <= CInt(toKikan)) Then
                        idoTo = toKikan
                    Else
                        idoTo = ""
                    End If
                End If
            Else
                idoTo = ""
            End If

            If idoFrom = "" Or idoTo = "" Then
                Return C_Cancel
            End If

            Dim yearFromIdo As String = idoFrom.Substring(0, 4)
            Dim monthFromIdo As String = CInt(idoFrom.Substring(4, 2)).ToString
            Dim yearToIdo As String = idoTo.Substring(0, 4)
            Dim monthToIdo As String = CInt(idoTo.Substring(4, 2)).ToString

            'センサス番号取得
            Dim CensusNo = ComUtil.Chosahyo.GetCensusNo(dcChosahyoCheck)

            Dim dtSame As DataTable
            Dim dtChosahyo As DataTable

            '登録対象データ取得
            Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)
            dcChosahyo = ComUtil.Tresa.GetSheetDataIdoTresaTarget(dtItem, xlSheets, CType(Me, ComObjectProcess), idoFrom, idoTo, Kessan)

            '同一期間(異動年月)調査票データの登録済みチェック
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try
                    dtSame = DAOOther.GetSameIdoKikanChosahyo(db, ChosaNen, CommonInfo.Chosakubun, yearFromIdo, monthFromIdo, yearToIdo, monthToIdo, Kessan, CensusNo)
                    db.BeginTrans()

                    If dtSame.Rows.Count <> 0 Then

                        'REV START 2022/05/20
                        '牛乳の場合
                        If ComConst.調査区分.牛乳生産費統計_個別 = CommonInfo.Chosakubun Or ComConst.調査区分.経営分析調査_牛乳生産費 = CommonInfo.Chosakubun Then

                            '通常分だけがある状態、成畜分だけがある状態、で処理を変化させるのでチェック
                            Dim dtKotaiShikibetsu As DataTable
                            Dim bolMessage As Boolean = False 'Falseならメッセージ出さない
                            '通常or成畜
                            If filetypr = 0 Then
                                '通常
                                dtKotaiShikibetsu = GetKotaiShikibetsu(db, CommonInfo.Chosakubun, ChosaNen, yearFromIdo, monthFromIdo, yearToIdo, monthToIdo, CensusNo, "Q11020101")
                                If 0 < dtKotaiShikibetsu.Rows.Count Then
                                    '通常分がある状態
                                    bolMessage = True
                                End If
                            Else
                                '成畜
                                dtKotaiShikibetsu = GetKotaiShikibetsu(db, CommonInfo.Chosakubun, ChosaNen, yearFromIdo, monthFromIdo, yearToIdo, monthToIdo, CensusNo, "Q11020701")
                                If 0 < dtKotaiShikibetsu.Rows.Count Then
                                    '成畜分がある状態
                                    bolMessage = True
                                End If
                            End If


                            If bolMessage Then
                                'メッセージを表示
                                msgRet = Message.ShowMsgBox(MessageID.MSG_Q_043, {yearFromIdo, monthFromIdo, yearToIdo, monthToIdo}, MsgBoxStyle.YesNo)
                                If msgRet = MsgBoxResult.Yes Then
                                    'メッセージで「はい」押下時は、対象データを削除
                                    For Each dr As DataRow In dtSame.Rows
                                        DAOOther.DeleteSameIdoKikanChosahyo(db, CommonInfo.Chosakubun, dr("調査年").ToString, dr("明細番号").ToString, CensusNo, filetypr)
                                    Next
                                Else
                                    db.RollBackTrans()
                                    Return C_Cancel
                                End If
                            End If
                        Else
                            '牛乳ではない場合
                            '登録されたデータが存在したら、メッセージを表示
                            msgRet = Message.ShowMsgBox(MessageID.MSG_Q_043, {yearFromIdo, monthFromIdo, yearToIdo, monthToIdo}, MsgBoxStyle.YesNo)
                            If msgRet = MsgBoxResult.Yes Then
                                'メッセージで「はい」押下時は、対象データを削除
                                For Each dr As DataRow In dtSame.Rows
                                    DAOOther.DeleteSameIdoKikanChosahyo(db, CommonInfo.Chosakubun, dr("調査年").ToString, dr("明細番号").ToString, CensusNo, filetypr)
                                Next
                            Else
                                db.RollBackTrans()
                                Return C_Cancel
                            End If
                        End If
                        'REV END 2022/05/20
                    End If

                    '電子調査票の登録済みチェック
                    dtChosahyo = DAOOther.GetImportChosahyo(db, CommonInfo.Chosakubun, ChosaNen, CensusNo)
                    If dtChosahyo.Rows.Count = 0 Then
                        Message.ShowMsgBox(MessageID.MSG_E_089, {CensusNo}, MsgBoxStyle.OkOnly)
                        db.RollBackTrans()
                        Return C_Cancel
                    End If

                    '成畜ファイル処理
                    If filetypr = 1 Then
                        '取込時使用する明細番号を取得する。
                        Dim MeisaiNo As String = ""
                        Dim dtMeisai = DAOOther.GetSeitikuMaxMeisaiNo(db, CommonInfo.Chosakubun, ChosaNen, CensusNo)
                        If dtMeisai.Rows.Count <> 0 Then
                            If dtMeisai.Rows(0)("明細番号最大値").ToString = "" Then
                                MeisaiNo = "0"
                            Else
                                MeisaiNo = dtMeisai.Rows(0)("明細番号最大値").ToString
                            End If
                        Else
                            db.RollBackTrans()
                            Return C_NG
                        End If

                        '主キー設定
                        Dim pKey As DAOChosahyo.PrimaryKey = New DAOChosahyo.PrimaryKey(
                                                                    ComUtil.Chosahyo.GetChosaNen(dcChosahyo),
                                                                    ComUtil.Chosahyo.GetCensusNo(dcChosahyo))

                        '調査票データ追加
                        'REV_001↓ 所有牛情報が更新されるため、職員の平均的な飼養頭数をクリアする
                        'DAOOther.InsertIdoTresaChosahyoTable(db, pKey, CommonInfo.Chosakubun, MeisaiNo, dcChosahyo)
                        Dim clearList = ComUtil.GetShiyotosuClearList()
                        DAOOther.InsertIdoTresaChosahyoTable(db, pKey, CommonInfo.Chosakubun, MeisaiNo, dcChosahyo, clearList)
                        'REV_001↑

                        'データの並び替え(異動月順)
                        Dim dtSort As DataTable
                        '並び替え用ダミーデータ登録
                        DAOOther.InsertSeitikuDummyIdoTresaChosahyoTable(db, CommonInfo.Chosakubun, ChosaNen, CensusNo)

                        '基データ削除
                        DAOOther.SeitikuDeleteChosahyo(db, CommonInfo.Chosakubun, ChosaNen, CensusNo)

                        'ダミーデータのソート
                        dtSort = DAOOther.GetSeitikuSortDummyChosahyo(db, CommonInfo.Chosakubun, ChosaNen, CensusNo, Kessan)

                        'ソートデータの追加
                        Dim dtTarget As DataTable
                        For Each row As DataRow In dtSort.Rows
                            dtTarget = DAOOther.GetSeitikuSortTaget(db, CommonInfo.Chosakubun, ChosaNen, CensusNo, row.Item("明細番号").ToString())
                            DAOOther.InsertSeitikuSortChosahyoTable(db, CommonInfo.Chosakubun, ChosaNen, dtTarget)
                        Next
                    Else

                        'REV START 2022/05/20
                        '通常ファイル処理（牛乳の場合）
                        If (ComConst.調査区分.牛乳生産費統計_個別 = CommonInfo.Chosakubun Or ComConst.調査区分.経営分析調査_牛乳生産費 = CommonInfo.Chosakubun) Then
                            '取込時使用する明細番号を取得する。
                            Dim MeisaiNo As String = ""
                            Dim dtMeisai = DAOOther.GetMaxMeisaiNo(db, CommonInfo.Chosakubun, ChosaNen, CensusNo)
                            If dtMeisai.Rows.Count <> 0 Then
                                If dtMeisai.Rows(0)("明細番号最大値").ToString = "" Then
                                    MeisaiNo = "0"
                                Else
                                    MeisaiNo = dtMeisai.Rows(0)("明細番号最大値").ToString
                                End If
                            Else
                                db.RollBackTrans()
                                Return C_NG
                            End If

                            '主キー設定
                            Dim pKey As DAOChosahyo.PrimaryKey = New DAOChosahyo.PrimaryKey(
                                                                    ComUtil.Chosahyo.GetChosaNen(dcChosahyo),
                                                                    ComUtil.Chosahyo.GetCensusNo(dcChosahyo))

                            '調査票データ追加
                            'REV_001↓ 所有牛情報が更新されるため、職員の平均的な飼養頭数をクリアする
                            'DAOOther.InsertIdoTresaChosahyoTable(db, pKey, CommonInfo.Chosakubun, MeisaiNo, dcChosahyo)
                            Dim clearList = ComUtil.GetShiyotosuClearList()
                            DAOOther.InsertIdoTresaChosahyoTable(db, pKey, CommonInfo.Chosakubun, MeisaiNo, dcChosahyo, clearList)
                            'REV_001↑

                            'データの並び替え(異動月順)
                            Dim dtSort As DataTable
                            '並び替え用ダミーデータ登録
                            DAOOther.InsertDummyIdoTresaChosahyoTable(db, CommonInfo.Chosakubun, ChosaNen, CensusNo)

                            '基データ削除
                            DAOOther.DeleteChosahyo(db, CommonInfo.Chosakubun, ChosaNen, CensusNo)

                            'ダミーデータのソート
                            dtSort = DAOOther.GetSortDummyChosahyo(db, CommonInfo.Chosakubun, ChosaNen, CensusNo, Kessan)
                            '最後尾に持ってくるための保持用
                            Dim lastHoji As DataTable 'dtSortの値をベースに同じものを作ってここから減らす
                            lastHoji = DAOOther.GetSortDummyChosahyo(db, CommonInfo.Chosakubun, ChosaNen, CensusNo, Kessan)

                            'ソートデータの追加
                            Dim dtTarget As DataTable
                            Dim meisai As String = "1"
                            Dim index As Integer = 0  '最後尾カウント用

                            For Each row As DataRow In dtSort.Rows
                                If String.IsNullOrEmpty(row.Item("Q11020101").ToString()) Then
                                    '使わなかった時だけカウントして次行を参照（残す行は最後尾に）
                                    index = index + 1
                                Else
                                    dtTarget = DAOOther.GetSortTaget(db, CommonInfo.Chosakubun, ChosaNen, CensusNo, row.Item("明細番号").ToString())
                                    DAOOther.InsertSortChosahyoTable(db, CommonInfo.Chosakubun, ChosaNen, meisai, dtTarget)
                                    meisai = (CInt(meisai) + 1).ToString
                                    '使ったものは消す
                                    lastHoji.Rows.RemoveAt(index)
                                End If
                            Next
                            '最後尾実施（成畜データ）
                            For Each row As DataRow In lastHoji.Rows
                                dtTarget = DAOOther.GetSortTaget(db, CommonInfo.Chosakubun, ChosaNen, CensusNo, row.Item("明細番号").ToString())
                                DAOOther.InsertSortChosahyoTable(db, CommonInfo.Chosakubun, ChosaNen, meisai, dtTarget)
                                meisai = (CInt(meisai) + 1).ToString
                            Next

                        Else

                            '牛乳以外の通常の場合
                            '取込時使用する明細番号を取得する。
                            Dim MeisaiNo As String = ""
                            Dim dtMeisai = DAOOther.GetMaxMeisaiNo(db, CommonInfo.Chosakubun, ChosaNen, CensusNo)
                            If dtMeisai.Rows.Count <> 0 Then
                                If dtMeisai.Rows(0)("明細番号最大値").ToString = "" Then
                                    MeisaiNo = "0"
                                Else
                                    MeisaiNo = dtMeisai.Rows(0)("明細番号最大値").ToString
                                End If
                            Else
                                db.RollBackTrans()
                                Return C_NG
                            End If

                            '主キー設定
                            Dim pKey As DAOChosahyo.PrimaryKey = New DAOChosahyo.PrimaryKey(
                                                                    ComUtil.Chosahyo.GetChosaNen(dcChosahyo),
                                                                    ComUtil.Chosahyo.GetCensusNo(dcChosahyo))

                            '調査票データ追加
                            'REV_001↓ 所有牛情報が更新されるため、職員の平均的な飼養頭数をクリアする
                            'DAOOther.InsertIdoTresaChosahyoTable(db, pKey, CommonInfo.Chosakubun, MeisaiNo, dcChosahyo)
                            Dim clearList = ComUtil.GetShiyotosuClearList()
                            DAOOther.InsertIdoTresaChosahyoTable(db, pKey, CommonInfo.Chosakubun, MeisaiNo, dcChosahyo, clearList)
                            'REV_001↑

                            'データの並び替え(異動月順)
                            Dim dtSort As DataTable
                            '並び替え用ダミーデータ登録
                            DAOOther.InsertDummyIdoTresaChosahyoTable(db, CommonInfo.Chosakubun, ChosaNen, CensusNo)

                            '基データ削除
                            DAOOther.DeleteChosahyo(db, CommonInfo.Chosakubun, ChosaNen, CensusNo)

                            'ダミーデータのソート
                            dtSort = DAOOther.GetSortDummyChosahyo(db, CommonInfo.Chosakubun, ChosaNen, CensusNo, Kessan)

                            'ソートデータの追加
                            Dim dtTarget As DataTable
                            Dim meisai As String = "1"
                            For Each row As DataRow In dtSort.Rows
                                dtTarget = DAOOther.GetSortTaget(db, CommonInfo.Chosakubun, ChosaNen, CensusNo, row.Item("明細番号").ToString())

                                DAOOther.InsertSortChosahyoTable(db, CommonInfo.Chosakubun, ChosaNen, meisai, dtTarget)

                                meisai = (CInt(meisai) + 1).ToString
                            Next
                            'REV END 2022/05/20

                        End If


                    End If
                    'ダミーデータ削除
                    DAOOther.DeleteChosahyo(db, CommonInfo.Chosakubun, (CInt(ChosaNen) + 100000).ToString, CensusNo)
                    db.CommitTrans()
                    ret = C_OK
                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using
            Return ret
        End Function

        ''' <summary>
        ''' Excelシート存在チェック
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function CheckExcelSheetExist(dt As DataTable, xlSheets As Excel.Sheets) As Boolean
            Dim ret As Boolean = True

            'シート名取得（項目マスタ）
            Dim dtSheetNames As New List(Of String)
            Dim sheetList = ComConst.牛トレサデータ.異動ファイルシート(CommonInfo.Chosakubun)

            For Each sheet In sheetList
                dtSheetNames.Add(sheet)
            Next

            'シート名取得（シートオブジェクト）
            Dim xlSheetNames As New List(Of String)

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                For Each xlSheet In xlSheets
                    If xlSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible Then
                        xlSheetNames.Add(xlSheet.Name)
                        ReleaseComObject(xlSheet)
                    End If
                Next
            Finally
                ReleaseComObject(xlSheet)
            End Try

            'シート名比較
            If xlSheetNames.Count <> dtSheetNames.Count Then
                ret = False
            Else
                For Each sheetName As String In dtSheetNames
                    If Not xlSheetNames.Contains(sheetName) Then
                        ret = False
                        Exit For
                    End If
                Next
            End If
            Return ret
        End Function

        ''' <summary>
        ''' Excel内データのエラーチェック
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="dcChosahyo"></param>
        ''' <param name="details"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function CheckError(xlSheets As Excel.Sheets, dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目), ByRef details As List(Of String)) As Boolean
            Dim ret As Boolean = True

            Const max As Integer = ComConst.ERR_MESSAGE_MAX

            Dim msg As String
            Dim cnt As Integer = 0

            '1）操作しているユーザが専門調査員の場合、その専門調査員が扱えるセンサス番号かどうか。
            msg = "{0}件目：操作可能なセンサス番号ではありません。"
            If CommonInfo.SenmonChosain Then
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    If Not DAOOther.CheckSenmonChosainKyakutaiExist(db, CommonInfo.UserId, ComUtil.Chosahyo.GetCensusNo(dcChosahyo)) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg, cnt.ToString.PadLeft(2)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End Using
            End If

            If Not ret Then Return ret

            '2）必須項目が入力されているか。
            msg = "{0}件目：{1}の入力がございません。"

            '調査年
            If String.IsNullOrEmpty(ComUtil.Chosahyo.GetChosaNen(dcChosahyo)) Then
                cnt = cnt + 1
                details.Add(String.Format(msg, cnt.ToString.PadLeft(2), "調査年"))
                ret = False
                If cnt = max Then Return ret
            End If

            Select Case CommonInfo.Kubun2
                Case ComConst.区分２.営農類型別経営統計
                    '営農類型
                    If String.IsNullOrEmpty(ComUtil.Chosahyo.GetEinouRuike(dcChosahyo)) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg, cnt.ToString.PadLeft(2), "営農類型"))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                Case ComConst.区分２.農産物生産費
                    '対象品目
                    If String.IsNullOrEmpty(ComUtil.Chosahyo.GetTaishoHinmoku(dcChosahyo)) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg, cnt.ToString.PadLeft(2), "対象品目"))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '経営種類
                    If String.IsNullOrEmpty(ComUtil.Chosahyo.GetKeieiShurui(dcChosahyo)) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg, cnt.ToString.PadLeft(2), "経営種類"))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                Case ComConst.区分２.畜産物生産費
                    If CommonInfo.Chosakubun = ComConst.調査区分.乳用雄育成牛生産費統計_個別 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.交雑種育成牛生産費統計_個別 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.去勢若齢肥育牛生産費統計_個別 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.乳用雄肥育牛生産費統計_個別 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.交雑種肥育牛生産費統計_個別 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_乳用雄育成牛生産費 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_交雑種育成牛生産費 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_交雑種肥育牛生産費 Then

                        '生産費区分
                        If String.IsNullOrEmpty(ComUtil.Chosahyo.GetSeisanhiKubun(dcChosahyo)) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg, cnt.ToString.PadLeft(2), "生産費区分"))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If
            End Select

            msg = "{0}件目：センサス番号の各項目について入力がない箇所がございます。"
            If String.IsNullOrEmpty(ComUtil.Chosahyo.GetTodofuken(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetShikuchoson(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetKyuShikuchoson(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetNogyoShuraku(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetChosaku(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetKyakutaiNo(dcChosahyo)) Then
                cnt = cnt + 1
                details.Add(String.Format(msg, cnt.ToString.PadLeft(2)))
                ret = False
                If cnt = max Then Return ret
            End If

            If Not ret Then Return ret

            '3）調査情報入力（実査設置拠点）画面で設定した、調査区分の電子調査票であるか。
            msg = "{0}件目：「調査情報入力画面」で設定した調査区分と、取込む電子調査票の調査区分が異なっています。"

            Dim sheet As String = String.Empty
            Dim adress As String = String.Empty
            Dim str As String = String.Empty

            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                sheet = "00_表紙"
                adress = "C20"
                str = "（個人経営体用）"
            End If
            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                sheet = "00_表紙"
                adress = "C20"
                str = "（法人経営体用）"
            End If
            If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then
                sheet = "指標部入力"
                adress = "E12"
                str = CommonInfo.ChosakubunName
            End If
            If CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費 Then
                sheet = "表紙"
                adress = "BJ10"
                str = CommonInfo.ChosakubunName
            End If

            Dim xlSheet As Excel.Worksheet = Nothing
            Try
                xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

                Dim rng As Excel.Range = Nothing
                Try
                    rng = xlSheet.Range(adress)
                    Dim val As String = If(rng.Value Is Nothing, String.Empty, rng.Value.ToString)
                    If Not val = str Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg, cnt.ToString.PadLeft(2)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                Finally
                    ReleaseComObject(rng)
                End Try
            Finally
                ReleaseComObject(xlSheet)
            End Try

            If Not ret Then Return ret

            '4）農業経営統計調査システムにログインしている実査設置拠点を管轄している都道府県と、電子調査票の都道府県（センサス番号内）が一致しているか。
            msg = "{0}件目：自管轄の都道府県と、取込む電子調査票の都道府県が異なっています。"
            If Not CommonInfo.Jimusyo = ComUtil.Chosahyo.ConvJimusyoNo(dcChosahyo).PadLeft(2, "0"c) Then
                cnt = cnt + 1
                details.Add(String.Format(msg, cnt.ToString.PadLeft(2)))
                ret = False
                If cnt = max Then Return ret
            End If

            If Not ret Then Return ret

            '5）電子調査票の各項目が、調査票項目マスタの型と一致しているか。
            msg = "{0}件目：シート名：{1}　{2}の型が一致しません。"
            For Each kv As KeyValuePair(Of String, DAOChosahyo.調査票項目) In dcChosahyo
                If Not String.IsNullOrEmpty(kv.Value.値) Then
                    If kv.Value.型区分 = ComConst.型区分.数値型 Then
                        Dim val As Decimal
                        If Not Decimal.TryParse(kv.Value.値, val) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg, cnt.ToString.PadLeft(2), kv.Value.シート名, kv.Key))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If
                End If
            Next

            '6）電子調査票の各項目が、データベースの桁数に収まっているか。
            msg = "{0}件目：シート名：{1}　{2}の桁数がデータベースの桁数を超えています。"
            For Each kv As KeyValuePair(Of String, DAOChosahyo.調査票項目) In dcChosahyo
                If Not String.IsNullOrEmpty(kv.Value.値) Then
                    If kv.Value.型区分 = ComConst.型区分.数値型 Then
                        Dim val As Decimal
                        If Decimal.TryParse(kv.Value.値, val) Then
                            If kv.Value.有効桁数 > 0 Then
                                Dim pattern As String
                                If kv.Value.小数点以下桁数 > 0 Then
                                    pattern = "^-?[0-9]{1," & kv.Value.有効桁数 - kv.Value.小数点以下桁数 & "}(\.[0-9]{1," & kv.Value.小数点以下桁数 & "})?$"
                                Else
                                    pattern = "^-?[0-9]{1," & kv.Value.有効桁数 & "}$"
                                End If
                                If Not Regex.IsMatch(kv.Value.値, pattern) Then
                                    cnt = cnt + 1
                                    details.Add(String.Format(msg, cnt.ToString.PadLeft(2), kv.Value.シート名, kv.Key))
                                    ret = False
                                    If cnt = max Then Return ret
                                End If
                            End If
                        End If
                    End If
                End If
            Next

            Return ret

        End Function

        ''' <summary>
        ''' 可変データより特定列のデータを取得
        ''' </summary>
        ''' <param name="db"></param>
        ''' <param name="chosaKubun">調査区分</param>
        ''' <param name="strYear">調査年</param>
        ''' <param name="strYearFrom">調査開始年</param>
        ''' <param name="strMonthFrom">調査開始月</param>
        ''' <param name="strYearTo">調査終了年</param>
        ''' <param name="strMonthTo">調査終了月</param>
        ''' <param name="strCensusNo">センサス番号</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function GetKotaiShikibetsu(db As DBAccess, chosaKubun As String, strYear As String, strYearFrom As String, strMonthFrom As String, strYearTo As String, strMonthTo As String, strCensusNo As String, strKoban As String) As DataTable

            Dim ret As DataTable
            Dim sb As System.Text.StringBuilder = Nothing
            Dim para As List(Of DBAccess.Parameter) = Nothing

            'EXCELとSQLではシリアル値が2日ずれるため、3から開始する
            Dim dtfrom As New DateTime(CInt(strYearFrom), CInt(strMonthFrom), 3)
            Dim dtto As DateTime
            Dim tmpYearTo As String
            Dim tmpMonthTo As String
            If CInt(strMonthTo) = 12 Then
                tmpYearTo = CStr(CInt(strYearTo) + 1)
                tmpMonthTo = "1"
            Else
                tmpYearTo = strYearTo
                tmpMonthTo = CStr(CInt(strMonthTo) + 1)
            End If
            dtto = New DateTime(CInt(tmpYearTo), CInt(tmpMonthTo), 3)

            Try
                Dim query = From val In ComConst.調査票.テーブル名称(chosaKubun) Where val.Contains("＿可変")
                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT A.* FROM ")
                    .AppendLine(" ( ")
                    .AppendLine("SELECT * FROM " & query(0).ToString)
                    .AppendLine(" WHERE ")
                    .AppendLine(" 調査年 = @調査年 ")
                    .AppendLine(" AND センサス番号 = @センサス番号 ")
                    .AppendLine(" AND 項目番号 = @項目番号 ")
                    .AppendLine(" ) ")
                    .AppendLine(" As A ")
                    .AppendLine(" INNER JOIN ")
                    .AppendLine(" ( ")
                    .AppendLine("SELECT * FROM " & query(0).ToString)
                    .AppendLine(" WHERE ")
                    .AppendLine(" 調査年 = @調査年 ")
                    .AppendLine(" AND センサス番号 = @センサス番号 ")
                    .AppendLine(" AND 項目番号 = @異動年月項目番号 ")
                    .AppendLine("  AND (値 >= CONVERT(varchar, (CONVERT(float, CONVERT(datetime,@開始年月)))) ")
                    .AppendLine("  AND 値 < CONVERT(varchar,CONVERT(float, CONVERT(datetime,@終了年月))))")
                    .AppendLine(" ) ")
                    .AppendLine(" As B ON A.明細番号 = B.明細番号 ")
                End With
                para.Add(db.CreateParameter("@調査年", SqlDbType.VarChar, strYear))
                para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, strCensusNo))
                para.Add(db.CreateParameter("@項目番号", SqlDbType.VarChar, strKoban))
                para.Add(db.CreateParameter("@異動年月項目番号", SqlDbType.VarChar, ComConst.牛トレサデータ.異動年月_項目番号(chosaKubun)))
                para.Add(db.CreateParameter("@開始年月", SqlDbType.VarChar, dtfrom.ToString))
                para.Add(db.CreateParameter("@終了年月", SqlDbType.VarChar, dtto.ToString))

                ret = db.GetDataTable(sb.ToString, para)
            Catch ex As Exception
                Throw ex
            End Try

            Return ret
        End Function

    End Class

End Class
