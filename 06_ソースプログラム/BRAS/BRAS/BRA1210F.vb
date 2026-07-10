Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions

''' <summary>
''' 電子調査票取込画面
''' </summary>
''' <remarks></remarks>
''' 
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.12.07 |TSP)                | フェーズ3 追加要件No4修正
'//  REV_002   | 2021.04.14 |TSP)                | 稼働障害修正
'//  REV_003   | 2021.12.07 |日本コンピュータシステム  | 要件No1-③
'//  REV_004   | 2025.09.11 |GCU                 | 要件No9 調査票取込に調査対象経営体と担当名称を追加
'//            |            |                    |
'//*************************************************************************************************
Public Class BRA1210F

    Private Sub BRA1210F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Dim year As String = txtYear.Text

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(year, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            'フォルダパス取得
            Dim folderPath As String = ComUtil.GetFolderPath(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainInPath, IniFileInfo.ExcelInPath))

            If folderPath.Equals(String.Empty) Then
                Exit Sub
            End If

            'ファイル拡張子
            Dim searchPattern As String() = {"*.xlsx", "*.xlsm"}

            Dim files As IEnumerable(Of String) = {}
            For Each pattern As String In searchPattern
                files = files.Concat(System.IO.Directory.EnumerateFiles(folderPath, pattern, System.IO.SearchOption.TopDirectoryOnly))
            Next

            'ファイル存在チェック
            If files.Count = 0 Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_007, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim fileNames As List(Of String)() = Nothing
            '調査票取込クラス生成
            Using ImportChosahyo = New ImportChosahyo(year)
                '処理実行
                fileNames = ImportChosahyo.Execute(files, Me)
            End Using

            Dim msgPara() As String = {String.Join(vbCrLf, fileNames(0)), String.Join(vbCrLf, fileNames(1))}

            'メッセージフォーム表示
            Message.ShowMsgForm(Me, MessageID.MSG_I_004, msgPara)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="year"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(year As String, ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        '未入力チェック
        If year.Equals(String.Empty) Then
            msgId = MessageID.MSG_E_004
            Return ret
        End If

        '半角数字チェック
        If Not Regex.IsMatch(year, "^[0-9]+$") Then
            msgId = MessageID.MSG_E_022
            Return ret
        End If

        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' 調査票取込クラス
    ''' </summary>
    ''' <remarks></remarks>
    Private Class ImportChosahyo
        Inherits ExcelProcess

        ''' <summary>調査年</summary>
        Private _chosaNen As String
        ''' <summary>拠点キー</summary>
        Private _kKey As DAOChosahyo.KotenKey

        ''' <summary>進捗ダイアログ</summary>
        Private ProgressDialog As New ProgressDialog()

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()

        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="chosaNen"></param>
        ''' <remarks></remarks>
        Public Sub New(chosaNen As String)
            MyBase.New(True)

            _chosaNen = chosaNen
            _kKey = New DAOChosahyo.KotenKey(CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
        End Sub

        ''' <summary>
        ''' 処理実行
        ''' </summary>
        ''' <param name="files"></param>
        ''' <param name="form"></param>
        ''' <remarks></remarks>
        Public Function Execute(files As IEnumerable(Of String), myParent As Form) As List(Of String)()
            Dim ret As List(Of String)() = {Nothing, Nothing}
            Dim ok As New List(Of String)
            Dim ng As New List(Of String)

            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Sheets = Nothing

            '調査票項目マスタ取得
            Dim dtItem As DataTable
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))



                dtItem = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)
            End Using

            'Excelアプリ無効
            Me.DisableExcelApp()

            Try
                '進捗ダイアログを表示する
                ProgressDialog.Maximum = files.Count * 2
                ProgressDialog.Show(myParent)

                For Each filePath As String In files
                    Try
                        '進捗を進める
                        ProgressDialog.AddValue = 1

                        'Workbookを開く
                        xlBook = xlBooks.Open(filePath)

                        Try
                            xlSheets = xlBook.Worksheets

                            Dim msgId As String = String.Empty
                            Dim msgPara As String() = {}

                            '取込処理
                            If Me.Import(xlSheets, dtItem, msgId, msgPara) Then
                                ok.Add(IO.Path.GetFileName(filePath))
                            Else
                                ng.Add(IO.Path.GetFileName(filePath))
                                Select Case msgId
                                    Case MessageID.MSG_E_008
                                        If msgPara.Length > 0 Then
                                            ProgressDialog.ShowMsgForm(msgId, {IO.Path.GetFileName(filePath)}.Concat(msgPara).ToArray)
                                        End If
                                    Case MessageID.MSG_E_033
                                        ProgressDialog.ShowMsgBox(msgId, {filePath}, MsgBoxStyle.OkOnly)
                                        'REV001_ADD START---
                                    Case MessageID.MSG_E_090
                                        ProgressDialog.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                                        'REV001_ADD END---
                                End Select
                            End If

                            '進捗を進める
                            ProgressDialog.AddValue = 1
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
                Next
            Finally
                If Not ProgressDialog Is Nothing Then
                    '進捗ダイアログを閉じる
                    ProgressDialog.endDispose()
                    ProgressDialog = Nothing
                End If
            End Try

            'Excelアプリ有効
            Me.EnableExcelApp()

            ret(0) = ok
            ret(1) = ng

            Return ret
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
        Private Function Import(xlSheets As Excel.Sheets, dtItem As DataTable, ByRef msgId As String, ByRef msgPara As String()) As Boolean
            Dim ret As Boolean = False

            'Excelシート存在チェック
            If Not Me.CheckExcelSheetExist(dtItem, xlSheets) Then
                'メッセージ設定
                msgId = MessageID.MSG_E_033
                Return ret
            End If

            'REV001_ADD START---
            'Excelファイルテンプレートチェック
            If Not Me.CheckExcelFileTemplates(xlSheets) Then
                'メッセージ設定
                msgId = MessageID.MSG_E_090
                Return ret
            End If
            'REV001_ADD END----

            Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)

            '調査票シートデータ取得
            'REV-003 START-------------------
            dcChosahyo = ComUtil.Chosahyo.GetSheetData(dtItem, xlSheets, CType(Me, ComObjectProcess), _chosaNen)
            'REV-003 END-------------------

            'エラーチェック
            Dim details As New List(Of String)
            If Not Me.CheckError(xlSheets, dcChosahyo, details) Then
                'メッセージ設定
                msgId = MessageID.MSG_E_008
                msgPara = {String.Join(vbCrLf, details)}
                Return ret
            End If

            '主キー設定
            Dim pKey As DAOChosahyo.PrimaryKey = New DAOChosahyo.PrimaryKey( _
                                                        ComUtil.Chosahyo.GetChosaNen(dcChosahyo), _
                                                        ComUtil.Chosahyo.GetCensusNo(dcChosahyo))

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                '調査票データ実査設置拠点取得
                Dim koten As DAOChosahyo.KotenKey
                koten = DAOChosahyo.GetChosahyoKoten(db, pKey)

                '実査設置拠点存在チェック
                If koten Is Nothing Then
                    Try
                        db.BeginTrans()

                        '>>>2022/01/27
                        '調査票データ追加
                        DAOChosahyo.InsertChosahyoTable(db, pKey, _kKey, dcChosahyo, Nothing, Nothing)
                        'DAOChosahyo.InsertChosahyoTable(db, pKey, _kKey, dcChosahyo)
                        '<<<2022/01/27
                        db.CommitTrans()

                        ret = True
                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try
                Else
                    If Not (CInt(koten.kyoku) = CInt(_kKey.kyoku) And CInt(koten.jimusho) = CInt(_kKey.jimusho) And CInt(koten.kyoten) = CInt(_kKey.kyoten)) Then
                        'エラーメッセージ
                        ProgressDialog.ShowMsgBox(MessageID.MSG_E_009, {pKey.chosaNen, MasterDao.GetJimusyoName(koten.jimusho), MasterDao.GetCenterName(koten.jimusho, koten.kyoten), pKey.censusNo}, MsgBoxStyle.OkOnly)
                    Else
                        '確認メッセージ
                        If ProgressDialog.ShowMsgBox(MessageID.MSG_Q_005, {pKey.censusNo}, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                            '個別結果表存在チェック
                            Dim pKobet As DAOKobetsuKekkahyo.PrimaryKey = New DAOKobetsuKekkahyo.PrimaryKey(pKey.chosaNen, pKey.censusNo)
                            Dim kKobet As DAOKobetsuKekkahyo.KyotenKey = New DAOKobetsuKekkahyo.KyotenKey(_kKey.kyoku, _kKey.jimusho, _kKey.kyoten)
                            Dim bln As Boolean = DAOKobetsuKekkahyo.CheckExist(db, pKobet, kKobet)

                            If bln Then
                                '確認メッセージ
                                If ProgressDialog.ShowMsgBox(MessageID.MSG_Q_015, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                                    Return ret
                                End If
                            End If

                            Try
                                db.BeginTrans()

                                If bln Then
                                    '個別結果表データ削除
                                    DAOKobetsuKekkahyo.DeleteTable(db, pKobet, kKobet)
                                End If

                                'REV_004
                                Dim chosaSoshiki As String = String.Empty
                                Dim tantoMeisho As String = String.Empty
                                '調査対象経営体と担当名称取得
                                Dim result As DataTable = DAOChosahyo.GetKeieitaiAndTantoName(db, pKey, _kKey)
                                If result IsNot Nothing AndAlso result.Rows.Count > 0 Then
                                    chosaSoshiki = result.Rows(0)("Q00000701").ToString()
                                    tantoMeisho = result.Rows(0)("Q00000801").ToString()
                                End If
                                '調査票データ削除
                                DAOChosahyo.DeleteChosahyoTable(db, pKey, _kKey)
                                '>>>2022/01/27
                                '調査票データ追加
                                DAOChosahyo.InsertChosahyoTable(db, pKey, _kKey, dcChosahyo, chosaSoshiki, tantoMeisho)
                                'DAOChosahyo.InsertChosahyoTable(db, pKey, _kKey, dcChosahyo)
                                '<<<2022/01/27


                                db.CommitTrans()

                                ret = True
                            Catch ex As Exception
                                db.RollBackTrans()
                                Throw ex
                            End Try
                        End If
                    End If
                End If
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
            Dim dtSheetNames As List(Of String) = (From dr In dt Group dr By dr!シート名 Into Group Select CType(シート名, String)).ToList

            'シート名取得（シートオブジェクト）
            Dim xlSheetNames As New List(Of String)

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                For Each xlSheet In xlSheets
                    xlSheetNames.Add(xlSheet.Name)
                    ReleaseComObject(xlSheet)
                Next
            Finally
                ReleaseComObject(xlSheet)
            End Try

            'シート名比較
            For Each sheetName As String In dtSheetNames
                If Not xlSheetNames.Contains(sheetName) Then
                    ret = False
                    Exit For
                End If
            Next

            Return ret
        End Function

        ''' <summary>
        ''' エラーチェック
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

            '3）当画面の調査年（産）と、電子調査票の調査年が一致しているか。
            msg = "{0}件目：画面の調査年（産）と、取込む電子調査票の調査年が異なっています。"
            If Not _chosaNen = ComUtil.Chosahyo.GetChosaNen(dcChosahyo) Then
                cnt = cnt + 1
                details.Add(String.Format(msg, cnt.ToString.PadLeft(2)))
                ret = False
                If cnt = max Then Return ret
            End If

            If Not ret Then Return ret

            '4）調査情報入力（実査設置拠点）画面で設定した、調査区分の電子調査票であるか。
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
            If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 And ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2015 Then
                sheet = "指標部入力"
                adress = "E12"
                str = CommonInfo.ChosakubunName
            ElseIf CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 And ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2020 Then
                sheet = "指標部入力"
                adress = "E17"
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

            '5）農業経営統計調査システムにログインしている実査設置拠点を管轄している都道府県と、電子調査票の都道府県（センサス番号内）が一致しているか。
            msg = "{0}件目：自管轄の都道府県と、取込む電子調査票の都道府県が異なっています。"
            If Not CommonInfo.Jimusyo = ComUtil.Chosahyo.ConvJimusyoNo(dcChosahyo).PadLeft(2, "0"c) Then
                cnt = cnt + 1
                details.Add(String.Format(msg, cnt.ToString.PadLeft(2)))
                ret = False
                If cnt = max Then Return ret
            End If

            If Not ret Then Return ret

            '6）電子調査票の各項目が、調査票項目マスタの型と一致しているか。
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

            '7）電子調査票の各項目が、データベースの桁数に収まっているか。
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

        'REV001_ADD START---
        ''' <summary>
        ''' 電子調査票テンプレートチェック
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <returns></returns>
        Private Function CheckExcelFileTemplates(xlSheets As Excel.Sheets) As Boolean
            Dim hantei As ComConst.調査票.取り込みファイル判定.判定内容
            Dim ret As Boolean = True
            Dim xlSheet As Excel.Worksheet = Nothing
            Dim rng As Excel.Range = Nothing

            'REV002_MOD START---
            'hantei = ComConst.調査票.取り込みファイル判定.判定内容一覧(CommonInfo.Chosakubun)
            If ComConst.調査票.取り込みファイル判定.判定内容一覧.ContainsKey(CommonInfo.Chosakubun) Then
                hantei = ComConst.調査票.取り込みファイル判定.判定内容一覧(CommonInfo.Chosakubun)
            Else
                Return ret
            End If
            'REV002_MOD END---

            Try
                xlSheet = DirectCast(xlSheets.Item(hantei.シート名), Excel.Worksheet)
                Try
                    rng = xlSheet.Range(hantei.セル番号)

                    Dim strValue As String = CStr(rng.Value)

                    strValue = strValue.Replace(vbLf, "")

                    If strValue.Equals(hantei.セル内容) Then
                        ret = False
                    Else
                        ret = True
                    End If
                Finally
                    ReleaseComObject(rng)
                End Try

            Finally
                ReleaseComObject(xlSheet)
            End Try

            Return ret
        End Function
        'REV001_ADD END---
    End Class
End Class
