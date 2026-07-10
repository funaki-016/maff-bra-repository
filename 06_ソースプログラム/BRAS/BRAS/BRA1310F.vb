'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_002   | 2023.10.30 |大興電子通信        | 変更要件No.6
'//  REV_003   | 2025.09.17 |GCU                 | 変更要件No.1④負担割合関連修正
'//            |            |                    |
'//*************************************************************************************************
Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions

''' <summary>
''' 中間集計表取込画面
''' </summary>
''' <remarks></remarks>
Public Class BRA1310F

    Private Sub BRA1310F_Load(sender As Object, e As EventArgs) Handles Me.Load
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

            'ファイルパス取得
            Dim filePath As String = GetFilePath(Of OpenFileDialog)(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainInPath, IniFileInfo.ExcelInPath))

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            Dim fileNames As List(Of String)() = Nothing
            '調査票取込クラス生成
            Using ImportChosahyo = New ImportChosahyo(year)
                '処理実行
                fileNames = ImportChosahyo.Execute(filePath, Me)
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
    ''' ファイルパス取得
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="frm"></param>
    ''' <param name="selectedPath"></param>
    ''' <param name="fileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFilePath(Of T As {FileDialog, New})(frm As Form, selectedPath As String, Optional fileName As String = Nothing) As String
        Dim ret As String = String.Empty
        Dim fdg As New T

        fdg.InitialDirectory = selectedPath
        fdg.Filter = "Excelファイル (*.xlsm)|*.xlsm"
        fdg.FilterIndex = 1
        fdg.RestoreDirectory = True
        If Not fileName Is Nothing Then
            fdg.FileName = fileName
        End If
        If TypeOf fdg Is SaveFileDialog Then
            Dim tp As Type = fdg.GetType()
            Dim pInf As Reflection.PropertyInfo = tp.GetProperty("OverwritePrompt")
            pInf.SetValue(fdg, False)
        End If

        If Not System.IO.Directory.Exists(selectedPath) Then
            System.IO.Directory.CreateDirectory(selectedPath)
        End If

        If fdg.ShowDialog(frm) = DialogResult.OK Then
            ret = fdg.FileName
        End If

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
        ''' <summary>主キー</summary>
        Private _pKey As DAOChosahyo.PrimaryKey

        ''' <summary>進捗ダイアログ</summary>
        Private ProgressDialog As New ProgressDialog()

        ''' <summary>非可変項目集計用</summary>
        Private Class HutanWariaiSum
            Public sum As Decimal
            Public cnt As Integer
        End Class

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
        Public Function Execute(files As String, myParent As Form) As List(Of String)()
            Dim ret As List(Of String)() = {Nothing, Nothing}
            Dim ok As New List(Of String)
            Dim ng As New List(Of String)

            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Sheets = Nothing

            '調査票、中間集計表項目マスタ取得
            Dim dtItem As DataTable
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                dtItem = DAOOther.GetTyukanSyukeiItemMaster(db, CommonInfo.Chosakubun, ComConst.中間集計表.登録削除区分.enm.登録, _chosaNen)
            End Using

            '調査区分に対応した生産費区分名の設定
            Select Case CommonInfo.Chosakubun
                Case ComConst.調査区分.なたね生産費統計_個別, ComConst.調査区分.経営分析調査_なたね生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.なたね

                Case ComConst.調査区分.そば生産費統計_個別, ComConst.調査区分.経営分析調査_そば生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.そば

                Case ComConst.調査区分.二条大麦生産費統計_個別, ComConst.調査区分.経営分析調査_二条大麦生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.二条大麦

                Case ComConst.調査区分.六条大麦生産費統計_個別, ComConst.調査区分.経営分析調査_六条大麦生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.六条大麦

                Case ComConst.調査区分.はだか麦生産費統計_個別, ComConst.調査区分.経営分析調査_はだか麦生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.はだか麦

                Case ComConst.調査区分.米生産費統計_個別, ComConst.調査区分.米生産費統計_組織法人
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.米

                Case ComConst.調査区分.小麦生産費統計_個別, ComConst.調査区分.小麦生産費統計_組織法人
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.小麦

                Case ComConst.調査区分.大豆生産費統計_個別, ComConst.調査区分.大豆生産費統計_組織法人
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.大豆

                Case ComConst.調査区分.原料用かんしょ生産費統計_個別
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.原料用かんしょ

                Case ComConst.調査区分.原料用ばれいしょ生産費統計_個別, ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.原料用ばれいしょ

                Case ComConst.調査区分.さとうきび生産費統計_個別, ComConst.調査区分.経営分析調査_さとうきび生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.さとうきび

                Case ComConst.調査区分.てんさい生産費統計_個別, ComConst.調査区分.経営分析調査_てんさい生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.てんさい

                Case ComConst.調査区分.牛乳生産費統計_個別, ComConst.調査区分.経営分析調査_牛乳生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.牛乳

                Case ComConst.調査区分.子牛生産費統計_個別, ComConst.調査区分.経営分析調査_子牛生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.肉用子牛

                Case ComConst.調査区分.乳用雄育成牛生産費統計_個別, ComConst.調査区分.経営分析調査_乳用雄育成牛生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.乳用雄育成牛

                Case ComConst.調査区分.交雑種育成牛生産費統計_個別, ComConst.調査区分.経営分析調査_交雑種育成牛生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.交雑種育成牛

                Case ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.去勢若齢肥育牛

                Case ComConst.調査区分.乳用雄肥育牛生産費統計_個別, ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.乳用雄肥育牛

                Case ComConst.調査区分.交雑種肥育牛生産費統計_個別, ComConst.調査区分.経営分析調査_交雑種肥育牛生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.交雑種肥育牛

                Case ComConst.調査区分.肥育豚生産費統計_個別, ComConst.調査区分.経営分析調査_肥育豚生産費
                    ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME = ComConst.中間集計表.中間集計表生産費区分名称.肥育豚

            End Select

            'Excelアプリ無効
            Me.DisableExcelApp()

            Try
                '進捗ダイアログを表示する
                ProgressDialog.Maximum = 4
                ProgressDialog.Show(myParent)

                Try
                    '進捗を進める
                    ProgressDialog.AddValue = 1

                    'Workbookを開く
                    xlBook = xlBooks.Open(files)

                    Try
                        xlSheets = xlBook.Worksheets

                        Dim msgId As String = String.Empty
                        Dim msgPara As String() = {}

                        '取込処理
                        If Me.Import(xlSheets, dtItem, msgId, msgPara) Then
                            ok.Add(IO.Path.GetFileName(files))

                            '電子調査票の再保存
                            ProgressDialog.AddValue = 1
                            Saveagain(_pKey)
                            ProgressDialog.AddValue = 1
                        Else
                            ProgressDialog.AddValue = 1
                            ProgressDialog.AddValue = 1

                            ng.Add(IO.Path.GetFileName(files))
                            Select Case msgId
                                Case MessageID.MSG_E_008
                                    If msgPara.Length > 0 Then
                                        ProgressDialog.ShowMsgForm(msgId, {IO.Path.GetFileName(files)}.Concat(msgPara).ToArray)
                                    End If
                                Case MessageID.MSG_E_033
                                    ProgressDialog.ShowMsgBox(msgId, {files}, MsgBoxStyle.OkOnly)
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

            Dim dcTyukanSyukei As Dictionary(Of String, DAOChosahyo.中間集計表項目)

            '中間集計表シートデータ取得
            dcTyukanSyukei = GetSheetData(dtItem, xlSheets, CType(Me, ComObjectProcess), ComConst.中間集計表.登録削除区分.enm.登録)

            Dim dcChosahyo As New Dictionary(Of String, DAOChosahyo.調査票項目)
            For Each dc As KeyValuePair(Of String, DAOChosahyo.中間集計表項目) In dcTyukanSyukei
                dcChosahyo.Add(dc.Key, DirectCast(dc.Value, DAOChosahyo.調査票項目))
            Next

            'エラーチェック
            Dim details As New List(Of String)
            If Not Me.CheckError(xlSheets, dcChosahyo, details) Then
                'メッセージ設定
                msgId = MessageID.MSG_E_008
                msgPara = {String.Join(vbCrLf, details)}
                Return ret
            End If

            '主キー設定
            Dim pKey As DAOChosahyo.PrimaryKey = New DAOChosahyo.PrimaryKey(
                                                        ComUtil.Chosahyo.GetChosaNen(dcChosahyo),
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
                        '個別結果表存在チェック
                        Dim pKobet As DAOKobetsuKekkahyo.PrimaryKey = New DAOKobetsuKekkahyo.PrimaryKey(pKey.chosaNen, pKey.censusNo)
                        Dim kKobet As DAOKobetsuKekkahyo.KyotenKey = New DAOKobetsuKekkahyo.KyotenKey(_kKey.kyoku, _kKey.jimusho, _kKey.kyoten)
                        Dim bln As Boolean = DAOKobetsuKekkahyo.CheckExist(db, pKobet, kKobet)

                        If bln Then
                            '個別結果表存在時確認メッセージ
                            If ProgressDialog.ShowMsgBox(MessageID.MSG_Q_016, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                                '「はい」の場合の処理は、トランザクション内に記述
                                Return ret
                            End If
                        End If

                        '上書き確認メッセージ
                        If ProgressDialog.ShowMsgBox(MessageID.MSG_Q_055, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                            Return ret
                        End If

                        Try
                            db.BeginTrans()

                            If bln Then
                                '個別結果表データ削除
                                DAOKobetsuKekkahyo.DeleteTable(db, pKobet, kKobet)
                            End If

                            '調査票、中間集計表項目マスタ取得
                            Dim dtsakujoItem As DataTable
                            Using dbsakujo As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                                dtsakujoItem = DAOOther.GetTyukanSyukeiItemMaster(db, CommonInfo.Chosakubun, ComConst.中間集計表.登録削除区分.enm.削除, _chosaNen)
                            End Using
                            Dim dcTyukanSyukeisakujo As Dictionary(Of String, DAOChosahyo.中間集計表項目)

                            '中間集計表シートデータ取得
                            dcTyukanSyukeisakujo = GetSheetData(dtsakujoItem, xlSheets, CType(Me, ComObjectProcess), ComConst.中間集計表.登録削除区分.enm.削除)


                            '調査票データ削除
                            DAOChosahyo.DeleteUpdateChosahyoTable_TyukanSyukei(db, pKey, _kKey, dcTyukanSyukeisakujo)
                            ''調査票データ追加
                            DAOChosahyo.InsertUpdateChosahyoTable_TyukanSyukei(db, pKey, _kKey, dcTyukanSyukei)

                            db.CommitTrans()

                            ret = True
                        Catch ex As Exception
                            db.RollBackTrans()
                            Throw ex
                        End Try
                    End If
                End If

                _pKey = pKey
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
            dtSheetNames.Add("経営体概要")
            dtSheetNames.Add("前年対比（収入）")
            dtSheetNames.Add("前年対比（支出）")

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
            msg = "{0}件目：「調査情報入力画面」で設定した調査区分の品目と、取込む中間集計表の生産費１～３のいずれかが一致しません。"

            Dim sheet As String = String.Empty
            Dim adress As String = String.Empty
            Dim str As String = String.Empty

            sheet = "経営体概要"
            str = CommonInfo.Chosakubun

            Select Case CommonInfo.Chosakubun
                Case ComConst.調査区分.なたね生産費統計_個別, ComConst.調査区分.経営分析調査_なたね生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.なたね

                Case ComConst.調査区分.そば生産費統計_個別, ComConst.調査区分.経営分析調査_そば生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.そば

                Case ComConst.調査区分.二条大麦生産費統計_個別, ComConst.調査区分.経営分析調査_二条大麦生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.二条大麦

                Case ComConst.調査区分.六条大麦生産費統計_個別, ComConst.調査区分.経営分析調査_六条大麦生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.六条大麦

                Case ComConst.調査区分.はだか麦生産費統計_個別, ComConst.調査区分.経営分析調査_はだか麦生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.はだか麦

                Case ComConst.調査区分.米生産費統計_個別, ComConst.調査区分.米生産費統計_組織法人
                    str = ComConst.中間集計表.中間集計表生産費区分.米

                Case ComConst.調査区分.小麦生産費統計_個別, ComConst.調査区分.小麦生産費統計_組織法人
                    str = ComConst.中間集計表.中間集計表生産費区分.小麦

                Case ComConst.調査区分.大豆生産費統計_個別, ComConst.調査区分.大豆生産費統計_組織法人
                    str = ComConst.中間集計表.中間集計表生産費区分.大豆

                Case ComConst.調査区分.原料用かんしょ生産費統計_個別
                    str = ComConst.中間集計表.中間集計表生産費区分.原料用かんしょ

                Case ComConst.調査区分.原料用ばれいしょ生産費統計_個別, ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.原料用ばれいしょ

                Case ComConst.調査区分.さとうきび生産費統計_個別, ComConst.調査区分.経営分析調査_さとうきび生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.さとうきび

                Case ComConst.調査区分.てんさい生産費統計_個別, ComConst.調査区分.経営分析調査_てんさい生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.てんさい

                Case ComConst.調査区分.牛乳生産費統計_個別, ComConst.調査区分.経営分析調査_牛乳生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.牛乳

                Case ComConst.調査区分.子牛生産費統計_個別, ComConst.調査区分.経営分析調査_子牛生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.肉用子牛

                Case ComConst.調査区分.乳用雄育成牛生産費統計_個別, ComConst.調査区分.経営分析調査_乳用雄育成牛生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.乳用雄育成牛

                Case ComConst.調査区分.交雑種育成牛生産費統計_個別, ComConst.調査区分.経営分析調査_交雑種育成牛生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.交雑種育成牛

                Case ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.去勢若齢肥育牛

                Case ComConst.調査区分.乳用雄肥育牛生産費統計_個別, ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.乳用雄肥育牛

                Case ComConst.調査区分.交雑種肥育牛生産費統計_個別, ComConst.調査区分.経営分析調査_交雑種肥育牛生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.交雑種肥育牛

                Case ComConst.調査区分.肥育豚生産費統計_個別, ComConst.調査区分.経営分析調査_肥育豚生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.肥育豚

            End Select

            Dim xlSheet As Excel.Worksheet = Nothing
            Try
                xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

                Dim rng As Excel.Range = Nothing
                Try
                    adress = "D5"
                    rng = xlSheet.Range(adress)
                    Dim val As String = If(rng.Value Is Nothing, String.Empty, rng.Value.ToString)
                    adress = "D6"
                    rng = xlSheet.Range(adress)
                    Dim val2 As String = If(rng.Value Is Nothing, String.Empty, rng.Value.ToString)
                    adress = "D7"
                    rng = xlSheet.Range(adress)
                    Dim val3 As String = If(rng.Value Is Nothing, String.Empty, rng.Value.ToString)
                    If Not (val = str OrElse val2 = str OrElse val3 = str) Then
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

        ''' <summary>
        ''' 調査票の表示、データの保存
        ''' </summary>
        ''' <param name="pKey"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Sub Saveagain(pKey As DAOChosahyo.PrimaryKey)

            Dim kKey As New DAOChosahyo.KotenKey

            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Sheets = Nothing

            Try

                'Workbookを開く
                'REV-001 START-------------------
                'REV_002↓調査票テンプレートは読取専用で開く
                'xlBook = xlBooks.Open(System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.調査票.入力用ファイル名称(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun)))))
                xlBook = xlBooks.Open(System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.調査票.入力用ファイル名称(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun)))), ReadOnly:=True)
                'REV_002↑
                'REV-001 END-----------------
                xlSheets = xlBook.Worksheets

                SetData(xlSheets, pKey)

                Dim dtItem As DataTable
                Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                    '調査票項目マスタ取得
                    'REV-001 START-------------------
                    dtItem = DAOOther.GetChosahyoItemMasterSeidoUketori(db, CommonInfo.Chosakubun, _chosaNen)
                    'REV-001 END-----------------

                    '調査票シートデータ取得
                    'REV-001 START-------------------
                    dcChosahyo = ComUtil.Chosahyo.GetSheetData(dtItem, xlSheets, CType(Me, ComObjectProcess), _chosaNen)
                    'REV-001 END-----------------

                    '↓ 2022/03/01 調査対象経営体・担当名称取得
                    Dim dataChosahyo As DataTable           '調査票データ
                    Dim chosaSoshiki As String = Nothing    '調査対象経営体
                    Dim tantoMeisho As String = Nothing     '担当名称

                    dataChosahyo = DAOChosahyo.GetChosahyoTable_KeieiTanto(db, pKey)
                    For Each row As DataRow In dataChosahyo.Rows
                        chosaSoshiki = row(0).ToString
                        tantoMeisho = row(1).ToString
                    Next
                    '↑ 2022/03/01 調査対象経営体・担当名称取得

                    Try
                        db.BeginTrans()

                        '調査票データ削除
                        DAOChosahyo.DeleteChosahyoTable(db, pKey, _kKey)

                        '>>>2022/01/27
                        '調査票データ追加
                        DAOChosahyo.InsertChosahyoTable(db, pKey, _kKey, dcChosahyo, chosaSoshiki, tantoMeisho)
                        'DAOChosahyo.InsertChosahyoTable(db, pKey, _kKey, dcChosahyo)
                        '<<<2022/01/27

                        db.CommitTrans()
                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try

                End Using


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

        End Sub

        ''' <summary>
        ''' データを設定する
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub SetData(xlSheets As Excel.Sheets, pKey As DAOChosahyo.PrimaryKey)
            Try

                Dim dtItem As DataTable
                Dim dtChosahyo As Dictionary(Of String, DataTable)
                Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    '調査票項目マスタ取得
                    dtItem = DAOOther.GetChosahyoItemMasterSeidoUketori(db, CommonInfo.Chosakubun, _chosaNen, ComConst.数式区分.数式ではない)

                    '調査票テーブル取得
                    dtChosahyo = DAOChosahyo.GetChosahyoTable(db, pKey)
                End Using

                '調査票項目取得
                dcChosahyo = ComUtil.Chosahyo.GetItem(dtItem, dtChosahyo)

                xlApp.Interactive = False

                'Excel前処理
                BeforeExcel()

                '調査票シートデータ設定
                ComUtil.Chosahyo.SetSheetData(dcChosahyo, xlSheets, CType(Me, ComObjectProcess))

            Catch ex As Exception
                Throw
            Finally
                'Excel後処理
                AfterExcel()
                xlApp.Interactive = True
            End Try
        End Sub

        ''' <summary>
        ''' 項目番号と収支Codeの対応Datatable取得
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <returns></returns>
        Private Function GetSyushiCodeDataTable(xlSheets As Excel.Sheets) As DataTable
            Dim syushicodeTable As DataTable
            Dim seisanhiNum As Integer

            '取込調査区分が生産1～3のどれか確認
            Dim sheet As String = String.Empty
            Dim adress As String = String.Empty
            Dim str As String = String.Empty
            sheet = "経営体概要"
            str = CommonInfo.Chosakubun
            Select Case CommonInfo.Chosakubun
                Case ComConst.調査区分.なたね生産費統計_個別, ComConst.調査区分.経営分析調査_なたね生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.なたね

                Case ComConst.調査区分.そば生産費統計_個別, ComConst.調査区分.経営分析調査_そば生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.そば

                Case ComConst.調査区分.二条大麦生産費統計_個別, ComConst.調査区分.経営分析調査_二条大麦生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.二条大麦

                Case ComConst.調査区分.六条大麦生産費統計_個別, ComConst.調査区分.経営分析調査_六条大麦生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.六条大麦

                Case ComConst.調査区分.はだか麦生産費統計_個別, ComConst.調査区分.経営分析調査_はだか麦生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.はだか麦

                Case ComConst.調査区分.米生産費統計_個別, ComConst.調査区分.米生産費統計_組織法人
                    str = ComConst.中間集計表.中間集計表生産費区分.米

                Case ComConst.調査区分.小麦生産費統計_個別, ComConst.調査区分.小麦生産費統計_組織法人
                    str = ComConst.中間集計表.中間集計表生産費区分.小麦

                Case ComConst.調査区分.大豆生産費統計_個別, ComConst.調査区分.大豆生産費統計_組織法人
                    str = ComConst.中間集計表.中間集計表生産費区分.大豆

                Case ComConst.調査区分.原料用かんしょ生産費統計_個別
                    str = ComConst.中間集計表.中間集計表生産費区分.原料用かんしょ

                Case ComConst.調査区分.原料用ばれいしょ生産費統計_個別, ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.原料用ばれいしょ

                Case ComConst.調査区分.さとうきび生産費統計_個別, ComConst.調査区分.経営分析調査_さとうきび生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.さとうきび

                Case ComConst.調査区分.てんさい生産費統計_個別, ComConst.調査区分.経営分析調査_てんさい生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.てんさい

                Case ComConst.調査区分.牛乳生産費統計_個別, ComConst.調査区分.経営分析調査_牛乳生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.牛乳

                Case ComConst.調査区分.子牛生産費統計_個別, ComConst.調査区分.経営分析調査_子牛生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.肉用子牛

                Case ComConst.調査区分.乳用雄育成牛生産費統計_個別, ComConst.調査区分.経営分析調査_乳用雄育成牛生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.乳用雄育成牛

                Case ComConst.調査区分.交雑種育成牛生産費統計_個別, ComConst.調査区分.経営分析調査_交雑種育成牛生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.交雑種育成牛

                Case ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.去勢若齢肥育牛

                Case ComConst.調査区分.乳用雄肥育牛生産費統計_個別, ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.乳用雄肥育牛

                Case ComConst.調査区分.交雑種肥育牛生産費統計_個別, ComConst.調査区分.経営分析調査_交雑種肥育牛生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.交雑種肥育牛

                Case ComConst.調査区分.肥育豚生産費統計_個別, ComConst.調査区分.経営分析調査_肥育豚生産費
                    str = ComConst.中間集計表.中間集計表生産費区分.肥育豚

            End Select

            Dim xlSheet As Excel.Worksheet = Nothing
            Try
                xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

                Dim rng As Excel.Range = Nothing
                Try
                    adress = "D5"
                    rng = xlSheet.Range(adress)
                    Dim val As String = If(rng.Value Is Nothing, String.Empty, rng.Value.ToString)
                    adress = "D6"
                    rng = xlSheet.Range(adress)
                    Dim val2 As String = If(rng.Value Is Nothing, String.Empty, rng.Value.ToString)
                    adress = "D7"
                    rng = xlSheet.Range(adress)
                    Dim val3 As String = If(rng.Value Is Nothing, String.Empty, rng.Value.ToString)

                    If val = str Then
                        seisanhiNum = 1
                    ElseIf val2 = str Then
                        seisanhiNum = 2
                    ElseIf val3 = str Then
                        seisanhiNum = 3
                    End If

                Finally
                    ReleaseComObject(rng)
                End Try
            Finally
                ReleaseComObject(xlSheet)
            End Try

            ' 入力用データ取込用データテーブル作成
            Dim syushicodePath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & System.IO.Path.DirectorySeparatorChar & "Syuushicode.csv"
            syushicodeTable = New DataTable

            syushicodeTable.Columns.Add("調査区分", Type.GetType("System.String"))
            syushicodeTable.Columns.Add("項目番号", Type.GetType("System.String"))
            syushicodeTable.Columns.Add("収入Code", Type.GetType("System.String"))
            syushicodeTable.Columns.Add("支出Code", Type.GetType("System.String"))
            syushicodeTable.Columns.Add("出力内容区分", Type.GetType("System.String"))
            syushicodeTable.Columns.Add("収支区分", Type.GetType("System.String"))
            syushicodeTable.Columns.Add("列", Type.GetType("System.Int32"))

            Using reader As New System.IO.StreamReader(syushicodePath, System.Text.Encoding.Default)
                While (reader.Peek() >= 0)
                    Dim buf As String = reader.ReadLine()
                    Dim item() As String = buf.Split(","c)

                    Dim row As DataRow = syushicodeTable.NewRow()
                    row("調査区分") = item(0)
                    row("項目番号") = item(1)
                    row("収入Code") = item(2)
                    row("支出Code") = item(3)
                    row("出力内容区分") = item(4)
                    row("収支区分") = item(5)

                    If CInt(item(4)) = ComConst.中間集計表.入力用シート取込.出力内容区分.数量 _
                        Or CInt(item(4)) = ComConst.中間集計表.入力用シート取込.出力内容区分.金額 _
                        Or CInt(item(4)) = ComConst.中間集計表.入力用シート取込.出力内容区分.負担割合 Then
                        row("列") = ComConst.中間集計表.入力用シート取込.出力内容列(CInt(item(4)))(seisanhiNum - 1)
                    Else
                        row("列") = ComConst.中間集計表.入力用シート取込.出力内容列(CInt(item(4)))(0)
                    End If
                    syushicodeTable.Rows.Add(row)
                End While

            End Using

            Return syushicodeTable
        End Function

        ''' <summary>
        ''' 調査票シートデータ取得
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <param name="tourokusakujoKubun"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSheetData(dt As DataTable, xlSheets As Excel.Sheets, comObject As ComObjectProcess, tourokusakujoKubun As Integer) As Dictionary(Of String, DAOChosahyo.中間集計表項目)
            Dim ret As New Dictionary(Of String, DAOChosahyo.中間集計表項目)

            Dim sheets = From dr In dt Group dr By dr!シート名 Into Group Select シート名
            Dim xlSheet As Excel.Worksheet = Nothing
            Dim syushicodeTable As DataTable

            syushicodeTable = GetSyushiCodeDataTable(xlSheets)

            For Each sheet In sheets
                Try
                    If sheet Is DBNull.Value Then
                        Dim query = From dr In dt Where dr("シート名").ToString = sheet.ToString Select dr

                        For Each dr As DataRow In query
                            Dim item As New DAOChosahyo.中間集計表項目

                            With item
                                .値 = Nothing
                                .型区分 = dr("型区分").ToString
                                If Not String.IsNullOrEmpty(dr("登録削除区分").ToString) Then
                                    .登録削除区分 = Integer.Parse(dr("登録削除区分").ToString)
                                Else
                                    .登録削除区分 = 0
                                End If
                            End With
                            If dr("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                                ret.Add(dr("項目番号").ToString, item)
                            Else
                                If (Not String.IsNullOrEmpty(item.値)) OrElse tourokusakujoKubun = ComConst.中間集計表.登録削除区分.enm.削除 Then
                                    ret.Add(dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & 1, item)
                                End If
                            End If

                        Next
                    ElseIf sheet.ToString = "入力用" Or sheet.ToString = "入力用（収入）" Or sheet.ToString = "入力用（支出）" Then
                        Dim meisaiNoDictionary As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
                        Dim hutanWariaiSumDictionary As Dictionary(Of String, HutanWariaiSum) = New Dictionary(Of String, HutanWariaiSum)
                        Dim cel As Excel.Range = Nothing
                        Dim rng1 As Excel.Range = Nothing
                        Dim rng2 As Excel.Range = Nothing
                        Dim rng3 As Excel.Range = Nothing
                        Dim rng4 As Excel.Range = Nothing
                        Dim rngArr As Excel.Range = Nothing

                        xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

                        'シート保護確認
                        If xlSheet.ProtectContents Then
                            xlSheet.Unprotect()
                        End If

                        Try
                            '出力する生産費の変更、シートの更新
                            Dim adress As String = String.Empty
                            Dim prng As Excel.Range = Nothing

                            adress = "I11:BL10310"
                            prng = xlSheet.Range(adress)
                            prng.Calculate()

                            Dim arrData(,) As Object
                            Dim last As Integer

                            '1行目が空の場合、continue
                            rng1 = xlSheet.Range(ComConst.中間集計表.入力用シート取込.Col.First & ComConst.中間集計表.入力用シート取込.Row.First)
                            If rng1.Value Is Nothing Then
                                Continue For
                            End If

                            '最終行を取得
                            rng2 = xlSheet.Range(ComConst.中間集計表.入力用シート取込.Col.First & ComConst.中間集計表.入力用シート取込.Row.First + 1)
                            If Not rng2.Value Is Nothing Then
                                rng3 = rng1.End(Excel.XlDirection.xlDown)
                                last = rng3.Row
                            Else
                                last = rng1.Row
                            End If

                            '入力用シートデータ取得
                            rngArr = xlSheet.Range(ComConst.中間集計表.入力用シート取込.Col.First & ComConst.中間集計表.入力用シート取込.Row.First & ":" _
                                        & ComConst.中間集計表.入力用シート取込.Col.Last & last)
                            arrData = DirectCast(rngArr.Value, Object(,))

                            Dim query = From dr In dt Where dr("シート名").ToString = sheet.ToString Select dr
                            cel = xlSheet.Cells
                            For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                                ' 収支Code0(=未入力行)の場合、continue
                                If arrData(i, 48).ToString = "0" Then
                                    Continue For
                                End If

                                'ADD START 2023/03/14
                                If arrData(i, 8) Is Nothing Then
                                    Continue For
                                End If
                                'ADD END 2023/03/14

                                ' 取込対象年度データでない場合、continue
                                If Not arrData(i, 8).ToString = _chosaNen Then
                                    Continue For
                                End If

                                ' 収支Codeから取込先の項目番号(複数の可能性有)を取得する
                                Dim idx = i
                                Dim itemName As String = If(arrData(i, 4).ToString = "収入", "収入Code", "支出Code")
                                Dim itemNos = From syushicode In syushicodeTable
                                              Where syushicode("調査区分").ToString = CommonInfo.Chosakubun _
                                              And syushicode(itemName).ToString = arrData(idx, 48).ToString
                                              Select syushicode

                                'REV_003↓
                                '生産と自給牧草の合計割合を保つ為のマップ
                                Dim seisanBokusouWariai As New Dictionary(Of Integer, Decimal)()
                                'REV_003↑
                                '項目番号が取得できなかった場合(取込対象調査区分データでない場合)、continue
                                If itemNos.Count = 0 Then
                                    Continue For
                                End If

                                For Each itemNo In itemNos
                                    ' 中間集計表の取込対象列取得
                                    For Each dr As DataRow In query
                                        ' 対象項目番号でない場合、continue
                                        If Not dr("項目番号").ToString = itemNo("項目番号").ToString Then
                                            Continue For
                                        End If

                                        ' 中間集計表項目生成
                                        Dim item As New DAOChosahyo.中間集計表項目
                                        With item
                                            .シート名 = dr("シート名").ToString
                                            .追加項目 = dr("追加項目").ToString
                                            .行位置 = ComConst.中間集計表.入力用シート取込.Row.First + i - 1
                                            .列位置 = CInt(itemNo("列"))
                                            .型区分 = dr("型区分").ToString
                                            If dr("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                                                .有効桁数 = Integer.Parse(dr("有効桁数").ToString)
                                                .小数点以下桁数 = Integer.Parse(dr("小数点以下桁数").ToString)
                                                .明細番号 = Integer.Parse(dr("明細番号").ToString)
                                            Else
                                                '可変の場合、項目番号は明細番号付きとする
                                                If meisaiNoDictionary.ContainsKey(dr("項目番号").ToString) Then
                                                    meisaiNoDictionary(dr("項目番号").ToString) += 1
                                                Else
                                                    meisaiNoDictionary.Add(dr("項目番号").ToString, 1)
                                                End If
                                                .明細番号 = meisaiNoDictionary(dr("項目番号").ToString)
                                            End If
                                            If Not String.IsNullOrEmpty(dr("ラウンド桁数").ToString) Then
                                                .ラウンド桁数 = Integer.Parse(dr("ラウンド桁数").ToString)
                                            Else
                                                .ラウンド桁数 = 0
                                            End If
                                            If Not String.IsNullOrEmpty(dr("乗算値").ToString) Then
                                                .乗算値 = Integer.Parse(dr("乗算値").ToString)
                                            Else
                                                .乗算値 = 0
                                            End If
                                            .登録削除区分 = Integer.Parse(dr("登録削除区分").ToString)

                                        End With

                                        'データ整形
                                        Dim rng As Excel.Range = Nothing
                                        'REV_003↓
                                        Dim rngBokuso As Excel.Range = Nothing
                                        Dim targetCols() As Integer = {44, 45, 46}
                                        Dim seisanBokusouWariaiSum As Decimal
                                        Dim jikyubokusowariaiCol As Integer = 47
                                        Dim targetCols2() As Integer = {57, 59, 61}
                                        Dim targetCols3() As Integer = {58, 60, 62}
                                        'REV_003↑

                                        Try
                                            rng = DirectCast(cel.Item(item.行位置, item.列位置), Excel.Range)

                                            'REV_003↓
                                            Dim fValue As String
                                            '自給牧草割合の値を取得する
                                            rngBokuso = DirectCast(cel.Item(item.行位置, jikyubokusowariaiCol), Excel.Range)
                                            Dim bValue As String = If(rngBokuso.Value Is Nothing, Nothing, rngBokuso.Value.ToString)

                                            '自給牧草割合の値により、各割合を計算する
                                            If String.IsNullOrEmpty(bValue) Then
                                                If targetCols.Contains(item.列位置) Then
                                                    fValue = "100"
                                                Else
                                                    fValue = If(rng.Value Is Nothing, Nothing, rng.Value.ToString)
                                                End If
                                            Else
                                                If targetCols.Contains(item.列位置) Then
                                                    Dim seisanWariaiValue As Decimal
                                                    If (rng.Value Is Nothing) Then
                                                        seisanWariaiValue = 0
                                                    Else
                                                        seisanWariaiValue = Decimal.Parse(rng.Value.ToString)
                                                    End If

                                                    '最終割合を計算するため、Decimalに変換する
                                                    Dim bokusouWariaiValue As Decimal
                                                    If bValue Is Nothing OrElse Not Decimal.TryParse(bValue, bokusouWariaiValue) Then
                                                        bokusouWariaiValue = 0
                                                    End If

                                                    seisanBokusouWariaiSum = seisanWariaiValue + bokusouWariaiValue
                                                    seisanBokusouWariai(item.列位置) = seisanBokusouWariaiSum
                                                    fValue = Math.Round((seisanWariaiValue / seisanBokusouWariaiSum) * 100, 1, MidpointRounding.AwayFromZero).ToString("F1")

                                                ElseIf item.列位置 = jikyubokusowariaiCol Then
                                                    '自給牧草最終割合を計算する
                                                    Dim bokusouSaisyuWariai As String = Nothing
                                                    '最終割合を計算するため、Decimalに変換する
                                                    Dim bokusouWariaiValue As Decimal = Decimal.Parse(bValue)
                                                    For Each kv In seisanBokusouWariai
                                                        seisanBokusouWariaiSum = kv.Value
                                                        If seisanBokusouWariaiSum > 0 Then
                                                            bokusouSaisyuWariai = Math.Round((bokusouWariaiValue / seisanBokusouWariaiSum) * 100, 1, MidpointRounding.AwayFromZero).ToString("F1")
                                                            Exit For
                                                        End If
                                                    Next
                                                    fValue = bokusouSaisyuWariai
                                                Else
                                                    fValue = If(rng.Value Is Nothing, Nothing, rng.Value.ToString)
                                                End If
                                            End If

                                            If CInt(itemNo("出力内容区分")) = CInt(ComConst.中間集計表.入力用シート取込.出力内容区分.自給項目総数量) Then
                                                If targetCols2.Contains(item.列位置) Then
                                                    Dim rngSeisanSuryo As Excel.Range = DirectCast(cel(item.行位置, item.列位置), Excel.Range)
                                                    Dim rngBokusoSuryo As Excel.Range = DirectCast(cel(item.行位置, 63), Excel.Range)

                                                    Dim seisanSuryouValue As Decimal = Convert.ToDecimal(If(rngSeisanSuryo.Value, 0))
                                                    Dim bokusouSuryouValue As Decimal = Convert.ToDecimal(If(rngBokusoSuryo.Value, 0))

                                                    fValue = (seisanSuryouValue + bokusouSuryouValue).ToString()
                                                End If
                                            End If

                                            If CInt(itemNo("出力内容区分")) = CInt(ComConst.中間集計表.入力用シート取込.出力内容区分.利子農具金額) Then
                                                If targetCols3.Contains(item.列位置) Then
                                                    Dim rngSeisanRishi As Excel.Range = DirectCast(cel(item.行位置, item.列位置), Excel.Range)
                                                    Dim rngBokusoRishi As Excel.Range = DirectCast(cel(item.行位置, 64), Excel.Range)

                                                    Dim seisanRishiValue As Decimal = Convert.ToDecimal(If(rngSeisanRishi.Value, 0))
                                                    Dim bokusouRishiuValue As Decimal = Convert.ToDecimal(If(rngBokusoRishi.Value, 0))

                                                    fValue = (seisanRishiValue + bokusouRishiuValue).ToString()
                                                End If
                                            End If

                                            'Dim sValue As String = If(rng.Value Is Nothing, Nothing, rng.Value.ToString)
                                            Dim sValue As String = fValue
                                            'REV_003↑

                                            Dim val As Decimal
                                            Dim multiValue As Double = 0

                                            If Decimal.TryParse(sValue, val) Then
                                                If item.乗算値 > 0 Then
                                                    multiValue = Decimal.Parse(sValue) * item.乗算値
                                                Else
                                                    multiValue = Decimal.Parse(sValue)
                                                End If
                                                sValue = multiValue.ToString

                                                If item.ラウンド桁数 > 0 Then
                                                    multiValue = ToHalfAdjust(Decimal.Parse(sValue), item.ラウンド桁数)
                                                Else
                                                    multiValue = Decimal.Parse(sValue)
                                                End If
                                                sValue = multiValue.ToString
                                            End If

                                            If String.IsNullOrEmpty(dr("追加項目").ToString) Then
                                                item.値 = sValue
                                            Else
                                                If String.IsNullOrEmpty(sValue) OrElse sValue = "0" Then
                                                    item.値 = Nothing
                                                ElseIf (sValue.Contains("稲わら") OrElse sValue.Contains("自給牧草")) Then
                                                    item.値 = "1"
                                                Else
                                                    item.値 = dr("追加項目").ToString
                                                End If
                                            End If
                                        Finally
                                            ReleaseComObject(rng)
                                        End Try

                                        ' データ追加
                                        If dr("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                                            '可変でない(＝合算が必要)な値は前年対比シートから取得するが、
                                            '負担割合のみ入力用シートから平均値を取る
                                            Dim parsedValue As Decimal

                                            '値がない場合、加算しない
                                            If String.IsNullOrEmpty(item.値) Then
                                                Exit For
                                            End If

                                            'decimal変換できない場合、加算しない
                                            If Not Decimal.TryParse(item.値, parsedValue) Then
                                                Exit For
                                            End If

                                            '牧草の値がある場合、牧草の値も加算
                                            Dim bokusouValue As Decimal
                                            If Decimal.TryParse(arrData(i, 62).ToString, bokusouValue) And
                                                 CInt(itemNo("出力内容区分")) = CInt(ComConst.中間集計表.入力用シート取込.出力内容区分.数量) Then

                                                parsedValue += bokusouValue
                                            ElseIf Decimal.TryParse(arrData(i, 63).ToString, bokusouValue) And
                                                CInt(itemNo("出力内容区分")) = CInt(ComConst.中間集計表.入力用シート取込.出力内容区分.金額) Then

                                                parsedValue += bokusouValue
                                            End If

                                            '値を追加／更新
                                            If hutanWariaiSumDictionary.ContainsKey(dr("項目番号").ToString) Then
                                                hutanWariaiSumDictionary(dr("項目番号").ToString).sum += parsedValue
                                                hutanWariaiSumDictionary(dr("項目番号").ToString).cnt += 1

                                                If CInt(itemNo("出力内容区分")) = CInt(ComConst.中間集計表.入力用シート取込.出力内容区分.負担割合) Or
                                                    CInt(itemNo("出力内容区分")) = CInt(ComConst.中間集計表.入力用シート取込.出力内容区分.自給牧草負担割合) Then
                                                    '負担割合の場合、平均値を設定
                                                    ret(dr("項目番号").ToString).値 = ToHalfAdjust(hutanWariaiSumDictionary(dr("項目番号").ToString).sum / hutanWariaiSumDictionary(dr("項目番号").ToString).cnt, item.ラウンド桁数).ToString
                                                Else
                                                    '負担割合以外の場合、合計値を設定
                                                    ret(dr("項目番号").ToString).値 = hutanWariaiSumDictionary(dr("項目番号").ToString).sum.ToString
                                                End If

                                            Else
                                                Dim hutanWariaiSum = New HutanWariaiSum()
                                                hutanWariaiSum.sum = parsedValue
                                                hutanWariaiSum.cnt = 1
                                                hutanWariaiSumDictionary.Add(dr("項目番号").ToString, hutanWariaiSum)

                                                item.値 = parsedValue.ToString
                                                ret.Add(dr("項目番号").ToString, item)
                                            End If

                                        Else
                                            If (Not String.IsNullOrEmpty(item.値)) OrElse tourokusakujoKubun = ComConst.中間集計表.登録削除区分.enm.削除 Then
                                                ret.Add(dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & item.明細番号, item)
                                            End If
                                        End If

                                        Exit For
                                    Next
                                Next
                            Next
                        Finally
                            ReleaseComObject(cel)
                            ReleaseComObject(rng1)
                            ReleaseComObject(rng2)
                            ReleaseComObject(rng3)
                            ReleaseComObject(rng4)
                            ReleaseComObject(rngArr)
                        End Try
                    Else

                        xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

                        'シート保護確認
                        If xlSheet.ProtectContents Then
                            xlSheet.Unprotect()
                        End If

                        Dim cel As Excel.Range = Nothing
                        Try
                            '出力する生産費の変更、シートの更新
                            Dim psheet As String = String.Empty
                            Dim adress As String = String.Empty
                            Dim prng As Excel.Range = Nothing
                            psheet = "前年対比（収入）"
                            If sheet.ToString = psheet Then
                                adress = "AC2"
                                xlSheet = DirectCast(xlSheets.Item(psheet), Excel.Worksheet)
                                prng = xlSheet.Range(adress)
                                prng.Value = ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME

                                adress = "AD7:AL321"
                                prng = xlSheet.Range(adress)
                                prng.Calculate()
                            End If
                            psheet = "前年対比（支出）"
                            If sheet.ToString = psheet Then
                                adress = "AC2"
                                xlSheet = DirectCast(xlSheets.Item(psheet), Excel.Worksheet)
                                prng = xlSheet.Range(adress)
                                prng.Value = ComConst.中間集計表.中間集計表生産費区分名称.Seisanhi_NAME

                                adress = "AD7:AR393"
                                prng = xlSheet.Range(adress)
                                prng.Calculate()
                            End If

                            xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

                            cel = xlSheet.Cells

                            Dim query = From dr In dt Where dr("シート名").ToString = sheet.ToString Select dr
                            For Each dr As DataRow In query

                                Dim item As New DAOChosahyo.中間集計表項目
                                With item
                                    .シート名 = dr("シート名").ToString
                                    .追加項目 = dr("追加項目").ToString
                                    .行位置 = Integer.Parse(dr("行位置").ToString)
                                    .列位置 = Integer.Parse(dr("列位置").ToString)
                                    .型区分 = dr("型区分").ToString
                                    If dr("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                                        .有効桁数 = Integer.Parse(dr("有効桁数").ToString)
                                        .小数点以下桁数 = Integer.Parse(dr("小数点以下桁数").ToString)
                                    End If
                                    If Not String.IsNullOrEmpty(dr("ラウンド桁数").ToString) Then
                                        .ラウンド桁数 = Integer.Parse(dr("ラウンド桁数").ToString)
                                    Else
                                        .ラウンド桁数 = 0
                                    End If
                                    If Not String.IsNullOrEmpty(dr("乗算値").ToString) Then
                                        .乗算値 = Integer.Parse(dr("乗算値").ToString)
                                    Else
                                        .乗算値 = 0
                                    End If
                                    .登録削除区分 = Integer.Parse(dr("登録削除区分").ToString)
                                    .明細番号 = Integer.Parse(dr("明細番号").ToString)
                                End With

                                Dim rng As Excel.Range = Nothing
                                Try
                                    rng = DirectCast(cel.Item(item.行位置, item.列位置), Excel.Range)

                                    Dim sValue As String = If(rng.Value Is Nothing, Nothing, rng.Value.ToString)
                                    Dim val As Decimal
                                    Dim multiValue As Double = 0

                                    If Not item.シート名 = "経営体概要" Then
                                        If Decimal.TryParse(sValue, val) Then
                                            If item.乗算値 > 0 Then
                                                multiValue = Decimal.Parse(sValue) * item.乗算値
                                            Else
                                                multiValue = Decimal.Parse(sValue)
                                            End If
                                            sValue = multiValue.ToString

                                            If item.ラウンド桁数 > 0 Then
                                                multiValue = ToHalfAdjust(Decimal.Parse(sValue), item.ラウンド桁数)
                                            Else
                                                multiValue = Decimal.Parse(sValue)
                                            End If
                                            sValue = multiValue.ToString
                                        End If
                                    End If

                                    If String.IsNullOrEmpty(dr("追加項目").ToString) Then
                                        If Not item.シート名 = "経営体概要" Then
                                            If String.IsNullOrEmpty(sValue) OrElse sValue = "0" Then
                                                item.値 = Nothing
                                            Else
                                                item.値 = sValue
                                            End If
                                        Else
                                            item.値 = sValue
                                        End If
                                    Else
                                        If String.IsNullOrEmpty(sValue) OrElse sValue = "0" Then
                                            item.値 = Nothing
                                        Else
                                            item.値 = dr("追加項目").ToString
                                        End If
                                    End If
                                Finally
                                    ReleaseComObject(rng)
                                End Try
                                If dr("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                                    ret.Add(dr("項目番号").ToString, item)
                                Else
                                    If (Not String.IsNullOrEmpty(item.値)) OrElse tourokusakujoKubun = ComConst.中間集計表.登録削除区分.enm.削除 Then
                                        ret.Add(dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & Integer.Parse(dr("明細番号").ToString), item)
                                    End If
                                End If

                            Next
                        Finally
                            ReleaseComObject(cel)
                        End Try
                    End If
                Finally
                    ReleaseComObject(xlSheet)
                End Try
            Next

            Return ret
        End Function

        '@(f)
        '機能       ：ToHalfAdjust
        '返り値     ：四捨五入された数値
        '引き数     ：丸め対象のDecimal値、戻り値の有効桁数の精度
        '機能説明   ：指定した精度の数値に四捨五入
        '備考       ：
        Friend Shared Function ToHalfAdjust(ByVal dValue As Decimal, ByVal iDigits As Integer) As Decimal
            Dim dCoef As Decimal
            If iDigits <= 1 Then
                dCoef = 1
            Else
                dCoef = CDec((10 ^ (iDigits - 1)))
            End If
            If dValue > 0 Then
                Return CDec(Int(CDec(dValue * dCoef + 0.5)) / dCoef)
            Else
                Return CDec(Fix(CDec(dValue * dCoef - 0.5)) / dCoef)
            End If
        End Function

        ''' <summary>
        ''' Excel前処理
        ''' </summary>
        ''' <remarks></remarks>
        Protected Sub BeforeExcel()
            Try

                ''手動計算にする
                xlApp.Calculation = Excel.XlCalculation.xlCalculationManual
                '画面更新を無効にする
                xlApp.ScreenUpdating = False
                'イベント発生を無効にする
                xlApp.EnableEvents = False
                'アラート表示を無効にする
                xlApp.DisplayAlerts = False

            Catch ex As Exception
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' Excel後処理
        ''' </summary>
        ''' <remarks></remarks>
        Protected Sub AfterExcel()
            Try

                'アラート表示を有効にする
                xlApp.DisplayAlerts = True
                'イベント発生を有効にする
                xlApp.EnableEvents = True
                '画面更新を有効にする
                xlApp.ScreenUpdating = True
                ''自動計算にする
                xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic

            Catch ex As Exception
                Throw
            End Try
        End Sub
    End Class
End Class
