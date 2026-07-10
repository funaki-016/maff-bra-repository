Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions

''' <summary>
''' 労働時間整理ファイル取込画面
''' </summary>
''' <remarks></remarks>
''' 
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_000   | 2020.12.26 |日本ｺﾝﾋﾟｭｰﾀｼｽﾃﾑ     | 新規追加
'//  REV_001   | 2022.10.06 |大興電子通信        | 要件No.20 BRACom化(本省、局工程でも利用)
'//  REV_002   | 2023.10.30 |大興電子通信        | 変更要件No.6
'//            |            |                    |
'//*************************************************************************************************
Public Class BRA2710F

    Private Shared sensasu As String = ""
    Private Shared chosaNen As Integer
    ' REV_001↓
    Private Shared kyoku As String = ""
    Private Shared jimusyo As String = ""
    Private Shared kyoten As String = ""
    ' REV_001↑
    Private Shared excelName As String = ""
    Private Shared chosaKbn As String = Nothing

    Private Sub BRA2710F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '調査年コンボボックス設定
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.Chosahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
            End Using

            '局コンボボックス設定(REV_001)
            ComUtil.SetKyokuComboBox(cboKyoku)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnTmpOut_Click(sender As Object, e As EventArgs) Handles btnTmpOut.Click
        'テンプレート出力ボタンクリック時の処理
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckErrorTmp(msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            'ファイル保存ダイアログ表示
            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath),
                                                                            cboRoudoFileShurui.SelectedItem.ToString + ".xlsx", "EXCELブック（*.xlsx）|*.xlsx")
            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            'ファイルコピー
            System.IO.File.Copy(System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), cboRoudoFileShurui.SelectedItem.ToString + ".xlsx"), filePath, True)

            'メッセージフォーム表示
            Message.ShowMsgBox(MessageID.MSG_I_002, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            'カーソルを戻す
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnReference_Click(sender As Object, e As EventArgs) Handles btnReference.Click
        '参照ボタンクリック時の処理
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'フォルダパス取得
            Dim folderPath As String = ComUtil.GetFolderPath(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainInPath, IniFileInfo.ExcelInPath))

            If folderPath.Equals(String.Empty) Then
                Exit Sub
            End If

            '取得したパスをセット
            Me.txtRoudoFile.Text = folderPath

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            'カーソルを戻す
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        '取込ボタンクリック時の処理
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'Excel件数格納要変数
            Dim excelCount As Integer

            '事前エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckErrorImport(msgId, excelCount) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '問い合わせメッセージを表示→「いいえ」なら処理終了
            If Message.ShowMsgBox(MessageID.MSG_Q_053, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                Exit Sub
            End If

            '調査年取得
            Dim year As String = cboChosaNen.SelectedValue.ToString

            'ファイル拡張子
            Dim searchPattern As String() = {"*.xlsx"}

            'ファイル
            Dim files As IEnumerable(Of String) = {}
            For Each pattern As String In searchPattern
                files = files.Concat(System.IO.Directory.EnumerateFiles(Me.txtRoudoFile.Text, pattern, System.IO.SearchOption.TopDirectoryOnly))
            Next

            Dim bolImpRet As Boolean
            '労働時間整理ファイル取込クラス生成
            Using ImportRoudoFile = New ImportRoudoFile(year)
                '処理実行
                bolImpRet = ImportRoudoFile.Execute(files, Me)
            End Using

            If bolImpRet Then
                '完了メッセージフォーム表示
                '↓MOD MS 2022/01/25
                'Message.ShowMsgForm(Me, MessageID.MSG_I_049, files.ToArray)
                Dim outMsgVal As String() = {String.Join(vbCrLf, files)}
                Message.ShowMsgForm(Me, MessageID.MSG_I_049, outMsgVal.ToArray)
                '↑MOD MS 2022/01/25
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
    ''' エラーチェック_テンプレート出力ボタン
    ''' </summary>
    ''' <param name="num"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckErrorTmp(ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        '労働時間整理ファイル種類の入力チェック
        If cboRoudoFileShurui.SelectedItem Is Nothing Then
            'エラーメッセージ
            msgId = MessageID.MSG_E_106
            Return ret
        End If

        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' エラーチェック_取込ボタン
    ''' </summary>
    ''' <param name="num"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckErrorImport(ByRef msgId As String, ByRef excelCount As Integer) As Boolean
        Dim ret As Boolean = False

        '調査年選択チェック
        If cboChosaNen.SelectedValue Is Nothing Then
            'エラーメッセージ
            msgId = MessageID.MSG_E_004
            Return ret
        Else
            chosaNen = CInt(cboChosaNen.SelectedValue.ToString)
        End If

        '農政局選択チェック(REV_001)
        If cboKyoku.SelectedValue Is DBNull.Value Then
            'エラーメッセージ
            msgId = MessageID.MSG_E_117
            Return ret
        Else
            kyoku = CStr(cboKyoku.SelectedValue)
        End If

        '実査設置拠点選択チェック(REV_001)
        If cboKyoten.SelectedValue Is DBNull.Value Then
            'エラーメッセージ
            msgId = MessageID.MSG_E_118
            Return ret
        Else
            jimusyo = CStr(CType(cboKyoten.SelectedItem, DataRowView)("事務所番号"))
            kyoten = CStr(cboKyoten.SelectedValue)
        End If

        '労働時間整理ファイル項目欄入力チェック
        If txtRoudoFile.Text Is Nothing Or "" = txtRoudoFile.Text Then
            '未入力エラー　エラーメッセージ
            msgId = MessageID.MSG_E_107
            Return ret
        End If

        'ファイル数取得
        Dim dirFileCount As Integer = System.IO.Directory.GetFiles(Me.txtRoudoFile.Text, "*", System.IO.SearchOption.TopDirectoryOnly).Length
        If dirFileCount = 0 Then
            'エクセルファイルなしエラー　エラーメッセージ
            msgId = MessageID.MSG_E_108
            Return ret
        ElseIf dirFileCount > 20 Then
            'エクセルファイル数最大値オーバーエラー　エラーメッセージ
            msgId = MessageID.MSG_E_109
            Return ret
        End If

        'ファイル数セット
        excelCount = dirFileCount
        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' 労働時間整理ファイル取込クラス
    ''' </summary>
    ''' <remarks></remarks>
    Private Class ImportRoudoFile
        Inherits ExcelProcess

        ''' <summary>調査年</summary>
        Private _chosaNen As String
        ''' <summary>拠点キー</summary>
        Private _kKey As DAOOther.RoudouKyotenKey

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
            ' REV_001↓
            '_kKey = New DAOOther.RoudouKyotenKey(CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
            _kKey = New DAOOther.RoudouKyotenKey(kyoku, jimusyo, kyoten)
            ' REV_001↑
        End Sub

        ''' <summary>
        ''' 処理実行
        ''' </summary>
        ''' <param name="files"></param>
        ''' <param name="form"></param>
        ''' <remarks></remarks>
        Public Function Execute(files As IEnumerable(Of String), myParent As Form) As Boolean
            Dim ret As List(Of String)() = {Nothing, Nothing}

            Dim blnError As Boolean = False                 'INS MS 2022/01/25

            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Sheets = Nothing

            'Excelアプリ無効
            Me.DisableExcelApp()

            Try
                '進捗ダイアログを表示する
                ProgressDialog.Maximum = files.Count * 4
                ProgressDialog.Show(myParent)

                'エラーメッセージリスト
                Dim errMsgList As List(Of String) = New List(Of String)
                Dim outMsgList As String() = {}

                'エラーチェック実施
                For Each filePath As String In files
                    Try
                        '進捗を進める
                        ProgressDialog.AddValue = 1

                        'Workbookを開く（ファイル名も取得）
                        xlBook = xlBooks.Open(filePath)
                        excelName = xlBook.Name

                        Try
                            xlSheets = xlBook.Worksheets

                            Dim msgId As String = String.Empty
                            Dim msgPara As String() = {}

                            'エラーチェック
                            If Not Me.errCheck(xlSheets, errMsgList) Then
                                '↓MOD MS 2022/01/25
                                'outMsgList = {excelName}.Concat(errMsgList).ToArray
                                outMsgList = outMsgList.Concat({excelName}.Concat(errMsgList)).ToArray
                                errMsgList = New List(Of String)
                                blnError = True
                                '↑MOD MS 2022/01/25
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

                'エラーメッセージが出力された場合、こちらで処理する（IDはBRAE-110）
                'DEL MS 2022/01/25  If 0 < errMsgList.Count Then
                If blnError Then                    'INS MS 2022/01/25
                    'DEL MS 2022/01/25  ProgressDialog.ShowMsgForm(MessageID.MSG_E_110, outMsgList.ToArray)
                    Dim outMsgVal As String() = {String.Join(vbCrLf, outMsgList)}
                    ProgressDialog.ShowMsgForm(MessageID.MSG_E_024, outMsgVal.ToArray)              'INS MS 2022/01/25
                    'Message.ShowMsgForm(myParent, MessageID.MSG_E_024, outMsgList.ToArray)          'INS MS 2022/01/25
                    Return False
                End If

                For Each filePath As String In files
                    Try
                        '進捗を進める
                        ProgressDialog.AddValue = 1

                        'Workbookを開く（ファイル名も取得）
                        xlBook = xlBooks.Open(filePath)
                        excelName = xlBook.Name

                        Try
                            xlSheets = xlBook.Worksheets

                            Dim msgId As String = String.Empty
                            Dim msgPara As String() = {}

                            '↓INS MS 2022/01/25
                            '再度エラーチェックを実行して調査区分・センサス番号を取得する
                            Dim dmyMsgList As List(Of String) = New List(Of String)
                            If Not Me.errCheck(xlSheets, dmyMsgList) Then
                                Return False
                            End If
                            '↑INS MS 2022/01/25

                            '取込処理
                            If Not Me.Import(xlSheets, msgId, msgPara, errMsgList) Then
                                If msgPara.Length > 0 Then
                                    'エラーメッセージが出力された場合、こちらで処理する（IDはBRAE-110）
                                    ProgressDialog.ShowMsgForm(msgId, {IO.Path.GetFileName(filePath)}.Concat(msgPara).ToArray)
                                    Return False
                                End If
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

            Return True
        End Function

        Private Function errCheck(xlSheets As Excel.Sheets, ByRef errMsgList As List(Of String)) As Boolean
            Dim ret As Boolean = False

            'Excelシート存在チェック
            If Not Me.CheckExcelSheetExist(xlSheets) Then
                'メッセージ設定
                'DEL MS 2022/01/25  errMsgList.Add("1")
                errMsgList.Add("件数:1")                                  'INS MS 2022/01/25
                errMsgList.Add("労働時間整理ファイルではありません。")
                errMsgList.Add("")                                        'INS MS 2022/01/25
                Return ret
            End If

            'Excelファイルテンプレートチェック
            Dim strKbn = Me.CheckExcelFileTemplates(xlSheets)
            If strKbn = "0" Then
                'メッセージ設定
                'DEL MS 2022/01/25  errMsgList.Add("1")
                errMsgList.Add("件数:1")                                  'INS MS 2022/01/25
                errMsgList.Add("調査区分が一致していません。")
                errMsgList.Add("")                                        'INS MS 2022/01/25
                Return ret
            ElseIf strKbn = "9" Then
                'メッセージ設定
                'DEL MS 2022/01/25  errMsgList.Add("1")
                errMsgList.Add("件数:1")                                  'INS MS 2022/01/25
                errMsgList.Add("設定シートのセンサス番号の入力に誤りがあります。")
                errMsgList.Add("")                                        'INS MS 2022/01/25
                Return ret
            End If

            'エラーチェック
            Dim details As New List(Of String)
            Dim cnt = Me.CheckError(xlSheets, details, strKbn)
            If 0 < cnt Then
                'メッセージ設定
                'DEL MS 2022/01/25  errMsgList.Add(cnt.ToString)
                errMsgList.Add(String.Format("件数:{0}", cnt.ToString))   'INS MS 2022/01/25
                errMsgList.Add(String.Join(vbCrLf, details))
                errMsgList.Add("")                                        'INS MS 2022/01/25
                Return ret
            End If

            'エラーなし
            Return True
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
        Private Function Import(xlSheets As Excel.Sheets, ByRef msgId As String, ByRef msgPara As String(), ByRef errMsgList As List(Of String)) As Boolean
            Dim ret As Boolean = False

            '調査票シートデータ取得
            Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)

            '主キー設定
            Dim pKey As DAOChosahyo.PrimaryKey = New DAOChosahyo.PrimaryKey(chosaNen.ToString, sensasu)

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                '個別結果表存在チェック
                Dim pKobet As DAOKobetsuKekkahyo.PrimaryKey = New DAOKobetsuKekkahyo.PrimaryKey(pKey.chosaNen, pKey.censusNo)
                Dim kKobet As DAOKobetsuKekkahyo.KyotenKey = New DAOKobetsuKekkahyo.KyotenKey(_kKey.kyoku, _kKey.jimusho, _kKey.kyoten)
                Dim bln As Boolean = DAOKobetsuKekkahyo.CheckExistRoudou(db, pKobet, kKobet, chosaKbn) '労働時間整理ファイル用

                If bln Then
                    '個別結果表存在時確認メッセージ
                    If ProgressDialog.ShowMsgBox(MessageID.MSG_Q_054, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                        '「はい」の場合の処理は、トランザクション内に記述
                        Return ret
                    End If
                    '上書き確認確認メッセージ　※メッセージ056がメッセージ一覧にない⇒確認
                    If ProgressDialog.ShowMsgBox(MessageID.MSG_Q_056, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                        Return ret
                    End If

                End If

                Try
                    db.BeginTrans()

                    If bln Then
                        '個別結果表データ削除
                        DAOKobetsuKekkahyo.DeleteTable_Roudou(db, pKobet, kKobet, chosaKbn)
                    End If

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

                    '↓INS MS 2022/03/09  初回削除前に調査票データ取得
                    Dim dtPrevChosahyo As Dictionary(Of String, DataTable)
                    dtPrevChosahyo = DAOChosahyo.GetChosahyoTable(db, pKey)
                    '↑INS MS 2022/03/09  初回削除前に調査票データ取得
                    '調査票データ削除
                    DAOChosahyo.DeleteChosahyoTable_Roudou(db, pKey, _kKey, chosaKbn, False)

                    '①調査票項目マスタ情報取得
                    Dim dtItem As DataTable
                    Dim dtChosahyo As Dictionary(Of String, DataTable)

                    '調査票項目マスタ取得
                    '↓MOD MS 2022/01/25
                    'dtItem = DAOOther.GetChosahyoItemMasterRoudou(db, chosaKbn, chosaNen.ToString, ComConst.数式区分.数式ではない)
                    dtItem = MyGetChosahyoItemMasterRoudou(db, chosaKbn, chosaNen.ToString, ComConst.数式区分.数式ではない)
                    '↑MOD MS 2022/01/25
                    '調査票用主キー設定⇒設定済み
                    '調査票テーブル取得
                    dtChosahyo = DAOChosahyo.GetChosahyoTable(db, pKey)
                    '調査票項目取得
                    dcChosahyo = ComUtil.Chosahyo.GetItem(dtItem, dtChosahyo)

                    '労働時間整理ファイルシートデータ取得（②労働時間整理ファイルの入力データ取得）
                    Dim dcRoudouFile As Dictionary(Of String, DAOOther.労働時間整理ファイル項目)
                    dcRoudouFile = ComUtil.RoudouFile.GetRoudouData(dtItem, xlSheets, CType(Me, ComObjectProcess), chosaNen.ToString, chosaKbn, excelName)

                    '↓INS MS 2022/01/25
                    '調査票データ取得・削除
                    Dim dtPrev As Dictionary(Of String, DataTable)
                    dtPrev = MyGetDeleteChosahyoTable_Roudou(db, pKey, _kKey, chosaKbn)
                    '↑INS MS 2022/01/25

                    '調査票データ追加
                    '↓UPD MS 2022/03/09
                    'DAOChosahyo.InsertChosahyoTable_Roudou(db, pKey, _kKey, dcChosahyo, dcRoudouFile, chosaKbn, dtPrev, chosaSoshiki, tantoMeisho)
                    DAOChosahyo.InsertChosahyoTable_Roudou(db, pKey, _kKey, dcChosahyo, dcRoudouFile, chosaKbn, dtPrev, dtPrevChosahyo, chosaSoshiki, tantoMeisho)
                    '↑UPD MS 2022/03/09

                    db.CommitTrans()

                    '電子調査票の数式再計算処理
                    Saveagain(pKey)

                    ret = True
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
        ''' <param name="xlSheets">Excelシート</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function CheckExcelSheetExist(xlSheets As Excel.Sheets) As Boolean
            Dim ret As Boolean = True

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

            'シート名比較（「設定」シートと比較）
            If Not xlSheetNames.Contains("設定") Then
                ret = False
            End If

            Return ret
        End Function

        ''' <summary>
        ''' エラーチェック
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="details"></param>
        ''' <param name="strKbn"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function CheckError(xlSheets As Excel.Sheets, ByRef details As List(Of String), strKbn As String) As Integer

            Const max As Integer = ComConst.ERR_MESSAGE_MAX

            Dim msg As String
            Dim sheet As String = String.Empty
            Dim adress As String = String.Empty
            Dim xlSheet As Excel.Worksheet = Nothing

            '⑤経営種類及び生産費区分が対象調査区分と一致しているかを確認する。（畜産物生産費、農産物生産費の労働時間整理ファイルの場合のみ実施）
            If strKbn = "3" Or strKbn = "4" Then
                sheet = "設定"

                Dim renge As Excel.Range = Nothing
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

                    '経営種類の確認
                    adress = "D3"
                    renge = xlSheet.Range(adress)
                    Dim valKeiei As String = If(renge.Value Is Nothing, String.Empty, renge.Value.ToString)
                    If Not ComConst.労働時間整理ファイル.経営種類及び生産費区分.確認("経営種類").Contains(valKeiei) Then
                        '経営種類が存在しない場合
                        msg = "調査区分が一致していません。"
                        details.Add(msg)
                        Return 1
                    End If

                    '生産費区分の確認
                    Dim strSeisanHi As String
                    '値取得
                    adress = "F3"
                    renge = xlSheet.Range(adress)

                    '判定
                    Dim valSeisan = If(renge.Value Is Nothing, String.Empty, renge.Value.ToString)
                    If ComConst.労働時間整理ファイル.経営種類及び生産費区分.確認("畜生生産費区分").Contains(valSeisan) Then
                        '調査区分設定
                        chosaKbn = ChosaKbnHantei(valKeiei, valSeisan)
                    ElseIf ComConst.労働時間整理ファイル.経営種類及び生産費区分.確認("農生生産費区分").Contains(valSeisan) Then
                        '調査区分設定
                        chosaKbn = ChosaKbnHantei(valKeiei, valSeisan)
                    End If

                    If chosaKbn = ComConst.調査区分.牛乳生産費統計_個別 Or chosaKbn = ComConst.調査区分.経営分析調査_牛乳生産費 Or
                       chosaKbn = ComConst.調査区分.子牛生産費統計_個別 Or chosaKbn = ComConst.調査区分.経営分析調査_子牛生産費 Or
                       chosaKbn = ComConst.調査区分.去勢若齢肥育牛生産費統計_個別 Or chosaKbn = ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費 Or
                       chosaKbn = ComConst.調査区分.乳用雄育成牛生産費統計_個別 Or chosaKbn = ComConst.調査区分.経営分析調査_乳用雄育成牛生産費 Or
                       chosaKbn = ComConst.調査区分.乳用雄肥育牛生産費統計_個別 Or chosaKbn = ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費 Or
                       chosaKbn = ComConst.調査区分.交雑種育成牛生産費統計_個別 Or chosaKbn = ComConst.調査区分.経営分析調査_交雑種育成牛生産費 Or
                       chosaKbn = ComConst.調査区分.交雑種肥育牛生産費統計_個別 Or chosaKbn = ComConst.調査区分.経営分析調査_交雑種肥育牛生産費 Or
                       chosaKbn = ComConst.調査区分.肥育豚生産費統計_個別 Or chosaKbn = ComConst.調査区分.経営分析調査_肥育豚生産費 Then
                        '畜生の場合
                        strSeisanHi = "畜生生産費区分"
                    Else
                        '農生の場合
                        strSeisanHi = "農生生産費区分"
                    End If

                    If Not ComConst.労働時間整理ファイル.経営種類及び生産費区分.確認(strSeisanHi).Contains(valSeisan) Then
                        '生産費区分が存在しない場合
                        msg = "調査区分が一致していません。"
                        details.Add(msg)
                        Return 1
                    End If
                Finally
                    ReleaseComObject(renge)
                    ReleaseComObject(xlSheet)
                End Try
            Else
                '営農個人・営農法人の場合はそのままセット
                chosaKbn = strKbn
            End If

            '⑥電子調査票ですでに登録されたセンサス番号であるか確認を行う
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                If DAOChosahyo.GetChosahyoTable(db, chosaNen, sensasu, chosaKbn) <= 0 Then
                    '0件以下の場合はエラー
                    msg = "指定のセンサス番号の電子調査票は登録されていません。"
                    details.Add(msg)
                    Return 1
                End If
            End Using

            '⑦操作しているユーザが専門調査員の場合、その専門調査員が扱えるセンサス番号かどうか。
            If CommonInfo.SenmonChosain Then
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    If Not DAOOther.CheckSenmonChosainKyakutaiExist(db, CommonInfo.UserId, sensasu) Then
                        msg = "操作可能なセンサス番号ではありません。"
                        details.Add(msg)
                        Return 1
                    End If
                End Using
            End If

            '⑨調査票項目マスタ情報取得
            '※処理順上、先に取得する
            Dim dtItem As DataTable
            Dim dtChosahyo As Dictionary(Of String, DataTable)
            Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '調査票項目マスタ取得
                '↓MOD MS 2022/01/25
                'dtItem = DAOOther.GetChosahyoItemMasterRoudou(db, chosaKbn, chosaNen.ToString, ComConst.数式区分.数式ではない)
                dtItem = MyGetChosahyoItemMasterRoudou(db, chosaKbn, chosaNen.ToString, ComConst.数式区分.数式ではない)
                '↑MOD MS 2022/01/25
                '調査票用主キー設定
                Dim pKey = New DAOChosahyo.PrimaryKey(chosaNen.ToString, sensasu)
                '調査票テーブル取得
                dtChosahyo = DAOChosahyo.GetChosahyoTable(db, pKey)
            End Using

            '調査票項目取得
            dcChosahyo = ComUtil.Chosahyo.GetItem(dtItem, dtChosahyo)

            ' REV_001↓
            ''⑧農業経営統計調査システムにログインしている実査設置拠点を管轄している都道府県と、労働時間整理ファイルの都道府県（センサス番号内）が一致しているか。
            'If Not CommonInfo.Jimusyo = ComUtil.RoudouFile.ConvJimusyoNoRoudou(sensasu).PadLeft(2, "0"c) Then
            '⑧入力された実査設置拠点を管轄している都道府県と、労働時間整理ファイルの都道府県（センサス番号内）が一致しているか。
            If Not jimusyo.PadLeft(2, "0"c) = ComUtil.RoudouFile.ConvJimusyoNoRoudou(sensasu).PadLeft(2, "0"c) Then
                msg = "自管轄の都道府県と、取込む電子調査票の都道府県が異なっています。"
                details.Add(msg)
                Return 1
            End If
            ' REV_001↑

            '⑩労働時間整理ファイルの[労働時間シート※1]の名前確認
            'シート名は、労働時間整理ファイルから取得して、固定の名前で比較する
            '↓INS MS 2022/01/25
            '"）"の位置を取得
            Dim lastPosition As Integer = excelName.IndexOf("）", 0)
            'Excelファイルの拡張子を取得
            Dim ext = System.IO.Path.GetExtension(excelName)
            '比較するファイル名を形成
            Dim targetExcelName As String = excelName.Substring(0, lastPosition + 1) & ext
            Dim sheets = ComConst.労働時間整理ファイル.シートリスト.リスト(targetExcelName)
            For Each sheet In sheets
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)
                Catch ex As Exception
                    'シート名不一致エラー
                    msg = "シート名に誤りがあります。"
                    details.Add(msg)
                    Return 1
                End Try
            Next
            '↑INS MS 2022/01/25

            '⑪労働時間整理ファイルシートデータ取得（⑪労働時間整理ファイルの入力データ取得）
            Dim dcRoudouFile As Dictionary(Of String, DAOOther.労働時間整理ファイル項目)
            dcRoudouFile = ComUtil.RoudouFile.GetRoudouData(dtItem, xlSheets, CType(Me, ComObjectProcess), chosaNen.ToString, chosaKbn, excelName)

            Dim bolSheet1 = False
            Dim bolSheet2 = False
            Dim flg = False
            For Each kv As KeyValuePair(Of String, DAOOther.労働時間整理ファイル項目) In dcRoudouFile
                'DEL MS 2022/01/25  不要行    Dim aaa = kv.Value.シート名
                'DEL MS 2022/01/25  不要行    Dim bbb = ComConst.労働時間整理ファイル.労働時間シート名.シート名(chosaKbn)
                If Not String.IsNullOrEmpty(kv.Value.シート名) Then
                    If chosaKbn = ComConst.調査区分.営農類型別経営統計_個人 Then
                        '営農個人の場合、2シート分のチェックが必要のため別処理
                        If kv.Value.シート名 = "10_労働" Then
                            bolSheet1 = True
                        End If
                        If kv.Value.シート名 = "11_労働（指定品目）" Then
                            bolSheet2 = True
                        End If
                        If bolSheet1 And bolSheet2 Then
                            '一致するシート名あり
                            flg = True
                            Exit For
                        End If
                    ElseIf chosaKbn = ComConst.調査区分.営農類型別経営統計_法人 Then
                        '営農法人の場合、2シート分のチェックが必要のため別処理
                        If kv.Value.シート名 = "12_01_労働の概況" Then
                            bolSheet1 = True
                        End If
                        If kv.Value.シート名 = "12_02_労働の概況" Then
                            bolSheet2 = True
                        End If
                        If bolSheet1 And bolSheet2 Then
                            '一致するシート名あり
                            flg = True
                            Exit For
                        End If
                        '↓INS MS 2022/01/25
                    ElseIf chosaKbn = ComConst.調査区分.牛乳生産費統計_個別 OrElse chosaKbn = ComConst.調査区分.子牛生産費統計_個別 Or
                           chosaKbn = ComConst.調査区分.乳用雄育成牛生産費統計_個別 OrElse chosaKbn = ComConst.調査区分.交雑種育成牛生産費統計_個別 Or
                           chosaKbn = ComConst.調査区分.去勢若齢肥育牛生産費統計_個別 OrElse chosaKbn = ComConst.調査区分.乳用雄肥育牛生産費統計_個別 Or
                           chosaKbn = ComConst.調査区分.交雑種肥育牛生産費統計_個別 OrElse chosaKbn = ComConst.調査区分.肥育豚生産費統計_個別 Or
                           chosaKbn = ComConst.調査区分.経営分析調査_牛乳生産費 Or chosaKbn = ComConst.調査区分.経営分析調査_子牛生産費 Or
                           chosaKbn = ComConst.調査区分.経営分析調査_乳用雄育成牛生産費 OrElse chosaKbn = ComConst.調査区分.経営分析調査_交雑種育成牛生産費 Or
                           chosaKbn = ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費 OrElse chosaKbn = ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費 Or
                           chosaKbn = ComConst.調査区分.経営分析調査_交雑種肥育牛生産費 OrElse chosaKbn = ComConst.調査区分.経営分析調査_肥育豚生産費 Then
                        If kv.Value.シート名 = "【12】労働時間" Then
                            bolSheet1 = True
                        End If
                        If kv.Value.シート名 = ComConst.労働時間整理ファイル.労働時間シート名.シート名(chosaKbn) Then
                            bolSheet2 = True
                        End If
                        If bolSheet1 OrElse bolSheet2 Then
                            '一致するシート名あり
                            flg = True
                            Exit For
                        End If                        '↑INS MS 2022/01/25
                    ElseIf kv.Value.シート名 = ComConst.労働時間整理ファイル.労働時間シート名.シート名(chosaKbn) Then
                        '上記以外は1シートのため、これでチェック可能
                        '固定値確認のため、上記で実施
                        flg = True
                        Exit For
                    End If
                End If
            Next
            If Not flg Then
                'シート名不一致エラー
                msg = "シート名に誤りがあります。"
                details.Add(msg)
                Return 1
            End If

            Dim cnt As Integer = 0
            '⑫労働時間整理ファイルの各項目が、調査票項目マスタの型と一致しているか。
            msg = "シート名：「{0}」行数 {1}　列数 {2}　の型が一致しません。"
            For Each kv As KeyValuePair(Of String, DAOOther.労働時間整理ファイル項目) In dcRoudouFile
                If Not String.IsNullOrEmpty(kv.Value.値) Then
                    If kv.Value.型区分 = ComConst.型区分.数値型 Then
                        Dim val As Decimal
                        If Not Decimal.TryParse(kv.Value.値, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, val) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg, kv.Value.シート名, kv.Value.行位置, kv.Value.列位置))
                            If cnt = max Then Return cnt
                        Else
                            '指数表記となっている場合、小数点以下に変換する。
                            If val <> 0 And val < 0.0001 Then
                                kv.Value.値 = val.ToString("0.#####")
                            End If
                        End If
                    End If
                End If
            Next

            '⑬労働時間整理ファイルの各項目が、データベースの桁数に収まっているか。
            msg = "シート名：「{0}」行数 {1}　列数 {2}　の桁数がデータベースの桁数を超えています。"
            For Each kv As KeyValuePair(Of String, DAOOther.労働時間整理ファイル項目) In dcRoudouFile
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
                                    details.Add(String.Format(msg, kv.Value.シート名, kv.Value.行位置, kv.Value.列位置))
                                    If cnt = max Then Return cnt
                                End If
                            End If
                        End If
                    End If
                End If
            Next

            Return cnt
        End Function

        ''' <summary>
        ''' 労働時間整理ファイルテンプレートチェック
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <returns></returns>
        Private Function CheckExcelFileTemplates(xlSheets As Excel.Sheets) As String
            Dim hantei As ComConst.労働時間整理ファイル.取り込みファイル判定.判定内容
            Dim ret As String = "0"
            Dim xlSheet As Excel.Worksheet = Nothing
            Dim rng As Excel.Range = Nothing
            Dim kbn As String = Nothing

            '③設定シートのC1セルを参照する
            Try
                xlSheet = DirectCast(xlSheets.Item("設定"), Excel.Worksheet)
                Try
                    rng = xlSheet.Range("C1")

                    Dim strValue As String = CStr(rng.Value)
                    strValue = strValue.Replace(vbLf, "")

                    If strValue = "営農類型別経営統計（個人経営体）　設定シート" Then
                        'ファイルが「労働時間整理ファイル（営農類型別経営統計（個人経営体））」の場合
                        '※trueでなくどの区分かを返す必要がある
                        kbn = "1"
                    ElseIf strValue = "営農類型別経営統計（法人経営体）　設定シート" Then
                        'ファイルが「労働時間整理ファイル（営農類型別経営統計（法人経営体））」の場合
                        kbn = "2"
                    ElseIf strValue = "畜産物生産費統計　設定シート" Then
                        'ファイルが「労働時間整理ファイル（畜産物生産費統計）」の場合
                        kbn = "3"
                    ElseIf strValue = "農産物生産費統計　設定シート" Then
                        'ファイルが「労働時間整理ファイル（農産物生産費統計）」の場合
                        kbn = "4"
                    Else
                        'どれにも当てはまらないからエラー
                        kbn = "err"
                        'Return ret 'retには0入っているはず
                    End If
                Finally
                    ReleaseComObject(rng)
                End Try

            Finally
                ReleaseComObject(xlSheet)
            End Try

            If kbn = "err" Then
                Return ret
            End If

            '③センサス番号チェック
            If ComConst.労働時間整理ファイル.取り込みファイル判定.センサス桁数判定内容一覧.ContainsKey(kbn) Then
                hantei = ComConst.労働時間整理ファイル.取り込みファイル判定.センサス桁数判定内容一覧(kbn)
            Else
                Return "9"
            End If

            Dim blnSheetError As Boolean = False                'INS MS 2022/01/25
            Try
                xlSheet = DirectCast(xlSheets.Item(hantei.シート名), Excel.Worksheet)
                Try
                    rng = xlSheet.Range(hantei.セル番号)

                    Dim strValue As String = CStr(rng.Value)

                    'DEL MS 2022/01/25  strValue = strValue.Replace(vbLf, "")

                    'センサス番号が16桁か
                    If String.IsNullOrEmpty(strValue) Then      'INS MS 2022/01/25
                        ret = "9"                               'INS MS 2022/01/25
                    Else                                        'INS MS 2022/01/25
                        strValue = strValue.Replace(vbLf, "")   'INS MS 2022/01/25
                        If strValue.Length <> 16 Then
                            ret = "9"
                        Else
                            sensasu = strValue
                            ret = kbn
                        End If
                    End If                                      'INS MS 2022/01/25

                Finally
                    ReleaseComObject(rng)
                End Try

            Catch ex As Exception                               'INS MS 2022/01/25
                blnSheetError = True                            'INS MS 2022/01/25
            Finally
                If Not blnSheetError Then                       'INS MS 2022/01/25
                    ReleaseComObject(xlSheet)
                End If                                          'INS MS 2022/01/25
            End Try

            Return ret
        End Function

        ''' <summary>
        ''' 調査区分算出
        ''' </summary>
        ''' <param name="strKeieishurui"></param>
        ''' <param name="strSeisanhiKbn"></param>
        ''' <returns></returns>
        Private Function ChosaKbnHantei(strKeieishurui As String, strSeisanhiKbn As String) As String
            Dim strRet = ""
            If strKeieishurui = "個別経営体" Then
                '個別側
                Select Case strSeisanhiKbn
                    Case "米"
                        strRet = ComConst.調査区分.米生産費統計_個別
                    Case "小麦"
                        strRet = ComConst.調査区分.小麦生産費統計_個別
                    Case "二条大麦"
                        strRet = ComConst.調査区分.二条大麦生産費統計_個別
                    Case "六条大麦"
                        strRet = ComConst.調査区分.六条大麦生産費統計_個別
                    Case "はだか麦"
                        strRet = ComConst.調査区分.はだか麦生産費統計_個別
                    Case "そば"
                        strRet = ComConst.調査区分.そば生産費統計_個別
                    Case "大豆"
                        strRet = ComConst.調査区分.大豆生産費統計_個別
                    Case "原料用かんしょ"
                        strRet = ComConst.調査区分.原料用かんしょ生産費統計_個別
                    Case "原料用ばれいしょ"
                        strRet = ComConst.調査区分.原料用ばれいしょ生産費統計_個別
                    Case "なたね"
                        strRet = ComConst.調査区分.なたね生産費統計_個別
                    Case "てんさい"
                        strRet = ComConst.調査区分.てんさい生産費統計_個別
                    Case "さとうきび"
                        strRet = ComConst.調査区分.さとうきび生産費統計_個別
                    Case "牛乳"
                        strRet = ComConst.調査区分.牛乳生産費統計_個別
                    Case "子牛"
                        strRet = ComConst.調査区分.子牛生産費統計_個別
                    Case "乳用雄育成牛"
                        strRet = ComConst.調査区分.乳用雄育成牛生産費統計_個別
                    Case "交雑種育成牛"
                        strRet = ComConst.調査区分.交雑種育成牛生産費統計_個別
                    Case "去勢若齢肥育牛"
                        strRet = ComConst.調査区分.去勢若齢肥育牛生産費統計_個別
                    Case "乳用雄肥育牛"
                        strRet = ComConst.調査区分.乳用雄肥育牛生産費統計_個別
                    Case "交雑種肥育牛"
                        strRet = ComConst.調査区分.交雑種肥育牛生産費統計_個別
                    Case "肥育豚"
                        strRet = ComConst.調査区分.肥育豚生産費統計_個別
                End Select

            ElseIf strKeieishurui = "組織法人経営体" Then
                '組織法人側
                Select Case strSeisanhiKbn
                    Case "米"
                        strRet = ComConst.調査区分.米生産費統計_組織法人
                    Case "小麦"
                        strRet = ComConst.調査区分.小麦生産費統計_組織法人
                    Case "二条大麦"
                        strRet = ComConst.調査区分.経営分析調査_二条大麦生産費
                    Case "六条大麦"
                        strRet = ComConst.調査区分.経営分析調査_六条大麦生産費
                    Case "はだか麦"
                        strRet = ComConst.調査区分.経営分析調査_はだか麦生産費
                    Case "そば"
                        strRet = ComConst.調査区分.経営分析調査_そば生産費
                    Case "大豆"
                        strRet = ComConst.調査区分.大豆生産費統計_組織法人
                    Case "原料用ばれいしょ"
                        strRet = ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費
                    Case "なたね"
                        strRet = ComConst.調査区分.経営分析調査_なたね生産費
                    Case "てんさい"
                        strRet = ComConst.調査区分.経営分析調査_てんさい生産費
                    Case "さとうきび"
                        strRet = ComConst.調査区分.経営分析調査_さとうきび生産費
                    Case "牛乳"
                        strRet = ComConst.調査区分.経営分析調査_牛乳生産費
                    Case "子牛"
                        strRet = ComConst.調査区分.経営分析調査_子牛生産費
                    Case "乳用雄育成牛"
                        strRet = ComConst.調査区分.経営分析調査_乳用雄育成牛生産費
                    Case "交雑種育成牛"
                        strRet = ComConst.調査区分.経営分析調査_交雑種育成牛生産費
                    Case "去勢若齢肥育牛"
                        strRet = ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費
                    Case "乳用雄肥育牛"
                        strRet = ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費
                    Case "交雑種肥育牛"
                        strRet = ComConst.調査区分.経営分析調査_交雑種肥育牛生産費
                    Case "肥育豚"
                        strRet = ComConst.調査区分.経営分析調査_肥育豚生産費
                End Select

            End If

            Return strRet
        End Function

        ''' <summary>
        ''' 調査票の表示、データの保存
        ''' </summary>
        ''' <param name="pKey"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Sub Saveagain(pKey As DAOChosahyo.PrimaryKey)

            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Sheets = Nothing

            Try

                'Workbookを開く
                'REV_002↓調査票テンプレートは読取専用で開く
                'xlBook = xlBooks.Open(System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.調査票.入力用ファイル名称(Tuple.Create(chosaKbn, ComUtil.getVersionKubun(_chosaNen, chosaKbn)))))
                xlBook = xlBooks.Open(System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.調査票.入力用ファイル名称(Tuple.Create(chosaKbn, ComUtil.getVersionKubun(_chosaNen, chosaKbn)))), ReadOnly:=True)
                'REV_002↑
                xlSheets = xlBook.Worksheets

                SetData(xlSheets, pKey)

                Dim dtItem As DataTable
                Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                    '調査票項目マスタ取得
                    '↓MOD MS 2022/01/25
                    'dtItem = DAOOther.GetChosahyoItemMasterRoudou(db, chosaKbn, _chosaNen)
                    dtItem = MyGetChosahyoItemMasterRoudou(db, chosaKbn, _chosaNen)
                    '↑MOD MS 2022/01/25

                    '調査票シートデータ取得
                    '↓MOD MS 2022/01/25
                    'dcChosahyo = ComUtil.Chosahyo.GetSheetData(dtItem, xlSheets, CType(Me, ComObjectProcess), _chosaNen)
                    dcChosahyo = MyGetSheetData(dtItem, xlSheets, CType(Me, ComObjectProcess), _chosaNen)
                    '↑MOD MS 2022/01/25

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
                        DAOChosahyo.DeleteChosahyoTable_Roudou(db, pKey, _kKey, chosaKbn, True)

                        '調査票データ追加
                        DAOChosahyo.InsertChosahyoTable_RoudouRe(db, pKey, _kKey, dcChosahyo, chosaKbn, chosaSoshiki, tantoMeisho)

                        db.CommitTrans()
                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try

                End Using


            Catch ex As Exception
                Throw ex
            Finally
                'Sheetsの解放                   'INS MS 2022/01/25
                ReleaseComObject(xlSheets)      'INS MS 2022/01/25 
                'Workbookを閉じる
                If xlBook IsNot Nothing Then
                    xlBook.Saved = True
                    'DEL MS 2022/01/25  xlBook.Close()
                    xlBook.Close(False, Type.Missing, Type.Missing)     'INS MS 2022/01/25
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
                    '↓MOD MS 2022/01/25
                    'dtItem = DAOOther.GetChosahyoItemMasterRoudou(db, chosaKbn, _chosaNen, ComConst.数式区分.数式ではない)
                    dtItem = MyGetChosahyoItemMasterRoudou(db, chosaKbn, _chosaNen, ComConst.数式区分.数式ではない)
                    '↑MOD MS 2022/01/25

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
        ''' Excel前処理
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub BeforeExcel()
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
        Private Sub AfterExcel()
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

        '↓INS MS 2022/01/25
        ''' <summary>
        ''' 調査票項目マスタ取得
        ''' </summary>
        ''' <param name="db"></param>
        ''' <param name="chosaKubun"></param>
        ''' <param name="suushikiKubun"></param>
        ''' <returns></returns>
        ''' <remarks>'REV_005 ADD</remarks>
        Private Function MyGetChosahyoItemMasterRoudou(db As DBAccess, chosaKubun As String, chosaNen As String, Optional suushikiKubun As String = Nothing) As DataTable

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
        ''' 調査票データ取得・削除
        ''' </summary>
        ''' <param name="db"></param>
        ''' <param name="pKey"></param>
        ''' <param name="kKey"></param>
        ''' <param name="chosakbn"></param>
        ''' <returns></returns>
        Private Function MyGetDeleteChosahyoTable_Roudou(db As DBAccess,
                                                         pKey As DAOChosahyo.PrimaryKey,
                                                         kKey As DAOOther.RoudouKyotenKey,
                                                         chosakbn As String) As Dictionary(Of String, DataTable)

            '調査票データ取得
            Dim ret As New Dictionary(Of String, DataTable)
            Dim sbGet As System.Text.StringBuilder = Nothing
            Dim sbDel As System.Text.StringBuilder = Nothing
            Dim para As List(Of DBAccess.Parameter) = Nothing

            Try

                For Each tableName As String In ComConst.調査票.テーブル名称(chosakbn)
                    If Not tableName.Contains("＿可変") Then
                        sbGet = New System.Text.StringBuilder
                        sbDel = New System.Text.StringBuilder
                        para = New List(Of DBAccess.Parameter)

                        ' データ取得SQL文の設定
                        With sbGet
                            .AppendLine("SELECT * ")
                            .AppendLine(String.Format("  FROM ""{0}""", tableName))
                            .AppendLine("WHERE  調査年         = @調査年 ")
                            .AppendLine("AND    センサス番号   = @センサス番号 ")
                            If tableName = ComConst.調査票.テーブル名称(chosakbn)(0) Then
                                .AppendLine("AND    農政局         = @農政局 ")
                                .AppendLine("AND    都道府県       = @都道府県 ")
                                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                            End If
                        End With

                        ' データ削除SQL文の設定
                        With sbDel
                            .AppendLine("DELETE ")
                            .AppendLine(String.Format("FROM ""{0}""", tableName))
                            .AppendLine("WHERE  調査年         = @調査年 ")
                            .AppendLine("AND    センサス番号   = @センサス番号 ")
                            If tableName = ComConst.調査票.テーブル名称(chosakbn)(0) Then
                                .AppendLine("AND    農政局         = @農政局 ")
                                .AppendLine("AND    都道府県       = @都道府県 ")
                                .AppendLine("AND    実査設置拠点   = @実査設置拠点 ")
                            End If
                        End With

                        para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pKey.chosaNen))
                        para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pKey.censusNo))
                        If tableName = ComConst.調査票.テーブル名称(chosakbn)(0) Then
                            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kKey.kyoku))
                            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kKey.jimusho))
                            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kKey.kyoten))
                        End If

                        ' データ取得
                        Dim dt As New DataTable
                        dt = db.GetDataTable(sbGet.ToString, para)
                        ret.Add(tableName, dt)

                        ' データ削除
                        db.ExecuteNonQuery(sbDel.ToString, para)
                    End If
                Next

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
                        'REV-005 START-----------------------
                        '↓MOD MS 2022/01/25
                        'rng = xlSheet.Range(ComConst.調査票.シートデータ範囲(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun)))(sheet.ToString))
                        rng = xlSheet.Range(ComConst.調査票.シートデータ範囲(Tuple.Create(chosaKbn, ComUtil.getVersionKubun(_chosaNen, chosaKbn)))(sheet.ToString))
                        '↑MOD MS 2022/01/25
                        'REV-005 END-----------------------

                        arrData = DirectCast(rng.Value, Object(,))

                        Dim query = From dr In dt Where dr("シート名").ToString = sheet.ToString Select dr

                        Dim seidouketori_count As Integer = 9010001 'REV-010 ADD

                        For Each dr As DataRow In query
                            If dr("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                                Dim item As New DAOChosahyo.調査票項目
                                With item
                                    'REV-010 START-----------------------
                                    '↓MOD MS 2022/01/25
                                    'If dr("項目番号").ToString = "Q" + seidouketori_count.ToString("00000000") Then

                                    '    seidouketori_count = seidouketori_count + 1

                                    '    Continue For
                                    'End If
                                    '"Q090100"で始まる項目番号は処理しない
                                    If dr("項目番号").ToString.StartsWith("Q090100") Then
                                        Continue For
                                    End If
                                    '↑MOD MS 2022/01/25
                                    'REV-010 END  -----------------------

                                    .シート名 = dr("シート名").ToString
                                    .行位置 = Integer.Parse(dr("行位置").ToString)

                                    .列位置 = Integer.Parse(dr("列位置").ToString)

                                    'REV-005 START-----------------------
                                    If .行位置 = 0 Or .列位置 = 0 Then
                                        .値 = Nothing
                                    Else
                                        .値 = If(arrData(.行位置, .列位置) Is Nothing, Nothing, If(String.IsNullOrEmpty(arrData(.行位置, .列位置).ToString), Nothing, arrData(.行位置, .列位置).ToString))
                                        If .値 IsNot Nothing Then
                                            ComUtil.Chosahyo.EscapeToiawasesaki(dr, .値)
                                        End If
                                    End If
                                    'REV-005 END-----------------------

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
        '↑INS MS 2022/01/25

    End Class

    ''' <summary>
    ''' 農政局プルダウン選択変更(REV_001)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cboKyoku_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKyoku.SelectedIndexChanged
        Try
            '拠点コンボボックス設定
            ComUtil.SetKyotenComboBox(cboKyoku, cboKyoten)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 調査年（産）プルダウン選択変更(REV_001)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cboChosaNen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboChosaNen.SelectedIndexChanged
        Try
            ' 実査設置拠点工程の場合
            If CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Center Then
                '何もしない
                Return
            End If

            ' 本省、局工程の場合
            ' ～2021：利用不可、2022～：利用可
            Dim enabled = CInt(If(cboChosaNen.SelectedValue, 0)) >= 2022
            cboRoudoFileShurui.Enabled = enabled
            txtRoudoFile.Enabled = enabled
            btnTmpOut.Enabled = enabled
            btnReference.Enabled = enabled
            btnImport.Enabled = enabled

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
End Class
