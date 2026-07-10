Imports Microsoft.Office.Interop

''' <summary>
''' 調査票審査論理管理（基本・追加）画面
''' </summary>
''' <remarks></remarks>
Public Class BRA1710F

    ''' <summary>タイトル文字列</summary>
    Private Const TITLE As String = "調査票審査論理管理（{0}）"

    ''' <summary>エラーチェック種別</summary>
    Private _errType As ComConst.エラーチェック種別.enm

    Public Sub New(errType As ComConst.エラーチェック種別.enm)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

        _errType = errType

        lblSyori.Text = String.Format(TITLE, ComConst.エラーチェック種別.リスト(_errType))
    End Sub

    Private Sub BRA1710F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '更新日時設定
            Me.UpdateDate(txtUpdateDate)
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
        Dim xlForm As BRA1720X = Nothing

        Try

            xlForm = New BRA1720X(Me, _errType)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally

        End Try
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Try
            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of OpenFileDialog)(Me, IniFileInfo.ExcelInPath)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Dim ret As Boolean
            '調査票審査論理管理（基本・追加）取込クラス生成
            Using ExcelImport = New ImportChosahyoShinsaRonri(_errType)
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

    Private Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click
        Try

            Dim fileName As String = ""
            '本省の場合は局、事務所、センター番号を表示させない。
            If _errType = ComConst.エラーチェック種別.enm.基本 Then
                fileName = String.Format(ComConst.調査票審査論理.出力用ファイル名称.reportName, ComConst.エラーチェック種別.リスト(_errType)) & "_" _
                                    & CommonInfo.ChosakubunName & ".xlsx"
            Else
                fileName = String.Format(ComConst.調査票審査論理.出力用ファイル名称.reportName, ComConst.エラーチェック種別.リスト(_errType)) & "_" _
                                    & CommonInfo.ChosakubunName & "_" _
                                    & CommonInfo.Kyoku & "_" _
                                    & CommonInfo.Jimusyo & "_" _
                                    & CommonInfo.Center & ".xlsx"
            End If

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '調査票審査論理（基本・追加）出力
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New BRA1710R(_errType, filePath)
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
            dtn = DAOOther.GetChosahyoShinsaRonriUpdateDate(db, _errType)
        End Using

        If Not dtn.Equals(String.Empty) Then
            txt.Text = DateTime.Parse(dtn).ToString(ComConst.DATETIME_FORMAT)
        Else
            txt.Text = String.Empty
        End If
    End Sub

    ''' <summary>
    ''' 調査票審査論理管理（基本・追加）取込クラス
    ''' </summary>
    ''' <remarks></remarks>
    Private Class ImportChosahyoShinsaRonri
        Inherits ExcelProcess

        ''' <summary>エラーチェック種別</summary>
        Private _errType As ComConst.エラーチェック種別.enm

        ''' <summary>進捗ダイアログ</summary>
        Private ProgressDialog As New ProgressDialog()

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New(ByVal errType As ComConst.エラーチェック種別.enm)
            MyBase.New(True)

            _errType = errType
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
            Dim lstDc As List(Of Dictionary(Of String, String)) = ComUtil.ChosahyoShinsaRonri.GetSheetData(xlSheets, CType(Me, ComObjectProcess))

            'エラーチェック
            Dim details As New List(Of String)
            If Not ComUtil.ChosahyoShinsaRonri.CheckError(lstDc, details, _errType) Then
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                'メッセージフォーム表示
                Message.ShowMsgForm(myParent, MessageID.MSG_E_024, {String.Join(vbCrLf, details)})
                Return ret
            End If

            'プログレスバーの最大値
            ProgressDialog.Maximum = lstDc.Count

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try

                    db.BeginTrans()

                    '調査票審査論理削除
                    DAOOther.DeleteChosahyoShinsaRonri(db, _errType)

                    For Each dc As Dictionary(Of String, String) In lstDc
                        '調査票審査論理追加
                        DAOOther.InsertChosahyoShinsaRonri(db, dc, _errType)

                        '進捗を進める
                        ProgressDialog.AddValue = 1
                    Next

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

            Const HEADER_TITLE As String = "調査票審査論理入力・修正（{0}）"

            'ファイル形式
            Dim FileFormat As New Dictionary(Of String, String) From { _
                  {"B1", CommonInfo.ChosakubunName} _
                , {"B2", String.Format(HEADER_TITLE, ComConst.エラーチェック種別.リスト(_errType))} _
                , {"B4", "エラーサイン"} _
                , {"G4", "区分"} _
            }

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シート名チェック
                Dim bln As Boolean = False
                For Each wsh As Excel.Worksheet In xlSheets
                    Try
                        If wsh.Name.Equals(ComConst.調査票審査論理.出力用ファイル名称.SheetName) Then
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
                xlSheet = DirectCast(xlSheets.Item(ComConst.調査票審査論理.出力用ファイル名称.SheetName), Excel.Worksheet)

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
