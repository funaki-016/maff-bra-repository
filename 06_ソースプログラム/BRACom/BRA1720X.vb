Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop.Forms
Imports System.Runtime.InteropServices
Imports System.Reflection

''' <summary>
''' 調査票審査論理入力・修正（基本・追加）（EXCEL）
''' </summary>
''' <remarks></remarks>

Public Class BRA1720X
    Inherits ExcelInputBaseClass

    ''' <summary>保存ボタン</summary>
    Private WithEvents btnSaveClose As CommandButton
    ''' <summary>戻るボタン</summary>
    Private WithEvents btnNoSaveClose As CommandButton

    ''' <summary>エラーチェック種別</summary>
    Private _errType As ComConst.エラーチェック種別.enm
    ''' <summary>ヘッダータイトル文字列</summary>
    Private Const HEADER_TITLE As String = "調査票審査論理入力・修正（{0}）画面"

    ''' <summary>Excelユーザーフォームハンドル</summary>
    Private _formHwnd As Win32WindowWrapper

    ''' <summary>更新日時設定デリゲート</summary>
    Private Delegate Sub UpdateDateDelegate(txt As System.Windows.Forms.TextBox)
    ''' <summary>調査票審査論理管理（基本・追加）画面</summary>
    Private _Form As BRA1710F

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef frm As BRA1710F, ByVal errType As ComConst.エラーチェック種別.enm)
        MyBase.New(frm, System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.調査票審査論理.入力用ファイル名称), True)
        Try
            _Form = frm
            _errType = errType

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
                Dim lstDc As List(Of Dictionary(Of String, String)) = ComUtil.ChosahyoShinsaRonri.GetSheetData(xlSheets, CType(Me, ComObjectProcess))

                '進捗を進める
                ProgressDialog.AddValue = 1

                'エラーチェック
                Dim details As New List(Of String)
                If Not ComUtil.ChosahyoShinsaRonri.CheckError(lstDc, details, _errType) Then
                    '進捗ダイアログを閉じる
                    ProgressDialog.endDispose()
                    'メッセージフォーム表示
                    Message.ShowMsgForm(_formHwnd, MessageID.MSG_E_024, {String.Join(vbCrLf, details)})
                    Exit Sub
                End If

                '進捗を進める
                ProgressDialog.AddValue = 1

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    Try

                        db.BeginTrans()

                        '調査票審査論理削除
                        DAOOther.DeleteChosahyoShinsaRonri(db, _errType)

                        For Each dc As Dictionary(Of String, String) In lstDc
                            '調査票審査論理追加
                            DAOOther.InsertChosahyoShinsaRonri(db, dc, _errType)
                        Next

                        db.CommitTrans()

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
            SetExcelValue(xlSheet, 2, 2, (String.Format(HEADER_TITLE, ComConst.エラーチェック種別.リスト(_errType))))

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
                '調査票審査論理取得
                dt = DAOOther.GetChosahyoShinsaRonri(db, _errType)
            End Using

            'シートデータ設定
            ComUtil.ChosahyoShinsaRonri.SetSheetData(dt, xlSheets, CType(Me, ExcelProcess))
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
