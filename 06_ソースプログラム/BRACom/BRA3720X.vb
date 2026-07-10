Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop.Forms
Imports System.Runtime.InteropServices
Imports System.Reflection

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.10.11 |Daiko               | 要件No1 バージョン区分追加
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 個別結果表審査論理入力・修正（基本・追加）（EXCEL）
''' </summary>
''' <remarks></remarks>

Public Class BRA3720X
    Inherits ExcelInputBaseClass

    ''' <summary>保存ボタン</summary>
    Private WithEvents btnSaveClose As CommandButton
    ''' <summary>戻るボタン</summary>
    Private WithEvents btnNoSaveClose As CommandButton

    ''' <summary>エラーチェック種別</summary>
    Private _errType As ComConst.エラーチェック種別.enm
    ''' <summary>ヘッダータイトル文字列</summary>
    Private Const HEADER_TITLE As String = "個別結果表審査論理入力・修正（{0}）画面"

    ''' <summary>Excelユーザーフォームハンドル</summary>
    Private _formHwnd As Win32WindowWrapper

    ' REV_001↓
    ''' <summary>バージョン区分</summary>
    Private _versionKbn As String
    ' REV_001↓

    ''' <summary>更新日時設定デリゲート</summary>
    ' REV_001↓
    'Private Delegate Sub UpdateDateDelegate(txt As System.Windows.Forms.TextBox)
    Private Delegate Sub UpdateDateDelegate(txt As System.Windows.Forms.TextBox, cbo As System.Windows.Forms.ComboBox)
    ' REV_001↑
    ''' <summary>個別結果表審査論理管理（基本・追加）画面</summary>
    Private _Form As BRA3710F

    ' REV_001↓
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <param name="errType"></param>
    ''' <param name="versionKbn"></param>
    'Public Sub New(ByRef frm As BRA3710F, ByVal errType As ComConst.エラーチェック種別.enm)
    Public Sub New(ByRef frm As BRA3710F, versionKbn As String, ByVal errType As ComConst.エラーチェック種別.enm)
        ' REV_001↑
        MyBase.New(frm, System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.個別結果表審査論理.入力用ファイル名称), True)
        Try
            _Form = frm
            _errType = errType

            ' REV_001↓
            _versionKbn = versionKbn
            ' REV_001↑

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
                Dim lstDc As List(Of Dictionary(Of String, String)) = ComUtil.KobetsuKekkahyoShinsaRonri.GetSheetData(xlSheets, CType(Me, ComObjectProcess))

                '進捗を進める
                ProgressDialog.AddValue = 1

                'エラーチェック
                Dim details As New List(Of String)
                ' REV_001↓
                'If Not ComUtil.KobetsuKekkahyoShinsaRonri.CheckError(lstDc, details, _errType) Then
                If Not ComUtil.KobetsuKekkahyoShinsaRonri.CheckError(lstDc, details, _errType, _versionKbn) Then
                    ' REV_001↓
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

                        '個別結果表審査論理削除
                        ' REV_001↓
                        'DAOOther.DeleteKobetsuKekkahyoShinsaRonri(db, _errType)
                        DAOOther.DeleteKobetsuKekkahyoShinsaRonri(db, _versionKbn, _errType)
                        ' REV_001↑

                        For Each dc As Dictionary(Of String, String) In lstDc
                            '個別結果表審査論理追加
                            ' REV_001↓
                            'DAOOther.InsertKobetsuKekkahyoShinsaRonri(db, dc, _errType)
                            DAOOther.InsertKobetsuKekkahyoShinsaRonri(db, _versionKbn, dc, _errType)
                            ' REV_001↑
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
                '個別結果表審査論理取得
                ' REV_001↓
                'dt = DAOOther.GetKobetsuKekkahyoShinsaRonri(db, _errType)
                dt = DAOOther.GetKobetsuKekkahyoShinsaRonri(db, _versionKbn, _errType)
                ' REV_001↑
            End Using

            'シートデータ設定
            ComUtil.KobetsuKekkahyoShinsaRonri.SetSheetData(dt, xlSheets, CType(Me, ExcelProcess))
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
        ' REV_001↓
        '_Form.Invoke(New UpdateDateDelegate(AddressOf _Form.UpdateDate), _Form.txtUpdateDate)
        _Form.Invoke(New UpdateDateDelegate(AddressOf _Form.UpdateDate), _Form.txtUpdateDate, _Form.cboVerKubun)
        ' REV_001↑
    End Sub
End Class
