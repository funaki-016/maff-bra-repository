Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop.Forms
Imports System.Runtime.InteropServices
Imports System.Reflection

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.12.14 |Daiko               | 要件No4 バージョン区分追加
'//  REV_002   | 2022.12.20 |Daiko               | 要件No15 集計結果検討表（報告用）追加
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 集計結果検討表作成論理入力・修正（EXCEL）
''' </summary>
''' <remarks></remarks>
Public Class BRA5420X
    Inherits ExcelInputBaseClass

    ''' <summary>保存ボタン</summary>
    Private WithEvents btnSaveClose As CommandButton
    ''' <summary>戻るボタン</summary>
    Private WithEvents btnNoSaveClose As CommandButton

    ' REV_002↓
    ''' <summary>ヘッダータイトル文字列</summary>
    'Private Const HEADER_TITLE As String = "集計結果検討表作成論理入力・修正画面"
    Private Shared HEADER_TITLE As New Dictionary(Of Integer, String) From {
        {ComConst.集計結果検討表作成論理.論理種別.集計結果検討表, "集計結果検討表作成論理入力・修正画面"} _
      , {ComConst.集計結果検討表作成論理.論理種別.集計結果検討表_報告用, "集計結果検討表（報告用）作成論理入力・修正画面"}
    }
    ' REV_002↑

    ''' <summary>Excelユーザーフォームハンドル</summary>
    Private _formHwnd As Win32WindowWrapper

    ' REV_001↓
    ''' <summary>バージョン区分</summary>
    Private _versionKbn As String
    ' REV_001↑

    ''' <summary>更新日時設定デリゲート</summary>
    ' REV_001↓
    'Private Delegate Sub UpdateDateDelegate(txt As System.Windows.Forms.TextBox, einouKeieitai As String)
    Private Delegate Sub UpdateDateDelegate(txt As System.Windows.Forms.TextBox, einouKeieitai As String, cbo As System.Windows.Forms.ComboBox)
    ' REV_001↑
    ''' <summary>集計結果検討表作成論理管理画面</summary>
    Private _Form As BRA5410F

    ''' <summary>営農経営体区分</summary>
    Private _einouKeieitai As String
    ''' <summary>調査区分</summary>
    Private _chosakubun As String

    ' REV_002↓
    ''' <summary>論理種別</summary>
    Private _logicType As ComConst.集計結果検討表作成論理.論理種別
    ' REV_002↑

    ' REV_001 REV_002↓
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="einouKeieitai"></param>
    ''' <param name="logicType"></param>
    'Public Sub New(ByRef frm As BRA5410F, einouKeieitai As String)
    Public Sub New(ByRef frm As BRA5410F, versionKbn As String, einouKeieitai As String, logicType As ComConst.集計結果検討表作成論理.論理種別)
        'MyBase.New(frm, System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.集計結果検討表作成論理.入力用ファイル名称), True)
        MyBase.New(frm, System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.集計結果検討表作成論理.入力用ファイル名称(logicType)), True)
        ' REV_001 REV_002↑
        Try
            _Form = frm
            ' REV_001↓
            _versionKbn = versionKbn
            ' REV_001↑ 
            _einouKeieitai = einouKeieitai
            _chosakubun = ComUtil.GetChosakubun(_einouKeieitai)
            ' REV_002↓
            _logicType = logicType
            ' REV_002↑ 

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
                Dim lstDc As List(Of Dictionary(Of String, String)) = ComUtil.SyukeiKekkaKentohyoSakuseiRonri.GetSheetData(xlSheets, CType(Me, ComObjectProcess))

                '進捗を進める
                ProgressDialog.AddValue = 1

                'エラーチェック
                Dim details As New List(Of String)
                ' REV_001 REV_002↓
                'If Not ComUtil.SyukeiKekkaKentohyoSakuseiRonri.CheckError(lstDc, details, _chosakubun) Then
                If Not ComUtil.SyukeiKekkaKentohyoSakuseiRonri.CheckError(lstDc, details, _chosakubun, _versionKbn, _logicType) Then
                    ' REV_001 REV_002↑
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

                        '集計結果検討表作成論理削除
                        ' REV_002↓
                        If _logicType = ComConst.集計結果検討表作成論理.論理種別.集計結果検討表 Then
                            ' REV_001↓
                            'DAOOther.DeleteSyukeiKekkaKentohyoSakuseiRonri(db, _chosakubun)
                            DAOOther.DeleteSyukeiKekkaKentohyoSakuseiRonri(db, _chosakubun, _versionKbn)
                            ' REV_001↑
                        Else
                            DAOOther.DeleteSyukeiKekkaKentohyoHoukokuyoSakuseiRonri(db, _chosakubun, _versionKbn)
                        End If
                        ' REV_002↑

                        For Each dc As Dictionary(Of String, String) In lstDc
                            '集計結果検討表作成論理追加
                            ' REV_002↓
                            If _logicType = ComConst.集計結果検討表作成論理.論理種別.集計結果検討表 Then
                                ' REV_001↓
                                'DAOOther.InsertSyukeiKekkaKentohyoSakuseiRonri(db, _chosakubun, dc)
                                DAOOther.InsertSyukeiKekkaKentohyoSakuseiRonri(db, _chosakubun, _versionKbn, dc)
                                ' REV_001↑
                            Else
                                DAOOther.InsertSyukeiKekkaKentohyoHoukokuyoSakuseiRonri(db, _chosakubun, _versionKbn, dc)
                            End If
                            ' REV_002↑
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

            SetExcelValue(xlSheet, 1, 2, ComUtil.GetChosakubunName(_chosakubun))
            ' REV_002↓
            'SetExcelValue(xlSheet, 2, 2, HEADER_TITLE)
            SetExcelValue(xlSheet, 2, 2, HEADER_TITLE(_logicType))
            ' REV_001↑

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
                ' REV_002↓
                If _logicType = ComConst.集計結果検討表作成論理.論理種別.集計結果検討表 Then
                    '集計結果検討表作成論理取得
                    ' REV_001↓
                    'dt = DAOOther.GetSyukeiKekkaKentohyoSakuseiRonri(db, _chosakubun)
                    dt = DAOOther.GetSyukeiKekkaKentohyoSakuseiRonri(db, _chosakubun, _versionKbn)
                    ' REV_001↑
                Else
                    dt = DAOOther.GetSyukeiKekkaKentohyoHoukokuyoSakuseiRonri(db, _chosakubun, _versionKbn)
                End If
                ' REV_002↑
            End Using

            'シートデータ設定
            ComUtil.SyukeiKekkaKentohyoSakuseiRonri.SetSheetData(dt, xlSheets, CType(Me, ExcelProcess))
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
        '_Form.Invoke(New UpdateDateDelegate(AddressOf _Form.UpdateDate), _Form.txtUpdateDate, _einouKeieitai)
        _Form.Invoke(New UpdateDateDelegate(AddressOf _Form.UpdateDate), _Form.txtUpdateDate, _einouKeieitai, _Form.cboVerKubun)
        ' REV_001↑
    End Sub
End Class
