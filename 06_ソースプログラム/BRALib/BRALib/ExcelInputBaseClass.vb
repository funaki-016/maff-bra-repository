Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop.Forms
Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.12.01 |TSP)                | フェーズ3 不具合対応No.12修正
'//  REV_002   | 2021.12.21 |日本ｺﾝﾋﾟｭｰﾀｼｽﾃﾑ　　 | 要件No1-⑤
'//  REV_003   | 2021.12.29 |日本ｺﾝﾋﾟｭｰﾀｼｽﾃﾑ　　 | 要件No18
'//            |            |                    |
'//*************************************************************************************************

''' <summary>
''' Excel入力・修正画面クラス
''' </summary>
''' <remarks></remarks>
Public MustInherit Class ExcelInputBaseClass
    Inherits ExcelProcess
    Implements IDisposable

    ''' <summary>エラーチェック種別</summary>
    Public Enum enmErorCheckType
        絶対エラー = 1
        警告エラー = 2
        範囲エラー = 3
    End Enum

    ''' <summary>コメント情報</summary>
    Public Class CommentInfo
        Public SheetName As String              'シート名
        Public Row As Integer                   '行位置
        Public Col As Integer                   '列位置
        Public Comment As String                'コメント内容
        Public ErrChkType As enmErorCheckType   'エラーチェック種別
    End Class

    ''' <summary>
    ''' セルの背景色
    ''' </summary>
    ''' <remarks></remarks>
    Public Class enmCellColor
        Public Const Red As Integer = 255
        Public Const ThinRed As Integer = 13408767
        Public Const Yellow As Integer = 65535
        Public Const Blue As Integer = 16711680
        Public Const White As Integer = 16777215
    End Class

    ''' <summary>Excelアプリケーション</summary>
    Protected Shadows WithEvents xlApp As Excel.Application

    ''' <summary>ブック</summary>
    Protected xlBook As Excel.Workbook
    ''' <summary>シートコレクション</summary>
    Protected xlSheets As Excel.Sheets

    'Excelを閉じることができるかのフラグ(ユーザーが閉じられないようにする)
    Private isClosable As Boolean

    ''' <summary>フォーム表示デリゲート</summary>
    Protected Delegate Sub FormVisibleDelegate()

    ''' <summary>ブックパス</summary>
    Private _bookPath As String
    ''' <summary>Excelアプリケーションハンドル</summary>
    Protected _hwnd As Win32WindowWrapper
    ''' <summary>Excel入力画面選択基本フォーム</summary>
    Private _Form As ExcelInputBaseForm

    ''' <summary>処理待ち画面フォーム</summary>
    Protected _WaitForm As WaitForm

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub New()
    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="frm">呼び出し元フォーム</param>
    ''' <remarks></remarks>
    Protected Sub New(frm As ExcelInputBaseForm, ByVal pFilePath As String, Optional waitFormShow As Boolean = False)
        _Form = frm

        Me.ShowExcel(pFilePath, waitFormShow)

        '呼び出し画面の非表示
        _Form.Visible = False
    End Sub

    ''' <summary>
    ''' 解放処理
    ''' </summary>
    ''' <param name="disposing"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not isDisposed Then

            'Excelを閉じられるようにする
            isClosable = True
            '問合せダイアログを非表示に設定
            xlApp.DisplayAlerts = False
            '変更無効化
            If xlBook IsNot Nothing Then
                xlBook.Saved = True
            End If

            'ショートカットキー有効化
            Me.EnableShortcutKey(True)

            'リボン表示
            xlApp.ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"",True)")

            '画面更新無効解除
            xlApp.ScreenUpdating = True

            'ウィンドウメニューボックス表示
            Me.ActSystemMenu()

            '処理待ち画面を閉じる
            Me.CloseWaitForm()

            '呼び出し元画面表示
            Me.FormVisible()

            'Sheetsの解放
            ReleaseComObject(xlSheets)
            'Workbookを閉じる
            If xlBook IsNot Nothing Then
                xlBook.Close()
            End If
            'Workbookの解放
            ReleaseComObject(xlBook)
            'Workbooksの解放
            ReleaseComObject(xlBooks)
            'ガベージコレクト強制
            GCCollect()
            'Excelの終了
            xlApp.Quit()
            'ExcelApplicationの解放
            ReleaseComObject(xlApp)
            'ガベージコレクト強制
            GCCollect()
        End If
        isDisposed = True
    End Sub

    ''' <summary>
    ''' Excel画面を開く
    ''' </summary>
    ''' <param name="pFilePath"></param>
    ''' <remarks></remarks>
    Private Sub ShowExcel(ByVal pFilePath As String, waitFormShow As Boolean)
        Try

            _bookPath = pFilePath
            isClosable = False

            'Excelアプリケーションの作成
            If xlApp Is Nothing Then
                xlApp = New Excel.Application
            End If

            'Excelアプリケーションハンドルの取得
            _hwnd = New Win32WindowWrapper(xlApp.Hwnd)

            '処理待ち画面表示
            If waitFormShow Then
                _WaitForm = New WaitForm
                _WaitForm.Show(_hwnd)
                _WaitForm.Refresh()
            End If

            'Excelアプリケーションの設定
            xlApp.Visible = True
            xlApp.UserControl = True

            'リボン非表示
            xlApp.ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"",False)")
            '画面更新無効
            xlApp.ScreenUpdating = False

            'Excelウィンドウの表示サイズを変更(最大化)
            xlApp.WindowState = Excel.XlWindowState.xlMaximized

            'ウィンドウメニューボックス削除
            Me.DelSystemMenu()

            'ショートカットキーの無効化
            Me.EnableShortcutKey(False)

            'Excelオブジェクトの設定
            xlBooks = xlApp.Workbooks
            xlBook = xlBooks.Open(_bookPath)
            xlSheets = xlBook.Worksheets
            Dim xlSheet As Excel.Worksheet = Nothing
            Try
                xlSheet = DirectCast(xlSheets.Item(1), Excel.Worksheet)
                xlSheet.Activate()
            Finally
                ReleaseComObject(xlSheet)
            End Try

        Catch ex As Exception
            If xlApp IsNot Nothing Then
                CloseExcel()
            End If
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' ExcelのWorkbookOpenイベント
    ''' </summary>
    ''' <param name="pWb"></param>
    ''' <remarks></remarks>
    Private Sub xlApp_WorkbookOpen(ByVal pWb As Microsoft.Office.Interop.Excel.Workbook) Handles xlApp.WorkbookOpen
        Try
            Dim bookFullName As String = pWb.FullName

            '入力用ブック確認
            If bookFullName <> _bookPath Then
                '強制終了
                pWb.Close()
            End If
        Catch ex As Exception
            Throw
        Finally
            ReleaseComObject(pWb)
        End Try
    End Sub

    ''' <summary>
    ''' ExcelのWorkbookBeforeCloseイベント
    ''' ユーザがボタン操作以外でExcelを閉じられないようにする
    ''' </summary>
    ''' <param name="pWb"></param>
    ''' <param name="pCancel"></param>
    ''' <remarks></remarks>
    Private Sub xlApp_WorkbookBeforeClose(ByVal pWb As Excel.Workbook, ByRef pCancel As Boolean) Handles xlApp.WorkbookBeforeClose
        Try

            If isClosable = False Then
                'ユーザがExcelを終了した場合はキャンセル
                pCancel = True
            Else
                'プログラム上から指定の場合は閉じる
                pCancel = False
            End If

        Catch ex As Exception
            Throw
        Finally
            ReleaseComObject(pWb)
        End Try
    End Sub

    ''' <summary>
    ''' Excel画面を閉じる
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overridable Sub CloseExcel()
        Try
            '解放処理
            Me.Dispose()

        Catch ex As Exception
            Throw
        End Try

    End Sub

    ''' <summary>
    ''' 呼び出し元画面表示
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overridable Sub FormVisible()
        'デリゲートの実行
        _Form.Invoke(New FormVisibleDelegate(AddressOf _Form.FormVisible))
    End Sub

    ''' <summary>
    ''' Excelファイルを保護する
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub ProtectExcel()
        Try

            For i As Integer = 1 To xlSheets.Count
                Dim xlSheet As Excel.Worksheet = DirectCast(xlSheets.Item(i), Excel.Worksheet)

                'REV_003 MOD START---
                'xlSheet.Protect()
                If Me._Form.Name = "BRA1410F" Then
                    '呼び元がBRA1410F（電子調査票）だった場合、フィルタの保護のみ外す
                    xlSheet.Protect(AllowFiltering:=True)
                Else
                    xlSheet.Protect()
                End If
                'REV_003 MOD END---

                ReleaseComObject(xlSheet)
            Next

            xlBook.Protect(Structure:=True)

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Excelファイルの保護を解除する
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub UnprotectExcel()
        Try

            xlBook.Unprotect()

            For i As Integer = 1 To xlSheets.Count
                Dim xlSheet As Excel.Worksheet = DirectCast(xlSheets.Item(i), Excel.Worksheet)
                xlSheet.Unprotect()
                ReleaseComObject(xlSheet)
            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Excel前処理
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub BeforeExcel()
        Try

            ''手動計算にする
            'xlApp.Calculation = Excel.XlCalculation.xlCalculationManual
            '画面更新を無効にする
            xlApp.ScreenUpdating = False
            'イベント発生を無効にする
            xlApp.EnableEvents = False
            'アラート表示を無効にする
            xlApp.DisplayAlerts = False
            'Excelファイルの保護を解除する
            Me.UnprotectExcel()

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

            'Excelファイルを保護する
            Me.ProtectExcel()
            'アラート表示を有効にする
            xlApp.DisplayAlerts = True
            'イベント発生を有効にする
            xlApp.EnableEvents = True
            '画面更新を有効にする
            xlApp.ScreenUpdating = True
            ''自動計算にする
            'xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' コメントを追加する
    ''' </summary>
    ''' <param name="CommentInfoList">コメント情報リスト</param>
    ''' <param name="isClear">コメント追加時にクリアするかどうか</param>
    ''' <param name="isAutoSize">コメントの自動サイズ調整するかどうか</param>
    ''' <remarks></remarks>
    Protected Sub SetCommentInfo(ByVal CommentInfoList As List(Of CommentInfo), Optional ByVal isClear As Boolean = True, Optional isAutoSize As Boolean = True)
        Try
            Dim xlSheet As Excel.Worksheet = Nothing
            Dim xlCells As Excel.Range = Nothing
            Dim xlTab As Excel.Tab = Nothing

            '重複しないシート名リストを取得
            Dim sheetNameList As List(Of String) = (From n In CommentInfoList Select n.SheetName Distinct).ToList
            For Each sheetName As String In sheetNameList
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheetName), Excel.Worksheet)
                    xlCells = DirectCast(xlSheet.Cells, Excel.Range)
                    xlTab = DirectCast(xlSheet.Tab, Excel.Tab)

                    'シート見出しの色を変更
                    xlTab.Color = enmCellColor.Red

                    'シート名が一致するコメント情報リストでループさせる
                    For Each CommentInfo In (From n In CommentInfoList Where n.SheetName = sheetName).ToList

                        Dim xlCell As Excel.Range = Nothing
                        Dim xlInterior As Excel.Interior = Nothing

                        Try

                            If (Not (CommentInfo.Row = 0 Or CommentInfo.Col = 0)) Then  'REV-002 ADD

                            xlCell = DirectCast(xlCells(CommentInfo.Row, CommentInfo.Col), Excel.Range)
                            xlInterior = DirectCast(xlCell.Interior, Excel.Interior)

                            If isClear Then
                                'コメントクリア
                                xlCell.ClearComments()
                            End If
                            'コメント追加
                            xlCell.AddComment(CommentInfo.Comment)
                            'セルの背景色を変更
                            xlInterior.Color = GetErrCellColor(CommentInfo.ErrChkType)

                            If isAutoSize Then
                                Dim xlComment As Excel.Comment = Nothing
                                Dim xlShape As Excel.Shape = Nothing
                                Dim xlTextFrame As Excel.TextFrame = Nothing
                                Try
                                    xlComment = DirectCast(xlCell.Comment, Excel.Comment)
                                    xlShape = DirectCast(xlComment.Shape, Excel.Shape)
                                    xlTextFrame = DirectCast(xlShape.TextFrame, Excel.TextFrame)
                                    '自動サイズ調整
                                    xlTextFrame.AutoSize = True
                                    'REV_001 ADD START---
                                    xlTextFrame.Characters.Font.Size = 16
                                    'REV_001 ADD END---
                                Finally
                                    ReleaseComObject(xlTextFrame)
                                    ReleaseComObject(xlShape)
                                    ReleaseComObject(xlComment)
                                End Try
                            End If
                            End If                                                      'REV-002 ADD

                        Finally
                            ReleaseComObject(xlInterior)
                            ReleaseComObject(xlCell)
                        End Try
                    Next
                Finally
                    ReleaseComObject(xlTab)
                    ReleaseComObject(xlCells)
                    ReleaseComObject(xlSheet)
                End Try
            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 処理待ち画面を閉じる
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub CloseWaitForm()
        '処理待ち画面解放
        If Not _WaitForm Is Nothing Then
            _WaitForm.Close()
            _WaitForm = Nothing
        End If
    End Sub

    ''' <summary>
    ''' エラー時のセルの色を取得する
    ''' </summary>
    ''' <param name="errorCheckType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetErrCellColor(ByVal errorCheckType As enmErorCheckType) As Integer
        Dim ret As Integer = Nothing

        Try

            Select Case (errorCheckType)
                Case enmErorCheckType.絶対エラー
                    ret = enmCellColor.Red
                Case enmErorCheckType.警告エラー
                    ret = enmCellColor.ThinRed
                    'REV_002 start--------------------
                Case enmErorCheckType.範囲エラー
                    ret = enmCellColor.Red
                    'REV_002 end--------------------
            End Select

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' ショートカットキーの有効化・無効化を設定する
    ''' </summary>
    ''' <param name="isEnable">True：有効化、False：無効化</param>
    ''' <remarks></remarks>
    Private Sub EnableShortcutKey(ByVal isEnable As Boolean)
        If isEnable Then
            'キー有効
            xlApp.OnKey("{F12}")
            xlApp.OnKey("+{F2}")
            xlApp.OnKey("+{F12}")
            xlApp.OnKey("^7")
            xlApp.OnKey("^{PGUP}")
            xlApp.OnKey("^{PGDN}")
            xlApp.OnKey("^n")
            xlApp.OnKey("^o")
            xlApp.OnKey("^p")
            xlApp.OnKey("^{F2}")
            xlApp.OnKey("^{F12}")
            xlApp.OnKey("^+{F12}")
            xlApp.OnKey("^s")
            xlApp.OnKey("^w")
            xlApp.OnKey("^{F4}")
            xlApp.OnKey("%{F4}")
            xlApp.OnKey("^+@")
        Else
            'キー無効
            xlApp.OnKey("{F12}", "")
            xlApp.OnKey("+{F2}", "")
            xlApp.OnKey("+{F12}", "")
            xlApp.OnKey("^7", "")
            xlApp.OnKey("^{PGUP}", "")
            xlApp.OnKey("^{PGDN}", "")
            xlApp.OnKey("^n", "")
            xlApp.OnKey("^o", "")
            xlApp.OnKey("^p", "")
            xlApp.OnKey("^{F2}", "")
            xlApp.OnKey("^{F12}", "")
            xlApp.OnKey("^+{F12}", "")
            xlApp.OnKey("^s", "")
            xlApp.OnKey("^w", "")
            xlApp.OnKey("^{F4}", "")
            xlApp.OnKey("%{F4}", "")
            xlApp.OnKey("^+@", "")
        End If
    End Sub

    ''' <summary>
    ''' ブックパス
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected ReadOnly Property BookPath() As String
        Get
            Return Me._bookPath
        End Get
    End Property

    ''' <summary>
    ''' Excelアプリケーションハンドル
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected ReadOnly Property Hwnd() As Win32WindowWrapper
        Get
            Return Me._hwnd
        End Get
    End Property

    ''' <summary>
    ''' ウィンドウラッパクラス
    ''' </summary>
    ''' <remarks></remarks>
    Protected Class Win32WindowWrapper
        Implements System.Windows.Forms.IWin32Window

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="windowHandle"></param>
        ''' <remarks></remarks>
        Public Sub New(windowHandle As Integer)
            Me.New(New IntPtr(windowHandle))
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="windowHandle"></param>
        ''' <remarks></remarks>
        Private Sub New(windowHandle As IntPtr)
            Me.m_handle = windowHandle
        End Sub

        Private ReadOnly m_handle As IntPtr

        ''' <summary>
        ''' ハンドル
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private ReadOnly Property Handle() As IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return Me.m_handle
            End Get
        End Property
    End Class

    ''' <summary>
    ''' ウィンドウ情報取得
    ''' </summary>
    ''' <param name="hWnd">ウィンドウのハンドル</param>
    ''' <param name="nIndex">取得する値のオフセット</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function GetWindowLong( _
        hWnd As IntPtr, _
        <MarshalAs(UnmanagedType.I4)> nIndex As Integer) As Integer
    End Function

    ''' <summary>
    ''' ウィンドウ情報設定
    ''' </summary>
    ''' <param name="hWnd">ウィンドウのハンドル</param>
    ''' <param name="nIndex">設定する値のオフセット</param>
    ''' <param name="dwNewLong">新しい値</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll")> _
    Private Shared Function SetWindowLong( _
        hWnd As IntPtr, _
        <MarshalAs(UnmanagedType.I4)> nIndex As Integer, _
        dwNewLong As IntPtr) As Integer
    End Function

    ''' <summary>
    ''' ウィンドウメニューバー再描画
    ''' </summary>
    ''' <param name="hWnd">ウィンドウのハンドル</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll")> _
    Shared Function DrawMenuBar( _
        ByVal hWnd As IntPtr) As Boolean
    End Function

    ''' <summary>ウィンドウスタイル</summary>
    Private Const WS_SYSMENU As Long = &H80000
    ''' <summary>タイトルバー上にウィンドウメニューボックスを持つウィンドウ</summary>
    Private Const GWL_STYLE As Long = -16

    ''' <summary>
    ''' ウィンドウメニューボックス削除
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DelSystemMenu()
        Dim hWnd As Long
        Dim lStyle As Long

        hWnd = xlApp.Hwnd
        lStyle = GetWindowLong(CType(hWnd, System.IntPtr), GWL_STYLE)
        lStyle = lStyle And Not WS_SYSMENU
        SetWindowLong(CType(hWnd, System.IntPtr), GWL_STYLE, CType(lStyle, System.IntPtr))
        DrawMenuBar(CType(hWnd, System.IntPtr))
    End Sub

    ''' <summary>
    ''' ウィンドウメニューボックス表示
    ''' </summary>
    ''' <remarks></remarks>
    Sub ActSystemMenu()
        Dim hWnd As Long
        Dim lStyle As Long

        hWnd = xlApp.Hwnd
        lStyle = GetWindowLong(CType(hWnd, System.IntPtr), GWL_STYLE)
        lStyle = lStyle Or WS_SYSMENU
        SetWindowLong(CType(hWnd, System.IntPtr), GWL_STYLE, CType(lStyle, System.IntPtr))
        DrawMenuBar(CType(hWnd, System.IntPtr))
    End Sub
End Class
