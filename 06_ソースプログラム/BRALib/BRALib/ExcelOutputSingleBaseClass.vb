Imports Microsoft.Office.Interop
Imports System.IO

''' <summary>
''' Excel出力クラス（単一ブック出力用）
''' </summary>
''' <remarks></remarks>
Public MustInherit Class ExcelOutputSingleBaseClass
    Inherits ExcelOutputBaseClass
    Implements IDisposable

    ''' <summary>Workbookオブジェクト</summary>
    Protected xlBook As Excel.Workbook
    ''' <summary>Sheetsオブジェクト</summary>
    Protected xlSheets As Excel.Sheets

    ''' <summary>帳票テンプレートファイル名</summary>
    Private _templateFile As String

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub New()

    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub New(templateFile As String, excelCheck As Boolean, reportCheck As Boolean, ByVal title As String, ByVal outPath As String, ByVal removalMacro As Boolean)
        MyBase.New(excelCheck, reportCheck, title, Path.GetDirectoryName(outPath), outPath, removalMacro)

        _templateFile = templateFile

        Me.OpenExcel()
    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub New(templateFile As String, excelCheck As Boolean, reportCheck As Boolean, ByVal title As String, ByVal outPath As String, ByVal removalMacro As Boolean, ByVal openExcel As Boolean)
        MyBase.New(excelCheck, reportCheck, title, Path.GetDirectoryName(outPath), outPath, removalMacro)

        _templateFile = templateFile

        If openExcel Then
            Me.OpenExcel()
        End If
    End Sub

    ''' <summary>
    ''' 解放処理
    ''' </summary>
    ''' <param name="disposing"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not isDisposed Then
            If xlApp IsNot Nothing Then
                'Sheetsの解放
                ReleaseComObject(CObj(xlSheets))
                'Workbookを閉じる
                If xlBook IsNot Nothing Then
                    xlBook.Saved = True
                    xlBook.Close()
                End If
                'Workbooksの解放
                ReleaseComObject(xlBooks)
                xlApp.DisplayAlerts = True
                'ガベージコレクト強制
                GCCollect()
                'Excelの終了
                xlApp.Quit()
                'ExcelApplicationの解放
                ReleaseComObject(xlApp)
                'ガベージコレクト強制
                GCCollect()

                'ストップウォッチを止める
                sw.Stop()

                'ユーザログ出力
                OutputLog.WriteUserLog(OutputLog.LogLevel.Info, processName, OutputLog.LogType.SyoriEnd, sw.Elapsed.TotalSeconds.ToString("#0"))
            End If
        End If
        isDisposed = True
    End Sub

    ''' <summary>
    ''' Excelアプリケーションオープン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OpenExcel()
        Try
            'ストップウォッチを開始する
            sw.Start()

            'ユーザログ出力
            OutputLog.WriteUserLog(OutputLog.LogLevel.Info, processName, OutputLog.LogType.SyoriStart)

            'Excelアプリケーションの作成
            If xlApp Is Nothing Then
                xlApp = New Excel.Application
            End If

            xlApp.DisplayAlerts = False
            'Excelオブジェクトの設定
            xlBooks = xlApp.Workbooks
            xlBook = xlBooks.Open(IniFileInfo.ExcelReportPath() & Path.DirectorySeparatorChar & _templateFile)
            xlSheets = xlBook.Worksheets

        Catch ex As Exception
            MyBase.Dispose()
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 処理実行
    ''' </summary>
    ''' <param name="pgd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function Execute(overWriteMessageID As String, pgd As ProgressDialog) As enmOutputReturn
        Return Me.Execute(Nothing, Nothing, overWriteMessageID, pgd)
    End Function

    ''' <summary>
    ''' 処理実行
    ''' </summary>
    ''' <param name="pgdParent"></param>
    ''' <param name="pgdCount"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function Execute(pgdParent As IWin32Window, pgdCount As Integer) As enmOutputReturn
        Return Me.Execute(pgdParent, pgdCount, Nothing)
    End Function

    ''' <summary>
    ''' 処理実行
    ''' </summary>
    ''' <param name="overWriteMessageID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function Execute(overWriteMessageID As String) As enmOutputReturn
        Return Me.Execute(Nothing, 0, overWriteMessageID)
    End Function

    ''' <summary>
    ''' 処理実行
    ''' </summary>
    ''' <remarks></remarks>
    Public Overloads Function Execute(Optional pgdParent As IWin32Window = Nothing, Optional pgdCount As Integer = 0, Optional overWriteMessageID As String = Nothing, Optional pgd As ProgressDialog = Nothing) As enmOutputReturn
        Dim ret As enmOutputReturn = Nothing
        Try
            'Excelアプリ無効
            Me.DisableExcelApp()

            If Not pgd Is Nothing Then
                ProgressDialog = pgd
            End If
            If Not pgdParent Is Nothing Then
                ProgressDialog = New ProgressDialog()
                '進捗ダイアログを表示する
                ProgressDialog.Maximum = 2 + pgdCount
                ProgressDialog.Show(pgdParent)
            End If

            Try

                '帳票編集
                Me.ReportEdit(xlSheets)

                If Not pgdParent Is Nothing Then
                    '進捗加増
                    Me.ProgressAddValue = 1
                End If

                '出力
                ret = Me.OutPut(xlBook, overWriteMessageID)
                If ret = enmOutputReturn.ERR_SAVEAS Then
                    Throw New SaveAsException()
                End If

                If Not pgdParent Is Nothing Then
                    '進捗加増
                    Me.ProgressAddValue = 1
                End If

            Finally
                If Not pgdParent Is Nothing Then
                    '進捗ダイアログを閉じる
                    ProgressDialog.endDispose()
                    ProgressDialog = Nothing
                End If
            End Try
        Catch ex As Exception
            Throw ex
        Finally
            'Excelアプリ有効
            Me.EnableExcelApp()
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected MustOverride Sub ReportEdit(xlSheets As Excel.Sheets)

End Class
