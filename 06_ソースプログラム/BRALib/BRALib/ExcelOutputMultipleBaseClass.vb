'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2023.10.30 |大興電子通信        | 変更要件No.6
'//            |            |                    |
'//*************************************************************************************************

Imports Microsoft.Office.Interop
Imports System.IO

''' <summary>
''' Excel出力クラス（複数ブック出力用）
''' </summary>
''' <remarks></remarks>
Public MustInherit Class ExcelOutputMultipleBaseClass
    Inherits ExcelOutputBaseClass
    Implements IDisposable

    ''' <summary>帳票テンプレートファイル名</summary>
    Private _templateFile As String

    ''' <summary>
    ''' 読取専用で開くか否か(REV_001)
    ''' </summary>
    Protected IsReadOnly As Boolean = False

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
    Protected Sub New(templateFile As String, excelCheck As Boolean, reportCheck As Boolean, ByVal title As String, ByVal outDir As String, ByVal removalMacro As Boolean)
        MyBase.New(excelCheck, reportCheck, title, outDir, removalMacro)

        _templateFile = templateFile

        Me.OpenExcel()
    End Sub

    ''' <summary>
    ''' 解放処理
    ''' </summary>
    ''' <param name="disposing"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not isDisposed Then
            If xlApp IsNot Nothing Then
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

        Catch ex As Exception
            MyBase.Dispose()
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 処理実行
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="lst"></param>
    ''' <param name="pgdParent"></param>
    ''' <param name="pgdCount"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function Execute(Of T As Class)(lst As List(Of T), pgdParent As IWin32Window, pgdCount As Integer) As enmOutputReturn
        Return Me.Execute(lst, pgdParent, pgdCount, Nothing)
    End Function

    ''' <summary>
    ''' 処理実行
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="lst"></param>
    ''' <param name="overWriteMessageID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function Execute(Of T As Class)(lst As List(Of T), overWriteMessageID As String) As enmOutputReturn
        Return Me.Execute(lst, Nothing, 0, overWriteMessageID)
    End Function

    ''' <summary>
    ''' 処理実行
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="lst"></param>
    ''' <param name="pgdParent"></param>
    ''' <param name="pgdCount"></param>
    ''' <param name="overWriteMessageID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function Execute(Of T As Class)(lst As List(Of T), Optional pgdParent As IWin32Window = Nothing, Optional pgdCount As Integer = 0, Optional overWriteMessageID As String = Nothing) As enmOutputReturn
        Dim ret As enmOutputReturn = enmOutputReturn.CANCEL

        Dim outPut As enmOutputReturn = Nothing

        Dim xlBook As Excel.Workbook = Nothing
        Dim xlSheets As Excel.Sheets = Nothing

        Try
            'Excelアプリ無効
            Me.DisableExcelApp()

            If Not pgdParent Is Nothing Then
                ProgressDialog = New ProgressDialog()
                '進捗ダイアログを表示する
                ProgressDialog.Maximum = lst.Count * 2 + lst.Count * pgdCount
                ProgressDialog.Show(pgdParent)
            End If

            Try

                For Each unit As T In lst
                    Try
                        '進捗加増
                        Me.ProgressAddValue = 1

                        'Workbookを開く
                        'REV_001↓調査票テンプレートは読取専用で開く
                        'xlBook = xlBooks.Open(IniFileInfo.ExcelReportPath() & Path.DirectorySeparatorChar & _templateFile)
                        xlBook = xlBooks.Open(IniFileInfo.ExcelReportPath() & Path.DirectorySeparatorChar & _templateFile, ReadOnly:=IsReadOnly)
                        'REV_001↑

                        Try
                            xlSheets = xlBook.Worksheets

                            '帳票編集
                            Me.ReportEdit(xlSheets, unit)
                            '出力
                            outPut = Me.OutPut(xlBook, overWriteMessageID)
                            Select Case outPut
                                Case enmOutputReturn.OK
                                    ret = outPut
                                Case enmOutputReturn.ERR_SAVEAS
                                    Throw New SaveAsException()
                            End Select

                            '進捗加増
                            Me.ProgressAddValue = 1
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
    ''' <typeparam name="T"></typeparam>
    ''' <param name="xlSheets"></param>
    ''' <param name="unit"></param>
    ''' <remarks></remarks>
    Protected MustOverride Sub ReportEdit(Of T As Class)(xlSheets As Excel.Sheets, unit As T)

End Class
