Imports Microsoft.Office.Interop
Imports System.IO

''' <summary>
''' Excelプロセスクラス
''' </summary>
''' <remarks></remarks>
Public MustInherit Class ExcelProcess
    Inherits ComObjectProcess
    Implements IDisposable

    ''' <summary>Excelアプリケーション</summary>
    Protected xlApp As Excel.Application
    ''' <summary>ブックコレクション</summary>
    Protected xlBooks As Excel.Workbooks

    ''' <summary>Disposeしたかの判定用</summary>
    Protected isDisposed As Boolean

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub New()

    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="open"></param>
    ''' <remarks></remarks>
    Protected Sub New(open As Boolean)
        If open Then
            Me.OpenExcel()
        End If
    End Sub

    ''' <summary>
    ''' 終了処理
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub Finalize()
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ''' <summary>
    ''' 解放処理
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overridable Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    ''' <summary>
    ''' 解放処理
    ''' </summary>
    ''' <param name="disposing"></param>
    ''' <remarks></remarks>
    Protected Overridable Sub Dispose(disposing As Boolean)
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
            'Excelアプリケーションの作成
            If xlApp Is Nothing Then
                xlApp = New Excel.Application
            End If

            xlApp.DisplayAlerts = False
            'Excelオブジェクトの設定
            xlBooks = xlApp.Workbooks

        Catch ex As Exception
            Me.Dispose()
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Excelアプリ無効
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub DisableExcelApp()
        Try

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
    ''' Excelアプリ有効
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub EnableExcelApp()
        Try

            'アラート表示を有効にする
            xlApp.DisplayAlerts = True
            'イベント発生を有効にする
            xlApp.EnableEvents = True
            '画面更新を有効にする
            xlApp.ScreenUpdating = True

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Excelに値をセットする
    ''' </summary>
    ''' <param name="xlSht"></param>
    ''' <param name="row"></param>
    ''' <param name="col"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Protected Sub SetExcelValue(ByRef xlSht As Excel.Worksheet, ByVal row As Integer, ByVal col As Integer, ByVal value As String)
        Dim xlCel As Excel.Range = Nothing
        Try
            xlCel = xlSht.Cells

            Dim xlRng As Excel.Range = Nothing
            Try
                xlRng = DirectCast(xlCel.Item(row, col), Excel.Range)
                xlRng.Value = value
            Finally
                ReleaseComObject(xlRng)
            End Try
        Finally
            ReleaseComObject(xlCel)
        End Try
    End Sub

    ''' <summary>
    ''' Excelの行を削除する
    ''' </summary>
    ''' <param name="xlSht"></param>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Protected Sub DeleteRow(ByRef xlSht As Excel.Worksheet, ByVal row As Integer)
        Dim xlRng As Excel.Range = Nothing
        Try
            xlRng = xlSht.Range(row & ":" & row)
            xlRng.Delete()
        Finally
            ReleaseComObject(xlRng)
        End Try
    End Sub

    ''' <summary>
    ''' Excelの行を削除する
    ''' </summary>
    ''' <param name="xlSht"></param>
    ''' <param name="first"></param>
    ''' <param name="last"></param>
    ''' <remarks></remarks>
    Protected Sub DeleteRow(ByRef xlSht As Excel.Worksheet, ByVal first As Integer, ByVal last As Integer)
        Dim xlRng As Excel.Range = Nothing
        Try
            xlRng = xlSht.Range(first & ":" & last)
            xlRng.Delete()
        Finally
            ReleaseComObject(xlRng)
        End Try
    End Sub

    ''' <summary>
    ''' Excelの列を削除する
    ''' </summary>
    ''' <param name="xlSht"></param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Protected Sub DeleteExcelCol(ByRef xlSht As Excel.Worksheet, ByVal col As Integer)
        Dim xlCol As Excel.Range = Nothing
        Try
            xlCol = xlSht.Columns

            Dim xlRng As Excel.Range = Nothing
            Try
                xlRng = DirectCast(xlCol.Item(col), Excel.Range)
                xlRng.Delete()
            Finally
                ReleaseComObject(xlRng)
            End Try
        Finally
            ReleaseComObject(xlCol)
        End Try
    End Sub

    ''' <summary>
    ''' 罫線設定
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <remarks></remarks>
    Protected Sub SetBorders(rng As Excel.Range)
        Dim xlBorders As Excel.Borders = Nothing
        Dim xlBorder As Excel.Border = Nothing

        Try
            xlBorders = rng.Borders

            Dim indexs As Excel.XlBordersIndex() = {Excel.XlBordersIndex.xlEdgeLeft, _
                                                Excel.XlBordersIndex.xlEdgeTop, _
                                                Excel.XlBordersIndex.xlEdgeBottom, _
                                                Excel.XlBordersIndex.xlEdgeRight, _
                                                Excel.XlBordersIndex.xlInsideVertical, _
                                                Excel.XlBordersIndex.xlInsideHorizontal}
            For Each index As Excel.XlBordersIndex In indexs
                Try
                    xlBorder = xlBorders.Item(index)
                    With xlBorder
                        .LineStyle = Excel.XlLineStyle.xlContinuous
                        .ColorIndex = Excel.Constants.xlAutomatic
                        .TintAndShade = 0
                        .Weight = Excel.XlBorderWeight.xlThin
                    End With
                Finally
                    ReleaseComObject(xlBorder)
                End Try
            Next
        Finally
            ReleaseComObject(xlBorders)
        End Try
    End Sub

    ''' <summary>
    ''' 行の高さを自動調整
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <remarks></remarks>
    Public Sub SetAutoFit(rng As Excel.Range)
        Dim rngRows As Excel.Range = Nothing
        Try
            rngRows = rng.Rows
            rngRows.AutoFit()
        Finally
            ReleaseComObject(rngRows)
        End Try
    End Sub
End Class
