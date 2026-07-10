Imports Microsoft.Office.Interop
Imports System.IO

Public Class ExcelDirImportBaseForm
    ''' <summary>Excel最大行数</summary>
    Protected Const ExcelMaxRow As Integer = 1048576

    ''' <summary>インポートファイル情報</summary>
    Protected Class ImportFileInfo
        ''' <summary>ファイル名称</summary>
        Public Shared FileName As String
        ''' <summary>データ先頭行</summary>
        Public Shared DataFirstRow As Integer
        ''' <summary>一度に読み込む最大データ件数</summary>
        Public Shared MaxRowCount As Integer = 100000
        ''' <summary>先頭項目</summary>
        Public Class FirstField
            Public Shared Row As Integer
            Public Shared Col As String
            Public Shared Name As String
        End Class
        ''' <summary>最終項目</summary>
        Public Class LastField
            Public Shared Row As Integer
            Public Shared Col As String
            Public Shared Name As String
        End Class
    End Class

    ''' <summary>値情報</summary>
    Protected Class ValueInfor
        Public type As SqlDbType
        Public value As String
    End Class

    ''' <summary>Excelアプリケーション</summary>
    Protected xlApp As Excel.Application
    ''' <summary>ブックコレクション</summary>
    Protected xlBooks As Excel.Workbooks
    ''' <summary>ブック</summary>
    Protected xlBook As Excel.Workbook
    ''' <summary>シートコレクション</summary>
    Protected xlSheets As Excel.Sheets
    ''' <summary>シート</summary>
    Protected xlSheet As Excel.Worksheet

    ''' <summary>
    ''' 画面起動
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ExcelImportBaseClass_Load(sender As Object, e As System.EventArgs) Handles Me.Load
    End Sub


    ''' <summary>
    ''' Excelファイルオープン
    ''' </summary>
    ''' <param name="file">Excelファイルフルパス</param>
    ''' <remarks></remarks>
    Protected Sub OpenExcel(ByVal file As String)
        'Excelアプリケーションの作成
        If xlApp Is Nothing Then
            xlApp = New Excel.Application
        End If

        'Excelオブジェクトの設定
        xlBooks = xlApp.Workbooks
        xlBook = xlBooks.Open(file)
        xlSheets = xlBook.Worksheets
        xlSheet = CType(xlSheets.Item(1), Excel.Worksheet)
        CType(xlSheet, Excel._Worksheet).Activate()
    End Sub

    ''' <summary>
    ''' Excelファイルクローズ
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub ExcelQuit()
        ReleaseComObject(CType(xlSheet, Object))
        ReleaseComObject(CType(xlSheets, Object))
        'Workbookを閉じる
        If xlBook IsNot Nothing Then
            xlBook.Close(SaveChanges:=False)
        End If
        ReleaseComObject(CType(xlBook, Object))
        ReleaseComObject(CType(xlBooks, Object))
        'Excelの終了
        xlApp.Quit()
        ReleaseComObject(CType(xlApp, Object))
        xlApp = Nothing
    End Sub

    ''' <summary>
    ''' COMオブジェクトを解放する
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <remarks></remarks>
    Protected Sub ReleaseComObject(ByRef obj As Object)
        Try
            If obj Is Nothing Then
                Return
            End If
            If System.Runtime.InteropServices.Marshal.IsComObject(obj) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
        Finally
            obj = Nothing
        End Try
    End Sub


    ''' <summary>
    ''' ファイル整合性チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function CheckConsistencyFile() As Boolean
        Dim ret As Boolean = False

        Dim rngFirst As Excel.Range = xlSheet.Range(ImportFileInfo.FirstField.Col & ImportFileInfo.FirstField.Row)
        Dim rngLast As Excel.Range = xlSheet.Range(ImportFileInfo.LastField.Col & ImportFileInfo.LastField.Row)

        If rngFirst.Value.ToString <> ImportFileInfo.FirstField.Name _
            Or rngLast.Value.ToString <> ImportFileInfo.LastField.Name Then
            ret = False
        Else
            ret = True
        End If

        ReleaseComObject(CType(rngLast, Object))
        ReleaseComObject(CType(rngFirst, Object))

        Return ret
    End Function

    ''' <summary>
    ''' データ存在チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function CheckExitData() As Boolean
        Dim ret As Boolean = False

        Dim rngFirst As Excel.Range = xlSheet.Range(ImportFileInfo.FirstField.Col & ExcelMaxRow)
        Dim rngEnd As Excel.Range = rngFirst.End(Excel.XlDirection.xlUp)

        If rngEnd.Row < ImportFileInfo.DataFirstRow Then
            ret = False
        Else
            ret = True
        End If

        ReleaseComObject(CType(rngEnd, Object))
        ReleaseComObject(CType(rngFirst, Object))

        Return ret
    End Function
End Class