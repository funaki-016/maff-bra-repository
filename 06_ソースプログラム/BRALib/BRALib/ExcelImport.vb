'------------------------------------------------------------------------------------------
'| REV | 変更年月日 | 変更者             | 変更内容
'------------------------------------------------------------------------------------------
'| 001 | 2019/05/15 | Daiko              | 新規作成
'------------------------------------------------------------------------------------------
Imports Microsoft.Office.Interop
Imports System.IO

''' <summary>
''' Excelファイル取込クラス
''' </summary>
''' <remarks></remarks>
Public Class ExcelImport
    Inherits ComObjectProcess
    Implements IDisposable

    ''' <summary>Excel最大行数</summary>
    Public Const ExcelMaxRow As Integer = 1048576

    ''' <summary>項目情報</summary>
    Public Class Field
        Public Row As Integer
        Public Col As Integer
        Public Name As String
    End Class

    ''' <summary>インポートファイル情報</summary>
    Public Class ImportFileInfo
        ''' <summary>データ先頭行</summary>
        Public DataFirstRow As Integer
        ''' <summary>先頭項目</summary>
        Public FirstField As Field
        ''' <summary>最終項目</summary>
        Public LastField As Field
    End Class

    ''' <summary>Excelアプリケーション</summary>
    Private xlApp As Excel.Application
    ''' <summary>Workbooksオブジェクト</summary>
    Private xlBooks As Excel.Workbooks
    ''' <summary>Workbookオブジェクト</summary>
    Private xlBook As Excel.Workbook
    ''' <summary>Sheetsオブジェクト</summary>
    Public xlSheets As Excel.Sheets
    ''' <summary>Worksheetオブジェクト</summary>
    Public xlSheet As Excel.Worksheet

    ''' <summary>インポートファイル情報</summary>
    Public ImportFile As ImportFileInfo

    ''' <summary>最終行</summary>
    Private xlLastRow As Integer

    ''' <summary>選択されたファイルのフルパス</summary>
    Public filePath As String

    ''' <summary>Disposeしたかの判定用</summary>
    Private isDisposed As Boolean

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="impFileInf"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal impFileInf As ImportFileInfo)
        ImportFile = impFileInf
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
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    ''' <summary>
    ''' 解放処理
    ''' </summary>
    ''' <param name="disposing"></param>
    ''' <remarks></remarks>
    Public Overridable Sub Dispose(disposing As Boolean)
        If Not isDisposed Then
            If xlApp IsNot Nothing Then
                'Sheetの解放
                If xlSheet IsNot Nothing Then
                    ReleaseComObject(CObj(xlSheet))
                End If
                'Sheetsの解放
                ReleaseComObject(CObj(xlSheets))
                'Workbookを閉じる
                If xlBook IsNot Nothing Then
                    xlBook.Saved = True
                    xlBook.Close()
                End If
                'Workbookの解放
                ReleaseComObject(CObj(xlBook))
                'Workbooksの解放
                ReleaseComObject(CObj(xlBooks))
                xlApp.DisplayAlerts = True
                'Excelの終了
                xlApp.Quit()
                'ExcelApplicationの解放
                ReleaseComObject(CObj(xlApp))
            End If

        End If
        isDisposed = True
    End Sub

    ''' <summary>
    ''' ファイル選択ダイアログを表示する
    ''' </summary>
    ''' <param name="openFileDlg"></param>
    ''' <param name="form"></param>
    ''' <remarks></remarks>
    Public Sub OpenFileDialog(ByVal openFileDlg As OpenFileDialog, ByVal form As Form)
        Dim ret As String = Nothing

        With openFileDlg
            'ダイアログのタイトルを設定する
            .Title = "ファイル選択"
            'ファイルのフィルタを設定する
            .Filter = "Microsoft Excelブック(*.xls;*.xlsx)|*.xls;*.xlsx|すべてのファイル(*.*)|*.*"
            'ファイルの種類の初期設定を1番目に設定する
            .FilterIndex = 1
            '初期表示するファイル名を設定する
            .FileName = ""
            '初期表示するディレクトリを設定する
            .InitialDirectory = IniFileInfo.ExcelInPath
            'ディレクトリが存在しなければ作成する
            If Not Directory.Exists(.InitialDirectory) Then
                Try
                    Directory.CreateDirectory(.InitialDirectory)
                Catch ex As Exception
                End Try
            End If

            If System.Windows.Forms.DialogResult.OK.Equals(.ShowDialog(form)) Then
                filePath = .FileName
            End If
        End With

    End Sub

    ''' <summary>
    ''' Excelファイルオープン
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub OpenExcel()
        Try
            'Excelアプリケーションの作成
            If xlApp Is Nothing Then
                xlApp = New Excel.Application
            End If

            xlApp.DisplayAlerts = False
            'Excelオブジェクトの設定
            xlBooks = xlApp.Workbooks
            xlBook = xlBooks.Open(filePath)
            xlSheets = xlBook.Worksheets
            xlSheet = DirectCast(xlSheets.Item(1), Excel.Worksheet)

        Catch ex As Exception
            Me.Dispose()
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' ファイル整合性チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckConsistencyFile() As Boolean
        Dim ret As Boolean = False

        Dim xlRngFirst As Excel.Range = Nothing
        Dim xlRngLast As Excel.Range = Nothing
        Try
            xlRngFirst = xlSheet.Range(ConvertColNoToLetter(ImportFile.FirstField.Col) & ImportFile.FirstField.Row)
            xlRngLast = xlSheet.Range(ConvertColNoToLetter(ImportFile.LastField.Col) & ImportFile.LastField.Row)

            If CStr(xlRngFirst.Value) = ImportFile.FirstField.Name AndAlso
               CStr(xlRngLast.Value) = ImportFile.LastField.Name Then
                ret = True
            End If
        Finally
            ReleaseComObject(CObj(xlRngFirst))
            ReleaseComObject(CObj(xlRngLast))
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' データ存在チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckExistData() As Boolean
        Dim ret As Boolean = False

        Dim xlRngMaxRow As Excel.Range = Nothing
        Dim xlRngEnd As Excel.Range = Nothing
        Try
            xlRngMaxRow = xlSheet.Range(ConvertColNoToLetter(ImportFile.FirstField.Col) & ExcelMaxRow)
            xlRngEnd = xlRngMaxRow.End(Excel.XlDirection.xlUp)

            xlLastRow = xlRngEnd.Row

            If xlRngEnd.Row >= ImportFile.DataFirstRow Then
                ret = True
            End If
        Finally
            ReleaseComObject(CObj(xlRngMaxRow))
            ReleaseComObject(CObj(xlRngEnd))
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 値一致チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckValue(ByVal row As Integer, ByVal col As Integer, ByVal value As String, Optional ByVal isJogaiSpace As Boolean = False) As Boolean
        Dim ret As Boolean = False

        Dim xlRange As Excel.Range = Nothing
        Try
            xlRange = xlSheet.Range(ConvertColNoToLetter(col) & row)

            If String.IsNullOrEmpty(CStr(xlRange.Value)) AndAlso String.IsNullOrEmpty(value) Then
                ret = True
            ElseIf xlRange.Value IsNot Nothing AndAlso If(isJogaiSpace, Me.DelSpace(CStr(xlRange.Value)), CStr(xlRange.Value)) = value Then
                ret = True
            End If
        Finally
            ReleaseComObject(CObj(xlRange))
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' Excelのデータを取得する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetExcelData() As Object(,)
        Dim ret(,) As Object = Nothing
        Dim rngFirst As Excel.Range = Nothing
        Dim rngLast As Excel.Range = Nothing
        Dim rngAll As Excel.Range = Nothing

        Try
            rngFirst = xlSheet.Range(ConvertColNoToLetter(ImportFile.FirstField.Col) & ImportFile.DataFirstRow)
            rngLast = xlSheet.Range(ConvertColNoToLetter(ImportFile.LastField.Col) & xlLastRow)
            rngAll = xlSheet.Range(rngFirst, rngLast)

            'セル値配列取得
            ret = DirectCast(rngAll.Value, Object(,))
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rngFirst)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rngLast)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rngAll)
            rngFirst = Nothing
            rngLast = Nothing
            rngAll = Nothing
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' Excelの列番号を列英字に変換する
    ''' </summary>
    ''' <param name="col"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConvertColNoToLetter(ByVal col As Integer) As String
        Dim ret As String = Nothing
        Dim alpha As Integer
        Dim remainder As Integer
        alpha = CInt(Math.Truncate((col - 1) / 26))
        remainder = col - (alpha * 26)
        If alpha > 0 Then
            ret = Chr(alpha + 64)
        End If
        If remainder > 0 Then
            ret = ret & Chr(remainder + 64)
        End If
        Return ret
    End Function

    ''' <summary>
    ''' Excelの列英字を列番号に変換する
    ''' </summary>
    ''' <param name="colStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConvertLetterToColNo(ByVal colStr As String) As Integer
        Dim ret As Integer = 0
        For i As Integer = 1 To Len(colStr)
            ret = ret * 26 + (Asc(UCase(Mid(colStr, i, 1))) - 64)
        Next
        Return ret
    End Function

    ''' <summary>
    ''' 半角空白及び全角空白を削除した文字列を返す
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DelSpace(ByVal value As String) As String
        Return value.Replace(" ", "").Replace("　", "")
    End Function
End Class
