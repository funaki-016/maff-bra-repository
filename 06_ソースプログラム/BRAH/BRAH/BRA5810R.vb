Imports Microsoft.Office.Interop

''' <summary>
''' 欠測値適用状況一覧表出力
''' </summary>
''' <remarks></remarks>
Public Class BRA5810R
    Inherits ExcelOutputSingleBaseClass

    ''' <summary>調査年</summary>
    Private ChosaNen As String = String.Empty

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="outPath"></param>
    ''' <param name="chosaNen"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal outPath As String, ByVal chosaNen As String)

        MyBase.New(ComConst.欠測値.適用状況一覧表出力用ファイル.Report.tempFileName, True, False, ComConst.欠測値.適用状況一覧表出力用ファイル.Report.reportName, outPath, True)

        Me.ChosaNen = chosaNen

        'ヘッダーを設定する
        Me.SetHeader()

    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(xlSheets As Excel.Sheets)

        '【処理詳細仕様 ②】帳票出力データを検索する
        Dim rowList = ComUtil.GetAverageResultTableRowList(ChosaNen, orderBy:="ORDER BY センサス番号")

        'シートデータ設定
        ComUtil.SetSheetData(rowList, xlSheets, CType(Me, ExcelProcess))

        '帳票フォーマット設定
        SetReportFormat(xlSheets, rowList.Count)

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

            SetExcelValue(xlSheet, 1, 1, ChosaNen)

            xlSheet.Protect()
        Catch ex As Exception
            MyBase.Dispose()
            Throw
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub

    ''' <summary>
    ''' 帳票フォーマット設定
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <param name="dataCount"></param>
    ''' <remarks></remarks>
    Private Sub SetReportFormat(xlSheets As Excel.Sheets, dataCount As Integer)
        Dim xlSheet As Excel.Worksheet = Nothing

        Try
            'シートの設定
            xlSheet = DirectCast(xlSheets.Item(ComConst.欠測値.適用状況一覧表出力用ファイル.Report.SheetName), Excel.Worksheet)

            xlSheet.Unprotect()

            '行削除
            If dataCount < ComConst.欠測値.適用状況一覧表出力用ファイル.Row.Max Then
                Me.DeleteRow(xlSheet, dataCount + ComConst.欠測値.適用状況一覧表出力用ファイル.Row.First, ComConst.欠測値.適用状況一覧表出力用ファイル.Row.Last)
            End If

            '枠線を描画
            If dataCount < ComConst.欠測値.適用状況一覧表出力用ファイル.Row.Max Then
                Me.SetBorders(xlSheet, dataCount + ComConst.欠測値.適用状況一覧表出力用ファイル.Row.First)
            End If

            xlSheet.Protect()
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub

    ''' <summary>
    ''' 罫線設定
    ''' </summary>
    ''' <param name="xlSht"></param>
    ''' <param name="first"></param>
    ''' <remarks></remarks>
    Private Overloads Sub SetBorders(ByRef xlSht As Excel.Worksheet, ByVal first As Integer)
        Dim xlRng As Excel.Range = Nothing
        Dim xlRngBorders As Excel.Borders = Nothing
        Dim xlRngBorder As Excel.Border = Nothing
        Try
            xlRng = xlSht.Range("A" & first & ":" & "I" & first)
            xlRngBorders = xlRng.Borders
            xlRngBorder = xlRngBorders.Item(Excel.XlBordersIndex.xlEdgeTop)

            xlRngBorder.LineStyle = Excel.XlLineStyle.xlContinuous
            xlRngBorder.Weight = Excel.XlBorderWeight.xlMedium
        Finally
            ReleaseComObject(xlRng)
            ReleaseComObject(xlRngBorder)
            ReleaseComObject(xlRngBorders)
        End Try
    End Sub

End Class
