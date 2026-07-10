Imports Microsoft.Office.Interop

''' <summary>
''' 毎月勤労統計の月別データ出力
''' </summary>
''' <remarks></remarks>
Public Class BRA8110R
    Inherits ExcelOutputSingleBaseClass

    ''' <summary>事務所</summary>
    Private _jimusho As String

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="jimusho"></param>
    ''' <param name="outDir"></param>
    ''' <remarks></remarks>
    Public Sub New(jimusho As String, ByVal outPath As String)
        MyBase.New(ComConst.毎月勤労統計.出力用ファイル名称.tempFileName, True, False, ComConst.毎月勤労統計.出力用ファイル名称.reportName, outPath, True)

        _jimusho = jimusho
    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(xlSheets As Excel.Sheets)

        Dim dt As DataTable

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '毎月勤労統計取得
            dt = DAOOther.GetMaitsukiKinrouToukei(db, _jimusho)
        End Using

        'シートデータ設定
        ComUtil.MaitsukiKinrouToukei.SetSheetData(dt, _jimusho, xlSheets, CType(Me, ComObjectProcess))

        '帳票フォーマット設定
        Me.SetReportFormat(xlSheets, dt.Rows.Count)
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
            xlSheet = DirectCast(xlSheets.Item(ComConst.毎月勤労統計.出力用ファイル名称.SheetName), Excel.Worksheet)

            xlSheet.Unprotect()

            '行削除
            If dataCount < ComConst.毎月勤労統計.出力用ファイル名称.Row.Max Then
                Me.DeleteRow(xlSheet, dataCount + ComConst.毎月勤労統計.出力用ファイル名称.Row.First, ComConst.毎月勤労統計.出力用ファイル名称.Row.Last)
            End If

            xlSheet.Protect()
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub
End Class
