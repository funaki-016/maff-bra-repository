Imports Microsoft.Office.Interop

''' <summary>
''' 個別結果表審査論理（範囲）出力
''' </summary>
''' <remarks></remarks>
Public Class BRA3810R
    Inherits ExcelOutputSingleBaseClass

    ''' <summary>ヘッダータイトル文字列</summary>
    Private Const HEADER_TITLE As String = "個別結果表審査論理入力・修正（範囲）"

    ' REV_001↓
    ''' <summary>バージョン区分</summary>
    Private _versionKbn As String
    ' REV_001↓

    ' REV_001↓
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="outPath"></param>
    ''' <param name="versionKbn"></param>
    'Public Sub New(ByVal outPath As String)
    Public Sub New(ByVal outPath As String, versionKbn As String)
        ' REV_001↑
        MyBase.New(ComConst.個別結果表審査論理範囲.出力用ファイル名称.tempFileName, True, False, ComConst.個別結果表審査論理範囲.出力用ファイル名称.reportName, outPath, True)

        ' REV_001↓
        _versionKbn = versionKbn
        ' REV_001↑

        'ヘッダーを設定する
        Me.SetHeader()
    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(xlSheets As Excel.Sheets)

        Dim dt As DataTable

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '個別結果表審査論理範囲取得
            ' REV_001↓
            'dt = DAOOther.GetKobetsuKekkahyoShinsaRonriRange(db)
            dt = DAOOther.GetKobetsuKekkahyoShinsaRonriRange(db, _versionKbn)
            ' REV_001↑
        End Using

        'シートデータ設定
        ComUtil.KobetsuKekkahyoShinsaRonriRange.SetSheetData(dt, xlSheets, CType(Me, ComObjectProcess))

        '帳票フォーマット設定
        Me.SetReportFormat(xlSheets, dt.Rows.Count)
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
            SetExcelValue(xlSheet, 2, 2, HEADER_TITLE)

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
            xlSheet = DirectCast(xlSheets.Item(ComConst.個別結果表審査論理範囲.出力用ファイル名称.SheetName), Excel.Worksheet)

            xlSheet.Unprotect()

            '行削除
            If dataCount < ComConst.個別結果表審査論理範囲.出力用ファイル名称.Row.Max Then
                Me.DeleteRow(xlSheet, dataCount + ComConst.個別結果表審査論理範囲.出力用ファイル名称.Row.First, ComConst.個別結果表審査論理範囲.出力用ファイル名称.Row.Last)
            End If

            xlSheet.Protect()
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub
End Class
