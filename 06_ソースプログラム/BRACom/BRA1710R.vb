Imports Microsoft.Office.Interop

''' <summary>
''' 調査票審査論理（基本・追加）出力
''' </summary>
''' <remarks></remarks>
Public Class BRA1710R
    Inherits ExcelOutputSingleBaseClass

    ''' <summary>エラーチェック種別</summary>
    Private _errType As ComConst.エラーチェック種別.enm
    ''' <summary>ヘッダータイトル文字列</summary>
    Private Const HEADER_TITLE As String = "調査票審査論理入力・修正（{0}）"

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="errType"></param>
    ''' <param name="outPath"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal errType As ComConst.エラーチェック種別.enm, ByVal outPath As String)
        MyBase.New(ComConst.調査票審査論理.出力用ファイル名称.tempFileName, True, False, ComConst.調査票審査論理.出力用ファイル名称.reportName, outPath, True)

        _errType = errType

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
            '調査票審査論理
            dt = DAOOther.GetChosahyoShinsaRonri(db, _errType)
        End Using

        'シートデータ設定
        ComUtil.ChosahyoShinsaRonri.SetSheetData(dt, xlSheets, CType(Me, ExcelProcess))

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
            SetExcelValue(xlSheet, 2, 2, (String.Format(HEADER_TITLE, ComConst.エラーチェック種別.リスト(_errType))))

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
            xlSheet = DirectCast(xlSheets.Item(ComConst.調査票審査論理.出力用ファイル名称.SheetName), Excel.Worksheet)

            xlSheet.Unprotect()

            '行削除
            If dataCount < ComConst.調査票審査論理.出力用ファイル名称.Row.Max Then
                Me.DeleteRow(xlSheet, dataCount + ComConst.調査票審査論理.出力用ファイル名称.Row.First, ComConst.調査票審査論理.出力用ファイル名称.Row.Last)
            End If

            xlSheet.Protect()
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub
End Class
