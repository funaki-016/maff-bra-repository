Imports Microsoft.Office.Interop


'//*************************************************************************************************
'//  修正履歴
'// ------------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前           |                修  正  内  容
'// -----------+------------+------------------------+----------------------------------------------
'//  REV_001   | 2024.05.31 |Daiko                   | 要件No.1 新規作成
'//            |            |                        |
'//*************************************************************************************************

''' <summary>
''' 農業地域類型マスタ一覧出力
''' </summary>
''' <remarks></remarks>
Public Class BRA9610R
    Inherits ExcelOutputSingleBaseClass

    ''' <summary>農業地域類型マスタ調査年</summary>
    Private _year As String

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="outPath"></param>
    ''' <param name="year"></param>
    Public Sub New(ByVal outPath As String, year As String)
        MyBase.New(ComConst.農業地域類型マスタ管理.ファイル情報.tempFileName, True, False, ComConst.農業地域類型マスタ管理.ファイル情報.reportName, outPath, False)

        _year = year

        '調査年を設定する
        Me.SetYear()
    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(xlSheets As Excel.Sheets)
        Dim dt As DataTable

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '農業地域類型マスタ調査年一覧取得
            dt = DAOOther.GetNogyoChikiMstYearList(db, _year)
        End Using

        'シートデータ設定
        ComUtil.NogyoChikiMstMainte.SetSheetData(dt, xlSheets, CType(Me, ExcelProcess))

        'フォーマット設定
        Me.SetReportFormat(xlSheets, dt.Rows.Count)

    End Sub

    ''' <summary>
    ''' 農業地域類型マスタ調査年を設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetYear()
        Dim xlSheet As Excel.Worksheet = Nothing

        Try
            xlSheet = DirectCast(xlSheets.Item(1), Excel.Worksheet)

            'シート保護設定
            Dim protect As Boolean = xlSheet.ProtectContents
            If protect Then
                xlSheet.Unprotect()
            End If

            '農業地域類型マスタ調査年設定
            SetExcelValue(xlSheet, ComConst.農業地域類型マスタ管理.ファイル情報.Row.Year, ComConst.農業地域類型マスタ管理.ファイル情報.Col.Year, _year)

            'シート保護設定
            If protect Then
                xlSheet.Protect()
            End If

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
            xlSheet = DirectCast(xlSheets.Item(ComConst.農業地域類型マスタ管理.ファイル情報.SheetName), Excel.Worksheet)

            'シート保護設定
            Dim protect As Boolean = xlSheet.ProtectContents
            If protect Then
                xlSheet.Unprotect()
            End If

            '行削除
            If dataCount < ComConst.農業地域類型マスタ管理.ファイル情報.Row.Max Then
                Me.DeleteRow(xlSheet, dataCount + ComConst.農業地域類型マスタ管理.ファイル情報.Row.First, ComConst.農業地域類型マスタ管理.ファイル情報.Row.Last)
            End If

            'シート保護設定
            If protect Then
                xlSheet.Protect()
            End If

        Finally
            ReleaseComObject(xlSheet)
        End Try

    End Sub

End Class
