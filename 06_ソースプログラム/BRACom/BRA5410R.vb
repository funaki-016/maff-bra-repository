Imports Microsoft.Office.Interop

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.12.15 |Daiko               | 要件No4 バージョン区分追加
'//  REV_002   | 2022.12.20 |Daiko               | 要件No15 集計結果検討表（報告用）追加
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 集計結果検討表作成論理出力
''' </summary>
''' <remarks></remarks>
Public Class BRA5410R
    Inherits ExcelOutputSingleBaseClass

    ''' <summary>営農経営体区分</summary>
    Private _einouKeieitai As String
    ''' <summary>調査区分</summary>
    Private _chosakubun As String

    ' REV_002↓
    ''' <summary>ヘッダータイトル文字列</summary>
    'Private Const HEADER_TITLE As String = "集計結果検討表作成論理入力・修正"
    Private Shared HEADER_TITLE As New Dictionary(Of Integer, String) From {
        {ComConst.集計結果検討表作成論理.論理種別.集計結果検討表, "集計結果検討表作成論理入力・修正"} _
      , {ComConst.集計結果検討表作成論理.論理種別.集計結果検討表_報告用, "集計結果検討表（報告用）作成論理入力・修正"}
    }
    ' REV_002↑

    ' REV_001 REV_002↓
    ''' <summary>バージョン区分</summary>
    Private _versionKbn As String
    ''' <summary>論理種別</summary>
    Private _logicType As ComConst.集計結果検討表作成論理.論理種別
    ' REV_001 REV_002↑

    ' REV_001 REV_002↓
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="outPath"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="einouKeieitai"></param>
    ''' <param name="logicType"></param>
    'Public Sub New(ByVal outPath As String, einouKeieitai As String)
    Public Sub New(ByVal outPath As String, versionKbn As String, einouKeieitai As String, logicType As ComConst.集計結果検討表作成論理.論理種別)
        'MyBase.New(ComConst.集計結果検討表作成論理.出力用ファイル名称.tempFileName, True, False, ComConst.集計結果検討表作成論理.出力用ファイル名称.reportName, outPath, True)
        MyBase.New(ComConst.集計結果検討表作成論理.出力用ファイル名称.tempFileName(logicType), True, False, ComConst.集計結果検討表作成論理.出力用ファイル名称.reportName(logicType), outPath, True)
        ' REV_001 REV_002↑

        ' REV_001 REV_002↓
        _versionKbn = versionKbn
        _logicType = logicType
        ' REV_001 REV_002↑
        _einouKeieitai = einouKeieitai
        _chosakubun = ComUtil.GetChosakubun(_einouKeieitai)

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
            '集計結果検討表作成論理取得
            ' REV_002↓
            If _logicType = ComConst.集計結果検討表作成論理.論理種別.集計結果検討表 Then
                ' REV_001↓
                'dt = DAOOther.GetSyukeiKekkaKentohyoSakuseiRonri(db, _chosakubun)
                dt = DAOOther.GetSyukeiKekkaKentohyoSakuseiRonri(db, _chosakubun, _versionKbn)
                ' REV_001↑
            Else
                dt = DAOOther.GetSyukeiKekkaKentohyoHoukokuyoSakuseiRonri(db, _chosakubun, _versionKbn)
            End If
            ' REV_002↑
        End Using

        'シートデータ設定
        ComUtil.SyukeiKekkaKentohyoSakuseiRonri.SetSheetData(dt, xlSheets, CType(Me, ExcelProcess))

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

            SetExcelValue(xlSheet, 1, 2, ComUtil.GetChosakubunName(_chosakubun))
            ' REV_002↓
            'SetExcelValue(xlSheet, 2, 2, HEADER_TITLE)
            SetExcelValue(xlSheet, 2, 2, HEADER_TITLE(_logicType))
            ' REV_001↑

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
            xlSheet = DirectCast(xlSheets.Item(ComConst.集計結果検討表作成論理.出力用ファイル名称.SheetName), Excel.Worksheet)

            xlSheet.Unprotect()

            '行削除
            If dataCount < ComConst.集計結果検討表作成論理.出力用ファイル名称.Row.Max Then
                Me.DeleteRow(xlSheet, dataCount + ComConst.集計結果検討表作成論理.出力用ファイル名称.Row.First, ComConst.集計結果検討表作成論理.出力用ファイル名称.Row.Last)
            End If

            xlSheet.Protect()
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub
End Class
