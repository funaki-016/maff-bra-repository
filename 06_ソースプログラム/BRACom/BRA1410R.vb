Imports Microsoft.Office.Interop

'//*************************************************************************************************
'//  修正履歴
'// ------------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前           |                修  正  内  容
'// -----------+------------+------------------------+----------------------------------------------
'//  REV_001   | 2023.11.27 |Daiko                   | 要件No.20 エラーリスト出力処理追加
'//            |            |                        |
'//*************************************************************************************************

''' <summary>
''' 調査票チェックリスト_複数客体出力
''' </summary>
''' <remarks></remarks>
Public Class BRA1410R
    Inherits ExcelOutputSingleBaseClass

    ''' <summary>エラーチェック種別</summary>
    Private _errType As ComConst.エラーチェック種別.enm
    ''' <summary>調査年</summary>
    Private _chosaNen As String
    ''' <summary>チェックリスト</summary>
    Private _checkList As List(Of Dictionary(Of String, String))
    ''' <summary>エラーチェック対象シート</summary>
    Private _sheet As List(Of String)

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="errType"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="checkList"></param>
    ''' <param name="outPath"></param>
    ''' <remarks></remarks>
    Public Sub New(errType As ComConst.エラーチェック種別.enm, chosaNen As String, checkList As List(Of Dictionary(Of String, String)), outPath As String, sheet As List(Of String))
        MyBase.New(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.tempFileName, True, False, ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.reportName(errType), outPath, True)

        _errType = errType
        _chosaNen = chosaNen
        _checkList = checkList
        _sheet = sheet

    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(xlSheets As Excel.Sheets)

        'シートデータ設定
        SetSheetData(_checkList, xlSheets)

    End Sub

    ''' <summary>
    ''' シートデータ設定
    ''' </summary>
    ''' <param name="lst"></param>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Private Sub SetSheetData(lst As List(Of Dictionary(Of String, String)), xlSheets As Excel.Sheets)
        Dim xlSheet As Excel.Worksheet = Nothing

        Try
            'シートの設定
            xlSheet = DirectCast(xlSheets.Item(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.SheetName), Excel.Worksheet)

            Dim rng As Excel.Range = Nothing
            Try
                'タイトル
                rng = xlSheet.Range(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.Title)
                rng.Value = ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.reportName(_errType)
                ReleaseComObject(rng)

                '調査年欄
                rng = xlSheet.Range(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.ChosaNen)
                rng.Value = _chosaNen
                ReleaseComObject(rng)

                '調査区分欄
                rng = xlSheet.Range(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.Chosakubun)
                rng.Value = ComUtil.GetChosakubunName(CommonInfo.Chosakubun)
                ReleaseComObject(rng)

                'エラーチェック対象シート名称
                rng = xlSheet.Range(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.Sheet)

                If _sheet.Count = ComConst.調査票.シートデータ範囲(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun))).Count Then
                    rng.Value = "全シート"
                Else
                    rng.Value = String.Join(ComConst.調査票エラーチェックリスト一覧複数客体.delimiter, _sheet)
                End If
                ReleaseComObject(rng)

                '拠点欄
                rng = xlSheet.Range(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.Kyoten)
                Select Case CommonInfo.Koutei
                    Case CommonInfo.KouteiKubun.Code.Honsyo
                        rng.Value = String.Format("[{0}]", CommonInfo.HONSYONAME)
                    Case CommonInfo.KouteiKubun.Code.Kyoku
                        rng.Value = String.Format("[{0}]", CommonInfo.KyokuName)
                    Case CommonInfo.KouteiKubun.Code.Center
                        rng.Value = String.Format("[{0}]", CommonInfo.CenterName)
                End Select
                ReleaseComObject(rng)

                '明細一覧
                Dim arrData(,) As Object

                rng = xlSheet.Range(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Col.First & ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Row.First & ":" _
                                    & ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Col.Last & lst.Count + ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Row.First - 1)

                arrData = DirectCast(rng.Formula, Object(,))

                For i As Integer = 1 To lst.Count
                    For Each kv As KeyValuePair(Of Integer, String) In ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field
                        arrData(i, kv.Key) = lst(i - 1)(kv.Value).ToString
                    Next
                    If i = ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Row.Max Then
                        Exit For
                    End If
                Next

                rng.Value = arrData
                rng.Value = rng.Formula

                '罫線設定
                SetBorders(rng)
                '行の高さを自動調整
                SetAutoFit(rng)
            Finally
                ReleaseComObject(rng)
            End Try
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub
End Class
