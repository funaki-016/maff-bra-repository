Imports Microsoft.Office.Interop

''' <summary>
''' 個別結果表チェックリスト（基本・追加）出力
''' </summary>
''' <remarks></remarks>
Public Class BRA3610R
    Inherits ExcelOutputSingleBaseClass

    ''' <summary>エラーチェック種別</summary>
    Private _errType As ComConst.エラーチェック種別.enm
    ''' <summary>調査年</summary>
    Private _chosaNen As String
    ''' <summary>チェックリスト</summary>
    Private _checkList As List(Of Dictionary(Of String, String))

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="errType"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="checkList"></param>
    ''' <param name="outPath"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal errType As ComConst.エラーチェック種別.enm, ByVal chosaNen As String, ByVal checkList As List(Of Dictionary(Of String, String)), ByVal outPath As String)
        MyBase.New(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.tempFileName, True, False, ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.reportName, outPath, True)

        _errType = errType
        _chosaNen = chosaNen
        _checkList = checkList
    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(xlSheets As Excel.Sheets)

        'シートデータ設定
        Me.SetSheetData(_checkList, xlSheets)

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
            xlSheet = DirectCast(xlSheets.Item(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.SheetName), Excel.Worksheet)

            Dim rng As Excel.Range = Nothing
            Try
                '調査年欄
                rng = xlSheet.Range(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.ChosaNen)
                rng.Value = String.Format("{0}年（産）　農業経営統計調査", _chosaNen)
                ReleaseComObject(rng)

                '調査区分欄
                rng = xlSheet.Range(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Chosakubun)
                rng.Value = String.Format("個別結果表{0}チェックリスト　{1}", ComConst.エラーチェック種別.リスト(_errType), CommonInfo.ChosakubunName)
                ReleaseComObject(rng)

                '拠点欄
                rng = xlSheet.Range(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Kyoten)
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

                rng = xlSheet.Range(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Col.First & ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Row.First & ":" _
                                    & ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Col.Last & lst.Count + ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Row.First - 1)

                arrData = DirectCast(rng.Formula, Object(,))

                For i As Integer = 1 To lst.Count
                    For Each kv As KeyValuePair(Of Integer, String) In ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field
                        arrData(i, kv.Key) = lst(i - 1)(kv.Value).ToString
                    Next
                Next

                rng.Value = arrData
                rng.Value = rng.Formula

                '罫線設定
                Me.SetBorders(rng)
                '行の高さを自動調整
                Me.SetAutoFit(rng)
            Finally
                ReleaseComObject(rng)
            End Try
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub
End Class
