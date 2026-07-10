Imports Microsoft.Office.Interop

''' <summary>
''' 労賃単価結果表出力
''' </summary>
''' <remarks></remarks>
Public Class BRA8210R
    Inherits ExcelOutputSingleBaseClass

    ''' <summary>生産費</summary>
    Private _seisanhi As String
    ''' <summary>調査年</summary>
    Private _chosaNen As String
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
    ''' <param name="seisanhi"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="outDir"></param>
    ''' <remarks></remarks>
    Public Sub New(seisanhi As String, chosaNen As String, jimusho As String, ByVal outPath As String)
        MyBase.New(ComConst.労賃単価.出力用ファイル名称.tempFileName, True, False, ComConst.労賃単価.出力用ファイル名称.reportName, outPath, False)

        _seisanhi = seisanhi
        _chosaNen = chosaNen
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
            '労賃単価取得
            dt = DAOOther.GetRouchinTanka(db, _seisanhi, _chosaNen, _jimusho)
        End Using

        'シートデータ設定
        Me.SetSheetData(dt, _seisanhi, _chosaNen, _jimusho, xlSheets)
    End Sub

    ''' <summary>
    ''' シートデータ設定
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="seisanhi"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Private Sub SetSheetData(dt As DataTable, seisanhi As String, chosaNen As String, jimusho As String, xlSheets As Excel.Sheets)
        Dim xlSheet As Excel.Worksheet = Nothing

        Try
            'シートの設定
            xlSheet = DirectCast(xlSheets.Item(ComConst.労賃単価.出力用ファイル名称.SheetName), Excel.Worksheet)

            Dim rng As Excel.Range = Nothing
            Try
                For Each dr As DataRow In dt.Rows
                    '生産費欄
                    rng = xlSheet.Range(ComConst.労賃単価.出力用ファイル名称.Seisanhi)
                    rng.Value = ComConst.生産費区分.リスト(seisanhi)
                    ReleaseComObject(rng)

                    '調査年欄
                    rng = xlSheet.Range(ComConst.労賃単価.出力用ファイル名称.ChosaNen)
                    rng.Value = chosaNen
                    ReleaseComObject(rng)

                    '都道府県欄
                    rng = xlSheet.Range(ComConst.労賃単価.出力用ファイル名称.Todofuken)
                    rng.Value = MasterDao.GetJimusyoName(jimusho)
                    ReleaseComObject(rng)

                    '明細欄
                    For Each kv As KeyValuePair(Of String, String) In ComConst.労賃単価.出力用ファイル名称.Field
                        rng = xlSheet.Range(kv.Value)
                        rng.Value = dr(kv.Key)
                        ReleaseComObject(rng)
                    Next
                Next

            Finally
                ReleaseComObject(rng)
            End Try
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub
End Class
