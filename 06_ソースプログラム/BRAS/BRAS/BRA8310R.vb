Imports Microsoft.Office.Interop

''' <summary>
''' 専門調査員及び担当調査客体一覧出力
''' </summary>
''' <remarks></remarks>
Public Class BRA8310R
    Inherits ExcelOutputSingleBaseClass

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="outPath"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal outPath As String)
        MyBase.New(ComConst.専門調査員及び担当調査客体一覧.出力用ファイル名称.tempFileName, True, False, ComConst.専門調査員及び担当調査客体一覧.出力用ファイル名称.reportName, outPath, True)
    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(xlSheets As Excel.Sheets)

        Dim dt As DataTable

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '専門調査員及び担当調査客体一覧取得
            dt = DAOOther.GetSenmonChosainKyakutaiList(db)
        End Using

        'シートデータ設定
        Me.SetSheetData(dt, xlSheets)
    End Sub

    ''' <summary>
    ''' シートデータ設定
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Private Sub SetSheetData(dt As DataTable, xlSheets As Excel.Sheets)
        Dim xlSheet As Excel.Worksheet = Nothing

        Try
            'シートの設定
            xlSheet = DirectCast(xlSheets.Item(ComConst.専門調査員及び担当調査客体一覧.出力用ファイル名称.SheetName), Excel.Worksheet)

            Dim rng As Excel.Range = Nothing
            Try
                '実査設置拠点欄
                rng = xlSheet.Range(ComConst.専門調査員及び担当調査客体一覧.出力用ファイル名称.Kyoten)
                rng.Value = CommonInfo.CenterName
                ReleaseComObject(rng)

                '明細一覧
                Dim arrData(,) As Object

                rng = xlSheet.Range(ComConst.専門調査員及び担当調査客体一覧.出力用ファイル名称.Col.First & ComConst.専門調査員及び担当調査客体一覧.出力用ファイル名称.Row.First & ":" _
                                    & ComConst.専門調査員及び担当調査客体一覧.出力用ファイル名称.Col.Last & dt.Rows.Count + ComConst.専門調査員及び担当調査客体一覧.出力用ファイル名称.Row.First - 1)

                arrData = DirectCast(rng.Formula, Object(,))

                For i As Integer = 1 To dt.Rows.Count
                    arrData(i, 1) = i
                    arrData(i, 2) = dt.Rows(i - 1)("ユーザーID").ToString
                    If Not IsDBNull(dt.Rows(i - 1)("氏名")) Then
                        arrData(i, 3) = dt.Rows(i - 1)("氏名").ToString
                    End If
                    If Not IsDBNull(dt.Rows(i - 1)("更新日付")) Then
                        arrData(i, 4) = DateTime.Parse(dt.Rows(i - 1)("更新日付").ToString).ToString(ComConst.DATETIME_FORMAT)
                    End If
                    If Not IsDBNull(dt.Rows(i - 1)("センサス番号")) Then
                        arrData(i, 5) = ComUtil.GetTodofuken(dt.Rows(i - 1)("センサス番号").ToString)
                        arrData(i, 6) = ComUtil.GetShikuchoson(dt.Rows(i - 1)("センサス番号").ToString)
                        arrData(i, 7) = ComUtil.GetKyuShikuchoson(dt.Rows(i - 1)("センサス番号").ToString)
                        arrData(i, 8) = ComUtil.GetNogyoShuraku(dt.Rows(i - 1)("センサス番号").ToString)
                        arrData(i, 9) = ComUtil.GetChosaku(dt.Rows(i - 1)("センサス番号").ToString)
                        arrData(i, 10) = ComUtil.GetKyakutaiNo(dt.Rows(i - 1)("センサス番号").ToString)
                    End If
                Next

                rng.Value = arrData
                rng.Value = rng.Formula

                '罫線設定
                Me.SetBorders(rng)
            Finally
                ReleaseComObject(rng)
            End Try
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub
End Class
