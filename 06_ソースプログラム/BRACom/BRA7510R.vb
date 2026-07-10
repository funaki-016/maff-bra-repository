Imports Microsoft.Office.Interop

Public Class BRA7510R
    Inherits ExcelOutputSingleBaseClass

    Private Const CHOSAHYO_TITLE As String = "電子調査票差分データ一覧"
    Private Const KOBETSU_TITLE As String = "個別結果表差分データ一覧"
    Private Const FILE_NAME As String = "電子調査票・個別結果表差分データ一覧.xlsx"

    Private Const CELL_NAME_SOUSHIN As String = "前回送信データ"
    Private Const CELL_NAME_JUSHIN As String = "受信データ"

    Private Enum データ識別区分
        電子調査票 = 1
        個別結果表
    End Enum

    Private Const TITLE_COL1 As String = "B1"
    Private Const TITLE_COL2 As String = "B2"
    Private Const KYOTEN As String = "K3"
    Private Const SOUJUSHIN As String = "J4"

    Private Const START_COL_X As Integer = 2
    Private Const START_COL_Y As Integer = 6

    Private _chosanen As String
    Private _kyoku As String
    Private _jimusho As String
    Private _kyoten As String
    Private _soujushin As Integer

    Private Enum 送受信判定
        送信
        受信
    End Enum

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="outDir"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pOutputPath As String, pChosaNen As String, pKyoku As String, pJimusho As String, pKyoten As String, pSoujushinColName As String)
        MyBase.New(FILE_NAME, True, False, "差分データ出力", pOutputPath & "\" & FILE_NAME, False)

        _chosanen = pChosaNen
        _kyoku = pKyoku
        _jimusho = pJimusho
        _kyoten = pKyoten

        If pSoujushinColName = "送信日時" Then
            _soujushin = 送受信判定.送信
        ElseIf pSoujushinColName = "受信日時" Then
            _soujushin = 送受信判定.受信
        End If
    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(xlSheets As Excel.Sheets)

        Dim dtItem As DataTable
        Dim dtChosahyo() As DataRow
        Dim dtKobetsuChosahyo() As DataRow
        Dim xlSheet As Excel.Worksheet = Nothing
        Dim delSheet As New ArrayList

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '差分データテーブル取得
            dtItem = DAOOther.getSabunDataTable(db, _chosanen, _kyoku, _jimusho, _kyoten)
        End Using

        dtChosahyo = dtItem.Select("データ識別区分 ='" & データ識別区分.電子調査票 & "'")
        dtKobetsuChosahyo = dtItem.Select("データ識別区分 ='" & データ識別区分.個別結果表 & "'")
        Try
            xlSheet = DirectCast(xlSheets.Item(1), Excel.Worksheet)
            If dtChosahyo.Length <> 0 Then
                setChohyo(xlSheet, dtChosahyo)
            Else
                delSheet.Add(xlSheet.Name)
            End If
            ReleaseComObject(xlSheet)
            xlSheet = Nothing

            xlSheet = DirectCast(xlSheets.Item(2), Excel.Worksheet)
            If dtKobetsuChosahyo.Length <> 0 Then
                setChohyo(xlSheet, dtKobetsuChosahyo)
            Else
                delSheet.Add(xlSheet.Name)
            End If
            ReleaseComObject(xlSheet)
            xlSheet = Nothing

            '不要なシートの削除
            If delSheet.Count <> 0 Then
                For Each sheetName As String In delSheet
                    xlSheet = DirectCast(xlSheets.Item(sheetName), Excel.Worksheet)
                    xlSheet.Delete()
                    ReleaseComObject(xlSheet)
                    xlSheet = Nothing
                Next
            End If

        Finally
            ReleaseComObject(xlSheet)
            xlSheet = Nothing
        End Try
    End Sub

    Private Sub setChohyo(pXlsheet As Excel.Worksheet, dtRow() As DataRow)
        Dim rng As Excel.Range = Nothing

        Try
            '調査年の設定
            rng = pXlsheet.Range(TITLE_COL1)
            rng.Value = CStr(rng.Value).Replace("9999", _chosanen)
            rng = Nothing

            'タイトルの設定
            rng = pXlsheet.Range(TITLE_COL2)
            rng.Value = CStr(rng.Value).Replace("cccc", ComConst.調査区分.リスト(CommonInfo.Chosakubun).名称)
            rng = Nothing

            Dim koutei As String = String.Empty
            Select Case CommonInfo.Koutei
                Case CommonInfo.KouteiKubun.Code.Honsyo
                    koutei = "本省"
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    koutei = CommonInfo.KyokuName
                Case CommonInfo.KouteiKubun.Code.Center
                    koutei = CommonInfo.CenterName
            End Select

            rng = pXlsheet.Range(KYOTEN)
            rng.Value = CStr(rng.Value).Replace("kkkk", koutei)
            rng = Nothing

            rng = pXlsheet.Range(SOUJUSHIN)
            If _soujushin = 送受信判定.送信 Then
                rng.Value = CELL_NAME_SOUSHIN
            ElseIf _soujushin = 送受信判定.受信 Then
                rng.Value = CELL_NAME_JUSHIN
            End If
            rng = Nothing

            Dim count As Integer = dtRow.Length - 1
            Dim arrData(,) As Object
            Dim i As Integer = 1

            rng = pXlsheet.Range(pXlsheet.Cells(START_COL_Y, START_COL_X), pXlsheet.Cells(START_COL_Y + count, START_COL_X + 9))

            arrData = DirectCast(rng.Formula, Object(,))

            For Each d1 As DataRow In dtRow
                arrData(i, 1) = i
                arrData(i, 2) = ComUtil.GetTodofuken(d1.Item("センサス番号").ToString)
                arrData(i, 3) = ComUtil.GetShikuchoson(d1.Item("センサス番号").ToString)
                arrData(i, 4) = ComUtil.GetKyuShikuchoson(d1.Item("センサス番号").ToString)
                arrData(i, 5) = ComUtil.GetNogyoShuraku(d1.Item("センサス番号").ToString)
                arrData(i, 6) = ComUtil.GetChosaku(d1.Item("センサス番号").ToString)
                arrData(i, 7) = ComUtil.GetKyakutaiNo(d1.Item("センサス番号").ToString)
                arrData(i, 8) = d1.Item("項目番号").ToString
                arrData(i, 9) = ComUtil.GetBlankOrString(d1.Item("前回送受信データ"))
                arrData(i, 10) = ComUtil.GetBlankOrString(d1.Item("最新データ"))
                i += 1
            Next

            rng.Value = arrData
            rng.Value = rng.Formula


            '罫線設定
            Me.SetBorders(rng)
            rng = Nothing

            '印刷範囲設定
            rng = pXlsheet.Range(pXlsheet.Cells(1, 1), pXlsheet.Cells(START_COL_Y + count, START_COL_X + 10))
            pXlsheet.PageSetup.PrintArea = rng.Address
            rng = Nothing

        Finally
            ReleaseComObject(rng)
            rng = Nothing
        End Try
 
    End Sub
End Class
