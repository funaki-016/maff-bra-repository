'------------------------------------------------------------------------------------------
'| REV | 変更年月日 | 変更者             | 変更内容
'------------------------------------------------------------------------------------------
'| 001 | 2021/05/24 | TSP                | 稼働障害　330205/7806対応
'| 002 | 2021/10/19 | TSP                | 稼働障害　8369対応
'| 003 | 2023/01/25 | Daiko              | 要件No4 バージョン区分追加
'| 004 | 2024/01/05 | Daiko              | 既存不具合対応
'------------------------------------------------------------------------------------------
Imports Microsoft.Office.Interop

Public Class BRA7420R
    Inherits ExcelOutputSingleBaseClass

    ''' <summary>
    ''' 還元資料設定値テーブル
    ''' </summary>
    ''' <remarks></remarks>
    Protected mSetteing As DataTable
    ''' <summary>
    ''' 個別結果表テーブル
    ''' </summary>
    ''' <remarks></remarks>
    Protected mKobetsu As Dictionary(Of String, Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目))
    ''' <summary>
    ''' 集計結果表テーブル
    ''' </summary>
    ''' <remarks></remarks>
    Protected mShukei As Dictionary(Of String, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目))
    ''' <summary>
    ''' 還元資料固定項目定義マスタテーブル
    ''' </summary>
    ''' <remarks></remarks>
    Protected mMaster As DataTable
    ''' <summary>
    ''' 還元資料個別結果表対応マスタ
    ''' </summary>
    ''' <remarks></remarks>
    Protected mTaiou As DataTable
    ''' <summary>
    ''' 還元資料項目名マスタテーブル
    ''' </summary>
    ''' <remarks></remarks>
    Protected mKomoku As DataTable
    ' REV_003↓
    ''' <summary>
    ''' 制度受取金・積立金等項目テーブル
    ''' </summary>
    ''' <remarks></remarks>
    Protected mSeidoUketoriItem As DataTable
    ' REV_003↑
    ''' <summary>
    ''' 出力シート
    ''' </summary>
    ''' <remarks></remarks>
    Protected msheetList As ArrayList
    ''' <summary>
    ''' 経年個別結果表 選択年度
    ''' </summary>
    ''' <remarks></remarks>
    Protected mChosanenKey As ArrayList
    Protected printFlg As Boolean = False

    Protected Const PROCESS_NAME As String = "還元資料出力"
    Protected Const CENCUS As String = "K000003"
    Protected Const SHUKEI_CHIIKI As String = "S000010"
    Protected Const FOLDER_NAME As String = "還元資料_画像"

    Protected Const SHUKEIKEKKA As String = "2"
    Protected Const Name As Integer = 0
    Protected Const tani As Integer = 1

    Protected Shared EINOU_PICTURE_FILE As New Dictionary(Of String, String) From {
        {ComConst.営農類型区分.水田作, "01_営農類型_水田作.png"},
        {ComConst.営農類型区分.畑作, "02_営農類型_畑作.png"},
        {ComConst.営農類型区分.露地野菜作, "03_営農類型_露地野菜作.png"},
        {ComConst.営農類型区分.施設野菜作, "04_営農類型_施設野菜作.png"},
        {ComConst.営農類型区分.果樹作, "05_営農類型_果樹作.png"},
        {ComConst.営農類型区分.露地花き作, "06_営農類型_露地花き作.png"},
        {ComConst.営農類型区分.施設花き作, "07_営農類型_施設花き作.png"},
        {ComConst.営農類型区分.酪農, "08_営農類型_畜産.png"},
        {ComConst.営農類型区分.繁殖牛, "08_営農類型_畜産.png"},
        {ComConst.営農類型区分.肥育牛, "08_営農類型_畜産.png"},
        {ComConst.営農類型区分.養豚, "08_営農類型_畜産.png"},
        {ComConst.営農類型区分.採卵養鶏, "08_営農類型_畜産.png"},
        {ComConst.営農類型区分.ブロイラー養鶏, "08_営農類型_畜産.png"},
        {ComConst.営農類型区分.その他, "09_営農類型_その他.png"}
}

    Public Class 順位
        Public Property 項番 As String
        ''REV 001 MOD START---
        'Public Property 値 As Integer
        Public Property 値 As Long
        ''REV 001 MOD END---

        'Public Sub New(ByVal No As String, ByVal Value As Integer)
        Public Sub New(ByVal No As String, ByVal Value As Long)
            ''REV 001 MOD END---
            Me.項番 = No
            Me.値 = Value
        End Sub
    End Class

    'REV_004↓
    'Public Shared 表示単位 As New Dictionary(Of String, String) From {
    '    {ComConst.営農類型区分.水田作, "水田作作付延べ面積{0}{1}ha"},
    '    {ComConst.営農類型区分.畑作, "畑作作付延べ面積{0}{1}ha"},
    '    {ComConst.営農類型区分.露地野菜作, "露地野菜作作付延べ面積{0}{1}ha"},
    '    {ComConst.営農類型区分.施設野菜作, "施設野菜作作付延べ面積{0}{1}㎡"},
    '    {ComConst.営農類型区分.果樹作, "果樹植栽面積{0}{1}ha"},
    '    {ComConst.営農類型区分.露地花き作, "露地花き作作付延べ面積{0}999,999,{0}ha"},
    '    {ComConst.営農類型区分.施設花き作, "施設花き作作付延べ面積{0}{1}㎡"},
    '    {ComConst.営農類型区分.酪農, "搾乳牛飼養頭数{0}{1}頭"},
    '    {ComConst.営農類型区分.繁殖牛, "繁殖雌牛飼養頭数{0}{1}頭"},
    '    {ComConst.営農類型区分.肥育牛, "肥育牛飼養頭数{0}{1}頭"},
    '    {ComConst.営農類型区分.養豚, "肥育豚飼養頭数{0}{1}頭"},
    '    {ComConst.営農類型区分.採卵養鶏, "採卵鶏飼養羽数{0}{1}羽"},
    '    {ComConst.営農類型区分.ブロイラー養鶏, "ブロイラー販売羽数{0}{1}羽"},
    '    {ComConst.営農類型区分.その他, ""}
    '    }
    Public Shared 表示単位 As New Dictionary(Of String, String) From {
        {ComConst.営農類型区分.水田作, "水田作作付延べ面積{0}{1}ha"},
        {ComConst.営農類型区分.畑作, "畑作作付延べ面積{0}{1}ha"},
        {ComConst.営農類型区分.露地野菜作, "露地野菜作作付延べ面積{0}{1}a"},
        {ComConst.営農類型区分.施設野菜作, "施設野菜作作付延べ面積{0}{1}㎡"},
        {ComConst.営農類型区分.果樹作, "果樹植栽面積{0}{1}a"},
        {ComConst.営農類型区分.露地花き作, "露地花き作作付延べ面積{0}{1}a"},
        {ComConst.営農類型区分.施設花き作, "施設花き作作付延べ面積{0}{1}㎡"},
        {ComConst.営農類型区分.酪農, "搾乳牛飼養頭数{0}{1}頭"},
        {ComConst.営農類型区分.繁殖牛, "繁殖雌牛飼養頭数{0}{1}頭"},
        {ComConst.営農類型区分.肥育牛, "肥育牛飼養頭数{0}{1}頭"},
        {ComConst.営農類型区分.養豚, "肥育豚飼養頭数{0}{1}頭"},
        {ComConst.営農類型区分.採卵養鶏, "採卵鶏飼養羽数{0}{1}羽"},
        {ComConst.営農類型区分.ブロイラー養鶏, "ブロイラー販売羽数{0}{1}羽"},
        {ComConst.営農類型区分.その他, ""}
        }
    'REV_004↑

    Public Shared 規模階層コード_個人経営体 As New Dictionary(Of String, String) From {
        {ComConst.営農類型区分.水田作, "1-001"},
        {ComConst.営農類型区分.畑作, "1-002"},
        {ComConst.営農類型区分.露地野菜作, "1-013"},
        {ComConst.営農類型区分.施設野菜作, "1-014"},
        {ComConst.営農類型区分.果樹作, "1-015"},
        {ComConst.営農類型区分.露地花き作, "1-018"},
        {ComConst.営農類型区分.施設花き作, "1-019"},
        {ComConst.営農類型区分.酪農, "1-020"},
        {ComConst.営農類型区分.繁殖牛, "1-022"},
        {ComConst.営農類型区分.肥育牛, "1-023"},
        {ComConst.営農類型区分.養豚, "1-024"},
        {ComConst.営農類型区分.採卵養鶏, "1-025"},
        {ComConst.営農類型区分.ブロイラー養鶏, "1-026"},
        {ComConst.営農類型区分.その他, ""}
        }

    Public Shared 規模階層コード_法人経営体 As New Dictionary(Of String, String) From {
        {ComConst.営農類型区分.水田作, "1-101"},
        {ComConst.営農類型区分.畑作, "1-102"},
        {ComConst.営農類型区分.露地野菜作, "1-104"},
        {ComConst.営農類型区分.施設野菜作, "1-105"},
        {ComConst.営農類型区分.果樹作, "1-106"},
        {ComConst.営農類型区分.露地花き作, "1-108"},
        {ComConst.営農類型区分.施設花き作, "1-109"},
        {ComConst.営農類型区分.酪農, "1-110"},
        {ComConst.営農類型区分.繁殖牛, "1-112"},
        {ComConst.営農類型区分.肥育牛, "1-113"},
        {ComConst.営農類型区分.養豚, "1-114"},
        {ComConst.営農類型区分.採卵養鶏, "1-115"},
        {ComConst.営農類型区分.ブロイラー養鶏, "1-116"},
        {ComConst.営農類型区分.その他, ""}
    }

    '集計結果表で「円」となるもの
    'REV 002 MOD START---
    'Public Shared shukeiEn_eiko() As String = {"S010513", "S010523", "S010527", "S010528", "S010529", "S010530", "S010537", "S010538", "S010542", "S031041", "S031046", "S031047", "S031048", "S031049", "S031329", _
    '                                      "S031330", "S031331", "S031332", "S031333", "S031334", "S031335", "S031336", "S031337", "S031338", "S031339", "S031340", "S031341", "S031342", "S031343", "S031344", _
    '                                      "S031345", "S031346", "S031347", "S031348", "S031349", "S031350", "S031351", "S031352"}

    'Public Shared shukeiEn_eiho() As String = {"S010529", "S010533", "S010543", "S010548"}

    Public Shared shukeiEn_eiko() As String = {"S010513", "S010523", "S010527", "S010528", "S010529", "S010530", "S010537", "S010538", "S010542", "S010543", "S031041", "S031046", "S031047", "S031048", "S031049", "S031329",
                                          "S031330", "S031331", "S031332", "S031333", "S031334", "S031335", "S031336", "S031337", "S031338", "S031339", "S031340", "S031341", "S031342", "S031343", "S031344",
                                          "S031345", "S031346", "S031347", "S031348", "S031349", "S031350", "S031351", "S031352"}

    Public Shared shukeiEn_eiho() As String = {"S010519", "S010529", "S010533", "S010534", "S010535", "S010536", "S010543", "S010544", "S010548", "S010549"}
    'REV 002 MOD END---

    Public Sub New()


    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="outDir"></param>
    ''' <remarks></remarks>
    Public Sub New(pCheckBox As ArrayList, pPrintFlg As Boolean, pOutputFlg As Boolean, pOutputPath As String, pSetting As DataTable,
                   pKobetsu As Dictionary(Of String, Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)),
                   pShukei As Dictionary(Of String, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)),
                   pMaster As DataTable, pTaiouMaster As DataTable, pKoumokuMaster As DataTable, pChosaNenKey As ArrayList,
                   pSeidoUketoriItem As DataTable) ' REV 003

        MyBase.New(ComConst.還元資料.テンプレートファイル名称(CommonInfo.Chosakubun), pOutputFlg, pPrintFlg, PROCESS_NAME, pOutputPath, False)

        mSetteing = pSetting '還元資料設定値テーブル
        mKobetsu = pKobetsu '個別結果表テーブル
        mShukei = pShukei '集計結果表テーブル
        mMaster = pMaster '還元資料固定項目定義マスタテーブル
        mKomoku = pKoumokuMaster '還元資料項目名マスタテーブル
        mTaiou = pTaiouMaster '還元資料個別結果表対応マスタ
        mSeidoUketoriItem = pSeidoUketoriItem '制度受取金・積立金等項目テーブル REV 003
        printFlg = pPrintFlg

        mChosanenKey = New ArrayList(pChosaNenKey) '経年個別結果表 選択年度
        msheetList = New ArrayList(pCheckBox) '出力シート
    End Sub



    ''' <summary>
    ''' 作成実行
    ''' </summary>
    ''' <param name="pCensusNo">センサス番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(xlSheets As Excel.Sheets)

    End Sub

    '年号変換
    Protected Function getNengo(pSeireki As String, pRyakugo As Boolean) As String
        Dim ret As String
        Dim dt As DateTime = New DateTime(CInt(pSeireki), 12, 31)
        Dim calendarJp = New System.Globalization.JapaneseCalendar()

        Dim era As Integer = calendarJp.GetEra(dt)

        Dim cultureJp = New System.Globalization.CultureInfo("ja-JP", False)
        cultureJp.DateTimeFormat.Calendar = calendarJp

        If pRyakugo = False Then
            '和暦の場合
            ret = dt.ToString("ggy年", cultureJp)
        Else
            '略号の場合
            Dim eraTable = New Dictionary(Of Integer, String)()

            For code = AscW("A"c) To AscW("Z"c)
                Dim e As String = ChrW(code)
                Dim eraIndex As Integer = cultureJp.DateTimeFormat.GetEra(e)
                If (eraIndex > 0) Then
                    eraTable.Add(eraIndex, e)
                End If
            Next

            ret = eraTable(era) + dt.ToString("yy", cultureJp)
        End If

        Return ret

    End Function

    '数字変換
    ''REV 001 MOD START---
    'Protected Function changeNum(pNumber As String) As Integer
    '    Dim ret As Integer
    Protected Function changeNum(pNumber As String) As Long
        Dim ret As Long
        'REV 001 MOD END---
        If pNumber = String.Empty Then
            ret = 0
        Else
            'REV 001 MOD START---
            'ret= CInt(pNumber)
            ret = CLng(pNumber)
            'REV 001 MOD END---
        End If

        Return ret
    End Function

    Protected Function createLank(kobetsuList() As String) As List(Of 順位)
        Dim ret As New List(Of 順位)

        For i = 0 To kobetsuList.Count - 1
            ''REV 001 MOD START---
            'Dim value As Integer = changeNum(mKobetsu(CStr(mChosanenKey(0))).Item(kobetsuList(i)).値)
            Dim value As Long = changeNum(mKobetsu(CStr(mChosanenKey(0))).Item(kobetsuList(i)).値)
            ''REV 001 MOD END---
            Dim no As String = kobetsuList(i)

            ret.Add(New 順位(no, value))
        Next

        ''REV 001 MOD START---
        'ret.Sort(Function(a, b) b.値 - a.値)
        ret.Sort(Function(a, b)
                     Dim tmpRet As Integer
                     If b.値 - a.値 = 0 Then
                         tmpRet = 0
                     ElseIf b.値 - a.値 > 0 Then
                         tmpRet = 1
                     ElseIf b.値 - a.値 < 0 Then
                         tmpRet = -1
                     End If
                     Return tmpRet
                 End Function)
        ''REV 001 MOD END---

        Return ret


    End Function

    ''' <summary>
    ''' 還元資料固定項目の設定(派生先でのカスタム項目)
    ''' </summary>
    ''' <param name="sheetNo"></param>
    ''' <param name="arrayNo"></param>
    ''' <param name="kobetsuArray"></param>
    ''' <param name="shukeiArray"></param>
    ''' <remarks></remarks>
    Protected Overridable Sub SetKangenShiryoFixedAdditional(sheetNo As Integer, arrayNo As Integer, ByRef kobetsuArray(,) As Object, ByRef shukeiArray(,) As Object)

    End Sub

    ''' <summary>
    ''' 還元資料固定項目の設定
    ''' </summary>
    ''' <param name="xlSheet"></param>
    ''' <param name="sheetNo"></param>
    ''' <param name="arrayNo"></param>
    ''' <param name="cell"></param>
    ''' <param name="shukeiCell"></param>
    ''' <param name="cellSpace"></param>
    ''' <param name="shukeiCellSpace"></param>
    ''' <remarks></remarks>
    Protected Sub SetKangenShiryoFixed(xlSheet As Excel.Worksheet, sheetNo As Integer, arrayNo As Integer, rowCount As Integer,
                                     cell As String, Optional shukeiCell As String = "",
                                     Optional cellSpace As UInteger = 0, Optional shukeiCellSpace As UInteger = 0)
        Dim d1 As DataRow()
        Dim dataArray(rowCount - 1, CInt(1 + cellSpace)) As Object
        Dim dataShukeiArray(rowCount - 1, CInt(1 + shukeiCellSpace)) As Object

        d1 = mMaster.Select("シート番号 = " + sheetNo.ToString + " AND 配列番号 = " + arrayNo.ToString)

        For Each d2 As DataRow In d1

            Dim kobetsuNo As String = d2.Item("個別結果表項番").ToString

            If mKobetsu(CStr(mChosanenKey(0))).ContainsKey(d2.Item("個別結果表項番").ToString) Then

                ' 個別 当年
                dataArray(CInt(d2.Item("ROW")), CInt(d2.Item("COL"))) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(kobetsuNo))

                ' 集計
                If String.IsNullOrEmpty(shukeiCell) = False Then
                    Dim d3 As DataRow()
                    Dim seisanHeikin As String = "0"
                    Dim shukeiNo As String = ""

                    d3 = mTaiou.Select("個別結果表項番 = '" + kobetsuNo + "'")

                    '対応する集計結果表を設定
                    GetTaiouShukeiNo(kobetsuNo, seisanHeikin, shukeiNo)

                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１ & "_" & seisanHeikin) Then
                        ' 集計 当年
                        dataShukeiArray(CInt(d2.Item("ROW")), CInt(d2.Item("COL"))) =
                            ShukeiDataHenkan(kobetsuNo, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                    End If

                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年 & "_" & seisanHeikin) Then
                        ' 集計 前年
                        dataShukeiArray(CInt(d2.Item("ROW")), CInt(CInt(d2.Item("COL")) + 1 + shukeiCellSpace)) =
                             ShukeiDataHenkan(kobetsuNo, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)
                    End If

                End If
            End If

            Dim zennen As Boolean = False
            If mChosanenKey.Count <> 1 Then
                zennen = True
            End If

            ' 個別 前年
            If zennen Then
                If mKobetsu(CStr(mChosanenKey(1))).ContainsKey(kobetsuNo) Then
                    dataArray(CInt(d2.Item("ROW")), CInt(CInt(d2.Item("COL")) + 1 + cellSpace)) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(1))).Item(kobetsuNo))
                End If
            End If
        Next

        '派生先で必要に応じて設定する項目
        If String.IsNullOrEmpty(shukeiCell) = True Then
            SetKangenShiryoFixedAdditional(sheetNo, arrayNo, dataArray, Nothing)
        Else
            SetKangenShiryoFixedAdditional(sheetNo, arrayNo, dataArray, dataShukeiArray)
        End If


        Dim range As Excel.Range = Nothing

        Try
            ' 個別結果表 設定値
            range = xlSheet.Range(cell)
            range.Value = dataArray
            ReleaseComObject(range)
            range = Nothing

            If String.IsNullOrEmpty(shukeiCell) = False Then
                ' 集計結果表 設定値
                range = xlSheet.Range(shukeiCell)
                range.Value = dataShukeiArray
                ReleaseComObject(range)
                range = Nothing
            End If
        Finally
            ReleaseComObject(range)
            range = Nothing
        End Try

    End Sub

    ''' <summary>
    ''' 配列に格納されたデータ形式で、セルの書式設定を行う
    ''' </summary>
    ''' <param name="range"></param>
    ''' <param name="array"></param>
    ''' <remarks></remarks>
    Protected Sub SetCellFormat(ByRef range As Excel.Range, dataArray(,) As Object)

        Dim decWork As Decimal

        For cnt1 As Integer = 0 To dataArray.GetLength(0) - 1
            For cnt2 As Integer = 0 To dataArray.GetLength(1) - 1
                If dataArray(cnt1, cnt2) Is Nothing Then
                    Continue For
                End If
                If Not Decimal.TryParse(dataArray(cnt1, cnt2).ToString, decWork) Then
                    Continue For
                End If

                Dim formatString As String = dataArray(cnt1, cnt2).ToString
                If formatString.Contains("."c) Then
                    '小数点以降を0に置き換え
                    Dim digit As Integer = formatString.Length - formatString.IndexOf("."c) - 1
                    formatString = formatString.Remove(formatString.IndexOf("."c) + 1)
                    For strcnt As Integer = 1 To digit
                        formatString += "0"c
                    Next
                    '小数点以前は3桁おきのカンマ区切り
                    formatString = formatString.Remove(0, formatString.IndexOf("."c))
                    formatString = formatString.Insert(0, "###,##0")
                Else
                    '小数でない場合はカンマ区切りのみ
                    formatString = "###,##0"
                End If
                '負数の設定
                formatString += (";""△""" + formatString)

                DirectCast(range.Cells(cnt1 + 1, cnt2 + 1), Excel.Range).NumberFormatLocal = formatString
            Next
        Next


    End Sub

    ''' <summary>
    ''' 還元資料任意項目の設定
    ''' </summary>
    ''' <param name="xlSheet"></param>
    ''' <param name="itemNo"></param>
    ''' <param name="rowCount"></param>
    ''' <param name="titleCell"></param>
    ''' <param name="dataCell"></param>
    ''' <param name="shukeiDataCell"></param>
    ''' <param name="titleCellSpace"></param>
    ''' <param name="dataCellSpace"></param>
    ''' <param name="shukeiDataCell"></param>
    ''' <remarks></remarks>
    Protected Sub SetKangenShryoOptional(xlSheet As Excel.Worksheet, itemNo As String, rowCount As Integer,
                                       titleCell As String, dataCell As String, Optional shukeiDataCell As String = "",
                                       Optional titleCellSpace As UInteger = 0, Optional dataCellSpace As UInteger = 0, Optional shukeiDataCellSpace As UInteger = 0)

        '配列の作成 (セル結合によりセル間にインターバルがある場合の調整をしたサイズ)
        Dim titleList(rowCount - 1, CInt(1 + titleCellSpace)) As Object
        Dim dataArray(rowCount - 1, CInt(1 + dataCellSpace)) As Object
        Dim dataShukeiArray(rowCount - 1, CInt(1 + shukeiDataCellSpace)) As Object
        Dim formatArray(rowCount - 1, CInt(1 + dataCellSpace)) As Object

        '個別結果表設定テーブルの取得
        Dim d1 As DataRow()
        d1 = mSetteing.Select("項目番号 = '" + itemNo + "'")

        Dim i As Integer = 0
        For Each d2 As DataRow In d1
            'タイトルを配列に設定する
            titleList(i, 0) = getKoumokuName(d2.Item("設定値").ToString, Name)
            titleList(i, CInt(1 + titleCellSpace)) = getKoumokuName(d2.Item("設定値").ToString, tani)

            Dim kobetsuNo As String = d2.Item("設定値").ToString

            If mKobetsu(CStr(mChosanenKey(0))).ContainsKey(kobetsuNo) Then

                ' 個別 当年
                Dim tmp As String = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(kobetsuNo))
                If tmp.Equals(String.Empty) Then
                    dataArray(i, 0) = String.Empty
                Else
                    dataArray(i, 0) = tmp
                End If

                If String.IsNullOrEmpty(shukeiDataCell) = False Then
                    Dim seisanHeikin As String = "0"
                    Dim shukeiNo As String = ""

                    '対応する集計結果表を設定
                    GetTaiouShukeiNo(kobetsuNo, seisanHeikin, shukeiNo)

                    ' 集計 当年
                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１ & "_" & seisanHeikin) Then
                        dataShukeiArray(i, 0) =
                            ShukeiDataHenkan(kobetsuNo, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                    End If

                    ' 集計 前年
                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年 & "_" & seisanHeikin) Then
                        dataShukeiArray(i, CInt(1 + shukeiDataCellSpace)) =
                            ShukeiDataHenkan(kobetsuNo, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)
                    End If

                    '個別結果表と同じ書式を設定する
                    Dim item As DAOKobetsuKekkahyo.個別結果表項目 = mKobetsu(CStr(mChosanenKey(0))).Item(kobetsuNo)
                    Dim format As String = String.Empty
                    If Not String.IsNullOrEmpty(item.表示単位) Then
                        Dim unit As Decimal = Decimal.Parse(item.表示単位)
                        If ComConst.個別結果表作成論理.表示単位.リスト.ContainsKey(unit) Then
                            format = 1D.ToString(ComConst.個別結果表作成論理.表示単位.リスト(unit))
                            'フォーマット設定用配列に1をフォーマットした値を設定
                            formatArray(i, 0) = format
                            formatArray(i, CInt(1 + shukeiDataCellSpace)) = format
                        End If
                    End If
                End If
            End If

            Dim zennen As Boolean = False
            If mChosanenKey.Count <> 1 Then
                zennen = True
            End If

            ' 個別 前年
            If zennen Then
                If mKobetsu(CStr(mChosanenKey(1))).ContainsKey(kobetsuNo) Then

                    Dim tmp As String = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(1))).Item(kobetsuNo))
                    If tmp.Equals(String.Empty) Then
                        dataArray(i, CInt(1 + dataCellSpace)) = String.Empty
                    Else
                        dataArray(i, CInt(1 + dataCellSpace)) = tmp
                    End If
                End If
            End If

            i += 1
        Next

        Dim range As Excel.Range = Nothing

        Try
            ' 項目名、単位
            range = xlSheet.Range(titleCell)
            range.Value = titleList
            ReleaseComObject(range)
            range = Nothing

            ' 個別結果表 設定値
            range = xlSheet.Range(dataCell)
            range.Value = dataArray
            SetCellFormat(range, dataArray)
            ReleaseComObject(range)
            range = Nothing

            If String.IsNullOrEmpty(shukeiDataCell) = False Then
                ' 集計結果表 設定値
                range = xlSheet.Range(shukeiDataCell)
                range.Value = dataShukeiArray
                SetCellFormat(range, formatArray) 'セル書式は個別結果表のものを使用する
                ReleaseComObject(range)
                range = Nothing
            End If
        Finally
            ReleaseComObject(range)
            range = Nothing
        End Try

    End Sub


    ''' <summary>
    ''' 還元資料個別結果表対応マスタから対応する集計結果表の生産費平均値種類、項番を取得
    ''' </summary>
    ''' <param name="kobetsuNo"></param>
    ''' <param name="seisanHeikin"></param>
    ''' <param name="shukeiNo"></param>
    ''' <remarks></remarks>
    Protected Sub GetTaiouShukeiNo(kobetsuNo As String, ByRef seisanHeikin As String, ByRef shukeiNo As String)
        Dim d As DataRow()

        d = mTaiou.Select("個別結果表項番 = '" + kobetsuNo + "'")

        '対応する集計結果表を設定
        If d.Count >= 1 Then
            seisanHeikin = d(0).Item("生産費平均値種類").ToString
            shukeiNo = d(0).Item("集計結果表項番").ToString
        Else
            ' 対応するデータがなかった場合、先頭文字置き換え(K→S)
            seisanHeikin = "0"
            shukeiNo = kobetsuNo.Replace("K"c, "S"c)
        End If

    End Sub

    ''' <summary>
    ''' 集計結果表の値取得
    ''' </summary>
    ''' <param name="shukeiMeisaiNo"></param>
    ''' <param name="seisanHeikin"></param>
    ''' <param name="shukeiNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function GetShukeiKekkahyoValue(shukeiMeisaiNo As Integer, seisanHeikin As String, shukeiNo As String) As String

        Dim key As String = shukeiMeisaiNo.ToString + "_" + seisanHeikin

        If mShukei.ContainsKey(key) Then
            If mShukei(key).ContainsKey(shukeiNo) Then
                Return mShukei(key).Item(shukeiNo).値
            End If
        End If

        Return Nothing

    End Function

    '還元資料 年の設定
    Protected Sub SetKangenShiryoYear(xlSheet As Excel.Worksheet, thisYearCell As String, prevYearCell As String)
        Dim range As Excel.Range = Nothing
        Dim yearString As String = "年"

        Try
            If ComConst.調査区分.リスト(CommonInfo.Chosakubun).区分２ = ComConst.区分２.農産物生産費 Then
                '農産物生産費の場合は、「年産」とする
                yearString += "産"
            End If


            range = xlSheet.Range(thisYearCell)
            range.Value = getNengo(mChosanenKey(0).ToString, True) + yearString
            ReleaseComObject(range)
            range = Nothing

            'If mChosanenKey.Count >= 2 Then
            '前年
            If mChosanenKey.Count <> 1 Then
                range = xlSheet.Range(prevYearCell)
                range.Value = getNengo(mChosanenKey(1).ToString, True) + yearString
                ReleaseComObject(range)
                range = Nothing
            End If

            'End If
        Finally
            ReleaseComObject(range)
            range = Nothing
        End Try

    End Sub

    Protected Function getKoumokuName(pKobetsukouban As String, Optional type As Integer = 0) As String
        Dim ret As String = String.Empty
        Dim d1 As DataRow()

        d1 = mKomoku.Select("個別結果表項番 = '" & pKobetsukouban & "'", " 営農類型 DESC")

        If d1.Count <> 0 Then
            If type = Name Then
                ret = d1(0).Item("項目名").ToString
                ' REV_003↓
                '項目番号(K010101)を検索
                Dim matchList As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ret, CreateKobetsu.C_SearchKakko)
                For Each mat As System.Text.RegularExpressions.Match In matchList
                    Dim d2 As DataRow()
                    d2 = mSeidoUketoriItem.Select("項目番号 = '" & mat.Value.Trim("["c, "]"c) & "'")
                    If d2.Count <> 0 Then
                        ret = ret.Replace(mat.Value, d2(0).Item("出力項目名").ToString)
                    End If
                Next
                ' REV_003↑
            ElseIf type = tani Then
                ret = d1(0).Item("単位").ToString
            End If
        End If


        Return ret
    End Function

    Protected Sub SetShukeiHeader(pXlSheet As Excel.Worksheet, thisYearCell As String, PrevYearCell As String, heikinCell As String, mensekiCell As String)
        Dim range As Excel.Range = Nothing
        Dim title(1, 0) As String
        Dim nen(0, 1) As String
        Dim seisanHeikin As String

        Dim d As DataRow()

        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 OrElse CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            seisanHeikin = "0"
        Else
            seisanHeikin = String.Empty
        End If


        Try
            If mShukei.Count <> 0 Then
                Dim Heikin As String = GetShukeiKekkahyoValue(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１, seisanHeikin, SHUKEI_CHIIKI)
                Dim zennen As Boolean = False

                If IsNothing(Heikin) Then
                    Heikin = GetShukeiKekkahyoValue(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年, seisanHeikin, SHUKEI_CHIIKI)
                    zennen = True
                End If

                Dim heikinName As String = String.Empty
                If Not IsNothing(Heikin) Then
                    heikinName = ComConst.地域.リスト(Heikin).名称
                    heikinName = heikinName.Replace("平均", String.Empty)
                    heikinName = heikinName & "平均"
                End If

                title(0, 0) = heikinName

                title(1, 0) = GetHyoujiTani(False, zennen)

            End If

            nen(0, 0) = getNengo(CStr(CInt(mChosanenKey(0))), True) + "年"

            d = mSetteing.Select("項目番号 = " & ComConst.還元資料.集計結果表 & " AND 明細番号 = " & ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)

            If d.Count <> 0 Then
                If Not String.IsNullOrEmpty(CStr(d(0).Item("設定値"))) Then
                    nen(0, 1) = getNengo(CStr(CInt(mChosanenKey(0)) - 1), True) + "年"
                End If
            End If

            range = pXlSheet.Range(heikinCell, mensekiCell)
            range.Value = title
            ReleaseComObject(range)
            range = Nothing

            range = pXlSheet.Range(thisYearCell, PrevYearCell)
            range.Value = nen
            ReleaseComObject(range)
            range = Nothing

        Finally
            ReleaseComObject(range)
            range = Nothing
        End Try

    End Sub

    Protected Function GetHyoujiTani(Optional suujiOnlyFlg As Boolean = False, Optional zennen As Boolean = False) As String
        Dim kaisouCode As String = String.Empty

        Dim max As String = String.Empty
        Dim min As String = String.Empty
        Dim ret As String = String.Empty

        Dim d As DataRow()
        Dim KaisouTable As DataTable

        If zennen = True Then
            d = mSetteing.Select("項目番号 = " & ComConst.還元資料.集計結果表 & " AND 明細番号 = " & ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)
        Else
            d = mSetteing.Select("項目番号 = " & ComConst.還元資料.集計結果表 & " AND 明細番号 = " & ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
        End If

        If d.Count <> 0 Then

            Dim values() As String = d(0).Item("設定値").ToString.Split("_"c)
            Dim kaisouKibo As String = values(values.Count - 1)
            '階層コードマスタを取得する
            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                kaisouCode = 規模階層コード_個人経営体(mKobetsu(CStr(mChosanenKey(0))).Item("K000005").値)
            ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                kaisouCode = 規模階層コード_法人経営体(mKobetsu(CStr(mChosanenKey(0))).Item("K000006").値)
            End If

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                KaisouTable = DAOOther.getKiboKaisouMaster(db, CommonInfo.Chosakubun, kaisouCode)
            End Using

            If Not String.IsNullOrEmpty(kaisouKibo) Then
                Dim d2 As DataRow()
                d2 = KaisouTable.Select("規模階層 = " & kaisouKibo)
                If d2.Count = 1 Then
                    max = CStr(d2(0).Item("上限"))
                    min = CStr(d2(0).Item("下限"))
                    If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                        ret = 表示単位(mKobetsu(CStr(mChosanenKey(0))).Item("K000005").値)
                    ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                        ret = 表示単位(mKobetsu(CStr(mChosanenKey(0))).Item("K000006").値)
                    End If
                End If
            End If
        End If

        If Not ret = String.Empty Then
            If suujiOnlyFlg Then
                ret = ret.Remove(0, ret.IndexOf("{"))
            End If

            If min = "0.0" Then
                ret = ret.Replace("{0}", String.Empty)
                ret = ret.Replace("{1}", max)
                ret = ret & "以下"
            ElseIf max = "999999.9" Then
                ret = ret.Replace("{0}", String.Empty)
                ret = ret.Replace("{1}", max)
                ret = ret & "以上"
            Else
                ret = ret.Replace("{0}", min & "～")
                ret = ret.Replace("{1}", max)
            End If
        End If

        Return ret
    End Function

    Protected Function ShukeiDataHenkan(kobetsuKouban As String, MeisaiNo As Integer) As String
        Dim ret As String = String.Empty
        Dim Value As Decimal
        Dim shukeiValue As String
        Dim seisanHeikin As String = "0"
        Dim shukeiNo As String = String.Empty
        Dim enFlg As Boolean = False
        Dim d1 As DataRow()

        GetTaiouShukeiNo(kobetsuKouban, seisanHeikin, shukeiNo)
        d1 = mSetteing.Select("項目番号 = " & ComConst.還元資料.集計結果表 & " AND 明細番号 = " & MeisaiNo)
        If Not d1.Count = 0 Then
            If Not String.IsNullOrEmpty(CStr(d1(0).Item("設定値"))) Then
                shukeiValue = GetShukeiKekkahyoValue(MeisaiNo, seisanHeikin, shukeiNo)
            Else
                shukeiValue = String.Empty
            End If
        Else
            Return ret
        End If


        If Not String.IsNullOrEmpty(shukeiValue) Then
            If getKoumokuName(kobetsuKouban, 1) = "円" Then
                If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                    For i As Integer = 0 To shukeiEn_eiko.Count - 1
                        If shukeiEn_eiko(i) = shukeiNo Then
                            enFlg = True
                            Exit For
                        End If
                    Next
                ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                    For i As Integer = 0 To shukeiEn_eiho.Count - 1
                        'REV 002 MOD START---
                        'If shukeiEn_eiko(i) = shukeiNo Then
                        If shukeiEn_eiho(i) = shukeiNo Then
                            'REV 002 MOD END---
                            enFlg = True
                            Exit For
                        End If
                    Next
                Else
                    enFlg = True
                End If
                If enFlg = False Then
                    Value = CDec(shukeiValue) * 1000
                    If Value <> 0 Then
                        ret = CStr(Value)
                    Else
                        ret = String.Empty
                    End If
                Else
                    ret = shukeiValue
                End If
            Else
                ret = shukeiValue
            End If
        Else
            ret = String.Empty
        End If

        Return ret
    End Function

    Protected Sub Ranking_Title(pXlSheet As Excel.Worksheet, Rank As List(Of 順位), row As Integer, col As Integer, titleCellSpace As String)
        Dim i As Integer
        Dim RankName(row, col) As Object
        Dim Range As Excel.Range = Nothing

        For i = 0 To Rank.Count - 1
            If Rank(i).値 <> 0 Then
                RankName(i, 0) = getKoumokuName(Rank(i).項番, Name)
            Else
                RankName(i, 0) = String.Empty
            End If
            If i = row Then
                Exit For
            End If
        Next


        Range = pXlSheet.Range(titleCellSpace)
        Range.Value = RankName
        ReleaseComObject(Range)
        Range = Nothing

    End Sub

    Public Function getSum(ParamArray number() As Integer) As String
        Dim sum As Decimal = 0
        Dim ret As String = String.Empty

        For i As Integer = 0 To number.Count - 1
            sum = sum + number(i)
        Next

        If sum = 0 Then
            ret = String.Empty
        Else
            ret = CStr(ret)
        End If

        Return ret
    End Function

    Protected Function getHeikinchishurui(shukeiDataMeisai As Integer) As String
        Dim d1() As DataRow
        Dim heikin As String = String.Empty
        Dim ret As String = String.Empty

        d1 = mSetteing.Select("項目番号 = " & ComConst.還元資料.集計結果表 & " AND 明細番号 = " & shukeiDataMeisai)

        If d1.Count <> 0 Then
            Dim tmp() As String
            tmp = d1(0).Item("設定値").ToString.Split("_"c)
            If tmp.Count = 4 Then
                heikin = tmp(2)
            End If

            If Not String.IsNullOrEmpty(heikin) Then
                ret = ComConst.地域.リスト(heikin).名称
                ret = ret.Replace("平均", String.Empty)
                ret = ret & "平均"
            End If
        End If

        Return ret

    End Function

    ''' <summary>
    ''' グラフの凡例が空白の場合、凡例を非表示にする関数
    ''' </summary>
    ''' <param name="pXlsheet"></param>
    ''' <param name="pChartObjectName"></param>
    ''' <remarks></remarks>
    Protected Sub setGraphSetting(pXlsheet As Excel.Worksheet, pChartObjectName As String)
        Dim xlSeriesCollection As Excel.SeriesCollection
        Dim chartObject As Excel.ChartObject
        Dim chart As Excel.Chart
        Dim i As Integer = 1

        chartObject = DirectCast(pXlsheet.ChartObjects(pChartObjectName), Excel.ChartObject)
        chart = chartObject.Chart

        xlSeriesCollection = DirectCast(chart.SeriesCollection, Excel.SeriesCollection)

        For Each xlSeries As Excel.Series In xlSeriesCollection
            If String.IsNullOrEmpty(xlSeries.Name) Then
                xlSeries.IsFiltered = True
            End If
        Next

        chartObject = Nothing
        chart = Nothing
        xlSeriesCollection = Nothing

    End Sub

    ''' <summary>
    ''' 文字列を指定桁で区切り、改行コードを挿入する
    ''' </summary>
    ''' <param name="pKoumokuName">文字列</param>
    ''' <param name="splitCount">区切り桁数</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function changeSplit(pKoumokuName As String, splitCount As Integer) As String
        Dim sString As String = String.Empty

        For i As Integer = splitCount To pKoumokuName.Length + splitCount Step splitCount
            If i = splitCount Then
                sString = Mid(pKoumokuName, i - splitCount + 1, splitCount)
            Else
                sString = sString & vbLf & Mid(pKoumokuName, i - splitCount + 1, splitCount)
            End If
        Next

        Return sString
    End Function


End Class
