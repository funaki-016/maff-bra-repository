Imports System.Text.RegularExpressions
Imports System.Runtime.Serialization.Formatters.Binary

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.10.11 |Daiko               | 要件No1 バージョン区分追加
'//  REV_002   | 2024.05.31 |Daiko               | 要件No1対応
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 個別結果表作成エラークラス
''' </summary>
''' <remarks></remarks>
Public Class CreateKobetsuException
    Inherits Exception

    ''' <summary>センサス番号</summary>
    Private _censusNo As String

    ''' <summary>センサス番号</summary>
    Public ReadOnly Property CensusNo As String
        Get
            Return _censusNo
        End Get
    End Property

    ''' <summary>項目番号</summary>
    Private _itemNo As String

    ''' <summary>項目番号</summary>
    Public ReadOnly Property ItemNo As String
        Get
            Return _itemNo
        End Get
    End Property

    Public Sub New(inner As Exception, ByVal pCensusNo As String, ByVal pItemNo As String, ByVal pFormula As String, ByVal pFormulaCnv As String)
        MyBase.New(String.Format("個別結果表作成エラー(項目番号:{0} 計算式:{1} 計算式(変換後):{2})", pItemNo, pFormula, pFormulaCnv), inner)
        _censusNo = pCensusNo
        _itemNo = pItemNo
    End Sub
End Class

''' <summary>
''' 個別結果表・個別結果検討表作成クラス
''' </summary>
''' <remarks></remarks>
Public Class CreateKobetsu

    ''' <summary>
    ''' 作成種別
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum enmCreateType
        個別結果表作成
        個別結果表再計算
        個別結果検討表作成
    End Enum

    ''' <summary>
    ''' 項目種別
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum enmItemType
        調査票
        個別結果表
        個別結果検討表
    End Enum

    ''' <summary>
    ''' 項目情報クラス
    ''' </summary>
    ''' <remarks></remarks>
    <Serializable()> Public Class ItemInfo
        ''' <summary>項目番号</summary>
        Public ItemNo As String
        ''' <summary>項目名</summary>
        Public ItemName As String
        ''' <summary>優先順位</summary>
        Public Priority As Decimal
        ''' <summary>計算式</summary>
        Public Formula As String
        ''' <summary>計算式(項番以外変換)</summary>
        Public FormulaConv As String
        ''' <summary>計算式(値変換)</summary>
        Public FormulaConvValue As String
        ''' <summary>表示単位</summary>
        Public HyojiTani As Decimal?
        ''' <summary>シート名</summary>
        Public SheetName As String
        ''' <summary>行位置</summary>
        Public Row As Integer
        ''' <summary>列位置</summary>
        Public Col As Integer
        ''' <summary>型区分</summary>
        Public ValueType As String
        ''' <summary>裏項番区分</summary>
        Public IsHidden As Boolean
        ''' <summary>再計算区分</summary>
        Public IsReCalc As Boolean
        ''' <summary>項目種別</summary>
        Public ItemType As enmItemType
        ''' <summary>値</summary>
        Public Value As Object
    End Class

    ''' <summary>
    ''' 順位クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class RankInfo
        ''' <summary>先頭からの連番</summary>
        Public Number As Integer
        ''' <summary>項目番号</summary>
        Public ItemNo As String
        ''' <summary>値</summary>
        Public Value As Decimal
    End Class

    ''' <summary>正規表現検索文字列([]で囲まれた項目番号の検索)</summary>
    Public Const C_SearchKakko As String = "\[前?[QKC]+[0-9_]+\]"
    ''' <summary>正規表現検索文字列(RANK()で囲まれた文字列の検索)</summary>
    Private Const C_SearchRank As String = "RANK\([^\(\)]*(\([^\(\)]*(\([^\(\)]*[^\(\)]*\)[^\(\)]*)*[^\(\)]*\)[^\(\)]*)*[^\(\)]*\)"
    ''' <summary>正規表現検索文字列({}で囲まれた文字列の検索)</summary>
    Private Const C_SearchRankKakko As String = "\{[^\(\)]*[^\(\)]*\}"
    ''' <summary>正規表現検索文字列([]で囲まれた調査票項目番号の検索)</summary>
    Private Const C_SearchChosa As String = "\[前?[Q]+[0-9_]+\]"
    ''' <summary>正規表現検索文字列([]で囲まれた可変項目番号の検索)</summary>
    Private Const C_SearchVariable As String = "\[前?[Q]+[0-9]+[_]+[0-9]+\]"

    ''' <summary>DBAccessオブジェクト</summary>
    Private _db As DBAccess

    ''' <summary>調査区分</summary>
    Private _chosaKubun As String

    ''' <summary>調査年</summary>
    Private _chosaYear As String

    ''' <summary>作成種別</summary>
    Private _createType As enmCreateType

    ''' <summary>調査票項目マスタ</summary>
    Private _dtChosaItemMst As DataTable

    ''' <summary>個別結果表項目マスタ</summary>
    Private _dtKobetsuItemMst As DataTable

    ''' <summary>個別結果検討表項目マスタ</summary>
    Private _dtKobetsuKentoItemMst As DataTable

    ''' <summary>項目情報リスト</summary>
    Private _ItemInfoList As List(Of ItemInfo)

    ''' <summary>エラーチェックかどうか</summary>
    Private _isErrCheck As Boolean

    ''' <summary>作成論理の計算式に可変項目が存在するかどうか</summary>
    Private _isExistVariable As Boolean

    ''' <summary>作成論理の計算式に調査票項目が存在するかどうか</summary>
    Private _isExistChosahyoItemNo As Boolean

    ''' <summary>進捗ダイアログ</summary>
    Private _progressDialog As ProgressDialog

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="pDb">DBAccessオブジェクト</param>
    ''' <param name="pChosaKubun">調査区分</param>
    ''' <param name="pChosaYear">調査年</param>
    ''' <param name="pCreateType">作成種別</param>
    ''' <param name="pDtChosaItemMst">調査票項目マスタDataTable</param>
    ''' <param name="pDtKobetsuItemMst">個別結果表項目マスタDataTable</param>
    ''' <param name="pDtKobetsuKentoItemMst">個別結果検討表項目マスタDataTable</param>
    ''' <param name="pDtCreateRonri">作成論理DataTable(個別結果表or個別結果検討表)</param>
    ''' <param name="pKobetsuList">個別結果表値リスト(個別結果表再計算時のみ)</param>
    ''' <param name="pProgressDialog">進捗ダイアログ</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pDb As DBAccess,
                   ByVal pChosaKubun As String,
                   ByVal pChosaYear As String,
                   ByVal pCreateType As enmCreateType,
                   ByVal pDtChosaItemMst As DataTable,
                   ByVal pDtKobetsuItemMst As DataTable,
                   ByVal pDtKobetsuKentoItemMst As DataTable,
                   ByVal pDtCreateRonri As DataTable,
                   ByVal pKobetsuList As Dictionary(Of String, Object),
                   ByVal pProgressDialog As ProgressDialog)
        Try
            _ItemInfoList = New List(Of ItemInfo)

            'DBAccessオブジェクト
            _db = pDb
            '調査区分
            _chosaKubun = pChosaKubun
            '調査年
            _chosaYear = pChosaYear
            '作成種別
            _createType = pCreateType
            '進捗ダイアログ
            _progressDialog = pProgressDialog

            '作成論理の計算式に可変項目が存在するかどうか判定する
            _isExistVariable = Me.IsExistVariable(pDtCreateRonri)

            '作成論理の計算式に調査票項目が存在するかどうか判定する
            _isExistChosahyoItemNo = Me.IsExistChosahyoItemNo(pDtCreateRonri)

            '調査票項目マスタ
            If pDtChosaItemMst IsNot Nothing Then
                For Each dr As DataRow In pDtChosaItemMst.Rows
                    If CInt(dr("可変区分")) = 0 Then
                        _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString,
                                                             .ValueType = dr("型区分").ToString})
                    Else
                        '作成論理に可変項目が存在する場合
                        If _isExistVariable Then
                            For meisaiNo As Integer = 1 To CInt(dr("可変最大数"))
                                _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & meisaiNo,
                                                                     .ValueType = dr("型区分").ToString,
                                                                     .ItemType = enmItemType.調査票})
                            Next
                        End If
                    End If
                Next
            End If

            '個別結果表項目マスタ
            If pDtKobetsuItemMst IsNot Nothing Then
                For Each dr As DataRow In pDtKobetsuItemMst.Rows
                    _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString,
                                                         .SheetName = dr("シート名").ToString,
                                                         .Row = CInt(dr("行位置")),
                                                         .Col = CInt(dr("列位置")),
                                                         .ValueType = dr("型区分").ToString,
                                                         .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                         .ItemType = enmItemType.個別結果表})
                Next
            End If

            '個別結果検討表項目マスタ
            If pDtKobetsuKentoItemMst IsNot Nothing Then
                For Each dr As DataRow In pDtKobetsuKentoItemMst.Rows
                    _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString,
                                                         .SheetName = dr("シート名").ToString,
                                                         .Row = CInt(dr("行位置")),
                                                         .Col = CInt(dr("列位置")),
                                                         .ValueType = dr("型区分").ToString,
                                                         .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                         .ItemType = enmItemType.個別結果検討表})
                Next
            End If

            '作成論理の情報更新
            For Each dr As DataRow In pDtCreateRonri.Rows
                Dim info As ItemInfo = (From n In _ItemInfoList Where n.ItemNo = dr("項目番号").ToString).First
                info.ItemName = dr("項目名").ToString
                info.Priority = CDec(If(IsDBNull(dr("優先順位")), 0, dr("優先順位")))
                info.Formula = dr("計算式").ToString
                If Not IsDBNull(dr("表示単位")) Then
                    info.HyojiTani = CDec(dr("表示単位"))
                End If
                info.IsReCalc = If(_createType = enmCreateType.個別結果表再計算, True, False)
            Next

            '計算式の文字列置換(項番以外)
            For Each item As ItemInfo In (From n In _ItemInfoList Where Not String.IsNullOrEmpty(n.Formula)).ToList
                item.FormulaConv = Me.ConvertFormula(item.Formula)
            Next

            Select Case (_createType)
                Case enmCreateType.個別結果表再計算
                    '項目情報リストに値をセットする
                    For Each info As ItemInfo In (From n In _ItemInfoList Where n.ItemType = enmItemType.個別結果表).ToList
                        If pKobetsuList.ContainsKey(info.ItemNo) Then
                            Dim value As Object = pKobetsuList(info.ItemNo)
                            info.Value = value
                        End If
                    Next

                Case enmCreateType.個別結果検討表作成
                    '個別結果表の前年値項目情報を追加
                    Dim itemInfoListZen As New List(Of ItemInfo)
                    For Each info As ItemInfo In (From n In _ItemInfoList Where n.ItemType = enmItemType.個別結果表).ToList
                        '前年値の数式は実行しないのでNothingとする
                        itemInfoListZen.Add(New ItemInfo With {.ItemNo = "前" & info.ItemNo,
                                                               .ItemName = info.ItemName,
                                                               .Priority = info.Priority,
                                                               .Formula = Nothing,
                                                               .FormulaConv = Nothing,
                                                               .FormulaConvValue = Nothing,
                                                               .HyojiTani = info.HyojiTani,
                                                               .ValueType = info.ValueType,
                                                               .IsHidden = info.IsHidden,
                                                               .ItemType = info.ItemType})
                    Next
                    _ItemInfoList.AddRange(itemInfoListZen)
            End Select

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' コンストラクタ(エラーチェック用)
    ''' </summary>
    ''' <param name="pDb">DBAccessオブジェクト</param>
    ''' <param name="pChosaKubun">調査区分</param>
    ''' <param name="pChosaYear">調査年</param>
    ''' <param name="pCreateType">作成種別</param>
    ''' <param name="pDtChosaItemMst">調査票項目マスタDataTable</param>
    ''' <param name="pDtKobetsuItemMst">個別結果表項目マスタDataTable</param>
    ''' <param name="pDtKobetsuKentoItemMst">個別結果検討表項目マスタDataTable</param>
    ''' <param name="pDtCreateRonri">作成論理DataTable(個別結果表or個別結果検討表)</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pDb As DBAccess,
                   ByVal pChosaKubun As String,
                   ByVal pChosaYear As String,
                   ByVal pCreateType As enmCreateType,
                   ByVal pDtChosaItemMst As DataTable,
                   ByVal pDtKobetsuItemMst As DataTable,
                   ByVal pDtKobetsuKentoItemMst As DataTable,
                   ByVal pDtCreateRonri As DataTable)
        Try
            _ItemInfoList = New List(Of ItemInfo)

            'DBAccessオブジェクト
            _db = pDb
            '調査区分
            _chosaKubun = pChosaKubun
            '調査年
            _chosaYear = pChosaYear
            '作成種別
            _createType = pCreateType
            '調査票項目マスタ
            _dtChosaItemMst = pDtChosaItemMst
            '個別結果表項目マスタ
            _dtKobetsuItemMst = pDtKobetsuItemMst
            '個別結果検討表項目マスタ
            _dtKobetsuKentoItemMst = pDtKobetsuKentoItemMst
            'エラーチェックかどうか
            _isErrCheck = True

            '作成論理の計算式に可変項目が存在するかどうか判定する
            _isExistVariable = Me.IsExistVariable(pDtCreateRonri)

            '調査票項目マスタ
            If pDtChosaItemMst IsNot Nothing Then
                For Each dr As DataRow In pDtChosaItemMst.Rows
                    If CInt(dr("可変区分")) = 0 Then
                        _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString,
                                                             .ValueType = dr("型区分").ToString})
                    Else
                        '作成論理に可変項目が存在する場合
                        If _isExistVariable Then
                            For meisaiNo As Integer = 1 To CInt(dr("可変最大数"))
                                _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & meisaiNo,
                                                                     .ValueType = dr("型区分").ToString,
                                                                     .ItemType = enmItemType.調査票})
                            Next
                        End If
                    End If
                Next
            End If

            '個別結果表項目マスタ
            If pDtKobetsuItemMst IsNot Nothing Then
                For Each dr As DataRow In pDtKobetsuItemMst.Rows
                    _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString,
                                                         .SheetName = dr("シート名").ToString,
                                                         .Row = CInt(dr("行位置")),
                                                         .Col = CInt(dr("列位置")),
                                                         .ValueType = dr("型区分").ToString,
                                                         .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                         .ItemType = enmItemType.個別結果表})
                Next
            End If

            '個別結果検討表項目マスタ
            If pDtKobetsuKentoItemMst IsNot Nothing Then
                For Each dr As DataRow In pDtKobetsuKentoItemMst.Rows
                    _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString,
                                                         .SheetName = dr("シート名").ToString,
                                                         .Row = CInt(dr("行位置")),
                                                         .Col = CInt(dr("列位置")),
                                                         .ValueType = dr("型区分").ToString,
                                                         .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                         .ItemType = enmItemType.個別結果検討表})
                Next
            End If

            If _createType = enmCreateType.個別結果検討表作成 Then
                '個別結果表の前年値項目情報を追加
                Dim itemInfoListZen As New List(Of ItemInfo)
                For Each info In (From n In _ItemInfoList Where n.ItemType = enmItemType.個別結果表).ToList
                    '前年値の数式は実行しないのでNothingとする
                    itemInfoListZen.Add(New ItemInfo With {.ItemNo = "前" & info.ItemNo,
                                                           .ItemName = info.ItemName,
                                                           .Priority = info.Priority,
                                                           .Formula = Nothing,
                                                           .FormulaConv = Nothing,
                                                           .FormulaConvValue = Nothing,
                                                           .HyojiTani = info.HyojiTani,
                                                           .ValueType = info.ValueType,
                                                           .IsHidden = info.IsHidden,
                                                           .ItemType = info.ItemType})
                Next
                _ItemInfoList.AddRange(itemInfoListZen)
            End If

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 作成実行
    ''' </summary>
    ''' <param name="pCensusNo">センサス番号</param>
    ''' <param name="isGetTaisyakuKubun">営農個人の貸借対照表区分取得時の呼び出し</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Execute(ByVal pCensusNo As String, Optional isGetTaisyakuKubun As Boolean = False) As List(Of ItemInfo) 'REV_002 CNG DAIKO
        Dim itemInfoListRonri As List(Of ItemInfo)
        Dim rochinTanka As Integer
        Dim nogyoChikiKbn1 As Integer
        Dim nogyoChikiKbn2 As Integer
        Dim nogyoChikiKbn As Integer
        Dim dtNogyoChikiKbn As DataTable

        Try

            '-----------------------------------------------------------------------------------------------------
            '初期化
            '-----------------------------------------------------------------------------------------------------
            If Not _isErrCheck Then
                Select Case (_createType)
                    Case enmCreateType.個別結果表作成, enmCreateType.個別結果検討表作成
                        For Each ItemInfo In _ItemInfoList
                            ItemInfo.Value = Nothing
                        Next
                End Select
            End If

            '作成論理が存在するリストを作成
            itemInfoListRonri = (From n In _ItemInfoList Where Not String.IsNullOrEmpty(n.Formula)).ToList
            For Each itemInfoRonri In itemInfoListRonri
                itemInfoRonri.FormulaConvValue = itemInfoRonri.FormulaConv
            Next

            '-----------------------------------------------------------------------------------------------------
            '労賃単価、農業地域区分の置換
            '-----------------------------------------------------------------------------------------------------

            If Not _isErrCheck And Not isGetTaisyakuKubun Then 'REV_002 CNG DAIKO 営農個人の貸借対照表区分取得時の呼び出しでは不要な処理のため条件追加
                '事務所番号取得
                Dim jimusho As String = ComUtil.ConvJimusyoNo(ComUtil.GetTodofuken(pCensusNo))
                '市区町村取得
                Dim shikuchoson As String = ComUtil.GetShikuchoson(pCensusNo)
                '旧市区町村取得
                Dim kyushikuchoson As String = ComUtil.GetKyuShikuchoson(pCensusNo)
                '都道府県取得
                Dim todofuken As String = ComUtil.GetTodofuken(pCensusNo)

                '労賃単価取得
                rochinTanka = GetRochinTanka(_db, jimusho)

                'REV_002 ADD START
                '調査年を農業地域類型マスタ調査年に変換
                Dim year = ComUtil.NogyoChikiMstMainte.GetNogyoChikiMstYear(_chosaYear, _chosaKubun)
                'REV_002 ADD END

                dtNogyoChikiKbn = DAOOther.GetNogyoChikiKbn(_db, todofuken, shikuchoson, kyushikuchoson, year) 'REV_002 CNG
                If dtNogyoChikiKbn.Rows.Count > 0 Then
                    nogyoChikiKbn1 = CInt(dtNogyoChikiKbn.Rows(0).Item("第１次分類"))
                    nogyoChikiKbn2 = CInt(dtNogyoChikiKbn.Rows(0).Item("第２次分類"))
                    nogyoChikiKbn = CInt(nogyoChikiKbn1 & nogyoChikiKbn2)
                End If
            End If

            For Each info As ItemInfo In (From n In itemInfoListRonri Where n.Formula.Contains("【"))
                '『【労賃単価】』を値に置換
                info.FormulaConvValue = info.FormulaConvValue.Replace("【労賃単価】", rochinTanka.ToString)
                '『【農業地域区分１次】』を値に置換
                info.FormulaConvValue = info.FormulaConvValue.Replace("【農業地域区分１次】", nogyoChikiKbn1.ToString)
                '『【農業地域区分２次】』を値に置換
                info.FormulaConvValue = info.FormulaConvValue.Replace("【農業地域区分２次】", nogyoChikiKbn2.ToString)
                '『【農業地域区分】』を値に置換
                info.FormulaConvValue = info.FormulaConvValue.Replace("【農業地域区分】", nogyoChikiKbn.ToString)
            Next

            '-----------------------------------------------------------------------------------------------------
            '当年と前年のデータ取得
            '-----------------------------------------------------------------------------------------------------
            If Not _isErrCheck Then
                Select Case (_createType)
                    Case enmCreateType.個別結果表作成
                        '調査票データを取得し、項目情報リストに値をセットする
                        Me.GetChosahyo(_db, New DAOChosahyo.PrimaryKey(_chosaYear, pCensusNo), _ItemInfoList)

                    Case enmCreateType.個別結果表再計算
                        '作成論理に調査票項目が存在する場合
                        If _isExistChosahyoItemNo Then
                            '調査票データを取得し、項目情報リストに値をセットする
                            Me.GetChosahyo(_db, New DAOChosahyo.PrimaryKey(_chosaYear, pCensusNo), _ItemInfoList)
                        End If

                    Case enmCreateType.個別結果検討表作成
                        '個別結果表データを取得し、項目情報リストに値をセットする
                        Me.GetKobetsukekkahyo(_db, New DAOKobetsuKekkahyo.PrimaryKey(_chosaYear, pCensusNo), _ItemInfoList)
                        '個別結果表データを取得し、項目情報リストに値をセットする(前年データ)
                        ' REV_001↓
                        ''(前回センサス番号から前年値を取得する必要はない)
                        'Me.GetKobetsukekkahyo(_db, New DAOKobetsuKekkahyo.PrimaryKey(CStr(CInt(_chosaYear) - 1), pCensusNo), _ItemInfoList, True)
                        Dim censusNo As String = pCensusNo
                        If (CInt(_chosaYear) = 2022 And ComConst.令和４年体系.対象調査区分2022.IndexOf(_chosaKubun) > -1) _
                            Or (CInt(_chosaYear) = 2023 And ComConst.令和４年体系.対象調査区分2023.IndexOf(_chosaKubun) > -1) Then

                            Dim info As ItemInfo = (From n In _ItemInfoList Where n.ItemNo = ComConst.個別結果表.前回センサス番号(_chosaKubun)).FirstOrDefault
                            censusNo = If(IsDBNull(info.Value) OrElse info.Value Is Nothing, pCensusNo, CType(info.Value, String))
                        End If
                        Me.GetKobetsukekkahyo(_db, New DAOKobetsuKekkahyo.PrimaryKey(CStr(CInt(_chosaYear) - 1), censusNo), _ItemInfoList, True)
                        ' REV_001↑
                End Select
            End If

            '作成論理を「優先順位(昇順)」で並び替える
            Dim grpItemInfoList As List(Of ItemInfo) = (From n In itemInfoListRonri Order By n.Priority).ToList
            For Each grpInfo In grpItemInfoList
                Try
                    '計算式文字列の項目番号を値に置換する
                    Me.ConvertFormulaItemNoToValue(_ItemInfoList, grpInfo)

                    '計算式文字列の順位関連文字列を値に置換する
                    Me.ConvertFormulaRankToValue(grpInfo)

                    '作成論理をSQLで実行し、値を取得する
                    Dim value As Object = Me.ExecuteSQL(_db, grpInfo)
                    '型区分が数値の場合
                    If grpInfo.ValueType = ComConst.型区分.数値型 Then
                        If IsDBNull(value) OrElse value Is Nothing Then
                            grpInfo.Value = Nothing
                        Else
                            If grpInfo.HyojiTani Is Nothing Then
                                grpInfo.Value = CDec(value)
                            Else
                                '表示単位で四捨五入する
                                Dim roundValue As Decimal = Me.Round(CDec(value), grpInfo.HyojiTani)
                                If roundValue = 0 Then
                                    '値が「0」の場合、NULLを格納する
                                    grpInfo.Value = Nothing
                                Else
                                    grpInfo.Value = roundValue
                                End If
                            End If
                        End If
                    Else
                        grpInfo.Value = If(IsDBNull(value) OrElse String.IsNullOrEmpty(CStr(value)), Nothing, value)
                    End If
                Catch ex As Exception
                    '個別結果表作成エラー
                    Throw New CreateKobetsuException(ex, pCensusNo, grpInfo.ItemNo, grpInfo.Formula, grpInfo.FormulaConvValue)
                End Try

                If _progressDialog IsNot Nothing Then
                    '進捗を進める
                    _progressDialog.AddValue = 1
                End If
            Next

            Debug.Print("作成完了")

        Catch ex As Exception
            Throw
        End Try

        Return _ItemInfoList
    End Function

    ''' <summary>
    ''' 「計算式」で使用されている項目番号が項目マスタに存在するかをチェックする
    ''' </summary>
    ''' <param name="pItemNo"></param>
    ''' <param name="pFormula"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckExistItemNo(ByVal pItemNo As String, ByVal pFormula As String) As Boolean
        Dim ret As Boolean = True
        Dim dt As DataTable = Nothing

        Try

            If _createType = enmCreateType.個別結果表作成 Then
                '項目マスタがデータ無しの場合
                If _dtChosaItemMst Is Nothing OrElse _dtChosaItemMst.Rows.Count = 0 OrElse
                   _dtKobetsuItemMst Is Nothing OrElse _dtKobetsuItemMst.Rows.Count = 0 Then
                    ret = False
                    Exit Try
                End If
            Else
                '項目マスタがデータ無しの場合
                If _dtKobetsuItemMst Is Nothing OrElse _dtKobetsuItemMst.Rows.Count = 0 OrElse
                   _dtKobetsuKentoItemMst Is Nothing OrElse _dtKobetsuKentoItemMst.Rows.Count = 0 Then
                    ret = False
                    Exit Try
                End If
            End If

            '『[』『]』で囲まれた文字列検索する
            Dim matchList As MatchCollection = Regex.Matches(pFormula, C_SearchKakko)
            For Each mat As Match In matchList
                Dim itemNo As String = mat.Value.Trim("["c, "]"c).Replace("前", "")
                Dim isVariable As Boolean = False

                If _createType = enmCreateType.個別結果表作成 Then
                    Select Case (itemNo.Substring(0, 1))
                        Case "Q"
                            dt = _dtChosaItemMst
                            If itemNo.Contains(ComConst.ITEM_NO_DELIMITER) Then
                                '可変項目である
                                isVariable = True
                            End If

                        Case "K"
                            dt = _dtKobetsuItemMst
                        Case Else
                            ret = False
                            Exit For
                    End Select
                Else
                    Select Case (itemNo.Substring(0, 1))
                        Case "K"
                            dt = _dtKobetsuItemMst
                        Case "C"
                            dt = _dtKobetsuKentoItemMst
                        Case Else
                            ret = False
                            Exit For
                    End Select
                End If

                '可変項目の場合
                If isVariable Then
                    Dim variItemNo As String = itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(0)
                    Dim meisaiNo As Integer = CType(itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(1), Integer)
                    '可変項目番号が項目マスタに存在しない場合
                    If Not (From n In dt.AsEnumerable Where n("項目番号").ToString = variItemNo).Any Then
                        ret = False
                        Exit For
                    End If
                    '可変項目番号が項目マスタに存在し、明細番号が可変最大数を超えている場合
                    If (From n In dt.AsEnumerable Where n("項目番号").ToString = variItemNo And CInt(n("可変最大数")) < meisaiNo).Any Then
                        ret = False
                        Exit For
                    End If
                Else
                    '項目番号が項目マスタに存在しない場合
                    If Not (From n In dt.AsEnumerable Where n("項目番号").ToString = itemNo).Any Then
                        ret = False
                        Exit For
                    End If
                End If
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 「計算式」がSQLで実行可能かをチェックする
    ''' </summary>
    ''' <param name="pItemNo"></param>
    ''' <param name="pFormula"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckExecutableSQL(ByVal pItemNo As String, ByVal pFormula As String) As Boolean
        Dim ret As Boolean
        Dim itemNo As String
        Dim infoTarget As ItemInfo

        Try

            itemNo = pItemNo.Trim("["c, "]"c)

            '「計算式」で使用されている項目番号が項目マスタに存在しない場合
            If Not Me.CheckExistItemNo(itemNo, pFormula) Then
                ret = False
                Exit Try
            End If

            infoTarget = (From n In _ItemInfoList Where n.ItemNo = itemNo).FirstOrDefault
            If infoTarget Is Nothing Then
                '項目情報リスト(項目マスタ)に存在しない場合、項目の型が不明なのでエラーとして返す
                ret = False
                Exit Try
            Else
                infoTarget.Formula = pFormula
                '計算式の文字列置換(項番以外)
                infoTarget.FormulaConv = Me.ConvertFormula(pFormula)
            End If

            Try

                '作成実行
                Execute(Nothing)
                ret = True

            Catch ex As Exception
                ret = False
            Finally
                '次回実行されないように数式をクリアする
                infoTarget.Formula = Nothing
                infoTarget.FormulaConv = Nothing
                infoTarget.FormulaConvValue = Nothing
            End Try

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 労賃単価を取得する
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pJimusho"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetRochinTanka(ByVal pDb As DBAccess, ByVal pJimusho As String) As Integer
        Dim ret As Integer

        Try

            '生産費区分取得
            Dim seisanhiKbn As String = ComUtil.RouchinTanka.GetSeisanhiKbn(_chosaKubun)
            If Not String.IsNullOrEmpty(seisanhiKbn) Then
                '労賃単価取得
                Dim dt As DataTable = DAOOther.GetRouchinTanka(pDb, seisanhiKbn, _chosaYear, pJimusho)
                If dt.Rows.Count > 0 Then
                    ret = CInt(dt.Rows(0).Item("男女平均＿評価単価"))
                End If
            End If

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 計算式の文字列置換する(項目番号以外)
    ''' </summary>
    ''' <param name="pShinsaNaiyo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConvertFormula(ByVal pFormula As String) As String
        Dim ret As String = Nothing

        Try
            ret = pFormula

            '『]＝空白』 ⇒ 『]@ IS NULL』
            ret = ret.Replace("]＝空白", "]@ IS NULL ")
            '『]≠空白』 ⇒ 『]@ IS NOT NULL』
            ret = ret.Replace("]≠空白", "]@ IS NOT NULL ")
            '『空白』 ⇒ 『NULL』
            ret = ret.Replace("空白", "NULL")
            '全角不等号『＜』『＞』 ⇒ 半角不等号『<』『>』
            ret = ret.Replace("＜", "<").Replace("＞", ">")
            '全角等号付き不等号『≦』『≧』 ⇒ 『<=』『>=』
            ret = ret.Replace("≦", "<=").Replace("≧", ">=")
            '全角等号『＝』 ⇒ 半角等号『=』
            ret = ret.Replace("＝", "=")
            '全角等号否定『≠』 ⇒ 半角山括弧『<>』
            ret = ret.Replace("≠", "<>")
            '全角記号『＋－÷×』 ⇒ 半角記号『+-/*』
            '『－』は後ろの数値がマイナスの時に『--』となる(コメント扱いになる)のを防ぐため後ろに半角スペースを入れる
            ret = ret.Replace("＋", "+").Replace("－", "- ").Replace("-", "- ").Replace("÷", "/").Replace("×", "*")

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 調査票データを取得し、項目情報リストに値をセットする
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pPKey"></param>
    ''' <param name="pItemInfoList"></param>
    ''' <remarks></remarks>
    Private Sub GetChosahyo(ByVal pDb As DBAccess, ByVal pPKey As DAOChosahyo.PrimaryKey, ByRef pItemInfoList As List(Of ItemInfo))
        Dim dicChosahyo As Dictionary(Of String, DataTable) = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim dc As DataColumn = Nothing

        Try

            '調査票データリスト取得
            dicChosahyo = DAOChosahyo.GetChosahyoTable(pDb, pPKey)
            For Each kv As KeyValuePair(Of String, DataTable) In dicChosahyo
                Dim tableName As String = kv.Key
                dt = kv.Value
                If Not tableName.Contains("＿可変") Then
                    If dt.Rows.Count > 0 Then
                        dr = dt.Rows(0)
                        For Each dc In dt.Columns
                            '項目情報リストの項目番号が一致するものに値をセット
                            SetItemInfoListValue(pItemInfoList, dc.ColumnName, dr(dc.ColumnName))
                        Next
                    End If
                Else
                    '作成論理に可変項目が存在する場合
                    If _isExistVariable Then
                        For Each dr In dt.Rows
                            '項目情報リストの項目番号が一致するものに値をセット
                            SetItemInfoListValue(pItemInfoList, dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & dr("明細番号").ToString, dr("値"))
                        Next
                    End If
                End If
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 個別結果表データを取得し、項目情報リストに値をセットする
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pPKey"></param>
    ''' <param name="pItemInfoList"></param>
    ''' <param name="isZennen">前年データかどうか</param>
    ''' <remarks></remarks>
    Private Sub GetKobetsukekkahyo(ByVal pDb As DBAccess,
                                   ByVal pPKey As DAOKobetsuKekkahyo.PrimaryKey,
                                   ByRef pItemInfoList As List(Of ItemInfo),
                                   Optional ByVal isZennen As Boolean = False)
        Dim dicKobetsu As Dictionary(Of String, DataTable) = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim dc As DataColumn = Nothing

        Try

            '個別結果表テーブルリスト取得
            dicKobetsu = DAOKobetsuKekkahyo.GetTable(pDb, pPKey)
            For Each kv As KeyValuePair(Of String, DataTable) In dicKobetsu
                dt = kv.Value
                If dt.Rows.Count > 0 Then
                    dr = dt.Rows(0)
                    For Each dc In dt.Columns
                        '項目情報リストの項目番号が一致するものに値をセット
                        SetItemInfoListValue(pItemInfoList, If(isZennen, "前" & dc.ColumnName, dc.ColumnName), dr(dc.ColumnName))
                    Next
                End If
            Next

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' 指定した項目番号と一致する項目情報リストのデータに値をセットする
    ''' </summary>
    ''' <param name="pItemInfoList"></param>
    ''' <param name="ItemNo"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Private Sub SetItemInfoListValue(ByRef pItemInfoList As List(Of ItemInfo), ByVal ItemNo As String, ByVal value As Object)
        Dim info As ItemInfo = (From n In pItemInfoList Where n.ItemNo = ItemNo).FirstOrDefault
        If info IsNot Nothing Then
            info.Value = value
        End If
    End Sub

    ''' <summary>
    ''' 計算式文字列の項目番号を値に置換する
    ''' </summary>
    ''' <param name="pItemInfo"></param>
    ''' <param name="pFormula"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Sub ConvertFormulaItemNoToValue(ByVal pItemInfoList As List(Of ItemInfo), ByRef pItemInfo As ItemInfo)
        Dim matchList As MatchCollection = Nothing

        '括弧付き項目番号『[010101]』『[010101_1]』を検索
        matchList = Regex.Matches(pItemInfo.FormulaConvValue, C_SearchKakko)
        '重複しない括弧付き項目番号リストを作成
        Dim itemNoKakkoList As String() = (From m In matchList.Cast(Of Match)() Select m.Value).Distinct().ToArray()
        For Each itemNoKakko As String In itemNoKakkoList
            '『[010101]』『[010101_1]』から項目番号『010101』『010101_1』を抽出する
            Dim itemNo As String = itemNoKakko.Trim("["c, "]"c)
            Dim info As ItemInfo = (From n In pItemInfoList Where n.ItemNo = itemNo).FirstOrDefault
            'データが存在する場合
            If info IsNot Nothing Then
                'エラーチェック時は文字列型には任意の文字をセットする(文字列と数値の比較『''<0』等がエラーにならないため)
                Dim value As Object = If(info.ValueType = ComConst.型区分.文字型 AndAlso _isErrCheck, "a", info.Value)
                '@付き項目番号『[010101]@』『[010101_1]@』を値に置換
                If info.ValueType = ComConst.型区分.数値型 Then
                    pItemInfo.FormulaConvValue = pItemInfo.FormulaConvValue.Replace("[" & itemNo & "]@", If(IsDBNull(value) OrElse value Is Nothing, "NULL", Me.CnvIntToDec(If(TypeOf value Is Decimal AndAlso Decimal.Truncate(CDec(value)) = CDec(value), Decimal.Truncate(CDec(value)).ToString, value.ToString))))
                Else
                    pItemInfo.FormulaConvValue = pItemInfo.FormulaConvValue.Replace("[" & itemNo & "]@", If(IsDBNull(value) OrElse value Is Nothing, "NULL", "'" & value.ToString & "'"))
                End If
                '項目番号『[010101]』『[010101_1]』を値に置換
                If info.ValueType = ComConst.型区分.数値型 Then
                    pItemInfo.FormulaConvValue = pItemInfo.FormulaConvValue.Replace("[" & itemNo & "]", If(IsDBNull(value) OrElse value Is Nothing OrElse value Is String.Empty, "0", Me.CnvIntToDec(If(TypeOf value Is Decimal AndAlso Decimal.Truncate(CDec(value)) = CDec(value), Decimal.Truncate(CDec(value)).ToString, value.ToString))))
                Else
                    pItemInfo.FormulaConvValue = pItemInfo.FormulaConvValue.Replace("[" & itemNo & "]", If(IsDBNull(value) OrElse value Is Nothing OrElse value Is String.Empty, "''", "'" & value.ToString & "'"))
                End If
            End If
        Next

    End Sub

    ''' <summary>
    ''' 計算式文字列の順位関連文字列を値に置換する
    ''' </summary>
    ''' <param name="pItemInfo"></param>
    ''' <param name="pFormula"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Sub ConvertFormulaRankToValue(ByRef pItemInfo As ItemInfo)
        Dim matchItemNoList As MatchCollection = Nothing
        Dim matchValueList As MatchCollection = Nothing

        '値に変換する前の計算式から『RANK(』『)』で囲まれた文字列を検索する(大文字と小文字を区別しない)
        matchItemNoList = Regex.Matches(pItemInfo.FormulaConv, C_SearchRank, RegexOptions.IgnoreCase)
        '値に変換した後の計算式から『RANK(』『)』で囲まれた文字列を検索する(大文字と小文字を区別しない)
        matchValueList = Regex.Matches(pItemInfo.FormulaConvValue, C_SearchRank, RegexOptions.IgnoreCase)
        For i As Integer = 0 To matchValueList.Count - 1
            '「項目の順位」を取得する文字列を値に置換する
            pItemInfo.FormulaConvValue = Me.ConvertRankToValue(pItemInfo.FormulaConvValue, matchItemNoList(i).Value, matchValueList(i).Value)
        Next
    End Sub

    ''' <summary>
    ''' 「項目の順位」を取得する文字列を値に置換する
    ''' </summary>
    ''' <param name="pFormulaConvValue"></param>
    ''' <param name="pRankItemNo">『RANK(』『)』で囲まれた計算式(値に変換する前)</param>
    ''' <param name="pRankValue">『RANK(』『)』で囲まれた計算式(値に変換した後)</param>
    ''' <returns></returns>
    ''' <remarks>対象項目番号の値が複数存在した場合、計算式に出現した順番とする(同順位を許さない)</remarks>
    Public Function ConvertRankToValue(ByVal pFormulaConvValue As String, ByVal pRankItemNo As String, ByVal pRankValue As String) As String
        Dim ret As String = Nothing
        Dim myStr As String
        Dim baseItemNo As String
        Dim baseValue As Decimal
        Dim matchList As MatchCollection = Nothing
        Dim itemNoList() As String = Nothing
        Dim valueList() As String = Nothing
        Dim value As String = Nothing
        Dim rank As Integer
        Dim list As New List(Of RankInfo)
        Dim sort As Integer

        Try
            '順位付けする項目番号を取得(『RANK(』『)』で囲まれた文字列の『RANK(』の後ろから『,』の前までの文字列を取得)
            baseItemNo = pRankItemNo.Substring(pRankItemNo.IndexOf("(") + 1, pRankItemNo.IndexOf(",") - (pRankItemNo.IndexOf("(") + 1))
            '順位対象の項目番号リストを取得(『RANK(』『)』で囲まれた文字列の『,』の後ろから『{』『}』で囲まれた文字列を取得)
            matchList = Regex.Matches(pRankItemNo.Substring(pRankItemNo.IndexOf(",") + 1), C_SearchRankKakko)
            '前後の『{』『}』を除去した文字列を『,』で分割
            itemNoList = matchList(0).Value.Trim("{"c, "}"c).Split(","c)

            '順位付けする値を取得(『RANK(』『)』で囲まれた文字列の『RANK(』の後ろから『,』の前までの文字列を取得)
            myStr = pRankValue.Substring(pRankValue.IndexOf("(") + 1, pRankValue.IndexOf(",") - (pRankValue.IndexOf("(") + 1))
            baseValue = CDec(myStr)
            '順位対象の値リストを取得(『RANK(』『)』で囲まれた文字列の『,』の後ろから『{』『}』で囲まれた文字列を取得)
            matchList = Regex.Matches(pRankValue.Substring(pRankValue.IndexOf(",") + 1), C_SearchRankKakko)
            '前後の『{』『}』を除去した文字列を『,』で分割
            valueList = matchList(0).Value.Trim("{"c, "}"c).Split(","c)
            '順位対象値リストを作成
            For i As Integer = 0 To valueList.Count - 1
                If Decimal.TryParse(valueList(i), New Decimal) Then
                    list.Add(New RankInfo With {.Number = i + 1, .ItemNo = itemNoList(i), .Value = Decimal.Parse(valueList(i))})
                End If
            Next

            '並び順を取得(『RANKNAME(』『)』で囲まれた文字列の末尾『)』を除去した文字列を後ろから検索して、『,』の後ろの文字列を取得)
            myStr = pRankItemNo.Trim(")"c).Substring(pRankItemNo.LastIndexOf(",") + 1)
            sort = CInt(myStr)
            '降順の場合
            If sort = 0 Then
                list = (From n In list Order By n.Value Descending, n.Number).ToList
            Else
                list = (From n In list Order By n.Value, n.Number).ToList
            End If

            For i As Integer = 0 To list.Count - 1
                '順位付けする項目番号と値を比較
                If list(i).ItemNo = baseItemNo AndAlso list(i).Value = baseValue Then
                    rank = i + 1
                    Exit For
                End If
            Next

            '『RANK(』『)』で囲まれた文字列を順位に置換する
            ret = pFormulaConvValue.Replace(pRankValue, rank.ToString)

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 作成論理をSQLで実行し、値を取得する
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pShinsaInfoList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ExecuteSQL(ByVal pDb As DBAccess, ByVal pItemInfo As ItemInfo) As Object
        Dim ret As Object
        Dim dt As DataTable = Nothing
        Dim sb As System.Text.StringBuilder = Nothing

        Try

            sb = New System.Text.StringBuilder

            'SQL文の設定
            sb.AppendLine("SET ANSI_WARNINGS OFF")
            sb.Append("SELECT " & pItemInfo.FormulaConvValue & " AS [" & pItemInfo.ItemNo & "]")

            dt = pDb.GetDataTable(sb.ToString)
            ret = dt.Rows(0)(0)

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 値を表示単位で四捨五入する
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="HyojiTani"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Round(ByVal value As Decimal, ByVal HyojiTani As Decimal?) As Decimal
        Dim iDigits As Integer
        Select Case (HyojiTani)
            Case 9D
                iDigits = 0
            Case 9.9D
                iDigits = 1
            Case 9.99D
                iDigits = 2
            Case 9.999D
                iDigits = 3
            Case Else
                iDigits = 0
        End Select
        Return Math.Round(value, iDigits, MidpointRounding.AwayFromZero)
    End Function

    ''' <summary>
    ''' Integer型の値をDecimal型に変換する
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CnvIntToDec(ByVal value As String) As String
        Dim ret As String

        'Decimal型に変換可能で、『.』が含まれていない場合
        If Decimal.TryParse(value, New Decimal) AndAlso Not value.Contains(".") Then
            ret = value & ".0"
        Else
            ret = value
        End If

        Return ret
    End Function

    ''' <summary>
    ''' オブジェクトをディープコピーする
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="target"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DeepCopy(Of T)(ByVal target As T) As T
        Dim ret As T
        Dim b As New BinaryFormatter
        Dim mem As New System.IO.MemoryStream()

        Try
            b.Serialize(mem, target)
            mem.Position = 0
            ret = CType(b.Deserialize(mem), T)
        Finally
            mem.Close()
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 作成論理の計算式に可変項目が存在するかどうか判定する
    ''' </summary>
    ''' <param name="pDtCreateRonri"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsExistVariable(ByVal pDtCreateRonri As DataTable) As Boolean
        Dim ret As Boolean = False
        Dim matchList As MatchCollection = Nothing

        If pDtCreateRonri IsNot Nothing Then
            For Each dr As DataRow In pDtCreateRonri.Rows
                '[]で囲まれた可変項目番号の検索
                matchList = Regex.Matches(dr("計算式").ToString, C_SearchVariable)
                If matchList.Count > 0 Then
                    ret = True
                    Exit For
                End If
            Next
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 作成論理の計算式に調査票項目が存在するかどうか判定する
    ''' </summary>
    ''' <param name="pDtCreateRonri"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsExistChosahyoItemNo(ByVal pDtCreateRonri As DataTable) As Boolean
        Dim ret As Boolean = False
        Dim matchList As MatchCollection = Nothing

        If pDtCreateRonri IsNot Nothing Then
            For Each dr As DataRow In pDtCreateRonri.Rows
                '[]で囲まれた調査票項目番号の検索
                matchList = Regex.Matches(dr("計算式").ToString, C_SearchChosa)
                If matchList.Count > 0 Then
                    ret = True
                    Exit For
                End If
            Next
        End If

        Return ret
    End Function
End Class
