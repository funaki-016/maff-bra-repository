'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2023.01.13 |大興電子通信        | 要件No.4－⑦ 営農の集計結果検討表
'//  REV_002   | 2023.01.12 |Daiko               | 要件No.4 制度受取金・積立金等項目名称設定処理追加
'//  REV_003   | 2023.01.25 |Daiko               | 要件No15 集計結果検討表（報告用）追加
'//*************************************************************************************************

Imports System.Text.RegularExpressions
Imports System.Runtime.Serialization.Formatters.Binary

''' <summary>
''' 集計結果検討表作成エラークラス
''' </summary>
''' <remarks></remarks>
Public Class CreateSyukeiKentohyoException
    Inherits Exception

    ''' <summary>集計番号</summary>
    Private _syukeiNo As String

    ''' <summary>集計番号</summary>
    Public ReadOnly Property SyukeiNo As String
        Get
            Return _syukeiNo
        End Get
    End Property

    ''' <summary>連番</summary>
    Private _groupKey As String

    ''' <summary>連番</summary>
    Public ReadOnly Property GroupKey As String
        Get
            Return _groupKey
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

    Public Sub New(inner As Exception, ByVal pSyukeiNo As String, ByVal pGroupKey As String, ByVal pItemNo As String, ByVal pFormula As String, ByVal pFormulaCnv As String)
        MyBase.New(String.Format("集計結果検討表作成エラー(項目番号:{0} 計算式:{1} 計算式(変換後):{2})", pItemNo, pFormula, pFormulaCnv), inner)
        _syukeiNo = pSyukeiNo
        _groupKey = pGroupKey
        _itemNo = pItemNo
    End Sub
End Class

''' <summary>
''' 集計結果検討表作成クラス
''' </summary>
''' <remarks></remarks>
Public Class CreateSyukeiKentohyo

    ''' <summary>
    ''' 項目種別
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum enmItemType
        集計結果表
        集計結果検討表
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
        ''' <summary>項目種別</summary>
        Public ItemType As enmItemType
        ''' <summary>値</summary>
        Public Value As Object
    End Class

    ''' <summary>正規表現検索文字列([]で囲まれた項目番号の検索)</summary>
    Public Const C_SearchKakko As String = "\[前?[DS]+[0-9_]+\]"

    ''' <summary>DBAccessオブジェクト</summary>
    Private _db As DBAccess

    ''' <summary>調査区分</summary>
    Private _chosaKubun As String

    ''' <summary>集計結果表項目マスタ</summary>
    Private _dtSyukeiItemMst As DataTable

    ''' <summary>集計結果検討表項目マスタ</summary>
    Private _dtSyukeiKentoItemMst As DataTable

    ''' <summary>項目情報リスト</summary>
    Private _ItemInfoList As List(Of ItemInfo)

    ''' <summary>エラーチェックかどうか</summary>
    Private _isErrCheck As Boolean

    ''' <summary>主キー（本年）</summary>
    Private _pkeyThis As DAOSyukeiKekkahyo.PrimaryKey

    ''' <summary>論理種別</summary>
    Private _logicType As ComConst.集計結果検討表作成論理.論理種別

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="pDb">DBAccessオブジェクト</param>
    ''' <param name="pChosaKubun">調査区分</param>
    ''' <param name="pDtSyukeiItemMst">集計結果表項目マスタDataTable</param>
    ''' <param name="pDtSyukeiKentoItemMst">集計結果検討表項目マスタDataTable</param>
    ''' <param name="pDtCreateRonri">集計結果検討表作成論理DataTable</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pDb As DBAccess,
                   ByVal pChosaKubun As String,
                   ByVal pDtSyukeiItemMst As DataTable,
                   ByVal pDtSyukeiKentoItemMst As DataTable,
                   ByVal pDtCreateRonri As DataTable,
                   ByVal isChikusan As Boolean,
                   ByVal nousanFlg As Boolean,　'REV_002
                   ByVal logicType As ComConst.集計結果検討表作成論理.論理種別)　'REV_003
        Try
            _ItemInfoList = New List(Of ItemInfo)

            'DBAccessオブジェクト
            _db = pDb
            '調査区分
            _chosaKubun = pChosaKubun
            '論理種別
            _logicType = logicType　'REV_003

            '集計結果表項目マスタ
            If pDtSyukeiItemMst IsNot Nothing Then
                If isChikusan Then
                    For i = 1 To If(_chosaKubun.Equals(ComConst.調査区分.牛乳生産費統計_個別) Or _chosaKubun.Equals(ComConst.調査区分.経営分析調査_牛乳生産費), 5, 4)
                        For Each dr As DataRow In pDtSyukeiItemMst.Rows
                            _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & i,
                                                                 .SheetName = dr("シート名").ToString,
                                                                 .Row = CInt(dr("行位置")),
                                                                 .Col = CInt(dr("列位置")),
                                                                 .ValueType = dr("型区分").ToString,
                                                                 .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                                 .ItemType = enmItemType.集計結果表})
                        Next
                    Next
                    ' REV_003↓
                ElseIf _logicType.Equals(ComConst.集計結果検討表作成論理.論理種別.集計結果検討表_報告用) AndAlso ComUtil.IsNousan() Then
                    For i = 1 To 4
                        For Each dr As DataRow In pDtSyukeiItemMst.Rows
                            _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & i,
                                                                 .SheetName = dr("シート名").ToString,
                                                                 .Row = CInt(dr("行位置")),
                                                                 .Col = CInt(dr("列位置")),
                                                                 .ValueType = dr("型区分").ToString,
                                                                 .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                                 .ItemType = enmItemType.集計結果表})
                        Next
                    Next
                    ' REV_003↑
                Else
                    For Each dr As DataRow In pDtSyukeiItemMst.Rows
                        _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString,
                                                             .SheetName = dr("シート名").ToString,
                                                             .Row = CInt(dr("行位置")),
                                                             .Col = CInt(dr("列位置")),
                                                             .ValueType = dr("型区分").ToString,
                                                             .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                             .ItemType = enmItemType.集計結果表})
                    Next
                End If
            End If



            '集計結果検討表項目マスタ
            If pDtSyukeiKentoItemMst IsNot Nothing Then
                For Each dr As DataRow In pDtSyukeiKentoItemMst.Rows
                    _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString,
                                                         .SheetName = dr("シート名").ToString,
                                                         .Row = CInt(dr("行位置")),
                                                         .Col = CInt(dr("列位置")),
                                                         .ValueType = dr("型区分").ToString,
                                                         .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                         .ItemType = enmItemType.集計結果検討表})
                Next
            End If

            '作成論理の情報更新
            For Each dr As DataRow In pDtCreateRonri.Rows
                Dim info As ItemInfo = (From n In _ItemInfoList Where n.ItemNo = dr("項目番号").ToString).First
                info.ItemName = dr("項目名").ToString
                info.Priority = CDec(If(IsDBNull(dr("優先順位")), 0, dr("優先順位")))
                info.Formula = dr("計算式").ToString
                'REV_002 START---------------
                'R4体系（農産）で出力項目名がある場合、任意項目名称を計算式に設定する
                If nousanFlg Then
                    If dr("出力項目名").ToString <> "" Then
                        info.Formula = "'" + dr("出力項目名").ToString + "'"
                    ElseIf Left(Trim((dr("計算式").ToString).Replace("'", "")), 4) = ComConst.制度受取金積立金等項目.任意項目名称.調査票項番 Then
                        info.Formula = ""
                    End If
                End If
                'REV_002 END---------------
                If Not IsDBNull(dr("表示単位")) Then
                    info.HyojiTani = CDec(dr("表示単位"))
                End If
            Next

            '計算式の文字列置換(項番以外)
            For Each item As ItemInfo In (From n In _ItemInfoList Where Not String.IsNullOrEmpty(n.Formula)).ToList
                item.FormulaConv = Me.ConvertFormula(item.Formula)
            Next

            '集計結果表の前年値項目情報を追加
            Dim itemInfoListZen As New List(Of ItemInfo)
            For Each info As ItemInfo In (From n In _ItemInfoList Where n.ItemType = enmItemType.集計結果表).ToList
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

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' コンストラクタ(エラーチェック用)
    ''' </summary>
    ''' <param name="pDb">DBAccessオブジェクト</param>
    ''' <param name="pChosaKubun">調査区分</param>
    ''' <param name="pDtSyukeiItemMst">集計結果表項目マスタDataTable</param>
    ''' <param name="pDtSyukeiKentoItemMst">集計結果検討表項目マスタDataTable</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pDb As DBAccess,
                   ByVal pChosaKubun As String,
                   ByVal pDtSyukeiItemMst As DataTable,
                   ByVal pDtSyukeiKentoItemMst As DataTable,
                   ByVal isChikusan As Boolean,
                   ByVal logicType As ComConst.集計結果検討表作成論理.論理種別)　'REV_003
        Try
            _ItemInfoList = New List(Of ItemInfo)

            'DBAccessオブジェクト
            _db = pDb
            '調査区分
            _chosaKubun = pChosaKubun
            '集計結果表項目マスタ
            _dtSyukeiItemMst = pDtSyukeiItemMst
            '集計結果検討表項目マスタ
            _dtSyukeiKentoItemMst = pDtSyukeiKentoItemMst
            'エラーチェックかどうか
            _isErrCheck = True
            '論理種別
            _logicType = logicType　'REV_003

            '集計結果表項目マスタ
            If pDtSyukeiItemMst IsNot Nothing Then
                If isChikusan Then
                    For i = 1 To If(_chosaKubun.Equals(ComConst.調査区分.牛乳生産費統計_個別) Or _chosaKubun.Equals(ComConst.調査区分.経営分析調査_牛乳生産費), 5, 4)
                        For Each dr As DataRow In pDtSyukeiItemMst.Rows
                            _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & i,
                                                                 .SheetName = dr("シート名").ToString,
                                                                 .Row = CInt(dr("行位置")),
                                                                 .Col = CInt(dr("列位置")),
                                                                 .ValueType = dr("型区分").ToString,
                                                                 .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                                 .ItemType = enmItemType.集計結果表})
                        Next
                    Next
                    ' REV_003↓
                ElseIf _logicType.Equals(ComConst.集計結果検討表作成論理.論理種別.集計結果検討表_報告用) AndAlso ComUtil.IsNousan() Then
                    For i = 1 To 4
                        For Each dr As DataRow In pDtSyukeiItemMst.Rows
                            _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & i,
                                                                 .SheetName = dr("シート名").ToString,
                                                                 .Row = CInt(dr("行位置")),
                                                                 .Col = CInt(dr("列位置")),
                                                                 .ValueType = dr("型区分").ToString,
                                                                 .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                                 .ItemType = enmItemType.集計結果表})
                        Next
                    Next
                    ' REV_003↑
                Else
                    For Each dr As DataRow In pDtSyukeiItemMst.Rows
                        _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString,
                                                             .SheetName = dr("シート名").ToString,
                                                             .Row = CInt(dr("行位置")),
                                                             .Col = CInt(dr("列位置")),
                                                             .ValueType = dr("型区分").ToString,
                                                             .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                             .ItemType = enmItemType.集計結果表})
                    Next
                End If
            End If

            '集計結果検討表項目マスタ
            If pDtSyukeiKentoItemMst IsNot Nothing Then
                For Each dr As DataRow In pDtSyukeiKentoItemMst.Rows
                    _ItemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString,
                                                         .SheetName = dr("シート名").ToString,
                                                         .Row = CInt(dr("行位置")),
                                                         .Col = CInt(dr("列位置")),
                                                         .ValueType = dr("型区分").ToString,
                                                         .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                         .ItemType = enmItemType.集計結果検討表})
                Next
            End If

            '集計結果表の前年値項目情報を追加
            Dim itemInfoListZen As New List(Of ItemInfo)
            For Each info In (From n In _ItemInfoList Where n.ItemType = enmItemType.集計結果表).ToList
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

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 作成実行
    ''' </summary>
    ''' <param name="pPkeyThis">主キー(本年)</param>
    ''' <param name="pPkeyLast">主キー(前年)</param>
    ''' <param name="pPkeyLast">項目キー</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Execute(ByVal pPkeyThis As DAOSyukeiKekkahyo.PrimaryKey, ByVal pPkeyLast As DAOSyukeiKekkahyo.PrimaryKey, Optional ByVal pKkey As DAOSyukeiKekkahyo.KomokuKey = Nothing) As List(Of ItemInfo)
        Dim itemInfoListRonri As List(Of ItemInfo)
        Dim syukeiNoThis As String = Nothing
        Dim syukeiNoLast As String = Nothing
        Dim syukeiNameThis As String = Nothing
        Dim syukeiNameLast As String = Nothing
        'REV_001↓
        Dim heikinShuruiThis As String = Nothing
        'REV_001↑

        Try

            '-----------------------------------------------------------------------------------------------------
            '初期化
            '-----------------------------------------------------------------------------------------------------
            If Not _isErrCheck Then
                For Each ItemInfo In _ItemInfoList
                    ItemInfo.Value = Nothing
                Next
            End If

            '作成論理が存在するリストを作成
            itemInfoListRonri = (From n In _ItemInfoList Where Not String.IsNullOrEmpty(n.Formula)).ToList
            For Each itemInfoRonri In itemInfoListRonri
                itemInfoRonri.FormulaConvValue = itemInfoRonri.FormulaConv
            Next

            '-----------------------------------------------------------------------------------------------------
            '集計番号、集計名称の置換
            '-----------------------------------------------------------------------------------------------------
            If Not _isErrCheck Then
                '当年データの取得
                Dim dtThis As DataTable = DAOSyukeiKekkahyo.GetTableMain(_db, _chosaKubun, pPkeyThis)
                If dtThis.Rows.Count > 0 Then
                    syukeiNoThis = dtThis.Rows(0).Item("集計番号").ToString
                    syukeiNameThis = dtThis.Rows(0).Item("集計名称").ToString
                    'REV_001↓
                    heikinShuruiThis = dtThis.Rows(0).Item("平均種類").ToString
                    'REV_001↑
                End If
                '前年データの取得
                Dim dtLast As DataTable = DAOSyukeiKekkahyo.GetTableMain(_db, _chosaKubun, pPkeyLast)
                If dtLast.Rows.Count > 0 Then
                    syukeiNoLast = dtLast.Rows(0).Item("集計番号").ToString
                    syukeiNameLast = dtLast.Rows(0).Item("集計名称").ToString
                End If
            End If
            For Each info As ItemInfo In (From n In itemInfoListRonri Where n.Formula.Contains("【"))
                '『【集計番号】』を値に置換
                info.FormulaConvValue = info.FormulaConvValue.Replace("【集計番号】", If(String.IsNullOrEmpty(syukeiNoThis), "''", "'" & syukeiNoThis & "'"))
                '『【集計名称】』を値に置換
                info.FormulaConvValue = info.FormulaConvValue.Replace("【集計名称】", If(String.IsNullOrEmpty(syukeiNameThis), "''", "'" & syukeiNameThis & "'"))
                '『【前集計番号】』を値に置換
                info.FormulaConvValue = info.FormulaConvValue.Replace("【前集計番号】", If(String.IsNullOrEmpty(syukeiNoLast), "''", "'" & syukeiNoLast & "'"))
                '『【前集計名称】』を値に置換
                info.FormulaConvValue = info.FormulaConvValue.Replace("【前集計名称】", If(String.IsNullOrEmpty(syukeiNameLast), "''", "'" & syukeiNameLast & "'"))
                'REV_001↓
                '『【平均種類】』を値に置換
                info.FormulaConvValue = info.FormulaConvValue.Replace("【平均種類】", If(String.IsNullOrEmpty(heikinShuruiThis), "''", "'" & heikinShuruiThis & "'"))
                'REV_001↑
            Next

            '-----------------------------------------------------------------------------------------------------
            '当年と前年のデータ取得
            '-----------------------------------------------------------------------------------------------------
            If Not _isErrCheck Then
                If pKkey Is Nothing Then
                    '集計結果表データを取得し、項目情報リストに値をセットする
                    Me.GetSyukeikekkahyo(_db, _chosaKubun, pPkeyThis, _ItemInfoList)
                    '集計結果表データを取得し、項目情報リストに値をセットする(前年データ)
                    Me.GetSyukeikekkahyo(_db, _chosaKubun, pPkeyLast, _ItemInfoList, True)
                Else
                    '集計結果表データを取得し、項目情報リストに値をセットする
                    Me.GetSyukeikekkahyoAllHeikin(_db, _chosaKubun, pPkeyThis, pKkey, _ItemInfoList)
                    '集計結果表データを取得し、項目情報リストに値をセットする(前年データ)
                    Me.GetSyukeikekkahyoAllHeikin(_db, _chosaKubun, pPkeyLast, pKkey, _ItemInfoList, True)
                End If
            End If

            '作成論理を「優先順位(昇順)」で並び替える
            Dim grpItemInfoList As List(Of ItemInfo) = (From n In itemInfoListRonri Order By n.Priority).ToList
            For Each grpInfo In grpItemInfoList
                Try
                    '計算式文字列の項目番号を値に置換する
                    If Me.ConvertFormulaItemNoToValue(_ItemInfoList, grpInfo) Then
                        '式中全ての項目がNullだった場合は式をNullとする。
                        grpInfo.FormulaConvValue = "NULL"
                    End If

                    '作成論理をSQLで実行し、値を取得する
                    Dim value As Object = Me.ExecuteSQL(_db, grpInfo)
                    '型区分が数値の場合
                    If grpInfo.ValueType = ComConst.型区分.数値型 Then
                        If IsDBNull(value) OrElse value Is Nothing Then
                            grpInfo.Value = Nothing
                        Else
                            Dim decValue As Decimal = CDec(value)
                            If grpInfo.HyojiTani Is Nothing Then
                                grpInfo.Value = decValue
                            Else
                                '表示単位で四捨五入する
                                Dim roundValue As Decimal = Me.Round(decValue, grpInfo.HyojiTani)
                                grpInfo.Value = roundValue
                            End If
                        End If
                    Else
                        grpInfo.Value = If(IsDBNull(value) OrElse String.IsNullOrEmpty(CStr(value)), Nothing, value)
                    End If
                Catch ex As Exception
                    '集計結果検討表作成エラー
                    Throw New CreateSyukeiKentohyoException(ex, pPkeyThis.syukeiNo, pPkeyThis.groupKey, grpInfo.ItemNo, grpInfo.Formula, grpInfo.FormulaConvValue)
                End Try
            Next

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

            '項目マスタがデータ無しの場合
            If _dtSyukeiItemMst Is Nothing OrElse _dtSyukeiItemMst.Rows.Count = 0 OrElse
                _dtSyukeiKentoItemMst Is Nothing OrElse _dtSyukeiKentoItemMst.Rows.Count = 0 Then
                ret = False
                Exit Try
            End If

            '『[』『]』で囲まれた文字列検索する
            Dim matchList As MatchCollection = Regex.Matches(pFormula, C_SearchKakko)
            For Each mat As Match In matchList
                Dim itemNo As String = mat.Value.Trim("["c, "]"c).Replace("前", "")
                Dim isVariable As Boolean = False

                Select Case (itemNo.Substring(0, 1))
                    Case "S"
                        dt = _dtSyukeiItemMst
                    Case "D"
                        dt = _dtSyukeiKentoItemMst
                    Case Else
                        ret = False
                        Exit For
                End Select

                '調査区分が「畜産物生産費」で、かつ集計結果表項目の場合
                If ComUtil.IsChikusan() AndAlso itemNo.Substring(0, 1) = "S" Then
                    Dim drItemNo As DataRow = (From n In dt.AsEnumerable Where n("項目番号").ToString = itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(0)).FirstOrDefault
                    If Not drItemNo Is Nothing Then
                        Dim ItemNoSeisanhiHeikinChi As New List(Of String)
                        For i As Integer = 1 To If(_chosaKubun.Equals(ComConst.調査区分.牛乳生産費統計_個別) Or _chosaKubun.Equals(ComConst.調査区分.経営分析調査_牛乳生産費), 5, 4)
                            ItemNoSeisanhiHeikinChi.Add(drItemNo("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & i)
                        Next

                        '項目番号の後ろの「生産費平均値種類」が範囲外の場合
                        If Not ItemNoSeisanhiHeikinChi.Contains(itemNo) Then
                            ret = False
                            Exit For
                        End If
                    Else
                        '項目番号が項目マスタに存在しない場合
                        ret = False
                        Exit For
                    End If
                    ' REV_003↓
                    '集計結果検討表（報告用）、かつ調査区分が「農産物生産費」、かつ集計結果表項目の場合
                ElseIf _logicType.Equals(ComConst.集計結果検討表作成論理.論理種別.集計結果検討表_報告用) AndAlso ComUtil.IsNousan() AndAlso itemNo.Substring(0, 1) = "S" Then
                    Dim drItemNo As DataRow = (From n In dt.AsEnumerable Where n("項目番号").ToString = itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(0)).FirstOrDefault
                    If Not drItemNo Is Nothing Then
                        Dim ItemNoSeisanhiHeikinChi As New List(Of String)
                        For i As Integer = 1 To If(_chosaKubun.Equals(ComConst.調査区分.牛乳生産費統計_個別) Or _chosaKubun.Equals(ComConst.調査区分.経営分析調査_牛乳生産費), 5, 4)
                            ItemNoSeisanhiHeikinChi.Add(drItemNo("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & i)
                        Next

                        '項目番号の後ろの「生産費平均値種類」が範囲外の場合
                        If Not ItemNoSeisanhiHeikinChi.Contains(itemNo) Then
                            ret = False
                            Exit For
                        End If
                    Else
                        '項目番号が項目マスタに存在しない場合
                        ret = False
                        Exit For
                    End If
                    ' REV_003↑
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
                Execute(Nothing, Nothing)
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
            ret = ret.Replace("＋", "+").Replace("－", "- ").Replace("÷", "/").Replace("×", "*")

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計結果表データを取得し、項目情報リストに値をセットする
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pChosaKubun"></param>
    ''' <param name="pPKey"></param>
    ''' <param name="pItemInfoList"></param>
    ''' <param name="isZennen">前年データかどうか</param>
    ''' <remarks></remarks>
    Private Sub GetSyukeikekkahyo(ByVal pDb As DBAccess,
                                  ByVal pChosaKubun As String,
                                  ByVal pPKey As DAOSyukeiKekkahyo.PrimaryKey,
                                  ByRef pItemInfoList As List(Of ItemInfo),
                                  Optional ByVal isZennen As Boolean = False)
        Dim dicSyukei As Dictionary(Of String, DataTable) = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim dc As DataColumn = Nothing

        Try

            '集計結果表テーブルリスト取得
            dicSyukei = DAOSyukeiKekkahyo.GetTable(pDb, pChosaKubun, pPKey)
            For Each kv As KeyValuePair(Of String, DataTable) In dicSyukei
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
    ''' 集計結果表データを取得し、項目情報リストに値をセットする(生産費平均値種類１～５)
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pChosaKubun"></param>
    ''' <param name="pPKey"></param>
    ''' <param name="pItemInfoList"></param>
    ''' <param name="isZennen">前年データかどうか</param>
    ''' <remarks></remarks>
    Private Sub GetSyukeikekkahyoAllHeikin(ByVal pDb As DBAccess,
                                  ByVal pChosaKubun As String,
                                  ByVal pPKey As DAOSyukeiKekkahyo.PrimaryKey,
                                  ByVal pKKey As DAOSyukeiKekkahyo.KomokuKey,
                                  ByRef pItemInfoList As List(Of ItemInfo),
                                  Optional ByVal isZennen As Boolean = False)

        Dim dicSyukei As Dictionary(Of String, DataTable) = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim dc As DataColumn = Nothing

        Try

            '集計結果表テーブルリスト取得
            dicSyukei = DAOSyukeiKekkahyo.GetTableAllHeikin(pDb, pChosaKubun, pPKey, pKKey)
            For Each kv As KeyValuePair(Of String, DataTable) In dicSyukei
                dt = kv.Value
                If dt.Rows.Count > 0 Then

                    Dim seisanhiHeikinchiSyurui As String = Nothing
                    For Each dr In dt.Rows
                        seisanhiHeikinchiSyurui = dr("生産費平均値種類").ToString
                        For Each dc In dt.Columns
                            '項目情報リストの項目番号が一致するものに値をセット
                            SetItemInfoListValue(pItemInfoList, If(isZennen, "前" & dc.ColumnName & ComConst.ITEM_NO_DELIMITER & seisanhiHeikinchiSyurui, dc.ColumnName & ComConst.ITEM_NO_DELIMITER & seisanhiHeikinchiSyurui), dr(dc.ColumnName))
                        Next
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
    Private Function ConvertFormulaItemNoToValue(ByVal pItemInfoList As List(Of ItemInfo), ByRef pItemInfo As ItemInfo) As Boolean
        Dim matchList As MatchCollection = Nothing

        '括弧付き項目番号『[010101]』『[010101_1]』を検索
        matchList = Regex.Matches(pItemInfo.FormulaConvValue, C_SearchKakko)

        '式中の全ての項番がNull(皆無)かどうかを判定するフラグ
        Dim isAllItemNoNull As Boolean = If(matchList.Count > 0, True, False)
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

                '式中に一つでもデータが入っていればFalse
                If (Not (IsDBNull(value) OrElse value Is Nothing)) Then
                    isAllItemNoNull = False
                End If

                '@付き項目番号『[010101]@』『[010101_1]@』を値に置換
                If info.ValueType = ComConst.型区分.数値型 Then
                    pItemInfo.FormulaConvValue = pItemInfo.FormulaConvValue.Replace("[" & itemNo & "]@", If(IsDBNull(value) OrElse value Is Nothing, "NULL", Me.CnvIntToDec(value.ToString)))
                Else
                    pItemInfo.FormulaConvValue = pItemInfo.FormulaConvValue.Replace("[" & itemNo & "]@", If(IsDBNull(value) OrElse value Is Nothing, "NULL", "'" & value.ToString & "'"))
                End If
                '項目番号『[010101]』『[010101_1]』を値に置換
                If info.ValueType = ComConst.型区分.数値型 Then
                    pItemInfo.FormulaConvValue = pItemInfo.FormulaConvValue.Replace("[" & itemNo & "]", If(IsDBNull(value) OrElse value Is Nothing, "0", Me.CnvIntToDec(value.ToString)))
                Else
                    pItemInfo.FormulaConvValue = pItemInfo.FormulaConvValue.Replace("[" & itemNo & "]", If(IsDBNull(value) OrElse value Is Nothing, "''", "'" & value.ToString & "'"))
                End If
            End If
        Next

        Return isAllItemNoNull
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
End Class
