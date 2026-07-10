Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Runtime.Serialization.Formatters.Binary

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2023.01.06 |Daiko               | 要件No4 バージョン区分追加
'//  REV_002   | 2023.01.12 |Daiko               | 要件No8 バージョン区分追加
'//  REV_003   | 2025.08.28 |GCU                 | 要件No2 継続区分追加
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 集計結果表作成エラークラス
''' </summary>
''' <remarks></remarks>
Public Class CreateSyukeiException
    Inherits Exception

    ''' <summary>集計番号</summary>
    Private _syukeiNo As String

    ''' <summary>センサス番号</summary>
    Public ReadOnly Property SyukeiNo As String
        Get
            Return _syukeiNo
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

    Public Sub New(inner As Exception, ByVal pSyukeiNo As String, ByVal pItemNo As String, ByVal pFormula As String, ByVal pFormulaCnv As String)
        MyBase.New(String.Format("集計結果表作成エラー(項目番号:{0} 計算式:{1} 計算式(変換後):{2})", pItemNo, pFormula, pFormulaCnv), inner)
        _syukeiNo = pSyukeiNo
        _itemNo = pItemNo
    End Sub
End Class

''' <summary>
''' 集計結果表・集計結果検討表作成クラス
''' </summary>
''' <remarks></remarks>
Public Class CreateSyukei

    ''' <summary>
    ''' 項目種別
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum enmItemType
        個別結果表
        集計結果表
    End Enum

    ''' <summary>
    ''' 項目情報クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ItemInfo
        ''' <summary>調査区分</summary>
        Public Chosakubun As String
        ''' <summary>項目番号</summary>
        Public ItemNo As String
        ''' <summary>項目名</summary>
        Public ItemName As String
        ''' <summary>優先順位</summary>
        Public Priority As Decimal
        ''' <summary>計算式</summary>
        Public Formula As String
        ''' <summary>計算式(記号変換)</summary>
        Public FormulaConv As String
        ''' <summary>計算式(値変換)</summary>
        Public FormulaConvValue As String
        ''' <summary>表示単位</summary>
        Public HyojiTani As Decimal
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
        ''' <summary>集計倍率区分</summary>
        Public SyukeibairitsuKbn As String
        ''' <summary>判定対象外区分</summary>
        Public HanteiTaisyougaiKbn As String
        ''' <summary>項目種別</summary>
        Public ItemType As enmItemType
        ''' <summary>値</summary>
        Public Value As List(Of Object)
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

    ''' <summary>正規表現検索文字列</summary>
    Public Const C_SearchKakko As String = "\[前?[QKCS]+[0-9_]+\]"
    Protected Const C_SearchRank As String = "RANK\([^\(\)]*(\([^\(\)]*(\([^\(\)]*[^\(\)]*\)[^\(\)]*)*[^\(\)]*\)[^\(\)]*)*[^\(\)]*\)"
    Protected Const C_SearchRankKakko As String = "\{[^\(\)]*[^\(\)]*\}"
    Protected Const C_SearchStackAverageFormula As String = "Σ\([^\(\)]*(\([^\(\)]*(\([^\(\)]*[^\(\)]*\)[^\(\)]*)*[^\(\)]*\)[^\(\)]*)*[^\(\)]*\)"
    Protected Const C_SearchStackFormula As String = "σ\([^\(\)]*(\([^\(\)]*(\([^\(\)]*[^\(\)]*\)[^\(\)]*)*[^\(\)]*\)[^\(\)]*)*[^\(\)]*\)"

    ''' <summary>DBAccessオブジェクト</summary>
    Protected _db As DBAccess

    ''' <summary>調査区分</summary>
    Protected _chosaKubun As String

    ''' <summary>調査年</summary>
    Protected _chosaYear As String

    ''' <summary>項目情報リスト</summary>
    Protected _ItemInfoList As Dictionary(Of String, ItemInfo)

    ''' <summary>エラーチェックかどうか</summary>
    Protected _isErrCheck As Boolean

    ''' <summary>集計条件</summary>
    Protected _syukeiInfo As DAOKobetsuKekkahyo.SyukeiInfo

    ''' <summary>営農農業経営体集計の場合のみ使用：個人・法人それぞれのレコード数</summary>
    Protected _einouSplitRecordCount As Dictionary(Of String, Integer)

    ''' <summary>集計倍率区分：営農経営体集計倍率</summary>
    Protected Const KBN_KEIEITAIBAIRITSU = "1"
    ''' <summary>集計倍率区分：部門集計倍率</summary>
    Protected Const KBN_BUMONBAIRITSU = "2"
    ''' <summary>１回で積み上げる最大のレコード数</summary>
    Protected Const STACK_MAX = 2000
    ''' <summary>集計倍率の積み上げ値</summary>
    Protected _syukeiBairitsu As String
    ''' <summary>部門集計倍率の積み上げ値</summary>
    Protected _bumonSyukeiBairitsu As String
    'REV_003↓
    ''' <summary>継続区分</summary>
    Protected _keizokuKubun As String

    ''' <summary>前回センサス番号使用</summary>
    Protected _zenkaiCensus As Boolean
    'REV_003↑

    Private _saisen1KiboNValues As New Dictionary(Of String, Dictionary(Of String, Boolean))

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pChosaKubun"></param>
    ''' <param name="pChosaYear"></param>
    ''' <param name="pCreateType"></param>
    ''' <param name="pDtKobetsuItemMst"></param>
    ''' <param name="pDtSyukeiItemMst"></param>
    ''' <param name="pDtCreateRonri"></param>
    ''' <param name="pSyukeiInfo"></param>
    ''' <param name="isErrCheck"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pDb As DBAccess,
               ByVal pChosaKubun As String,
               ByVal pChosaYear As String,
               ByVal pDtKobetsuItemMst As DataTable,
               ByVal pDtSyukeiItemMst As DataTable,
               ByVal pDtCreateRonri As DataTable,
               Optional isErrCheck As Boolean = False,
               Optional pKeizokuKubun As String = Nothing,　'REV_003←
               Optional pZenkaiCensus As Boolean = False)　'REV_003←

        Try
            _isErrCheck = isErrCheck
            _ItemInfoList = New Dictionary(Of String, ItemInfo)

            'DBAccessオブジェクト
            _db = pDb
            '調査区分
            _chosaKubun = pChosaKubun
            '調査年
            _chosaYear = pChosaYear
            'REV_003↓
            '継続区分
            _keizokuKubun = pKeizokuKubun
            '前回センサス番号使用
            _zenkaiCensus = pZenkaiCensus
            'REV_003↑
            '個別結果表項目マスタ
            If pDtKobetsuItemMst IsNot Nothing Then
                For Each dr As DataRow In pDtKobetsuItemMst.Rows
                    _ItemInfoList.Add(dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & dr("調査区分").ToString, New ItemInfo With {.Chosakubun = dr("調査区分").ToString,
                                                         .ItemNo = dr("項目番号").ToString,
                                                         .SheetName = dr("シート名").ToString,
                                                         .Row = CInt(dr("行位置")),
                                                         .Col = CInt(dr("列位置")),
                                                         .ValueType = dr("型区分").ToString,
                                                         .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                         .ItemType = enmItemType.個別結果表,
                                                         .Value = New List(Of Object)})
                Next
            End If

            '集計結果表項目マスタ
            If pDtSyukeiItemMst IsNot Nothing Then
                For Each dr As DataRow In pDtSyukeiItemMst.Rows
                    _ItemInfoList.Add(dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & dr("調査区分").ToString, New ItemInfo With {.Chosakubun = dr("調査区分").ToString,
                                                         .ItemNo = dr("項目番号").ToString,
                                                         .SheetName = dr("シート名").ToString,
                                                         .Row = CInt(dr("行位置")),
                                                         .Col = CInt(dr("列位置")),
                                                         .ValueType = dr("型区分").ToString,
                                                         .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                         .SyukeibairitsuKbn = dr("集計倍率区分").ToString,
                                                         .HanteiTaisyougaiKbn = dr("判定対象外区分").ToString,
                                                         .ItemType = enmItemType.集計結果表,
                                                         .Value = New List(Of Object)})
                Next
                If Not ComUtil.IsEinou() Then
                    For Each dr As DataRow In pDtSyukeiItemMst.Rows
                        _ItemInfoList.Add(dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & "1" & ComConst.ITEM_NO_DELIMITER & dr("調査区分").ToString, New ItemInfo With {.Chosakubun = dr("調査区分").ToString,
                                                             .ItemNo = dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & "1",
                                                             .SheetName = dr("シート名").ToString,
                                                             .Row = CInt(dr("行位置")),
                                                             .Col = CInt(dr("列位置")),
                                                             .ValueType = dr("型区分").ToString,
                                                             .IsHidden = If(CInt(dr("裏項番区分")) = 1, True, False),
                                                             .SyukeibairitsuKbn = dr("集計倍率区分").ToString,
                                                             .HanteiTaisyougaiKbn = dr("判定対象外区分").ToString,
                                                             .ItemType = enmItemType.集計結果表,
                                                             .Value = New List(Of Object)})
                    Next
                End If
            End If

            Me.addCreateRonriInfo(pDtCreateRonri)

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 集計を実行する
    ''' </summary>
    ''' <param name="pChiiki"></param>
    ''' <param name="pKibokaisou"></param>
    ''' <param name="pEinouChusyutsuJouken"></param>
    ''' <param name="pSeisanhiChusyutsuJouken"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Execute(pProgressDialog As ProgressDialog,
                            pChiiki As String,
                            pKibokaisou As DAOKobetsuKekkahyo.Kibokaisou,
                            pEinouChusyutsuJouken As DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu,
                            pSeisanhiChusyutsuJouken As DAOKobetsuKekkahyo.SeisanhiChusyutsu,
                            pSyukeiInfo As DAOKobetsuKekkahyo.SyukeiInfo,
                            Optional pSousuRecord As Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目) = Nothing) As Dictionary(Of String, ItemInfo)

        Dim itemInfoListRonri As List(Of ItemInfo)

        _syukeiInfo = pSyukeiInfo
        _einouSplitRecordCount = New Dictionary(Of String, Integer)

        Try
            Me.initializeItemValue()
            '作成論理が存在するリストを作成
            itemInfoListRonri = (From n In _ItemInfoList.Values Where Not String.IsNullOrEmpty(n.Formula)).ToList

            '-----------------------------------------------------------------------------------------------------
            'データ取得
            '-----------------------------------------------------------------------------------------------------
            Dim kobetsuRecordCount As Integer = 0
            Dim weight As String = String.Empty


            '営農の場合
            If ComUtil.IsEinou() Then
                '個別結果表データを項目情報リストにセットし、個別結果表のレコード数を取得する。
                kobetsuRecordCount = Me.GetKobetsu(_db,
                                                    New DAOKobetsuKekkahyo.PrimaryKey(_chosaYear),
                                                    If(CommonInfo.Koutei.Equals("H"), Nothing, New DAOKobetsuKekkahyo.KyotenKey(CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)),
                                                    _syukeiInfo,
                                                    _ItemInfoList,
                                                    pKibokaisou,
                                                    pEinouChusyutsuJouken,
                                                    pSeisanhiChusyutsuJouken)
            Else
                '生産費で総数集計の場合
                If _syukeiInfo.seisanhiHeikin.Equals("1") Then
                    '個別結果表データを項目情報リストにセットし、個別結果表のレコード数を取得する。
                    kobetsuRecordCount = Me.GetKobetsu(_db,
                                                        New DAOKobetsuKekkahyo.PrimaryKey(_chosaYear),
                                                        If(CommonInfo.Koutei.Equals("H"), Nothing, New DAOKobetsuKekkahyo.KyotenKey(CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)),
                                                        _syukeiInfo,
                                                        _ItemInfoList,
                                                        pKibokaisou,
                                                        pEinouChusyutsuJouken,
                                                        pSeisanhiChusyutsuJouken)
                ElseIf Not pSousuRecord Is Nothing Then
                    '集計結果表（生産費平均値種類１）のデータを項目情報リストにセットする。
                    For Each kv As KeyValuePair(Of String, DAOSyukeiKekkahyo.集計結果表項目) In pSousuRecord
                        SetItemInfoListValues(_ItemInfoList, kv.Key & ComConst.ITEM_NO_DELIMITER & "1", kv.Value.値, _chosaKubun)
                    Next

                    '生産費平均値種類Weightを取得
                    weight = getWeight(_syukeiInfo.seisanhiHeikin)
                    Dim info As ItemInfo = New ItemInfo
                    info.FormulaConvValue = weight
                    Me.ConvertFormulaItemNoToValue(_ItemInfoList, info, False, 0, _chosaKubun)
                    weight = info.FormulaConvValue
                End If
            End If

            '集計倍率の積み上げを取得する
            _syukeiBairitsu = getSyukeiBairitsu()
            _bumonSyukeiBairitsu = If(_chosaKubun.Equals(ComConst.調査区分.営農類型別経営統計_個人), getBumonSyukeiBairitsu(), Nothing)

            '作成論理を「優先順位(昇順)」で並び替える
            Dim grpItemInfoList As List(Of ItemInfo) = (From n In itemInfoListRonri Order By n.Priority).ToList

            If kobetsuRecordCount > 0 OrElse Not pSousuRecord Is Nothing Then
                '作成論理 １行毎
                For Each grpInfo In grpItemInfoList
                    Try
                        Dim value As Object
                        'Σ( ) で囲われた計算を積み上げと判断する
                        If grpInfo.Formula.Contains("Σ") Then
                            '積み上げ実行
                            value = Me.calcStackAndAvarage(grpInfo, pChiiki, kobetsuRecordCount, pEinouChusyutsuJouken, pSeisanhiChusyutsuJouken)
                            If grpInfo.ItemNo = "S020216" Then
                                If CommonInfo.Chosakubun.Equals(ComConst.調査区分.米生産費統計_個別) OrElse CommonInfo.Chosakubun.Equals(ComConst.調査区分.米生産費統計_組織法人) Then
                                    If Not _saisen1KiboNValues.ContainsKey(_syukeiInfo.kiboKaisou) Then
                                        _saisen1KiboNValues.Add(_syukeiInfo.kiboKaisou, New Dictionary(Of String, Boolean))
                                    End If
                                    _saisen1KiboNValues(_syukeiInfo.kiboKaisou)(pChiiki) = IsDBNull(value) OrElse value Is Nothing OrElse CDec(value) = 0
                                End If
                            End If
                        ElseIf grpInfo.Formula.Contains("σ") Then
                            'σ( ) で囲われた計算は単純積み上げ（加重累積）と判断する
                            value = Me.calcStack(grpInfo, pChiiki, kobetsuRecordCount, pEinouChusyutsuJouken, pSeisanhiChusyutsuJouken)
                            If grpInfo.ItemNo = "S020216" Then
                                If CommonInfo.Chosakubun.Equals(ComConst.調査区分.米生産費統計_個別) OrElse CommonInfo.Chosakubun.Equals(ComConst.調査区分.米生産費統計_組織法人) Then
                                    If Not _saisen1KiboNValues.ContainsKey(_syukeiInfo.kiboKaisou) Then
                                        _saisen1KiboNValues.Add(_syukeiInfo.kiboKaisou, New Dictionary(Of String, Boolean))
                                    End If
                                    _saisen1KiboNValues(_syukeiInfo.kiboKaisou)(pChiiki) = IsDBNull(value) OrElse value Is Nothing OrElse CDec(value) = 0
                                End If
                            End If
                        Else
                            '生産費平均値種類によってＷの値を置換する
                            grpInfo.FormulaConvValue = grpInfo.FormulaConv.Replace("Ｗ", weight)

                            '計算式文字列の項目番号を値に置換する
                            Dim isFormulaNull As Boolean = Me.ConvertFormulaItemNoToValue(_ItemInfoList, grpInfo, False, 0, _chosaKubun)

                            '計算式文字列の順位関連文字列を値に置換する
                            Me.ConvertFormulaRankToValue(grpInfo)

                            '固定文字で示された論理を置換する
                            grpInfo.FormulaConvValue = replaceRonri(grpInfo.FormulaConvValue, pChiiki, kobetsuRecordCount, pEinouChusyutsuJouken, pSeisanhiChusyutsuJouken)

                            '指標部以外で全ての項目で置換ができなかった（Nullまたは０）場合は式全体をNullとする
                            If isFormulaNull And (Not ComConst.集計結果表.指標部.Item(_chosaKubun).Contains(grpInfo.ItemNo)) Then
                                grpInfo.FormulaConvValue = Nothing
                            End If
                            '作成論理をSQLで実行し、値を取得する
                            value = Me.ExecuteSQL(_db, grpInfo)
                        End If

                        '計算結果を項目リストにセット
                        '型区分が数値の場合
                        If grpInfo.ValueType = ComConst.型区分.数値型 Then
                            If IsDBNull(value) OrElse value Is Nothing Then
                                AddValueInItemInfo(grpInfo, value)
                            Else
                                Dim decValue As Decimal = CDec(value)
                                '表示単位がある場合 表示単位で四捨五入する
                                If Not grpInfo.HyojiTani = 0 Then
                                    decValue = Me.Round(decValue, grpInfo.HyojiTani)
                                End If

                                If grpInfo.ItemNo = "S020216" Then
                                    If CommonInfo.Chosakubun.Equals(ComConst.調査区分.米生産費統計_個別) OrElse CommonInfo.Chosakubun.Equals(ComConst.調査区分.米生産費統計_組織法人) Then
                                        Dim isNullOrZero As Boolean = False
                                        If _saisen1KiboNValues.ContainsKey(_syukeiInfo.kiboKaisou) AndAlso _saisen1KiboNValues(_syukeiInfo.kiboKaisou).ContainsKey(pChiiki) Then
                                            isNullOrZero = _saisen1KiboNValues(_syukeiInfo.kiboKaisou)(pChiiki)
                                        End If
                                        
                                        If isNullOrZero Then
                                            AddValueInItemInfo(grpInfo, Nothing)
                                        Else
                                            AddValueInItemInfo(grpInfo, decValue)
                                        End If
                                    Else
                                        AddValueInItemInfo(grpInfo, decValue)
                                    End If
                                Else
                                    AddValueInItemInfo(grpInfo, decValue)
                                End If
                            End If
                        Else
                            AddValueInItemInfo(grpInfo, If(IsDBNull(value) OrElse String.IsNullOrEmpty(CStr(value)), Nothing, value))
                        End If

                        '進捗を進める
                        If pProgressDialog IsNot Nothing Then
                            pProgressDialog.AddValue = 1
                        End If

                    Catch ex As Exception
                        '個別結果表作成エラー
                        Throw New CreateSyukeiException(ex, _syukeiInfo.syukeiNo, grpInfo.ItemNo, grpInfo.Formula, grpInfo.FormulaConvValue)
                    End Try
                Next

                Return _ItemInfoList
            Else
                '進捗を進める
                If pProgressDialog IsNot Nothing Then
                    pProgressDialog.AddValue = grpItemInfoList.Count
                End If

                Return Nothing
            End If
        Catch ex As Exception
            Throw
        End Try

    End Function


    ''' <summary>
    ''' 【生産費】出力用のレコードを作成する（生産費平均値種類０）
    ''' </summary>
    ''' <param name="pChiiki"></param>
    ''' <param name="pSeisanhiChusyutsuJouken"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteMerge(pProgressDialog As ProgressDialog, tplSyukeiList As List(Of ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)))) As Dictionary(Of String, ItemInfo)

        '平均値種類１～５のレコードを分割して管理する
        Dim syukeiRecord_sousu As New Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)
        Dim syukeiRecord_heikinchi2 As New Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)
        Dim syukeiRecord_heikinchi3 As New Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)
        Dim syukeiRecord_heikinchi4 As New Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)
        Dim syukeiRecord_heikinchi5 As New Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)

        '出力用編集マスタを取得する
        ' REV_002↓
        'Dim dtSyutsuryokuRonri As DataTable = DAOOther.GetSeisanhiSyutsuryokuHensyu(_db, _chosaKubun)
        Dim dtSyutsuryokuRonri As DataTable = DAOOther.GetSeisanhiSyutsuryokuHensyu(_db, _chosaKubun, ComUtil.getVersionKubunTaikei(_chosaYear, _chosaKubun))
        ' REV_002↑

        If tplSyukeiList.Any Then

            For Each tplSyukei As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)) In tplSyukeiList

                Select Case tplSyukei.Item2.seisanhiHeikin
                    Case "1"
                        syukeiRecord_sousu = tplSyukei.Item3
                    Case "2"
                        syukeiRecord_heikinchi2 = tplSyukei.Item3
                    Case "3"
                        syukeiRecord_heikinchi3 = tplSyukei.Item3
                    Case "4"
                        syukeiRecord_heikinchi4 = tplSyukei.Item3
                    Case "5"
                        syukeiRecord_heikinchi5 = tplSyukei.Item3
                End Select
            Next

            '指標部は生産費平均値種類１～５と同じ値が入る（S000006「平均種類」のみ「0」が入る）
            For Each shihyoNo As String In ComConst.集計結果表.指標部.Item(_chosaKubun)
                Dim info As ItemInfo = _ItemInfoList(shihyoNo & ComConst.ITEM_NO_DELIMITER & _chosaKubun)

                Dim value As String = String.Empty
                If shihyoNo.Equals(ComConst.集計結果表.生産費平均値種類) Then
                    value = "0"
                Else
                    value = syukeiRecord_sousu.Item(shihyoNo).値
                End If
                AddValueInItemInfo(info, value)
            Next

            '出力用編集マスタ１レコード毎に計算を行う
            For Each dr As DataRow In dtSyutsuryokuRonri.Rows
                Dim itemNo As String = dr("項目番号").ToString
                Dim seisanhiHeikin As String = dr("生産費平均値種類").ToString
                Dim editItemNo As String = dr("編集元項番").ToString

                Dim value As Object = Nothing
                Dim info As ItemInfo
                Select Case seisanhiHeikin
                    Case "2"
                        value = syukeiRecord_heikinchi2.Item(editItemNo).値
                    Case "3"
                        value = syukeiRecord_heikinchi3.Item(editItemNo).値
                    Case "4"
                        value = syukeiRecord_heikinchi4.Item(editItemNo).値
                    Case "5"
                        value = syukeiRecord_heikinchi5.Item(editItemNo).値
                End Select

                info = _ItemInfoList(itemNo & ComConst.ITEM_NO_DELIMITER & _chosaKubun)

                AddValueInItemInfo(info, If(IsDBNull(value) OrElse String.IsNullOrEmpty(CStr(value)), Nothing, value))

                '進捗を進める
                If pProgressDialog IsNot Nothing Then
                    pProgressDialog.AddValue = 1
                End If

            Next

            Return _ItemInfoList
        Else
            '進捗を進める
            If pProgressDialog IsNot Nothing Then
                pProgressDialog.AddValue = dtSyutsuryokuRonri.Rows.Count
            End If
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' 計算式の文字列置換する(項目番号以外)
    ''' </summary>
    ''' <param name="pFormula"></param>
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
            ret = ret.Replace("＋", "+").Replace("－", "- ").Replace("-", "- ").Replace("÷", "/").Replace("×", "*")
            '全角括弧『（, ）』 ⇒ 半角括弧『(, )』
            ret = ret.Replace("（", "(").Replace("）", ")")
        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function


    ''' <summary>
    ''' 計算式文字列の項目番号を値に置換する
    ''' </summary>
    ''' <param name="pItemInfoList"></param>
    ''' <param name="pItemInfo"></param>
    ''' <param name="idx"></param>
    ''' <param name="pChosakubun"></param>
    ''' <remarks></remarks>
    ''' <returns>式全体がNullかどうか 全体がNullの場合「True」 ひとつでも値がある場合「False」</returns>
    Private Function ConvertFormulaItemNoToValue(ByVal pItemInfoList As Dictionary(Of String, ItemInfo), ByRef pItemInfo As ItemInfo, ByVal isStack As Boolean, ByVal idx As Integer, ByVal pChosakubun As String) As Boolean
        Dim matchList As MatchCollection = Nothing

        '括弧付き項目番号『[010101]』『[010101_1]』を検索
        matchList = Regex.Matches(pItemInfo.FormulaConvValue, C_SearchKakko)
        Dim isAllItemNoNull As Boolean = If(matchList.Count > 0, True, False)

        '重複しない括弧付き項目番号リストを作成
        Dim itemNoKakkoList As String() = (From m In matchList.Cast(Of Match)() Select m.Value).Distinct().ToArray()
        For Each itemNoKakko As String In itemNoKakkoList
            '『[010101]』『[010101_1]』から項目番号『010101』『010101_1』を抽出する
            Dim itemNo As String = itemNoKakko.Trim("["c, "]"c)
            Dim infoChosakubun As String = If(pChosakubun Is Nothing, _chosaKubun, pChosakubun)

            Dim info As ItemInfo = Nothing

            Dim serchItemNo As String = itemNo & ComConst.ITEM_NO_DELIMITER & pChosakubun
            If pItemInfoList.ContainsKey(serchItemNo) Then
                info = pItemInfoList(serchItemNo)
            End If

            'データが存在する場合
            If info IsNot Nothing Then

                'エラーチェック時は文字列型には任意の文字をセットする(文字列と数値の比較『''<0』等がエラーにならないため)
                Dim value As Object
                If info.ValueType = ComConst.型区分.文字型 AndAlso _isErrCheck Then
                    value = "a"
                ElseIf info.Value.Count > 0 Then
                    value = info.Value(idx)
                Else
                    value = Nothing
                End If

                '値が入っているかどうか（個別結果表の集計倍率と集計結果表の指標部はカウントしない）
                If Not isAllItemNoNull OrElse Not _isErrCheck Then
                    If isStack Then
                        If (Not (IsDBNull(value) OrElse value Is Nothing OrElse CDec(value) = 0)) AndAlso
                           (infoChosakubun.Equals(ComConst.営農経営体区分.農業経営体) OrElse Not itemNo.Equals(ComConst.個別結果表.集計倍率(infoChosakubun))) AndAlso
                           (Not ComConst.集計結果表.指標部.Item(infoChosakubun).Contains(itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(0))) Then
                            isAllItemNoNull = False
                        End If
                    Else
                        If (Not (IsDBNull(value) OrElse value Is Nothing)) AndAlso
                           (infoChosakubun.Equals(ComConst.営農経営体区分.農業経営体) OrElse Not itemNo.Equals(ComConst.個別結果表.集計倍率(infoChosakubun))) AndAlso
                           (Not ComConst.集計結果表.指標部.Item(infoChosakubun).Contains(itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(0))) Then

                            '表上計算で総数以外なら 項番が判定対象外区分(0かNullかを判別すべき項番か)を見る
                            If Not _syukeiInfo.seisanhiHeikin.Equals("1") Then
                                If (info.ItemType.Equals(enmItemType.集計結果表) AndAlso info.HanteiTaisyougaiKbn.Equals("0")) Then
                                    isAllItemNoNull = False
                                End If
                            Else
                                isAllItemNoNull = False
                            End If
                        End If
                    End If
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
    ''' 個別結果表のデータを取得し、項目情報リストに値をセットする
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pPKey"></param>
    ''' <param name="syukeiInfo"></param>
    ''' <param name="pItemInfoList"></param>
    ''' <returns>個別結果表のレコード数</returns>
    ''' <remarks></remarks>
    Private Function GetKobetsu(ByVal pDb As DBAccess,
                                ByVal pPKey As DAOKobetsuKekkahyo.PrimaryKey,
                                ByVal pkKey As DAOKobetsuKekkahyo.KyotenKey,
                                ByRef pSyukeiInfo As DAOKobetsuKekkahyo.SyukeiInfo,
                                ByRef pItemInfoList As Dictionary(Of String, ItemInfo),
                                ByRef pKibokaisou As DAOKobetsuKekkahyo.Kibokaisou,
                                ByRef pEinoukeieitaiChusyutsuJoken As DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu,
                                ByRef pSeisanhiChusyutsuJoken As DAOKobetsuKekkahyo.SeisanhiChusyutsu) As Integer

        Dim dicKobetsu As Dictionary(Of String, DataTable) = Nothing
        Dim dt As DataTable = Nothing
        Dim dc As DataColumn = Nothing
        Dim lst As List(Of DataRow) = Nothing
        Dim recordCount As Integer = 0

        Try
            '集計用センサス番号管理テーブル取得
            Dim syukeiCensusNoList As List(Of String) = Me.getSyukeiCensusNoList(pDb)

            'データリスト取得
            If _chosaKubun.Equals(ComConst.営農経営体区分.農業経営体) Then
                '農業経営体の場合の条件『【個人の抽出条件】@【法人の抽出条件】』と格納されている
                Dim joukenList As String()
                If pEinoukeieitaiChusyutsuJoken.jouken Is Nothing Then
                    joukenList = {Nothing, Nothing}
                Else
                    joukenList = pEinoukeieitaiChusyutsuJoken.jouken.Split("@"c)
                End If
                For i As Integer = 1 To joukenList.Count
                    Dim einouChosaKbn As String = CStr(i)
                    'REV_003↓
                    dicKobetsu = DAOKobetsuKekkahyo.GetKobetsuTable(pDb, einouChosaKbn, pPKey, pkKey, pSyukeiInfo, pKibokaisou, pEinoukeieitaiChusyutsuJoken, pSeisanhiChusyutsuJoken, syukeiCensusNoList, joukenList(i - 1), ComConst.欠測値補完.無, _keizokuKubun, _zenkaiCensus)
                    For Each kv As KeyValuePair(Of String, DataTable) In dicKobetsu
                        Dim tableName As String = kv.Key
                        dt = kv.Value
                        For Each dr As DataRow In dt.Rows
                            For Each dc In dt.Columns
                                '項目情報リストの項目番号が一致するものに値をセット
                                SetItemInfoListValues(pItemInfoList, dc.ColumnName, dr(dc.ColumnName), einouChosaKbn)
                            Next
                        Next
                    Next
                    recordCount += If(dicKobetsu.Count = 0, 0, dt.Rows.Count)
                    _einouSplitRecordCount.Add(einouChosaKbn, If(dicKobetsu.Count = 0, 0, dt.Rows.Count))
                Next
            Else
                '営農個別のみ欠測値保管後のテーブルを見る必要がある
                If _chosaKubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) Then
                    'REV_003↓
                    dicKobetsu = DAOKobetsuKekkahyo.GetKobetsuTable(pDb, _chosaKubun, pPKey, pkKey, pSyukeiInfo, pKibokaisou, pEinoukeieitaiChusyutsuJoken, pSeisanhiChusyutsuJoken, syukeiCensusNoList, , ComConst.欠測値補完.有, _keizokuKubun, _zenkaiCensus)
                Else
                    'REV_003↓
                    dicKobetsu = DAOKobetsuKekkahyo.GetKobetsuTable(pDb, _chosaKubun, pPKey, pkKey, pSyukeiInfo, pKibokaisou, pEinoukeieitaiChusyutsuJoken, pSeisanhiChusyutsuJoken, syukeiCensusNoList, , ComConst.欠測値補完.無, _keizokuKubun, _zenkaiCensus)
                End If
                For Each kv As KeyValuePair(Of String, DataTable) In dicKobetsu
                    Dim tableName As String = kv.Key
                    dt = kv.Value
                    lst = dt.AsEnumerable().ToList()
                    Dim colArr(dt.Columns.Count - 1) As DataColumn
                    dt.Columns.CopyTo(colArr, 0)
                    For Each dr As DataRow In lst
                        For Each dc In colArr
                            '項目情報リストの項目番号が一致するものに値をセット
                            SetItemInfoListValues(pItemInfoList, dc.ColumnName, dr(dc.ColumnName), _chosaKubun)
                        Next
                    Next
                Next
                recordCount = If(dicKobetsu.Count = 0, 0, lst.Count)
            End If
        Catch ex As Exception
            Throw ex
        End Try

        Return recordCount
    End Function

    ''' <summary>
    ''' 指定した項目番号と一致する項目情報リストのデータに値をセットする
    ''' </summary>
    ''' <param name="pItemInfoList"></param>
    ''' <param name="ItemNo"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Private Sub SetItemInfoListValues(ByRef pItemInfoList As Dictionary(Of String, ItemInfo), ByVal ItemNo As String, ByVal value As Object, chosakubun As String)
        Dim info As ItemInfo = Nothing

        Dim serchItemNo As String = ItemNo & ComConst.ITEM_NO_DELIMITER & chosakubun
        If pItemInfoList.ContainsKey(serchItemNo) Then
            info = pItemInfoList(serchItemNo)
        End If

        If info IsNot Nothing Then
            If serchItemNo.Contains("S") Then
                AddValueInItemInfo(info, value)
            Else
                info.Value.Add(value)
            End If
        End If

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
        matchItemNoList = Regex.Matches(pItemInfo.Formula, C_SearchRank, RegexOptions.IgnoreCase)
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
            sb.Append("SELECT " & If(String.IsNullOrEmpty(pItemInfo.FormulaConvValue), "NULL", pItemInfo.FormulaConvValue) & " AS [" & pItemInfo.ItemNo & "]")

            dt = pDb.GetDataTable(sb.ToString)
            ret = dt.Rows(0)(0)

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 積み上げをSQLで実行し、値を取得する
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pShinsaInfoList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ExecuteStackSQL(ByVal pDb As DBAccess, ByVal formula As String, ByVal itemNo As String) As Object
        Dim ret As Object
        Dim dt As DataTable = Nothing
        Dim sb As System.Text.StringBuilder = Nothing

        Try

            sb = New System.Text.StringBuilder

            'SQL文の設定
            sb.AppendLine("SET ANSI_WARNINGS OFF")
            sb.AppendLine("SELECT ")
            sb.AppendLine(If(String.IsNullOrEmpty(formula), "NULL", formula))
            sb.AppendLine(" AS [" & itemNo & "]")

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
    Private Function Round(ByVal value As Decimal, ByVal HyojiTani As Decimal) As Decimal
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
    ''' 集計結果表の項目情報リストに値を代入する
    ''' </summary>
    ''' <param name="itemInfo"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Private Sub AddValueInItemInfo(ByRef itemInfo As ItemInfo, ByVal value As Object)
        If itemInfo.Value.Count = 0 Then
            itemInfo.Value.Add(value)
        Else
            itemInfo.Value(0) = value
        End If
    End Sub

    ''' <summary>
    ''' 積み上げ計算式に集計倍率を埋め込む
    ''' </summary>
    ''' <param name="iteminfo"></param>
    ''' <param name="stackformula"></param>
    ''' <param name="chosakubun"></param>
    ''' <remarks></remarks>
    Private Sub SetSyukeiBairitsu(ByRef iteminfo As ItemInfo, stackformula As String, chosakubun As String)


        Select Case _syukeiInfo.heikinSyurui
            Case ComConst.平均種類.加重
                If iteminfo.SyukeibairitsuKbn.Equals(KBN_KEIEITAIBAIRITSU) Then
                    iteminfo.FormulaConvValue = "(" & stackformula & " * [" & ComConst.個別結果表.集計倍率(chosakubun) & "])"
                Else
                    iteminfo.FormulaConvValue = "(" & stackformula & " * [" & ComConst.個別結果表.部門集計倍率(chosakubun) & "])"
                End If

            Case ComConst.平均種類.総和
                iteminfo.FormulaConvValue = stackformula

            Case ComConst.平均種類.累積
                If iteminfo.SyukeibairitsuKbn.Equals(KBN_KEIEITAIBAIRITSU) Then
                    iteminfo.FormulaConvValue = stackformula & " * [" & ComConst.個別結果表.集計倍率(chosakubun) & "]"
                Else
                    iteminfo.FormulaConvValue = stackformula & " * [" & ComConst.個別結果表.部門集計倍率(chosakubun) & "]"
                End If

        End Select
    End Sub

    ''' <summary>
    ''' 集計倍率の積み上げ値を取得する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getSyukeiBairitsu() As String

        Dim iteminfo As ItemInfo = Nothing
        Dim total As Decimal = 0

        If _chosaKubun.Equals(ComConst.営農経営体区分.農業経営体) Then
            iteminfo = _ItemInfoList(ComConst.個別結果表.集計倍率(ComConst.調査区分.営農類型別経営統計_個人) & ComConst.ITEM_NO_DELIMITER & ComConst.調査区分.営農類型別経営統計_個人)
            For Each value As Object In iteminfo.Value
                If Not IsDBNull(value) Then
                    total += CDec(value)
                End If
            Next
            iteminfo = _ItemInfoList(ComConst.個別結果表.集計倍率(ComConst.調査区分.営農類型別経営統計_法人) & ComConst.ITEM_NO_DELIMITER & ComConst.調査区分.営農類型別経営統計_法人)
            For Each value As Object In iteminfo.Value
                If Not IsDBNull(value) Then
                    total += CDec(value)
                End If
            Next

        Else
            iteminfo = _ItemInfoList(ComConst.個別結果表.集計倍率(_chosaKubun) & ComConst.ITEM_NO_DELIMITER & _chosaKubun)
            For Each value As Object In iteminfo.Value
                If Not IsDBNull(value) Then
                    total += CDec(value)
                End If
            Next
        End If

        Return total.ToString
    End Function

    ''' <summary>
    ''' 【営農個別のみ】部門集計倍率の積み上げ値を取得する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getBumonSyukeiBairitsu() As String
        Dim total As Decimal = 0

        If _chosaKubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) Then
            Dim iteminfo As ItemInfo = Nothing
            Dim serchItemNo As String = ComConst.個別結果表.部門集計倍率(_chosaKubun) & ComConst.ITEM_NO_DELIMITER & _chosaKubun

            iteminfo = _ItemInfoList(serchItemNo)
            For Each value As Object In iteminfo.Value
                If Not IsDBNull(value) Then
                    total += CDec(value)
                End If
            Next
        End If

        Return total.ToString
    End Function

    ''' <summary>
    ''' 式に含まれている項番すべてが数値項目かを判定する
    ''' </summary>
    ''' <param name="formula"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isNumberItem(formula As String) As Boolean
        Dim ret As Boolean = True
        Dim matchList As MatchCollection = Nothing

        '括弧付き項目番号『[010101]』『[010101_1]』を検索
        matchList = Regex.Matches(formula, C_SearchKakko)

        For Each mat As Match In matchList
            '『[010101]』『[010101_1]』から項目番号『010101』『010101_1』を抽出する
            Dim itemNo As String = Regex.Match(mat.Value, C_SearchKakko).Value.Trim("["c, "]"c)
            Dim info As ItemInfo = _ItemInfoList(itemNo & ComConst.ITEM_NO_DELIMITER & If(_chosaKubun.Equals(ComConst.営農経営体区分.農業経営体) And itemNo.Contains("K"), ComConst.調査区分.営農類型別経営統計_個人, _chosaKubun))

            If info.ValueType = ComConst.型区分.文字型 Then
                ret = False
                Exit For
            End If
        Next

        Return ret

    End Function

    ''' <summary>
    ''' 調査区分と生産費平均値種類に該当する「総数を割る数」を取得する
    ''' </summary>
    ''' <param name="seisanhiHeikin"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getWeight(seisanhiHeikin As String) As String
        Dim val As String = Nothing

        For Each Key In ComConst.生産費平均値種類Weight.リスト.Keys
            If ComConst.生産費平均値種類Weight.リスト(Key).調査区分.Contains(_chosaKubun) And ComConst.生産費平均値種類Weight.リスト(Key).平均値種類.Contains(seisanhiHeikin) Then
                val = ComConst.生産費平均値種類Weight.リスト(Key).割る数
            End If
        Next

        Return val

    End Function

    ''' <summary>
    ''' 【】で括られた固定文字列論理を実データに置換する
    ''' </summary>
    ''' <param name="info"></param>
    ''' <param name="pChiiki"></param>
    ''' <param name="kobetsuRecordCount"></param>
    ''' <param name="pEinouChusyutsuJouken"></param>
    ''' <param name="pSeisanhiChusyutsuJouken"></param>
    ''' <remarks></remarks>
    Private Function replaceRonri(str As String, pChiiki As String, kobetsuRecordCount As Integer, pEinouChusyutsuJouken As DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu, pSeisanhiChusyutsuJouken As DAOKobetsuKekkahyo.SeisanhiChusyutsu) As String
        '『【調査年】』を値に置換
        str = str.Replace("【調査年】", _chosaYear)
        '『【集計番号】』を値に置換
        str = str.Replace("【集計番号】", "'" & _syukeiInfo.syukeiNo & "'")
        '『【集計名称】』を値に置換
        str = str.Replace("【集計名称】", "'" & _syukeiInfo.syukeiName & "'")
        '『【集計パターン】』を値に置換
        str = str.Replace("【集計パターン】", "'" & _syukeiInfo.syukeipattern & "'")
        ' REV_002↓
        '『【基本・詳細項目集計】』を値に置換
        str = str.Replace("【基本・詳細項目集計】", "'" & _syukeiInfo.kihonSyosaiSyukei & "'")
        '『【集計１】』を値に置換
        str = str.Replace("【集計１】", If(_syukeiInfo.syukei1 Is Nothing, "NULL", "'" & ComConst.集計１.リスト(ComUtil.getVersionKubunTaikei(_chosaYear, _chosaKubun))(_syukeiInfo.syukei1).名称 & "'"))
        '『【集計２】』を値に置換
        str = str.Replace("【集計２】", If(_syukeiInfo.syukei2 Is Nothing, "NULL", "'" & ComConst.集計２.リスト(ComUtil.getVersionKubunTaikei(_chosaYear, _chosaKubun))(_syukeiInfo.syukei2).名称 & "'"))
        '『【集計３】』を値に置換
        str = str.Replace("【集計３】", If(_syukeiInfo.syukei3 Is Nothing, "NULL", "'" & ComConst.集計３.リスト(ComUtil.getVersionKubunTaikei(_chosaYear, _chosaKubun))(_syukeiInfo.syukei3).名称 & "'"))
        '『【集計４】』を値に置換
        str = str.Replace("【集計４】", If(_syukeiInfo.syukei4 Is Nothing, "NULL", "'" & ComConst.集計４.リスト(ComUtil.getVersionKubunTaikei(_chosaYear, _chosaKubun))(_syukeiInfo.syukei4).名称 & "'"))
        ' REV_002↑
        '『【平均種類】』を値に置換
        str = str.Replace("【平均種類】", _syukeiInfo.heikinSyurui)
        '『【生産費平均値種類】』を値に置換
        str = str.Replace("【生産費平均値種類】", _syukeiInfo.seisanhiHeikin)
        '『【地域】』を値に置換
        str = str.Replace("【地域】", pChiiki)
        '『【規模階層】』を値に置換
        str = str.Replace("【規模階層】", _syukeiInfo.kiboKaisou)
        '『【個別結果表レコード数】』を値に置換
        str = str.Replace("【個別結果表レコード数】", kobetsuRecordCount.ToString)
        '『【集計倍率】』を値に変換
        str = str.Replace("【集計倍率】", _syukeiBairitsu)

        '生産費のみ
        If Not pSeisanhiChusyutsuJouken Is Nothing Then
            '『【集計区分】』を値に変換
            str = str.Replace("【集計区分】", pSeisanhiChusyutsuJouken.syukeiKbn)
            '『【田畑区分】』を値に変換
            str = str.Replace("【田畑区分】", pSeisanhiChusyutsuJouken.tahataKbn)
            '『【ビール麦販売区分】』を値に変換
            str = str.Replace("【ビール麦販売区分】", pSeisanhiChusyutsuJouken.beerKbn)
            '『【てんさい栽培区分】』を値に変換
            str = str.Replace("【てんさい栽培区分】", pSeisanhiChusyutsuJouken.tensaiKbn)
        End If

        '営農のみ
        If Not pEinouChusyutsuJouken Is Nothing Then
            '『【部門集計】』を値に置換
            str = str.Replace("【部門集計】", If(pEinouChusyutsuJouken.bumonCd Is Nothing, "0", "1"))
            '『【部門コード】』を値に置換
            str = str.Replace("【部門コード】", If(pEinouChusyutsuJouken.bumonCd Is Nothing, "0", pEinouChusyutsuJouken.bumonCd))
            '『【営農類型】』を値に置換
            str = str.Replace("【営農類型】", "'" & pEinouChusyutsuJouken.einouruikei & "'")
            If _chosaKubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) Then
                '『【指定部門集計倍率】』を値に置換
                str = str.Replace("【指定部門集計倍率】", _bumonSyukeiBairitsu)
            End If
        End If


        '『【抽出条件】』を値に置換 抽出条件内に【】がある場合がある為、最後に置換
        str = str.Replace("【抽出条件】", "'" & _syukeiInfo.niniSyukeiShihyou & "'")

        Return str
    End Function

    ''' <summary>
    ''' 固定文字列論理を論理式のエラーチェック用に全て０で置換する
    ''' </summary>
    ''' <param name="pFormula"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function replaceRonriToErrorCheck(pFormula As String) As String

        '積み上げ計算式判定用のΣを除去
        pFormula = pFormula.Replace("Σ", "")
        '積み上げ計算式判定用のσを除去
        pFormula = pFormula.Replace("σ", "")

        '【生産費の総数以外】『Ｗ』を0に変換
        pFormula = pFormula.Replace("Ｗ", "0")

        '【】で括られた固定文字列論理を0に変換
        '【調査年】
        pFormula = pFormula.Replace("【調査年】", "0")
        '『【集計番号】』
        pFormula = pFormula.Replace("【集計番号】", "'a'")
        '『【集計名称】』
        pFormula = pFormula.Replace("【集計名称】", "'a'")
        '『【集計パターン】』
        pFormula = pFormula.Replace("【集計パターン】", "'a'")
        '『【集計１】』
        pFormula = pFormula.Replace("【集計１】", "'a'")
        '『【集計２】』
        pFormula = pFormula.Replace("【集計２】", "'a'")
        '『【集計３】』
        pFormula = pFormula.Replace("【集計３】", "'a'")
        '『【集計４】』
        pFormula = pFormula.Replace("【集計４】", "'a'")
        '『【平均種類】』
        pFormula = pFormula.Replace("【平均種類】", "0")
        '『【生産費平均値種類】』
        pFormula = pFormula.Replace("【生産費平均値種類】", "0")
        '『【地域】』
        pFormula = pFormula.Replace("【地域】", "0")
        '『【規模階層】』
        pFormula = pFormula.Replace("【規模階層】", "0")
        '『【個別結果表レコード数】』
        pFormula = pFormula.Replace("【個別結果表レコード数】", "0")
        '『【集計倍率】』
        pFormula = pFormula.Replace("【集計倍率】", "0")
        '『【集計区分】』
        pFormula = pFormula.Replace("【集計区分】", "0")
        '『【田畑区分】』
        pFormula = pFormula.Replace("【田畑区分】", "0")
        '『【ビール麦販売区分】』
        pFormula = pFormula.Replace("【ビール麦販売区分】", "0")
        '『【てんさい栽培区分】』
        pFormula = pFormula.Replace("【てんさい栽培区分】", "0")
        '『【部門集計】』
        pFormula = pFormula.Replace("【部門集計】", "0")
        '『【部門コード】』
        pFormula = pFormula.Replace("【部門コード】", "0")
        '『【営農類型】』
        pFormula = pFormula.Replace("【営農類型】", "'a'")
        '『【抽出条件】』
        pFormula = pFormula.Replace("【抽出条件】", "'a'")
        '『【指定部門集計倍率】』
        pFormula = pFormula.Replace("【指定部門集計倍率】", "0")
        ' REV_002↓
        '『【基本・詳細項目集計】』
        pFormula = pFormula.Replace("【基本・詳細項目集計】", "0")
        ' REV_002↑

        Return pFormula
    End Function

    ''' <summary>
    ''' 集計用センサス番号管理テーブルを取得しリストにして返す
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getSyukeiCensusNoList(pDb As DBAccess) As List(Of String)
        Dim dt As DataTable = DAOOther.GetSyukeiCensusNo(pDb, _chosaYear)
        Dim list As List(Of String) = Nothing

        If dt.Rows.Count > 0 Then
            list = New List(Of String)
            For Each dr As DataRow In dt.Rows
                list.Add(dr("センサス番号").ToString)
            Next
        End If

        Return list
    End Function

    ''' <summary>
    ''' 作成論理の「計算式」で使用されている項目番号が項目マスタに存在するかをチェックする
    ''' </summary>
    ''' <param name="pFormula"></param>
    ''' <param name="pDtKobetsuItemMst"></param>
    ''' <param name="pDtSyukeiItemMst"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckExistItemNo(pFormula As String, pDtKobetsuItemMst As DataTable, pDtSyukeiItemMst As DataTable, pSeisanhiHeikin As String) As Boolean
        Dim ret As Boolean = True
        Dim dt As DataTable = Nothing

        Try

            '項目マスタがデータ無しの場合
            If pDtKobetsuItemMst Is Nothing OrElse pDtKobetsuItemMst.Rows.Count = 0 OrElse
               pDtSyukeiItemMst Is Nothing OrElse pDtSyukeiItemMst.Rows.Count = 0 Then
                ret = False
                Exit Try
            End If

            '『[』『]』で囲まれた文字列検索する
            Dim matchList As MatchCollection = Regex.Matches(pFormula, C_SearchKakko)
            For Each mat As Match In matchList
                Dim itemNo As String = mat.Value.Trim("["c, "]"c).Replace("前", "")
                Dim isVariable As Boolean = False
                Dim itemNoInitial As String = itemNo.Substring(0, 1)

                Select Case itemNoInitial
                    Case "K"
                        dt = pDtKobetsuItemMst
                    Case "S"
                        dt = pDtSyukeiItemMst
                    Case Else
                        ret = False
                        Exit For
                End Select

                '項目番号が項目マスタに存在しない場合
                If itemNoInitial.Equals("S") And pSeisanhiHeikin.Equals("総数以外") Then
                    Dim itemNoSplit As String() = itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))
                    If Not (From n In dt.AsEnumerable Where n("項目番号").ToString = itemNoSplit(0)).Any Then
                        ret = False
                        Exit For
                    End If

                    If itemNoSplit.Length > 1 Then
                        If (Not itemNoSplit.Length = 2) Or (Not itemNoSplit(1).Equals("1")) Then
                            ret = False
                            Exit For
                        End If
                    End If

                Else
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
    Public Function CheckExecutableSQL(pItemNo As String, pFormula As String, pDtKobetsuItemMst As DataTable, pDtSyukeiItemMst As DataTable, pSeisanhiHeikin As String) As Boolean
        Dim ret As Boolean
        Dim itemNo As String
        Dim infoTarget As ItemInfo

        Try
            _isErrCheck = True
            itemNo = pItemNo.Trim("["c, "]"c)

            '「計算式」で使用されている項目番号が項目マスタに存在しない場合
            If Not Me.CheckExistItemNo(pFormula, pDtKobetsuItemMst, pDtSyukeiItemMst, pSeisanhiHeikin) Then
                ret = False
                Exit Try
            End If

            '項目情報リスト(項目マスタ)に存在しない場合、項目の型が不明なのでエラーとして返す
            infoTarget = (From n In _ItemInfoList.Values Where n.ItemNo = itemNo).FirstOrDefault
            If infoTarget Is Nothing Then
                ret = False
                Exit Try
            End If

            '積み上げ計算
            If pFormula.Contains("Σ") Or pFormula.Contains("σ") Then
                '生産費平均値種類が「総数以外」ならエラーとして返す
                If pSeisanhiHeikin.Equals("総数以外") Then
                    ret = False
                    Exit Try
                End If

                '積み上げ式の指標が２つ以上ある場合エラーとして返す
                If pFormula.Split("Σ"c).Length + pFormula.Split("σ"c).Length > 3 Then
                    ret = False
                    Exit Try
                End If

                '積み上げ式で文字型の項番が含まれている場合エラーとして返す
                If Not isNumberItem(pFormula) Then
                    ret = False
                    Exit Try
                End If

                '積み上げ式でRANKが含まれている場合エラーとして返す
                Dim matchList As MatchCollection = Regex.Matches(pFormula, C_SearchRank, RegexOptions.IgnoreCase)
                If matchList.Count > 0 Then
                    ret = False
                    Exit Try
                End If
            End If

            'ウエイト計算
            If pFormula.Contains("Ｗ") Then
                '営農ならエラーとして返す
                If ComUtil.IsEinou() Then
                    ret = False
                    Exit Try
                End If
                '生産費平均値種類が「総数以外」以外ならエラーとして返す
                If Not pSeisanhiHeikin.Equals("総数以外") Then
                    ret = False
                    Exit Try
                End If
            End If

            infoTarget.Formula = pFormula

            '計算式の文字列置換(項番以外)
            infoTarget.FormulaConv = Me.ConvertFormula(pFormula)

            '計算式の固定文字列論理を置換
            infoTarget.FormulaConvValue = Me.replaceRonriToErrorCheck(infoTarget.FormulaConv)

            '計算式内の項番を任意の文字・数値に置換
            ConvertFormulaItemNoToValue(_ItemInfoList, infoTarget, False, 0, _chosaKubun)
            If _chosaKubun.Equals(ComConst.営農経営体区分.農業経営体) Then
                ConvertFormulaItemNoToValue(_ItemInfoList, infoTarget, False, 0, ComConst.調査区分.営農類型別経営統計_個人)
            End If

            '計算式文字列の順位関連文字列を値に置換する
            Me.ConvertFormulaRankToValue(infoTarget)

            Try

                '作成実行
                ExecuteSQL(_db, infoTarget)
                ret = True

            Catch ex As Exception
                ret = False
            Finally
                '次回実行されないように数式をクリアする
                infoTarget.Formula = Nothing
                infoTarget.FormulaConvValue = Nothing
            End Try

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 【作成論理：Σ】積み上げを行う（加重・総和・累積の集計倍率をかける）
    ''' </summary>
    ''' <param name="grpInfo"></param>
    ''' <param name="pChiiki"></param>
    ''' <param name="pkobetsuRecordCount"></param>
    ''' <param name="pEinouChusyutsuJouken"></param>
    ''' <param name="pSeisanhiChusyutsuJouken"></param>
    ''' <param name="pStackFormula"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function calcStackAndAvarage(grpInfo As ItemInfo,
                                         pChiiki As String,
                                         pkobetsuRecordCount As Integer,
                                         pEinouChusyutsuJouken As DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu,
                                         pSeisanhiChusyutsuJouken As DAOKobetsuKekkahyo.SeisanhiChusyutsu) As Object

        Dim stackedValue As Decimal = 0

        '積み上げ対象の個別結果表 １レコード毎
        Dim stackFormula As String = Regex.Match(grpInfo.FormulaConv, C_SearchStackAverageFormula).Value.Replace("Σ", "")
        stackFormula = Me.ConvertFormula(stackFormula)

        grpInfo.FormulaConvValue = grpInfo.FormulaConv.Replace("Σ", "")

        Dim tmpFormula As String = grpInfo.FormulaConvValue

        '固定文字で示された論理を置換する
        tmpFormula = replaceRonri(tmpFormula, pChiiki, pkobetsuRecordCount, pEinouChusyutsuJouken, pSeisanhiChusyutsuJouken)

        Dim sb As New StringBuilder
        Dim isFormulaNull As Boolean = True
        If _chosaKubun.Equals(ComConst.営農経営体区分.農業経営体) Then
            '農業経営体の場合は個人と法人で分けて積み上げを行う
            For Each kv As KeyValuePair(Of String, Integer) In _einouSplitRecordCount
                Dim chosaKubun As String = kv.Key
                Dim count As Integer = kv.Value

                For i As Integer = 0 To count - 1
                    '積み上げ計算式に集計倍率を埋め込む
                    Me.SetSyukeiBairitsu(grpInfo, stackFormula, chosaKubun)

                    '計算式文字列の項目番号を値に置換する
                    If Not Me.ConvertFormulaItemNoToValue(_ItemInfoList, grpInfo, True, i, chosaKubun) Then
                        isFormulaNull = False
                    End If

                    '積み上げ計算式に追加する
                    If sb.Length > 0 Then
                        sb.Append(" + ")
                        sb.AppendLine(grpInfo.FormulaConvValue)
                    Else
                        sb.AppendLine(grpInfo.FormulaConvValue)
                    End If
                    '2000回に1回SQLを実行しその結果をSQLの途中式とする（SQLエラー防止）
                    If i <> 0 And i Mod STACK_MAX = 0 Then
                        Dim tmpValue As Object = Me.ExecuteStackSQL(_db, sb.ToString, grpInfo.ItemNo)
                        sb.Clear()
                        sb.AppendLine(If(tmpValue Is Nothing, "0", tmpValue.ToString))
                    End If
                Next
            Next
        Else

            For i As Integer = 0 To pkobetsuRecordCount - 1
                '積み上げ計算式に集計倍率を埋め込む
                Me.SetSyukeiBairitsu(grpInfo, stackFormula, _chosaKubun)

                '計算式文字列の項目番号を値に置換する
                If Not Me.ConvertFormulaItemNoToValue(_ItemInfoList, grpInfo, True, i, _chosaKubun) Then
                    isFormulaNull = False
                End If

                '積み上げ計算式に追加する
                If sb.Length > 0 Then
                    sb.Append(" + ")
                    sb.AppendLine(grpInfo.FormulaConvValue)
                Else
                    sb.AppendLine(grpInfo.FormulaConvValue)
                End If
                '2000回に1回SQLを実行しその結果をSQLの途中式とする（SQLエラー防止）
                If i <> 0 And i Mod STACK_MAX = 0 Then
                    Dim tmpValue As Object = Me.ExecuteStackSQL(_db, sb.ToString, grpInfo.ItemNo)
                    sb.Clear()
                    sb.AppendLine(If(tmpValue Is Nothing, "0", tmpValue.ToString))
                End If

            Next
        End If

        '積み上げ後の値の計算
        sb.Insert(0, "(").Append(")")
        If (_chosaKubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) Or _chosaKubun.Equals(ComConst.調査区分.営農類型別経営統計_法人) Or _chosaKubun.Equals(ComConst.営農経営体区分.農業経営体)) Then

            Select Case _syukeiInfo.heikinSyurui
                Case ComConst.平均種類.加重
                    '営農で加重平均の場合のみ、積み上げ後の値を集計倍率の積み上げ値で除算する
                    If grpInfo.SyukeibairitsuKbn.Equals(KBN_KEIEITAIBAIRITSU) Then
                        sb.Append(" / ").Append(_syukeiBairitsu)
                    Else
                        If _chosaKubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) Then
                            sb.Append(" / ").Append(_bumonSyukeiBairitsu)
                        End If
                    End If

                Case ComConst.平均種類.総和
                    '営農で総和平均の場合、積み上げ後の値を集計に使用した個別結果表のレコード数で除算する
                    sb.Append(" / ").Append(pkobetsuRecordCount)
            End Select

        End If
        grpInfo.FormulaConvValue = tmpFormula.Replace(stackFormula, If(pkobetsuRecordCount > 0, sb.ToString, "0"))

        '積み上げ以外の箇所の計算式文字列の項目番号を値に置換する
        Me.ConvertFormulaItemNoToValue(_ItemInfoList, grpInfo, False, 0, _chosaKubun)

        Dim stackedRet As Object = Me.ExecuteStackSQL(_db, If(isFormulaNull, Nothing, grpInfo.FormulaConvValue), grpInfo.ItemNo)

        Return If(IsDBNull(stackedRet), stackedRet, CDec(stackedRet))
    End Function

    ''' <summary>
    ''' 【作成論理：σ】積み上げを行う（加重累積値）
    ''' </summary>
    ''' <param name="grpInfo"></param>
    ''' <param name="pChiiki"></param>
    ''' <param name="pkobetsuRecordCount"></param>
    ''' <param name="pEinouChusyutsuJouken"></param>
    ''' <param name="pSeisanhiChusyutsuJouken"></param>
    ''' <param name="pStackFormula"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function calcStack(grpInfo As ItemInfo,
                               pChiiki As String,
                               pkobetsuRecordCount As Integer,
                               pEinouChusyutsuJouken As DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu,
                               pSeisanhiChusyutsuJouken As DAOKobetsuKekkahyo.SeisanhiChusyutsu) As Object

        Dim stackedValue As Decimal = 0

        '積み上げ対象の個別結果表 １レコード毎
        Dim stackFormula As String = Regex.Match(grpInfo.FormulaConv, C_SearchStackFormula).Value.Replace("σ", "")
        stackFormula = Me.ConvertFormula(stackFormula)

        grpInfo.FormulaConvValue = grpInfo.FormulaConv.Replace("σ", "")

        Dim tmpFormula As String = grpInfo.FormulaConvValue
        '固定文字で示された論理を置換する
        tmpFormula = replaceRonri(tmpFormula, pChiiki, pkobetsuRecordCount, pEinouChusyutsuJouken, pSeisanhiChusyutsuJouken)

        Dim sb As New StringBuilder
        Dim isFormulaNull As Boolean = True
        If _chosaKubun.Equals(ComConst.営農経営体区分.農業経営体) Then
            '農業経営体の場合は個人と法人で分けて積み上げを行う
            For Each kv As KeyValuePair(Of String, Integer) In _einouSplitRecordCount
                Dim chosaKubun As String = kv.Key
                Dim count As Integer = kv.Value

                For i As Integer = 0 To count - 1
                    grpInfo.FormulaConvValue = stackFormula
                    '計算式文字列の項目番号を値に置換する
                    If Not Me.ConvertFormulaItemNoToValue(_ItemInfoList, grpInfo, True, i, chosaKubun) Then
                        isFormulaNull = False
                    End If

                    '積み上げ計算式に追加する
                    If sb.Length > 0 Then
                        sb.Append(" + ")
                        sb.AppendLine(grpInfo.FormulaConvValue)
                    Else
                        sb.AppendLine(grpInfo.FormulaConvValue)
                    End If
                    '2000回に1回SQLを実行しその結果をSQLの途中式とする（SQLエラー防止）
                    If i <> 0 And i Mod STACK_MAX = 0 Then
                        Dim tmpValue As Object = Me.ExecuteStackSQL(_db, sb.ToString, grpInfo.ItemNo)
                        sb.Clear()
                        sb.AppendLine(If(tmpValue Is Nothing, "0", tmpValue.ToString))
                    End If
                Next
            Next
        Else

            For i As Integer = 0 To pkobetsuRecordCount - 1
                grpInfo.FormulaConvValue = stackFormula
                '計算式文字列の項目番号を値に置換する
                If Not Me.ConvertFormulaItemNoToValue(_ItemInfoList, grpInfo, True, i, _chosaKubun) Then
                    isFormulaNull = False
                End If

                '積み上げ計算式に追加する
                If sb.Length > 0 Then
                    sb.Append(" + ")
                    sb.AppendLine(grpInfo.FormulaConvValue)
                Else
                    sb.AppendLine(grpInfo.FormulaConvValue)
                End If
                '2000回に1回SQLを実行しその結果をSQLの途中式とする（SQLエラー防止）
                If i <> 0 And i Mod STACK_MAX = 0 Then
                    Dim tmpValue As Object = Me.ExecuteStackSQL(_db, sb.ToString, grpInfo.ItemNo)
                    sb.Clear()
                    sb.AppendLine(If(tmpValue Is Nothing, "0", tmpValue.ToString))
                End If
            Next
        End If

        sb.Insert(0, "(").Append(")")
        grpInfo.FormulaConvValue = tmpFormula.Replace(stackFormula, If(pkobetsuRecordCount > 0, sb.ToString, "0"))

        '積み上げ以外の箇所の計算式文字列の項目番号を値に置換する
        Me.ConvertFormulaItemNoToValue(_ItemInfoList, grpInfo, False, 0, _chosaKubun)
        Dim stackedRet As Object = Me.ExecuteStackSQL(_db, If(isFormulaNull, Nothing, grpInfo.FormulaConvValue), grpInfo.ItemNo)

        Return If(IsDBNull(stackedRet), stackedRet, CDec(stackedRet))
    End Function

    ''' <summary>
    ''' 作成論理の情報を項目リストへ
    ''' </summary>
    ''' <param name="pDtCreateRonri"></param>
    ''' <remarks></remarks>
    Public Sub addCreateRonriInfo(pDtCreateRonri As DataTable)

        If Not pDtCreateRonri Is Nothing Then
            For Each dr As DataRow In pDtCreateRonri.Rows
                Dim serchItemNo = dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & dr("調査区分").ToString
                Dim info As ItemInfo = _ItemInfoList(serchItemNo)
                info.ItemName = dr("項目名").ToString
                info.Priority = CDec(If(IsDBNull(dr("優先順位")), 0, dr("優先順位")))
                info.Formula = dr("計算式").ToString
                info.HyojiTani = CDec(If(IsDBNull(dr("表示単位")), 0, dr("表示単位")))
            Next
        End If

        '計算式の文字列置換(項番以外)
        For Each item As ItemInfo In (From n In _ItemInfoList.Values Where Not String.IsNullOrEmpty(n.Formula)).ToList
            item.FormulaConv = Me.ConvertFormula(item.Formula)
        Next

    End Sub

    ''' <summary>
    ''' 項目情報リストの値を初期化する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initializeItemValue()
        For Each info As ItemInfo In _ItemInfoList.Values
            info.Value = New List(Of Object)
        Next
    End Sub

End Class
