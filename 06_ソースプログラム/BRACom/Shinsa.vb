Imports System.Text.RegularExpressions
Imports System.Runtime.Serialization.Formatters.Binary

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.10.07 |大興電子通信        | 要件No.16 GetCountifSumifValueの返り値をDecimalに変更
'//  REV_002   | 2022.10.11 |大興電子通信        | 要件No1 バージョン区分追加
'//  REV_003   | 2023.08.08 |大興電子通信        | 要件No.8 個別結果表審査論理で前年の参照を可能にする
'//  REV_004   | 2023.11.27 |大興電子通信        | 要件No.20 メモリ不足対応,DBNUll対応
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 審査エラークラス
''' </summary>
''' <remarks></remarks>
Public Class ShinsaException
    Inherits Exception

    ''' <summary>センサス番号</summary>
    Private _censusNo As String

    ''' <summary>センサス番号</summary>
    Public ReadOnly Property CensusNo As String
        Get
            Return _censusNo
        End Get
    End Property

    ''' <summary>エラーサイン</summary>
    Private _errSign As String

    ''' <summary>エラーサイン</summary>
    Public ReadOnly Property ErrSign As String
        Get
            Return _errSign
        End Get
    End Property

    Public Sub New(inner As Exception, ByVal pCensusNo As String, ByVal pErrSign As String, ByVal pErrJoken As String, ByVal pErrJokenCnv As String)
        MyBase.New(String.Format("審査エラー(エラーサイン:{0} エラーとなる条件:{1} エラーとなる条件(変換後):{2})", pErrSign, pErrJoken, pErrJokenCnv), inner)
        _censusNo = pCensusNo
        _errSign = pErrSign
    End Sub
End Class

''' <summary>
''' 審査クラス
''' </summary>
''' <remarks></remarks>
Public Class Shinsa

    ''' <summary>
    ''' 審査情報クラス
    ''' </summary>
    ''' <remarks></remarks>
    <Serializable()> Public Class ShinsaInfo
        ''' <summary>エラーサイン</summary>
        Public ErrSign As String
        ''' <summary>チェック項目名</summary>
        Public CheckItemName As String
        ''' <summary>エラー内容</summary>
        Public ErrNaiyo As String
        ''' <summary>エラー区分</summary>
        Public ErrKubun As String
        ''' <summary>エラーとなる条件</summary>
        Public ErrJoken As String
        ''' <summary>エラーとなる条件(可変項目追加)</summary>
        Public ErrJokenConv As String
        ''' <summary>エラーとなる条件(可変項目追加＆値変換)</summary>
        Public ErrJokenConvValue As String
        ''' <summary>繰り返しかどうか</summary>
        Public IsRepeat As Boolean
        ''' <summary>明細番号</summary>
        Public MeisaiNo As Integer
        ''' <summary>審査結果(True:エラーなし、False：エラーあり)</summary>
        Public Result As Boolean = True
    End Class

    ''' <summary>
    ''' 審査結果クラス
    ''' </summary>
    ''' <remarks></remarks>
    <Serializable()> Public Class ShinsaKekka
        ''' <summary>センサス番号</summary>
        Public CensusNo As String
        ''' <summary>審査情報リスト</summary>
        Public ShinsaInfoList As List(Of ShinsaInfo)
    End Class

    ''' <summary>
    ''' 項目情報クラス
    ''' </summary>
    ''' <remarks>アンダースコア付きの可変項目番号はここには入らない</remarks>
    <Serializable()> Public Class ItemInfo
        Implements System.IComparable
        ''' <summary>項目番号</summary>
        Public ItemNo As String
        ''' <summary>項目の型</summary>
        Public Type As String
        ''' <summary>可変項目かどうか</summary>
        Public IsVariable As Boolean
        ''' <summary>可変最大数</summary>
        Public VariableMax As Integer

        ''' <summary>
        ''' 値を比較した結果を返す
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CompareTo(ByVal obj As Object) As Integer _
            Implements System.IComparable.CompareTo

            '項目番号を比較する
            Return Me.ItemNo.CompareTo(DirectCast(obj, ItemInfo).ItemNo)
        End Function
    End Class

    ''' <summary>
    ''' 審査対象データクラス
    ''' </summary>
    ''' <remarks></remarks>
    <Serializable()> Public Class TargetData
        ''' <summary>項目番号</summary>
        Public ItemNo As String
        ''' <summary>値</summary>
        Public Value As Object
        ''' <summary>明細番号(可変項目の場合のみ)</summary>
        Public MeisaiNo As Integer
    End Class

    ''' <summary>一度に実行(SELECT)する審査の数</summary>
    Private Const C_ShinsaUnitCount As Integer = 1000

    ''' <summary>正規表現検索文字列</summary>
    Public Const C_SearchKakko As String = "\[前?[QKC]+[0-9_]+\]"
    Private Const C_SearchCount As String = "COUNT\(" & C_SearchKakko & "\)"
    Private Const C_SearchCountif As String = "COUNT\(IIF\([^\(\)]*(\([^\(\)]*(\([^\(\)]*[^\(\)]*\)[^\(\)]*)*[^\(\)]*\)[^\(\)]*)*[^\(\)]*\)\)"
    Private Const C_SearchSum As String = "SUM\(" & C_SearchKakko & "\)"
    Private Const C_SearchSumif As String = "SUM\(IIF\([^\(\)]*(\([^\(\)]*(\([^\(\)]*[^\(\)]*\)[^\(\)]*)*[^\(\)]*\)[^\(\)]*)*[^\(\)]*\)\)"

    ''' <summary>審査論理の繰り返し判定用文字</summary>
    Private Const C_RepeatStr As String = "○"

    ''' <summary>DB接続文字列</summary>
    Private _connectionString As String

    ''' <summary>DBAccessオブジェクト</summary>
    Private _db As DBAccess

    ''' <summary>調査年</summary>
    Private _chosaYear As String

    ''' <summary>事務所番号</summary>
    Private _jimusho As String

    ''' <summary>審査データ種別</summary>
    Private _shinsaDataType As String

    ''' <summary>審査論理種別</summary>
    Private _shinsaRonriType As String

    ''' <summary>調査区分</summary>
    Private _chosaKubun As String

    ''' <summary>審査結果リスト</summary>
    Private _shinsaKekkaList As List(Of ShinsaKekka)

    ''' <summary>項目情報リスト</summary>
    Private _itemInfoList As List(Of ItemInfo)

    ''' <summary>審査対象データリスト</summary>
    Private _targetDataList As List(Of TargetData)

    ''' <summary>項目マスタ</summary>
    Private _dtItemMst As DataTable

    ''' <summary>エラーチェックかどうか</summary>
    Private _isErrCheck As Boolean

    ''' <summary>進捗ダイアログ</summary>
    Private _progressDialog As ProgressDialog

    ''' <summary>最大エラー件数</summary>
    Private _errMaxCount As Integer

    ' REV_002↓
    ''' <summary>バージョン区分</summary>
    Private _versionKbn As String
    ' REV_002↓

    ''' <summary>エラーなし情報削除フラグ</summary>
    Private _delResultTrue As Boolean

    ''' <summary>
    ''' コンストラクタ(審査論理エラーチェック用)
    ''' </summary>
    ''' <param name="pDb">DBAccessオブジェクト</param>
    ''' <param name="pChosaKubun">調査区分</param>
    ''' <param name="pShinsaDataType">審査論理データ種別(Comconst.審査論理データ種別)</param>
    ''' <param name="pShinsaRonriType">審査論理種別(Comconst.審査論理種別)</param>
    ''' <param name="pDtItemMst">項目マスタ</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pDb As DBAccess,
                   ByVal pChosaKubun As String,
                   ByVal pShinsaDataType As String,
                   ByVal pShinsaRonriType As String,
                   ByVal pDtItemMst As DataTable)
        Try
            _shinsaKekkaList = New List(Of ShinsaKekka)

            'DBAccessオブジェクト
            _db = pDb
            '調査区分
            _chosaKubun = pChosaKubun
            '審査データ種別
            _shinsaDataType = pShinsaDataType
            '審査論理種別
            _shinsaRonriType = pShinsaRonriType
            '審査結果リスト
            _shinsaKekkaList.Add(New ShinsaKekka With {.CensusNo = String.Empty, .ShinsaInfoList = New List(Of ShinsaInfo)})
            '項目マスタ
            _dtItemMst = pDtItemMst
            'エラーチェックかどうか
            _isErrCheck = True

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' コンストラクタ(入力・修正画面用)
    ''' </summary>
    ''' <param name="pConnectionString">DB接続文字列</param>
    ''' <param name="pChosaKubun">調査区分</param>
    ''' <param name="pShinsaDataType">審査論理データ種別(Comconst.審査論理データ種別)</param>
    ''' <param name="pShinsaRonriType">審査論理種別(Comconst.審査論理種別)</param>
    ''' <param name="pChosaYear">調査年</param>
    ''' <param name="pJimusho">事務所番号</param>
    ''' <param name="pTargetDataList">審査対象データリスト</param>
    ''' <param name="pProgressDialog">進捗ダイアログ</param>
    ''' <param name="pErrMaxCount">最大エラー件数</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pConnectionString As String,
                   ByVal pChosaKubun As String,
                   ByVal pShinsaDataType As String,
                   ByVal pShinsaRonriType As String,
                   ByVal pChosaYear As String,
                   ByVal pJimusho As String,
                   ByVal pTargetDataList As List(Of TargetData),
                   ByVal pProgressDialog As ProgressDialog,
                   Optional ByVal pErrMaxCount As Integer = 0)
        Try
            _shinsaKekkaList = New List(Of ShinsaKekka)

            'DB接続文字列
            _connectionString = pConnectionString
            '調査区分
            _chosaKubun = pChosaKubun
            '審査データ種別
            _shinsaDataType = pShinsaDataType
            '審査論理種別
            _shinsaRonriType = pShinsaRonriType
            '調査年
            _chosaYear = pChosaYear
            '事務所番号
            _jimusho = pJimusho
            '審査対象データリスト
            _targetDataList = pTargetDataList
            '審査結果リスト
            _shinsaKekkaList.Add(New ShinsaKekka With {.CensusNo = String.Empty, .ShinsaInfoList = New List(Of ShinsaInfo)})
            '進捗ダイアログ
            _progressDialog = pProgressDialog
            '最大エラー件数
            _errMaxCount = pErrMaxCount

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' コンストラクタ(エラーチェックリスト用)
    ''' </summary>
    ''' <param name="pConnectionString">DB接続文字列</param>
    ''' <param name="pChosaKubun">調査区分</param>
    ''' <param name="pShinsaDataType">審査論理データ種別(Comconst.審査論理データ種別)</param>
    ''' <param name="pShinsaRonriType">審査論理種別(Comconst.審査論理種別)</param>
    ''' <param name="pChosaYear">調査年</param>
    ''' <param name="pCensusNoList">センサス番号リスト</param>
    ''' <param name="pProgressDialog">進捗ダイアログ</param>
    ''' <param name="pDelResultTrue">エラーなし情報削除フラグ</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pConnectionString As String,
                   ByVal pChosaKubun As String,
                   ByVal pShinsaDataType As String,
                   ByVal pShinsaRonriType As String,
                   ByVal pChosaYear As String,
                   ByVal pCensusNoList As List(Of String),
                   ByVal pProgressDialog As ProgressDialog,
                   Optional ByVal pDelResultTrue As Boolean = False)
        Try
            _shinsaKekkaList = New List(Of ShinsaKekka)

            'DB接続文字列
            _connectionString = pConnectionString
            '調査区分
            _chosaKubun = pChosaKubun
            '審査データ種別
            _shinsaDataType = pShinsaDataType
            '審査論理種別
            _shinsaRonriType = pShinsaRonriType
            '調査年
            _chosaYear = pChosaYear
            ' REV_002↓
            _versionKbn = ComUtil.getVersionKubunTaikei(_chosaYear, CommonInfo.Chosakubun)
            ' REV_002↑

            '審査結果リスト
            For Each CensusNo In pCensusNoList
                _shinsaKekkaList.Add(New ShinsaKekka With {.CensusNo = CensusNo, .ShinsaInfoList = New List(Of ShinsaInfo)})
            Next
            '進捗ダイアログ
            _progressDialog = pProgressDialog
            'エラーなし情報削除フラグ
            _delResultTrue = pDelResultTrue

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ' REV_002↓
    ''' <summary>
    ''' コンストラクタ(個別結果表審査論理エラーチェック用)
    ''' </summary>
    ''' <param name="pDb">DBAccessオブジェクト</param>
    ''' <param name="pChosaKubun">調査区分</param>
    ''' <param name="pShinsaDataType">審査論理データ種別(Comconst.審査論理データ種別)</param>
    ''' <param name="pShinsaRonriType">審査論理種別(Comconst.審査論理種別)</param>
    ''' <param name="pDtItemMst">項目マスタ</param>
    ''' <param name="versionKbn">バージョン区分</param>
    Public Sub New(ByVal pDb As DBAccess,
                   ByVal pChosaKubun As String,
                   ByVal pShinsaDataType As String,
                   ByVal pShinsaRonriType As String,
                   ByVal pDtItemMst As DataTable,
                   ByVal versionKbn As String)
        Try
            _shinsaKekkaList = New List(Of ShinsaKekka)

            'DBAccessオブジェクト
            _db = pDb
            '調査区分
            _chosaKubun = pChosaKubun
            '審査データ種別
            _shinsaDataType = pShinsaDataType
            '審査論理種別
            _shinsaRonriType = pShinsaRonriType
            '審査結果リスト
            _shinsaKekkaList.Add(New ShinsaKekka With {.CensusNo = String.Empty, .ShinsaInfoList = New List(Of ShinsaInfo)})
            '項目マスタ
            _dtItemMst = pDtItemMst
            'エラーチェックかどうか
            _isErrCheck = True

            ' REV_002↓
            'バージョン区分
            _versionKbn = versionKbn
            ' REV_002↑

        Catch ex As Exception
            Throw
        End Try
    End Sub
    ' REV_002↓

    ''' <summary>
    ''' 審査実行
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Execute() As List(Of ShinsaKekka)
        '審査情報リスト
        Dim shinsaInfoList As List(Of ShinsaInfo) = Nothing
        '審査対象データリスト
        Dim targetDataList As List(Of TargetData) = Nothing

        Try
            Using db As New DBAccess(_connectionString)

                '項目定義情報取得
                _itemInfoList = GetItemInfo(db)
                If _itemInfoList.Count = 0 Then
                    Exit Try
                End If

                'REV_003↓個別結果表審査論理の場合、項目定義情報(前年)追加
                If _shinsaDataType = ComConst.審査論理データ種別.個別結果表 Then
                    _itemInfoList.AddRange(_itemInfoList.Select(Function(x) New ItemInfo With {
                            .ItemNo = "前" & x.ItemNo,
                            .Type = x.Type,
                            .IsVariable = x.IsVariable,
                            .VariableMax = x.VariableMax}).ToList())
                End If
                'REV_003↑

                '項目番号でソートする
                _itemInfoList.Sort()

                '審査情報取得
                shinsaInfoList = GetShinsaInfo(db)
                If shinsaInfoList.Count = 0 Then
                    Exit Try
                End If

                'エラーとなる条件の文字列変換(項番以外)
                For Each ShinsaInfo As ShinsaInfo In shinsaInfoList
                    ShinsaInfo.ErrJokenConvValue = ConvertShinsaNaiyo(ShinsaInfo.ErrJoken)
                Next

                For Each ShinsaKekka In _shinsaKekkaList

                    '審査情報をセット
                    ShinsaKekka.ShinsaInfoList = New List(Of ShinsaInfo)
                    ShinsaKekka.ShinsaInfoList.AddRange(DeepCopy(shinsaInfoList))

                    'REV_003↓
                    Dim hasPrevData = False
                    'REV_003↑

                    '入力・修正画面からの審査の場合
                    If String.IsNullOrEmpty(ShinsaKekka.CensusNo) Then
                        targetDataList = _targetDataList
                    Else
                        '審査対象データ取得
                        If IsChosahyo() Then
                            targetDataList = GetChosahyo(db, New DAOChosahyo.PrimaryKey(_chosaYear, ShinsaKekka.CensusNo), _itemInfoList)
                        Else
                            targetDataList = GetKobetsukekkahyo(db, New DAOKobetsuKekkahyo.PrimaryKey(_chosaYear, ShinsaKekka.CensusNo), _itemInfoList)
                            'REV_003↓前年個別結果表データ追加
                            Dim targetDataListPrev = GetKobetsukekkahyo(db, New DAOKobetsuKekkahyo.PrimaryKey(_chosaYear, ShinsaKekka.CensusNo), _itemInfoList, True)
                            If targetDataListPrev.Count > 0 Then
                                hasPrevData = True
                                targetDataList.AddRange(targetDataListPrev)
                            End If
                            'REV_003↑
                        End If
                    End If

                    '審査内容文字列変換(項番)
                    'REV_003↓
                    'ShinsaKekka.ShinsaInfoList = ConvertShinsaNaiyo(db, ShinsaKekka, targetDataList, _itemInfoList)
                    ShinsaKekka.ShinsaInfoList = ConvertShinsaNaiyo(db, ShinsaKekka, targetDataList, _itemInfoList, hasPrevData)
                    'REV_003↑

                    Try
                        '審査実行
                        ExecuteShinsa(db, ShinsaKekka.ShinsaInfoList)

                    Catch ex As Exception
                        '審査実行
                        ExecuteShinsaSingle(db, ShinsaKekka.CensusNo, ShinsaKekka.ShinsaInfoList)
                    End Try

                    If _delResultTrue Then
                        'エラーなし情報削除
                        ShinsaKekka.ShinsaInfoList.RemoveAll(Function(s As ShinsaInfo) s.Result = True)
                    End If

                Next

            End Using

        Catch ex As Exception
            Throw
        End Try

        Return _shinsaKekkaList
    End Function

    ''' <summary>
    ''' 「エラーとなる条件」で使用されている項目番号が項目マスタに存在するかをチェックする
    ''' </summary>
    ''' <param name="pErrSign"></param>
    ''' <param name="pErrJoken"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckExistItemNo(ByVal pErrSign As String, ByVal pErrJoken As String) As Boolean
        Dim ret As Boolean = True

        Try

            '項目マスタがデータ無しの場合
            If _dtItemMst Is Nothing OrElse _dtItemMst.Rows.Count = 0 Then
                ret = False
                Exit Try
            End If

            '『[』『]』で囲まれた文字列検索する
            Dim matchList As MatchCollection = Regex.Matches(pErrJoken, Shinsa.C_SearchKakko)
            For Each mat As Match In matchList
                Dim itemNo As String = mat.Value.Trim("["c, "]"c)
                'REV_003↓個別結果表審査論理の場合、前付項番を許容する
                If _shinsaDataType = ComConst.審査論理データ種別.個別結果表 AndAlso itemNo.StartsWith("前") Then
                    itemNo = itemNo.Substring(1)
                End If
                'REV_003↑
                If itemNo.Contains(ComConst.ITEM_NO_DELIMITER) Then
                    Dim variItemNo As String = itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(0)
                    Dim meisaiNo As Integer = CType(itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(1), Integer)
                    '可変項目番号が項目マスタに存在しない場合
                    If Not (From n In _dtItemMst.AsEnumerable Where n("項目番号").ToString = variItemNo).Any Then
                        ret = False
                        Exit For
                    End If
                    '可変項目番号が項目マスタに存在し、明細番号が可変最大数を超えている場合
                    If (From n In _dtItemMst.AsEnumerable Where n("項目番号").ToString = variItemNo And CInt(n("可変最大数")) < meisaiNo).Any Then
                        ret = False
                        Exit For
                    End If
                Else
                    '項目番号が項目マスタに存在しない場合
                    If Not (From n In _dtItemMst.AsEnumerable Where n("項目番号").ToString = itemNo).Any Then
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
    ''' 「繰り返し」が「○」の審査論理において「エラーとなる条件」に可変項目が存在するかをチェックする
    ''' </summary>
    ''' <param name="pErrSign"></param>
    ''' <param name="pErrJoken"></param>
    ''' <param name="pRepeat"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckRepeat(ByVal pErrSign As String, ByVal pErrJoken As String, ByVal pRepeat As String) As Boolean
        Dim ret As Boolean = False

        Try

            '項目マスタがデータ無しの場合
            If _dtItemMst Is Nothing OrElse _dtItemMst.Rows.Count = 0 Then
                ret = False
                Exit Try
            End If

            '繰り返しではない場合
            If pRepeat <> C_RepeatStr Then
                ret = True
                Exit Try
            End If

            '『[』『]』で囲まれた文字列検索する
            Dim matchList As MatchCollection = Regex.Matches(pErrJoken, Shinsa.C_SearchKakko)
            For Each mat As Match In matchList
                Dim itemNo As String = mat.Value.Trim("["c, "]"c)
                '可変区分=1の項目番号が項目マスタに存在する場合
                If (From n In _dtItemMst.AsEnumerable Where CInt(n("可変区分")) = 1 And n("項目番号").ToString = itemNo).Any Then
                    ret = True
                    Exit For
                End If
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 「エラーとなる条件」がSQLで実行可能かをチェックする
    ''' </summary>
    ''' <param name="pErrSign"></param>
    ''' <param name="pErrJoken"></param>
    ''' <param name="pRepeat"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckExecutableSQL(ByVal pErrSign As String, ByVal pErrJoken As String, Optional ByVal pRepeat As String = Nothing) As Boolean
        Dim ret As Boolean

        Try

            '「エラーとなる条件」で使用されている項目番号が項目マスタに存在しない場合
            If Not Me.CheckExistItemNo(pErrSign, pErrJoken) Then
                ret = False
                Exit Try
            End If

            _itemInfoList = New List(Of ItemInfo)
            For Each dr As DataRow In _dtItemMst.Rows
                '項目定義情報リスト作成
                _itemInfoList.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString,
                                                     .Type = dr("型区分").ToString,
                                                     .IsVariable = If(IsChosahyo(), CBool(dr("可変区分")), False),
                                                     .VariableMax = If(IsChosahyo(), CInt(dr("可変最大数")), 0)})
            Next

            'REV_003↓個別結果表審査論理の場合、項目定義情報(前年)追加
            Dim hasPrevData = False
            If _shinsaDataType = ComConst.審査論理データ種別.個別結果表 Then
                hasPrevData = True
                _itemInfoList.AddRange(_itemInfoList.Select(Function(x) New ItemInfo With {
                            .ItemNo = "前" & x.ItemNo,
                            .Type = x.Type,
                            .IsVariable = x.IsVariable,
                            .VariableMax = x.VariableMax}).ToList())
            End If
            'REV_003↑

            '審査情報取得
            _shinsaKekkaList(0).ShinsaInfoList = New List(Of ShinsaInfo)
            _shinsaKekkaList(0).ShinsaInfoList.Add(New ShinsaInfo With {.ErrSign = pErrSign, .ErrJoken = pErrJoken, .IsRepeat = If(pRepeat = C_RepeatStr, True, False)})

            'エラーとなる条件の文字列変換(項番以外)
            For Each ShinsaInfo As ShinsaInfo In _shinsaKekkaList(0).ShinsaInfoList
                ShinsaInfo.ErrJokenConvValue = ConvertShinsaNaiyo(ShinsaInfo.ErrJoken)
            Next

            Try

                '審査内容文字列変換(項番)
                'REV_003↓
                '_shinsaKekkaList(0).ShinsaInfoList = ConvertShinsaNaiyo(_db, _shinsaKekkaList(0), New List(Of TargetData), _itemInfoList)
                _shinsaKekkaList(0).ShinsaInfoList = ConvertShinsaNaiyo(_db, _shinsaKekkaList(0), New List(Of TargetData), _itemInfoList, hasPrevData)
                'REV_003↑

                '審査実行
                ExecuteShinsaSingle(_db, Nothing, _shinsaKekkaList(0).ShinsaInfoList)
                ret = True

            Catch ex As Exception
                ret = False
            End Try

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 審査論理データ種別が調査票かどうかを判定する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsChosahyo() As Boolean
        Return If(_shinsaDataType = ComConst.審査論理データ種別.調査票, True, False)
    End Function

    ''' <summary>
    ''' 審査論理データ種別が個別結果表かどうかを判定する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsKobetsukekkahyo() As Boolean
        Return If(_shinsaDataType = ComConst.審査論理データ種別.個別結果表, True, False)
    End Function

    ''' <summary>
    ''' 項目情報を取得する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetItemInfo(ByVal pDb As DBAccess) As List(Of ItemInfo)
        Dim ret As List(Of ItemInfo) = Nothing
        Dim dt As DataTable = Nothing

        Try
            ret = New List(Of ItemInfo)

            '調査票の場合
            If IsChosahyo() Then
                dt = DAOOther.GetChosahyoItemMaster(pDb, _chosaKubun, _chosaYear)
            Else
                ' REV_002↓
                'dt = DAOOther.GetKobetsuKekkahyoItemMaster(pDb, _chosaKubun)
                dt = DAOOther.GetKobetsuKekkahyoItemMaster(pDb, _chosaKubun, _versionKbn)
                ' REV_002↑
            End If

            For Each dr As DataRow In dt.Rows
                ret.Add(New ItemInfo With {.ItemNo = dr("項目番号").ToString,
                                           .Type = dr("型区分").ToString,
                                           .IsVariable = If(IsChosahyo(), CBool(dr("可変区分")), False),
                                           .VariableMax = If(IsChosahyo(), CInt(dr("可変最大数")), 0)})
            Next

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 審査情報を取得する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetShinsaInfo(ByVal pDb As DBAccess) As List(Of ShinsaInfo)
        Dim ret As List(Of ShinsaInfo) = Nothing
        Dim dt As DataTable = Nothing

        Try
            ret = New List(Of ShinsaInfo)

            If IsChosahyo() Then
                If _shinsaRonriType = ComConst.審査論理種別.基本チェック Then
                    dt = DAOOther.GetChosahyoShinsaRonri(pDb, ComConst.エラーチェック種別.enm.基本)
                ElseIf _shinsaRonriType = ComConst.審査論理種別.追加チェック Then
                    dt = DAOOther.GetChosahyoShinsaRonri(pDb, ComConst.エラーチェック種別.enm.追加)
                End If
            Else
                If _shinsaRonriType = ComConst.審査論理種別.基本チェック Then
                    ' REV_002↓
                    'dt = DAOOther.GetKobetsuKekkahyoShinsaRonri(pDb, ComConst.エラーチェック種別.enm.基本)
                    dt = DAOOther.GetKobetsuKekkahyoShinsaRonri(pDb, _versionKbn, ComConst.エラーチェック種別.enm.基本)
                    ' REV_002↓
                ElseIf _shinsaRonriType = ComConst.審査論理種別.追加チェック Then
                    ' REV_002↓
                    'dt = DAOOther.GetKobetsuKekkahyoShinsaRonri(pDb, ComConst.エラーチェック種別.enm.追加)
                    dt = DAOOther.GetKobetsuKekkahyoShinsaRonri(pDb, _versionKbn, ComConst.エラーチェック種別.enm.追加)
                    ' REV_002↓
                End If
            End If

            For Each dr As DataRow In dt.Rows
                ret.Add(New ShinsaInfo With {.ErrSign = dr("エラーサイン").ToString,
                                                .CheckItemName = dr("チェック項目名").ToString,
                                                .ErrNaiyo = dr("エラー内容").ToString,
                                                .ErrKubun = dr("エラー区分").ToString,
                                                .ErrJoken = dr("エラーとなる条件").ToString,
                                                .IsRepeat = If(IsChosahyo(), If(dr("繰り返し").ToString = C_RepeatStr, True, False), False)})
            Next

        Catch ex As Exception
            Throw ex
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
    ''' 調査票審査情報を取得する
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pPKey"></param>
    ''' <param name="pItemInfoList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChosahyo(ByVal pDb As DBAccess, ByVal pPKey As DAOChosahyo.PrimaryKey, ByVal pItemInfoList As List(Of ItemInfo)) As List(Of TargetData)
        Dim ret As List(Of TargetData) = Nothing
        Dim dicChosahyo As Dictionary(Of String, DataTable) = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim chosahyoItemList As List(Of ItemInfo) = Nothing

        Try
            ret = New List(Of TargetData)
            '可変項目以外の調査票項目リストを取得
            chosahyoItemList = (From n In pItemInfoList Where n.IsVariable = False).ToList

            '調査票テーブルリスト取得
            dicChosahyo = DAOChosahyo.GetChosahyoTable(pDb, pPKey)
            For Each kv As KeyValuePair(Of String, DataTable) In dicChosahyo
                Dim tableName As String = kv.Key
                dt = kv.Value
                If Not tableName.Contains("＿可変") Then
                    If dt.Rows.Count > 0 Then
                        dr = dt.Rows(0)
                        For i As Integer = 0 To chosahyoItemList.Count - 1
                            '対象テーブルに項目番号が存在する場合
                            If dt.Columns.Contains(chosahyoItemList(i).ItemNo) Then
                                '項目番号を追加する
                                '↓REV_004
                                ret.Add(New TargetData With {.ItemNo = chosahyoItemList(i).ItemNo, .Value = If(IsDBNull(dr(chosahyoItemList(i).ItemNo)), Nothing, dr(chosahyoItemList(i).ItemNo))})
                                '↑REV_004
                            End If
                        Next
                    End If
                Else
                    For Each dr In dt.Rows
                        '可変項目の項目番号を『項目番号 + '_' + 明細番号』の形式で追加する
                        ret.Add(New TargetData With {.ItemNo = dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & dr("明細番号").ToString,
                                                        .Value = If(IsDBNull(dr("値")), Nothing, dr("値")),
                                                        .MeisaiNo = CInt(dr("明細番号"))})
                    Next
                End If
            Next

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 個別結果表審査情報を取得する
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pPKey"></param>
    ''' <param name="pItemInfoList"></param>
    ''' <param name="isPrev">REV_003</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetKobetsukekkahyo(ByVal pDb As DBAccess, ByVal pPKey As DAOKobetsuKekkahyo.PrimaryKey, ByVal pItemInfoList As List(Of ItemInfo), Optional isPrev As Boolean = False) As List(Of TargetData)
        Dim ret As List(Of TargetData) = Nothing
        Dim dicKobetsu As Dictionary(Of String, DataTable) = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim kekkahyoItemList As List(Of ItemInfo) = Nothing

        Try
            ret = New List(Of TargetData)
            '個別結果表項目リストを取得
            kekkahyoItemList = (From n In pItemInfoList Where n.IsVariable = False).ToList

            '個別結果表テーブルリスト取得
            'REV_003↓
            If isPrev Then
                pPKey.chosaNen = (CInt(pPKey.chosaNen) - 1).ToString
            End If
            'REV_003↑
            dicKobetsu = DAOKobetsuKekkahyo.GetTable(pDb, pPKey)
            For Each kv As KeyValuePair(Of String, DataTable) In dicKobetsu
                dt = kv.Value
                If dt.Rows.Count > 0 Then
                    dr = dt.Rows(0)
                    For i As Integer = 0 To kekkahyoItemList.Count - 1
                        '対象テーブルに項目番号が存在する場合
                        If dt.Columns.Contains(kekkahyoItemList(i).ItemNo) Then
                            '項目番号を追加する
                            'REV_003↓
                            'ret.Add(New TargetData With {.ItemNo = kekkahyoItemList(i).ItemNo, .Value = dr(kekkahyoItemList(i).ItemNo)})
                            ret.Add(New TargetData With {.ItemNo = If(isPrev, "前", "") & kekkahyoItemList(i).ItemNo, .Value = dr(kekkahyoItemList(i).ItemNo)})
                            'REV_003↑
                        End If
                    Next
                End If
            Next

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 審査内容文字列を変換する(項目番号以外の置換)
    ''' </summary>
    ''' <param name="pShinsaNaiyo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConvertShinsaNaiyo(ByVal pShinsaNaiyo As String) As String
        Dim ret As String = Nothing

        Try
            ret = pShinsaNaiyo

            '半角スペース、全角スペースを削除
            ret = ret.Replace(" "c, "").Replace("　"c, "")
            '改行を削除
            ret = ret.Replace(vbCr, "").Replace(vbLf, "")
            '全角角括弧『［』『］』 ⇒ 半角角括弧『[』『]』
            ret = ret.Replace("［"c, "["c).Replace("］"c, "]"c)
            '『]＝空白』 ⇒ 『]@ IS NULL』
            ret = ret.Replace("]＝空白", "]@ IS NULL")
            '『]≠空白』 ⇒ 『]@ IS NOT NULL』
            ret = ret.Replace("]≠空白", "]@ IS NOT NULL")
            '全角不等号『＜』『＞』 ⇒ 半角不等号『<』『>』
            ret = ret.Replace("＜", "<").Replace("＞", ">")
            '全角等号付き不等号『≦』『≧』 ⇒ 『<=』『>=』
            ret = ret.Replace("≦", "<=").Replace("≧", ">=")
            '全角丸括弧『（』『）』 ⇒ 半角丸括弧『(』『)』
            ret = ret.Replace("（", "(").Replace("）", ")")
            '全角等号『＝』 ⇒ 半角等号『=』
            ret = ret.Replace("＝", "=")
            '全角等号否定『≠』 ⇒ 半角山括弧『<>』
            ret = ret.Replace("≠", "<>")
            '全角記号『＋－÷×』 ⇒ 半角記号『+-/*』
            '『－』は後ろの数値がマイナスの時に『--』となる(コメント扱いになる)のを防ぐため後ろに半角スペースを入れる
            ret = ret.Replace("＋", "+").Replace("－", "- ").Replace("-", "- ").Replace("÷", "/").Replace("×", "*")
            '全角記号『／』『＊』 ⇒ 半角記号『/』『*』
            ret = ret.Replace("／", "/").Replace("＊", "*")
            '全角単一引用符『’』 ⇒ 半角単一引用符『'』
            ret = ret.Replace("’", "'")
            '小文字を大文字に変換
            ret = ret.ToUpper()
            '『AND』 ⇒ 『 AND 』(前後に半角スペースを付加)
            ret = ret.Replace("AND", " AND ")
            '『OR』 ⇒ 『 OR 』(前後に半角スペースを付加)
            ret = ret.Replace("OR", " OR ")
            '『NOTIN』 ⇒ 『 NOTIN』(前に半角スペースを付加)
            ret = ret.Replace("NOTIN", " NOTIN")
            '『IN』 ⇒ 『 IN』(前に半角スペースを付加)
            ret = ret.Replace("IN", " IN")

            '項目番号の置換はここではやらない
            '(可変データの数により審査論理が増減するため)

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 審査内容文字列を変換する(項目番号の置換)
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pShinsaKekka"></param>
    ''' <param name="pTargetDataList"></param>
    ''' <param name="pItemInfoList"></param>
    ''' <param name="hasPrevData">REV_003</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConvertShinsaNaiyo(ByVal pDb As DBAccess,
                                        ByVal pShinsaKekka As ShinsaKekka,
                                        ByVal pTargetDataList As List(Of TargetData),
                                        ByVal pItemInfoList As List(Of ItemInfo),
                                        hasPrevData As Boolean) As List(Of ShinsaInfo)
        Dim ret As List(Of ShinsaInfo) = Nothing
        Dim shinsaInfoList As List(Of ShinsaInfo) = Nothing
        Dim jimusho As String
        Dim rochinTanka As Integer

        Try

            ret = New List(Of ShinsaInfo)
            shinsaInfoList = New List(Of ShinsaInfo)(pShinsaKekka.ShinsaInfoList)

            '-----------------------------------------------------------------------------------------------------
            '労賃単価の置換
            '-----------------------------------------------------------------------------------------------------
            If _isErrCheck Then
                rochinTanka = 0
            Else
                '入力・修正画面からの審査の場合
                If String.IsNullOrEmpty(pShinsaKekka.CensusNo) Then
                    jimusho = _jimusho
                Else
                    '事務所番号取得
                    jimusho = ComUtil.ConvJimusyoNo(ComUtil.GetTodofuken(pShinsaKekka.CensusNo))
                End If
                '労賃単価取得
                rochinTanka = GetRochinTanka(pDb, jimusho)
            End If
            For Each ShinsaInfo In shinsaInfoList
                '『【労賃単価】』を値に置換
                ShinsaInfo.ErrJokenConvValue = ShinsaInfo.ErrJokenConvValue.Replace("【労賃単価】", rochinTanka.ToString)
            Next

            '検索処理高速化のため項目番号と値のリストを作成
            Dim itemNoValueDict As New Dictionary(Of String, Object)
            For Each targetData In pTargetDataList
                itemNoValueDict.Add(targetData.ItemNo, targetData.Value)
            Next

            For Each ShinsaInfo In shinsaInfoList
                'REV_003↓前付項番を使用していて前年データがない審査はスキップする
                If ShinsaInfo.ErrJoken.Contains("[前") AndAlso Not hasPrevData Then
                    If _progressDialog IsNot Nothing Then
                        '進捗を進める
                        _progressDialog.AddValue = 1
                    End If
                    Continue For
                End If
                'REV_003↑
                '繰り返しの場合
                If ShinsaInfo.IsRepeat Then
                    'エラー条件文字列から可変項目番号リストを取得する
                    Dim variableItemNoList As List(Of String) = GetVariableItemNoList(pItemInfoList, ShinsaInfo.ErrJokenConvValue)
                    Dim variableCount As Integer = GetVariableCount(ShinsaInfo.ErrJoken, pTargetDataList)
                    '繰り返しありの審査を明細の数だけ増加
                    For i As Integer = 1 To variableCount
                        Dim tmpErrJokenConv As String = ShinsaInfo.ErrJokenConvValue
                        For Each variableItemNo In variableItemNoList
                            '審査論理の可変項目番号を『項目番号 + '_' + 明細番号』に置換する
                            tmpErrJokenConv = tmpErrJokenConv.Replace(variableItemNo, variableItemNo & ComConst.ITEM_NO_DELIMITER & i)
                        Next

                        Dim sInfo As ShinsaInfo = New ShinsaInfo With {.ErrSign = ShinsaInfo.ErrSign,
                                                                        .ErrNaiyo = ShinsaInfo.ErrNaiyo,
                                                                        .ErrKubun = ShinsaInfo.ErrKubun,
                                                                        .ErrJoken = ShinsaInfo.ErrJoken,
                                                                        .ErrJokenConvValue = tmpErrJokenConv,
                                                                        .IsRepeat = ShinsaInfo.IsRepeat,
                                                                        .MeisaiNo = i,
                                                                        .Result = ShinsaInfo.Result}

                        sInfo.ErrJokenConv = sInfo.ErrJokenConvValue

                        'エラー条件文字列の集計関連文字列を値に置換する
                        Me.ConvertFormulaCountAndSum(pDb, sInfo, pItemInfoList, itemNoValueDict, pTargetDataList)

                        'エラー条件文字列の項目番号を値に置換する
                        sInfo.ErrJokenConvValue = ConvertJokenItemNoToValue(itemNoValueDict, sInfo.ErrJokenConvValue)

                        ret.Add(sInfo)
                    Next
                Else
                    ShinsaInfo.ErrJokenConv = ShinsaInfo.ErrJokenConvValue

                    'エラー条件文字列の集計関連文字列を値に置換する
                    Me.ConvertFormulaCountAndSum(pDb, ShinsaInfo, pItemInfoList, itemNoValueDict, pTargetDataList)

                    'エラー条件文字列の項目番号を値に置換する
                    ShinsaInfo.ErrJokenConvValue = ConvertJokenItemNoToValue(itemNoValueDict, ShinsaInfo.ErrJokenConvValue)

                    ret.Add(ShinsaInfo)
                End If

                If _progressDialog IsNot Nothing Then
                    '進捗を進める
                    _progressDialog.AddValue = 1
                End If
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' エラー条件文字列の集計関連文字列を値に置換する
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pShinsaInfo"></param>
    ''' <param name="pItemInfoList"></param>
    ''' <param name="pItemNoValueDict"></param>
    ''' <param name="pTargetDataList"></param>
    ''' <remarks></remarks>
    Private Sub ConvertFormulaCountAndSum(ByVal pDb As DBAccess,
                                          ByRef pShinsaInfo As ShinsaInfo,
                                          ByVal pItemInfoList As List(Of ItemInfo),
                                          ByVal pItemNoValueDict As Dictionary(Of String, Object),
                                          ByVal pTargetDataList As List(Of TargetData))

        Dim matchList As MatchCollection = Nothing

        '------------------------------------------
        '項目番号->値の置換(条件なしのCOUNT、SUM)
        '------------------------------------------
        '『COUNT([』『])』で囲まれた文字列検索する
        matchList = Regex.Matches(pShinsaInfo.ErrJokenConvValue, C_SearchCount)
        For Each mat As Match In matchList
            '『COUNT([010101])』から項目番号『010101』を抽出する
            Dim itemNo As String = Regex.Match(mat.Value, C_SearchKakko).Value.Trim("["c, "]"c)
            '可変項目の値の総数を取得する
            Dim countValue As Decimal = (From n In pTargetDataList Where n.ItemNo.Contains(itemNo) Select CDec(n.Value)).Count
            '項目番号を総数に置換する
            pShinsaInfo.ErrJokenConvValue = pShinsaInfo.ErrJokenConvValue.Replace(mat.Value, countValue.ToString)
        Next
        '『SUM([』『])』で囲まれた文字列検索する
        matchList = Regex.Matches(pShinsaInfo.ErrJokenConvValue, C_SearchSum)
        For Each mat As Match In matchList
            '『SUM([010101])』から項目番号『010101』を抽出する
            Dim itemNo As String = Regex.Match(mat.Value, C_SearchKakko).Value.Trim("["c, "]"c)
            '可変項目の値の総和を取得する
            Dim sumValue As Decimal = (From n In pTargetDataList Where n.ItemNo.Contains(itemNo) Select CDec(n.Value)).Sum
            '項目番号を総和に置換する
            pShinsaInfo.ErrJokenConvValue = pShinsaInfo.ErrJokenConvValue.Replace(mat.Value, sumValue.ToString)
        Next

        '------------------------------------------
        '項目番号->値の置換(条件ありのCOUNT、SUM)
        '------------------------------------------
        '『COUNT(』『)』で囲まれた文字列検索する
        matchList = Regex.Matches(pShinsaInfo.ErrJokenConvValue, C_SearchCountif)
        For Each mat As Match In matchList
            '『COUNT(』『)』で囲まれた文字列を合計値に置換する
            pShinsaInfo.ErrJokenConvValue = ConvertCountifSumifToValue(pDb, pItemInfoList, pItemNoValueDict, pTargetDataList, pShinsaInfo.ErrJokenConvValue, mat.Value)
        Next
        '『SUM(』『)』で囲まれた文字列検索する
        matchList = Regex.Matches(pShinsaInfo.ErrJokenConvValue, C_SearchSumif)
        For Each mat As Match In matchList
            '『SUM(』『)』で囲まれた文字列を合計値に置換する
            pShinsaInfo.ErrJokenConvValue = ConvertCountifSumifToValue(pDb, pItemInfoList, pItemNoValueDict, pTargetDataList, pShinsaInfo.ErrJokenConvValue, mat.Value)
        Next

    End Sub

    ''' <summary>
    ''' エラー条件文字列から可変項目番号リストを取得する
    ''' </summary>
    ''' <param name="pItemInfoList"></param>
    ''' <param name="pJoken"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetVariableItemNoList(ByVal pItemInfoList As List(Of ItemInfo), ByVal pJoken As String) As List(Of String)
        Dim ret As New List(Of String)
        Dim matchList As MatchCollection = Nothing

        '半角角括弧『［』『］』で囲まれた文字列を『[010101_1]』に置換してデータの数だけ複数行作成する。
        matchList = Regex.Matches(pJoken, C_SearchKakko)
        For Each mat As Match In matchList
            '『[010101]』から項目番号『010101』を抽出する
            Dim itemNo As String = mat.Value.Trim("["c, "]"c)
            '抽出した項目番号が可変項目の場合
            Dim index As Integer = pItemInfoList.BinarySearch(New ItemInfo With {.ItemNo = itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(0)})
            If pItemInfoList(index).IsVariable Then
                If Not ret.Contains(itemNo) Then
                    ret.Add(itemNo)
                End If
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' エラー条件文字列のCOUNTIF、SUMIFを値に置換する
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pItemInfoList"></param>
    ''' <param name="pItemNoValueDict"></param>
    ''' <param name="pTargetDataList"></param>
    ''' <param name="pErrJoken"></param>
    ''' <param name="pErrJokenCountifSumif"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConvertCountifSumifToValue(ByVal pDb As DBAccess,
                                                ByVal pItemInfoList As List(Of ItemInfo),
                                                ByVal pItemNoValueDict As Dictionary(Of String, Object),
                                                ByVal pTargetDataList As List(Of TargetData),
                                                ByVal pErrJoken As String,
                                                ByVal pErrJokenCountifSumif As String) As String
        Dim ret As String = Nothing

        'エラー条件文字列から可変項目番号リストを取得する
        Dim variableItemNoList As List(Of String) = GetVariableItemNoList(pItemInfoList, pErrJokenCountifSumif)
        '増加する条件リスト
        Dim jokenList As New List(Of String)
        Dim variableCount As Integer = GetVariableCount(pErrJokenCountifSumif, pTargetDataList)
        For i As Integer = 1 To variableCount
            Dim tmpErrJokenConv As String = pErrJokenCountifSumif
            For Each variableItemNo In variableItemNoList
                '審査論理の可変項目番号を『項目番号 + '_' + 明細番号』に置換する
                tmpErrJokenConv = tmpErrJokenConv.Replace(variableItemNo, variableItemNo & ComConst.ITEM_NO_DELIMITER & i)
            Next
            jokenList.Add(tmpErrJokenConv)
        Next

        'エラー条件文字列の項目番号を値に置換する
        Dim JokenConvList As New List(Of String)
        For Each joken In jokenList
            JokenConvList.Add(ConvertJokenItemNoToValue(pItemNoValueDict, joken).Replace("空白", "NULL"))
        Next
        'COUNTIF、SUMIFの値を取得する
        Dim value As Decimal = GetCountifSumifValue(pDb, JokenConvList)
        'エラー条件文字列の項目番号を合計値に置換する
        ret = pErrJoken.Replace(pErrJokenCountifSumif, value.ToString)

        Return ret
    End Function

    ''' <summary>
    ''' エラー条件文字列の項目番号を値に置換する
    ''' </summary>
    ''' <param name="pItemNoValueDict"></param>
    ''' <param name="pJoken"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConvertJokenItemNoToValue(ByVal pItemNoValueDict As Dictionary(Of String, Object), ByVal pErrJoken As String) As String
        Dim ret As String = pErrJoken
        Dim matchList As MatchCollection = Nothing

        '括弧付き項目番号『[010101]』『[010101_1]』を検索
        matchList = Regex.Matches(ret, C_SearchKakko)
        '重複しない括弧付き項目番号リストを作成
        Dim itemNoKakkoList As String() = (From m In matchList.Cast(Of Match)() Select m.Value).Distinct().ToArray()
        For Each itemNoKakko As String In itemNoKakkoList
            '『[010101]』『[010101_1]』から項目番号『010101』『010101_1』を抽出する
            Dim itemNo As String = itemNoKakko.Trim("["c, "]"c)
            Dim itemType As String = GetItemType(itemNo)
            'データが存在する場合
            If pItemNoValueDict.ContainsKey(itemNo) Then
                Dim value As Object = pItemNoValueDict(itemNo)
                '@付き項目番号『[010101]@』『[010101_1]@』を値に置換
                If itemType = ComConst.型区分.数値型 Then
                    ret = ret.Replace(itemNoKakko & "@", If(IsDBNull(value) OrElse value Is Nothing, "NULL", Me.CnvIntToDec(value.ToString)))
                Else
                    ret = ret.Replace(itemNoKakko & "@", If(IsDBNull(value) OrElse value Is Nothing, "NULL", "'" & value.ToString & "'"))
                End If
                '項目番号『[010101]』『[010101_1]』を値に置換
                If itemType = ComConst.型区分.数値型 Then
                    ret = ret.Replace(itemNoKakko, If(IsDBNull(value) OrElse value Is Nothing, "0", Me.CnvIntToDec(value.ToString)))
                Else
                    ret = ret.Replace(itemNoKakko, If(IsDBNull(value) OrElse value Is Nothing, "''", "'" & value.ToString & "'"))
                End If
            Else
                '@付き項目番号『[010101]@』『[010101_1]@』を『NULL』に置換
                ret = ret.Replace(itemNoKakko & "@", "NULL")
                '項目番号『[010101]』『[010101_1]』を『0』or『''』に置換
                If itemType = ComConst.型区分.数値型 Then
                    ret = ret.Replace(itemNoKakko, "0")
                Else
                    'エラーチェック時は任意の文字をセットする(文字列と数値の比較『''<0』等がエラーにならないため)
                    ret = ret.Replace(itemNoKakko, If(_isErrCheck, "'a'", "''"))
                End If
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' COUNTIF、SUMIFの値を取得する
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pErrJokenList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ' REV_001↓
    'Private Function GetCountifSumifValue(ByVal pDb As DBAccess, ByVal pErrJokenList As List(Of String)) As Integer
    '    Dim ret As Integer
    Private Function GetCountifSumifValue(ByVal pDb As DBAccess, ByVal pErrJokenList As List(Of String)) As Decimal
        Dim ret As Decimal
        ' REV_001↑
        Dim dt As DataTable = Nothing
        Dim sb As System.Text.StringBuilder = Nothing

        Try
            'C_ShinsaUnitCountの数だけまとめて審査(SELECT)する
            For i As Integer = 0 To CInt(Math.Ceiling(pErrJokenList.Count / C_ShinsaUnitCount)) - 1
                Dim jokenList As List(Of String) = pErrJokenList.Skip(C_ShinsaUnitCount * i).Take(C_ShinsaUnitCount).ToList

                sb = New System.Text.StringBuilder

                '0除算をエラーとしない
                sb.AppendLine("SET ANSI_WARNINGS OFF")
                sb.AppendLine("SELECT ")
                For j As Integer = 0 To jokenList.Count - 1
                    If j > 0 Then
                        sb.Append(" ,")
                    End If
                    sb.Append("ISNULL(" & jokenList(j) & ",0) AS [" & j + 1 & "] ")
                Next
                dt = pDb.GetDataTable(sb.ToString)
                For col As Integer = 0 To dt.Columns.Count - 1
                    ' REV_001↓
                    'ret += CInt(dt.Rows(0)(col))
                    ret += CDec(dt.Rows(0)(col))
                    ' REV_001↑
                Next
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 審査を実行して審査結果をセットする
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pShinsaInfoList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteShinsa(ByVal pDb As DBAccess, ByRef pShinsaInfoList As List(Of ShinsaInfo))
        Dim dt As DataTable = Nothing
        Dim sb As System.Text.StringBuilder = Nothing
        Dim errCnt As Integer

        Try

            'C_ShinsaUnitCountの数だけまとめて審査(SELECT)する
            For i As Integer = 0 To CInt(Math.Ceiling(pShinsaInfoList.Count / C_ShinsaUnitCount)) - 1
                Dim shinsaInfoList As List(Of ShinsaInfo) = pShinsaInfoList.Skip(C_ShinsaUnitCount * i).Take(C_ShinsaUnitCount).ToList

                sb = New System.Text.StringBuilder

                '0除算をエラーとしない
                sb.AppendLine("SET ANSI_WARNINGS OFF")
                sb.AppendLine("SELECT ")
                For j As Integer = 0 To shinsaInfoList.Count - 1
                    If j > 0 Then
                        sb.Append(" ,")
                    End If
                    sb.Append("CASE WHEN " & shinsaInfoList(j).ErrJokenConvValue & " THEN 'FALSE' ELSE 'TRUE' END AS [")
                    If shinsaInfoList(j).IsRepeat Then
                        '繰り返しの場合はエラーサインに『アンダースコア + 行数』を付加する
                        sb.Append(shinsaInfoList(j).ErrSign & "_" & shinsaInfoList(j).MeisaiNo)
                    Else
                        sb.Append(shinsaInfoList(j).ErrSign)
                    End If
                    sb.AppendLine("] ")
                Next
                dt = pDb.GetDataTable(sb.ToString)
                For Each dc As DataColumn In dt.Columns
                    If CBool(dt(0).Item(dc.ColumnName)) = False Then
                        errCnt += 1
                    End If
                Next

                For Each ShinsaInfo In shinsaInfoList
                    If ShinsaInfo.IsRepeat Then
                        ShinsaInfo.Result = CBool(dt.Rows(0).Item(ShinsaInfo.ErrSign & "_" & ShinsaInfo.MeisaiNo))
                    Else
                        ShinsaInfo.Result = CBool(dt.Rows(0).Item(ShinsaInfo.ErrSign))
                    End If
                Next

                '審査エラー数が最大エラー件数に達した場合
                If _errMaxCount > 0 AndAlso errCnt >= _errMaxCount Then
                    Exit For
                End If
            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 審査を実行して審査結果をセットする
    ''' </summary>
    ''' <param name="pDb"></param>
    ''' <param name="pCensusNo"></param>
    ''' <param name="pShinsaInfoList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Sub ExecuteShinsaSingle(ByVal pDb As DBAccess, ByVal pCensusNo As String, ByRef pShinsaInfoList As List(Of ShinsaInfo))
        Dim dt As DataTable = Nothing
        Dim sb As System.Text.StringBuilder = Nothing

        Try

            For Each ShinsaInfo In pShinsaInfoList
                sb = New System.Text.StringBuilder

                'SQL文の設定
                sb.AppendLine("SET ANSI_WARNINGS OFF")
                sb.Append("SELECT CASE WHEN " & ShinsaInfo.ErrJokenConvValue & " THEN 'FALSE' ELSE 'TRUE' END AS [")
                If ShinsaInfo.IsRepeat Then
                    '繰り返しの場合はエラーサインに『アンダースコア + 行数』を付加する
                    sb.Append(ShinsaInfo.ErrSign & "_" & ShinsaInfo.MeisaiNo)
                Else
                    sb.Append(ShinsaInfo.ErrSign)
                End If
                sb.AppendLine("] ")

                Try
                    dt = pDb.GetDataTable(sb.ToString)
                    If ShinsaInfo.IsRepeat Then
                        ShinsaInfo.Result = CBool(dt.Rows(0).Item(ShinsaInfo.ErrSign & "_" & ShinsaInfo.MeisaiNo))
                    Else
                        ShinsaInfo.Result = CBool(dt.Rows(0).Item(ShinsaInfo.ErrSign))
                    End If
                Catch ex As Exception
                    '審査エラー
                    Throw New ShinsaException(ex, pCensusNo, ShinsaInfo.ErrSign, ShinsaInfo.ErrJoken, ShinsaInfo.ErrJokenConvValue)
                End Try
            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 審査論理に使用されている可変項目で一番大きい行数を取得する
    ''' </summary>
    ''' <param name="pNaiyo"></param>
    ''' <param name="pTargetDataList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetVariableCount(ByVal pNaiyo As String, ByVal pTargetDataList As List(Of TargetData)) As Integer
        Dim ret As Integer
        Dim countList As New List(Of Integer)
        Dim matchList As MatchCollection = Nothing

        If _isErrCheck Then
            ret = 1
        Else
            '審査論理に使用されている項目番号を抽出する
            matchList = Regex.Matches(pNaiyo, C_SearchKakko)
            For Each mat As Match In matchList
                '『[010101]』から項目番号『010101』を抽出する
                Dim itemNo As String = mat.Value.Trim("["c, "]"c)
                '可変項目の場合
                If (From n In _itemInfoList Where n.ItemNo = itemNo And n.IsVariable).Any Then
                    If (From n In pTargetDataList Where n.ItemNo.Contains(itemNo)).Any Then
                        '値リストの最大行数をセットする
                        countList.Add((From n In pTargetDataList Where n.ItemNo.Contains(itemNo) Select n.MeisaiNo).Max)
                    Else
                        countList.Add(0)
                    End If
                End If
            Next
            ret = countList.Max
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 項目番号の型を取得する
    ''' </summary>
    ''' <param name="pItemNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetItemType(ByVal pItemNo As String) As String
        'pItemNo：可変項目番号も含まれる([010101]、[010101_1])
        'ItemNo ：可変項目番号は含まない([010101])
        Dim index As Integer = _itemInfoList.BinarySearch(New ItemInfo With {.ItemNo = pItemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(0)})
        Return _itemInfoList(index).Type
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
    Public Function DeepCopy(Of T)(ByVal target As T) As T
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
