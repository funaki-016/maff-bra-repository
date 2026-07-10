'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_000   | 2023.01.13 |大興電子通信        | 要件No.5 新規作成
'//            |            |                    |
'//*************************************************************************************************
Imports Maff.BRA.HyohonConst

''' <summary>
''' 標本管理の共通機能クラス
''' </summary>
Public Class HyohonUtil

    ''' <summary>
    ''' ComboboxにKeyValueListを設定
    ''' </summary>
    ''' <param name="cbo"></param>
    ''' <param name="lst"></param>
    Public Shared Sub SetCombobox(Of TKey)(ByRef cbo As ComboBox, lst As List(Of KeyValuePair(Of TKey, String)))
        cbo.ValueMember = "Key"
        cbo.DisplayMember = "Value"
        cbo.DataSource = lst
    End Sub

    ''' <summary>
    ''' 経営形態区分リストを取得
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function Get経営形態区分リスト() As List(Of KeyValuePair(Of Integer, String))
        Return 経営形態区分辞書.OrderBy(Function(x) x.Key).Select(Function(x) New KeyValuePair(Of Integer, String)(x.Key, x.Value.名称)).ToList()
    End Function

    ''' <summary>
    ''' 生産費区分リストを取得
    ''' </summary>
    ''' <param name="経営形態区分"></param>
    ''' <returns></returns>
    Public Shared Function Get生産費区分リスト(経営形態区分 As 経営形態区分) As List(Of KeyValuePair(Of Integer, String))
        Return 生産費区分リスト辞書(経営形態区分)
    End Function

    ''' <summary>
    ''' Comboboxの選択値を取得
    ''' </summary>
    ''' <param name="cbo"></param>
    ''' <returns></returns>
    Public Shared Function GetComboValue(cbo As ComboBox) As Integer
        Return GetComboValue(cbo, 識別対象外値)
    End Function
    ''' <summary>
    ''' Comboboxの選択値を取得
    ''' </summary>
    ''' <param name="cbo"></param>
    ''' <param name="emptyValue"></param>
    ''' <returns></returns>
    Public Shared Function GetComboValue(Of T)(cbo As ComboBox, emptyValue As T) As T
        Return CType(If(cbo.SelectedValue Is Nothing OrElse cbo.SelectedValue Is DBNull.Value, emptyValue, cbo.SelectedValue), T)
    End Function

    ''' <summary>
    ''' Comboboxが空か否かを判定
    ''' </summary>
    ''' <param name="cbo"></param>
    ''' <returns></returns>
    Public Shared Function IsEmpty(cbo As ComboBox) As Boolean
        Return IsEmpty(cbo, 識別対象外値)
    End Function
    ''' <summary>
    ''' Comboboxが空か否かを判定
    ''' </summary>
    ''' <param name="cbo"></param>
    ''' <param name="emptyValue"></param>
    ''' <returns></returns>
    Public Shared Function IsEmpty(Of T)(cbo As ComboBox, emptyValue As T) As Boolean
        Return GetComboValue(cbo, emptyValue).Equals(emptyValue)
    End Function

    ''' <summary>
    ''' 列情報取得SQL
    ''' </summary>
    Private Const SQL_SELECT_COL_INFO =
        "SELECT" &
        " C.COLUMN_ID," &
        " C.NAME," &
        " C.IS_NULLABLE," &
        " UPPER(T.NAME) AS TYPE," &
        " CASE UPPER(T.NAME)" &
        " WHEN 'NVARCHAR' THEN C.MAX_LENGTH / 2" &
        " WHEN 'NUMERIC' THEN C.PRECISION + IIF(C.SCALE > 0, 1, 0)" &
        " ELSE C.MAX_LENGTH END AS MAX_LENGTH," &
        " C.PRECISION - C.SCALE AS PRECISION," &
        " C.SCALE" &
        " FROM SYS.OBJECTS O" &
        " JOIN SYS.COLUMNS C" &
        " ON C.OBJECT_ID = O.OBJECT_ID" &
        " JOIN SYS.TYPES T" &
        " ON T.USER_TYPE_ID = C.USER_TYPE_ID" &
        " WHERE O.NAME = @テーブル名" &
        " AND C.NAME LIKE 'NO%'" &
        " ORDER BY C.COLUMN_ID"
    ''' <summary>
    ''' 標本リストyyyyの列情報を取得
    ''' </summary>
    ''' <param name="censusNen"></param>
    ''' <returns></returns>
    Public Shared Function GetColInfo(censusNen As Integer) As (COLUMN_ID As Integer, NAME As String, IS_NULLABLE As Boolean, TYPE As String, MAX_LENGTH As Integer, PRECISION As Integer, SCALE As Integer)()
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Dim para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@テーブル名", SqlDbType.VarChar, String.Format("標本リスト{0}", censusNen)))
            Dim dt = db.GetDataTable(SQL_SELECT_COL_INFO, para)
            Dim colInfo As (COLUMN_ID As Integer, NAME As String, IS_NULLABLE As Boolean, TYPE As String, MAX_LENGTH As Integer, PRECISION As Integer, SCALE As Integer)() =
                dt.Rows.Cast(Of DataRow).Select(Function(x) (CInt(x("COLUMN_ID")), x("NAME").ToString(), CBool(x("IS_NULLABLE")), x("TYPE").ToString(), CInt(x("MAX_LENGTH")), CInt(x("PRECISION")), CInt(x("SCALE")))).ToArray()
            Return colInfo
        End Using
    End Function

    ''' <summary>
    ''' 標本リストの列番号を取得
    ''' </summary>
    ''' <param name="センサス実施年"></param>
    ''' <param name="標本リスト列情報キー"></param>
    ''' <returns></returns>
    Public Shared Function Get標本リスト列番号(センサス実施年 As Integer, 標本リスト列情報キー As 標本リスト列番号キー) As Integer
        Return 標本リスト列番号(センサス実施年)(標本リスト列情報キー)
    End Function

    ''' <summary>
    ''' 標本リスト比較列番号を取得
    ''' </summary>
    ''' <param name="センサス実施年"></param>
    ''' <param name="経営形態区分"></param>
    ''' <param name="営農類型区分"></param>
    ''' <param name="集落営農区分"></param>
    ''' <param name="生産費区分"></param>
    ''' <param name="田畑区分"></param>
    ''' <returns></returns>
    Public Shared Function Get標本リスト比較列番号(センサス実施年 As Integer, 経営形態区分 As 経営形態区分, 営農類型区分 As 営農類型区分, 集落営農区分 As 集落営農区分, 生産費区分 As 生産費区分, 田畑区分 As 田畑区分) As Integer
        Return 標本リスト比較列番号(センサス実施年)((経営形態区分, 営農類型区分, 集落営農区分, 生産費区分, 田畑区分))
    End Function

    ''' <summary>
    ''' リストから名称を取得
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="key"></param>
    ''' <returns></returns>
    Public Shared Function Getリスト名称(Of TKey)(list As List(Of KeyValuePair(Of TKey, String)), key As TKey) As String
        Return list.Where(Function(x) x.Key.Equals(key)).Select(Function(x) x.Value).FirstOrDefault
    End Function

    ''' <summary>
    ''' 営農類型・生産費区分（田畑区分）名称を取得
    ''' </summary>
    ''' <param name="keieiKeitai"></param>
    ''' <param name="einoRuikei"></param>
    ''' <param name="seisanhi"></param>
    ''' <param name="tahata"></param>
    ''' <returns></returns>
    Public Shared Function Get営農類型＿生産費区分名称(keieiKeitai As Integer, einoRuikei As Integer, seisanhi As Integer, tahata As Integer) As String
        Dim einoSeisanhiName = ""
        Select Case keieiKeitai
            Case HyohonConst.経営形態区分.個人経営体, HyohonConst.経営形態区分.法人経営体
                einoSeisanhiName = HyohonUtil.Getリスト名称(HyohonConst.営農類型区分リスト, einoRuikei)
            Case HyohonConst.経営形態区分.個別経営体, HyohonConst.経営形態区分.組織法人経営体
                einoSeisanhiName = HyohonUtil.Getリスト名称(HyohonUtil.Get生産費区分リスト(DirectCast(keieiKeitai, HyohonConst.経営形態区分)), seisanhi)
                Select Case tahata
                    Case HyohonConst.田畑区分.田
                        einoSeisanhiName &= "（うち、田）"
                    Case HyohonConst.田畑区分.畑
                        einoSeisanhiName &= "（うち、畑）"
                End Select
        End Select
        Return einoSeisanhiName
    End Function

    ''' <summary>
    ''' 一覧に行を追加
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <param name="values"></param>
    ''' <returns></returns>
    Public Shared Function AddDgvRow(dgv As DataGridView, values As Object()) As DataGridViewRow
        Dim row = New DataGridViewRow
        row.CreateCells(dgv)
        row.SetValues(values)
        row.Height = 20
        dgv.Rows.Add(row)
        Return row
    End Function

    ''' <summary>
    ''' 地方農政局等マスタ検索SQL
    ''' </summary>
    Private Const SQL_SELECT_地方農政局等マスタリスト =
        "SELECT * FROM 地方農政局等マスタ" _
        & " ORDER BY" _
        & " CASE WHEN 農政局 = @農政局＿全国 THEN 0 " _
        & "      WHEN 都道府県 = @都道府県＿全国＿農政局 THEN 99" _
        & "      ELSE 都道府県 END," _
        & " 農政局"
    ''' <summary>
    ''' 地方農政局等マスタリストを取得
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function Get地方農政局等マスタリスト() As List(Of KeyValuePair(Of (農政局 As Integer, 都道府県 As Integer), String))
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Dim para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@農政局＿全国", SqlDbType.Int, 農政局＿全国))
            para.Add(db.CreateParameter("@都道府県＿全国＿農政局", SqlDbType.Int, 都道府県＿全国＿農政局))
            Dim dt = db.GetDataTable(SQL_SELECT_地方農政局等マスタリスト, para)
            Dim list = New List(Of KeyValuePair(Of (Integer, Integer), String)) From {
                New KeyValuePair(Of (Integer, Integer), String)((識別対象外値, 識別対象外値), "")
            }
            list.AddRange(dt.Rows().Cast(Of DataRow).Select(Function(x) New KeyValuePair(Of (Integer, Integer), String)((CInt(x("農政局")), CInt(x("都道府県"))), x("名称").ToString())))
            Return list
        End Using
    End Function

    ''' <summary>
    ''' 標本リスト設定ID検索SQL
    ''' </summary>
    Private Const SQL_SELECT_標本リスト設定ID =
        "SELECT ID FROM 標本リスト設定" _
        & " WHERE 経営形態区分 = @経営形態区分" _
        & " AND 営農類型区分 = @営農類型区分" _
        & " AND 主副業区分 = @主副業区分" _
        & " AND 青色申告区分 = @青色申告区分" _
        & " AND 集落営農区分 = @集落営農区分" _
        & " AND 生産費区分 = @生産費区分" _
        & " AND 田畑区分 = @田畑区分"
    ''' <summary>
    ''' 標本リスト設定テーブルのIDを取得する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="経営形態区分"></param>
    ''' <param name="営農類型区分"></param>
    ''' <param name="主副業区分選択肢"></param>
    ''' <param name="青色申告区分"></param>
    ''' <param name="集落営農区分"></param>
    ''' <param name="生産費区分"></param>
    ''' <param name="田畑区分"></param>
    ''' <returns></returns>
    Public Shared Function Get標本リスト設定ID(db As DBAccess, 経営形態区分 As Integer, 営農類型区分 As Integer, 主副業区分選択肢 As Integer, 青色申告区分 As Integer, 集落営農区分 As Integer, 生産費区分 As Integer, 田畑区分 As Integer) As Integer
        Dim para = New List(Of DBAccess.Parameter)
        para.Add(db.CreateParameter("@経営形態区分", SqlDbType.Int, 経営形態区分))
        para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, 営農類型区分))
        para.Add(db.CreateParameter("@主副業区分", SqlDbType.Int, 主副業区分選択肢))
        para.Add(db.CreateParameter("@青色申告区分", SqlDbType.Int, 青色申告区分))
        para.Add(db.CreateParameter("@集落営農区分", SqlDbType.Int, 集落営農区分))
        para.Add(db.CreateParameter("@生産費区分", SqlDbType.Int, 生産費区分))
        para.Add(db.CreateParameter("@田畑区分", SqlDbType.Int, 田畑区分))
        Dim dt = db.GetDataTable(SQL_SELECT_標本リスト設定ID, para)
        If dt.Rows.Count > 0 Then
            Return CInt(dt.Rows(0)("ID"))
        End If
        Return 0

    End Function

    ''' <summary>
    ''' 標本リスト設定＿抽出条件検索SQL
    ''' </summary>
    Private Const SQL_SELECT_標本リスト設定＿抽出条件 =
        "SELECT " _
        & " BASE.規模番号," _
        & " CJ.上限," _
        & " CJ.下限" _
        & " FROM (SELECT TOP(@規模階層数) ROW_NUMBER() OVER(ORDER BY OBJECT_ID) AS 規模番号 FROM SYS.OBJECTS) BASE" _
        & " LEFT JOIN 標本リスト設定＿抽出条件 CJ" _
        & " ON CJ.ID = @ID" _
        & " AND CJ.規模番号 = BASE.規模番号" _
        & " ORDER BY BASE.規模番号"
    ''' <summary>
    ''' 標本リスト設定＿抽出条件検索
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="id"></param>
    ''' <returns></returns>
    Public Shared Function Select標本リスト設定＿抽出条件(db As DBAccess, id As Integer) As DataTable
        Dim para = New List(Of DBAccess.Parameter)
        para.Add(db.CreateParameter("@規模階層数", SqlDbType.Int, HyohonConst.規模階層区分リスト.Count - 1))
        para.Add(db.CreateParameter("@ID", SqlDbType.Int, id))
        Return db.GetDataTable(SQL_SELECT_標本リスト設定＿抽出条件, para)
    End Function

    ''' <summary>
    ''' 標本リスト設定＿期待値検索SQL
    ''' </summary>
    Private Const SQL_SELECT_標本リスト設定＿期待値 =
        "SELECT" _
        & " M.農政局," _
        & " M.名称," _
        & " BASE.規模番号," _
        & " SUM(ISNULL(K.標本サイズ, 0)) OVER(PARTITION BY K.農政局) AS 横計," _
        & " ISNULL(K.標本サイズ, 0) AS 標本サイズ" _
        & " FROM 地方農政局等マスタ M" _
        & " CROSS JOIN (SELECT TOP(@規模階層数) ROW_NUMBER() OVER(ORDER BY OBJECT_ID) AS 規模番号 FROM SYS.OBJECTS) BASE" _
        & " LEFT JOIN 標本リスト設定＿期待値 K" _
        & " ON K.ID = @ID" _
        & " AND K.農政局 = M.農政局" _
        & " AND K.規模番号 = BASE.規模番号" _
        & " WHERE M.農政局 = IIF(@農政局 = -1, M.農政局, @農政局)" _
        & " AND M.都道府県 = @都道府県＿全国＿農政局" _
        & " ORDER BY M.農政局, BASE.規模番号"
    ''' <summary>
    ''' 標本リスト設定＿期待値検索
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="id"></param>
    ''' <param name="kyoku"></param>
    ''' <returns></returns>
    Public Shared Function Select標本リスト設定＿期待値(db As DBAccess, id As Integer, Optional kyoku As Integer = -1) As DataTable
        Dim para = New List(Of DBAccess.Parameter)
        para.Add(db.CreateParameter("@規模階層数", SqlDbType.Int, HyohonConst.規模階層区分リスト.Count - 1))
        para.Add(db.CreateParameter("@ID", SqlDbType.Int, id))
        para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kyoku))
        para.Add(db.CreateParameter("@都道府県＿全国＿農政局", SqlDbType.Int, 都道府県＿全国＿農政局))
        Return db.GetDataTable(SQL_SELECT_標本リスト設定＿期待値, para)
    End Function

End Class
