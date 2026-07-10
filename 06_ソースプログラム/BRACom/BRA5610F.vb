'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2023.01.13 |大興電子通信        | フェーズ2 要件No.3
'//            |            |                    |
'//*************************************************************************************************
Imports System.IO
Imports System.Text

''' <summary>
''' 集計結果表CSV出力画面
''' </summary>
''' <remarks></remarks>
Public Class BRA5610F

    ''' <summary>調査年</summary>
    Private _chosaNen As String
    ''' <summary>営農経営体区分</summary>
    Private _einouKeieitai As String

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA5610F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '営農経営体区分コンボボックス設定
            ComUtil.SetEinouKeieitaiComboBox(lblEinouKeieitai, cboEinouKeieitai)

            '調査区分取得
            Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

            If Not (CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計) Then
                '調査年コンボボックス設定
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    ComUtil.SyukeiKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, chosakubun, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
                End Using
            End If

            'REV_001↓
            '出力形式プルダウン設定
            If CInt(CommonInfo.Chosakubun) >= CInt(ComConst.調査区分.経営分析調査_二条大麦生産費) Then
                LblOutType.Visible = False
                CboOutType.Visible = False
            End If
            'REV_001↑

            'DataGridView設定
            ComUtil.ConfigDgv(Me.dgvList)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        Try
            '調査年選択チェック
            If cboChosaNen.SelectedValue Is Nothing Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            _chosaNen = cboChosaNen.SelectedValue.ToString
            _einouKeieitai = If(cboEinouKeieitai.SelectedValue Is Nothing, Nothing, cboEinouKeieitai.SelectedValue.ToString)

            '一覧表示
            Me.ShowList(_chosaNen, _einouKeieitai, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 出力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '集計結果表主キー取得
            Dim sKey As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey) = Me.GetSyukeiKekkahyoSelectKey(_chosaNen)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(sKey, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            'フォルダパス取得
            Dim folderPath As String = ComUtil.GetFolderPath(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath))

            If folderPath.Equals(String.Empty) Then
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_029, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                'REV_001↓
                Dim outType = DirectCast(CboOutType.SelectedValue, 出力形式)
                'REV_001↑

                Dim chosakubun As String = ComUtil.GetChosakubun(_einouKeieitai)

                Dim pKey As DAOSyukeiKekkahyo.PrimaryKey = sKey.Item1
                Dim kkey As DAOSyukeiKekkahyo.KomokuKey = sKey.Item2

                Dim dc As Dictionary(Of String, DataTable)
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    'REV_001↓
                    'dc = DAOSyukeiKekkahyo.GetSyukeiNoTable(db, chosakubun, pKey, kkey)
                    If outType = 出力形式.分割 Then
                        dc = DAOSyukeiKekkahyo.GetSyukeiNoTable(db, chosakubun, pKey, kkey)
                    Else
                        dc = Select統合集計結果表(db, GetSelect統合集計結果表SQL(db, outType, _chosaNen, chosakubun), pKey, kkey)
                    End If
                    'REV_001↑
                End Using

                Dim fileNamePattern As String = Me.GetFileNamePattern(pKey.chosaNen, pKey.syukeiNo)

                'CSV出力
                Dim ret As ComConst.CSVファイル.enmOutputReturn = Me.PutCSV(folderPath, fileNamePattern, dc)

                Select Case ret
                    Case ComConst.CSVファイル.enmOutputReturn.OK
                        '完了メッセージ
                        Message.ShowMsgBox(MessageID.MSG_I_002, MsgBoxStyle.OkOnly)
                    Case ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS
                        'エラーメッセージ
                        Message.ShowMsgBox(MessageID.MSG_E_058, MsgBoxStyle.OkOnly)
                End Select
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub dgvList_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvList.CellValueChanged
        If e.ColumnIndex = 0 Then
            If CType(dgvList(e.ColumnIndex, e.RowIndex).Value, Boolean) = True Then
                For rowIndex = 0 To dgvList.Rows.Count - 1
                    If rowIndex <> e.RowIndex Then
                        dgvList(0, rowIndex).Value = False
                        dgvList(0, rowIndex).ReadOnly = False
                    End If
                Next
                dgvList(e.ColumnIndex, e.RowIndex).ReadOnly = True
            End If
        End If
    End Sub

    Private Sub dgvList_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvList.CurrentCellDirtyStateChanged
        If dgvList.CurrentCellAddress.X = 0 AndAlso dgvList.IsCurrentCellDirty Then
            dgvList.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub cboEinouKeieitai_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboEinouKeieitai.SelectedIndexChanged
        dgvList.Rows.Clear()

        '調査区分取得
        Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

        '調査年コンボボックス設定
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            ComUtil.SyukeiKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, chosakubun, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
        End Using
    End Sub

    ''' <summary>
    ''' 集計結果表主キー取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSyukeiKekkahyoSelectKey(chosaNen As String) As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey)
        Dim ret As New ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey)

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                Dim pkey As DAOSyukeiKekkahyo.PrimaryKey = New DAOSyukeiKekkahyo.PrimaryKey(chosaNen, dgvList.Rows(i).Cells(1).Value.ToString)
                Dim kkey As DAOSyukeiKekkahyo.KomokuKey = New DAOSyukeiKekkahyo.KomokuKey(dgvList.Rows(i).Cells(5).Value.ToString, dgvList.Rows(i).Cells(6).Value.ToString, dgvList.Rows(i).Cells(7).Value.ToString)
                ret = ValueTuple.Create(pkey, kkey)

                Exit For
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <param name="chosakubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <remarks></remarks>
    Public Sub ShowList(chosaNen As String, einouKeieitai As String, kyoku As String, jimusho As String, kyoten As String)
        Dim dt As DataTable = Nothing

        Dim chosakubun As String = ComUtil.GetChosakubun(einouKeieitai)

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Select Case CommonInfo.Koutei
                Case CommonInfo.KouteiKubun.Code.Center
                    dt = DAOSyukeiKekkahyo.GetList(db, chosakubun, chosaNen, kyoku, jimusho, kyoten)
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    dt = DAOSyukeiKekkahyo.GetList(db, chosakubun, chosaNen, kyoku, Nothing, Nothing)
                Case CommonInfo.KouteiKubun.Code.Honsyo
                    dt = DAOSyukeiKekkahyo.GetList(db, chosakubun, chosaNen, Nothing, Nothing, Nothing)
            End Select
        End Using

        dgvList.Rows.Clear()

        For Each row As DataRow In dt.Rows
            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(1).Value = row("集計番号").ToString
            dgvList.Rows(i).Cells(2).Value = row("集計名称").ToString
            dgvList.Rows(i).Cells(3).Value = DateTime.Parse(row("更新日付").ToString).ToString(ComConst.DATETIME_FORMAT)
            dgvList.Rows(i).Cells(4).Value = row("集計条件").ToString
            dgvList.Rows(i).Cells(5).Value = row("農政局").ToString
            dgvList.Rows(i).Cells(6).Value = row("都道府県").ToString
            dgvList.Rows(i).Cells(7).Value = row("実査設置拠点").ToString

            dgvList.Rows(i).Cells(4).Style.Alignment = DataGridViewContentAlignment.MiddleLeft
        Next
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="keys"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Function CheckError(sKey As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey), ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        '集計番号選択チェック
        If sKey.Item1 Is Nothing Then
            msgId = MessageID.MSG_E_045
            Return ret
        End If

        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' CSV出力
    ''' </summary>
    ''' <param name="folderPath"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PutCSV(outDir As String, fileName As String, dc As Dictionary(Of String, DataTable)) As ComConst.CSVファイル.enmOutputReturn
        Dim ret As ComConst.CSVファイル.enmOutputReturn = ComConst.CSVファイル.enmOutputReturn.CANCEL

        Dim sjisEnc As Encoding = Encoding.GetEncoding(ComConst.CSVファイル.CODEPAGE_SHIFT_JIS)

        Dim filePathTemp As String = System.IO.Path.Combine(outDir, fileName)

        If Not System.IO.Directory.Exists(outDir) Then
            Directory.CreateDirectory(outDir)
        End If

        'REV_001↓
        Dim outType = DirectCast(CboOutType.SelectedValue, 出力形式)
        Dim itemNames As List(Of (項目番号 As String, 項目名 As String)) = Nothing
        If outType <> 出力形式.分割 Then
            itemNames = GetItemNames(_chosaNen, _einouKeieitai)
        End If
        'REV_001↑

        For Each kv As KeyValuePair(Of String, DataTable) In dc
            Dim tableName As String = kv.Key
            Dim dt As DataTable = kv.Value

            Try
                Using sw As New System.IO.StreamWriter(String.Format(filePathTemp, tableName), False, sjisEnc)
                    Dim colArr(dt.Columns.Count - 1) As Object
                    dt.Columns.CopyTo(colArr, 0)
                    sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, colArr) & ComConst.CSVファイル.END_ADDITION)

                    'REV_001↓
                    '統合版は２行目に項目名出力
                    If outType <> 出力形式.分割 Then
                        Dim nameArr = dt.Columns.Cast(Of DataColumn).Select(Function(x) If(itemNames.Where(Function(y) y.項目番号 = x.ColumnName).Select(Function(y) y.項目名).FirstOrDefault(), x.ColumnName)).ToArray()
                        sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, nameArr) & ComConst.CSVファイル.END_ADDITION)
                    End If
                    'REV_001↑

                    For Each row As DataRow In dt.Rows
                        Dim arr As Object() = row.ItemArray().ToArray()
                        sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, arr) & ComConst.CSVファイル.END_ADDITION)
                    Next
                End Using
            Catch ex As Exception
                ret = ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS
                Return ret
            End Try
        Next

        ret = ComConst.CSVファイル.enmOutputReturn.OK

        Return ret
    End Function

    ''' <summary>
    ''' ファイル名パターン取得
    ''' </summary>
    ''' <param name="year"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetFileNamePattern(year As String, syukeiNo As String) As String
        Return "{0}_" & year & "_" & syukeiNo & ".csv"
    End Function

    ''' <summary>
    ''' 出力形式リスト(REV_001)
    ''' </summary>
    Private Shared ReadOnly 出力形式リスト As New List(Of KeyValuePair(Of Integer, String)) From {
        New KeyValuePair(Of Integer, String)(出力形式.分割, "分割"),
        New KeyValuePair(Of Integer, String)(出力形式.統合, "統合")
    }
    ''' <summary>
    ''' 出力形式リスト取得(REV_001)
    ''' </summary>
    ''' <param name="nen"></param>
    ''' <returns></returns>
    Private Shared Function Get出力形式リスト(version As String) As List(Of KeyValuePair(Of Integer, String))
        Return 出力形式リスト.Where(Function(x) x.Key = 出力形式.分割 OrElse version = ComConst.バージョン区分.結果表等項目2022).ToList()
    End Function

    ''' <summary>
    ''' 調査年選択(REV_001)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cboChosaNen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboChosaNen.SelectedIndexChanged
        Try
            Dim nen = If(cboChosaNen.SelectedValue?.ToString(), "0")
            If _chosaNen = nen Then
                Return
            End If
            _chosaNen = nen
            '一覧クリア
            dgvList.Rows.Clear()
            '出力形式プルダウン
            Dim version = ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun)
            CboOutType.ValueMember = "Key"
            CboOutType.DisplayMember = "Value"
            CboOutType.DataSource = Get出力形式リスト(version)
            CboOutType.SelectedIndex = 0

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 統合集計結果表検索SQL取得(REV_001)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="outType"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="chosakubun"></param>
    ''' <returns></returns>
    Private Shared Function GetSelect統合集計結果表SQL(db As DBAccess, outType As 出力形式, chosaNen As String, chosakubun As String) As List(Of (tableName As String, sql As String))
        Dim sqlList = New List(Of (String, String))
        Dim colInfoLists = Get統合集計結果表列リスト(db, outType, chosaNen, chosakubun)
        Dim sbSql As StringBuilder
        Dim isFirst As Boolean
        Dim firstIndex = 0
        For Each colInfo In colInfoLists
            sbSql = New StringBuilder()
            sbSql.Append("SELECT")
            isFirst = True
            For Each colData In colInfo.colDataList
                If isFirst Then
                    isFirst = False
                Else
                    sbSql.Append(",")
                End If
                sbSql.Append(String.Format(" NO{0}.{1}", colData.index, colData.colName))
            Next
            isFirst = True
            For Each index In colInfo.colDataList.Select(Function(x) x.index).Distinct().OrderBy(Function(x) x)
                If isFirst Then
                    sbSql.Append(String.Format(" FROM {0} NO{1}", ComConst.集計結果表.テーブル名称(chosakubun)(index), index))
                    firstIndex = index
                    isFirst = False
                Else
                    sbSql.Append(String.Format(" JOIN {0} NO{1}", ComConst.集計結果表.テーブル名称(chosakubun)(index), index))
                    sbSql.Append(String.Format(" ON NO{0}.調査年 = NO{1}.調査年", index, firstIndex))
                    sbSql.Append(String.Format(" AND NO{0}.集計番号 = NO{1}.集計番号", index, firstIndex))
                    sbSql.Append(String.Format(" AND NO{0}.連番 = NO{1}.連番", index, firstIndex))
                End If
            Next
            sbSql.Append(String.Format(" WHERE NO{0}.調査年 = @調査年", firstIndex))
            sbSql.Append(String.Format(" AND NO{0}.集計番号 = @集計番号", firstIndex))
            sbSql.Append(String.Format(" AND NO{0}.農政局 = @農政局", firstIndex))
            sbSql.Append(String.Format(" AND NO{0}.都道府県 = @都道府県", firstIndex))
            sbSql.Append(String.Format(" AND NO{0}.実査設置拠点 = @実査設置拠点", firstIndex))
            sbSql.Append(String.Format(" ORDER BY NO{0}.連番", firstIndex))
            sqlList.Add((colInfo.tableName, sbSql.ToString()))
        Next
        Return sqlList
    End Function

    ''' <summary>
    ''' 統合集計結果表検索(REV_001)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="sqlInfoList"></param>
    ''' <param name="pkey"></param>
    ''' <param name="kkey"></param>
    ''' <returns></returns>
    Private Shared Function Select統合集計結果表(db As DBAccess, sqlInfoList As List(Of (tableName As String, sql As String)), pkey As DAOSyukeiKekkahyo.PrimaryKey, kkey As DAOSyukeiKekkahyo.KomokuKey) As Dictionary(Of String, DataTable)
        Dim dc = New Dictionary(Of String, DataTable)
        Dim para As List(Of DBAccess.Parameter)
        For Each sqlInfo In sqlInfoList
            para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, pkey.chosaNen))
            para.Add(db.CreateParameter("@集計番号", SqlDbType.VarChar, pkey.syukeiNo))
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kkey.kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kkey.jimusho))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kkey.kyoten))
            dc.Add(sqlInfo.tableName, db.GetDataTable(sqlInfo.sql, para))
        Next
        Return dc
    End Function

    ''' <summary>
    ''' 統合集計結果表列検索SQL(REV_001)
    ''' </summary>
    Private Const SQL_SELECT_統合集計結果表列 =
        "SELECT" _
        & " {0} AS SORT_NO," _
        & " C.COLUMN_ID," _
        & " {1} AS IDX," _
        & " C.NAME AS COL_NAME" _
        & " FROM SYS.OBJECTS O" _
        & " JOIN SYS.COLUMNS C" _
        & " ON C.OBJECT_ID = O.OBJECT_ID" _
        & " WHERE O.NAME = '{2}'" _
        & " AND C.NAME {3} IN ({4})" _
        & " AND (C.NAME NOT LIKE 'S%'" _
        & "  OR EXISTS(SELECT 1 FROM 集計結果表項目マスタ M" _
        & "  WHERE M.バージョン区分 = {5}" _
        & "  AND M.調査区分 = {6}" _
        & "  AND M.項目番号 = C.NAME" _
        & "  AND M.裏項番区分 = 0))"
    ''' <summary>
    ''' 統合集計結果表列検索ORDER(REV_001)
    ''' </summary>
    Private Const SQL_SELECT_統合集計結果表列_ORDER = " ORDER BY SORT_NO,COLUMN_ID"
    ''' <summary>
    ''' キー情報列(REV_001)
    ''' </summary>
    Private Const キー情報列 = "'調査年','集計番号','連番'"
    ''' <summary>
    ''' 更新情報列(REV_001)
    ''' </summary>
    Private Const 更新情報列 = "'更新日付','更新者ID'"
    ''' <summary>
    ''' 統合集計結果表列リスト取得(REV_001)
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="outType"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="chosakubun"></param>
    ''' <returns></returns>
    Private Shared Function Get統合集計結果表列リスト(db As DBAccess, outType As 出力形式, chosaNen As String, chosakubun As String) As List(Of (tableName As String, colDataList As List(Of (index As Integer, colName As String))))
        Dim colLists = New List(Of (String, List(Of (Integer, String))))
        Dim indexes = GetIndexes(outType, chosakubun)
        Dim tableName = ComConst.集計結果表.テーブル名称(chosakubun)(indexes.first)
        If indexes.first <> indexes.last Then
            tableName &= "＿" & Strings.Right(ComConst.集計結果表.テーブル名称(chosakubun)(indexes.last), (indexes.last + 1).ToString().Length)
        End If
        Dim version = ComUtil.getVersionKubunTaikei(chosaNen, chosakubun)
        Dim sbSql = New StringBuilder()
        For i = indexes.first To indexes.last
            If i = indexes.first Then
                'キー情報は先頭から出力
                sbSql.Append(String.Format(SQL_SELECT_統合集計結果表列, -1, 0, ComConst.集計結果表.テーブル名称(chosakubun)(0), "", キー情報列, version, chosakubun))
            End If
            'キー情報、更新情報を除外
            sbSql.Append(" UNION ALL ")
            sbSql.Append(String.Format(SQL_SELECT_統合集計結果表列, i, i, ComConst.集計結果表.テーブル名称(chosakubun)(i), "NOT", キー情報列 & "," & 更新情報列, version, chosakubun))
        Next
        sbSql.Append(" UNION ALL ")
        '更新情報は先頭から出力
        sbSql.Append(String.Format(SQL_SELECT_統合集計結果表列, indexes.last + 1, 0, ComConst.集計結果表.テーブル名称(chosakubun)(0), "", 更新情報列, version, chosakubun))
        sbSql.Append(SQL_SELECT_統合集計結果表列_ORDER)
        colLists.Add((tableName, db.GetDataTable(sbSql.ToString()).Rows.Cast(Of DataRow).Select(Function(x) (CInt(x("IDX")), x("COL_NAME").ToString())).ToList()))

        Return colLists
    End Function

    ''' <summary>
    ''' 統合対象テーブル名称のIndexを取得(REV_001)
    ''' </summary>
    ''' <param name="outType"></param>
    ''' <param name="chosakubun"></param>
    ''' <returns></returns>
    Private Shared Function GetIndexes(outType As 出力形式, chosakubun As String) As (first As Integer, last As Integer)
        Dim first = 0
        Dim last = 0
        Select Case outType
            Case 出力形式.統合
                last = UBound(ComConst.集計結果表.テーブル名称(chosakubun))
        End Select
        Return (first, last)
    End Function

    ''' <summary>
    ''' 出力形式(REV_001)
    ''' </summary>
    Private Enum 出力形式
        分割
        統合
    End Enum

    ''' <summary>
    ''' 集計結果表項目名検索SQL(REV_001)
    ''' </summary>
    Private Const SQL_SELECT_ITEM_NAMES =
        "SELECT DISTINCT" _
        & " 項目番号," _
        & " REPLACE(項目名, CHAR(10), ' ') AS 項目名" _
        & " FROM 集計結果表作成論理" _
        & " WHERE バージョン区分 = @バージョン区分" _
        & " AND 調査区分 = @調査区分" _
        & " ORDER BY 1"
    ''' <summary>
    ''' 集計結果表項目名取得(REV_001)
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="einouKeieitai"></param>
    ''' <returns></returns>
    Private Shared Function GetItemNames(chosaNen As String, einouKeieitai As String) As List(Of (項目番号 As String, 項目名 As String))
        Dim version = ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun)
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Dim para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, version))
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, ComUtil.GetChosakubun(einouKeieitai)))
            Dim dt = db.GetDataTable(SQL_SELECT_ITEM_NAMES, para)
            Return dt.Rows.Cast(Of DataRow).Select(Function(x) (x("項目番号").ToString(), x("項目名").ToString())).ToList()
        End Using
    End Function
End Class
