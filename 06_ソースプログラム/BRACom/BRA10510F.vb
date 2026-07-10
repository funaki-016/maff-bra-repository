'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_000   | 2023.01.13 |大興電子通信        | 要件No.5 新規作成
'//            |            |                    |
'//*************************************************************************************************

''' <summary>
''' 標本リスト送受信画面
''' </summary>
''' <remarks></remarks>
Public Class BRA10510F

    ''' <summary>
    ''' 標本リスト送受信区分
    ''' </summary>
    Private ReadOnly SendRecv As HyohonConst.標本リスト送受信区分

    ''' <summary>
    ''' 上り下り区分
    ''' </summary>
    Private ReadOnly UpDown As String

    ''' <summary>
    ''' 送受信先データベース
    ''' </summary>
    Private ReadOnly DestDB As String

    ''' <summary>
    ''' センサス実施年
    ''' </summary>
    Private CensusNen As Integer

    ''' <summary>
    ''' 送受信可否
    ''' </summary>
    Private Enum 送受信可否
        送受信可 = 1
        送受信不可 = 0
    End Enum

    ''' <summary>
    ''' 一覧列INDEX
    ''' </summary>
    Private Enum DGV_COL
        選択 = 0
        農政局名 = 1
        実査設置拠点名 = 2
        経営形態区分名 = 3
        営農類型＿生産費区分名 = 4
        更新_送信日時 = 5
        送受信日時 = 6
        農政局 = 7
        都道府県 = 8
        実査設置拠点 = 9
        経営形態区分 = 10
        営農類型区分 = 11
        生産費区分 = 12
        田畑区分 = 13
        送受信可否 = 14
        送受信日時チェック用 = 15
    End Enum

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="sendRecv">標本リスト送受信区分</param>
    ''' <param name="upDown">標本リスト上り下り区分</param>
    ''' <remarks></remarks>
    Public Sub New(sendRecv As HyohonConst.標本リスト送受信区分, upDown As String)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

        Me.SendRecv = sendRecv
        Me.UpDown = upDown

        If Me.SendRecv = HyohonConst.標本リスト送受信区分.送信 Then
            lblSyori.Text = "標本リスト送信"
            dgvList.Columns(DGV_COL.更新_送信日時).HeaderText = "更新日時"
            dgvList.Columns(DGV_COL.送受信日時).HeaderText = "送信日時"
            BtnSendRecv.Text = "送信"
            If Me.UpDown = ComConst.上り下り区分.上位工程への送信 Then
                Select Case CommonInfo.Koutei
                    Case CommonInfo.KouteiKubun.Code.Kyoku
                        DestDB = "BRAH.dbo."
                        lblSyori.Text &= "（本省へ）"
                    Case CommonInfo.KouteiKubun.Code.Center
                        DestDB = "BRAN.dbo."
                        lblSyori.Text &= "（局へ）"
                    Case Else
                        Throw New Exception("パラメータエラー")
                End Select
            Else
                Select Case CommonInfo.Koutei
                    Case CommonInfo.KouteiKubun.Code.Kyoku
                        DestDB = "BRAS.dbo."
                        lblSyori.Text &= "（実査設置拠点へ）"
                    Case Else
                        Throw New Exception("パラメータエラー")
                End Select
            End If
        Else
            lblSyori.Text = "標本リスト受信"
            dgvList.Columns(DGV_COL.更新_送信日時).HeaderText = "送信日時"
            dgvList.Columns(DGV_COL.送受信日時).HeaderText = "受信日時"
            BtnSendRecv.Text = "受信"
            If Me.UpDown = ComConst.上り下り区分.上位工程への送信 Then
                Select Case CommonInfo.Koutei
                    Case CommonInfo.KouteiKubun.Code.Honsyo
                        DestDB = "BRAN.dbo."
                        lblSyori.Text &= "（局から）"
                    Case CommonInfo.KouteiKubun.Code.Kyoku
                        DestDB = "BRAS.dbo."
                        lblSyori.Text &= "（実査設置拠点から）"
                    Case Else
                        Throw New Exception("パラメータエラー")
                End Select
            Else
                Select Case CommonInfo.Koutei
                    Case CommonInfo.KouteiKubun.Code.Center
                        DestDB = "BRAN.dbo."
                        lblSyori.Text &= "（局から）"
                    Case Else
                        Throw New Exception("パラメータエラー")
                End Select
            End If
        End If

    End Sub

    ''' <summary>
    ''' 初期表示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA10510F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'センサス実施年設定
            HyohonUtil.SetCombobox(CboCensusNen, HyohonConst.センサス実施年リスト)
            CboCensusNen.SelectedIndex = 0
            '農政局プルダウン設定
            ComUtil.SetKyokuComboBox(CboKyoku)

            'DataGridView設定
            ComUtil.ConfigDgv(dgvList)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 農政局プルダウン選択
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CboKyoku_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboKyoku.SelectedIndexChanged
        Try
            '拠点コンボボックス設定
            ComUtil.SetKyotenComboBox(CboKyoku, CboKyoten)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 送信データ検索SQL
    ''' </summary>
    Private Const SQL_SELECT_送信データ =
        "SELECT" _
        & " KM.局名 AS 農政局名," _
        & " CM.センター名 AS 実査設置拠点名," _
        & " HL.農政局," _
        & " HL.都道府県," _
        & " HL.実査設置拠点," _
        & " HL.経営形態区分," _
        & " HL.営農類型区分," _
        & " HL.生産費区分," _
        & " HL.田畑区分," _
        & " FORMAT(HL.更新日時, 'yyyy/MM/dd HH:mm') AS 更新＿送信日時," _
        & " FORMAT(SJK.送信日時, 'yyyy/MM/dd HH:mm') AS 送受信日時," _
        & " FORMAT(ISNULL(SJK.送信日時, 0), 'yyyyMMddHHmmssfff') AS 送受信日時チェック用," _
        & " IIF(HL.更新日時 > SJK.送信日時 OR SJK.送信日時 IS NULL, @送受信可, @送受信不可) AS 送受信可否" _
        & " FROM" _
        & " (SELECT" _
        & " 農政局," _
        & " 都道府県," _
        & " 実査設置拠点," _
        & " 経営形態区分," _
        & " 営農類型区分," _
        & " 生産費区分," _
        & " 田畑区分," _
        & " MAX(更新日付) AS 更新日時" _
        & " FROM {0}" _
        & " WHERE 農政局 = IIF(@農政局 = @識別対象外値, 農政局, @農政局)" _
        & " AND 都道府県 = IIF(@都道府県 = @識別対象外値, 都道府県, @都道府県)" _
        & " AND 実査設置拠点 = IIF(@実査設置拠点 = @識別対象外値, 実査設置拠点, @実査設置拠点)" _
        & " GROUP BY" _
        & " 農政局," _
        & " 都道府県," _
        & " 実査設置拠点," _
        & " 経営形態区分," _
        & " 営農類型区分," _
        & " 生産費区分," _
        & " 田畑区分" _
        & " ) HL" _
        & " LEFT JOIN {1} SJK" _
        & "  ON SJK.農政局 = HL.農政局" _
        & " AND SJK.都道府県 = HL.都道府県" _
        & " AND SJK.実査設置拠点 = HL.実査設置拠点" _
        & " AND SJK.経営形態区分 = HL.経営形態区分" _
        & " AND SJK.営農類型区分 = HL.営農類型区分" _
        & " AND SJK.生産費区分 = HL.生産費区分" _
        & " AND SJK.田畑区分 = HL.田畑区分" _
        & " AND SJK.上り下り区分 = @上り下り区分" _
        & " JOIN COMM.dbo.局名マスタ KM" _
        & "  ON KM.局コード = HL.農政局" _
        & " JOIN センター名マスタ CM" _
        & "  ON CM.事務所番号 = IIF(HL.都道府県 = 1, 51, HL.都道府県)" _
        & " AND CM.センター番号 = HL.実査設置拠点" _
        & " ORDER BY" _
        & " 送受信可否 DESC," _
        & " HL.農政局," _
        & " HL.都道府県," _
        & " HL.実査設置拠点," _
        & " HL.経営形態区分," _
        & " HL.営農類型区分," _
        & " HL.生産費区分," _
        & " HL.田畑区分"

    ''' <summary>
    ''' 受信データ検索SQL
    ''' </summary>
    Private Const SQL_SELECT_受信データ =
        "SELECT" _
        & " KM.局名 AS 農政局名," _
        & " CM.センター名 AS 実査設置拠点名," _
        & " SJK.農政局," _
        & " SJK.都道府県," _
        & " SJK.実査設置拠点," _
        & " SJK.経営形態区分," _
        & " SJK.営農類型区分," _
        & " SJK.生産費区分," _
        & " SJK.田畑区分," _
        & " FORMAT(SJK.送信日時, 'yyyy/MM/dd HH:mm') AS 更新＿送信日時," _
        & " FORMAT(SJK.受信日時, 'yyyy/MM/dd HH:mm') AS 送受信日時," _
        & " FORMAT(ISNULL(SJK.受信日時, 0), 'yyyyMMddHHmmssfff') AS 送受信日時チェック用," _
        & " IIF(SJK.送信日時 > SJK.受信日時 OR SJK.受信日時 IS NULL, @送受信可, @送受信不可) AS 送受信可否" _
        & " FROM {0} SJK" _
        & " JOIN COMM.dbo.局名マスタ KM" _
        & "  ON KM.局コード = SJK.農政局" _
        & " JOIN センター名マスタ CM" _
        & "  ON CM.事務所番号 = IIF(SJK.都道府県 = 1, 51, SJK.都道府県)" _
        & " AND CM.センター番号 = SJK.実査設置拠点" _
        & " WHERE SJK.農政局 = IIF(@農政局 = @識別対象外値, 農政局, @農政局)" _
        & " AND SJK.都道府県 = IIF(@都道府県 = @識別対象外値, 都道府県, @都道府県)" _
        & " AND SJK.実査設置拠点 = IIF(@実査設置拠点 = @識別対象外値, 実査設置拠点, @実査設置拠点)" _
        & " AND SJK.上り下り区分 = @上り下り区分" _
        & " ORDER BY" _
        & " 送受信可否 DESC," _
        & " SJK.農政局," _
        & " SJK.都道府県," _
        & " SJK.実査設置拠点," _
        & " SJK.経営形態区分," _
        & " SJK.営農類型区分," _
        & " SJK.生産費区分," _
        & " SJK.田畑区分"

    ''' <summary>
    ''' 表示ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnShow_Click(sender As Object, e As EventArgs) Handles BtnShow.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            'センサス実施年退避
            CensusNen = HyohonUtil.GetComboValue(CboCensusNen)

            '送受信データ検索
            Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Dim para = New List(Of DBAccess.Parameter)
                para.Add(db.CreateParameter("@送受信可", SqlDbType.Int, 送受信可否.送受信可))
                para.Add(db.CreateParameter("@送受信不可", SqlDbType.Int, 送受信可否.送受信不可))
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, HyohonUtil.GetComboValue(CboKyoku)))
                para.Add(db.CreateParameter("@識別対象外値", SqlDbType.Int, HyohonConst.識別対象外値))
                Dim kyoten = HyohonUtil.GetComboValue(CboKyoten)
                para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kyoten))
                If kyoten = HyohonConst.識別対象外値 Then
                    para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, HyohonConst.識別対象外値))
                Else
                    Dim jimusho = CInt(CType(CboKyoten.SelectedItem, DataRowView)("事務所番号"))
                    para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, If(jimusho = 51, 1, jimusho)))
                End If
                para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, UpDown))
                Dim sql = ""
                If SendRecv = HyohonConst.標本リスト送受信区分.送信 Then
                    '送信時
                    sql = String.Format(SQL_SELECT_送信データ,
                                        String.Format("標本リスト{0}", CensusNen),
                                        String.Format("{0}標本リスト{1}送受信管理", DestDB, CensusNen))
                Else
                    '受信時
                    sql = String.Format(SQL_SELECT_受信データ, String.Format("標本リスト{0}送受信管理", CensusNen))
                End If

                Dim dt = db.GetDataTable(sql, para)

                '一覧表示
                dgvList.Rows.Clear()
                For Each row As DataRow In dt.Rows
                    Dim list = New List(Of Object)
                    Dim keieiKeitai = CInt(row("経営形態区分"))
                    Dim einoRuikei = CInt(row("営農類型区分"))
                    Dim seisanhi = CInt(row("生産費区分"))
                    Dim tahata = CInt(row("田畑区分"))
                    Dim sojushin = CInt(row("送受信可否"))
                    list.Add(False)
                    list.Add(row("農政局名"))
                    list.Add(row("実査設置拠点名"))
                    list.Add(HyohonUtil.Getリスト名称(HyohonUtil.Get経営形態区分リスト(), keieiKeitai))
                    list.Add(HyohonUtil.Get営農類型＿生産費区分名称(keieiKeitai, einoRuikei, seisanhi, tahata))
                    list.Add(row("更新＿送信日時"))
                    list.Add(row("送受信日時"))
                    list.Add(row("農政局"))
                    list.Add(row("都道府県"))
                    list.Add(row("実査設置拠点"))
                    list.Add(keieiKeitai)
                    list.Add(einoRuikei)
                    list.Add(seisanhi)
                    list.Add(tahata)
                    list.Add(sojushin)
                    list.Add(row("送受信日時チェック用"))
                    Dim addRow = HyohonUtil.AddDgvRow(dgvList, list.ToArray())
                    If sojushin = 送受信可否.送受信不可 Then
                        addRow.ReadOnly = True
                        addRow.DefaultCellStyle.BackColor = Color.WhiteSmoke
                    End If
                Next

                dgvList.ClearSelection()

            End Using

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 全選択ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnAllSelect_Click(sender As Object, e As EventArgs) Handles BtnAllSelect.Click
        Try
            ComUtil.SetDataGridViewAllCheckEnabledOnly(dgvList, True)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 全解除ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnAllRelease_Click(sender As Object, e As EventArgs) Handles BtnAllRelease.Click
        Try
            ComUtil.SetDataGridViewAllCheckEnabledOnly(dgvList, False)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 送受信済チェックSQL
    ''' </summary>
    Private Const SQL_SELECT_送受信済チェック =
        "SELECT 1 FROM {0}" _
        & " WHERE 農政局 = @農政局" _
        & " AND 都道府県 = @都道府県" _
        & " AND 実査設置拠点 = @実査設置拠点" _
        & " AND 経営形態区分 = @経営形態区分" _
        & " AND 営農類型区分 = @営農類型区分" _
        & " AND 生産費区分 = @生産費区分" _
        & " AND 田畑区分 = @田畑区分" _
        & " AND FORMAT(ISNULL({1}, 0), 'yyyyMMddHHmmssfff') <> @送受信日時"

    ''' <summary>
    ''' 送受信ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnSendRecv_Click(sender As Object, e As EventArgs) Handles BtnSendRecv.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            '入力チェック
            Dim selectedRows = dgvList.Rows().Cast(Of DataGridViewRow).Where(Function(x) CBool(x.Cells(DGV_COL.選択).Value)).ToList()
            If selectedRows.Count <= 0 Then
                Message.ShowMsgBox(If(SendRecv = HyohonConst.標本リスト送受信区分.送信, MessageID.MSG_E_135, MessageID.MSG_E_136), MsgBoxStyle.OkOnly)
                Return
            End If

            '実行確認
            If Message.ShowMsgBox(If(SendRecv = HyohonConst.標本リスト送受信区分.送信, MessageID.MSG_Q_063, MessageID.MSG_Q_064), MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If

            Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '進捗ダイアログ
                Dim progressDialog = New ProgressDialog()
                Try
                    '進捗ダイアログ表示
                    progressDialog.Maximum = selectedRows.Count
                    progressDialog.Show(Me)

                    Dim para As List(Of DBAccess.Parameter)
                    Dim sql = ""
                    Dim dbProcessFlg = False
                    For Each row In selectedRows
                        '処理済チェック
                        para = New List(Of DBAccess.Parameter)
                        para.Add(db.CreateParameter("@農政局", SqlDbType.Int, row.Cells(DGV_COL.農政局).Value))
                        para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, row.Cells(DGV_COL.都道府県).Value))
                        para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, row.Cells(DGV_COL.実査設置拠点).Value))
                        para.Add(db.CreateParameter("@経営形態区分", SqlDbType.Int, row.Cells(DGV_COL.経営形態区分).Value))
                        para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, row.Cells(DGV_COL.営農類型区分).Value))
                        para.Add(db.CreateParameter("@生産費区分", SqlDbType.Int, row.Cells(DGV_COL.生産費区分).Value))
                        para.Add(db.CreateParameter("@田畑区分", SqlDbType.Int, row.Cells(DGV_COL.田畑区分).Value))
                        para.Add(db.CreateParameter("@送受信日時", SqlDbType.VarChar, row.Cells(DGV_COL.送受信日時チェック用).Value))
                        If SendRecv = HyohonConst.標本リスト送受信区分.送信 Then
                            sql = String.Format(SQL_SELECT_送受信済チェック,
                                                String.Format("{0}標本リスト{1}送受信管理", DestDB, CensusNen),
                                                "送信日時")
                        Else
                            sql = String.Format(SQL_SELECT_送受信済チェック,
                                                String.Format("標本リスト{0}送受信管理", CensusNen),
                                                "受信日時")
                        End If
                        If db.GetDataTable(sql, para).Rows.Count <= 0 Then
                            'DB処理
                            DBProcess(db, row)
                            dbProcessFlg = True
                        End If

                        '進捗ダイアログを進める
                        progressDialog.AddValue = 1
                    Next

                    db.CommitTrans()

                    '送受信データ有無
                    If Not dbProcessFlg Then
                        progressDialog.ShowMsgBox(If(SendRecv = HyohonConst.標本リスト送受信区分.送信, MessageID.MSG_E_137, MessageID.MSG_E_138), MsgBoxStyle.OkOnly)
                        Return
                    End If

                    '表示ボタンクリック
                    BtnShow.PerformClick()

                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                Finally
                    If progressDialog IsNot Nothing Then
                        '進捗ダイアログを閉じる
                        progressDialog.endDispose()
                        progressDialog = Nothing
                    End If
                End Try
            End Using

            '完了メッセージ
            Message.ShowMsgBox(If(SendRecv = HyohonConst.標本リスト送受信区分.送信, MessageID.MSG_I_052, MessageID.MSG_I_053), MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 送受信済データ削除SQL
    ''' </summary>
    Private Const SQL_DELETE_送受信済データ =
        "DELETE FROM {0}" _
        & " WHERE 農政局 = @農政局" _
        & " AND 都道府県 = @都道府県" _
        & " AND 実査設置拠点 = @実査設置拠点" _
        & " AND 経営形態区分 = @経営形態区分" _
        & " AND 営農類型区分 = @営農類型区分" _
        & " AND 生産費区分 = @生産費区分" _
        & " AND 田畑区分 = @田畑区分" _
        & " OR NO1 IN (SELECT NO1 FROM {1}" _
        & " WHERE 農政局 = @農政局" _
        & " AND 都道府県 = @都道府県" _
        & " AND 実査設置拠点 = @実査設置拠点" _
        & " AND 経営形態区分 = @経営形態区分" _
        & " AND 営農類型区分 = @営農類型区分" _
        & " AND 生産費区分 = @生産費区分" _
        & " AND 田畑区分 = @田畑区分)" _
        & " AND 経営形態区分 = @経営形態区分" _
        & " AND 営農類型区分 = @営農類型区分" _
        & " AND 生産費区分 = @生産費区分" _
        & " AND 田畑区分 = @田畑区分"

    ''' <summary>
    ''' 送受信データ登録SQL
    ''' </summary>
    Private Const SQL_INSERT_送受信データ =
        "INSERT INTO {0}" _
        & " SELECT * FROM {1}" _
        & " WHERE 農政局 = @農政局" _
        & " AND 都道府県 = @都道府県" _
        & " AND 実査設置拠点 = @実査設置拠点" _
        & " AND 経営形態区分 = @経営形態区分" _
        & " AND 営農類型区分 = @営農類型区分" _
        & " AND 生産費区分 = @生産費区分" _
        & " AND 田畑区分 = @田畑区分"

    ''' <summary>
    ''' 送受信管理更新SQL
    ''' </summary>
    Private Const SQL_UPDATE_送受信管理 =
        "MERGE INTO {0} SJK" _
        & " USING (SELECT" _
        & " @農政局 AS 農政局," _
        & " @都道府県 AS 都道府県," _
        & " @実査設置拠点 AS 実査設置拠点," _
        & " @経営形態区分 AS 経営形態区分," _
        & " @営農類型区分 AS 営農類型区分," _
        & " @生産費区分 AS 生産費区分," _
        & " @田畑区分 AS 田畑区分," _
        & " @上り下り区分 AS 上り下り区分" _
        & " ) DAT" _
        & " ON SJK.農政局 = DAT.農政局" _
        & " AND SJK.都道府県 = DAT.都道府県" _
        & " AND SJK.実査設置拠点 = DAT.実査設置拠点" _
        & " AND SJK.経営形態区分 = DAT.経営形態区分" _
        & " AND SJK.営農類型区分 = DAT.営農類型区分" _
        & " AND SJK.生産費区分 = DAT.生産費区分" _
        & " AND SJK.田畑区分 = DAT.田畑区分" _
        & " AND SJK.上り下り区分 = DAT.上り下り区分" _
        & " WHEN MATCHED THEN" _
        & " UPDATE SET" _
        & " {1} = SYSDATETIME()," _
        & " 更新日付 = SYSDATETIME()," _
        & " 更新者ID = @更新者ID" _
        & " WHEN NOT MATCHED THEN" _
        & " INSERT VALUES(" _
        & " DAT.農政局," _
        & " DAT.都道府県," _
        & " DAT.実査設置拠点," _
        & " DAT.経営形態区分," _
        & " DAT.営農類型区分," _
        & " DAT.生産費区分," _
        & " DAT.田畑区分," _
        & " DAT.上り下り区分," _
        & " SYSDATETIME()," _
        & " NULL," _
        & " SYSDATETIME()," _
        & " @更新者ID);"
    ''' <summary>
    ''' DB処理
    ''' </summary>
    ''' <param name="db"></param>
    Private Sub DBProcess(db As DBAccess, row As DataGridViewRow)
        Dim sqlIns As String
        Dim sqlUpd As String
        Dim sqlDel As String
        If SendRecv = HyohonConst.標本リスト送受信区分.送信 Then
            '送信時
            sqlDel = String.Format(SQL_DELETE_送受信済データ,
                        String.Format("{0}受信＿標本リスト{1}", DestDB, CensusNen),
                        String.Format("標本リスト{0}", CensusNen))
            sqlIns = String.Format(SQL_INSERT_送受信データ,
                        String.Format("{0}受信＿標本リスト{1}", DestDB, CensusNen),
                        String.Format("標本リスト{0}", CensusNen))
            sqlUpd = String.Format(SQL_UPDATE_送受信管理,
                        String.Format("{0}標本リスト{1}送受信管理", DestDB, CensusNen),
                        "送信日時")
        Else
            '受信時
            sqlDel = String.Format(SQL_DELETE_送受信済データ,
                        String.Format("標本リスト{0}", CensusNen),
                        String.Format("受信＿標本リスト{0}", CensusNen))
            sqlIns = String.Format(SQL_INSERT_送受信データ,
                        String.Format("標本リスト{0}", CensusNen),
                        String.Format("受信＿標本リスト{0}", CensusNen))
            sqlUpd = String.Format(SQL_UPDATE_送受信管理,
                        String.Format("標本リスト{0}送受信管理", CensusNen),
                        "受信日時")
        End If
        Dim para = New List(Of DBAccess.Parameter)
        para.Add(db.CreateParameter("@農政局", SqlDbType.Int, row.Cells(DGV_COL.農政局).Value))
        para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, row.Cells(DGV_COL.都道府県).Value))
        para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, row.Cells(DGV_COL.実査設置拠点).Value))
        para.Add(db.CreateParameter("@経営形態区分", SqlDbType.Int, row.Cells(DGV_COL.経営形態区分).Value))
        para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, row.Cells(DGV_COL.営農類型区分).Value))
        para.Add(db.CreateParameter("@生産費区分", SqlDbType.Int, row.Cells(DGV_COL.生産費区分).Value))
        para.Add(db.CreateParameter("@田畑区分", SqlDbType.Int, row.Cells(DGV_COL.田畑区分).Value))

        '送受信済データ削除
        db.ExecuteNonQuery(sqlDel, para)

        '送受信データ登録
        db.ExecuteNonQuery(sqlIns, para)

        '送受信管理更新
        para.Add(db.CreateParameter("@更新者ID", SqlDbType.Int, CommonInfo.UserId))
        para.Add(db.CreateParameter("@上り下り区分", SqlDbType.Int, UpDown))
        db.ExecuteNonQuery(sqlUpd, para)
    End Sub
End Class
