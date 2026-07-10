'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.11.06 |TSP)                | フェーズ3 要件No.1、6修正
'//  REV_002   | 2021.03.19 |TSP)T.Minesaki      | フェーズ3 連絡票No696 システムエラー（Dictionary のキー重複） 対応　
'//  REV_003   | 2021.03.22 |TSP)T.Minesaki      | フェーズ3 連絡票No696 位置ずれ対応　
'//  REV_004   | 2023.01.04 |大興電子通信        | 要件No.19
'//  REV_005   | 2023.04.21 |大興電子通信        | 変更要件No.3
'//*************************************************************************************************

''' <summary>
''' 牛トレサプレプリント（総括データ）作成画面
''' </summary>
''' <remarks></remarks>
Public Class BRA2210F

    Private Class 調査票項目
        Public Shared 識別番号 As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, "Q11021801"},
            {ComConst.牛トレサデータ.調査区分分類.子牛, "Q02030301"},
            {ComConst.牛トレサデータ.調査区分分類.乳用, "Q02021401"},
            {ComConst.牛トレサデータ.調査区分分類.交雑, "Q02021401"},
            {ComConst.牛トレサデータ.調査区分分類.去勢, "Q02021401"}
        }

        'REV_003 MOD START
        'Public Shared 種類コード As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String) From {
        '    {ComConst.牛トレサデータ.調査区分分類.牛乳, If(CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別, "Q11022001", "Q11021901")},
        '    {ComConst.牛トレサデータ.調査区分分類.子牛, "Q02030501"},
        '    {ComConst.牛トレサデータ.調査区分分類.乳用, "Q02021501"},
        '    {ComConst.牛トレサデータ.調査区分分類.交雑, "Q02021501"},
        '    {ComConst.牛トレサデータ.調査区分分類.去勢, "Q02021501"}
        '}

        'Public Shared 品種コード As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String) From {
        '    {ComConst.牛トレサデータ.調査区分分類.牛乳, If(CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別, "Q11022101", "Q11022001")},
        '    {ComConst.牛トレサデータ.調査区分分類.子牛, "Q02030601"},
        '    {ComConst.牛トレサデータ.調査区分分類.乳用, "Q02021601"},
        '    {ComConst.牛トレサデータ.調査区分分類.交雑, "Q02021601"},
        '    {ComConst.牛トレサデータ.調査区分分類.去勢, "Q02021601"}
        '}
        'Public Shared 性別コード As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String) From {
        '    {ComConst.牛トレサデータ.調査区分分類.牛乳, If(CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別, "Q11022201", "Q11022101")},
        '    {ComConst.牛トレサデータ.調査区分分類.子牛, "Q02030801"},
        '    {ComConst.牛トレサデータ.調査区分分類.乳用, "Q02021701"},
        '    {ComConst.牛トレサデータ.調査区分分類.交雑, "Q02021701"},
        '    {ComConst.牛トレサデータ.調査区分分類.去勢, "Q02021701"}
        '}
        'Public Shared 生産年月 As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
        '    {ComConst.牛トレサデータ.調査区分分類.牛乳, If(CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別, {"Q11022301", "Q11022401"}, {"Q11022201", "Q11022301"})},
        '    {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02031001", "Q02031101"}},
        '    {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02021901", "Q02022001"}},
        '    {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02021901", "Q02022001"}},
        '    {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02021901", "Q02022001"}}
        '}
        'Public Shared 取得年月 As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
        '    {ComConst.牛トレサデータ.調査区分分類.牛乳, If(CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別, {"Q11022701", "Q11022801"}, {"Q11022601", "Q11022701"})},
        '    {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02031201", "Q02031301"}},
        '    {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02022101", "Q02022201"}},
        '    {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02022101", "Q02022201"}},
        '    {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02022101", "Q02022201"}}
        '}

        '一回目に牛乳生産費個別、アプリ終了せずに二回目牛乳生産費（法人）に変えた場合、牛乳生産費個別の値となり、１回目(CommonInfo.Chosakubun)の判断結果を保持し続けるため、方式変更
        'すべてにおいて、修正前後を残さないが、以下の方針で修正し、配列のどこを指すかはプログラムで判断
        '修正前：{ComConst.牛トレサデータ.調査区分分類.牛乳, If(CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別, "Q11022001", "Q11021901")}
        '修正後：{ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11022001", "Q11021901"}}
        Public Shared 種類コード As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11022001", "Dummy", "Q11021901", "Dummy"}},
            {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02030501"}},
            {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02021501"}},
            {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02021501"}},
            {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02021501"}}
        }

        Public Shared 品種コード As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11022101", "Dummy", "Q11022001", "Dummy"}},
            {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02030601"}},
            {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02021601"}},
            {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02021601"}},
            {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02021601"}}
        }
        Public Shared 性別コード As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11022201", "Dummy", "Q11022101", "Dummy"}},
            {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02030801"}},
            {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02021701"}},
            {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02021701"}},
            {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02021701"}}
        }
        Public Shared 生産年月 As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11022301", "Q11022401", "Q11022201", "Q11022301"}},
            {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02031001", "Q02031101"}},
            {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02021901", "Q02022001"}},
            {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02021901", "Q02022001"}},
            {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02021901", "Q02022001"}}
        }
        Public Shared 取得年月 As New Dictionary(Of ComConst.牛トレサデータ.調査区分分類, String()) From {
            {ComConst.牛トレサデータ.調査区分分類.牛乳, {"Q11022701", "Q11022801", "Q11022601", "Q11022701"}},
            {ComConst.牛トレサデータ.調査区分分類.子牛, {"Q02031201", "Q02031301"}},
            {ComConst.牛トレサデータ.調査区分分類.乳用, {"Q02022101", "Q02022201"}},
            {ComConst.牛トレサデータ.調査区分分類.交雑, {"Q02022101", "Q02022201"}},
            {ComConst.牛トレサデータ.調査区分分類.去勢, {"Q02022101", "Q02022201"}}
        }
        'REV_003 MOD END
    End Class



#Region "変数一覧"

    ''' <summary>調査年</summary>
    Private _chosaNen As String
    ''' <summary>局</summary>
    Private _kyoku As String
    ''' <summary>事務所</summary>
    Private _jimusho As String
    ''' <summary>拠点</summary>
    Private _kyoten As String
    ''' <summary>営農類型</summary>
    Private _einouRuikei As String

    ''' <summary>進捗ダイアログ</summary>
    Private ProgressDialog As ProgressDialog

#End Region

#Region "【処理詳細仕様 1】初期表示"

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA2210F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

            '【処理詳細仕様 1-2】調査年（産）設定

            '調査年コンボボックス設定
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.Chosahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
            End Using

            '【処理詳細仕様 1-3】農政局設定

            '局コンボボックス設定
            ComUtil.SetKyokuComboBox(cboKyoku)

            'REV-004 START------------------
            '全所有牛情報出力ボタン表示
            If CommonInfo.Kubun1 = ComConst.区分１.組織法人経営体に関する経営分析調査 Then
                btnMakeAll.Visible = False
            End If
            'REV-004 END------------------

            'DataGridView設定
            ComUtil.ConfigDgv(Me.dgvList)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 農政局コンボボックス選択時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboKyoku_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKyoku.SelectedIndexChanged
        Try
            '【処理詳細仕様 1-4】実査設置拠点設定

            '拠点コンボボックス設定
            ComUtil.SetKyotenComboBox(cboKyoku, cboKyoten)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

#End Region

#Region "【処理詳細仕様 3】「表示」ボタンクリック"

    ''' <summary>
    ''' 表示ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        Try
            '調査年選択チェック
            If cboChosaNen.SelectedValue Is Nothing Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            _chosaNen = cboChosaNen.SelectedValue.ToString
            _kyoku = If(IsDBNull(cboKyoku.SelectedValue), Nothing, cboKyoku.SelectedValue.ToString)
            _jimusho = If(IsDBNull(cboKyoten.SelectedValue), Nothing, CStr(CType(cboKyoten.SelectedItem, DataRowView)("事務所番号")))
            _kyoten = If(IsDBNull(cboKyoten.SelectedValue), Nothing, cboKyoten.SelectedValue.ToString)

            '一覧表示
            Me.ShowList(_chosaNen, _kyoku, _jimusho, _kyoten, _einouRuikei)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

#End Region

#Region "【処理詳細仕様 4】「作成」ボタンクリック"

    ''' <summary>
    ''' 作成ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnMake_Click(sender As Object, e As EventArgs) Handles btnMake.Click
        Try
            Dim keys As List(Of ValueTuple(Of DAOChosahyo.PrimaryKey, DAOChosahyo.KotenKey)) = Me.GetChosahyoKey(_chosaNen)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(keys, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_023, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If

            '作成結果を保持するための変数
            Dim resultList As New List(Of ValueTuple(Of Boolean, String))

            Try
                ProgressDialog = New ProgressDialog

                '進捗ダイアログを表示する
                ProgressDialog.Maximum = keys.Count
                ProgressDialog.Show(Me)

                '【処理詳細仕様 6-1】プレプリント（牛トレサ）調査票作成処理

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    Try
                        db.BeginTrans()

                        '選択したセンサス番号の数分処理
                        For Each key In keys

                            'プレプリント調査票を作成する
                            CreatePrePrintChosahyo(db, key, resultList)

                            '進捗を進める
                            ProgressDialog.AddValue = 1
                        Next

                        db.CommitTrans()

                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try
                End Using
            Finally
                If Not ProgressDialog Is Nothing Then
                    '進捗ダイアログを閉じる
                    ProgressDialog.endDispose()
                    ProgressDialog = Nothing
                End If
            End Try

            '【処理詳細仕様 6-2】完了メッセージ表示
            ShowMessage(resultList)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 完了メッセージ表示
    ''' </summary>
    ''' <param name="resultList"></param>
    ''' <remarks></remarks>
    Private Sub ShowMessage(ByVal resultList As List(Of ValueTuple(Of Boolean, String)))

        Dim ok As New List(Of String)
        Dim ng As New List(Of String)

        resultList.ForEach(Sub(row)
                               If row.Item1 Then
                                   ok.Add(row.Item2)
                               Else
                                   ng.Add(row.Item2)
                               End If
                           End Sub)

        Dim msgPara() As String = {String.Join(vbCrLf, ok), String.Join(vbCrLf, ng)}

        'メッセージフォーム表示
        Message.ShowMsgForm(Me, MessageID.MSG_I_042, msgPara)

    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="keys"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Function CheckError(keys As List(Of ValueTuple(Of DAOChosahyo.PrimaryKey, DAOChosahyo.KotenKey)), ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'センサス番号選択チェック
        If keys.Count = 0 Then
            msgId = MessageID.MSG_E_003
            Return ret
        End If

        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' プレプリント調査票を作成する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="key"></param>
    ''' <param name="resultList"></param>
    ''' <remarks></remarks>
    Private Sub CreatePrePrintChosahyo(ByVal db As DBAccess, ByVal key As ValueTuple(Of DAOChosahyo.PrimaryKey, DAOChosahyo.KotenKey), ByRef resultList As List(Of ValueTuple(Of Boolean, String)))

        '【処理詳細仕様 6-1 ①】プレプリント調査票作成先の総括データ存在チェック
        If DAOChosahyo.CheckTresaSummaryExist(db, _chosaNen, key.Item1.censusNo, CommonInfo.Chosakubun) Then
            '電子調査票に総括データが既に存在する場合

            'シート名を設定
            Dim sheetName As String = ComConst.調査票.牛総括データシート(CommonInfo.Chosakubun)

            '確認メッセージ
            If ProgressDialog.ShowMsgBox(MessageID.MSG_Q_042, {sheetName, key.Item1.censusNo}, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If
        End If


        '【処理詳細仕様 6-1 ②】牛トレサデータ取得処理
        '調査票データ取得
        Dim chosahyo = GetNowChosahyo(db, key.Item1.censusNo)
        '異動年月日を設定
        Dim idoDate As Integer = ComUtil.Tresa.GetIdoDateEnd(_chosaNen, chosahyo)
        If idoDate = -1 Then
            '決算月不明
            resultList.Add(ValueTuple.Create(False, key.Item1.censusNo))
            Return
        End If

        '農家団体コードをすべて取得
        Dim farmCodeList As List(Of String) = ComUtil.Tresa.GetFarmCodeList(chosahyo)

        '牛トレサ異動情報取得
        Dim dtTresa As DataTable = DAOOther.GetTresa(db, farmCodeList, 0, idoDate)

        '編成対象とする牛トレサデータを絞り込む
        Dim dtTresaHensei As DataTable = GetHenseiTresaData(dtTresa)

        '調査票の牛総括データを削除
        DAOChosahyo.DeleteTresaSoukatsu(db, _chosaNen, key.Item1.censusNo, CommonInfo.Chosakubun)

        Dim chosahyoInsert As Dictionary(Of String, DAOChosahyo.調査票項目)

        '牛トレサデータ⇒調査票のデータ変換
        chosahyoInsert = GetChosahyoData(db, dtTresaHensei, ComUtil.Tresa.GetChosaBeginDate(idoDate))

        'REV_005↓ 所有牛情報が更新されるため、職員の平均的な飼養頭数をクリアする
        'DAOChosahyo.InsertChosahyoTableKahen(db,
        '                                key.Item1,
        '                                New DAOChosahyo.KotenKey(_kyoku, _jimusho, _kyoten),
        '                                chosahyoInsert)
        '調査区分別にクリアする項番のリストを作成
        Dim clearList = ComUtil.GetShiyotosuClearList()
        '調査票データ更新
        DAOChosahyo.InsertChosahyoTableKahen(db,
                                        key.Item1,
                                        New DAOChosahyo.KotenKey(_kyoku, _jimusho, _kyoten),
                                        chosahyoInsert,
                                        clearList)
        'REV_005↑

        resultList.Add(ValueTuple.Create(True, key.Item1.censusNo))

    End Sub


    ''' <summary>
    ''' 画面で指定された調査年の調査票を取得する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetNowChosahyo(ByVal db As DBAccess, ByVal censusNo As String) As Dictionary(Of String, DAOChosahyo.調査票項目)

        '調査票項目マスタ取得
        Dim chosahyoItemMaster = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)

        '調査票テーブル取得
        Dim chosahyoTable = DAOChosahyo.GetChosahyoTable(db, New DAOChosahyo.PrimaryKey(_chosaNen, censusNo))

        '調査票項目取得
        Return ComUtil.Chosahyo.GetItem(chosahyoItemMaster, chosahyoTable)

    End Function

#End Region

#Region "「全所有牛情報出力」ボタンクリック"

    ''' <summary>
    ''' 全所有牛情報出力ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnMakeAll_Click(sender As Object, e As EventArgs) Handles btnMakeAll.Click
        Try
            Dim keys As List(Of ValueTuple(Of DAOChosahyo.PrimaryKey, DAOChosahyo.KotenKey)) = Me.GetChosahyoKey(_chosaNen)
            Dim CensusNo As String
            Dim fromKikan As String
            Dim toKikan As String
            Dim ChosaNen As String
            Dim lstHoseiList As New List(Of Dictionary(Of String, String))
            'Dim hoseiList = New List(Of BRA2120F.HoseiIdouInfo)

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(keys, msgId) Then
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
            If Message.ShowMsgBox(MessageID.MSG_Q_023, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If

            '検索条件で使用する、期間の作成
            ChosaNen = cboChosaNen.SelectedValue.ToString

            '作成結果を保持するための変数
            Dim resultList As New List(Of ValueTuple(Of Boolean, String))

            '画面で選択した調査票の分だけ以下の処理を行う。
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                For i As Integer = 0 To dgvList.Rows.Count - 1

                    If dgvList.Rows(i).Cells(0).Value IsNot Nothing _
                        And CBool(dgvList.Rows(i).Cells(0).Value) = True Then

                        'センサス番号取得
                        CensusNo = dgvList.Rows(i).Cells(7).Value.ToString

                        '調査票データ取得
                        Dim chosahyo = GetNowChosahyo(db, CensusNo)
                        '異動年月日を設定
                        Dim idoDate As Integer = ComUtil.Tresa.GetIdoDateEnd(_chosaNen, chosahyo)
                        If idoDate = -1 Then
                            '決算月不明
                            resultList.Add(ValueTuple.Create(False, CensusNo))
                            Return
                        End If

                        '農家団体コードをすべて取得
                        Dim farmCodeList As List(Of String) = ComUtil.Tresa.GetFarmCodeList(chosahyo)

                        '牛トレサ異動情報取得
                        Dim dtTresa As DataTable = DAOOther.GetTresa(db, farmCodeList, 0, idoDate)

                        '編成対象とする牛トレサデータを絞り込む
                        Dim dtTresaHensei As DataTable = GetHenseiTresaData(dtTresa)

                        Dim chosahyoInsert As Dictionary(Of String, DAOChosahyo.調査票項目)

                        '牛トレサデータ⇒調査票のデータ変換
                        chosahyoInsert = GetChosahyoData(db, dtTresaHensei, ComUtil.Tresa.GetChosaBeginDate(idoDate))
                    End If
                Next
            End Using

            Dim ret As ExcelOutputBaseClass.enmOutputReturn

            Try
                Dim pKeys As List(Of DAOChosahyo.PrimaryKey) = Me.GetChosahyoPrimaryKey(_chosaNen)
                Using ExcelOutput = New BRA2210R(folderPath, _chosaNen)
                    ret = ExcelOutput.Execute(pKeys, Me, 3)
                End Using
                If ret = ExcelOutputBaseClass.enmOutputReturn.OK Then
                    '完了メッセージ
                    Message.ShowMsgBox(MessageID.MSG_I_002, MsgBoxStyle.OkOnly)
                End If
            Catch ex As ExcelOutputBaseClass.SaveAsException
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_006, MsgBoxStyle.OkOnly)
            Catch ex As Exception
                Throw
            End Try
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 調査票主キー取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChosahyoPrimaryKey(chosaNen As String) As List(Of DAOChosahyo.PrimaryKey)
        Dim ret As New List(Of DAOChosahyo.PrimaryKey)

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                Dim pkey As DAOChosahyo.PrimaryKey = New DAOChosahyo.PrimaryKey(chosaNen, dgvList.Rows(i).Cells(7).Value.ToString)
                ret.Add(pkey)
            End If
        Next

        Return ret
    End Function

#End Region


#Region "処理全般"

    ''' <summary>
    ''' 全選択ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAllSelect_Click(sender As Object, e As EventArgs) Handles btnAllSelect.Click
        Try
            ComUtil.SetDataGridViewAllCheck(dgvList, True)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 全解除ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAllCancel_Click(sender As Object, e As EventArgs) Handles btnAllCancel.Click
        Try
            ComUtil.SetDataGridViewAllCheck(dgvList, False)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 調査票キー取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChosahyoKey(chosaNen As String) As List(Of ValueTuple(Of DAOChosahyo.PrimaryKey, DAOChosahyo.KotenKey))
        Dim ret As New List(Of ValueTuple(Of DAOChosahyo.PrimaryKey, DAOChosahyo.KotenKey))

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                Dim pkey As DAOChosahyo.PrimaryKey = New DAOChosahyo.PrimaryKey(chosaNen, dgvList.Rows(i).Cells(7).Value.ToString)
                Dim kkey As DAOChosahyo.KotenKey = New DAOChosahyo.KotenKey(dgvList.Rows(i).Cells(9).Value.ToString, dgvList.Rows(i).Cells(10).Value.ToString, dgvList.Rows(i).Cells(11).Value.ToString)
                ret.Add(ValueTuple.Create(pkey, kkey))
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="einouRuikei"></param>
    ''' <remarks></remarks>
    Private Sub ShowList(chosaNen As String, kyoku As String, jimusho As String, kyoten As String, einouRuikei As String)
        Dim dt As DataTable

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dt = DAOChosahyo.GetChosahyoList(db, chosaNen, kyoku, jimusho, kyoten, einouRuikei, "", "")
        End Using

        dgvList.Rows.Clear()

        For Each row As DataRow In dt.Rows
            Dim censusNo As String = row("センサス番号").ToString

            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(1).Value = ComUtil.GetTodofuken(censusNo)
            dgvList.Rows(i).Cells(2).Value = ComUtil.GetShikuchoson(censusNo)
            dgvList.Rows(i).Cells(3).Value = ComUtil.GetKyuShikuchoson(censusNo)
            dgvList.Rows(i).Cells(4).Value = ComUtil.GetNogyoShuraku(censusNo)
            dgvList.Rows(i).Cells(5).Value = ComUtil.GetChosaku(censusNo)
            dgvList.Rows(i).Cells(6).Value = ComUtil.GetKyakutaiNo(censusNo)
            dgvList.Rows(i).Cells(7).Value = censusNo
            dgvList.Rows(i).Cells(8).Value = DateTime.Parse(row("更新日付").ToString).ToString(ComConst.DATETIME_FORMAT)
            dgvList.Rows(i).Cells(9).Value = row("農政局").ToString
            dgvList.Rows(i).Cells(10).Value = row("都道府県").ToString
            dgvList.Rows(i).Cells(11).Value = row("実査設置拠点").ToString
        Next
    End Sub

#End Region

    ''' <summary>
    ''' 編成対象の牛トレサデータ取得
    ''' </summary>
    ''' <param name="dtTresa">牛トレサデータ</param>
    ''' <returns>編成対象の牛トレサデータ</returns>
    ''' <remarks></remarks>
    Private Function GetHenseiTresaData(dtTresa As DataTable) As DataTable
        Dim dtTresaHensei As DataTable = dtTresa.Clone
        Dim tresaHenseiIndex As Integer = 0

        '個体識別番号をすべて抽出
        Dim dtKotai As DataTable = dtTresa.DefaultView.ToTable(True, "個体識別番号")
        For Each row As DataRow In dtKotai.Rows

            '個体識別番号ごとに牛の存在を確認
            Dim tresaRows As DataRow() = dtTresa.Select("個体識別番号 = '" & row("個体識別番号").ToString & "'")

            If tresaRows.Count <= 0 Then
                '編成対象外
                Continue For
            End If

            If ComUtil.OrganiezCheck.IsCowOrganize(tresaRows, tresaHenseiIndex) = False Then
                Continue For
            End If

            '編成対象に追加
            dtTresaHensei.ImportRow(tresaRows(tresaHenseiIndex))
        Next

        Return dtTresaHensei
    End Function

    ''' <summary>
    ''' 調査票へ挿入するデータの取得
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="dtTresaHensei">調査票への編成対象の牛トレサデータ</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChosahyoData(db As DBAccess, dtTresaHensei As DataTable, beginDate As Date) As Dictionary(Of String, DAOChosahyo.調査票項目)

        Dim choShuruiCode As String '種類コード
        Dim choHinsyuCode As String '品種コード
        Dim choSexCode As String '性別コード
        Dim chosaBunrui As ComConst.牛トレサデータ.調査区分分類

        Dim meisaiNo As Integer = 1

        Dim chosahyoInsert As New Dictionary(Of String, DAOChosahyo.調査票項目)

        'REV_003 Add START
        '経営分析調査_牛乳生産費の場合のみ、参照する配列番号を＋２する
        Dim gyunyuKeieiHosei As Integer = 0
        If CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
            gyunyuKeieiHosei = 2
        End If
        'REV_003 Add END

        '編成対象となる牛トレサデータ分ループ
        For Each row As DataRow In dtTresaHensei.Rows

            Dim kindCode As Integer = Integer.Parse(row("牛の識別CD").ToString)
            Dim sexCode As Integer = Integer.Parse(row("性別コード").ToString)
            Dim idouFlag As Integer = Integer.Parse(row("異動フラグ").ToString)

            If ComConst.牛トレサデータ.調査区分分類変換テーブル.ContainsKey(CommonInfo.Chosakubun) Then
                chosaBunrui = ComConst.牛トレサデータ.調査区分分類変換テーブル(CommonInfo.Chosakubun)
            Else
                '対象外の調査区分
                Exit For
            End If

            '編成有無をチェック
            If ComUtil.Tresa.IsOrganizeData(chosaBunrui, kindCode, sexCode, idouFlag) = False Then
                Continue For
            End If

            '種類コードを変換
            choShuruiCode = ComUtil.Tresa.GetShuruiCode(chosaBunrui, kindCode, sexCode)

            '品種コードを変換
            choHinsyuCode = ComUtil.Tresa.GetHinshuCode(chosaBunrui, kindCode, sexCode)

            '性別コードを変換
            choSexCode = ComConst.牛トレサデータ.性別コード変換テーブル(sexCode)

            '［以下調査票への挿入データ格納］

            '識別番号
            AddChosahyoData(chosahyoInsert, 調査票項目.識別番号(chosaBunrui), meisaiNo, row("個体識別番号").ToString)

            '種類　コード
            'REV_003 MOD START
            'AddChosahyoData(chosahyoInsert, 調査票項目.種類コード(chosaBunrui), meisaiNo, choShuruiCode)
            AddChosahyoData(chosahyoInsert, 調査票項目.種類コード(chosaBunrui)(0 + gyunyuKeieiHosei), meisaiNo, choShuruiCode)
            'REV_003 MOD END

            '品種　コード
            'REV_003 MOD START
            'AddChosahyoData(chosahyoInsert, 調査票項目.品種コード(chosaBunrui), meisaiNo, choHinsyuCode)
            AddChosahyoData(chosahyoInsert, 調査票項目.品種コード(chosaBunrui)(0 + gyunyuKeieiHosei), meisaiNo, choHinsyuCode)
            'REV_003 MOD END

            '母畜の識別番号
            If chosaBunrui = ComConst.牛トレサデータ.調査区分分類.子牛 Then
                AddChosahyoData(chosahyoInsert, "Q02030701", meisaiNo, row("母牛個体識別番号").ToString)
            End If

            '性別区分　コード
            'REV_003 MOD START
            'AddChosahyoData(chosahyoInsert, 調査票項目.性別コード(chosaBunrui), meisaiNo, choSexCode)
            AddChosahyoData(chosahyoInsert, 調査票項目.性別コード(chosaBunrui)(0 + gyunyuKeieiHosei), meisaiNo, choSexCode)
            'REV_003 MOD END

            '生産年月
            'REV_003 MOD START
            'AddChosahyoData(chosahyoInsert, 調査票項目.生産年月(chosaBunrui)(0), meisaiNo, row("生年月日").ToString.Substring(0, 4))
            'AddChosahyoData(chosahyoInsert, 調査票項目.生産年月(chosaBunrui)(1), meisaiNo, row("生年月日").ToString.Substring(4, 2))
            AddChosahyoData(chosahyoInsert, 調査票項目.生産年月(chosaBunrui)(0 + gyunyuKeieiHosei), meisaiNo, row("生年月日").ToString.Substring(0, 4))
            AddChosahyoData(chosahyoInsert, 調査票項目.生産年月(chosaBunrui)(1 + gyunyuKeieiHosei), meisaiNo, row("生年月日").ToString.Substring(4, 2))
            'REV_003 MOD END

            '成畜年月
            If chosaBunrui = ComConst.牛トレサデータ.調査区分分類.牛乳 Then
                If ComUtil.Tresa.IsAdultTarget(choShuruiCode, choHinsyuCode) Then
                    Dim adultDate As String = ComUtil.Tresa.GetAdultDate(db, row("個体識別番号").ToString, row("生年月日").ToString, beginDate.ToString("yyyyMMdd"))

                    If adultDate.Length = 8 Then
                        Dim kobanYear As String = "Q11022501"
                        Dim kobanMonth As String = "Q11022601"
                        If CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
                            kobanYear = "Q11022401"
                            kobanMonth = "Q11022501"
                        Else

                        End If
                        AddChosahyoData(chosahyoInsert, kobanYear, meisaiNo, adultDate.Substring(0, 4))
                        AddChosahyoData(chosahyoInsert, kobanMonth, meisaiNo, adultDate.Substring(4, 2))
                    End If
                End If
            End If

            '取得年月(牛乳)、転入(購入)　年月(子牛)、導入時　購入　年月(育成・肥育)
            If row("異動フラグ").ToString.Equals(ComConst.牛トレサデータ.異動フラグ.転入搬入.ToString) Then
                '異動フラグが 4(転入/搬入)の場合
                'REV_003 MOD START
                'AddChosahyoData(chosahyoInsert, 調査票項目.取得年月(chosaBunrui)(0), meisaiNo, row("異動年月日").ToString.Substring(0, 4))
                'AddChosahyoData(chosahyoInsert, 調査票項目.取得年月(chosaBunrui)(1), meisaiNo, row("異動年月日").ToString.Substring(4, 2))
                AddChosahyoData(chosahyoInsert, 調査票項目.取得年月(chosaBunrui)(0 + gyunyuKeieiHosei), meisaiNo, row("異動年月日").ToString.Substring(0, 4))
                AddChosahyoData(chosahyoInsert, 調査票項目.取得年月(chosaBunrui)(1 + gyunyuKeieiHosei), meisaiNo, row("異動年月日").ToString.Substring(4, 2))
                'REV_003 MOD END
            End If
            meisaiNo += 1
        Next

        Return chosahyoInsert
    End Function

    ''' <summary>
    ''' 調査票データ（可変）に追加する
    ''' </summary>
    ''' <param name="dic"></param>
    ''' <param name="itemNo"></param>
    ''' <param name="meisaiNo"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Private Sub AddChosahyoData(ByRef dic As Dictionary(Of String, DAOChosahyo.調査票項目), itemNo As String, meisaiNo As Integer, value As String)
        Dim insertKey As String
        Dim insertData As New DAOChosahyo.調査票項目

        insertKey = GetChosahyoKey(itemNo, meisaiNo)
        insertData.値 = value

        'REV_002 Mod Start
        'dic.Add(insertKey, insertData)

        '重複キーがない場合にディクショナリに追加する。
        If Not dic.ContainsKey(insertKey) Then
            dic.Add(insertKey, insertData)
        End If
        'REV_002 Mod End

    End Sub

    ''' <summary>
    ''' 調査票データ（可変）のキーを取得
    ''' </summary>
    ''' <param name="itemNo"></param>
    ''' <param name="meisaiNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChosahyoKey(itemNo As String, meisaiNo As Integer) As String
        Dim ret As String = itemNo

        ret = ret & ComConst.ITEM_NO_DELIMITER & meisaiNo.ToString

        Return ret
    End Function



End Class
