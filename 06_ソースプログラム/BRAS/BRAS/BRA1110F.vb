'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.12.23 |TSP)                | フェーズ3 要件No.1修正
'//  REV_002   | 2021.03.17 |TSP)T.Minesaki      | フェーズ3 連絡票No.695 追加した項目が総括表から削除されない問題修正
'//  REV_003   | 2021.11.24 |TNCS                | 要件No.1-②修正
'//  REV_004   | 2023.08.07 |Daiko               | 要件No.3-① 前年の調査開始以前に売却（死亡）した繁殖雌牛を削除
'//  REV_005   | 2023.08.10 |Daiko               | 要件No.12 営農のプレプリント処理修正
'//  REV_006   | 2024.05.31 |Daiko               | 追加要件No.2 修正
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' プレプリント調査票作成画面
''' </summary>
''' <remarks></remarks>
Public Class BRA1110F

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

    Private Const 出_入牧 As String = "1：出（入牧）"
    Private Const 売却_売却 As String = "2：売却（売却）"
    Private Const 売却_死亡 As String = "2：売却（死亡）"

    '---REV_001_ADD START
    Private Const 転出_売却 As String = "2：転出（売却）"
    Private Const 転出_死亡 As String = "2：転出（死亡）"
    '---REV_001_ADD END

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
    Private Sub BRA1110F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '【処理詳細仕様 1-2】調査年（産）設定

            '調査年コンボボックス設定
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.Chosahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center, True)
            End Using

            '【処理詳細仕様 1-3】農政局設定

            '局コンボボックス設定
            ComUtil.SetKyokuComboBox(cboKyoku)

            '営農類型コンボボックス設定
            ComUtil.SetEinouRuikeiComboBox(lblEinouRuikei, cboEinouRuikei)

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
            If cboChosaNen.SelectedValue Is DBNull.Value Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            _chosaNen = cboChosaNen.SelectedValue.ToString
            _kyoku = If(IsDBNull(cboKyoku.SelectedValue), Nothing, cboKyoku.SelectedValue.ToString)
            _jimusho = If(IsDBNull(cboKyoten.SelectedValue), Nothing, CStr(CType(cboKyoten.SelectedItem, DataRowView)("事務所番号")))
            _kyoten = If(IsDBNull(cboKyoten.SelectedValue), Nothing, cboKyoten.SelectedValue.ToString)
            _einouRuikei = If(cboEinouRuikei.SelectedValue Is Nothing, Nothing, cboEinouRuikei.SelectedValue.ToString)

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

                '【処理詳細仕様 4-1】プレプリント調査票作成処理

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    Try
                        db.BeginTrans()

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

            '【処理詳細仕様 4-2】完了メッセージ表示
            ShowMessage(resultList)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
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

        '【処理詳細仕様 4-1 ①】プレプリント調査票作成先の本年電子調査票存在チェック
        If ExistsNextChosahyo(db, key.Item1.censusNo) Then

            '「調査票を作成しなかった」という結果を保持する
            resultList.Add(ValueTuple.Create(False, key.Item1.censusNo))

            Return
        End If

        '【処理詳細仕様 4-1 ②】前年調査票取得処理
        CreateNextChosahyo(db, key.Item1.censusNo)

        '「調査票を作成した」という結果を保持する
        resultList.Add(ValueTuple.Create(True, key.Item1.censusNo))

    End Sub

    ''' <summary>
    ''' 画面で指定された調査年の「翌年」の調査票が存在するか確認する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <remarks></remarks>
    Private Function ExistsNextChosahyo(ByVal db As DBAccess, ByVal censusNo As String) As Boolean

        Dim table = DAOChosahyo.GetChosahyo(db,
                                            CStr(CInt(_chosaNen) + 1),
                                            censusNo,
                                            New DAOChosahyo.KotenKey(_kyoku, _jimusho, _kyoten))

        Return If(0 < table.Rows.Count, True, False)

    End Function

    ''' <summary>
    ''' 画面で指定された調査年の「翌年」の調査票を作成する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <remarks></remarks>
    Private Sub CreateNextChosahyo(ByVal db As DBAccess, ByVal censusNo As String)

        '画面で指定された調査年の調査票を取得する
        Dim chosahyo = GetNowChosahyo(db, censusNo)

        '今年の調査票を翌年の調査票に変換する
        ExchangeNowChosahyoToNextChosahyo(chosahyo, db, censusNo)
        '>>>2022/01/27
        '調査票データ追加
        DAOChosahyo.InsertChosahyoTable(db,
                                        New DAOChosahyo.PrimaryKey(CStr(CInt(_chosaNen) + 1), censusNo),
                                        New DAOChosahyo.KotenKey(_kyoku, _jimusho, _kyoten),
                                        chosahyo, Nothing, Nothing)
        'DAOChosahyo.InsertChosahyoTable(db,
        '                                New DAOChosahyo.PrimaryKey(CStr(CInt(_chosaNen) + 1), censusNo),
        '                                New DAOChosahyo.KotenKey(_kyoku, _jimusho, _kyoten),
        '                                chosahyo)
        '<<<2022/01/27
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

    ''' <summary>
    ''' 今年の調査票を翌年の調査票に変換する
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <remarks></remarks>
    Private Sub ExchangeNowChosahyoToNextChosahyo(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目), ByVal db As DBAccess, ByVal censusNo As String)

        '集計論理に従った処理を実装する
        ExecuteAggregateLogic(chosahyoItemList, db, censusNo)

        'ブランクが設定されている調査票項目をNothingに置換する
        ReplaceBlankToNothing(chosahyoItemList)

        '可変テーブル内の値がNothingのデータを削除しDBに書き込まないようにする
        RemoveNothingDataForVariableTable(chosahyoItemList)

    End Sub

#Region "集計論理実行"

    ''' <summary>
    ''' 集計論理に従った処理を実施する
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteAggregateLogic(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目), ByVal db As DBAccess, ByVal censusNo As String)

        '一部の項目を集計論理に則った値に置換する
        ReplaceTargetItemToSomething(chosahyoItemList, db, censusNo)

        'コピー不要な調査票項目をNothingに置換する
        ReplaceNotCopyItemToNothing(chosahyoItemList)

    End Sub

    ''' <summary>
    ''' 一部の項目を集計論理に記載されている内容に則った値に置換する
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <remarks></remarks>
    Private Sub ReplaceTargetItemToSomething(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目), ByVal db As DBAccess, ByVal censusNo As String)

        Trace.WriteLine("集計論理処理 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        Trace.WriteLine("集計論理処理 ① 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ①】
        ExecuteRonri01(chosahyoItemList)

        Trace.WriteLine("集計論理処理 ② 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ②】
        ExecuteRonri02(chosahyoItemList)

        Trace.WriteLine("集計論理処理 ③ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ③】
        ExecuteRonri03(chosahyoItemList)

        Trace.WriteLine("集計論理処理 ④ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ④】
        ExecuteRonri04(chosahyoItemList)

        Trace.WriteLine("集計論理処理 ⑤ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ⑤】
        ExecuteRonri05(chosahyoItemList)

        Trace.WriteLine("集計論理処理 ⑥ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ⑥】
        ExecuteRonri06(chosahyoItemList)

        Trace.WriteLine("集計論理処理 ⑦ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ⑦】
        ExecuteRonri07(chosahyoItemList)

        Trace.WriteLine("集計論理処理 ⑧ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ⑧】
        ExecuteRonri08(chosahyoItemList)

        Trace.WriteLine("集計論理処理 ⑨ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ⑨】
        ExecuteRonri09(chosahyoItemList, db, censusNo)

        Trace.WriteLine("集計論理処理 ⑩ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ⑩】
        'REV_002 Mod Start
        'ExecuteRonri10(chosahyoItemList)
        ExecuteRonri10(chosahyoItemList, db)
        'REV_002 Mod End

        Trace.WriteLine("集計論理処理 ⑪ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ⑪】
        ExecuteRonri11(chosahyoItemList)

        Trace.WriteLine("集計論理処理 ⑫ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ⑫】
        ExecuteRonri12(chosahyoItemList)

        Trace.WriteLine("集計論理処理 ⑬ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ⑬】
        ExecuteRonri13(chosahyoItemList)

        Trace.WriteLine("集計論理処理 ⑭ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ⑭】
        ExecuteRonri14(chosahyoItemList)

        '---REV_001_ADD START
        Trace.WriteLine("集計論理処理 ⑮ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

        '【集計論理記述表 ⑮】
        'REV_002 Mod Start
        'ExecuteRonri15(chosahyoItemList)
        ExecuteRonri15(chosahyoItemList, db)
        'REV_002 Mod End
        '---REV_001_ADD END


        '---REV_003_ADD START

        Trace.WriteLine("集計論理処理 ⑯ 開始 " & DateTime.Now.ToString("yyyy/mm/dd/hh:mm:ss:fff"))
        '【集計論理記述表 ⑯】
        ExecuteRonri16(chosahyoItemList)


        Trace.WriteLine("集計論理処理 ⑰ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))
        '【集計論理記述表 ⑰】
        ExecuteRonri17(chosahyoItemList)


        Trace.WriteLine("集計論理処理 ⑱ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))
        '【集計論理記述表 ⑱】
        ExecuteRonri18(chosahyoItemList)


        Trace.WriteLine("集計論理処理 ⑲ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))
        '【集計論理記述表 ⑲】
        ExecuteRonri19(chosahyoItemList)


        Trace.WriteLine("集計論理処理 ⑳ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))
        '【集計論理記述表 ⑳】
        ExecuteRonri20(chosahyoItemList)


        Trace.WriteLine("集計論理処理 ㉑ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))
        '【集計論理記述表 ㉑】
        ExecuteRonri21(chosahyoItemList)



        Trace.WriteLine("集計論理処理 ㉒ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))
        '【集計論理記述表㉒】
        ExecuteRonri22(chosahyoItemList)


        '2022/03/02　ADD START　連絡票No.283対応
        Trace.WriteLine("集計論理処理 ㉓ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))
        '【集計論理記述表㉓】
        ExecuteRonri23(chosahyoItemList)
        '2022/03/02　ADD END　連絡票No.283対応

        '---REV_003_ADD END

        'REV_005↓
        Trace.WriteLine("集計論理処理 ㉔ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))
        '【集計論理記述表㉔】
        ExecuteRonri24(chosahyoItemList)

        Trace.WriteLine("集計論理処理 ㉕ 開始 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))
        '【集計論理記述表㉕】
        ExecuteRonri25(chosahyoItemList)
        'REV_005↑

        Trace.WriteLine("集計論理処理 終了 " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss:fff"))

    End Sub

    ''' <summary>
    ''' 【集計論理記述表 ①】
    ''' 指標部入力シート（農生）、表紙シート（畜生）
    ''' </summary>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri01(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '指定した年の翌年をセットする
        chosahyoItemList("Q00000101").値 = CStr(CInt(_chosaNen) + 1)

        '---REV_003_ADD START

        '問い合わせを削除対象
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            chosahyoItemList("Q00000401").値 = Nothing

        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Or CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費 Then
            chosahyoItemList("Q00000301").値 = Nothing

        Else
            chosahyoItemList("Q00000201").値 = Nothing

        End If


        '2020年センサス番号がない場合⇒新バージョンの調査票での2015センサス番号（Q00002001～Q00002006）がない場合として捉える
        '2015，2020センサス番号どちらにも旧バージョン調査票の2015センサス番号（Q00001001～Q00001006）を挿入する
        '2020年センサス番号がある場合は、前年の調査票からコピーされているので（値はすでに保持されているので）、特に何もせずに問題ない
        '---問合せ番号9711対応 START 2023/02/06
        'If ((CInt(_chosaNen)) >= 2021 And Not (CommonInfo.Chosakubun = ComConst.調査区分.小麦生産費統計_個別 Or
        '                                  CommonInfo.Chosakubun = ComConst.調査区分.小麦生産費統計_組織法人 Or
        '                                  CommonInfo.Chosakubun = ComConst.調査区分.二条大麦生産費統計_個別 Or
        '                                  CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_二条大麦生産費 Or
        '                                  CommonInfo.Chosakubun = ComConst.調査区分.六条大麦生産費統計_個別 Or
        '                                  CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_六条大麦生産費 Or
        '                                  CommonInfo.Chosakubun = ComConst.調査区分.はだか麦生産費統計_個別 Or
        '                                  CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_はだか麦生産費 Or
        '                                  CommonInfo.Chosakubun = ComConst.調査区分.なたね生産費統計_個別 Or
        '                                  CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_なたね生産費)) Then



        '    chosahyoItemList("Q00002001").値 = chosahyoItemList("Q00001001").値
        '    chosahyoItemList("Q00002002").値 = chosahyoItemList("Q00001002").値
        '    chosahyoItemList("Q00002003").値 = chosahyoItemList("Q00001003").値
        '    chosahyoItemList("Q00002004").値 = chosahyoItemList("Q00001004").値
        '    chosahyoItemList("Q00002005").値 = chosahyoItemList("Q00001005").値
        '    chosahyoItemList("Q00002006").値 = chosahyoItemList("Q00001006").値


        'ElseIf ((CInt(_chosaNen)) >= 2022) Then



        '    chosahyoItemList("Q00002001").値 = chosahyoItemList("Q00001001").値
        '    chosahyoItemList("Q00002002").値 = chosahyoItemList("Q00001002").値
        '    chosahyoItemList("Q00002003").値 = chosahyoItemList("Q00001003").値
        '    chosahyoItemList("Q00002004").値 = chosahyoItemList("Q00001004").値
        '    chosahyoItemList("Q00002005").値 = chosahyoItemList("Q00001005").値
        '    chosahyoItemList("Q00002006").値 = chosahyoItemList("Q00001006").値


        'End If

        '---REV_003_ADD END

        If ((CInt(_chosaNen)) = 2021 And Not (CommonInfo.Chosakubun = ComConst.調査区分.小麦生産費統計_個別 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.小麦生産費統計_組織法人 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.二条大麦生産費統計_個別 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_二条大麦生産費 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.六条大麦生産費統計_個別 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_六条大麦生産費 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.はだか麦生産費統計_個別 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_はだか麦生産費 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.なたね生産費統計_個別 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_なたね生産費)) Then



            chosahyoItemList("Q00002001").値 = chosahyoItemList("Q00001001").値
            chosahyoItemList("Q00002002").値 = chosahyoItemList("Q00001002").値
            chosahyoItemList("Q00002003").値 = chosahyoItemList("Q00001003").値
            chosahyoItemList("Q00002004").値 = chosahyoItemList("Q00001004").値
            chosahyoItemList("Q00002005").値 = chosahyoItemList("Q00001005").値
            chosahyoItemList("Q00002006").値 = chosahyoItemList("Q00001006").値


        ElseIf ((CInt(_chosaNen)) = 2022 And (CommonInfo.Chosakubun = ComConst.調査区分.小麦生産費統計_個別 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.小麦生産費統計_組織法人 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.二条大麦生産費統計_個別 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_二条大麦生産費 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.六条大麦生産費統計_個別 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_六条大麦生産費 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.はだか麦生産費統計_個別 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_はだか麦生産費 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.なたね生産費統計_個別 Or
                                          CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_なたね生産費)) Then



            chosahyoItemList("Q00002001").値 = chosahyoItemList("Q00001001").値
            chosahyoItemList("Q00002002").値 = chosahyoItemList("Q00001002").値
            chosahyoItemList("Q00002003").値 = chosahyoItemList("Q00001003").値
            chosahyoItemList("Q00002004").値 = chosahyoItemList("Q00001004").値
            chosahyoItemList("Q00002005").値 = chosahyoItemList("Q00001005").値
            chosahyoItemList("Q00002006").値 = chosahyoItemList("Q00001006").値


        End If

        '---問合せ番号9711対応 END 2023/02/06

    End Sub

    ''' <summary>
    ''' 【集計論理記述表 ②】
    ''' 経営の概況①シート（農生）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri02(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        If Not CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then
            Return
        End If

        '収量の合計値を算出する
        Dim getTotalSyuuryou = Function(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                                   Dim total = CDec(0)

                                   '区分ごとに使用する項目が異なる
                                   Select Case CommonInfo.Chosakubun

                                       Case ComConst.調査区分.米生産費統計_個別,
                                           ComConst.調査区分.米生産費統計_組織法人

                                           total = ComUtil.Sum(ComUtil.GetZeroOrDec(value("Q02010401").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010402").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010403").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010404").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010405").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010406").値))

                                       Case ComConst.調査区分.小麦生産費統計_個別,
                                           ComConst.調査区分.六条大麦生産費統計_個別,
                                           ComConst.調査区分.はだか麦生産費統計_個別,
                                           ComConst.調査区分.そば生産費統計_個別,
                                           ComConst.調査区分.小麦生産費統計_組織法人,
                                           ComConst.調査区分.経営分析調査_六条大麦生産費,
                                           ComConst.調査区分.経営分析調査_はだか麦生産費,
                                           ComConst.調査区分.経営分析調査_そば生産費

                                           total = ComUtil.Sum(ComUtil.GetZeroOrDec(value("Q02010401").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010402").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010403").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010404").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010405").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010406").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010407").値))

                                       Case ComConst.調査区分.二条大麦生産費統計_個別,
                                           ComConst.調査区分.経営分析調査_二条大麦生産費

                                           total = ComUtil.Sum(ComUtil.GetZeroOrDec(value("Q02020407").値),
                                                               ComUtil.GetZeroOrDec(value("Q02020408").値),
                                                               ComUtil.GetZeroOrDec(value("Q02020409").値),
                                                               ComUtil.GetZeroOrDec(value("Q02020410").値),
                                                               ComUtil.GetZeroOrDec(value("Q02020411").値),
                                                               ComUtil.GetZeroOrDec(value("Q02020412").値),
                                                               ComUtil.GetZeroOrDec(value("Q02020413").値))

                                       Case ComConst.調査区分.大豆生産費統計_個別,
                                           ComConst.調査区分.大豆生産費統計_組織法人

                                           total = ComUtil.Sum(ComUtil.GetZeroOrDec(value("Q02010401").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010402").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010403").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010404").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010405").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010406").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010407").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010408").値))

                                       Case ComConst.調査区分.原料用かんしょ生産費統計_個別,
                                           ComConst.調査区分.原料用ばれいしょ生産費統計_個別,
                                           ComConst.調査区分.なたね生産費統計_個別,
                                           ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費,
                                           ComConst.調査区分.経営分析調査_なたね生産費

                                           total = ComUtil.Sum(ComUtil.GetZeroOrDec(value("Q02010401").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010402").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010403").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010404").値))

                                       Case ComConst.調査区分.てんさい生産費統計_個別,
                                           ComConst.調査区分.経営分析調査_てんさい生産費

                                           total = ComUtil.Sum(ComUtil.GetZeroOrDec(value("Q02010501").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010502").値))

                                       Case ComConst.調査区分.さとうきび生産費統計_個別,
                                           ComConst.調査区分.経営分析調査_さとうきび生産費

                                           total = ComUtil.Sum(ComUtil.GetZeroOrDec(value("Q02010501").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010502").値),
                                                               ComUtil.GetZeroOrDec(value("Q02010503").値))

                                   End Select

                                   Return total

                               End Function

        '作付面積の合計値を算出する
        Dim getTotalMenseki = Function(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                                  Dim total = CDec(0)

                                  '区分ごとに使用する項目が異なる
                                  Select Case CommonInfo.Chosakubun

                                      Case ComConst.調査区分.米生産費統計_個別,
                                          ComConst.調査区分.米生産費統計_組織法人

                                          total = ComUtil.Sum(ComUtil.GetZeroOrDec(value("Q11011001").値), ComUtil.GetZeroOrDec(value("Q11021201").値))

                                      Case ComConst.調査区分.小麦生産費統計_個別,
                                          ComConst.調査区分.二条大麦生産費統計_個別,
                                          ComConst.調査区分.六条大麦生産費統計_個別,
                                          ComConst.調査区分.はだか麦生産費統計_個別,
                                          ComConst.調査区分.そば生産費統計_個別,
                                          ComConst.調査区分.大豆生産費統計_個別,
                                          ComConst.調査区分.原料用かんしょ生産費統計_個別,
                                          ComConst.調査区分.原料用ばれいしょ生産費統計_個別,
                                          ComConst.調査区分.なたね生産費統計_個別,
                                          ComConst.調査区分.てんさい生産費統計_個別,
                                          ComConst.調査区分.小麦生産費統計_組織法人,
                                          ComConst.調査区分.経営分析調査_二条大麦生産費,
                                          ComConst.調査区分.経営分析調査_六条大麦生産費,
                                          ComConst.調査区分.経営分析調査_はだか麦生産費,
                                          ComConst.調査区分.経営分析調査_そば生産費,
                                          ComConst.調査区分.大豆生産費統計_組織法人,
                                          ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費,
                                          ComConst.調査区分.経営分析調査_なたね生産費,
                                          ComConst.調査区分.経営分析調査_てんさい生産費

                                          total = ComUtil.Sum(ComUtil.GetZeroOrDec(value("Q11011501").値), ComUtil.GetZeroOrDec(value("Q11011502").値))

                                      Case ComConst.調査区分.さとうきび生産費統計_個別,
                                          ComConst.調査区分.経営分析調査_さとうきび生産費

                                          total = ComUtil.Sum(ComUtil.GetZeroOrDec(value("Q11011701").値), ComUtil.GetZeroOrDec(value("Q11011702").値))

                                  End Select

                                  Return total

                              End Function

        '10ａ当たり収量を算出する
        Dim calc10aSyuuryou = Function(syuuryou As Decimal, menseki As Decimal)

                                  '0除算となりえるか
                                  If IsZero(menseki) Then
                                      Return GetOutputStringForDivideByZero()
                                  End If

                                  '10ａ当たり収量＝ROUND(収量÷作付面積×10,0)
                                  Return CStr(ComUtil.Round(syuuryou / menseki * 10, 0))

                              End Function

        '「直近５か年の10ａ当たり収量」の1～5年前の各番地のアドレス
        Dim last1YearAddress = If(CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_組織法人, "Q01050101", "Q01040101")
        Dim last2YearAddress = If(CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_組織法人, "Q01050201", "Q01040201")
        Dim last3YearAddress = If(CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_組織法人, "Q01050301", "Q01040301")
        Dim last4YearAddress = If(CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_組織法人, "Q01050401", "Q01040401")
        Dim last5YearAddress = If(CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_組織法人, "Q01050501", "Q01040501")

        '5年前の値には前年データの「４年前」をセットする。
        chosahyoItemList(last5YearAddress).値 = chosahyoItemList(last4YearAddress).値

        '4年前〃「３年前」〃
        chosahyoItemList(last4YearAddress).値 = chosahyoItemList(last3YearAddress).値

        '3年前〃「２年前」〃
        chosahyoItemList(last3YearAddress).値 = chosahyoItemList(last2YearAddress).値

        '2年前〃「１年前」〃
        chosahyoItemList(last2YearAddress).値 = chosahyoItemList(last1YearAddress).値

        '収量の合計値
        Dim totalSyuuryou = getTotalSyuuryou(chosahyoItemList)

        '作付面積の合計値
        Dim totalMenseki = getTotalMenseki(chosahyoItemList)

        '1年前の値を算出する
        chosahyoItemList(last1YearAddress).値 = calc10aSyuuryou(totalSyuuryou, totalMenseki)

    End Sub

    ''' <summary>
    ''' 【集計論理記述表 ③】
    ''' 経営の概況③シート（農生（原料用かんしょを除く。））
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri03(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        If Not CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then
            Return
        End If

        If CommonInfo.Chosakubun = ComConst.調査区分.原料用かんしょ生産費統計_個別 Then
            Return
        End If

        Dim addressList As New List(Of String)

        '区分ごとに項目が異なる
        Select Case CommonInfo.Chosakubun

            Case ComConst.調査区分.米生産費統計_個別

                addressList.Add("Q01110104")
                addressList.Add("Q01110105")
                addressList.Add("Q01110106")
                addressList.Add("Q01110201")
                addressList.Add("Q01110204")
                addressList.Add("Q01110205")
                addressList.Add("Q01110206")
                addressList.Add("Q01110301")
                addressList.Add("Q01110302")
                addressList.Add("Q01110303")
                addressList.Add("Q01110304")
                addressList.Add("Q01110305")
                addressList.Add("Q01110306")

            Case ComConst.調査区分.小麦生産費統計_個別,
                ComConst.調査区分.二条大麦生産費統計_個別,
                ComConst.調査区分.六条大麦生産費統計_個別,
                ComConst.調査区分.はだか麦生産費統計_個別,
                ComConst.調査区分.大豆生産費統計_個別,
                ComConst.調査区分.小麦生産費統計_組織法人,
                ComConst.調査区分.大豆生産費統計_組織法人,
                ComConst.調査区分.経営分析調査_二条大麦生産費,
                ComConst.調査区分.経営分析調査_六条大麦生産費,
                ComConst.調査区分.経営分析調査_はだか麦生産費

                addressList.Add("Q01070106")
                addressList.Add("Q01070107")
                addressList.Add("Q01070108")
                addressList.Add("Q01070201")
                addressList.Add("Q01070206")
                addressList.Add("Q01070207")
                addressList.Add("Q01070208")
                addressList.Add("Q01070301")
                addressList.Add("Q01070302")
                addressList.Add("Q01070303")
                addressList.Add("Q01070304")
                addressList.Add("Q01070305")
                addressList.Add("Q01070306")
                addressList.Add("Q01070307")
                addressList.Add("Q01070308")

            Case ComConst.調査区分.経営分析調査_そば生産費,
                ComConst.調査区分.経営分析調査_なたね生産費

                addressList.Add("Q01070106")
                addressList.Add("Q01070107")
                addressList.Add("Q01070108")
                addressList.Add("Q01070201")
                addressList.Add("Q01070206")
                addressList.Add("Q01070207")
                addressList.Add("Q01070208")
                addressList.Add("Q01070301")
                addressList.Add("Q01070303")
                addressList.Add("Q01070304")
                addressList.Add("Q01070305")
                addressList.Add("Q01070306")
                addressList.Add("Q01070307")
                addressList.Add("Q01070308")

            Case ComConst.調査区分.そば生産費統計_個別,
                ComConst.調査区分.なたね生産費統計_個別,
                ComConst.調査区分.てんさい生産費統計_個別,
                ComConst.調査区分.経営分析調査_てんさい生産費

                addressList.Add("Q01070105")
                addressList.Add("Q01070106")
                addressList.Add("Q01070107")
                addressList.Add("Q01070201")
                addressList.Add("Q01070205")
                addressList.Add("Q01070206")
                addressList.Add("Q01070207")
                addressList.Add("Q01070301")
                addressList.Add("Q01070302")
                addressList.Add("Q01070303")
                addressList.Add("Q01070304")
                addressList.Add("Q01070305")
                addressList.Add("Q01070306")
                addressList.Add("Q01070307")

            Case ComConst.調査区分.原料用ばれいしょ生産費統計_個別,
                ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費

                addressList.Add("Q01060105")
                addressList.Add("Q01060106")
                addressList.Add("Q01060107")
                addressList.Add("Q01060201")
                addressList.Add("Q01060205")
                addressList.Add("Q01060206")
                addressList.Add("Q01060207")
                addressList.Add("Q01060301")
                addressList.Add("Q01060302")
                addressList.Add("Q01060303")
                addressList.Add("Q01060304")
                addressList.Add("Q01060305")
                addressList.Add("Q01060306")
                addressList.Add("Q01060307")

            Case ComConst.調査区分.さとうきび生産費統計_個別,
                ComConst.調査区分.経営分析調査_さとうきび生産費

                addressList.Add("Q01070201")
                addressList.Add("Q01070301")

            Case ComConst.調査区分.米生産費統計_組織法人

                addressList.Add("Q01130104")
                addressList.Add("Q01130105")
                addressList.Add("Q01130106")
                addressList.Add("Q01130201")
                addressList.Add("Q01130204")
                addressList.Add("Q01130205")
                addressList.Add("Q01130206")
                addressList.Add("Q01130301")
                addressList.Add("Q01130302")
                addressList.Add("Q01130303")
                addressList.Add("Q01130304")
                addressList.Add("Q01130305")
                addressList.Add("Q01130306")

        End Select

        '「制度受取金等の状況」を削除する
        For Each address In addressList
            chosahyoItemList(address).値 = Nothing
        Next

    End Sub


    ''' <summary>
    ''' 【集計論理記述表 ④】
    ''' 【５】土地改良及び水利費・【６】借入金シート（農生）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri04(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        If Not CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then
            Return
        End If

        Dim targetItem = If(CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_個別 OrElse CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_組織法人, "Q06010104", "Q06010103")

        '長期借入金の内訳の「資金名（用途）」欄以外の項目を取得する
        Dim query = chosahyoItemList.Where(Function(row) (row.Key.StartsWith("Q05") OrElse row.Key.StartsWith("Q06")) AndAlso
                                                         Not row.Key.StartsWith(targetItem)).ToList()

        '長期借入金の内訳の「資金名（用途）」欄以外を削除する
        For Each item In query
            chosahyoItemList(item.Key).値 = Nothing
        Next

    End Sub

    ''' <summary>
    ''' 【集計論理記述表 ⑤】
    ''' 【４】物件税～【６】借入金シート（牛乳生産費）、
    ''' 【４】物件税～【７】出荷経費シート（子牛生産費、育成牛・肥育牛生産費）、
    ''' 【４】物件税～【７】出荷に要した経費（肥育豚生産費）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri05(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        If Not CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費 Then
            Return
        End If

        '牛乳はQ07関連の項目が対象外なのでQ06を割り当てる
        Dim target = If(CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費, "Q06", "Q07")

        '「借入金種類」欄以外の項目を取得する
        Dim query = chosahyoItemList.Where(Function(row) (row.Key.StartsWith("Q04") OrElse row.Key.StartsWith("Q05") OrElse row.Key.StartsWith("Q06") OrElse row.Key.StartsWith(target)) AndAlso
                                                         Not row.Key.StartsWith("Q06010102")).ToList()

        '「借入金種類」欄以外を削除する
        For Each item In query
            chosahyoItemList(item.Key).値 = Nothing
        Next

    End Sub

    ''' <summary>
    ''' 【集計論理記述表 ⑥】
    ''' 建物及び構築物シート（農生）、建物シート（畜生）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri06(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（個人）または（法人）なら終了
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            Return
        End If

        Dim 延べ面積又は施設数Address = String.Empty
        Dim 取得価額Address = String.Empty
        Dim 農業経営基盤強化準備金Address = String.Empty
        Dim 修繕費Address = String.Empty
        Dim 保険料Address = String.Empty
        Dim 異動コードAddress = String.Empty
        Dim 異動に伴う発生金額Address = String.Empty
        Dim 部分取り壊しをした面積割合Address = String.Empty

        If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then

            取得価額Address = "Q07010801"
            農業経営基盤強化準備金Address = "Q07010901"
            修繕費Address = "Q07011001"
            保険料Address = "Q07011101"
            異動コードAddress = "Q07011401"
            異動に伴う発生金額Address = "Q07011501"
            部分取り壊しをした面積割合Address = "Q07011601"

        ElseIf CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費 Then

            If CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse
                CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then

                延べ面積又は施設数Address = "Q07010401"
                取得価額Address = "Q07010901"
                農業経営基盤強化準備金Address = "Q07011001"
                修繕費Address = "Q07011101"
                保険料Address = "Q07011201"
                異動コードAddress = "Q07011501"
                異動に伴う発生金額Address = "Q07011601"
                部分取り壊しをした面積割合Address = "Q07011701"

            ElseIf CommonInfo.Chosakubun = ComConst.調査区分.肥育豚生産費統計_個別 OrElse
                CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_肥育豚生産費 Then

                延べ面積又は施設数Address = "Q08010401"
                取得価額Address = "Q08010901"
                農業経営基盤強化準備金Address = "Q08011001"
                修繕費Address = "Q08011101"
                保険料Address = "Q08011201"
                異動コードAddress = "Q08011401"
                異動に伴う発生金額Address = "Q08011501"
                部分取り壊しをした面積割合Address = "Q08011601"

            Else

                延べ面積又は施設数Address = "Q08010401"
                取得価額Address = "Q08010901"
                農業経営基盤強化準備金Address = "Q08011001"
                修繕費Address = "Q08011101"
                保険料Address = "Q08011201"
                異動コードAddress = "Q08011501"
                異動に伴う発生金額Address = "Q08011601"
                部分取り壊しをした面積割合Address = "Q08011701"

            End If
        End If

        '構築物・建物シート内の可変項目を取得する
        Dim query1 = chosahyoItemList.Where(Function(row) row.Key.StartsWith(異動コードAddress.Substring(0, 3)) AndAlso
                                                          row.Key.IndexOf("_") = 9).ToList()

        If Not query1.Any() Then
            Return
        End If

        '「修繕費」、「保険料」欄を削除する
        Dim deleteFixdItem = Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                                 Dim query = query1.Where(Function(row) row.Key.StartsWith(修繕費Address) OrElse
                                                                        row.Key.StartsWith(保険料Address)).ToList()

                                 '削除する
                                 query.ForEach(Sub(row) value(row.Key).値 = Nothing)

                             End Sub

        '「修繕費」、「保険料」欄を削除する
        deleteFixdItem(chosahyoItemList)

        '「異動コード」が入力されている項目を抽出する
        Dim query2 = query1.Where(Function(row) row.Key.StartsWith(異動コードAddress) AndAlso Not String.IsNullOrWhiteSpace(row.Value.値)).ToList()

        If Not query2.Any() Then
            Return
        End If

        'まだ削除してない可変項目をすべて取得する
        Dim query3 = query1.Where(Function(row) Not row.Key.StartsWith(修繕費Address) AndAlso
                                                Not row.Key.StartsWith(保険料Address)).ToList()

        '「異動コード」が「1(取り壊し)」か「3(売却)」の資産を削除する
        Dim deleteTransferCode1Or3 = Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                                         '「異動コード」が[1]取り壊し、[3]売却の資産を取得する
                                         Dim getTransferCode1Or3RowList = Function()

                                                                              '「異動コード」が「1(取り壊し)」か「3(売却)」のレコードを抽出する
                                                                              Dim query = query2.Where(Function(row) row.Value.値 IsNot Nothing AndAlso
                                                                                                                     (row.Value.値.Contains("[1]取り壊し") OrElse
                                                                                                                      row.Value.値.Contains("1：取り壊し") OrElse
                                                                                                                      row.Value.値.Contains("[3]売　却") OrElse
                                                                                                                      row.Value.値.Contains("3：売　却"))).
                                                                                                 Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key)).
                                                                                                 ToList()

                                                                              Dim itemListList As New List(Of List(Of KeyValuePair(Of String, DAOChosahyo.調査票項目)))

                                                                              '「異動コード」が「1(取り壊し)」か「3(売却)」となっている項目と同じ枝番を持つ項目を抽出する
                                                                              query.ForEach(Sub(row) itemListList.Add(query3.Where(Function(row2) row2.Key.EndsWith(row)).ToList()))

                                                                              Return itemListList

                                                                          End Function

                                         '「異動コード」が「1(取り壊し)」か「3(売却)」の資産を取得する
                                         Dim deleteItemListList = getTransferCode1Or3RowList()

                                         If Not deleteItemListList.Any() Then
                                             Return
                                         End If

                                         '削除対象項目の数ぶんループする
                                         For Each deleteItemList In deleteItemListList

                                             '削除対象項目の数ぶんループする
                                             For Each deleteItem In deleteItemList

                                                 '「異動コード」が「1(取り壊し)」か「3(売却)」の資産を削除する
                                                 value(deleteItem.Key).値 = Nothing
                                             Next
                                         Next

                                     End Sub

        '「異動コード」が「2(部分取り壊し)」の資産を更新する
        Dim updateTransferCode2 = Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                                      '「異動コード」が「2(部分取り壊し)」の資産を取得する
                                      Dim getTransferCode2RowList = Function()

                                                                        '「異動コード」が「2(部分取り壊し)」のレコードを抽出する
                                                                        Dim query = query2.Where(Function(row) row.Value.値 IsNot Nothing AndAlso
                                                                                                               (row.Value.値.Contains("[2]部分取り壊し") OrElse
                                                                                                                row.Value.値.Contains("2：部分取り壊し"))).
                                                                                           Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key)).
                                                                                           ToList()

                                                                        Dim itemListList As New List(Of List(Of KeyValuePair(Of String, DAOChosahyo.調査票項目)))

                                                                        '「異動コード」が「2(部分取り壊し)」となっている項目と同じ枝番を持つ項目を抽出する
                                                                        query.ForEach(Sub(row) itemListList.Add(query3.Where(Function(row2) row2.Key.EndsWith(row)).ToList()))

                                                                        Return itemListList

                                                                    End Function

                                      '「異動コード」が[2]部分取り壊しの資産を取得する
                                      Dim updateItemListList = getTransferCode2RowList()

                                      If Not updateItemListList.Any() Then
                                          Return
                                      End If

                                      '「異動コード」、「異動に伴う発生金額」及び「部分取り壊しをした面積割合」欄のデータを削除する
                                      Dim deleteItem = Sub(itemList As List(Of KeyValuePair(Of String, DAOChosahyo.調査票項目)))

                                                           Dim 異動コード = itemList.Where(Function(row) row.Key.StartsWith(異動コードAddress)).FirstOrDefault()

                                                           If Not IsNothing(異動コード.Value) Then
                                                               value(異動コード.Key).値 = Nothing
                                                           End If

                                                           Dim 異動に伴う発生金額 = itemList.Where(Function(row) row.Key.StartsWith(異動に伴う発生金額Address)).FirstOrDefault()

                                                           If Not IsNothing(異動に伴う発生金額.Value) Then
                                                               value(異動に伴う発生金額.Key).値 = Nothing
                                                           End If

                                                           Dim 部分取り壊しをした面積割合 = itemList.Where(Function(row) row.Key.StartsWith(部分取り壊しをした面積割合Address)).FirstOrDefault()

                                                           If Not IsNothing(部分取り壊しをした面積割合.Value) Then
                                                               value(部分取り壊しをした面積割合.Key).値 = Nothing
                                                           End If

                                                       End Sub

                                      If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then

                                          '取得価額を算出する
                                          Dim calcSyutokuKagaku = Sub(itemList As List(Of KeyValuePair(Of String, DAOChosahyo.調査票項目)))

                                                                      Dim 取得価額 = itemList.Where(Function(row) row.Key.StartsWith(取得価額Address)).FirstOrDefault()

                                                                      If IsNothing(取得価額.Value) OrElse IsNothing(取得価額.Value.値) Then
                                                                          Return
                                                                      End If

                                                                      Dim 部分取り壊しをした面積割合 = itemList.Where(Function(row) row.Key.StartsWith(部分取り壊しをした面積割合Address)).FirstOrDefault()

                                                                      If IsNothing(部分取り壊しをした面積割合.Value) OrElse IsNothing(部分取り壊しをした面積割合.Value.値) Then
                                                                          value(取得価額.Key).値 = Nothing
                                                                          Return
                                                                      End If

                                                                      '「取得価額」＝「取得価額」×(1－「部分取り壊しをした面積割合」÷100)（小数点以下第１位を四捨五入）
                                                                      value(取得価額.Key).値 = CStr(ComUtil.Round(ComUtil.GetZeroOrDec(取得価額.Value.値) *
                                                                                                                  (1 - (ComUtil.GetZeroOrDec(部分取り壊しをした面積割合.Value.値)) / 100), 0))

                                                                  End Sub

                                          '農業経営基盤強化準備金を算出する
                                          Dim calcJunbiKin = Sub(itemList As List(Of KeyValuePair(Of String, DAOChosahyo.調査票項目)))

                                                                 Dim 農業経営基盤強化準備金 = itemList.Where(Function(row) row.Key.StartsWith(農業経営基盤強化準備金Address)).FirstOrDefault()

                                                                 If IsNothing(農業経営基盤強化準備金.Value) OrElse IsNothing(農業経営基盤強化準備金.Value.値) Then
                                                                     Return
                                                                 End If

                                                                 Dim 部分取り壊しをした面積割合 = itemList.Where(Function(row) row.Key.StartsWith(部分取り壊しをした面積割合Address)).FirstOrDefault()

                                                                 If IsNothing(部分取り壊しをした面積割合.Value) OrElse IsNothing(部分取り壊しをした面積割合.Value.値) Then
                                                                     value(農業経営基盤強化準備金.Key).値 = Nothing
                                                                     Return
                                                                 End If

                                                                 '「農業経営基盤強化準備金」＝「農業経営基盤強化準備金」×(1－「部分取り壊しをした面積割合」÷100)（小数点以下第１位を四捨五入）
                                                                 value(農業経営基盤強化準備金.Key).値 = CStr(ComUtil.Round(ComUtil.GetZeroOrDec(農業経営基盤強化準備金.Value.値) *
                                                                                                                           (1 - (ComUtil.GetZeroOrDec(部分取り壊しをした面積割合.Value.値)) / 100), 0))

                                                             End Sub

                                          '更新する項目のリストの数ぶんループする
                                          For Each updateItemList In updateItemListList

                                              '取得価額を算出する
                                              calcSyutokuKagaku(updateItemList)

                                              '農業経営基盤強化準備金を算出する
                                              calcJunbiKin(updateItemList)

                                              '「異動コード」、「異動に伴う発生金額」及び「部分取り壊しをした面積割合」欄のデータを削除する
                                              deleteItem(updateItemList)

                                          Next

                                      ElseIf CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費 Then

                                          '取得価額を算出する
                                          Dim calcSyutokuKagaku = Sub(itemList As List(Of KeyValuePair(Of String, DAOChosahyo.調査票項目)))

                                                                      Dim 取得価額 = itemList.Where(Function(row) row.Key.StartsWith(取得価額Address)).FirstOrDefault()

                                                                      If IsNothing(取得価額.Value) OrElse IsNothing(取得価額.Value.値) Then
                                                                          Return
                                                                      End If

                                                                      Dim 延べ面積又は施設数 = itemList.Where(Function(row) row.Key.StartsWith(延べ面積又は施設数Address)).FirstOrDefault()

                                                                      If IsNothing(延べ面積又は施設数.Value) OrElse IsNothing(延べ面積又は施設数.Value.値) Then
                                                                          value(取得価額.Key).値 = Nothing
                                                                          Return
                                                                      End If

                                                                      Dim 部分取り壊しをした面積割合 = itemList.Where(Function(row) row.Key.StartsWith(部分取り壊しをした面積割合Address)).FirstOrDefault()

                                                                      If IsNothing(部分取り壊しをした面積割合.Value) OrElse IsNothing(部分取り壊しをした面積割合.Value.値) Then
                                                                          value(取得価額.Key).値 = Nothing
                                                                          Return
                                                                      End If

                                                                      '「取得価額」＝「取得価額」×(「延べ面積（㎡）又は施設数（基）」－「部分取り壊しを行った面積（㎡）」）÷「延べ面積（㎡）又は施設数（基）」（小数点以下第１位を四捨五入）
                                                                      value(取得価額.Key).値 =
                                                                          CStr(ComUtil.Round(ComUtil.GetZeroOrDec(取得価額.Value.値) *
                                                                                             (ComUtil.GetZeroOrDec(延べ面積又は施設数.Value.値) - ComUtil.GetZeroOrDec(部分取り壊しをした面積割合.Value.値)) /
                                                                                             ComUtil.GetZeroOrDec(延べ面積又は施設数.Value.値),
                                                                                             0))

                                                                  End Sub

                                          '農業経営基盤強化準備金を算出する
                                          Dim calcJunbiKin = Sub(itemList As List(Of KeyValuePair(Of String, DAOChosahyo.調査票項目)))

                                                                 Dim 農業経営基盤強化準備金 = itemList.Where(Function(row) row.Key.StartsWith(農業経営基盤強化準備金Address)).FirstOrDefault()

                                                                 If IsNothing(農業経営基盤強化準備金.Value) OrElse IsNothing(農業経営基盤強化準備金.Value.値) Then
                                                                     Return
                                                                 End If

                                                                 Dim 延べ面積又は施設数 = itemList.Where(Function(row) row.Key.StartsWith(延べ面積又は施設数Address)).FirstOrDefault()

                                                                 If IsNothing(延べ面積又は施設数.Value) OrElse IsNothing(延べ面積又は施設数.Value.値) Then
                                                                     value(農業経営基盤強化準備金.Key).値 = Nothing
                                                                     Return
                                                                 End If

                                                                 Dim 部分取り壊しをした面積割合 = itemList.Where(Function(row) row.Key.StartsWith(部分取り壊しをした面積割合Address)).FirstOrDefault()

                                                                 If IsNothing(部分取り壊しをした面積割合.Value) OrElse IsNothing(部分取り壊しをした面積割合.Value.値) Then
                                                                     value(農業経営基盤強化準備金.Key).値 = Nothing
                                                                     Return
                                                                 End If

                                                                 '0除算となりえるか
                                                                 If IsZero(延べ面積又は施設数.Value.値) Then
                                                                     value(農業経営基盤強化準備金.Key).値 = GetOutputStringForDivideByZero()
                                                                     Return
                                                                 End If

                                                                 '「農業経営基盤強化準備金」＝「農業経営基盤強化準備金」×(「延べ面積（㎡）又は施設数（基）」－「部分取り壊しを行った面積（㎡）」）÷「延べ面積（㎡）又は施設数（基）」（小数点以下第１位を四捨五入）
                                                                 value(農業経営基盤強化準備金.Key).値 =
                                                                     CStr(ComUtil.Round(ComUtil.GetZeroOrDec(農業経営基盤強化準備金.Value.値) *
                                                                                        (ComUtil.GetZeroOrDec(延べ面積又は施設数.Value.値) - ComUtil.GetZeroOrDec(部分取り壊しをした面積割合.Value.値)) /
                                                                                        ComUtil.GetZeroOrDec(延べ面積又は施設数.Value.値),
                                                                                        0))

                                                             End Sub

                                          '延べ面積（㎡）又は施設数（基）を算出する
                                          Dim calcMensekiOrShisetsu = Sub(itemList As List(Of KeyValuePair(Of String, DAOChosahyo.調査票項目)))

                                                                          Dim 延べ面積又は施設数 = itemList.Where(Function(row) row.Key.StartsWith(延べ面積又は施設数Address)).FirstOrDefault()

                                                                          If IsNothing(延べ面積又は施設数.Value) OrElse IsNothing(延べ面積又は施設数.Value.値) Then
                                                                              Return
                                                                          End If

                                                                          Dim 部分取り壊しをした面積割合 = itemList.Where(Function(row) row.Key.StartsWith(部分取り壊しをした面積割合Address)).FirstOrDefault()

                                                                          If IsNothing(部分取り壊しをした面積割合.Value) OrElse IsNothing(部分取り壊しをした面積割合.Value.値) Then
                                                                              value(延べ面積又は施設数.Key).値 = Nothing
                                                                              Return
                                                                          End If

                                                                          '「延べ面積（㎡）又は施設数（基）」＝「延べ面積（㎡）又は施設数（基）」－「部分取り壊しを行った面積（㎡）」
                                                                          value(延べ面積又は施設数.Key).値 =
                                                                              CStr(ComUtil.Round(ComUtil.GetZeroOrDec(延べ面積又は施設数.Value.値) - ComUtil.GetZeroOrDec(部分取り壊しをした面積割合.Value.値), 0))

                                                                      End Sub

                                          '更新する項目のリストの数ぶんループする
                                          For Each updateItemList In updateItemListList

                                              '取得価額を算出する
                                              calcSyutokuKagaku(updateItemList)

                                              '農業経営基盤強化準備金を算出する
                                              calcJunbiKin(updateItemList)

                                              '延べ面積（㎡）又は施設数（基）を算出する
                                              calcMensekiOrShisetsu(updateItemList)

                                              '「異動コード」、「異動に伴う発生金額」及び「部分取り壊しをした面積割合」欄のデータを削除する
                                              deleteItem(updateItemList)

                                          Next

                                      End If

                                  End Sub

        '「異動コード」が「1(取り壊し)」か「3(売却)」の資産を削除する
        deleteTransferCode1Or3(chosahyoItemList)

        '「異動コード」が「2(部分取り壊し)」の資産を更新する
        updateTransferCode2(chosahyoItemList)

        '歯抜けのレコードを上に詰める
        chosahyoItemList = ComUtil.Chosahyo.MoveUpMissingRow(異動コードAddress, chosahyoItemList)

    End Sub

    ''' <summary>
    ''' 【集計論理記述表 ⑦】
    ''' 自動車シート（共通）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri07(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（個人）または（法人）なら終了
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            Return
        End If

        Dim 修繕費Address = String.Empty
        Dim 自動車保険料Address = String.Empty
        Dim 自動車_軽自動車税Address = String.Empty
        Dim 自動車重量税Address = String.Empty
        Dim 自賠責保険Address = String.Empty
        Dim 異動コードAddress = String.Empty
        Dim 売却金額Address = String.Empty

        If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then

            修繕費Address = "Q08010901"
            自動車保険料Address = "Q08011001"
            自動車_軽自動車税Address = "Q08011101"
            自動車重量税Address = "Q08011201"
            自賠責保険Address = "Q08011301"
            異動コードAddress = "Q08011601"
            売却金額Address = "Q08011701"

        ElseIf CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費 Then

            If CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse
                CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then

                修繕費Address = "Q08010901"
                自動車保険料Address = "Q08011001"
                自動車_軽自動車税Address = "Q08011101"
                自動車重量税Address = "Q08011201"
                自賠責保険Address = "Q08011301"
                異動コードAddress = "Q08011601"
                売却金額Address = "Q08011701"

            ElseIf CommonInfo.Chosakubun = ComConst.調査区分.肥育豚生産費統計_個別 OrElse
                CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_肥育豚生産費 Then

                修繕費Address = "Q09010901"
                自動車保険料Address = "Q09011001"
                自動車_軽自動車税Address = "Q09011101"
                自動車重量税Address = "Q09011201"
                自賠責保険Address = "Q09011301"
                異動コードAddress = "Q09011501"
                売却金額Address = "Q09011601"

            Else

                修繕費Address = "Q09010901"
                自動車保険料Address = "Q09011001"
                自動車_軽自動車税Address = "Q09011101"
                自動車重量税Address = "Q09011201"
                自賠責保険Address = "Q09011301"
                異動コードAddress = "Q09011601"
                売却金額Address = "Q09011701"

            End If
        End If

        '自動車シート内の可変項目を取得する
        Dim query1 = chosahyoItemList.Where(Function(row) row.Key.StartsWith(異動コードAddress.Substring(0, 3)) AndAlso
                                                          row.Key.IndexOf("_") = 9).ToList()

        If Not query1.Any() Then
            Return
        End If

        '「異動コード」が入力されている項目を抽出する
        Dim query2 = query1.Where(Function(row) row.Key.StartsWith(異動コードAddress) AndAlso Not String.IsNullOrWhiteSpace(row.Value.値)).ToList()

        '「修繕費」、「自動車保険料」、「自動車・軽自動車税」、「自動車重量税」、「自賠責保険」、「異動コード」及び「売却金額」欄を削除する
        Dim deleteFixdItem = Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                                 Dim query = query1.Where(Function(row) row.Key.StartsWith(修繕費Address) OrElse
                                                                        row.Key.StartsWith(自動車保険料Address) OrElse
                                                                        row.Key.StartsWith(自動車_軽自動車税Address) OrElse
                                                                        row.Key.StartsWith(自動車重量税Address) OrElse
                                                                        row.Key.StartsWith(自賠責保険Address) OrElse
                                                                        row.Key.StartsWith(異動コードAddress) OrElse
                                                                        row.Key.StartsWith(売却金額Address)).ToList()

                                 '削除する
                                 query.ForEach(Sub(row) value(row.Key).値 = Nothing)

                             End Sub

        '「修繕費」、「自動車保険料」、「自動車・軽自動車税」、「自動車重量税」、「自賠責保険」、「異動コード」及び「売却金額」欄を削除する
        deleteFixdItem(chosahyoItemList)

        If Not query2.Any() Then
            Return
        End If

        'まだ削除してない可変項目をすべて取得する
        Dim query3 = query1.Where(Function(row) Not row.Key.StartsWith(修繕費Address) AndAlso
                                                Not row.Key.StartsWith(自動車保険料Address) AndAlso
                                                Not row.Key.StartsWith(自動車_軽自動車税Address) AndAlso
                                                Not row.Key.StartsWith(自動車重量税Address) AndAlso
                                                Not row.Key.StartsWith(自賠責保険Address) AndAlso
                                                Not row.Key.StartsWith(異動コードAddress) AndAlso
                                                Not row.Key.StartsWith(売却金額Address)).ToList()

        If Not query3.Any() Then
            Return
        End If

        '「異動コード」が入力されている項目の数ぶんループする
        For Each item In query2

            '削除対象項目を取得する
            Dim query4 = query3.Where(Function(row) row.Key.EndsWith(ComUtil.Chosahyo.GetEdaNo(item.Key))).ToList()

            '「異動コード」の入力のある資産を削除する
            For Each deleteItem In query4
                chosahyoItemList(deleteItem.Key).値 = Nothing
            Next
        Next

        '歯抜けのレコードを上に詰める
        chosahyoItemList = ComUtil.Chosahyo.MoveUpMissingRow(異動コードAddress, chosahyoItemList)

    End Sub

    ''' <summary>
    ''' 【集計論理記述表 ⑧】
    ''' 農業機械シート（共通）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri08(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（個人）または（法人）なら終了
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            Return
        End If


        'REV 003 START 自動車重量税、自賠責保険をセットしない項目として追加
        Dim 修繕費Address = String.Empty
        Dim 保険料Address = String.Empty
        Dim 軽自動車税Address = String.Empty
        Dim 自動車重量税Address = String.Empty　'ADD REV003
        Dim 自賠責保険Address = String.Empty 'ADD REV003
        Dim 異動コードAddress = String.Empty
        Dim 売却金額Address = String.Empty

        If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then



            If CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_個別 OrElse
                CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_組織法人 Then
                修繕費Address = "Q09011001"
                保険料Address = "Q09011101"
                軽自動車税Address = "Q09011201"
                自動車重量税Address = "Q09013101"
                自賠責保険Address = "Q09013201"
                異動コードAddress = "Q09011501"
                売却金額Address = "Q09011601"
            Else
                修繕費Address = "Q09011001"
                保険料Address = "Q09011101"
                軽自動車税Address = "Q09011201"
                自動車重量税Address = "Q09012101"
                自賠責保険Address = "Q09012201"
                異動コードAddress = "Q09011501"
                売却金額Address = "Q09011601"

            End If


        ElseIf CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費 Then

            If CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse
                CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then

                修繕費Address = "Q09011001"
                保険料Address = "Q09011101"
                軽自動車税Address = "Q09011201"
                自動車重量税Address = "Q09013001" 'ADD REV003
                自賠責保険Address = "Q09013101" 'ADD REV003
                異動コードAddress = "Q09011501"
                売却金額Address = "Q09011601"

            ElseIf CommonInfo.Chosakubun = ComConst.調査区分.肥育豚生産費統計_個別 OrElse
                CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_肥育豚生産費 Then

                修繕費Address = "Q10011001"
                保険料Address = "Q10011101"
                軽自動車税Address = "Q10011201"
                自動車重量税Address = "Q10013101" 'ADD REV003
                自賠責保険Address = "Q10013201" 'ADD REV003
                異動コードAddress = "Q10011401"
                売却金額Address = "Q10011501"

            Else

                修繕費Address = "Q10011001"
                保険料Address = "Q10011101"
                軽自動車税Address = "Q10011201"
                自動車重量税Address = "Q10013101" 'ADD REV003
                自賠責保険Address = "Q10013201" 'ADD REV003
                異動コードAddress = "Q10011501"
                売却金額Address = "Q10011601"

            End If
        End If

        '農業機械シート内の可変項目を取得する
        Dim query1 = chosahyoItemList.Where(Function(row) row.Key.StartsWith(異動コードAddress.Substring(0, 3)) AndAlso
                                                          row.Key.IndexOf("_") = 9).ToList()

        If Not query1.Any() Then
            Return
        End If

        '「異動コード」が入力されている項目を抽出する
        Dim query2 = query1.Where(Function(row) row.Key.StartsWith(異動コードAddress) AndAlso Not String.IsNullOrWhiteSpace(row.Value.値)).ToList()

        '「修繕費」、「保険料」、「軽自動車税」、「自動車重量税」、「自賠責保険」、「異動コード」及び「売却金額」欄を削除する
        Dim deleteFixdItem = Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                                 Dim query = query1.Where(Function(row) row.Key.StartsWith(修繕費Address) OrElse
                                                                        row.Key.StartsWith(保険料Address) OrElse
                                                                        row.Key.StartsWith(軽自動車税Address) OrElse
                                                                        row.Key.StartsWith(自動車重量税Address) OrElse 'ADD REV003
                                                                        row.Key.StartsWith(自賠責保険Address) OrElse 'ADD REV003
                                                                        row.Key.StartsWith(異動コードAddress) OrElse
                                                                        row.Key.StartsWith(売却金額Address)).ToList()

                                 '削除する
                                 query.ForEach(Sub(row) value(row.Key).値 = Nothing)

                             End Sub

        '「修繕費」、「保険料」、「軽自動車税」、「自動車重量税」、「自賠責保険」、「異動コード」及び「売却金額」欄を削除する
        deleteFixdItem(chosahyoItemList)

        If Not query2.Any() Then
            Return
        End If

        'まだ削除してない可変項目をすべて取得する
        Dim query3 = query1.Where(Function(row) Not row.Key.StartsWith(修繕費Address) AndAlso
                                                Not row.Key.StartsWith(保険料Address) AndAlso
                                                Not row.Key.StartsWith(軽自動車税Address) AndAlso
                                                Not row.Key.StartsWith(自動車重量税Address) AndAlso 'ADD REV003
                                                Not row.Key.StartsWith(自賠責保険Address) AndAlso 'ADD REV003
                                                Not row.Key.StartsWith(異動コードAddress) AndAlso
                                                Not row.Key.StartsWith(売却金額Address)).ToList()

        If Not query3.Any() Then
            Return
        End If

        '「異動コード」が入力されている項目の数ぶんループする
        For Each item In query2

            '削除対象項目を取得する
            Dim query4 = query3.Where(Function(row) row.Key.EndsWith(ComUtil.Chosahyo.GetEdaNo(item.Key))).ToList()

            '「異動コード」の入力のある資産を削除する
            For Each deleteItem In query4
                chosahyoItemList(deleteItem.Key).値 = Nothing
            Next
        Next

        '歯抜けのレコードを上に詰める
        chosahyoItemList = ComUtil.Chosahyo.MoveUpMissingRow(異動コードAddress, chosahyoItemList)

        'REV 003 END 自動車重量税、自賠責保険の追加

    End Sub

    ''' <summary>
    ''' 【集計論理記述表 ⑨】
    ''' 【11】搾乳牛所有状況２シート（牛乳生産費）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri09(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目), ByVal db As DBAccess, ByVal censusNo As String)

        If chosahyoItemList Is Nothing Then
            Return
        End If

        If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then
            Return
        End If

        If Not CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 AndAlso
            Not CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then

            Return

        End If

        'REV_006 ADD STR DAIKO
        '各バージョン区分に対応する項番を取得する。
        Dim 評価額_売却額等項番 = If(ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2015, "Q11023101", "Q11023201")
        'REV_006 ADD END DAIKO
        Dim 評価額_売却額等Address = If(CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別, 評価額_売却額等項番, "Q11023001") 'REV_006 CHG DAIKO

        '「10日齢時点評価額・10日以前売却額等」の項目を取得する
        Dim query1 = chosahyoItemList.Where(Function(row) row.Key.IndexOf(評価額_売却額等Address) = 0).ToList()

        '「10日齢時点評価額・10日以前売却額等」欄を削除する
        For Each item In query1
            chosahyoItemList(item.Key).値 = Nothing
        Next

        '「搾乳牛所有状況２」シートの「個体識別番号」項目を抽出する
        Dim query2 = chosahyoItemList.Where(Function(row) row.Key.StartsWith("Q11021801")).ToList()

        If Not query2.Any() Then
            Return
        End If

        '「搾乳牛所有状況」シートの項目を抽出する
        Dim query3 = chosahyoItemList.Where(Function(row) row.Key.StartsWith("Q11020101") OrElse
                                                          row.Key.StartsWith("Q11020401") OrElse
                                                          row.Key.StartsWith("Q11020601") OrElse
                                                          row.Key.StartsWith("Q11020801")).ToList()

        If Not query3.Any() Then
            Return
        End If

        'REV_002 Add Start
        '調査票項目マスタから、総括データシートの項目番号を抽出する。



        Dim chosahyoItemMaster As DataTable = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)
        Dim SoukatsuDataRows As DataRow() = chosahyoItemMaster.Select("シート名='" & ComConst.調査票.牛総括データシート(CommonInfo.Chosakubun) & "'")
        'REV_002 Add End

        'レコードを削除する
        Dim deleteRecord = Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目), edaNoList As IEnumerable(Of String))

                               '抽出した資産の数ぶんループする
                               For Each edaNo In edaNoList

                                   'REV_002 Mod Start
                                   'For i = 11021801 To 11023101 Step 100

                                   '    'レコードが存在したら削除する
                                   '    If value.ContainsKey("Q" & i & edaNo) Then
                                   '        value("Q" & i & edaNo).値 = Nothing
                                   '    End If
                                   'Next

                                   '総括データシートの該当項目番号のデータをすべて削除する
                                   For Each SoukatsuDataRow In SoukatsuDataRows
                                       If value.ContainsKey(SoukatsuDataRow("項目番号").ToString & edaNo) Then
                                           value(SoukatsuDataRow("項目番号").ToString & edaNo).値 = Nothing
                                       End If
                                   Next
                                   'REV_002 Mod End
                               Next

                           End Sub

        '【集計論理記述表 ⑨-ア】
        '搾乳牛所有状況シートの「購入・売却」が「２：売却（売却）」又は「２：売却（死亡）」であり、
        '「育成牧場への預託」にデータがないレコードの「個体識別番号」と同一番号の資産を削除する
        Dim executeア = Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))
                           deleteRecord(value, DAOChosahyo.GetPrePrintDelData(db, New DAOChosahyo.PrimaryKey(_chosaNen, censusNo), "ア"))
                       End Sub

        '【集計論理記述表 ⑨-イ】
        '搾乳牛所有状況シートの「育成牧場への預託欄」が「1：出（入牧）」であり、かつ、同一の個体識別番号で
        '「育成牧場への預託」に「２：戻（下牧）」のレコードがないレコードの「個体識別番号」と同一番号の資産を削除する
        Dim executeイ = Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))
                           deleteRecord(value, DAOChosahyo.GetPrePrintDelData(db, New DAOChosahyo.PrimaryKey(_chosaNen, censusNo), "イ"))
                       End Sub

        '【集計論理記述表 ⑨-ウ】
        '搾乳牛所有状況シートの「異動年月日（短期売却区分）」欄が２又は３のレコードの「個体識別番号」と同一番号の資産を削除する
        Dim executeウ = Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                           '搾乳牛所有状況シートの「異動年月日（短期売却区分）」欄が２又は３のレコードの「個体識別番号」と同一番号の資産を持つレコードの枝番を取得する
                           Dim query4 = (From table1 In query2
                                         Join table2 In query3
                                         On table1.Value.値 Equals table2.Value.値
                                         Join table3 In query3
                                         On ComUtil.Chosahyo.GetDetailNo(table2.Key) Equals ComUtil.Chosahyo.GetDetailNo(table3.Key)
                                         Where table1.Key.StartsWith("Q11021801") AndAlso
                                               table2.Key.StartsWith("Q11020101") AndAlso
                                               table3.Key.StartsWith("Q11020801") AndAlso
                                               (table3.Value.値 = "2" OrElse table3.Value.値 = "3")
                                         Select ComUtil.Chosahyo.GetEdaNo(table1.Key)).Distinct()

                           'レコードを削除する
                           deleteRecord(value, query4)

                       End Sub

        'アの処理を実行する
        executeア(chosahyoItemList)

        'イの処理を実行する
        executeイ(chosahyoItemList)

        'ウの処理を実行する
        executeウ(chosahyoItemList)

        '歯抜けのレコードを上に詰める
        chosahyoItemList = ComUtil.Chosahyo.MoveUpMissingRow("Q11021801", chosahyoItemList, "【11】搾乳牛所有状況２")

    End Sub

    'REV_002 Mod Start (第２引数追加)
    ''' <summary>
    ''' 【集計論理記述表 ⑩】
    ''' 【２】牛２シート（育成牛・肥育牛生産費）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <param name="db"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri10(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目), ByVal db As DBAccess)
        '    Private Sub ExecuteRonri10(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))
        'REV_002 Mod End   (第２引数追加)

        If chosahyoItemList Is Nothing Then
            Return
        End If

        If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then
            Return
        End If

        '調査区分が営農類型（個人）または（法人）なら終了
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            Return
        End If

        If Not CommonInfo.Chosakubun = ComConst.調査区分.乳用雄育成牛生産費統計_個別 AndAlso
            Not CommonInfo.Chosakubun = ComConst.調査区分.交雑種育成牛生産費統計_個別 AndAlso
            Not CommonInfo.Chosakubun = ComConst.調査区分.去勢若齢肥育牛生産費統計_個別 AndAlso
            Not CommonInfo.Chosakubun = ComConst.調査区分.乳用雄肥育牛生産費統計_個別 AndAlso
            Not CommonInfo.Chosakubun = ComConst.調査区分.交雑種肥育牛生産費統計_個別 AndAlso
            Not CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_乳用雄育成牛生産費 AndAlso
            Not CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_交雑種育成牛生産費 AndAlso
            Not CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費 AndAlso
            Not CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費 AndAlso
            Not CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_交雑種肥育牛生産費 Then

            Return

        End If

        '「【２】牛１」シートの処理対象項目を抽出する
        Dim query1 = chosahyoItemList.Where(Function(row) row.Key.StartsWith("Q02020101") OrElse
                                                          row.Key.StartsWith("Q02020401")).ToList()

        If Not query1.Any() Then
            Return
        End If

        'REV_002 Add Start
        '調査票項目マスタから、総括データシートの項目番号を抽出する。



        Dim chosahyoItemMaster As DataTable = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)
        Dim SoukatsuDataRows As DataRow() = chosahyoItemMaster.Select("シート名='" & ComConst.調査票.牛総括データシート(CommonInfo.Chosakubun) & "'")
        'REV_002 Add End

        '「２：売却」が入力されているデータの削除処理を実行する
        Dim deleteRecord = Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                               '「購入・売却」列が「２：売却（売却）」又は「２：売却（死亡）」であるレコードを抽出する
                               Dim query2 = query1.Where(Function(row) row.Key.StartsWith("Q02020401") AndAlso
                                                                       (row.Value.値 = 売却_売却 OrElse row.Value.値 = 売却_死亡)).
                                                   Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key)).
                                                   ToList()

                               '「購入・売却」列が「２：売却（売却）」又は「２：売却（死亡）」であるレコードの数ぶんループする
                               For Each edaNo In query2

                                   '「個体識別番号」を抽出する
                                   Dim query3 = query1.Where(Function(row) row.Key.StartsWith("Q02020101") AndAlso row.Key.EndsWith(edaNo)).FirstOrDefault()

                                   If IsNothing(query3.Value) OrElse String.IsNullOrWhiteSpace(query3.Value.値) Then
                                       Continue For
                                   End If

                                   '先に取得した「個体識別番号」と同一番号の資産を抽出する
                                   Dim query4 = value.Where(Function(row) row.Key.StartsWith("Q02021401") AndAlso row.Value.値 = query3.Value.値).
                                                      Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key))

                                   '抽出した資産の数ぶんループする
                                   For Each edaNo2 In query4

                                       'REV_002 Mod Start
                                       'For i = 2021401 To 2023001 Step 100

                                       '    'レコードが存在したら削除する
                                       '    If value.ContainsKey("Q0" & i & edaNo2) Then
                                       '        value("Q0" & i & edaNo2).値 = Nothing
                                       '    End If
                                       'Next

                                       '総括データシートの該当項目番号のデータをすべて削除する
                                       For Each SoukatsuDataRow In SoukatsuDataRows
                                           If value.ContainsKey(SoukatsuDataRow("項目番号").ToString & edaNo2) Then
                                               value(SoukatsuDataRow("項目番号").ToString & edaNo2).値 = Nothing
                                           End If
                                       Next
                                       'REV_002 Mod End
                                   Next
                               Next

                           End Sub

        '「仮売却年・月」が入力されているデータの更新処理を実行する
        Dim updateRecord = Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                               '「仮売却年・月」が入力されているレコードを抽出する
                               Dim query2 = value.Where(Function(row) (row.Key.StartsWith("Q02022701") AndAlso Not String.IsNullOrWhiteSpace(row.Value.値)) OrElse
                                                                      (row.Key.StartsWith("Q02022801") AndAlso Not String.IsNullOrWhiteSpace(row.Value.値))).
                                                  Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key)).
                                                  Distinct().
                                                  ToList()

                               For Each edaNo In query2

                                   Dim itemNumber = "Q02021801" & edaNo

                                   'レコードが存在したら「2：乖離売却済み」をセットする
                                   If value.ContainsKey(itemNumber) Then
                                       value(itemNumber).値 = "2"
                                   End If
                               Next

                           End Sub

        '「２：売却」が入力されているデータの削除処理を実行する
        deleteRecord(chosahyoItemList)

        '「仮売却年・月」が入力されているデータの更新処理を実行する
        updateRecord(chosahyoItemList)

        '歯抜けのレコードを上に詰める
        chosahyoItemList = ComUtil.Chosahyo.MoveUpMissingRow("Q02020101", chosahyoItemList, "【２】牛２")

    End Sub

    ''' <summary>
    ''' 【集計論理記述表 ⑪】
    ''' 【13】地代・搾乳牛負担割合・自給牧草シート（牛乳生産費）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri11(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then
            Return
        End If

        If Not CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 AndAlso
            Not CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then

            Return

        End If

        '「実面積」欄以外の項目を取得する
        Dim query = chosahyoItemList.Where(Function(row) (row.Key.StartsWith("Q13") OrElse row.Key.StartsWith("Q14")) AndAlso
                                                         Not row.Key.StartsWith("Q1301010") AndAlso
                                                         Not row.Key.StartsWith("Q1302010")).ToList()

        '「実面積」欄以外を削除する
        For Each item In query
            chosahyoItemList(item.Key).値 = Nothing
        Next

    End Sub

    ''' <summary>
    ''' 【集計論理記述表 ⑫】
    ''' 【13】地代シート（牛乳生産費以外の畜生）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri12(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（個人）または（法人）なら終了
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            Return
        End If

        If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 Then
            Return
        End If

        If CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse
            CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then

            Return

        End If

        '「実面積」欄以外の項目を取得する
        Dim query = chosahyoItemList.Where(Function(row) row.Key.StartsWith("Q13") AndAlso
                                                         Not row.Key.StartsWith("Q1301010") AndAlso
                                                         Not row.Key.StartsWith("Q1302010")).ToList()

        '「実面積」欄以外を削除する
        For Each item In query
            chosahyoItemList(item.Key).値 = Nothing
        Next

    End Sub

    ''' <summary>
    ''' 【集計論理記述表 ⑬】
    ''' 【13】飼料用米①シート（米生産費（個別経営））
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri13(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        If CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費 Then
            Return
        End If

        If Not CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_個別 Then
            Return
        End If

        '「作付面積（ａ）」欄及び「飼料用米と食用米の10ａ当たり費用及び労働時間の違いの有無」欄以外の項目を取得する
        Dim query = chosahyoItemList.Where(Function(row) (row.Key.StartsWith("Q1301") OrElse
                                                          row.Key.StartsWith("Q1302") OrElse
                                                          row.Key.StartsWith("Q1303")) AndAlso
                                                         Not row.Key.StartsWith("Q1301010") AndAlso
                                                         Not row.Key.StartsWith("Q13030101")).ToList()

        '「作付面積（ａ）」欄及び「飼料用米と食用米の10ａ当たり費用及び労働時間の違いの有無」欄以外を削除する
        For Each item In query
            chosahyoItemList(item.Key).値 = Nothing
        Next

    End Sub

    ''' <summary>
    ''' 【集計論理記述表 ⑭】
    ''' 【13】飼料用米②シート（米生産費（個別経営））
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri14(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        If CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費 Then
            Return
        End If

        If Not CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_個別 Then
            Return
        End If

        '「費用の種類」の「その他」欄、「食用米との違い」欄及び「割合」欄以外の項目を取得する
        Dim query = chosahyoItemList.Where(Function(row) (row.Key.StartsWith("Q1304") OrElse
                                                          row.Key.StartsWith("Q1305")) AndAlso
                                                         Not row.Key.StartsWith("Q130401") AndAlso
                                                         Not row.Key.StartsWith("Q130402") AndAlso
                                                         Not row.Key.StartsWith("Q130403") AndAlso
                                                         Not row.Key.StartsWith("Q130501") AndAlso
                                                         Not row.Key.StartsWith("Q130502")).ToList()

        '「費用の種類」の「その他」欄、「食用米との違い」欄及び「割合」欄以外を削除する
        For Each item In query
            chosahyoItemList(item.Key).値 = Nothing
        Next

    End Sub

    '---REV_001_ADD START
    'REV_002 Mod Start (第２引数追加)
    ''' <summary>
    ''' 【集計論理記述表 ⑮】
    ''' 【2】対象蓄概要１シート（子牛（個別経営））
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <param name="db"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri15(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目), ByVal db As DBAccess)
        'Private Sub ExecuteRonri15(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))
        'REV_002 Mod End (第２引数追加)

        If chosahyoItemList Is Nothing Then
            Return
        End If

        If CommonInfo.Chosakubun <> ComConst.調査区分.子牛生産費統計_個別 And
            CommonInfo.Chosakubun <> ComConst.調査区分.経営分析調査_子牛生産費 Then
            Return
        End If

        '「【２】2牛取引情報」シートの処理対象項目を抽出する
        Dim query1 = chosahyoItemList.Where(Function(row) row.Key.StartsWith("Q02020101") OrElse
                                                          row.Key.StartsWith("Q02020401")).ToList()

        If Not query1.Any() Then
            Return
        End If

        'REV_002 Add Start
        '調査票項目マスタから、総括データシートの項目番号を抽出する。
        Dim chosahyoItemMaster As DataTable = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)
        Dim SoukatsuDataRows As DataRow() = chosahyoItemMaster.Select("シート名='" & ComConst.調査票.牛総括データシート(CommonInfo.Chosakubun) & "'")
        'REV_002 Add End

        '（１）条件に該当する牛を除外
        Dim deleteRecord =
            Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                '「異動の状況」列が「２：転出（売却）」又は「２：転出（死亡）」であるレコードを抽出する
                Dim query2 = query1.Where(Function(row) row.Key.StartsWith("Q02020401") AndAlso
                                                        (row.Value.値 = 転出_売却 OrElse row.Value.値 = 転出_死亡)).
                                    Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key)).
                                    ToList()

                '「異動の状況」列が「２：転出（売却）」又は「２：転出（死亡）」であるレコードの数ぶんループする
                For Each edaNo In query2

                    '「個体識別番号」を抽出する
                    Dim query3 = query1.Where(Function(row) row.Key.StartsWith("Q02020101") AndAlso row.Key.EndsWith(edaNo)).FirstOrDefault()

                    If IsNothing(query3.Value) OrElse String.IsNullOrWhiteSpace(query3.Value.値) Then
                        Continue For
                    End If

                    '先に取得した「個体識別番号」と同一番号の資産を抽出する
                    Dim query4 = value.Where(Function(row) row.Key.StartsWith("Q02030301") AndAlso row.Value.値 = query3.Value.値).
                                       Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key))

                    '　種類が「肉用牛」の牛の資産を抽出する
                    Dim queryType1 = value.Where(Function(row) row.Key.StartsWith("Q02030501") AndAlso row.Value.値 = "1").
                                           Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key))

                    '　種類が「繁殖牛」で、自身から生産された牛の中で前年調査終了時に調査対象畜（子牛）として飼養される牛がいない牛を抽出する
                    Dim queryType2 = value.Where(Function(row) row.Key.StartsWith("Q02030501") AndAlso row.Value.値 = "2").
                                           Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key))
                    queryType2 = value.Where(Function(row) row.Key.StartsWith("Q02034601") AndAlso (String.IsNullOrEmpty(row.Value.値) OrElse row.Value.値 = "0")).
                                           Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key)).Intersect(queryType2)

                    '条件に該当する牛のみ抽出
                    query4 = query4.Intersect(queryType1.Union(queryType2))

                    '抽出した資産の数ぶんループする
                    For Each edaNo2 In query4

                        'REV_002 Mod Start
                        'For i = 2030301 To 2032001 Step 100

                        '    'レコードが存在したら削除する
                        '    If value.ContainsKey("Q0" & i & edaNo2) Then
                        '        value("Q0" & i & edaNo2).値 = Nothing
                        '    End If
                        'Next
                        'For i = 2033901 To 2036501 Step 100

                        '    'レコードが存在したら削除する
                        '    If value.ContainsKey("Q0" & i & edaNo2) Then
                        '        value("Q0" & i & edaNo2).値 = Nothing
                        '    End If
                        'Next
                        'For i = 2039101 To 2039111 Step 1

                        '    'レコードが存在したら削除する
                        '    If value.ContainsKey("Q0" & i & edaNo2) Then
                        '        value("Q0" & i & edaNo2).値 = Nothing
                        '    End If
                        'Next

                        '総括データシートの該当項目番号のデータをすべて削除する
                        For Each SoukatsuDataRow In SoukatsuDataRows
                            If value.ContainsKey(SoukatsuDataRow("項目番号").ToString & edaNo2) Then
                                value(SoukatsuDataRow("項目番号").ToString & edaNo2).値 = Nothing
                            End If
                        Next
                        'REV_002 Mod End

                    Next
                Next

                'REV_004↓種類が「繁殖牛」かつ性別が「雌」で、前年の調査開始以前に売却又は死亡した牛を抽出する
                Dim judgeYM = String.Format("{0:0000}12", CInt(_chosaNen) - 1)
                Dim query5 = value.Where(Function(row) row.Key.StartsWith("Q02030501") AndAlso row.Value.値 = "2").
                                           Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key))
                query5 = value.Where(Function(row) row.Key.StartsWith("Q02030801") AndAlso row.Value.値 = "2").
                                           Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key)).Intersect(query5)
                query5 = value.Where(Function(row) row.Key.StartsWith("Q02035101") AndAlso Not String.IsNullOrEmpty(row.Value.値) AndAlso row.Value.値 <= judgeYM).
                                           Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key)).Intersect(query5)
                For Each edaNo3 In query5
                    '総括データシートの該当項目番号のデータをすべて削除する
                    For Each SoukatsuDataRow In SoukatsuDataRows
                        If value.ContainsKey(SoukatsuDataRow("項目番号").ToString & edaNo3) Then
                            value(SoukatsuDataRow("項目番号").ToString & edaNo3).値 = Nothing
                        End If
                    Next
                Next
                'REV_004↑

            End Sub

        '（２）前年の調査結果により出力内容が変わる項目の更新
        Dim updateItemList1 =
            Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                '「仮売却年月」にデータがある場合は「1」をセット
                Dim queryFrom = value.Where(Function(row) row.Key.StartsWith("Q02031701") AndAlso
                                                            Not String.IsNullOrEmpty(row.Value.値)).
                                                            Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key))

                Dim list As New List(Of String)(queryFrom)
                For Each edaNo In list
                    Dim key As String = "Q02030901" & edaNo
                    '「１：仮売却済み」をセットする
                    If Not value.ContainsKey(key) Then
                        value.Add(key, New DAOChosahyo.調査票項目)
                    End If
                    value(key).値 = "1"
                    value(key).シート名 = ComConst.調査票.牛総括データシート(CommonInfo.Chosakubun)
                Next

                '「前年の調査結果＿前回計算対象時年月」[AX, AY列]は、「今回子牛販売時年月」[AP, AQ列]にデータがある場合は当該データをセット、
                'それ以外は「前回計算対象時年月」[AN, AO列]のデータをセット。
                updateChosahyoItem(value, {"Q02035601", "Q02035401"}, "Q02034201", True)
                updateChosahyoItem(value, {"Q02035701", "Q02035501"}, "Q02034301", True)
            End Sub

        '（３）データの参照元と出力箇所が違う項目の更新
        Dim updateItemList2 =
            Sub(value As Dictionary(Of String, DAOChosahyo.調査票項目))

                '前年の調査結果＿購入価格
                updateChosahyoItem(value, {"Q02031401"}, "Q02033901")

                '前年の調査結果＿成畜年月
                Dim queryFrom = value.Where(Function(row) row.Key.StartsWith("Q02034401")).Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key))
                Dim list As New List(Of String)(queryFrom)
                For Each edaNo In list
                    Dim keyFrom As String = "Q02034401" & edaNo

                    If Not value.ContainsKey(keyFrom) Then
                        Continue For
                    End If
                    If String.IsNullOrEmpty(value(keyFrom).値) Or value(keyFrom).値 = "0" Then
                        Continue For
                    End If

                    If value(keyFrom).値.Length >= 6 Then
                        '年をセット
                        Dim keyTo As String = "Q02034001" & edaNo
                        If Not value.ContainsKey(keyTo) Then
                            value.Add(keyTo, New DAOChosahyo.調査票項目)
                        End If
                        value(keyTo).値 = value(keyFrom).値.Substring(0, 4)
                        value(keyTo).シート名 = value(keyFrom).シート名

                        '月をセット
                        keyTo = "Q02034101" & edaNo
                        If Not value.ContainsKey(keyTo) Then
                            value.Add(keyTo, New DAOChosahyo.調査票項目)
                        End If
                        value(keyTo).値 = value(keyFrom).値.Substring(4, 2)
                        value(keyTo).シート名 = value(keyFrom).シート名
                    End If
                Next

                '前年の調査結果＿調査開始以前の分べん計算の対象回数
                updateChosahyoItem(value, {"Q02034801"}, "Q02034901")

                '前年の調査結果＿自身の生んだ計算対象畜（子牛）のうち最も遅い分べん年月
                updateChosahyoItem(value, {"Q02034701"}, "Q02035001", True)

                '前年の調査結果＿調査期間より前に転出した繁殖雌牛＿転出年月
                updateChosahyoItem(value, {"Q02034501"}, "Q02035101", True)

                '前年の調査結果＿調査期間より前に転出した繁殖雌牛＿売却額
                updateChosahyoItem(value, {"Q02035301"}, "Q02035201")
            End Sub

        'レコードを削除する
        deleteRecord(chosahyoItemList)

        '調査票データの更新
        updateItemList1(chosahyoItemList)
        updateItemList2(chosahyoItemList)

        'プレプリント対象でない項目番号を取得する
        Dim query = chosahyoItemList.Where(Function(row) row.Key.StartsWith("Q02031401") OrElse
                                                         row.Key.StartsWith("Q02031601") OrElse
                                                         row.Key.StartsWith("Q02031701") OrElse
                                                         row.Key.StartsWith("Q02031801") OrElse
                                                         row.Key.StartsWith("Q02031901") OrElse
                                                         row.Key.StartsWith("Q02032001") OrElse
                                                         row.Key.StartsWith("Q02034401") OrElse
                                                         row.Key.StartsWith("Q02034501") OrElse
                                                         row.Key.StartsWith("Q02034601") OrElse
                                                         row.Key.StartsWith("Q02034701") OrElse
                                                         row.Key.StartsWith("Q02034801") OrElse
                                                         row.Key.StartsWith("Q02035301") OrElse
                                                         row.Key.StartsWith("Q02035401") OrElse
                                                         row.Key.StartsWith("Q02035501") OrElse
                                                         row.Key.StartsWith("Q02035601") OrElse
                                                         row.Key.StartsWith("Q02035701") OrElse
                                                         row.Key.StartsWith("Q02035801") OrElse
                                                         row.Key.StartsWith("Q02035901") OrElse
                                                         row.Key.StartsWith("Q02036001") OrElse
                                                         row.Key.StartsWith("Q02036101") OrElse
                                                         row.Key.StartsWith("Q02036201") OrElse
                                                         row.Key.StartsWith("Q02036301") OrElse
                                                         row.Key.StartsWith("Q02036401") OrElse
                                                         row.Key.StartsWith("Q02036501")).ToList()

        'プレプリント対象でない項目を削除する
        For Each item In query
            chosahyoItemList(item.Key).値 = Nothing
        Next

        '歯抜けのレコードを上に詰める
        chosahyoItemList = ComUtil.Chosahyo.MoveUpMissingRow("Q02030301", chosahyoItemList, "【２】対象畜概要１")

    End Sub
    '---REV_001_ADD END


    '---REV_003_ADD START
    ''' <summary>
    ''' 【集計論理記述表 ⑯】
    ''' 04_事業収支の概要・05_投資と資金シート（営農個人）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri16(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（個人）でないなら終了
        If Not CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            Return
        End If

        '04_事業収支の概要・05_投資と資金シートのうち【４】事業収支の項目を取得
        Dim query = chosahyoItemList.Where(Function(row) row.Key.StartsWith("Q04")).ToList()


        '【４】事業収支の項目を削除する（【５】の項目はプレプリントされる値として保持したままにする）
        For Each item In query
            chosahyoItemList(item.Key).値 = Nothing
        Next

    End Sub


    ''' <summary>
    ''' 【集計論理記述表 ⑰】
    ''' 08_01_生産概況、農畜産物収入（営農個人）、09_01_生産概況+農畜産物収入（営農法人））
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri17(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型でないなら終了
        If Not CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計 Then
            Return
        End If

        '営農個人の場合Q0801を、営農法人の場合はQ0901を割り当てる
        Dim not_target = If(CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人, "Q0801", "Q0901")

        'Nothingを入れる、項目名のための文字列を作成
        Dim 販売数量address1 = not_target & "04"
        Dim 販売数量address2 = not_target & "09"
        Dim 販売金額address1 = not_target & "05"
        Dim 販売金額address2 = not_target & "10"
        Dim 販売金額address3 = not_target & "12"


        'Nothingを入れる項目を取得する
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            Dim query = chosahyoItemList.Where(Function(row) (row.Key.StartsWith(販売数量address1) OrElse
                                                   row.Key.StartsWith(販売数量address2) OrElse
                                                   row.Key.StartsWith(販売金額address1) OrElse
                                                   row.Key.StartsWith(販売金額address2) OrElse
                                                   row.Key.StartsWith(販売金額address3) OrElse
                                                   row.Key.StartsWith("Q08011304"))).ToList()            '営農類型個人　指定品目（販売金額）のための追加


            'プレプリント対象外の項目にNothingを入れる
            For Each item In query
                chosahyoItemList(item.Key).値 = Nothing
            Next

        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            Dim query = chosahyoItemList.Where(Function(row) (row.Key.StartsWith(販売数量address1) OrElse
                                                       row.Key.StartsWith(販売数量address2) OrElse
                                                       row.Key.StartsWith(販売金額address1) OrElse
                                                       row.Key.StartsWith(販売金額address2) OrElse
                                                       row.Key.StartsWith(販売金額address3))).ToList()

            'プレプリント対象外の項目にNothingを入れる
            For Each item In query
                chosahyoItemList(item.Key).値 = Nothing
            Next

        End If

    End Sub



    ''' <summary>
    ''' 【集計論理記述表⑱】
    ''' 08_02_生産概況、農畜産物収入（営農個人）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri18(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（個人）でないなら終了
        If Not CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            Return
        End If


        'Nothingを入れる項目を取得する
        Dim query = chosahyoItemList.Where(Function(row) (row.Key.StartsWith("Q080202") OrElse
                                               row.Key.StartsWith("Q080203") OrElse
                                               row.Key.StartsWith("Q08020115") OrElse
                                               row.Key.StartsWith("Q08020103") OrElse
                                               row.Key.StartsWith("Q08020116") OrElse
                                               row.Key.StartsWith("Q08020117")) AndAlso
                                                         Not row.Key.StartsWith("Q08020214")).ToList()


        'プレプリント対象外の項目にNothingを入れる
        For Each item In query
            chosahyoItemList(item.Key).値 = Nothing
        Next


    End Sub

    ''' <summary>
    ''' 【集計論理記述表⑲】
    ''' 09_02_生産概況+農畜産物収入（営農法人）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri19(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（法人）でないなら終了
        If Not CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            Return
        End If


        'Nothingを入れる項目を取得する
        Dim query = chosahyoItemList.Where(Function(row) (row.Key.StartsWith("Q090202") OrElse
                                               row.Key.StartsWith("Q090203") OrElse
                                               row.Key.StartsWith("Q09020115") OrElse
                                               row.Key.StartsWith("Q09020116") OrElse
                                               row.Key.StartsWith("Q09020117")) AndAlso
                                                         Not row.Key.StartsWith("Q09020214")).ToList()


        'プレプリント対象外の項目にNothingを入れる
        For Each item In query
            chosahyoItemList(item.Key).値 = Nothing
        Next


    End Sub


    ''' <summary>
    ''' 【集計論理記述表⑳】
    ''' 10_労働（営農個人）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri20(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（個人）でないなら終了
        If Not CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            Return
        End If

        'REV_005↓経営主もプレプリント対象にする
        ''経営主の欄を対象外にする
        'chosahyoItemList("Q10010101").値 = Nothing
        'chosahyoItemList("Q10010201").値 = Nothing
        'chosahyoItemList("Q10010401").値 = Nothing
        'chosahyoItemList("Q10010501").値 = Nothing
        '経営主の年齢を＋１する。
        Dim value = chosahyoItemList("Q10010201").値
        chosahyoItemList("Q10010201").値 = If(IsNumeric(value), CStr(CInt(value) + 1), Nothing)
        'REV_005↑

        '可変項目を取得する
        Dim query1 = chosahyoItemList.Where(Function(row) row.Key.StartsWith("Q1001") AndAlso
                                                          row.Key.IndexOf("_") = 9).ToList()

        If Not query1.Any() Then
            Return
        End If

        '家族・雇用の別の可変項目を取得
        Dim query2 = chosahyoItemList.Where(Function(row) row.Key.StartsWith("Q10010302") AndAlso
                                                          row.Key.IndexOf("_") = 9).ToList()

        '家族・雇用の別分ループする(行数分ループする)
        For Each item As KeyValuePair(Of String, DAOChosahyo.調査票項目) In query2

            If (item.Value.値 = "1：家族") Then
                'REV_005↓可変のため存在チェックと数値チェックを追加
                'chosahyoItemList("Q10010202" & ComUtil.Chosahyo.GetEdaNo(item.Key)).値 = CStr(CInt(chosahyoItemList("Q10010202" & ComUtil.Chosahyo.GetEdaNo(item.Key)).値) + 1)
                If chosahyoItemList.ContainsKey("Q10010202" & ComUtil.Chosahyo.GetEdaNo(item.Key)) Then
                    value = chosahyoItemList("Q10010202" & ComUtil.Chosahyo.GetEdaNo(item.Key)).値
                    chosahyoItemList("Q10010202" & ComUtil.Chosahyo.GetEdaNo(item.Key)).値 = If(IsNumeric(value), CStr(CInt(value) + 1), Nothing)
                End If
                'REV_005↑
            Else

                '削除対象項目を取得する
                Dim query3 = query1.Where(Function(row) row.Key.EndsWith(ComUtil.Chosahyo.GetEdaNo(item.Key))).ToList()

                '「1:家族」以外の入力のある行を削除する
                For Each deleteItem In query3
                    chosahyoItemList(deleteItem.Key).値 = Nothing
                Next

            End If

        Next

        'REV_005↓空行を詰める
        chosahyoItemList = ComUtil.Chosahyo.MoveUpMissingRow("Q1001", chosahyoItemList)
        'REV_005↑

    End Sub


    ''' <summary>
    ''' 【集計論理記述表㉑】
    ''' 11_労働（指定品目）（営農個人）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri21(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（個人）でないなら終了
        If Not CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            Return
        End If

        'Nothingを入れる項目を取得する
        Dim query = chosahyoItemList.Where(Function(row) row.Key.StartsWith("Q11"))


        'プレプリント対象外の項目にNothingを入れる
        For Each item In query
            chosahyoItemList(item.Key).値 = Nothing
        Next

    End Sub


    ''' <summary>
    ''' 【集計論理記述表㉒】
    ''' 10_01_受託収入、10_02_受託収入（営農法人）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri22(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（法人）でないなら終了
        If Not CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            Return
        End If

        'Nothingを入れる項目を取得する
        Dim query = chosahyoItemList.Where(Function(row) (row.Key.StartsWith("Q100202") OrElse '連絡票No283対応（"Q1001"を削除）　2022/03/02
                                               row.Key.StartsWith("Q100203") OrElse
                                               row.Key.StartsWith("Q100205") OrElse
                                               row.Key.StartsWith("Q100206") OrElse
                                               row.Key.StartsWith("Q100208")) OrElse
                                               row.Key.StartsWith("Q100209")).ToList()


        'プレプリント対象外の項目にNothingを入れる
        For Each item In query
            chosahyoItemList(item.Key).値 = Nothing
        Next


    End Sub

    '2022/03/02　ADD START　連絡票No.283対応
    ''' <summary>
    ''' 【集計論理記述表㉓】
    ''' 10_01_受託収入、10_02_受託収入（営農法人）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteRonri23(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（法人）でないなら終了
        If Not CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            Return
        End If

        chosahyoItemList("Q01010107").値 = CStr(CInt(chosahyoItemList("Q01010107").値) + 1)

    End Sub
    '2022/03/02　END START　連絡票No.283対応


    '---REV_003_ADD END

    ''' <summary>
    ''' 【集計論理記述表㉔】
    ''' はじめに・01_現況（営農個人）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks>REV_005</remarks>
    Private Sub ExecuteRonri24(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（個人）でないなら終了
        If Not CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            Return
        End If

        '経営主の年齢を前年データ＋１してセットする。
        Dim value = chosahyoItemList("Q01010152").値
        chosahyoItemList("Q01010152").値 = If(IsNumeric(value), CStr(CInt(value) + 1), Nothing)

    End Sub

    ''' <summary>
    ''' 【集計論理記述表㉕】
    ''' 06_給与の状況（営農法人）
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks>REV_005</remarks>
    Private Sub ExecuteRonri25(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        If chosahyoItemList Is Nothing Then
            Return
        End If

        '調査区分が営農類型（法人）でないなら終了
        If Not CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            Return
        End If

        '有給役員の平均年齢を前年データ＋１してセットする。
        '男
        Dim value = chosahyoItemList("Q06020101").値
        chosahyoItemList("Q06020101").値 = If(IsNumeric(value), CStr(CInt(value) + 1), Nothing)
        '女
        value = chosahyoItemList("Q06020102").値
        chosahyoItemList("Q06020102").値 = If(IsNumeric(value), CStr(CInt(value) + 1), Nothing)

    End Sub

#End Region

    ''' <summary>
    ''' コピー不要な調査票項目をNothingに置換する
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ReplaceNotCopyItemToNothing(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        '2つの文字列の文字数を比較し、小さいほうの文字数を取得する
        Dim getLength = Function(value1 As String, value2 As String)
                            Return If(value1.Length < value2.Length, value1.Length, value2.Length)
                        End Function

        'シート名の一部を抜粋する
        Dim substringSheetName = Function(value1 As String, value2 As String)
                                     Return value1.Substring(0, getLength(value1, value2))
                                 End Function

        'コピー対象のシートかチェックする
        Dim isCopySheet = Function(sheetList As List(Of String), checkSheet As String)
                              Return If(0 <= sheetList.FindIndex(Function(sheet) 0 <= sheet.IndexOf(substringSheetName(checkSheet, sheet))), True, False)
                          End Function

        'コピーするシートの一覧を取得する
        Dim copyTargetSheetList As List(Of String) = GetCopyTargetSheetList()

        For Each chosahyoItem In chosahyoItemList

            'コピー対象外のシートか
            If Not isCopySheet(copyTargetSheetList, chosahyoItem.Value.シート名) Then

                'コピー対象外の項目にはNothingを代入する
                chosahyoItem.Value.値 = Nothing

            End If
        Next

        '牛乳だけ特別処理を行う
        If CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse
            CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then

            '「【11】搾乳牛所有状況」シートの項目数ぶんループする
            For Each chosahyoItem In chosahyoItemList.Where(Function(row) row.Value.シート名 = "【11】搾乳牛所有状況")

                'コピー対象外の項目にはNothingを代入する
                chosahyoItem.Value.値 = Nothing
            Next
        End If

    End Sub

    ''' <summary>
    ''' コピーするシートの一覧を取得する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCopyTargetSheetList() As List(Of String)

        Select Case CommonInfo.Kubun2


    '---REV_003_ADD START
            Case ComConst.区分２.営農類型別経営統計

                '区分ごとにコピーするシートが異なる
                Select Case CommonInfo.Chosakubun
                    Case ComConst.調査区分.営農類型別経営統計_個人

                        Dim copyTargetSheet As List(Of String) = New List(Of String) From {"00_表紙",
                                                                   "はじめに・01_現況",
                                                                   "04_事業収支の概要・05_投資と資金",
                                                                   "06_固定資産・07土地面積",
                                                                   "08_01_生産概況、農畜産物収入",
                                                                   "08_02_生産概況、農畜産物収入",
                                                                   "10_労働",
                                                                   "11_労働（指定品目）",
                                                                   "12_01_農産関連事業収支"}

                        Return copyTargetSheet

                    Case ComConst.調査区分.営農類型別経営統計_法人

                        Dim copyTargetSheet As List(Of String) = New List(Of String) From {"00_表紙",
                                                                   "01_02_はじめに",
                                                                   "03_投資と資金",
                                                                   "06_給与の状況、07_土地面積",
                                                                   "08_主要農業固定資産",
                                                                   "09_01_生産概況+農畜産物収入",
                                                                   "09_02_生産概況＋農畜産物収入",
                                                                   "10_01_受託収入",
                                                                   "10_02_受託収入",
                                                                   "13_01_生産関連事業収支"}

                        Return copyTargetSheet


                End Select





    '---REV_003_ADD END


            Case ComConst.区分２.農産物生産費

                '以下のシートの前年データをコピーする
                Dim copyTargetSheet As List(Of String) = New List(Of String) From {"指標部入力",
                                                                                           "【１】経営の概況①",
                                                                                           "【１】経営の概況②",
                                                                                           "【１】経営の概況③",
                                                                                           "【５】土地改良及び水利費・【６】借入金",
                                                                                           "【７】建物及び構築物",
                                                                                           "【８】自動車",
                                                                                           "【９】農業機械",
                                                                                           "【11】土地"}

                '「米（個別経営）」のみコピー対象のシートが多い
                If CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_個別 Then

                    copyTargetSheet.Add("【13】飼料用米①")
                    copyTargetSheet.Add("【13】飼料用米②")
                End If

                Return copyTargetSheet

            Case ComConst.区分２.畜産物生産費

                '以下のシートの前年データをコピーする
                Dim copyTargetSheet As List(Of String) = New List(Of String) From {"表紙",
                                                                                           "【１】経営概要",
                                                                                           "【４】物件税～",
                                                                                           "【13】地代"}

                '区分ごとにコピーするシートが異なる
                Select Case CommonInfo.Chosakubun

                    Case ComConst.調査区分.牛乳生産費統計_個別,
                                ComConst.調査区分.経営分析調査_牛乳生産費

                        copyTargetSheet.Add("【６】借入金")
                        copyTargetSheet.Add("【７】建物")
                        copyTargetSheet.Add("【8】自動車")
                        copyTargetSheet.Add("【８】自動車")
                        copyTargetSheet.Add("【９】農業機械")
                        copyTargetSheet.Add("【11】農家団体コード")
                        copyTargetSheet.Add("【11】搾乳牛所有状況２")

                    Case ComConst.調査区分.子牛生産費統計_個別,
                                ComConst.調査区分.経営分析調査_子牛生産費

                        copyTargetSheet.Add("【２】農家団体コード")
                        '---REV_001_ADD START
                        copyTargetSheet.Add("【２】対象畜概要１")
                        '---REV_001_ADD END
                        copyTargetSheet.Add("【８】建物")
                        copyTargetSheet.Add("【９】自動車")
                        copyTargetSheet.Add("【10】農業機械")

                    Case ComConst.調査区分.乳用雄育成牛生産費統計_個別,
                                ComConst.調査区分.交雑種育成牛生産費統計_個別,
                                ComConst.調査区分.去勢若齢肥育牛生産費統計_個別,
                                ComConst.調査区分.乳用雄肥育牛生産費統計_個別,
                                ComConst.調査区分.交雑種肥育牛生産費統計_個別,
                                ComConst.調査区分.経営分析調査_乳用雄育成牛生産費,
                                ComConst.調査区分.経営分析調査_交雑種育成牛生産費,
                                ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費,
                                ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費,
                                ComConst.調査区分.経営分析調査_交雑種肥育牛生産費

                        copyTargetSheet.Add("【２】農家団体コード")
                        copyTargetSheet.Add("【２】牛２")
                        copyTargetSheet.Add("【８】建物")
                        copyTargetSheet.Add("【９】自動車")
                        copyTargetSheet.Add("【10】農業機械")

                    Case ComConst.調査区分.肥育豚生産費統計_個別,
                                ComConst.調査区分.経営分析調査_肥育豚生産費

                        copyTargetSheet.Add("【８】建物")
                        copyTargetSheet.Add("【９】自動車")
                        copyTargetSheet.Add("【10】農業機械")

                End Select

                Return copyTargetSheet

            Case Else



        End Select

        Return Nothing

    End Function

    ''' <summary>
    ''' ブランクが設定されている調査票項目をNothingに置換する
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub ReplaceBlankToNothing(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        'ブランクが設定されている可変テーブル「以外」の項目を取得する
        Dim query = chosahyoItemList.Where(Function(row) row.Value.値 = "" AndAlso
                                                         row.Value.値 IsNot Nothing AndAlso
                                                         Not row.Key.IndexOf("_") = 9).
                                     Select(Function(row) row.Key)

        For Each key In query
            chosahyoItemList(key).値 = Nothing
        Next

    End Sub

    ''' <summary>
    ''' 可変テーブル内の値がNothingのデータを削除しDBに書き込まないようにする
    ''' </summary>
    ''' <param name="chosahyoItemList"></param>
    ''' <remarks></remarks>
    Private Sub RemoveNothingDataForVariableTable(ByRef chosahyoItemList As Dictionary(Of String, DAOChosahyo.調査票項目))

        '値が入力されている可変テーブルのデータ及び可変テーブル「以外」のデータを取得する
        chosahyoItemList = chosahyoItemList.Where(Function(row) (Not String.IsNullOrWhiteSpace(row.Value.値) AndAlso
                                                                 row.Key.IndexOf("_") = 9) OrElse
                                                                Not row.Key.IndexOf("_") = 9).
                                            ToDictionary(Function(row) row.Key, Function(row) row.Value)

    End Sub

    '---REV_001_ADD START
    ''' <summary>
    ''' 調査票項目の更新
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="koubanFrom"></param>
    ''' <param name="koubanTo"></param>
    ''' <returns>更新有無（koubanFromの項目が存在するか）</returns>
    ''' <remarks></remarks>
    Private Function updateChosahyoItem(ByRef value As Dictionary(Of String, DAOChosahyo.調査票項目), koubanFromList As String(), koubanTo As String, Optional isDate As Boolean = False) As Boolean

        '元データが存在する明細番号を抽出
        Dim queryFrom As IEnumerable(Of String) = {}
        For Each koubanFrom In koubanFromList
            queryFrom = queryFrom.Union(
                value.Where(Function(row) row.Key.StartsWith(koubanFrom)).
                Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key))
                )
        Next

        '抽出した資産の数ぶんループする
        Dim list As New List(Of String)(queryFrom)
        For Each edaNo In list
            Dim keyFrom As String = ""
            Dim keyTo As String = koubanTo & edaNo

            '値が空白でない項目番号を選定
            For Each key As String In koubanFromList
                If value.ContainsKey(key & edaNo) Then
                    If Not String.IsNullOrEmpty(value(key & edaNo).値) Then
                        If isDate = False Then
                            '日付でない場合はこの時点で確定
                            keyFrom = key & edaNo
                            Exit For
                        ElseIf value(key & edaNo).値 <> "0" Then
                            '日付の場合、0であった場合採用しない
                            keyFrom = key & edaNo
                            Exit For
                        End If
                    End If
                End If
            Next

            'プレプリント対象の値を代入(Fromがすべて存在しなかった場合、値は変わらない)
            If Not String.IsNullOrEmpty(keyFrom) Then
                If Not value.ContainsKey(keyTo) Then
                    value.Add(keyTo, New DAOChosahyo.調査票項目)
                End If
                value(keyTo).値 = value(keyFrom).値
                value(keyTo).シート名 = value(keyFrom).シート名
            End If
        Next

        Return True
    End Function
    '---REV_001_ADD END

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
        Message.ShowMsgForm(Me, MessageID.MSG_I_015, msgPara)

    End Sub

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

    ''' <summary>
    ''' パラメーターの値は0か
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns>
    ''' TRUE ：0である
    ''' FALSE：0ではない
    ''' </returns>
    ''' <remarks></remarks>
    Private Function IsZero(ByVal value As String) As Boolean

        'ブランクは0扱いとする
        If String.IsNullOrEmpty(value) Then
            Return True
        End If

        '数値でなかったら0扱いとする
        If Not ComUtil.TryParseToDecimal(value) Then
            Return True
        End If

        Return IsZero(CDec(value))

    End Function

    ''' <summary>
    ''' パラメーターの値は0か
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns>
    ''' TRUE ：0である
    ''' FALSE：0ではない
    ''' </returns>
    ''' <remarks></remarks>
    Private Function IsZero(ByVal value As Decimal) As Boolean

        Return If(value = 0, True, False)

    End Function

    ''' <summary>
    ''' 0除算が発生した項目に印字する文字列を返す
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetOutputStringForDivideByZero() As String

        Return ""

    End Function

#End Region

End Class
