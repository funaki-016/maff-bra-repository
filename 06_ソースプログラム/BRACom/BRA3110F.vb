'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.10.11 |Daiko               | 要件No1 バージョン区分追加
'//  REV_002   | 2025.09.11 |GCU                 | 要件No8 一覧に調査対象経営体と担当名称追加
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 個別結果表作成画面
''' </summary>
''' <remarks></remarks>
Public Class BRA3110F

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

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA3110F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '調査年コンボボックス設定
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.Chosahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
            End Using

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
            _einouRuikei = If(cboEinouRuikei.SelectedValue Is Nothing, Nothing, cboEinouRuikei.SelectedValue.ToString)

            '一覧表示
            Me.ShowList(_chosaNen, _kyoku, _jimusho, _kyoten, _einouRuikei)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 作成ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnMake_Click(sender As Object, e As EventArgs) Handles btnMake.Click
        Dim dtChoItemMst As DataTable
        Dim dtKobetsuItemMst As DataTable
        Dim dtCreateRonri As DataTable
        Dim itemInfoList As List(Of CreateKobetsu.ItemInfo)
        Dim ItemInfoListKobetsu As List(Of CreateKobetsu.ItemInfo)
        Dim pKey As DAOKobetsuKekkahyo.PrimaryKey
        Dim kKey As New DAOKobetsuKekkahyo.KyotenKey
        Dim dcKobetsu As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)
        Dim progressDialog As New ProgressDialog()
        Dim taisyakuKubun As String = Nothing

        Try
            Dim keys As List(Of ValueTuple(Of DAOKobetsuKekkahyo.PrimaryKey, DAOKobetsuKekkahyo.KyotenKey)) = Me.GetKobetsuKekkahyoKey(_chosaNen)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(keys, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_011, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))


                    '調査票項目マスタ取得
                    dtChoItemMst = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)
                    '個別結果表項目マスタ取得(裏項番含める)
                    ' REV_001↓
                    'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, True)
                    dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun), True)
                    ' REV_001↑
                    If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                        taisyakuKubun = ComConst.貸借対照表区分._1
                        '個別結果表作成論理＿営農個人取得
                        ' REV_001↓
                        'dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db, taisyakuKubun)
                        dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db, taisyakuKubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
                        ' REV_001↑
                    Else
                        '個別結果表作成論理取得
                        ' REV_001↓
                        'dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonri(db)
                        dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonri(db, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
                        ' REV_001↑
                    End If
                    '個別結果表作成論理のデータが１件も存在しない場合
                    If dtCreateRonri.Rows.Count = 0 Then
                        'エラーメッセージ
                        Message.ShowMsgBox(MessageID.MSG_E_017, MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If

                    '進捗ダイアログを表示する
                    progressDialog.Maximum = dtCreateRonri.Rows.Count * keys.Count
                    progressDialog.Show(Me)

                    For i As Integer = 0 To dgvList.Rows.Count - 1
                        If Not Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                            Continue For
                        End If

                        pKey = New DAOKobetsuKekkahyo.PrimaryKey(_chosaNen, dgvList.Rows(i).Cells(7).Value.ToString)
                        kKey.kyoku = dgvList.Rows(i).Cells(9).Value.ToString
                        kKey.jimusho = dgvList.Rows(i).Cells(10).Value.ToString
                        kKey.kyoten = dgvList.Rows(i).Cells(11).Value.ToString

                        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                            '営農個人の貸借対照表区分を取得する
                            taisyakuKubun = DAOOther.GetTaisyakuKubunEinokojin(db, _chosaNen, dtChoItemMst, dtKobetsuItemMst, pKey.censusNo)
                            '個別結果表作成論理＿営農個人取得
                            ' REV_001↓
                            dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db, taisyakuKubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
                            ' REV_001↑
                        Else
                            '個別結果表作成論理取得
                            ' REV_001↓
                            'dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonri(db)
                            dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonri(db, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
                            ' REV_001↑
                        End If
                        '個別結果表・個別結果検討表作成クラス
                        Dim kobetsu As CreateKobetsu = New CreateKobetsu(db, CommonInfo.Chosakubun,
                                                                         _chosaNen,
                                                                         CreateKobetsu.enmCreateType.個別結果表作成,
                                                                         dtChoItemMst,
                                                                         dtKobetsuItemMst,
                                                                         Nothing,
                                                                         dtCreateRonri,
                                                                         Nothing,
                                                                         progressDialog)
                        '個別結果表作成実行
                        itemInfoList = kobetsu.Execute(pKey.censusNo)
                        '個別結果表(当年データ、裏項番以外)で抽出
                        ItemInfoListKobetsu = (From n In itemInfoList Where n.ItemType = CreateKobetsu.enmItemType.個別結果表 And Not n.ItemNo.Contains("前") And n.IsHidden = False).ToList
                        dcKobetsu = New Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)
                        For Each info In ItemInfoListKobetsu
                            Dim item As New DAOKobetsuKekkahyo.個別結果表項目
                            With item
                                .シート名 = info.SheetName
                                .行位置 = info.Row
                                .列位置 = info.Col
                                .値 = If(info.Value Is Nothing, Nothing, info.Value.ToString)
                                .型区分 = info.ValueType
                            End With
                            dcKobetsu.Add(info.ItemNo, item)
                        Next

                        Try
                            db.BeginTrans()

                            '個別結果表データ削除
                            DAOKobetsuKekkahyo.DeleteTable(db, pKey, kKey)

                            '個別結果表データ追加
                            DAOKobetsuKekkahyo.InsertTable(db, pKey, kKey, dcKobetsu)

                            db.CommitTrans()

                            '更新日時の更新
                            dgvList.Rows(i).Cells(8).Value = Now.ToString(ComConst.DATETIME_FORMAT)

                        Catch ex As Exception
                            db.RollBackTrans()
                            Throw ex
                        End Try

                        dtCreateRonri.Clear()
                        dtCreateRonri = Nothing
                    Next

                    '進捗ダイアログを閉じる
                    progressDialog.endDispose()
                End Using

                '完了メッセージ
                Message.ShowMsgBox(MessageID.MSG_I_003, MsgBoxStyle.OkOnly)

            End If

        Catch ex As CreateKobetsuException
            '進捗ダイアログを閉じる
            progressDialog.endDispose()
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_036, {ex.CensusNo, ex.ItemNo}, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)

        Finally
            If Not progressDialog Is Nothing Then
                '進捗ダイアログを閉じる
                progressDialog.endDispose()
                progressDialog = Nothing
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 削除ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            Dim keys As List(Of ValueTuple(Of DAOKobetsuKekkahyo.PrimaryKey, DAOKobetsuKekkahyo.KyotenKey))
            keys = Me.GetKobetsuKekkahyoKey(_chosaNen)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(keys, msgId, True) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim key As ValueTuple(Of DAOKobetsuKekkahyo.PrimaryKey, DAOKobetsuKekkahyo.KyotenKey) = keys.First()
            Dim pKey As DAOKobetsuKekkahyo.PrimaryKey = key.Item1
            Dim kKey As DAOKobetsuKekkahyo.KyotenKey = key.Item2

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_003, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                    '個別結果表存在チェック
                    If Not DAOKobetsuKekkahyo.CheckExist(db, New DAOKobetsuKekkahyo.PrimaryKey(pKey.chosaNen, pKey.censusNo), _
                                                             New DAOKobetsuKekkahyo.KyotenKey(kKey.kyoku, kKey.jimusho, kKey.kyoten)) Then
                        'エラーメッセージ
                        Message.ShowMsgBox(MessageID.MSG_E_016, MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If

                    Try
                        db.BeginTrans()

                        '個別結果表データ削除
                        DAOKobetsuKekkahyo.DeleteTable(db, pKey, kKey)

                        db.CommitTrans()

                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try
                End Using

                '完了メッセージ
                Message.ShowMsgBox(MessageID.MSG_I_005, MsgBoxStyle.OkOnly)

                '一覧表示
                Me.ShowList(_chosaNen, _kyoku, _jimusho, _kyoten, _einouRuikei)
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

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
    ''' 農政局コンボボックス選択時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboKyoku_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKyoku.SelectedIndexChanged
        Try
            '拠点コンボボックス設定
            ComUtil.SetKyotenComboBox(cboKyoku, cboKyoten)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 個別結果表キー取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetKobetsuKekkahyoKey(chosaNen As String) As List(Of ValueTuple(Of DAOKobetsuKekkahyo.PrimaryKey, DAOKobetsuKekkahyo.KyotenKey))
        Dim ret As New List(Of ValueTuple(Of DAOKobetsuKekkahyo.PrimaryKey, DAOKobetsuKekkahyo.KyotenKey))

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                Dim pkey As DAOKobetsuKekkahyo.PrimaryKey = New DAOKobetsuKekkahyo.PrimaryKey(chosaNen, dgvList.Rows(i).Cells(7).Value.ToString)
                Dim kkey As DAOKobetsuKekkahyo.KyotenKey = New DAOKobetsuKekkahyo.KyotenKey(dgvList.Rows(i).Cells(9).Value.ToString, dgvList.Rows(i).Cells(10).Value.ToString, dgvList.Rows(i).Cells(11).Value.ToString)
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
        Dim dtChosahyo As DataTable
        Dim dtKobetsu As DataTable
        Dim senmonName As String = ""

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dtChosahyo = DAOChosahyo.GetChosahyoList(db, chosaNen, kyoku, jimusho, kyoten, einouRuikei, "", "")
            dtKobetsu = DAOKobetsuKekkahyo.GetList(db, chosaNen, kyoku, jimusho, kyoten, einouRuikei)
            senmonName = DAOChosahyo.GetSenmonChosainName(db)
        End Using

        dgvList.Rows.Clear()

        For Each row As DataRow In dtChosahyo.Rows
            Dim censusNo As String = row("センサス番号").ToString
            Dim updateDate As String = Nothing

            Dim query = (From dr In dtKobetsu Where dr("センサス番号").ToString = censusNo).Take(1).ToArray
            If query.Any() Then
                updateDate = query(0)("更新日付").ToString
            End If

            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(1).Value = ComUtil.GetTodofuken(censusNo)
            dgvList.Rows(i).Cells(2).Value = ComUtil.GetShikuchoson(censusNo)
            dgvList.Rows(i).Cells(3).Value = ComUtil.GetKyuShikuchoson(censusNo)
            dgvList.Rows(i).Cells(4).Value = ComUtil.GetNogyoShuraku(censusNo)
            dgvList.Rows(i).Cells(5).Value = ComUtil.GetChosaku(censusNo)
            dgvList.Rows(i).Cells(6).Value = ComUtil.GetKyakutaiNo(censusNo)
            dgvList.Rows(i).Cells(7).Value = censusNo
            dgvList.Rows(i).Cells(8).Value = If(String.IsNullOrEmpty(updateDate), updateDate, DateTime.Parse(updateDate).ToString(ComConst.DATETIME_FORMAT))
            dgvList.Rows(i).Cells(9).Value = row("農政局").ToString
            dgvList.Rows(i).Cells(10).Value = row("都道府県").ToString
            dgvList.Rows(i).Cells(11).Value = row("実査設置拠点").ToString
            'REV_002
            dgvList.Rows(i).Cells(12).Value = row("Q00000701").ToString
            If CommonInfo.SenmonChosain Then
                dgvList.Rows(i).Cells(13).Value = senmonName
            Else
                dgvList.Rows(i).Cells(13).Value = row("Q00000801").ToString
            End If
        Next
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="keys"></param>
    ''' <param name="msgId"></param>
    ''' <param name="one"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Function CheckError(keys As List(Of ValueTuple(Of DAOKobetsuKekkahyo.PrimaryKey, DAOKobetsuKekkahyo.KyotenKey)), ByRef msgId As String, Optional one As Boolean = False) As Boolean
        Dim ret As Boolean = False

        'センサス番号選択チェック
        If keys.Count = 0 Then
            msgId = MessageID.MSG_E_003
            Return ret
        End If

        If one Then
            'センサス番号単一選択チェック
            If keys.Count > 1 Then
                msgId = MessageID.MSG_E_038
                Return ret
            End If
        End If

        ret = True

        Return ret
    End Function
End Class
