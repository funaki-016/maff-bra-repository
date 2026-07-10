Imports Microsoft.Office.Interop

''' <summary>
''' 営農欠測値平均値代入画面
''' </summary>
''' <remarks></remarks>
''' 
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2021.11.02 |日本コンピュータシステム | 要件No1-③対応
'//  REV_002   | 2022.10.11 |daiko               | 要件No1  バージョン区分追加
'//  REV_003   | 2022.01.04 |daiko               | 要件No6  欠測値補完機能に係る修正
'//            |            |                    |
'//*************************************************************************************************
Public Class BRA5810F

#Region "【処理詳細仕様 1】初期表示"

    ''' <summary>
    ''' フォームのロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA5810F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '【処理詳細仕様 1-2】調査年（産）設定

            '調査年コンボボックス設定
            ComUtil.SetChosaNenComboBox(cboChosaNen, ComConst.欠測値.営農欠測値適用平均値データ)

            '更新日時設定
            ComUtil.UpdateDate(txtUpdateDate, cboChosaNen.Text, ComConst.欠測値.営農欠測値平均値代入結果)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

#End Region

#Region "【処理詳細仕様 2】「調査年（産）」コンボボックス選択"

    ''' <summary>
    ''' 「調査年（産）」コンボボックスの選択を変更した
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboChosaNen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboChosaNen.SelectedIndexChanged

        '更新日時設定
        ComUtil.UpdateDate(txtUpdateDate, cboChosaNen.Text, ComConst.欠測値.営農欠測値平均値代入結果)

    End Sub

#End Region

#Region "【処理詳細仕様 3】「代入」ボタンクリック"

    ''' <summary>
    ''' 代入ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSubstitute_Click(sender As Object, e As EventArgs) Handles btnSubstitute.Click

        Try
            If String.IsNullOrWhiteSpace(cboChosaNen.Text) Then
                Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
                Return
            End If

            If Message.ShowMsgBox(MessageID.MSG_Q_024, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If

            Dim progressDialog As New ProgressDialog()

            Using db As New DBAccess(My.Settings.BRAHConnectionString)
                Try
                    db.BeginTrans()

                    '【処理詳細仕様 3-1】調査年データ削除
                    DeleteChosaNenData(db)

                    '【処理詳細仕様 3-2】個別結果表データコピー
                    CopyKobetsuKekkaHyoData(db)

                    '【処理詳細仕様 3-3】代入対象検索
                    Dim syuukeiTable = DAOKobetsuKekkahyo.SelectKobetsuKekkaHyoSyuukeiTableWhereTaisyakuKubun3(db, cboChosaNen.Text)

                    '進捗ダイアログを表示する(処理詳細仕様 3-4～3-7をカウントの対象とする)
                    progressDialog.Maximum = syuukeiTable.Rows.Count * 4
                    progressDialog.Show(Me)

                    For Each tableRow As DataRow In syuukeiTable.Rows

                        '【処理詳細仕様 3-4】代入対象平均値データ検索
                        Dim 営農欠測値適用平均値データRow = SelectAverageData(tableRow)

                        'REV_003 START---------------
                        'データが検索できない場合、全国地を検索
                        If 営農欠測値適用平均値データRow Is Nothing OrElse 営農欠測値適用平均値データRow.集計数 = 0 Then
                            営農欠測値適用平均値データRow = SelectAverageDataZenkoku(tableRow)
                        End If
                        'REV_003 END-----------------

                        '進捗を進める
                        progressDialog.AddValue = 1

                        Dim censusNo = GetCensusNo(tableRow)
                        Dim kKey As New DAOKobetsuKekkahyo.KyotenKey
                        kKey = GetPkey(tableRow)

                        'データが検索できたか
                        If Not 営農欠測値適用平均値データRow Is Nothing AndAlso 1 <= 営農欠測値適用平均値データRow.集計数 Then

                            '【処理詳細仕様 3-5】平均値代入
                            DAOKobetsuKekkahyo.UpdateKobetsuKekkaHyo1SyuukeiTable(db, 営農欠測値適用平均値データRow, censusNo)

                            '進捗を進める
                            progressDialog.AddValue = 1

                            '【処理詳細仕様 3-6】個別結果表再計算
                            Recalculate(db, censusNo, kKey)

                            '進捗を進める
                            progressDialog.AddValue = 1

                        Else
                            '進捗を進める
                            progressDialog.AddValue = 2
                        End If

                        '【処理詳細仕様 3-7】平均値代入結果登録
                        InsertAverageResultTable(db, 営農欠測値適用平均値データRow, censusNo)

                        '進捗を進める
                        progressDialog.AddValue = 1
                    Next

                    '進捗ダイアログを閉じる
                    progressDialog.endDispose()

                    db.CommitTrans()

                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                Finally
                    If Not progressDialog Is Nothing Then
                        '進捗ダイアログを閉じる
                        progressDialog.endDispose()
                        progressDialog = Nothing
                    End If
                End Try
            End Using

            '【処理詳細仕様 3-8】完了メッセージ表示
            Message.ShowMsgBox(MessageID.MSG_I_024, MsgBoxStyle.OkOnly)

            '更新日時設定
            ComUtil.UpdateDate(txtUpdateDate, cboChosaNen.Text, ComConst.欠測値.営農欠測値平均値代入結果)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try

    End Sub

    ''' <summary>
    ''' 調査年データ削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <remarks></remarks>
    Private Sub DeleteChosaNenData(ByVal db As DBAccess)

        If db Is Nothing Then
            Return
        End If

        DAOOther.DeleteTable(db, cboChosaNen.Text, ComConst.欠測値.営農欠測値平均値代入結果)
        DAOOther.DeleteTable(db, cboChosaNen.Text, ComConst.欠測値.個別結果表＿農業経営＿営農類型＿個人１＿集計用)
        DAOOther.DeleteTable(db, cboChosaNen.Text, ComConst.欠測値.個別結果表＿農業経営＿営農類型＿個人２＿集計用)
        DAOOther.DeleteTable(db, cboChosaNen.Text, ComConst.欠測値.個別結果表＿農業経営＿営農類型＿個人３＿集計用)
        DAOOther.DeleteTable(db, cboChosaNen.Text, ComConst.欠測値.個別結果表＿農業経営＿営農類型＿個人４＿集計用)
        DAOOther.DeleteTable(db, cboChosaNen.Text, ComConst.欠測値.個別結果表＿農業経営＿営農類型＿個人５＿集計用)
        DAOOther.DeleteTable(db, cboChosaNen.Text, ComConst.欠測値.個別結果表＿農業経営＿営農類型＿個人６＿集計用)

    End Sub

    ''' <summary>
    ''' 個別結果表データコピー
    ''' </summary>
    ''' <param name="db"></param>
    ''' <remarks></remarks>
    Private Sub CopyKobetsuKekkaHyoData(ByVal db As DBAccess)

        If db Is Nothing Then
            Return
        End If

        '「個別結果表＿農業経営＿営農類型＿個人１～６」の各テーブルを操作する
        For i = 0 To ComConst.個別結果表.テーブル名称(ComConst.調査区分.営農類型別経営統計_個人).Count - 1

            DAOKobetsuKekkahyo.CopyTable(db, cboChosaNen.Text, i)
        Next

    End Sub

    ''' <summary>
    ''' 代入対象平均値データを検索する
    ''' </summary>
    ''' <param name="tableRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SelectAverageData(ByVal tableRow As DataRow) As ComUtil.営農欠測値適用平均値データ

        If tableRow Is Nothing Then
            Return Nothing
        End If

        Dim 農業地域コード = ComConst.欠測値.Get農業地域コード(ComUtil.GetNullOrInt(tableRow("K000002")))
        Dim 営農類型コード = ComUtil.GetNullOrInt(tableRow("K000005"))
        Dim 営農規模コード = ComUtil.GetNullOrInt(tableRow("K000006"))

        If 農業地域コード Is Nothing OrElse
            営農類型コード Is Nothing OrElse
            営農規模コード Is Nothing Then

            Return Nothing
        End If

        Return ComUtil.GetAverageDataRowList(cboChosaNen.Text,
                                             String.Format("AND 農業地域 = {0}", 農業地域コード) & " " &
                                             String.Format("AND 営農類型 = {0}", 営農類型コード) & " " &
                                             String.Format("AND 営農規模 = {0}", 営農規模コード)).FirstOrDefault()

    End Function

    ''' <summary>
    ''' 代入対象平均値データの全国地を検索する
    ''' </summary>
    ''' <param name="tableRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SelectAverageDataZenkoku(ByVal tableRow As DataRow) As ComUtil.営農欠測値適用平均値データ

        If tableRow Is Nothing Then
            Return Nothing
        End If

        Dim 営農類型コード = ComUtil.GetNullOrInt(tableRow("K000005"))
        Dim 営農規模コード = ComUtil.GetNullOrInt(tableRow("K000006"))

        If 営農類型コード Is Nothing OrElse
            営農規模コード Is Nothing Then

            Return Nothing
        End If

        Return ComUtil.GetAverageDataRowList(cboChosaNen.Text,
                                             "AND 農業地域 = 0 " &
                                             String.Format("AND 営農類型 = {0}", 営農類型コード) & " " &
                                             String.Format("AND 営農規模 = {0}", 営農規模コード)).FirstOrDefault()

    End Function

    ''' <summary>
    ''' センサス番号を取得する
    ''' </summary>
    ''' <param name="tableRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCensusNo(ByVal tableRow As DataRow) As String

        Return ComUtil.GetBlankOrString(tableRow("センサス番号"))

    End Function

    ''' <summary>
    ''' 農政局、都道府県、実査設置拠点を取得する
    ''' </summary>
    ''' <param name="tableRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetPkey(ByVal tableRow As DataRow) As DAOKobetsuKekkahyo.KyotenKey

        Dim kKey As New DAOKobetsuKekkahyo.KyotenKey

        Dim 農政局 = ComUtil.GetBlankOrString(tableRow("農政局"))
        Dim 都道府県 = ComUtil.GetBlankOrString(tableRow("都道府県"))
        Dim 実査設置拠点 = ComUtil.GetBlankOrString(tableRow("実査設置拠点"))

        kKey = New DAOKobetsuKekkahyo.KyotenKey(If(農政局 Is Nothing, String.Empty, 農政局), _
                                  If(都道府県 Is Nothing, String.Empty, 都道府県), _
                                  If(実査設置拠点 Is Nothing, String.Empty, 実査設置拠点))

        Return kKey
    End Function

    ''' <summary>
    ''' 平均値代入結果を登録する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="営農欠測値適用平均値データRow"></param>
    ''' <param name="censusNo"></param>
    ''' <remarks></remarks>
    Private Sub InsertAverageResultTable(ByVal db As DBAccess, ByVal 営農欠測値適用平均値データRow As ComUtil.営農欠測値適用平均値データ, ByVal censusNo As String)

        If db Is Nothing Then
            Return
        End If

        If String.IsNullOrWhiteSpace(censusNo) Then
            Return
        End If

        Dim 営農欠測値平均値代入結果Row As New ComUtil.営農欠測値平均値代入結果

        Cast(営農欠測値適用平均値データRow, censusNo, 営農欠測値平均値代入結果Row)

        DAOOther.InsertAverageResultTable(db, 営農欠測値平均値代入結果Row)

    End Sub

    ''' <summary>
    ''' 営農欠測値適用平均値クラスを営農欠測値平均値代入結果クラスに変換する
    ''' </summary>
    ''' <param name="営農欠測値適用平均値データRow"></param>
    ''' <param name="censusNo"></param>
    ''' <param name="営農欠測値平均値代入結果Row"></param>
    ''' <remarks></remarks>
    Private Sub Cast(ByVal 営農欠測値適用平均値データRow As ComUtil.営農欠測値適用平均値データ, ByVal censusNo As String, ByRef 営農欠測値平均値代入結果Row As ComUtil.営農欠測値平均値代入結果)

        If String.IsNullOrWhiteSpace(censusNo) Then
            Return
        End If

        If 営農欠測値平均値代入結果Row Is Nothing Then

            営農欠測値平均値代入結果Row = New ComUtil.営農欠測値平均値代入結果

        End If

        営農欠測値平均値代入結果Row.センサス番号 = censusNo

        'DBから平均値データが取得できたか
        If 営農欠測値適用平均値データRow IsNot Nothing Then

            営農欠測値平均値代入結果Row.調査年 = 営農欠測値適用平均値データRow.調査年
            営農欠測値平均値代入結果Row.営農類型 = 営農欠測値適用平均値データRow.営農類型
            営農欠測値平均値代入結果Row.営農規模 = 営農欠測値適用平均値データRow.営農規模

            '集計されたデータか
            If 1 <= 営農欠測値適用平均値データRow.集計数 Then

                営農欠測値平均値代入結果Row.適用＿営農類型 = 営農欠測値適用平均値データRow.営農類型
                営農欠測値平均値代入結果Row.適用＿営農規模 = 営農欠測値適用平均値データRow.営農規模
                営農欠測値平均値代入結果Row.適用＿農業地域 = 営農欠測値適用平均値データRow.農業地域
            Else
                営農欠測値平均値代入結果Row.適用＿不可 = 1
            End If
        Else
            '平均値データが取得できなかった
            営農欠測値平均値代入結果Row.調査年 = CInt(cboChosaNen.Text)
            営農欠測値平均値代入結果Row.適用＿不可 = 1
        End If

    End Sub

    ''' <summary>
    ''' 再計算処理を行う
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <param name="kkey"></param>
    ''' <remarks></remarks>
    Private Sub Recalculate(ByVal db As DBAccess, ByVal censusNo As String, ByVal kkey As DAOKobetsuKekkahyo.KyotenKey)
        Dim dtChoItemMst As DataTable
        Dim dtKobetsuItemMst As DataTable
        Dim dtCreateRonri As DataTable
        Dim dcKobetsu As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)
        Dim kobetsuList As Dictionary(Of String, Object)
        Dim itemInfoList As List(Of CreateKobetsu.ItemInfo)
        Dim ItemInfoListKobetsu As List(Of CreateKobetsu.ItemInfo)
        Dim dtKobetsu As Dictionary(Of String, DataTable)
        Dim _pKey As DAOKobetsuKekkahyo.PrimaryKey = Nothing
        _pKey = New DAOKobetsuKekkahyo.PrimaryKey(cboChosaNen.Text, censusNo)
        Dim taisyakuKubun As String = Nothing

        Try

            '調査票項目マスタ取得
            'REV-001 START --------------------------
            dtChoItemMst = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _pKey.chosaNen)
            'REV-001 END --------------------------
            '個別結果表項目マスタ取得(裏項番含める)
            ' REV_002↓
            'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, True)
            dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_pKey.chosaNen, CommonInfo.Chosakubun), True)
            ' REV_002↑
            '個別結果表テーブル取得
            dtKobetsu = DAOKobetsuKekkahyo.GetTable(db, _pKey, ComConst.欠測値補完.有)
            '個別結果表項目取得
            dcKobetsu = ComUtil.KobetsuKekkahyo.GetItem(dtKobetsuItemMst, dtKobetsu)
            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                taisyakuKubun = dcKobetsu(ComConst.個別結果表.貸借対照表(ComConst.調査区分.営農類型別経営統計_個人)).値
                '個別結果表作成論理＿営農個人取得(再計算する論理のみ取得する)
                ' REV_002↓
                dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db, taisyakuKubun, ComUtil.getVersionKubunTaikei(_pKey.chosaNen, CommonInfo.Chosakubun), True)
                ' REV_002↑
            Else
                '個別結果表作成論理取得(再計算する論理のみ取得する)
                ' REV_002↓
                'dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonri(db, True)
                dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonri(db, ComUtil.getVersionKubunTaikei(_pKey.chosaNen, CommonInfo.Chosakubun), True)
                ' REV_002↑
            End If

            kobetsuList = New Dictionary(Of String, Object)
            For Each kv As KeyValuePair(Of String, DAOKobetsuKekkahyo.個別結果表項目) In dcKobetsu
                kobetsuList.Add(kv.Key, kv.Value.値)
            Next

            '個別結果表・個別結果検討表作成クラス
            Dim kobetsu As CreateKobetsu = New CreateKobetsu(db,
                                                             CommonInfo.Chosakubun,
                                                             _pKey.chosaNen,
                                                             CreateKobetsu.enmCreateType.個別結果表再計算,
                                                             dtChoItemMst,
                                                             dtKobetsuItemMst,
                                                             Nothing,
                                                             dtCreateRonri,
                                                             kobetsuList,
                                                             Nothing)
            '個別結果表再計算実行
            itemInfoList = kobetsu.Execute(_pKey.censusNo)

            '個別結果表(当年データ、裏項番以外、再計算対象)で抽出
            ItemInfoListKobetsu = (From n In itemInfoList Where n.ItemType = CreateKobetsu.enmItemType.個別結果表 And Not n.ItemNo.Contains("前") And n.IsHidden = False And n.IsReCalc).ToList
            For Each info In ItemInfoListKobetsu
                '個別結果表シートデータを上書き
                If dcKobetsu.ContainsKey(info.ItemNo) Then
                    dcKobetsu(info.ItemNo).値 = If(info.Value Is Nothing Or IsDBNull(info.Value) Or info.Value Is String.Empty, Nothing, info.Value.ToString)
                End If
            Next

            '個別結果表データ削除
            DAOKobetsuKekkahyo.DeleteTable(db, _pKey, kkey, ComConst.欠測値補完.有)

            '個別結果表データ追加
            DAOKobetsuKekkahyo.InsertTable(db, _pKey, kkey, dcKobetsu, ComConst.欠測値補完.有)

        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region "【処理詳細仕様 4】「出力」ボタンクリック"

    ''' <summary>
    ''' 出力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click

        If String.IsNullOrWhiteSpace(cboChosaNen.Text) Then
            Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
            Return
        End If

        '帳票出力データを検索する
        Dim rowList = ComUtil.GetAverageResultTableRowList(cboChosaNen.Text)

        If rowList Is Nothing OrElse Not rowList.Any() Then
            Message.ShowMsgBox(MessageID.MSG_E_023, MsgBoxStyle.OkOnly)
            Return
        End If

        Try
            '【処理詳細仕様 ①】ファイル保存ダイアログを表示する
            Dim fileName As String = ComConst.欠測値.適用状況一覧表出力用ファイル.Report.reportName & "_" & cboChosaNen.Text & ".xlsx"

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            '【処理詳細仕様 ③】帳票を出力する
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New BRA5810R(filePath, cboChosaNen.Text)
                    ret = ExcelOutput.Execute(MessageID.MSG_Q_004)
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
        End Try

    End Sub

#End Region

End Class
