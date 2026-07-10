'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_002   | 2023.01.04 |大興電子通信        | 要件No.19
'//  REV_003   | 2023.08.07 |大興電子通信        | 要件No.13-① 牛乳生産費の場合に搾乳牛の成畜ボタン活性化
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 牛トレサデータプレプリント（異動データ）出力画面
''' </summary>
''' <remarks></remarks>
Public Class BRA2110F

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
    ''' <summary>牛トレサデータ期間開始年</summary>
    Private _kaishinen As String
    ''' <summary>牛トレサデータ期間開始月</summary>
    Private _kaishituki As String
    ''' <summary>牛トレサデータ期間終了年</summary>
    Private _syuryonen As String
    ''' <summary>牛トレサデータ期間終了月</summary>
    Private _syuryotuki As String

    Private Const YEAR_FROM As String = "0"
    Private Const YEAR_TO As String = "1"
    Private Const MONTH_FROM As String = "0"
    Private Const MONTH_TO As String = "1"

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
    Private Sub BRA2110F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '【処理詳細仕様 1-2】調査年（産）設定

            '調査年コンボボックス設定
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.Chosahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
            End Using

            '【処理詳細仕様 1-3】農政局設定

            '局コンボボックス設定
            ComUtil.SetKyokuComboBox(cboKyoku)

            '営農類型コンボボックス設定
            'ComUtil.SetEinouRuikeiComboBox(lblEinouRuikei, cboEinouRuikei)

            'REV-002 START------------------
            '全所有牛情報出力ボタン表示
            If CommonInfo.Kubun1 = ComConst.区分１.組織法人経営体に関する経営分析調査 Then
                btnMakeAll.Visible = False
            End If
            'REV-002 END------------------

            '↓REV_003
            '牛乳生産費の場合に搾乳牛の成畜ボタン活性化
            If CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 Then
                btnSeitiku.Enabled = True
            End If
            '↑REV_003

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
            Dim CensusNo As String
            Dim fromKikan As String
            Dim toKikan As String
            Dim ChosaNen As String
            Dim lstHoseiList As New List(Of Dictionary(Of String, String))
            Dim hoseiList = New List(Of BRA2120F.HoseiIdouInfo)

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(keys, msgId) Then
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
            fromKikan = txtYearFrom.Text & txtMonthFrom.Text.PadLeft(2, "0"c) & "01"
            Dim theDay = New DateTime(CInt(txtYearTo.Text), CInt(txtMonthTo.Text), 1)
            Dim days As Integer = DateTime.DaysInMonth(theDay.Year, theDay.Month)
            Dim lastDayOfMonth = New DateTime(theDay.Year, theDay.Month, days)
            toKikan = lastDayOfMonth.ToString("yyyyMMdd")

            '画面で選択した調査票の分だけ以下の処理を行う。
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                For i As Integer = 0 To dgvList.Rows.Count - 1

                    If dgvList.Rows(i).Cells(0).Value IsNot Nothing _
                        And CBool(dgvList.Rows(i).Cells(0).Value) = True Then

                        'センサス番号取得
                        CensusNo = dgvList.Rows(i).Cells(7).Value.ToString

                        'プレプリント出力対象牛トレサ情報の取得
                        Dim dtTarget = DAOOther.GetPrePrintTargetData(db, CommonInfo.Chosakubun, ChosaNen, CensusNo, fromKikan, toKikan)

                        '牛トレサ情報から農家団体コードを取得
                        Dim farmCode = dtTarget.DefaultView.ToTable(True, "農家団体コード")
                        Dim farmCodeList As List(Of String) = New List(Of String)

                        For Each row As DataRow In farmCode.Rows
                            farmCodeList.Add(row("農家団体コード").ToString)
                        Next

                        '牛トレサ異動情報取得
                        Dim dtTresa As DataTable = DAOOther.GetTresa(db, farmCodeList, 0, CInt(toKikan))

                        '個体識別番号リスト作成
                        Dim dtKotaiNo As DataTable = dtTarget.DefaultView.ToTable(True, "個体識別番号")

                        'プレプリント出力対象牛トレサ情報の内、補正異動情報一覧画面に表示するデータを選別する。
                        For Each dr As DataRow In dtKotaiNo.Rows
                            Dim dtIdoList As DataRow() = dtTarget.Select("個体識別番号 = '" & dr("個体識別番号").ToString & "'", "異動年月日,異動フラグ ASC")
                            Dim idoHoseiFlg As Boolean = False
                            '個体毎にデータの確認を行う。
                            For j As Integer = 0 To dtIdoList.Length - 1
                                Dim kindCode As Integer = Integer.Parse(dtIdoList(j)("牛の識別CD").ToString)
                                Dim sexCode As Integer = Integer.Parse(dtIdoList(j)("性別コード").ToString)
                                Dim idouFlag As Integer = Integer.Parse(dtIdoList(j)("異動フラグ").ToString)
                                Dim chosaBunrui As ComConst.牛トレサデータ.調査区分分類
                                Dim dicHosei As New Dictionary(Of String, String)

                                If ComConst.牛トレサデータ.調査区分分類変換テーブル.ContainsKey(CommonInfo.Chosakubun) Then
                                    chosaBunrui = ComConst.牛トレサデータ.調査区分分類変換テーブル(CommonInfo.Chosakubun)
                                Else
                                    '対象外の調査区分
                                    Exit For
                                End If
                                If ComUtil.Tresa.IsOrganizeData(chosaBunrui, kindCode, sexCode, idouFlag) = False Then
                                    Continue For
                                End If
                                '異動フラグが3の場合、次のレコードが4ならば異動補正の対象となる
                                If dtIdoList(j)("異動フラグ").ToString = "3" Then
                                    If j <> dtIdoList.Length - 1 Then
                                        If dtIdoList(j + 1)("異動フラグ").ToString = "4" Then
                                            '異動年月日が同日の場合、時系列で判断を行う
                                            If dtIdoList(j)("異動年月日").ToString = dtIdoList(j + 1)("異動年月日").ToString Then
                                                '個体識別番号で牛の存在を確認
                                                Dim tresaRows As DataRow() = dtTresa.Select("個体識別番号 = '" & dtIdoList(j).Item("個体識別番号").ToString & "' and 異動年月日 < '" & dtIdoList(j)("異動年月日").ToString & "'", "異動年月日 DESC")
                                                Dim tresaIndex As Integer = 0

                                                If ComUtil.OrganiezCheck.IsCowOrganize(tresaRows, tresaIndex) = True Then
                                                    idoHoseiFlg = True
                                                End If
                                            Else
                                                idoHoseiFlg = True
                                            End If

                                            If idoHoseiFlg = True Then
                                                dicHosei("センサス番号") = CensusNo
                                                dicHosei("個体識別番号") = dtIdoList(j)("個体識別番号").ToString
                                                dicHosei("転出異動年月日") = dtIdoList(j)("異動年月日").ToString
                                                dicHosei("転出異動フラグ") = dtIdoList(j)("異動フラグ").ToString
                                                dicHosei("転出農家団体コード") = dtIdoList(j)("農家団体コード").ToString
                                                dicHosei("転入異動年月日") = dtIdoList(j + 1).Item("異動年月日").ToString
                                                dicHosei("転入異動フラグ") = dtIdoList(j + 1).Item("異動フラグ").ToString
                                                dicHosei("転入農家団体コード") = dtIdoList(j + 1).Item("農家団体コード").ToString
                                                lstHoseiList.Add(dicHosei)

                                                j = j + 1
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                Next

                '補正異動情報一覧画面の遷移対象データがある場合には、補正異動情報一覧画面を表示する。
                If lstHoseiList.Count <> 0 Then
                    '異動年月日でソートする。
                    Dim sortedList = lstHoseiList.OrderBy(Function(n) n.Item("転出異動年月日")).ToList()

                    For Each dc As Dictionary(Of String, String) In sortedList
                        Dim hosei As New BRA2120F.HoseiIdouInfo
                        hosei.CensusNo = dc("センサス番号")
                        hosei.KotaiNo = dc("個体識別番号")
                        hosei.OutIdoDate = dc("転出異動年月日")
                        hosei.OutIdoFlag = dc("転出異動フラグ")
                        hosei.OutFarmCode = dc("転出農家団体コード")
                        hosei.InIdoDate = dc("転入異動年月日")
                        hosei.InIdoFlag = dc("転入異動フラグ")
                        hosei.InFarmCode = dc("転入農家団体コード")
                        hosei.IsDelete = False
                        hoseiList.Add(hosei)
                    Next

                    Dim frm As New BRA2120F(hoseiList)
                    Me.Hide()
                    frm.ShowDialog(Me)
                    Me.Show()
                End If
            End Using

            Dim ret As ExcelOutputBaseClass.enmOutputReturn

            Try
                Dim pKeys As List(Of DAOChosahyo.PrimaryKey) = Me.GetChosahyoPrimaryKey(_chosaNen)
                Dim filenamechange As Long = 0
                'REV-001 START------------------
                Using ExcelOutput = New BRA2110R(folderPath, txtYearFrom.Text & txtMonthFrom.Text.PadLeft(2, "0"c), txtYearTo.Text & txtMonthTo.Text.PadLeft(2, "0"c), hoseiList, _chosaNen, filenamechange)
                    'REV-001 END------------------
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
            'エラーメッセージ
            Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
            Return ret
        End If

        '画面入力項目単体チェック
        '開始年のチェック
        If YearCheck(txtYearFrom.Text, YEAR_FROM) <> 0 Then
            txtYearFrom.BackColor = Color.Red
            Return ret
        Else
            txtYearFrom.BackColor = Color.White
        End If

        '開始月のチェック
        If MonthCheck(txtMonthFrom.Text, MONTH_FROM) <> 0 Then
            txtMonthFrom.BackColor = Color.Red
            Return ret
        Else
            txtMonthFrom.BackColor = Color.White
        End If

        '終了年のチェック
        If YearCheck(txtYearTo.Text, YEAR_TO) <> 0 Then
            txtYearTo.BackColor = Color.Red
            Return ret
        Else
            txtYearTo.BackColor = Color.White
        End If

        '終了月のチェック
        If MonthCheck(txtMonthTo.Text, MONTH_TO) <> 0 Then
            txtMonthTo.BackColor = Color.Red
            Return ret
        Else
            txtMonthTo.BackColor = Color.White
        End If

        '画面入力項目間チェック
        '開始年月と終了年月の大小チェック
        If DateConsistentCheck(txtYearFrom.Text, txtMonthFrom.Text, txtYearTo.Text, txtMonthTo.Text) <> 0 Then
            txtYearFrom.BackColor = Color.Red
            txtMonthFrom.BackColor = Color.Red
            txtYearTo.BackColor = Color.Red
            txtMonthTo.BackColor = Color.Red
            Return ret
        End If

        ret = True

        Return ret
    End Function

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

#Region "【要件No.22】「搾乳牛の成畜」ボタンクリック"

    ''' <summary>
    ''' 作成ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSeitiku_Click(sender As Object, e As EventArgs) Handles btnSeitiku.Click
        Try
            Dim keys As List(Of ValueTuple(Of DAOChosahyo.PrimaryKey, DAOChosahyo.KotenKey)) = Me.GetChosahyoKey(_chosaNen)
            Dim CensusNo As String
            Dim fromKikan As String
            Dim fromKikan_y As String
            Dim fromKikan_m As String
            Dim toKikan As String
            Dim ChosaNen As String
            Dim lstHoseiList As New List(Of Dictionary(Of String, String))
            Dim hoseiList = New List(Of BRA2120F.HoseiIdouInfo)

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.SeitikuCheckError(keys, msgId) Then
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
            fromKikan = txtYearFrom.Text & txtMonthFrom.Text.PadLeft(2, "0"c) & "01"
            fromKikan_y = txtYearFrom.Text
            fromKikan_m = txtMonthFrom.Text
            Dim theDay = New DateTime(CInt(txtYearTo.Text), CInt(txtMonthTo.Text), 1)
            Dim days As Integer = DateTime.DaysInMonth(theDay.Year, theDay.Month)
            Dim lastDayOfMonth = New DateTime(theDay.Year, theDay.Month, days)
            toKikan = lastDayOfMonth.ToString("yyyyMMdd")

            '画面で選択した調査票の分だけ以下の処理を行う。
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                For i As Integer = 0 To dgvList.Rows.Count - 1

                    If dgvList.Rows(i).Cells(0).Value IsNot Nothing _
                        And CBool(dgvList.Rows(i).Cells(0).Value) = True Then

                        'センサス番号取得
                        CensusNo = dgvList.Rows(i).Cells(7).Value.ToString

                        'プレプリント出力対象牛トレサ情報の取得
                        Dim dtTarget = DAOOther.GetSeitikuPrePrintTargetData(db, CommonInfo.Chosakubun, ChosaNen, CensusNo, fromKikan, fromKikan_y, fromKikan_m, toKikan)
                        'Dim dtTarget = DAOOther.GetPrePrintTargetData(db, CommonInfo.Chosakubun, ChosaNen, CensusNo, fromKikan, toKikan)

                        '牛トレサ情報から農家団体コードを取得
                        Dim farmCode = dtTarget.DefaultView.ToTable(True, "農家団体コード")
                        Dim farmCodeList As List(Of String) = New List(Of String)

                        For Each row As DataRow In farmCode.Rows
                            farmCodeList.Add(row("農家団体コード").ToString)
                        Next

                        '牛トレサ異動情報取得(要件No22不要)
                        'Dim dtTresa As DataTable = DAOOther.GetTresa(db, farmCodeList, 0, CInt(toKikan))

                        '個体識別番号リスト作成
                        Dim dtKotaiNo As DataTable = dtTarget.DefaultView.ToTable(True, "個体識別番号")
#Region "要件No22不要"
                        ''プレプリント出力対象牛トレサ情報の内、補正異動情報一覧画面に表示するデータを選別する。
                        'For Each dr As DataRow In dtKotaiNo.Rows
                        '    Dim dtIdoList As DataRow() = dtTarget.Select("個体識別番号 = '" & dr("個体識別番号").ToString & "'", "異動年月日,異動フラグ ASC")
                        '    Dim idoHoseiFlg As Boolean = False
                        '    '個体毎にデータの確認を行う。
                        '    For j As Integer = 0 To dtIdoList.Length - 1
                        '        Dim kindCode As Integer = Integer.Parse(dtIdoList(j)("牛の識別CD").ToString)
                        '        Dim sexCode As Integer = Integer.Parse(dtIdoList(j)("性別コード").ToString)
                        '        Dim idouFlag As Integer = Integer.Parse(dtIdoList(j)("異動フラグ").ToString)
                        '        Dim chosaBunrui As ComConst.牛トレサデータ.調査区分分類
                        '        Dim dicHosei As New Dictionary(Of String, String)

                        '        If ComConst.牛トレサデータ.調査区分分類変換テーブル.ContainsKey(CommonInfo.Chosakubun) Then
                        '            chosaBunrui = ComConst.牛トレサデータ.調査区分分類変換テーブル(CommonInfo.Chosakubun)
                        '        Else
                        '            '対象外の調査区分
                        '            Exit For
                        '        End If
                        '        If ComUtil.Tresa.IsOrganizeData(chosaBunrui, kindCode, sexCode, idouFlag) = False Then
                        '            Continue For
                        '        End If
                        '        '異動フラグが3の場合、次のレコードが4ならば異動補正の対象となる
                        '        If dtIdoList(j)("異動フラグ").ToString = "3" Then
                        '            If j <> dtIdoList.Length - 1 Then
                        '                If dtIdoList(j + 1)("異動フラグ").ToString = "4" Then
                        '                    '異動年月日が同日の場合、時系列で判断を行う
                        '                    If dtIdoList(j)("異動年月日").ToString = dtIdoList(j + 1)("異動年月日").ToString Then
                        '                        '個体識別番号で牛の存在を確認
                        '                        Dim tresaRows As DataRow() = dtTresa.Select("個体識別番号 = '" & dtIdoList(j).Item("個体識別番号").ToString & "' and 異動年月日 < '" & dtIdoList(j)("異動年月日").ToString & "'", "異動年月日 DESC")
                        '                        Dim tresaIndex As Integer = 0

                        '                        If ComUtil.OrganiezCheck.IsCowOrganize(tresaRows, tresaIndex) = True Then
                        '                            idoHoseiFlg = True
                        '                        End If
                        '                    Else
                        '                        idoHoseiFlg = True
                        '                    End If

                        '                    If idoHoseiFlg = True Then
                        '                        dicHosei("センサス番号") = CensusNo
                        '                        dicHosei("個体識別番号") = dtIdoList(j)("個体識別番号").ToString
                        '                        dicHosei("転出異動年月日") = dtIdoList(j)("異動年月日").ToString
                        '                        dicHosei("転出異動フラグ") = dtIdoList(j)("異動フラグ").ToString
                        '                        dicHosei("転出農家団体コード") = dtIdoList(j)("農家団体コード").ToString
                        '                        dicHosei("転入異動年月日") = dtIdoList(j + 1).Item("異動年月日").ToString
                        '                        dicHosei("転入異動フラグ") = dtIdoList(j + 1).Item("異動フラグ").ToString
                        '                        dicHosei("転入農家団体コード") = dtIdoList(j + 1).Item("農家団体コード").ToString
                        '                        lstHoseiList.Add(dicHosei)

                        '                        j = j + 1
                        '                    End If
                        '                End If
                        '            End If
                        '        End If
                        '    Next
                        'Next
#End Region
                    End If
                Next

                ''補正異動情報一覧画面の遷移対象データがある場合には、補正異動情報一覧画面を表示する。（要件No22不要）
                'If lstHoseiList.Count <> 0 Then
                '    '異動年月日でソートする。
                '    Dim sortedList = lstHoseiList.OrderBy(Function(n) n.Item("転出異動年月日")).ToList()

                '    For Each dc As Dictionary(Of String, String) In sortedList
                '        Dim hosei As New BRA2120F.HoseiIdouInfo
                '        hosei.CensusNo = dc("センサス番号")
                '        hosei.KotaiNo = dc("個体識別番号")
                '        hosei.OutIdoDate = dc("転出異動年月日")
                '        hosei.OutIdoFlag = dc("転出異動フラグ")
                '        hosei.OutFarmCode = dc("転出農家団体コード")
                '        hosei.InIdoDate = dc("転入異動年月日")
                '        hosei.InIdoFlag = dc("転入異動フラグ")
                '        hosei.InFarmCode = dc("転入農家団体コード")
                '        hosei.IsDelete = False
                '        hoseiList.Add(hosei)
                '    Next

                '    Dim frm As New BRA2120F(hoseiList)
                '    Me.Hide()
                '    frm.ShowDialog(Me)
                '    Me.Show()
                'End If
            End Using

            Dim ret As ExcelOutputBaseClass.enmOutputReturn

            Try
                Dim pKeys As List(Of DAOChosahyo.PrimaryKey) = Me.GetSeitikuChosahyoPrimaryKey(_chosaNen)
                Dim filenamechange As Long = 1

                'REV-001 START------------------
                Using ExcelOutput = New BRA2110R(folderPath, txtYearFrom.Text & txtMonthFrom.Text.PadLeft(2, "0"c), txtYearTo.Text & txtMonthTo.Text.PadLeft(2, "0"c), hoseiList, _chosaNen, filenamechange)
                    'REV-001 END------------------
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
    ''' エラーチェック
    ''' </summary>
    ''' <param name="keys"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Function SeitikuCheckError(keys As List(Of ValueTuple(Of DAOChosahyo.PrimaryKey, DAOChosahyo.KotenKey)), ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'センサス番号選択チェック
        If keys.Count = 0 Then
            msgId = MessageID.MSG_E_003
            'エラーメッセージ
            Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
            Return ret
        End If

        '調査区分チェック
        If CommonInfo.Chosakubun <> "34" And CommonInfo.Chosakubun <> "18" Then
            msgId = MessageID.MSG_E_003
            'エラーメッセージ
            Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
            Return ret
        End If

        '画面入力項目単体チェック
        '開始年のチェック
        If YearCheck(txtYearFrom.Text, YEAR_FROM) <> 0 Then
            txtYearFrom.BackColor = Color.Red
            Return ret
        Else
            txtYearFrom.BackColor = Color.White
        End If

        '開始月のチェック
        If MonthCheck(txtMonthFrom.Text, MONTH_FROM) <> 0 Then
            txtMonthFrom.BackColor = Color.Red
            Return ret
        Else
            txtMonthFrom.BackColor = Color.White
        End If

        '終了年のチェック
        If YearCheck(txtYearTo.Text, YEAR_TO) <> 0 Then
            txtYearTo.BackColor = Color.Red
            Return ret
        Else
            txtYearTo.BackColor = Color.White
        End If

        '終了月のチェック
        If MonthCheck(txtMonthTo.Text, MONTH_TO) <> 0 Then
            txtMonthTo.BackColor = Color.Red
            Return ret
        Else
            txtMonthTo.BackColor = Color.White
        End If

        '画面入力項目間チェック
        '開始年月と終了年月の大小チェック
        If DateConsistentCheck(txtYearFrom.Text, txtMonthFrom.Text, txtYearTo.Text, txtMonthTo.Text) <> 0 Then
            txtYearFrom.BackColor = Color.Red
            txtMonthFrom.BackColor = Color.Red
            txtYearTo.BackColor = Color.Red
            txtMonthTo.BackColor = Color.Red
            Return ret
        End If

        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' 調査票主キー取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSeitikuChosahyoPrimaryKey(chosaNen As String) As List(Of DAOChosahyo.PrimaryKey)
        Dim ret As New List(Of DAOChosahyo.PrimaryKey)

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                Dim pkey As DAOChosahyo.PrimaryKey = New DAOChosahyo.PrimaryKey(chosaNen, dgvList.Rows(i).Cells(7).Value.ToString)
                ret.Add(pkey)
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 画面で指定された調査年の調査票を取得する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSeitikuNowChosahyo(ByVal db As DBAccess, ByVal censusNo As String) As Dictionary(Of String, DAOChosahyo.調査票項目)


        '調査票項目マスタ取得
        Dim chosahyoItemMaster = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)

        '調査票テーブル取得
        Dim chosahyoTable = DAOChosahyo.GetChosahyoTable(db, New DAOChosahyo.PrimaryKey(_chosaNen, censusNo))

        '調査票項目取得
        Return ComUtil.Chosahyo.GetItem(chosahyoItemMaster, chosahyoTable)

    End Function

#End Region

#Region "「全所有牛情報出力」ボタンクリック"

    'REV-002 START------------------
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
            Dim hoseiList = New List(Of BRA2120F.HoseiIdouInfo)

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(keys, msgId) Then
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
            fromKikan = txtYearFrom.Text & txtMonthFrom.Text.PadLeft(2, "0"c) & "01"
            Dim theDay = New DateTime(CInt(txtYearTo.Text), CInt(txtMonthTo.Text), 1)
            Dim days As Integer = DateTime.DaysInMonth(theDay.Year, theDay.Month)
            Dim lastDayOfMonth = New DateTime(theDay.Year, theDay.Month, days)
            toKikan = lastDayOfMonth.ToString("yyyyMMdd")

            '画面で選択した調査票の分だけ以下の処理を行う。
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                For i As Integer = 0 To dgvList.Rows.Count - 1

                    If dgvList.Rows(i).Cells(0).Value IsNot Nothing _
                        And CBool(dgvList.Rows(i).Cells(0).Value) = True Then

                        'センサス番号取得
                        CensusNo = dgvList.Rows(i).Cells(7).Value.ToString

                        'プレプリント出力対象牛トレサ情報の取得
                        Dim dtTarget = DAOOther.GetPrePrintTargetData(db, CommonInfo.Chosakubun, ChosaNen, CensusNo, fromKikan, toKikan)

                        '牛トレサ情報から農家団体コードを取得
                        Dim farmCode = dtTarget.DefaultView.ToTable(True, "農家団体コード")
                        Dim farmCodeList As List(Of String) = New List(Of String)

                        For Each row As DataRow In farmCode.Rows
                            farmCodeList.Add(row("農家団体コード").ToString)
                        Next

                        '牛トレサ異動情報取得
                        Dim dtTresa As DataTable = DAOOther.GetTresa(db, farmCodeList, 0, CInt(toKikan))

                        '個体識別番号リスト作成
                        Dim dtKotaiNo As DataTable = dtTarget.DefaultView.ToTable(True, "個体識別番号")

                        'プレプリント出力対象牛トレサ情報の内、補正異動情報一覧画面に表示するデータを選別する。
                        For Each dr As DataRow In dtKotaiNo.Rows
                            Dim dtIdoList As DataRow() = dtTarget.Select("個体識別番号 = '" & dr("個体識別番号").ToString & "'", "異動年月日,異動フラグ ASC")
                            Dim idoHoseiFlg As Boolean = False
                            '個体毎にデータの確認を行う。
                            For j As Integer = 0 To dtIdoList.Length - 1
                                Dim kindCode As Integer = Integer.Parse(dtIdoList(j)("牛の識別CD").ToString)
                                Dim sexCode As Integer = Integer.Parse(dtIdoList(j)("性別コード").ToString)
                                Dim idouFlag As Integer = Integer.Parse(dtIdoList(j)("異動フラグ").ToString)
                                Dim chosaBunrui As ComConst.牛トレサデータ.調査区分分類
                                Dim dicHosei As New Dictionary(Of String, String)

                                If ComConst.牛トレサデータ.調査区分分類変換テーブル.ContainsKey(CommonInfo.Chosakubun) Then
                                    chosaBunrui = ComConst.牛トレサデータ.調査区分分類変換テーブル(CommonInfo.Chosakubun)
                                Else
                                    '対象外の調査区分
                                    Exit For
                                End If
                                '異動フラグが3の場合、次のレコードが4ならば異動補正の対象となる
                                If dtIdoList(j)("異動フラグ").ToString = "3" Then
                                    If j <> dtIdoList.Length - 1 Then
                                        If dtIdoList(j + 1)("異動フラグ").ToString = "4" Then
                                            '異動年月日が同日の場合、時系列で判断を行う
                                            If dtIdoList(j)("異動年月日").ToString = dtIdoList(j + 1)("異動年月日").ToString Then
                                                '個体識別番号で牛の存在を確認
                                                Dim tresaRows As DataRow() = dtTresa.Select("個体識別番号 = '" & dtIdoList(j).Item("個体識別番号").ToString & "' and 異動年月日 < '" & dtIdoList(j)("異動年月日").ToString & "'", "異動年月日 DESC")
                                                Dim tresaIndex As Integer = 0

                                            Else
                                                idoHoseiFlg = True
                                            End If

                                            If idoHoseiFlg = True Then
                                                dicHosei("センサス番号") = CensusNo
                                                dicHosei("個体識別番号") = dtIdoList(j)("個体識別番号").ToString
                                                dicHosei("転出異動年月日") = dtIdoList(j)("異動年月日").ToString
                                                dicHosei("転出異動フラグ") = dtIdoList(j)("異動フラグ").ToString
                                                dicHosei("転出農家団体コード") = dtIdoList(j)("農家団体コード").ToString
                                                dicHosei("転入異動年月日") = dtIdoList(j + 1).Item("異動年月日").ToString
                                                dicHosei("転入異動フラグ") = dtIdoList(j + 1).Item("異動フラグ").ToString
                                                dicHosei("転入農家団体コード") = dtIdoList(j + 1).Item("農家団体コード").ToString
                                                lstHoseiList.Add(dicHosei)

                                                j = j + 1
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                Next

                '補正異動情報一覧画面の遷移対象データがある場合には、補正異動情報一覧画面を表示する。
                If lstHoseiList.Count <> 0 Then
                    '異動年月日でソートする。
                    Dim sortedList = lstHoseiList.OrderBy(Function(n) n.Item("転出異動年月日")).ToList()

                    For Each dc As Dictionary(Of String, String) In sortedList
                        Dim hosei As New BRA2120F.HoseiIdouInfo
                        hosei.CensusNo = dc("センサス番号")
                        hosei.KotaiNo = dc("個体識別番号")
                        hosei.OutIdoDate = dc("転出異動年月日")
                        hosei.OutIdoFlag = dc("転出異動フラグ")
                        hosei.OutFarmCode = dc("転出農家団体コード")
                        hosei.InIdoDate = dc("転入異動年月日")
                        hosei.InIdoFlag = dc("転入異動フラグ")
                        hosei.InFarmCode = dc("転入農家団体コード")
                        hosei.IsDelete = False
                        hoseiList.Add(hosei)
                    Next

                    Dim frm As New BRA2120F(hoseiList)
                    Me.Hide()
                    frm.ShowDialog(Me)
                    Me.Show()
                End If
            End Using

            Dim ret As ExcelOutputBaseClass.enmOutputReturn

            Try
                Dim pKeys As List(Of DAOChosahyo.PrimaryKey) = Me.GetChosahyoPrimaryKey(_chosaNen)
                Dim filenamechange As Long = 2
                'REV-001 START------------------
                Using ExcelOutput = New BRA2110R(folderPath, txtYearFrom.Text & txtMonthFrom.Text.PadLeft(2, "0"c), txtYearTo.Text & txtMonthTo.Text.PadLeft(2, "0"c), hoseiList, _chosaNen, filenamechange)
                    'REV-001 END------------------
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
    'REV-002 END------------------

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
    ''' 開始年/終了年 個別チェック
    ''' </summary>
    ''' <param name="txtYear">チェックする年の値</param>
    ''' <param name="iType">0:開始年、1:終了年</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function YearCheck(strYear As String, strType As String) As Integer
        Dim iSystemYear As Integer
        Dim iYear As Integer
        Dim iRet As Integer = -1

        Try
            '年が空の場合は、未入力エラー
            If strYear.Length = 0 Then
                If strType = YEAR_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_069, MsgBoxStyle.OkOnly)
                ElseIf strType = YEAR_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_073, MsgBoxStyle.OkOnly)
                End If
                Return -1
            End If

            'システム年を数値に変換
            iSystemYear = CInt(DateTime.Now.ToString("yyyy"))

            If Integer.TryParse(strYear, iYear) Then
                '入力年を数値に変換
                iYear = CInt(strYear)

                '入力年がシステム年-3～システム年+1の範囲かをチェック
                If iYear >= iSystemYear - 3 And iYear <= iSystemYear + 1 Then
                    iRet = 0
                Else
                    If strType = YEAR_FROM Then
                        Message.ShowMsgBox(MessageID.MSG_E_070, {(iSystemYear - 3).ToString, (iSystemYear + 1).ToString}, MsgBoxStyle.OkOnly)
                    ElseIf strType = YEAR_TO Then
                        Message.ShowMsgBox(MessageID.MSG_E_074, {(iSystemYear - 3).ToString, (iSystemYear + 1).ToString}, MsgBoxStyle.OkOnly)
                    End If
                    iRet = -1
                End If
            Else
                '入力内容が数値でない場合もエラー
                If strType = YEAR_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_070, {(iSystemYear - 3).ToString, (iSystemYear + 1).ToString}, MsgBoxStyle.OkOnly)
                ElseIf strType = YEAR_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_074, {(iSystemYear - 3).ToString, (iSystemYear + 1).ToString}, MsgBoxStyle.OkOnly)
                End If
                iRet = -1
            End If

        Catch ex As Exception
            Throw ex
        End Try
        Return iRet
    End Function

    ''' <summary>
    ''' 開始月/終了月 個別チェック
    ''' </summary>
    ''' <param name="txtMonth">チェックする月の値</param>
    ''' <param name="iType">0:開始月、1:終了月</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function MonthCheck(strMonth As String, strType As String) As Integer
        Dim iMonth As Integer
        Dim iRet As Integer = -1

        Try
            '月が空の場合は、未入力エラー
            If strMonth.Length = 0 Then
                If strType = MONTH_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_071, MsgBoxStyle.OkOnly)
                ElseIf strType = MONTH_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_075, MsgBoxStyle.OkOnly)
                End If
                Return -1
            End If

            If strMonth.Length = 2 AndAlso strMonth.Substring(0, 1) = "0" Then
                If strType = MONTH_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_072, MsgBoxStyle.OkOnly)
                ElseIf strType = MONTH_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_076, MsgBoxStyle.OkOnly)
                End If
                Return -1
            End If

            If Integer.TryParse(strMonth, iMonth) Then
                If iMonth >= 1 And iMonth <= 12 Then
                    iRet = 0
                Else
                    If strType = MONTH_FROM Then
                        Message.ShowMsgBox(MessageID.MSG_E_072, MsgBoxStyle.OkOnly)
                    ElseIf strType = MONTH_TO Then
                        Message.ShowMsgBox(MessageID.MSG_E_076, MsgBoxStyle.OkOnly)
                    End If
                    iRet = -1
                End If
            Else
                '入力内容が数値でない場合もエラー
                If strType = MONTH_FROM Then
                    Message.ShowMsgBox(MessageID.MSG_E_072, MsgBoxStyle.OkOnly)
                ElseIf strType = MONTH_TO Then
                    Message.ShowMsgBox(MessageID.MSG_E_076, MsgBoxStyle.OkOnly)
                End If
                iRet = -1
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return iRet
    End Function

    ''' <summary>
    ''' 開始/終了年月間整合性チェック
    ''' </summary>
    ''' <param name="txtYearFrom">開始年の値</param>
    ''' <param name="txtMonthFrom">開始月の値</param>
    ''' <param name="txtYearTo">終了年の値</param>
    ''' <param name="txtMonthTo">終了月の値</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DateConsistentCheck(strYearFrom As String, strMonthFrom As String, strYearTo As String, strMonthTo As String) As Integer
        Dim iRet As Integer = -1
        Dim strFrom As String
        Dim strTo As String

        Try
            strFrom = strYearFrom & strMonthFrom.PadLeft(2, "0"c)
            strTo = strYearTo & strMonthTo.PadLeft(2, "0"c)

            If strFrom > strTo Then
                Message.ShowMsgBox(MessageID.MSG_E_083, MsgBoxStyle.OkOnly)
                iRet = -1
            Else
                iRet = 0
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return iRet
    End Function
#End Region

End Class
