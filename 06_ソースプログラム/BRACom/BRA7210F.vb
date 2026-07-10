''' <summary>
''' 調査票・個別結果表受信画面
''' </summary>
''' <remarks></remarks>
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.10.26 |TSP)                | フェーズ3 要件No.6修正
'//            |            |                    |
'//*************************************************************************************************
Public Class BRA7210F

    ''' <summary>上り下り区分</summary>
    Private _upLow As String
    ''' <summary>送受信データ種別</summary>
    Private _dataType As String

    ''' <summary>調査年</summary>
    Private _chosaNen As String
    ''' <summary>局</summary>
    Private _kyoku As String
    ''' <summary>事務所</summary>
    Private _jimusho As String
    ''' <summary>拠点</summary>
    Private _kyoten As String

    ''' <summary>進捗ダイアログ</summary>
    Private ProgressDialog As ProgressDialog

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(upLow As String, dataType As String)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

        _upLow = upLow

        _dataType = dataType
    End Sub

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA7210F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '調査年コンボボックス設定（受信）
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.SendReceiveManagement.SetChosaNenJushinComboBox(cboChosaNen, db, CommonInfo.Koutei, _upLow, _dataType, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
            End Using

            '局コンボボックス設定
            ComUtil.SetKyokuComboBox(cboKyoku)

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
            _kyoku = If(IsDBNull(cboKyoku.SelectedValue), Nothing, cboKyoku.SelectedValue.ToString)
            _jimusho = If(IsDBNull(cboKyoten.SelectedValue), Nothing, CStr(CType(cboKyoten.SelectedItem, DataRowView)("事務所番号")))
            _kyoten = If(IsDBNull(cboKyoten.SelectedValue), Nothing, cboKyoten.SelectedValue.ToString)

            '一覧表示
            Me.ShowList(_chosaNen, _upLow, _dataType, _kyoku, _jimusho, _kyoten)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnAllSelect_Click(sender As Object, e As EventArgs) Handles btnAllSelect.Click
        Try
            ComUtil.SetDataGridViewAllCheckEnabledOnly(dgvList, True)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnAllCancel_Click(sender As Object, e As EventArgs) Handles btnAllCancel.Click
        Try
            ComUtil.SetDataGridViewAllCheckEnabledOnly(dgvList, False)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

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

    Private Sub btnReceive_Click(sender As Object, e As EventArgs) Handles btnReceive.Click
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'センサス番号リスト取得
            Dim lstDc As List(Of Dictionary(Of String, String)) = Me.GetCensusNoList(_chosaNen, _upLow, _dataType)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(lstDc, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_013, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                Try
                    ProgressDialog = New ProgressDialog

                    '進捗ダイアログを表示する
                    ProgressDialog.Maximum = lstDc.Count * 7
                    ProgressDialog.Show(Me)

                    Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                        Try

                            db.BeginTrans()

                            For Each dc As Dictionary(Of String, String) In lstDc
                                Dim pChosa As DAOChosahyo.PrimaryKey = New DAOChosahyo.PrimaryKey(dc("調査年"), dc("センサス番号"))
                                Dim kChosa As DAOChosahyo.KotenKey = New DAOChosahyo.KotenKey(dc("農政局"), dc("都道府県"), dc("実査設置拠点"))

                                Dim pKobet As DAOKobetsuKekkahyo.PrimaryKey = New DAOKobetsuKekkahyo.PrimaryKey(dc("調査年"), dc("センサス番号"))
                                Dim kKobet As DAOKobetsuKekkahyo.KyotenKey = New DAOKobetsuKekkahyo.KyotenKey(dc("農政局"), dc("都道府県"), dc("実査設置拠点"))

                                '調査票データ削除
                                DAOChosahyo.DeleteChosahyoTable(db, pChosa, kChosa)

                                '進捗を進める
                                ProgressDialog.AddValue = 1

                                '個別結果表データ削除
                                DAOKobetsuKekkahyo.DeleteTable(db, pKobet, kKobet)

                                '進捗を進める
                                ProgressDialog.AddValue = 1

                                '調査票データ追加
                                DAOChosahyo.InsertChosahyoTable(db, dc)

                                '進捗を進める
                                ProgressDialog.AddValue = 1

                                '個別結果表データ追加
                                DAOKobetsuKekkahyo.InsertTable(db, dc)

                                '進捗を進める
                                ProgressDialog.AddValue = 1

                                '調査票データ削除（受信）
                                '--- REV.001 DEL START
                                'DAOChosahyo.DeleteChosahyoTableJushin(db, dc)
                                '--- REV.001 DEL END

                                '進捗を進める
                                ProgressDialog.AddValue = 1

                                '個別結果表データ削除（受信）
                                ''--- REV.001 DEL START
                                'DAOKobetsuKekkahyo.DeleteTableJushin(db, dc)
                                ''--- REV.001 DEL END

                                '進捗を進める
                                ProgressDialog.AddValue = 1

                                '送受信管理受信日時更新
                                DAOOther.UpdateSendReceiveManagementReceiveDate(db, dc)

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

                '完了メッセージ
                Message.ShowMsgBox(MessageID.MSG_I_013, MsgBoxStyle.OkOnly)

                '一覧表示
                Me.ShowList(_chosaNen, _upLow, _dataType, _kyoku, _jimusho, _kyoten)
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' センサス番号リスト取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="upLow"></param>
    ''' <param name="dataType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCensusNoList(chosaNen As String, upLow As String, dataType As String) As List(Of Dictionary(Of String, String))
        Dim ret As New List(Of Dictionary(Of String, String))

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                Dim pkey As New Dictionary(Of String, String)

                pkey("調査年") = _chosaNen
                pkey("センサス番号") = dgvList.Rows(i).Cells(7).Value.ToString
                pkey("上り下り区分") = upLow
                pkey("送受信データ種別") = dataType
                pkey("農政局") = dgvList.Rows(i).Cells(10).Value.ToString
                pkey("都道府県") = dgvList.Rows(i).Cells(11).Value.ToString
                pkey("実査設置拠点") = dgvList.Rows(i).Cells(12).Value.ToString

                ret.Add(pkey)
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="upLow"></param>
    ''' <param name="dataType"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <remarks></remarks>
    Private Sub ShowList(chosaNen As String, upLow As String, dataType As String, kyoku As String, jimusho As String, kyoten As String)
        Dim dt As DataTable

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dt = DAOOther.GetSendReceiveManagement(db, chosaNen, upLow, dataType, kyoku, jimusho, kyoten)
        End Using

        txtCount.Text = dt.Rows.Count.ToString

        dgvList.Rows.Clear()

        For Each row As DataRow In dt.Rows
            Dim censusNo As String = row("センサス番号").ToString
            Dim sendDate As String = row("送信日時").ToString
            Dim receiveDate As String = row("受信日時").ToString

            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(1).Value = ComUtil.GetTodofuken(censusNo)
            dgvList.Rows(i).Cells(2).Value = ComUtil.GetShikuchoson(censusNo)
            dgvList.Rows(i).Cells(3).Value = ComUtil.GetKyuShikuchoson(censusNo)
            dgvList.Rows(i).Cells(4).Value = ComUtil.GetNogyoShuraku(censusNo)
            dgvList.Rows(i).Cells(5).Value = ComUtil.GetChosaku(censusNo)
            dgvList.Rows(i).Cells(6).Value = ComUtil.GetKyakutaiNo(censusNo)
            dgvList.Rows(i).Cells(7).Value = censusNo
            dgvList.Rows(i).Cells(8).Value = If(String.IsNullOrEmpty(sendDate), sendDate, DateTime.Parse(sendDate).ToString(ComConst.DATETIME_FORMAT))
            dgvList.Rows(i).Cells(9).Value = If(String.IsNullOrEmpty(receiveDate), receiveDate, DateTime.Parse(receiveDate).ToString(ComConst.DATETIME_FORMAT))
            dgvList.Rows(i).Cells(10).Value = row("農政局").ToString
            dgvList.Rows(i).Cells(11).Value = row("都道府県").ToString
            dgvList.Rows(i).Cells(12).Value = row("実査設置拠点").ToString

            If Not (String.IsNullOrEmpty(sendDate) Or String.IsNullOrEmpty(receiveDate)) Then
                If DateTime.Parse(sendDate) < DateTime.Parse(receiveDate) Then
                    dgvList.Rows(i).Cells(0).ReadOnly = True

                    For j As Integer = 0 To dgvList.ColumnCount - 1
                        dgvList.Rows(i).Cells(j).Style.ForeColor = SystemColors.GrayText
                        dgvList.Rows(i).Cells(j).Style.BackColor = SystemColors.InactiveBorder
                    Next
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="pKey"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(pKeys As List(Of Dictionary(Of String, String)), ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'センサス番号選択チェック
        If pKeys.Count = 0 Then
            msgId = MessageID.MSG_E_003
            Return ret
        End If

        ret = True

        Return ret
    End Function
End Class
