''' <summary>
''' 電子調査票・個別結果表差分比較画面
''' </summary>
''' <remarks></remarks>
Public Class BRA7510F
    ''' <summary>
    ''' 差分比較クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 差分比較項目
        Public 調査年 As String
        Public 調査区分 As String
        Public 農政局 As String
        Public 都道府県 As String
        Public 実査設置拠点 As String
        Public センサス番号 As String
        Public データ識別区分 As String
        Public 項目番号 As String
        Public 明細番号 As String
        Public 前回送受信データ As String
        Public 最新データ As String
    End Class

    ''' <summary>上り下り区分</summary>
    Private _upLow As String
    ''' <summary>送受信データ種別</summary>
    Private _dataType As String
    ''' <summary>自調査工程</summary>
    Private _chosaKouteiFrom As String
    ''' <summary>送信先工程 </summary>
    Private _chosaKouteiTo As String

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

    Private _soujushinColName As String

    ''' <summary>進捗ダイアログ</summary>
    Private ProgressDialog As ProgressDialog

    Private Const BRAH As String = "BRAH"
    Private Const BRAN As String = "BRAN"
    Private Const BRAS As String = "BRAS"

    Private Const 電子調査票 As String = "1"
    Private Const 個別結果表 As String = "2"
    Private Const 客体情報なし As String = "客体情報なし"

    Private Const 受信日時 As String = "受信日時"
    Private Const 送信日時 As String = "送信日時"
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(soujushinKubun As Integer)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        _dataType = "1"

        Select Case (soujushinKubun)
            Case ComConst.送受信区分.本省_局から_受信
                _upLow = ComConst.上り下り区分.上位工程への送信
                _chosaKouteiFrom = BRAH
                _chosaKouteiTo = BRAH
                _soujushinColName = 受信日時
            Case ComConst.送受信区分.本省_局へ_送信
                _upLow = ComConst.上り下り区分.下位工程への送信
                _chosaKouteiFrom = BRAH
                _chosaKouteiTo = BRAN
                _soujushinColName = 送信日時
            Case ComConst.送受信区分.局_本省から_受信
                _upLow = ComConst.上り下り区分.下位工程への送信
                _chosaKouteiFrom = BRAN
                _chosaKouteiTo = BRAN
                _soujushinColName = 受信日時
            Case ComConst.送受信区分.局_本省へ_送信
                _upLow = ComConst.上り下り区分.上位工程への送信
                _chosaKouteiFrom = BRAN
                _chosaKouteiTo = BRAH
                _soujushinColName = 送信日時
            Case ComConst.送受信区分.局_実査設置拠点から_受信
                _upLow = ComConst.上り下り区分.上位工程への送信
                _chosaKouteiFrom = BRAN
                _chosaKouteiTo = BRAN
                _soujushinColName = 受信日時
            Case ComConst.送受信区分.局_実査設置拠点へ_送信
                _upLow = ComConst.上り下り区分.下位工程への送信
                _chosaKouteiFrom = BRAN
                _chosaKouteiTo = BRAS
                _soujushinColName = 送信日時
            Case ComConst.送受信区分.実査設置拠点_局から_受信
                _upLow = ComConst.上り下り区分.下位工程への送信
                _chosaKouteiFrom = BRAS
                _chosaKouteiTo = BRAS
                _soujushinColName = 受信日時
            Case ComConst.送受信区分.実査設置拠点_局へ_送信
                _upLow = ComConst.上り下り区分.上位工程への送信
                _chosaKouteiFrom = BRAS
                _chosaKouteiTo = BRAN
                _soujushinColName = 送信日時
        End Select

    End Sub

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA7510F_Load(sender As Object, e As EventArgs) Handles Me.Load
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


            dgvList.Columns(9).HeaderText = _soujushinColName

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 差分比較ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnHikaku_Click(sender As Object, e As EventArgs) Handles btnHikaku.Click
        Dim dtManage As DataTable = New DataTable
        Dim dtJushinChosahyo As Dictionary(Of String, DataTable)
        Dim dtChosahyo As Dictionary(Of String, DataTable)

        Dim dtKobetsu As Dictionary(Of String, DataTable)
        Dim dtJushinKobetsu As Dictionary(Of String, DataTable)

        Dim sabunList As New List(Of 差分比較項目)

        _chosaNen = cboChosaNen.SelectedValue.ToString
        _kyoku = If(IsDBNull(cboKyoku.SelectedValue), Nothing, cboKyoku.SelectedValue.ToString)
        _jimusho = If(IsDBNull(cboKyoten.SelectedValue), Nothing, CStr(CType(cboKyoten.SelectedItem, DataRowView)("事務所番号")))
        _kyoten = If(IsDBNull(cboKyoten.SelectedValue), Nothing, cboKyoten.SelectedValue.ToString)
        _einouRuikei = If(cboEinouRuikei.SelectedValue Is Nothing, Nothing, cboEinouRuikei.SelectedValue.ToString)

        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            If errCheckSoujushin(_chosaKouteiFrom, _upLow, _dataType, _kyoku, _jimusho, _kyoten, _chosaKouteiTo, dtManage) Then
                Message.ShowMsgBox(MessageID.MSG_I_041, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Try
                ProgressDialog = New ProgressDialog

                '進捗ダイアログを表示する
                ProgressDialog.Maximum = 5
                ProgressDialog.Show(Me)

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    Try

                        db.BeginTrans()

                        '差分比較データテーブルを削除する
                        DAOOther.DeleteSabunTable(db)

                        '進捗を進める
                        ProgressDialog.AddValue = 1

                        '調査票データ（受信）取得
                        dtJushinChosahyo = DAOChosahyo.GetHikakuChosahyoTable(db, _chosaNen, _kyoku, _jimusho, _kyoten, _einouRuikei, _upLow, _chosaKouteiTo, _chosaKouteiFrom, True)
                        '個別結果票データ(受信)取得
                        dtJushinKobetsu = DAOKobetsuKekkahyo.GetKobetuTableHikaku(db, _chosaNen, _kyoku, _jimusho, _kyoten, _einouRuikei, _upLow, _chosaKouteiTo, _chosaKouteiFrom, True)

                        '進捗を進める
                        ProgressDialog.AddValue = 1

                        '調査票データ取得
                        dtChosahyo = DAOChosahyo.GetHikakuChosahyoTable(db, _chosaNen, _kyoku, _jimusho, _kyoten, _einouRuikei, _upLow, _chosaKouteiTo, _chosaKouteiFrom)
                        '個別結果票データ(受信)取得
                        dtKobetsu = DAOKobetsuKekkahyo.GetKobetuTableHikaku(db, _chosaNen, _kyoku, _jimusho, _kyoten, _einouRuikei, _upLow, _chosaKouteiTo, _chosaKouteiFrom)

                        '進捗を進める
                        ProgressDialog.AddValue = 1

                        '差分の比較
                        Call ChosahyoSabunHikaku(dtChosahyo, dtJushinChosahyo, sabunList)
                        Call kobetsuSabunHikaku(dtKobetsu, dtJushinKobetsu, sabunList)

                        '進捗を進める
                        ProgressDialog.AddValue = 1

                        '差分をDBに登録
                        If sabunList.Count <> 0 Then
                            DAOOther.InsertSabunTable(db, sabunList)
                        End If

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

            '一覧表示
            Me.ShowList(_chosaKouteiTo, _chosaNen, _upLow, _dataType, _kyoku, _jimusho, _kyoten)

            ''完了メッセージ
            'Message.ShowMsgBox(MessageID.MSG_I_012, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 農政局コンボボックス変更
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
    ''' 出力ボタン押下
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Try
            'エラーチェック
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                If Not DAOOther.sabunKensuCheck(db, _chosaNen, _kyoku, _jimusho, _kyoten) Then
                    'エラーメッセージ
                    Message.ShowMsgBox(MessageID.MSG_E_023, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
            End Using

            'フォルダパス取得
            Dim folderPath As String = ComUtil.GetFolderPath(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath))

            If folderPath.Equals(String.Empty) Then
                Exit Sub
            End If

            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn

                Using ExcelOutput = New BRA7510R(folderPath, _chosaNen, _kyoku, _jimusho, _kyoten, _soujushinColName)
                    ret = ExcelOutput.Execute(MessageID.MSG_Q_004)
                End Using

                If ret = ExcelOutputBaseClass.enmOutputReturn.OK Then
                    '完了メッセージ
                    Message.ShowMsgBox(MessageID.MSG_I_002, MsgBoxStyle.OkOnly)
                End If
            Catch ex As ExcelOutputBaseClass.SaveAsException
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_006, MsgBoxStyle.OkOnly)
            Catch ex As CreateKobetsuException
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Message.ShowMsgBox(MessageID.MSG_E_054, {ex.CensusNo, ex.ItemNo}, MsgBoxStyle.OkOnly)
            Catch ex As Exception
                Throw
            End Try

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)

        End Try
    End Sub

    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <param name="chosaKoutei"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <remarks></remarks>
    Private Sub ShowList(chosaKoutei As String, chosaNen As String, upLow As String, dataType As String, kyoku As String, jimusho As String, kyoten As String)
        Dim dt As DataTable

        '差分一覧を取得
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dt = DAOOther.GetSabunIchiran(db, chosaNen, kyoku, jimusho, kyoten, chosaKoutei, upLow, CommonInfo.Chosakubun)
        End Using

        '一覧表示
        dgvList.Rows.Clear()

        For Each dtRow As DataRow In dt.Rows
            Dim censusNo As String = dtRow.Item("センサス番号").ToString
            Dim chosahyoDate As String = dtRow.Item("調査票更新日時").ToString
            Dim kobetsuDate As String = dtRow.Item("個別結果表更新日時").ToString
            Dim sendDate As String = String.Empty

            If _soujushinColName = 送信日時 Then
                sendDate = dtRow.Item("送信日時").ToString
            Else
                sendDate = dtRow.Item("受信日時").ToString
            End If

            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(0).Value = ComUtil.GetTodofuken(censusNo)
            dgvList.Rows(i).Cells(1).Value = ComUtil.GetShikuchoson(censusNo)
            dgvList.Rows(i).Cells(2).Value = ComUtil.GetKyuShikuchoson(censusNo)
            dgvList.Rows(i).Cells(3).Value = ComUtil.GetNogyoShuraku(censusNo)
            dgvList.Rows(i).Cells(4).Value = ComUtil.GetChosaku(censusNo)
            dgvList.Rows(i).Cells(5).Value = ComUtil.GetKyakutaiNo(censusNo)
            dgvList.Rows(i).Cells(6).Value = censusNo
            dgvList.Rows(i).Cells(7).Value = If(String.IsNullOrEmpty(chosahyoDate), chosahyoDate, DateTime.Parse(chosahyoDate).ToString(ComConst.DATETIME_FORMAT))
            dgvList.Rows(i).Cells(8).Value = If(String.IsNullOrEmpty(kobetsuDate), kobetsuDate, DateTime.Parse(kobetsuDate).ToString(ComConst.DATETIME_FORMAT))
            dgvList.Rows(i).Cells(9).Value = If(String.IsNullOrEmpty(sendDate), sendDate, DateTime.Parse(sendDate).ToString(ComConst.DATETIME_FORMAT))
        Next

    End Sub

    ''' <summary>
    ''' 送受信管理テーブルのデータを確認する
    ''' </summary>
    ''' <param name="pChosaNen"></param>
    ''' <param name="pUpLow"></param>
    ''' <param name="pDataType"></param>
    ''' <param name="pKyoku"></param>
    ''' <param name="pJjimusho"></param>
    ''' <param name="pKyoten"></param>
    ''' <param name="pChosaKouteiTo"></param>
    ''' <param name="pSoujushin"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function errCheckSoujushin(pChosaNen As String, pUpLow As String, pDataType As String, pKyoku As String, pJjimusho As String, pKyoten As String, pChosaKouteiTo As String, ByRef pSoujushin As DataTable) As Boolean
        Dim ret As Boolean = False

        '送受信管理テーブルを取得する
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            pSoujushin = DAOOther.GetSendReceiveManagement(db, _chosaNen, _upLow, _dataType, _kyoku, _jimusho, _kyoten, _chosaKouteiTo)
        End Using

        If pSoujushin.Rows.Count = 0 Then
            ret = True
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 調査票差分比較
    ''' </summary>
    ''' <param name="pManage"></param>
    ''' <param name="pChosahyo"></param>
    ''' <param name="pJushinChosahyo"></param>
    ''' <param name="pSabunList"></param>
    ''' <remarks></remarks>
    Private Sub ChosahyoSabunHikaku(pChosahyo As Dictionary(Of String, DataTable), pJushinChosahyo As Dictionary(Of String, DataTable), ByRef pSabunList As List(Of 差分比較項目))
        Dim jushinData As String
        Dim myData As String
        Dim kyoku As String
        Dim jimusho As String
        Dim kyoten As String
        Dim kouban As String
        Dim meisai As String
        Dim census As String

        For Each tableName As String In ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)
            Dim chosahyo As DataTable = pChosahyo.Item(tableName)
            Dim jushinChosahyo As DataTable = pJushinChosahyo.Item(tableName)
            Dim firstTable As Boolean = False
            If tableName = ComConst.調査票.テーブル名称(CommonInfo.Chosakubun)(0) Then
                If jushinChosahyo.Rows.Count = 0 Then
                    '自工程の調査票のみ存在する場合、テーブルを登録してループを終了
                    For Each d1 As DataRow In chosahyo.Rows
                        census = d1.Item("センサス番号").ToString
                        kyoku = d1.Item("農政局").ToString
                        jimusho = d1.Item("都道府県").ToString
                        kyoten = d1.Item("実査設置拠点").ToString
                        kouban = "-"
                        meisai = String.Empty
                        jushinData = 客体情報なし
                        myData = String.Empty

                        setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 電子調査票, kouban, meisai, jushinData, myData, pSabunList)
                    Next
                    Exit For
                ElseIf chosahyo.Rows.Count = 0 Then
                    '受信の調査票のみ存在する場合
                    For Each d1 As DataRow In jushinChosahyo.Rows
                        census = d1.Item("センサス番号").ToString
                        kyoku = d1.Item("農政局").ToString
                        jimusho = d1.Item("都道府県").ToString
                        kyoten = d1.Item("実査設置拠点").ToString
                        kouban = "-"
                        meisai = String.Empty
                        jushinData = String.Empty
                        myData = 客体情報なし

                        setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 電子調査票, kouban, meisai, jushinData, myData, pSabunList)
                    Next
                    Exit For
                End If
                firstTable = True
            End If

            '2つのテーブルを比較する
            If Not tableName.Contains("＿可変") Then
                tableCompare(chosahyo, jushinChosahyo, 電子調査票, firstTable, pSabunList)
            Else
                kahenTableCompare(chosahyo, jushinChosahyo, pSabunList)
            End If

        Next
    End Sub

    ''' <summary>
    ''' 個別結果票差分比較
    ''' </summary>
    ''' <param name="pManage"></param>
    ''' <param name="pkobetsu"></param>
    ''' <param name="pJushinkobetsu"></param>
    ''' <param name="pSabunList"></param>
    ''' <remarks></remarks>
    Private Sub kobetsuSabunHikaku(pkobetsu As Dictionary(Of String, DataTable), pJushinkobetsu As Dictionary(Of String, DataTable), ByRef pSabunList As List(Of 差分比較項目))
        Dim jushinData As String
        Dim myData As String
        Dim kyoku As String
        Dim jimusho As String
        Dim kyoten As String
        Dim kouban As String
        Dim meisai As String
        Dim census As String

        For Each tableName As String In ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)
            Dim kobetsu As DataTable = pkobetsu.Item(tableName)
            Dim jushinkobetsu As DataTable = pJushinkobetsu.Item(tableName)
            Dim firstTable As Boolean = False

            If tableName = ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0) Then
                If jushinkobetsu.Rows.Count = 0 Then
                    '自工程の調査票のみ存在する場合、テーブルを登録してループを終了
                    For Each d1 As DataRow In kobetsu.Rows
                        census = d1.Item("センサス番号").ToString
                        kyoku = d1.Item("農政局").ToString
                        jimusho = d1.Item("都道府県").ToString
                        kyoten = d1.Item("実査設置拠点").ToString
                        kouban = "-"
                        meisai = String.Empty
                        jushinData = 客体情報なし
                        myData = String.Empty

                        setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 個別結果表, kouban, meisai, jushinData, myData, pSabunList)
                    Next
                    Exit For
                ElseIf kobetsu.Rows.Count = 0 Then
                    '受信の調査票のみ存在する場合
                    For Each d1 As DataRow In jushinkobetsu.Rows
                        census = d1.Item("センサス番号").ToString
                        kyoku = d1.Item("農政局").ToString
                        jimusho = d1.Item("都道府県").ToString
                        kyoten = d1.Item("実査設置拠点").ToString
                        kouban = "-"
                        meisai = String.Empty
                        jushinData = String.Empty
                        myData = 客体情報なし

                        setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 個別結果表, kouban, meisai, jushinData, myData, pSabunList)
                    Next
                    Exit For
                End If

                firstTable = True
            End If

            '2つのテーブルを比較する
            tableCompare(kobetsu, jushinkobetsu, 個別結果表, firstTable, pSabunList)

        Next
    End Sub

    ''' <summary>
    ''' 2つのテーブルを比較して差分を登録する
    ''' </summary>
    ''' <param name="Table1"></param>
    ''' <param name="Table2"></param>
    ''' <param name="datatype"></param>
    ''' <param name="psabunList"></param>
    ''' <remarks></remarks>
    Private Sub tableCompare(Table1 As DataTable, Table2 As DataTable, datatype As String, firstTable As Boolean, psabunList As List(Of 差分比較項目))
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim comCol As String() = {"調査年", "センサス番号", "農政局", "都道府県", "実査設置拠点", "更新日付", "更新者ID"}
        Dim jushinData As String
        Dim myData As String
        Dim kyoku As String
        Dim jimusho As String
        Dim kyoten As String
        Dim kouban As String
        Dim meisai As String
        Dim census As String

        Do Until (i = Table1.Rows.Count And j = Table2.Rows.Count)
            Dim row1 As DataRow = Nothing
            Dim row2 As DataRow = Nothing

            If i <> Table1.Rows.Count Then
                row1 = Table1.Rows(i)
            End If

            If j <> Table2.Rows.Count Then
                row2 = Table2.Rows(j)
            End If

            If Not IsNothing(row1) And Not IsNothing(row2) Then
                Select Case row1.Item("センサス番号").ToString
                    Case Is = row2.Item("センサス番号").ToString
                        'センサス番号が同じ場合、カラムの確認を行う
                        For Each col As DataColumn In Table1.Columns
                            'カラムリストの中のカラム名は比較を行わない
                            If Not comCol.Contains(col.ColumnName) Then
                                If Not Table1.Rows(i).Item(col.ColumnName).ToString = Table2.Rows(j).Item(col.ColumnName).ToString Then
                                    'カラムが違う場合、登録する
                                    myData = row1.Item(col.ColumnName).ToString
                                    jushinData = row2.Item(col.ColumnName).ToString
                                    kyoku = row1.Item("農政局").ToString
                                    jimusho = row1.Item("都道府県").ToString
                                    kyoten = row1.Item("実査設置拠点").ToString
                                    kouban = col.ColumnName
                                    meisai = "0"
                                    census = row1.Item("センサス番号").ToString

                                    setSabun(_chosaNen, kyoku, jimusho, kyoten, census, datatype, kouban, meisai, jushinData, myData, psabunList)
                                End If
                            End If
                        Next

                        i += 1
                        j += 1
                    Case Is > row2.Item("センサス番号").ToString
                        If firstTable = True Then
                            myData = 客体情報なし
                            jushinData = String.Empty
                            kyoku = row2.Item("農政局").ToString
                            jimusho = row2.Item("都道府県").ToString
                            kyoten = row2.Item("実査設置拠点").ToString
                            kouban = "-"
                            meisai = String.Empty
                            census = row2.Item("センサス番号").ToString

                            setSabun(_chosaNen, kyoku, jimusho, kyoten, census, datatype, kouban, meisai, jushinData, myData, psabunList)
                        End If
                        j += 1
                    Case Is < row2.Item("センサス番号").ToString
                        If firstTable = True Then
                            myData = String.Empty
                            jushinData = 客体情報なし
                            kyoku = row1.Item("農政局").ToString
                            jimusho = row1.Item("都道府県").ToString
                            kyoten = row1.Item("実査設置拠点").ToString
                            kouban = "-"
                            meisai = String.Empty
                            census = row1.Item("センサス番号").ToString

                            setSabun(_chosaNen, kyoku, jimusho, kyoten, census, datatype, kouban, meisai, jushinData, myData, psabunList)
                        End If
                        i += 1
                End Select
            ElseIf IsNothing(row1) Then
                If firstTable = True Then
                    myData = 客体情報なし
                    jushinData = String.Empty
                    kyoku = row2.Item("農政局").ToString
                    jimusho = row2.Item("都道府県").ToString
                    kyoten = row2.Item("実査設置拠点").ToString
                    kouban = "-"
                    meisai = String.Empty
                    census = row2.Item("センサス番号").ToString

                    setSabun(_chosaNen, kyoku, jimusho, kyoten, census, datatype, kouban, meisai, jushinData, myData, psabunList)
                End If
                j += 1
            ElseIf IsNothing(row2) Then
                If firstTable = True Then
                    myData = String.Empty
                    jushinData = 客体情報なし
                    kyoku = row1.Item("農政局").ToString
                    jimusho = row1.Item("都道府県").ToString
                    kyoten = row1.Item("実査設置拠点").ToString
                    kouban = "-"
                    meisai = String.Empty
                    census = row1.Item("センサス番号").ToString

                    setSabun(_chosaNen, kyoku, jimusho, kyoten, census, datatype, kouban, meisai, jushinData, myData, psabunList)
                End If
                i += 1
            End If
        Loop

    End Sub

    ''' <summary>
    ''' 可変テーブルの差分チェック
    ''' </summary>
    ''' <param name="table1"></param>
    ''' <param name="table2"></param>
    ''' <param name="pSabunList"></param>
    ''' <remarks></remarks>
    Private Sub kahenTableCompare(table1 As DataTable, table2 As DataTable, ByRef pSabunList As List(Of 差分比較項目))
        Dim i As Integer = 0
        Dim j As Integer = 0

        Do Until (i = table1.Rows.Count And j = table2.Rows.Count)
            Dim row1 As DataRow = Nothing
            Dim row2 As DataRow = Nothing
            Dim jushinData As String
            Dim myData As String
            Dim kyoku As String
            Dim jimusho As String
            Dim kyoten As String
            Dim kouban As String
            Dim meisai As String
            Dim census As String

            If i <> table1.Rows.Count Then
                row1 = table1.Rows(i)
            End If

            If j <> table2.Rows.Count Then
                row2 = table2.Rows(j)
            End If

            'センサス番号のチェック
            If Not IsNothing(row1) And Not IsNothing(row2) Then
                Select Case row1.Item("センサス番号").ToString
                    Case Is > row2.Item("センサス番号").ToString
                        myData = String.Empty
                        jushinData = row2.Item("値").ToString
                        kyoku = row2.Item("農政局").ToString
                        jimusho = row2.Item("都道府県").ToString
                        kyoten = row2.Item("実査設置拠点").ToString
                        kouban = row2.Item("項目番号").ToString
                        meisai = row2.Item("明細番号").ToString
                        census = row2.Item("センサス番号").ToString

                        setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 電子調査票, kouban, meisai, jushinData, myData, pSabunList)
                        j += 1

                        Continue Do
                    Case Is < row2.Item("センサス番号").ToString
                        myData = row1.Item("値").ToString
                        jushinData = String.Empty
                        kyoku = row1.Item("農政局").ToString
                        jimusho = row1.Item("都道府県").ToString
                        kyoten = row1.Item("実査設置拠点").ToString
                        kouban = row1.Item("項目番号").ToString
                        meisai = row1.Item("明細番号").ToString
                        census = row1.Item("センサス番号").ToString

                        setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 電子調査票, kouban, meisai, jushinData, myData, pSabunList)
                        i += 1
                        Continue Do
                    Case Else
                End Select

                '項目番号のチェック
                Select Case row1.Item("項目番号").ToString
                    Case Is > row2.Item("項目番号").ToString
                        myData = String.Empty
                        jushinData = row2.Item("値").ToString
                        kyoku = row2.Item("農政局").ToString
                        jimusho = row2.Item("都道府県").ToString
                        kyoten = row2.Item("実査設置拠点").ToString
                        kouban = row2.Item("項目番号").ToString
                        meisai = row2.Item("明細番号").ToString
                        census = row2.Item("センサス番号").ToString

                        setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 電子調査票, kouban, meisai, jushinData, myData, pSabunList)
                        j += 1

                        Continue Do
                    Case Is < row2.Item("項目番号").ToString
                        myData = row1.Item("値").ToString
                        jushinData = String.Empty
                        kyoku = row1.Item("農政局").ToString
                        jimusho = row1.Item("都道府県").ToString
                        kyoten = row1.Item("実査設置拠点").ToString
                        kouban = row1.Item("項目番号").ToString
                        meisai = row1.Item("明細番号").ToString
                        census = row1.Item("センサス番号").ToString

                        setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 電子調査票, kouban, meisai, jushinData, myData, pSabunList)
                        i += 1

                        Continue Do
                    Case Else
                End Select

                '明細番号のチェック
                Select Case row1.Item("明細番号").ToString
                    Case Is > row2.Item("明細番号").ToString
                        myData = String.Empty
                        jushinData = row2.Item("値").ToString
                        kyoku = row2.Item("農政局").ToString
                        jimusho = row2.Item("都道府県").ToString
                        kyoten = row2.Item("実査設置拠点").ToString
                        kouban = row2.Item("項目番号").ToString
                        meisai = row2.Item("明細番号").ToString
                        census = row2.Item("センサス番号").ToString

                        setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 電子調査票, kouban, meisai, jushinData, myData, pSabunList)
                        j += 1

                        Continue Do
                    Case Is < row2.Item("明細番号").ToString
                        myData = row1.Item("値").ToString
                        jushinData = String.Empty
                        kyoku = row1.Item("農政局").ToString
                        jimusho = row1.Item("都道府県").ToString
                        kyoten = row1.Item("実査設置拠点").ToString
                        kouban = row1.Item("項目番号").ToString
                        meisai = row1.Item("明細番号").ToString
                        census = row1.Item("センサス番号").ToString

                        setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 電子調査票, kouban, meisai, jushinData, myData, pSabunList)
                        i += 1

                        Continue Do
                    Case Else
                End Select

                '値のチェック
                Select Case row1.Item("値").ToString
                    Case Is <> row2.Item("値").ToString
                        myData = row1.Item("値").ToString
                        jushinData = row2.Item("値").ToString
                        kyoku = row1.Item("農政局").ToString
                        jimusho = row1.Item("都道府県").ToString
                        kyoten = row1.Item("実査設置拠点").ToString
                        kouban = row1.Item("項目番号").ToString
                        meisai = row1.Item("明細番号").ToString
                        census = row1.Item("センサス番号").ToString

                        setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 電子調査票, kouban, meisai, jushinData, myData, pSabunList)
                    Case Else
                End Select

                i += 1
                j += 1
            ElseIf IsNothing(row1) Then
                myData = String.Empty
                jushinData = row2.Item("値").ToString
                kyoku = row2.Item("農政局").ToString
                jimusho = row2.Item("都道府県").ToString
                kyoten = row2.Item("実査設置拠点").ToString
                kouban = row2.Item("項目番号").ToString
                meisai = row2.Item("明細番号").ToString
                census = row2.Item("センサス番号").ToString

                setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 電子調査票, kouban, meisai, jushinData, myData, pSabunList)

                j += 1
            ElseIf IsNothing(row2) Then
                myData = row1.Item("値").ToString
                jushinData = String.Empty
                kyoku = row1.Item("農政局").ToString
                jimusho = row1.Item("都道府県").ToString
                kyoten = row1.Item("実査設置拠点").ToString
                kouban = row1.Item("項目番号").ToString
                meisai = row1.Item("明細番号").ToString
                census = row1.Item("センサス番号").ToString

                setSabun(_chosaNen, kyoku, jimusho, kyoten, census, 電子調査票, kouban, meisai, jushinData, myData, pSabunList)

                i += 1
            End If
        Loop
    End Sub

    Private Sub setSabun(chosanen As String, kyoku As String, jimusho As String, kyoten As String, census As String, _
                         DetaType As String, koumokuNo As String, meisaiNo As String, jushinDeta As String, saishinDeta As String, ByRef pSabunList As List(Of 差分比較項目))

        Dim sabun As New 差分比較項目

        sabun.調査年 = chosanen
        sabun.調査区分 = CommonInfo.Chosakubun
        sabun.農政局 = kyoku
        sabun.都道府県 = jimusho
        sabun.実査設置拠点 = kyoten
        sabun.センサス番号 = census
        sabun.データ識別区分 = DetaType
        sabun.項目番号 = koumokuNo
        sabun.明細番号 = meisaiNo
        sabun.前回送受信データ = jushinDeta
        sabun.最新データ = saishinDeta

        pSabunList.Add(sabun)

    End Sub

    Private Sub btnReturn_Click(sender As Object, e As EventArgs) Handles btnReturn.Click
        '差分比較データテーブルを削除する
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            DAOOther.DeleteSabunTable(db)
        End Using
    End Sub
End Class
