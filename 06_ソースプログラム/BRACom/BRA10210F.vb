'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_000   | 2023.01.13 |大興電子通信        | 要件No.5 新規作成
'//            |            |                    |
'//*************************************************************************************************
Imports System.Data.SqlClient
Imports System.IO
Imports System.Reflection
Imports System.Text

''' <summary>
''' 標本リスト管理画面
''' </summary>
''' <remarks></remarks>
Public Class BRA10210F

    ''' <summary>
    ''' センサス実施年
    ''' </summary>
    Private CensusNen As Integer

    ''' <summary>
    ''' 標本リストyyyyの列情報
    ''' </summary>
    Private ColInfo As (COLUMN_ID As Integer, NAME As String, IS_NULLABLE As Boolean, TYPE As String, MAX_LENGTH As Integer, PRECISION As Integer, SCALE As Integer)()

    ''' <summary>
    ''' 表頭名配列
    ''' </summary>
    Private HyohonHeaders As String()

    ''' <summary>
    ''' 標本リスト列数
    ''' </summary>
    Private HyohonColCount As Integer

    ''' <summary>
    ''' 農政局（追加時のキー項目）
    ''' </summary>
    Private Kyoku As Integer
    ''' <summary>
    ''' 都道府県（追加時のキー項目）
    ''' </summary>
    Private Pref As Integer
    ''' <summary>
    ''' 実査設置拠点（追加時のキー項目用）
    ''' </summary>
    Private Kyoten As Integer
    ''' <summary>
    ''' 経営形態区分（追加時のキー項目）
    ''' </summary>
    Private KeieiKeitai As Integer
    ''' <summary>
    ''' 営農類型区分（追加時のキー項目）
    ''' </summary>
    Private EinoRuikei As Integer
    ''' <summary>
    ''' 生産費区分（追加時のキー項目）
    ''' </summary>
    Private Seisanhi As Integer
    ''' <summary>
    ''' 田畑区分（追加時のキー項目）
    ''' </summary>
    Private Tahata As Integer

    ''' <summary>
    ''' 表示時データ
    ''' </summary>
    Private BeforeData As List(Of List(Of Object))

    ''' <summary>
    ''' 削除情報
    ''' </summary>
    Private ReadOnly DelInfo As New HashSet(Of (ichirenNo As String, keieiKeitai As Integer, einoRuikei As Integer, seisanhi As Integer, tahata As Integer))

    ''' <summary>
    ''' 一連番号列Index
    ''' </summary>
    Private 一連番号列 As Integer
    ''' <summary>
    ''' 経営形態区分列Index
    ''' </summary>
    Private 経営形態区分列 As Integer
    ''' <summary>
    ''' 営農類型区分列Index
    ''' </summary>
    Private 営農類型区分列 As Integer
    ''' <summary>
    ''' 生産費区分列Index
    ''' </summary>
    Private 生産費区分列 As Integer
    ''' <summary>
    ''' 田畑区分列Index
    ''' </summary>
    Private 田畑区分列 As Integer
    ''' <summary>
    ''' 変更区分列
    ''' </summary>
    Private 変更区分列 As Integer

    ''' <summary>
    ''' 変更区分
    ''' </summary>
    Private Enum 変更区分 As Integer
        変更なし
        変更あり
        追加
    End Enum

    ''' <summary>
    ''' 初期表示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA10210F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'DGVちらつき防止
            dgvList.GetType().InvokeMember("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, Nothing, dgvList, New Object() {True})

            '本省工程は編集不可
            If CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Honsyo Then
                BtnShow.Enabled = False
            End If
            '表示するまで編集関係ボタンは使用不可
            SetEditButtonEnabled(False)

            'センサス実施年設定
            HyohonUtil.SetCombobox(CboCensusNen, HyohonConst.センサス実施年リスト)
            CboCensusNen.SelectedIndex = 0
            '農政局プルダウン設定
            ComUtil.SetKyokuComboBox(CboKyoku)
            '経営形態区分プルダウン設定
            HyohonUtil.SetCombobox(CboKeieiKeitai, HyohonUtil.Get経営形態区分リスト())
            '営農類型区分プルダウン設定
            HyohonUtil.SetCombobox(CboEinoRuikei, HyohonConst.営農類型区分リスト)
            '生産費区分プルダウン設定
            HyohonUtil.SetCombobox(CboSeisanhi, HyohonUtil.Get生産費区分リスト(HyohonConst.経営形態区分.識別対象外))
            '田畑区分プルダウン設定
            HyohonUtil.SetCombobox(CboTahata, HyohonConst.田畑区分リスト)
            '規模階層プルダウン設定
            HyohonUtil.SetCombobox(CboKiboKaiso, HyohonConst.規模階層区分リスト)
            '標本区分プルダウン設定
            HyohonUtil.SetCombobox(CboHyohon, HyohonConst.標本区分リスト)
            '標本候補区分プルダウン設定
            HyohonUtil.SetCombobox(CboHyohonKoho, HyohonConst.標本候補区分リスト)
            '経営形態区分の空を選択
            CboKeieiKeitai.SelectedValue = HyohonConst.経営形態区分.識別対象外

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' センサス実施年プルダウン選択
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CboCensusNen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboCensusNen.SelectedIndexChanged
        Try
            'センサス実施年
            Dim nen = CInt(CboCensusNen.SelectedValue)
            If CensusNen = nen Then
                Return
            End If
            CensusNen = nen
            '変更有無判定
            If IsModified() Then
                If Message.ShowMsgBox(MessageID.MSG_Q_058, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Return
                End If
            End If

            'テーブルの列情報取得
            ColInfo = HyohonUtil.GetColInfo(CensusNen)

            '一覧初期化
            dgvList.Rows.Clear()
            dgvList.Columns.Clear()
            '表示するまで編集関係ボタンは使用不可
            SetEditButtonEnabled(False)
            'センサス実施年の標本リスト表頭名から一覧の列を設定する
            HyohonHeaders = HyohonConst.標本リスト表頭名(CensusNen)
            HyohonColCount = HyohonHeaders.Length - 1
            Dim col As DataGridViewColumn
            '選択
            col = New DataGridViewCheckBoxColumn()
            col.Frozen = True
            col.Width = 30
            dgvList.Columns.Add(col)
            '標本リスト項目
            For i = 1 To HyohonColCount
                col = New DataGridViewTextBoxColumn()
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                col.HeaderText = i & "．" & HyohonHeaders(i)
                If i = HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.一連番号) Then
                    '一連番号
                    col.ReadOnly = True
                    col.Frozen = True
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                    col.Width = 120
                    一連番号列 = i
                End If
                Select Case i
                    Case HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.都道府県CD),
                         HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.市区町村CD),
                         HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.旧市区町村CD),
                         HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.農業集落CD),
                         HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.調査区CD),
                         HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.客体番号CD)
                        '都道府県CD～客体番号(振興局CD除く)
                        col.ReadOnly = True
                    Case 9 To 12
                        '経営体名称、経営主氏名、住所１、２
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                        col.Width = 150
                    Case 14
                        '電話番号
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                        col.Width = 90
                    Case 257
                        '備考
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                        col.Width = 400
                End Select
                Select Case ColInfo(i - 1).TYPE
                    Case "CHAR", "VARCHAR", "NVARCHAR"
                        col.CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    Case Else
                        col.CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                End Select
                dgvList.Columns.Add(col)
            Next
            '経営形態区分
            col = New DataGridViewTextBoxColumn()
            col.HeaderText = "経営形態区分"
            col.ReadOnly = True
            col.Width = 95
            col.CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvList.Columns.Add(col)
            '営農類型・生産費区分（田畑区分）
            col = New DataGridViewTextBoxColumn()
            col.HeaderText = "営農類型・生産費区分（田畑区分）"
            col.ReadOnly = True
            col.Width = 190
            col.CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvList.Columns.Add(col)
            '更新日時
            col = New DataGridViewTextBoxColumn()
            col.HeaderText = "更新日時"
            col.ReadOnly = True
            col.Width = 100
            col.CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvList.Columns.Add(col)
            '更新者ID
            col = New DataGridViewTextBoxColumn()
            col.HeaderText = "更新者ID"
            col.ReadOnly = True
            col.Width = 100
            col.CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvList.Columns.Add(col)
            '経営形態区分（非表示）
            col = New DataGridViewTextBoxColumn()
            col.Visible = False
            dgvList.Columns.Add(col)
            経営形態区分列 = col.Index
            '営農類型区分（非表示）
            col = New DataGridViewTextBoxColumn()
            col.Visible = False
            dgvList.Columns.Add(col)
            営農類型区分列 = col.Index
            '生産費区分（非表示）
            col = New DataGridViewTextBoxColumn()
            col.Visible = False
            dgvList.Columns.Add(col)
            生産費区分列 = col.Index
            '田畑区分（非表示）
            col = New DataGridViewTextBoxColumn()
            col.Visible = False
            dgvList.Columns.Add(col)
            田畑区分列 = col.Index
            '変更区分（非表示）
            col = New DataGridViewTextBoxColumn()
            col.Visible = False
            dgvList.Columns.Add(col)
            変更区分列 = col.Index

            'DGV共通設定
            ComUtil.ConfigDgvEditable(dgvList, 0)

            'セル選択で編集モードにする
            dgvList.EditMode = DataGridViewEditMode.EditOnEnter

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
    ''' 経営形態区分プルダウン選択
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CboKeieiKeitai_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboKeieiKeitai.SelectedIndexChanged
        Try
            Dim keieiKeitai = DirectCast(HyohonUtil.GetComboValue(CboKeieiKeitai), HyohonConst.経営形態区分)
            Select Case keieiKeitai
                Case HyohonConst.経営形態区分.識別対象外
                    '空
                    CboEinoRuikei.SelectedValue = HyohonConst.営農類型区分.識別対象外
                    CboEinoRuikei.Enabled = False
                    CboSeisanhi.SelectedValue = HyohonConst.生産費区分.識別対象外
                    CboSeisanhi.Enabled = False
                Case HyohonConst.経営形態区分.個人経営体, HyohonConst.経営形態区分.法人経営体
                    '個人経営体、法人経営体
                    CboEinoRuikei.Enabled = True
                    CboSeisanhi.SelectedValue = HyohonConst.生産費区分.識別対象外
                    CboSeisanhi.Enabled = False
                Case Else
                    '個別経営体、組織法人経営体
                    CboEinoRuikei.SelectedValue = HyohonConst.営農類型区分.識別対象外
                    CboEinoRuikei.Enabled = False
                    HyohonUtil.SetCombobox(CboSeisanhi, HyohonUtil.Get生産費区分リスト(DirectCast(HyohonUtil.GetComboValue(CboKeieiKeitai), HyohonConst.経営形態区分)))
                    CboSeisanhi.Enabled = True
                    CboSeisanhi.SelectedValue = HyohonConst.生産費区分.識別対象外
            End Select

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 生産費区分プルダウン選択
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CboSeisanhi_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboSeisanhi.SelectedIndexChanged
        Try
            Dim keieiKeitai = DirectCast(HyohonUtil.GetComboValue(CboKeieiKeitai), HyohonConst.経営形態区分)
            Dim seisanhi = DirectCast(HyohonUtil.GetComboValue(CboSeisanhi), HyohonConst.生産費区分)
            If keieiKeitai = HyohonConst.経営形態区分.個別経営体 _
                AndAlso (seisanhi = HyohonConst.生産費区分.小麦 OrElse seisanhi = HyohonConst.生産費区分.大豆) Then
                '小麦、大豆
                CboTahata.Enabled = True
            Else
                '小麦、大豆以外
                CboTahata.SelectedValue = HyohonConst.生産費区分.識別対象外
                CboTahata.Enabled = False
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' クリアボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        Try
            TxtTodofuken.Text = ""
            TxtShikuchoson.Text = ""
            CboKiboKaiso.SelectedValue = HyohonConst.識別対象外値
            CboHyohon.SelectedValue = HyohonConst.識別対象外値
            CboHyohonKoho.SelectedValue = HyohonConst.識別対象外値
            TxtIchirenNo.Text = ""
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' CSV出力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnCsvOutput_Click(sender As Object, e As EventArgs) Handles BtnCsvOutput.Click

        Try
            '入力チェック
            '都道府県CD
            If Not IsFitToSearchValue(TxtTodofuken.Text) Then
                Message.ShowMsgBox(MessageID.MSG_E_128, {"都道府県CD", "数字と記号(,-)のみ"}, MsgBoxStyle.OkOnly)
                TxtTodofuken.Focus()
                Return
            End If
            '市区町村CD
            If Not IsFitToSearchValue(TxtShikuchoson.Text) Then
                Message.ShowMsgBox(MessageID.MSG_E_128, {"市区町村CD", "数字と記号(,-)のみ"}, MsgBoxStyle.OkOnly)
                TxtShikuchoson.Focus()
                Return
            End If

            '実行確認
            Dim csvOutPattern = HyohonConst.CSV出力パターン.実査設置拠点
            If HyohonUtil.GetComboValue(CboKyoku) = HyohonConst.識別対象外値 Then
                If Message.ShowMsgBox(MessageID.MSG_Q_059, vbYesNo) = vbNo Then
                    Return
                End If
                csvOutPattern = HyohonConst.CSV出力パターン.全国
            ElseIf HyohonUtil.GetComboValue(CboKyoten) = HyohonConst.識別対象外値 Then
                If Message.ShowMsgBox(MessageID.MSG_Q_060, vbYesNo) = vbNo Then
                    Return
                End If
                csvOutPattern = HyohonConst.CSV出力パターン.農政局
            End If

            'フォルダパス取得
            Dim folderPath As String = ComUtil.GetFolderPath(Me, IniFileInfo.ExcelOutPath)
            If String.IsNullOrEmpty(folderPath) Then
                Return
            End If

            '標本リスト検索
            Cursor.Current = Cursors.WaitCursor

            '進捗ダイアログ
            Dim progressDialog As New ProgressDialog()
            Try
                '進捗ダイアログを表示
                progressDialog.Maximum = 200
                progressDialog.Show(Me)

                Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    '検索結果件数
                    Dim resultCount As Integer
                    Dim dr = Select標本リスト(db, progressDialog, resultCount)
                    If Not dr.HasRows Then
                        progressDialog.ShowMsgBox(MessageID.MSG_E_023, MsgBoxStyle.OkOnly)
                        Return
                    End If

                    '標本リストCSV出力
                    CsvOutput(csvOutPattern, dr, resultCount, folderPath, progressDialog)
                End Using
            Finally
                If progressDialog IsNot Nothing Then
                    '進捗ダイアログを閉じる
                    progressDialog.endDispose()
                    progressDialog = Nothing
                End If
            End Try

            '完了メッセージ
            Message.ShowMsgBox(MessageID.MSG_I_002, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 表示ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnShow_Click(sender As Object, e As EventArgs) Handles BtnShow.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            '変更有無判定
            If IsModified() Then
                If Message.ShowMsgBox(MessageID.MSG_Q_058, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Return
                End If
            End If

            '入力チェック
            '田畑区分（活性の場合、要否確認）
            Dim cbo = CboTahata
            If cbo.Enabled AndAlso HyohonUtil.IsEmpty(cbo) Then
                If Message.ShowMsgBox(MessageID.MSG_Q_057, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    cbo.Focus()
                    Return
                End If
            End If
            '都道府県CD
            If Not IsFitToSearchValue(TxtTodofuken.Text) Then
                Message.ShowMsgBox(MessageID.MSG_E_128, {"都道府県CD", "数字と記号(,-)のみ"}, MsgBoxStyle.OkOnly)
                TxtTodofuken.Focus()
                Return
            End If
            '市区町村CD
            If Not IsFitToSearchValue(TxtShikuchoson.Text) Then
                Message.ShowMsgBox(MessageID.MSG_E_128, {"市区町村CD", "数字と記号(,-)のみ"}, MsgBoxStyle.OkOnly)
                TxtShikuchoson.Focus()
                Return
            End If

            '進捗ダイアログ
            Dim progressDialog As New ProgressDialog()
            Try
                '進捗ダイアログを表示
                progressDialog.Maximum = 200
                progressDialog.Show(Me)

                '最大表示件数を取得
                Dim maxResult = My.Settings.HyohonMaxResult

                '標本リスト検索
                Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    Dim resultCount As Integer
                    Dim dr = Select標本リスト(db, progressDialog, resultCount, maxResult)
                    'キー項目を退避
                    Kyoku = HyohonUtil.GetComboValue(CboKyoku)
                    Kyoten = HyohonUtil.GetComboValue(CboKyoten)
                    If Kyoten <> HyohonConst.識別対象外値 Then
                        Pref = CInt(CType(CboKyoten.SelectedItem, DataRowView)("事務所番号"))
                        Pref = If(Pref = 51, 1, Pref)
                    Else
                        Pref = HyohonConst.識別対象外値
                    End If
                    KeieiKeitai = HyohonUtil.GetComboValue(CboKeieiKeitai)
                    EinoRuikei = HyohonUtil.GetComboValue(CboEinoRuikei)
                    Seisanhi = HyohonUtil.GetComboValue(CboSeisanhi)
                    Tahata = HyohonUtil.GetComboValue(CboTahata)
                    '削除情報クリア
                    DelInfo.Clear()
                    '一覧クリア
                    dgvList.Rows.Clear()
                    '都道府県CD～客体番号を読取専用にする
                    dgvList.Columns(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.都道府県CD)).ReadOnly = True
                    dgvList.Columns(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.市区町村CD)).ReadOnly = True
                    dgvList.Columns(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.旧市区町村CD)).ReadOnly = True
                    dgvList.Columns(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.農業集落CD)).ReadOnly = True
                    dgvList.Columns(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.調査区CD)).ReadOnly = True
                    dgvList.Columns(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.客体番号CD)).ReadOnly = True
                    '一覧表示
                    BeforeData = New List(Of List(Of Object))
                    Dim readCount = 0
                    Dim max = Math.Min(resultCount, maxResult)
                    While dr.Read() AndAlso readCount < max
                        readCount += 1
                        Dim values = New List(Of Object)
                        'キー項目取得
                        Dim keieiKeitai = CInt(dr("経営形態区分"))
                        Dim einoRuikei = CInt(dr("営農類型区分"))
                        Dim seisanhi = CInt(dr("生産費区分"))
                        Dim tahata = CInt(dr("田畑区分"))
                        '選択
                        values.Add(False)
                        '標本リスト項目
                        For i = 0 To HyohonColCount - 1
                            values.Add(dr(i))
                        Next
                        '経営形態区分
                        values.Add(HyohonUtil.Getリスト名称(HyohonUtil.Get経営形態区分リスト(), keieiKeitai))
                        '営農類型・生産費区分（田畑区分）
                        values.Add(HyohonUtil.Get営農類型＿生産費区分名称(keieiKeitai, einoRuikei, seisanhi, tahata))
                        '更新日時
                        values.Add(dr("更新日時"))
                        '更新者ID
                        values.Add(dr("更新者ID"))
                        '経営形態区分（非表示）
                        values.Add(keieiKeitai)
                        '営農類型区分（非表示）
                        values.Add(einoRuikei)
                        '生産費区分（非表示）
                        values.Add(seisanhi)
                        '田畑区分（非表示）
                        values.Add(tahata)
                        '変更区分（非表示）
                        values.Add(変更区分.変更なし)
                        '一覧に行を追加
                        HyohonUtil.AddDgvRow(dgvList, values.ToArray())
                        '表示時情報に追加
                        BeforeData.Add(values)
                        '進捗を進める
                        progressDialog.Value = CInt(Math.Truncate(readCount / max * 100)) + 100
                    End While

                    '進捗を進める
                    progressDialog.Value = 200

                    '編集関係ボタンを活性化
                    SetEditButtonEnabled(True)

                    dgvList.ClearSelection()

                    '最大表示件数を越えている場合、メッセージを表示
                    If maxResult < resultCount Then
                        progressDialog.ShowMsgBox(MessageID.MSG_I_050, {maxResult.ToString()}, MsgBoxStyle.OkOnly)
                    End If
                End Using

            Finally
                If progressDialog IsNot Nothing Then
                    '進捗ダイアログを閉じる
                    progressDialog.endDispose()
                    progressDialog = Nothing
                End If
            End Try

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' セル描画
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DgvList_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles dgvList.CellPainting
        Try
            If e.RowIndex < 0 Then
                Return
            End If
            Dim cell = dgvList(e.ColumnIndex, e.RowIndex)
            If cell.ReadOnly Then
                '読取専用セルを薄いグレーにする
                cell.Style.BackColor = Color.WhiteSmoke
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' セル未確定状態変更
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DgvList_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvList.CurrentCellDirtyStateChanged
        Try
            '選択列で未確定の場合、確定する
            If dgvList.CurrentCellAddress.X = 0 AndAlso dgvList.IsCurrentCellDirty Then
                dgvList.CommitEdit(DataGridViewDataErrorContexts.Commit)
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' セル値変更
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DgvList_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvList.CellValueChanged
        Try
            If e.ColumnIndex <> 0 Then
                '選択列以外は何もしない
                Return
            End If
            '削除ボタン制御
            Dim hasSelectedRow = dgvList.Rows.Cast(Of DataGridViewRow) _
                .Where(Function(x) CBool(x.Cells(0).Value)).FirstOrDefault IsNot Nothing
            BtnDel.Enabled = hasSelectedRow
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' セル編集開始
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DgvList_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvList.CellBeginEdit
        Try
            If e.ColumnIndex <= 0 Then
                '選択列は何もしない
                Return
            End If
            Dim cell = DirectCast(dgvList(e.ColumnIndex, e.RowIndex), DataGridViewTextBoxCell)
            '最大入力可能桁数を設定する
            cell.MaxInputLength = ColInfo(e.ColumnIndex - 1).MAX_LENGTH
            If ColInfo(e.ColumnIndex - 1).COLUMN_ID = HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.営農類型区分規模) Then
                '負数を許容するため、符号分+1
                cell.MaxInputLength += 1
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 一覧セル検証後
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DgvList_CellValidated(sender As Object, e As DataGridViewCellEventArgs) Handles dgvList.CellValidated
        Try
            If e.ColumnIndex <= 0 Then
                '選択列は何もしない
                Return
            End If

            Dim cell = DirectCast(dgvList(e.ColumnIndex, e.RowIndex), DataGridViewTextBoxCell)
            If cell.ReadOnly Then
                '読取専用は何もしない
                Return
            End If
            'Nothingは空文字にする
            If cell.Value Is Nothing Then
                cell.Value = ""
            End If
            '数字項目の場合、非数字は空にする
            If ColInfo(e.ColumnIndex - 1).TYPE = "NUMERIC" Then
                Dim d As Decimal
                If Decimal.TryParse(cell.Value.ToString, d) Then
                    Select Case e.ColumnIndex
                        Case HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.都道府県CD),
                             HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.振興局CD),
                             HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.市区町村CD),
                             HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.旧市区町村CD),
                             HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.農業集落CD),
                             HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.調査区CD),
                             HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.客体番号CD)
                            '都道府県CD～客体番号は0埋め
                            cell.Value = d.ToString().PadLeft(cell.MaxInputLength, "0"c)
                        Case Else
                            cell.Value = d.ToString()
                    End Select
                Else
                    cell.Value = ""
                End If
            End If
            '半角カンマを全角に置換する
            cell.Value = cell.Value.ToString().Replace(",", "、")
            '都道府県CD、市区町村CD、旧市区町村CD、農業集落CD、調査区CD、客体番号CDから一連番号を設定する
            Select Case e.ColumnIndex
                Case HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.都道府県CD),
                     HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.市区町村CD),
                     HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.旧市区町村CD),
                     HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.農業集落CD),
                     HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.調査区CD),
                     HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.客体番号CD)
                    dgvList(1, e.RowIndex).Value = "a" &
                        If(dgvList(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.都道府県CD), e.RowIndex).Value, 0).ToString.PadLeft(2, "0"c) &
                        If(dgvList(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.市区町村CD), e.RowIndex).Value, 0).ToString.PadLeft(3, "0"c) &
                        If(dgvList(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.旧市区町村CD), e.RowIndex).Value, 0).ToString.PadLeft(2, "0"c) &
                        If(dgvList(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.農業集落CD), e.RowIndex).Value, 0).ToString.PadLeft(3, "0"c) &
                        If(dgvList(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.調査区CD), e.RowIndex).Value, 0).ToString.PadLeft(3, "0"c) &
                        If(dgvList(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.客体番号CD), e.RowIndex).Value, 0).ToString.PadLeft(3, "0"c)
            End Select

            '変更有無判定
            Dim row = dgvList.Rows(e.RowIndex)
            If CInt(row.Cells(変更区分列).Value) = 変更区分.追加 Then
                Return
            End If
            Dim before = BeforeData.Where(Function(x) x(一連番号列).ToString() = row.Cells(一連番号列).Value.ToString()).FirstOrDefault()
            Dim hasDiff = before.Where(Function(x, index) index >= 一連番号列 AndAlso index <= HyohonColCount AndAlso x.ToString() <> row.Cells(index).Value.ToString()).FirstOrDefault() IsNot Nothing
            row.Cells(変更区分列).Value = If(hasDiff, 変更区分.変更あり, 変更区分.変更なし)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 全選択ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
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
    ''' <remarks></remarks>
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
    ''' 追加ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click
        Try
            '最大表示件数を取得
            Dim maxResult = My.Settings.HyohonMaxResult
            If dgvList.Rows.Count + 1 > maxResult Then
                '最大表示件数を越える
                Message.ShowMsgBox(MessageID.MSG_E_126, {maxResult.ToString()}, MsgBoxStyle.OkOnly)
                Return
            End If

            '選択、一連番号
            Dim values = New List(Of Object) From {False, "a0000000000000000"}
            '標本リスト項目（一連番号以外）（空）
            values.AddRange(Enumerable.Repeat("", HyohonColCount - 1))
            '経営形態区分
            values.Add(HyohonUtil.Getリスト名称(HyohonUtil.Get経営形態区分リスト(), KeieiKeitai))
            '営農類型・生産費区分（田畑区分）
            values.Add(HyohonUtil.Get営農類型＿生産費区分名称(KeieiKeitai, EinoRuikei, Seisanhi, Tahata))
            '更新日時（空）、更新者ID（空）
            values.AddRange({"", ""})
            '経営形態区分（非表示）
            values.Add(KeieiKeitai)
            '営農類型区分（非表示）
            values.Add(EinoRuikei)
            '生産費区分（非表示）
            values.Add(Seisanhi)
            '田畑区分（非表示）
            values.Add(Tahata)
            '変更区分（非表示）（追加）
            values.Add(変更区分.追加)
            Dim addRow = HyohonUtil.AddDgvRow(dgvList, values.ToArray())
            '都道府県CD～客体番号を入力可能にする
            addRow.Cells(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.都道府県CD)).ReadOnly = False
            addRow.Cells(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.市区町村CD)).ReadOnly = False
            addRow.Cells(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.旧市区町村CD)).ReadOnly = False
            addRow.Cells(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.農業集落CD)).ReadOnly = False
            addRow.Cells(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.調査区CD)).ReadOnly = False
            addRow.Cells(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.客体番号CD)).ReadOnly = False
            '都道府県CDをアクティブセルにする
            dgvList.CurrentCell = addRow.Cells(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.都道府県CD))
            dgvList.BeginEdit(True)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 削除ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnDel_Click(sender As Object, e As EventArgs) Handles BtnDel.Click
        Try
            '実行確認
            If Message.ShowMsgBox(MessageID.MSG_Q_061, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If
            '選択行取得
            Dim selectedRow = dgvList.Rows.Cast(Of DataGridViewRow).Where(Function(x) CBool(x.Cells(0).Value)) _
                .OrderByDescending(Function(x) x.Cells(0).RowIndex)
            For Each row In selectedRow
                If CInt(row.Cells(変更区分列).Value) <> 変更区分.追加 Then
                    '削除情報に追加
                    DelInfo.Add((row.Cells(一連番号列).Value.ToString(), CInt(row.Cells(経営形態区分列).Value), CInt(row.Cells(営農類型区分列).Value), CInt(row.Cells(生産費区分列).Value), CInt(row.Cells(田畑区分列).Value)))
                End If
                '一覧から削除
                dgvList.Rows.Remove(row)
            Next
            '削除ボタン非活性化
            BtnDel.Enabled = False
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 保存ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            '変更チェック
            If Not IsModified() Then
                Message.ShowMsgBox(MessageID.MSG_I_051, MsgBoxStyle.OkOnly)
                Return
            End If

            '入力チェック
            Dim cell = CheckForSave()
            If cell IsNot Nothing Then
                dgvList.CurrentCell = cell
                If Not cell.ReadOnly Then
                    dgvList.BeginEdit(True)
                End If
                Return
            End If

            '実行確認
            If Message.ShowMsgBox(MessageID.MSG_Q_001, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If

            'データベース反映処理
            Save標本リスト()

            '削除情報クリア
            DelInfo.Clear()
            '一覧クリア
            dgvList.Rows.Clear()
            '編集関係ボタンを非活性化
            SetEditButtonEnabled(False)

            '完了メッセージ
            Message.ShowMsgBox(MessageID.MSG_I_001, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 戻るボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Overrides Sub BtnReturn_Click(sender As Object, e As EventArgs)
        If IsModified() Then
            If Message.ShowMsgBox(MessageID.MSG_Q_058, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If
        End If
        MyBase.BtnReturn_Click(sender, e)
    End Sub

    ''' <summary>
    ''' フォームサイズ変更
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BRA10210F_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        '一覧サイズ
        dgvList.Width = Size.Width - 48
        dgvList.Height = Size.Height - 400
        'ボタン位置
        Dim top = Size.Height - 86
        BtnAllSelect.Top = top
        BtnAllRelease.Top = top
        BtnAdd.Top = top
        BtnDel.Top = top
        BtnSave.Top = top
        btnReturn.Top = top
    End Sub

    ''' <summary>
    ''' キー押下
    ''' 一連番号では、EnterでTabしないようにする。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub BaseForm_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs)
        If ActiveControl.Equals(TxtIchirenNo) Then
            Return
        End If
        MyBase.BaseForm_KeyDown(sender, e)
    End Sub

    ''' <summary>
    ''' 編集関係ボタンの活性／非活性設定
    ''' </summary>
    ''' <param name="enabled"></param>
    Private Sub SetEditButtonEnabled(enabled As Boolean)
        BtnAllSelect.Enabled = enabled
        BtnAllRelease.Enabled = enabled
        '削除の活性化はしない（行選択で制御）
        If Not enabled Then
            BtnDel.Enabled = enabled
        End If
        BtnSave.Enabled = enabled
        '追加はキー項目を全て指定している場合のみ可
        If enabled _
            AndAlso Kyoku <> HyohonConst.識別対象外値 _
            AndAlso Kyoten <> HyohonConst.識別対象外値 _
            AndAlso KeieiKeitai <> HyohonConst.識別対象外値 _
            AndAlso (Not CboEinoRuikei.Enabled OrElse EinoRuikei <> HyohonConst.識別対象外値) _
            AndAlso (Not CboSeisanhi.Enabled OrElse Seisanhi <> HyohonConst.識別対象外値) _
            AndAlso (Not CboTahata.Enabled OrElse Tahata <> HyohonConst.識別対象外値) Then
            BtnAdd.Enabled = True
        Else
            BtnAdd.Enabled = False
        End If

    End Sub

    ''' <summary>
    ''' 変更有無判定
    ''' </summary>
    ''' <returns></returns>
    Private Function IsModified() As Boolean
        '削除情報がある場合、変更あり
        If DelInfo.Count > 0 Then
            Return True
        End If

        '一覧の変更区分に「変更なし」以外がある場合、変更あり
        If dgvList.Rows.Cast(Of DataGridViewRow).Where(Function(x) x.Cells(変更区分列).Value?.ToString() <> 変更区分.変更なし.ToString()).Count > 0 Then
            Return True
        End If

        '変更なし
        Return False
    End Function

    ''' <summary>
    ''' 検索値(数字と記号(,-)のみ)に適合するか否かを判定
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <returns></returns>
    Private Shared Function IsFitToSearchValue(txt As String) As Boolean
        If String.IsNullOrEmpty(txt) Then
            Return True
        End If
        Return txt.Split(","c).SelectMany(Function(x) x.Split("-"c)) _
            .Where(Function(x) Not String.IsNullOrEmpty(x) _
                       AndAlso Not Integer.TryParse(x, Nothing)).Count = 0
    End Function

    ''' <summary>
    ''' ,-で指定された検索条件からSQLを作成
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="sbSql"></param>
    ''' <param name="para"></param>
    ''' <param name="col"></param>
    ''' <param name="searchValue"></param>
    Private Shared Sub AppendSearchValue(db As DBAccess, sbSql As StringBuilder, para As List(Of DBAccess.Parameter), col As String, searchValue As String)
        If String.IsNullOrEmpty(searchValue) Then
            Return
        End If
        Dim isFirst = True
        Dim cds = searchValue.Split(","c).ToArray()
        For i = 0 To UBound(cds)
            If String.IsNullOrEmpty(cds(i)) Then
                Continue For
            End If
            If cds(i).Contains("-") Then
                '範囲指定
                If cds(i).Replace("-", "").Length = 0 Then
                    '-のみ
                    Continue For
                End If
                Dim range = cds(i).Split("-"c)
                If isFirst Then
                    sbSql.Append(" AND (")
                    isFirst = False
                Else
                    sbSql.Append(" OR ")
                End If
                '下限
                Dim existsKagen = False
                If Not String.IsNullOrEmpty(range(0)) Then
                    sbSql.Append(String.Format(" HL.{0} >= @{0}下限{1}", col, i))
                    para.Add(db.CreateParameter(String.Format("@{0}下限{1}", col, i), SqlDbType.Int, range(0)))
                    existsKagen = True
                End If
                '上限
                If Not String.IsNullOrEmpty(range(range.Length - 1)) Then
                    If existsKagen Then
                        sbSql.Append(" AND ")
                    End If
                    sbSql.Append(String.Format(" HL.{0} <= @{0}上限{1}", col, i))
                    para.Add(db.CreateParameter(String.Format("@{0}上限{1}", col, i), SqlDbType.Int, range(range.Length - 1)))
                End If
            Else
                If isFirst Then
                    sbSql.Append(" AND (")
                    isFirst = False
                Else
                    sbSql.Append(" OR ")
                End If
                sbSql.Append(String.Format(" HL.{0} = @{0}_{1}", col, i))
                para.Add(db.CreateParameter(String.Format("@{0}_{1}", col, i), SqlDbType.Int, cds(i)))
            End If
        Next
        If Not isFirst Then
            sbSql.Append(")")
        End If
    End Sub

    ''' <summary>
    ''' 標本リスト検索
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="progressDialog"></param>
    ''' <param name="resultCount"></param>
    ''' <param name="maxResult"></param>
    ''' <returns></returns>
    Private Function Select標本リスト(db As DBAccess, progressDialog As ProgressDialog, ByRef resultCount As Integer, Optional maxResult As Integer = 0) As SqlDataReader

        'SELECT
        Dim sbSelect = New StringBuilder()
        Dim paraSelect = New List(Of DBAccess.Parameter)
        sbSelect.Append("SELECT")
        If maxResult > 0 Then
            '最大表示件数
            sbSelect.Append(String.Format(" TOP({0})", maxResult + 1))
        End If
        '標本リスト列（文字列化して取得）
        For i = 0 To UBound(ColInfo)
            Select Case ColInfo(i).COLUMN_ID
                Case HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.都道府県CD),
                     HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.振興局CD),
                     HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.市区町村CD),
                     HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.旧市区町村CD),
                     HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.農業集落CD),
                     HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.調査区CD),
                     HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.客体番号CD)
                    '都道府県CD～客体番号CDは0埋め
                    sbSelect.Append(String.Format(" CASE WHEN HL.{0} IS NULL THEN '' ELSE FORMAT(HL.{0},'{1}') END AS {0},", ColInfo(i).NAME, StrDup(ColInfo(i).MAX_LENGTH, "0"c)))
                Case Else
                    sbSelect.Append(String.Format(" ISNULL(CONVERT(NVARCHAR({0}), HL.{1}),'') AS {1},", ColInfo(i).MAX_LENGTH + 1, ColInfo(i).NAME))
            End Select
        Next
        '更新情報
        sbSelect.Append(" FORMAT(HL.更新日付, 'yyyy/MM/dd HH:mm') AS 更新日時,")
        sbSelect.Append(" HL.更新者ID,")
        'キー項目
        sbSelect.Append(" HL.農政局,")
        sbSelect.Append(" HL.都道府県,")
        sbSelect.Append(" HL.実査設置拠点,")
        sbSelect.Append(" HL.経営形態区分,")
        sbSelect.Append(" HL.営農類型区分,")
        sbSelect.Append(" HL.生産費区分,")
        sbSelect.Append(" HL.田畑区分")

        'FROM～WHERE
        Dim sbFromWhere = New StringBuilder()
        Dim paraFromWhere = New List(Of DBAccess.Parameter)

        sbFromWhere.Append(String.Format(" FROM 標本リスト{0} HL", CensusNen))
        '規模階層
        Dim kiboKaiso = HyohonUtil.GetComboValue(CboKiboKaiso)
        Dim tahata = HyohonUtil.GetComboValue(CboTahata)
        If kiboKaiso <> HyohonConst.識別対象外値 Then
            sbFromWhere.Append(" JOIN 標本リスト設定 S")
            sbFromWhere.Append(" ON S.経営形態区分 = HL.経営形態区分")
            sbFromWhere.Append(" AND S.営農類型区分 = HL.営農類型区分")
            sbFromWhere.Append(" AND S.生産費区分 = HL.生産費区分")
            '主副業区分（営農個人以外は識別対象外）
            sbFromWhere.Append(" AND S.主副業区分 = IIF(HL.経営形態区分 <> @主副業＿経営形態区分, @主副業＿識別対象外,")
            paraFromWhere.Add(db.CreateParameter("@主副業＿経営形態区分", SqlDbType.Int, HyohonConst.経営形態区分.個人経営体))
            paraFromWhere.Add(db.CreateParameter("@主副業＿識別対象外", SqlDbType.Int, HyohonConst.識別対象外値))
            sbFromWhere.Append(String.Format(" IIF(HL.NO{0} = @主副業＿副業, @主副業選択肢＿副業, @主副業選択肢＿主業＿準主業))", HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.主副業区分)))
            paraFromWhere.Add(db.CreateParameter("@主副業＿副業", SqlDbType.Int, HyohonConst.主副業区分.副業))
            paraFromWhere.Add(db.CreateParameter("@主副業選択肢＿副業", SqlDbType.Int, HyohonConst.主副業区分選択肢.副業))
            paraFromWhere.Add(db.CreateParameter("@主副業選択肢＿主業＿準主業", SqlDbType.Int, HyohonConst.主副業区分選択肢.主業＿準主業))
            '青色申告区分（営農個人副業以外は識別対象外）
            sbFromWhere.Append(String.Format(" AND S.青色申告区分 = IIF(HL.経営形態区分 <> @青色＿経営形態区分 OR HL.NO{0} <> @青色＿主副業区分, @青色＿識別対象外,", HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.主副業区分)))
            paraFromWhere.Add(db.CreateParameter("@青色＿経営形態区分", SqlDbType.Int, HyohonConst.経営形態区分.個人経営体))
            paraFromWhere.Add(db.CreateParameter("@青色＿主副業区分", SqlDbType.Int, HyohonConst.主副業区分.副業))
            paraFromWhere.Add(db.CreateParameter("@青色＿識別対象外", SqlDbType.Int, HyohonConst.識別対象外値))
            sbFromWhere.Append(String.Format(" HL.NO{0})", HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.青色申告区分)))
            '集落営農区分（営農法人水田作以外は識別対象外）
            sbFromWhere.Append(" AND S.集落営農区分 = IIF(HL.経営形態区分 <> @集落＿経営形態区分 OR HL.営農類型区分 <> @集落＿営農類型区分, @集落＿識別対象外,")
            paraFromWhere.Add(db.CreateParameter("@集落＿経営形態区分", SqlDbType.Int, HyohonConst.経営形態区分.法人経営体))
            paraFromWhere.Add(db.CreateParameter("@集落＿営農類型区分", SqlDbType.Int, HyohonConst.営農類型区分.水田作))
            paraFromWhere.Add(db.CreateParameter("@集落＿識別対象外", SqlDbType.Int, HyohonConst.識別対象外値))
            sbFromWhere.Append(String.Format(" HL.NO{0})", HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.集落営農区分)))
            If tahata <> HyohonConst.識別対象外値 Then
                sbFromWhere.Append(" AND S.田畑区分 = HL.田畑区分")
            Else
                sbFromWhere.Append(" AND S.田畑区分 = @田畑＿識別対象外")
                paraFromWhere.Add(db.CreateParameter("@田畑＿識別対象外", SqlDbType.Int, HyohonConst.識別対象外値))
            End If
            sbFromWhere.Append(" JOIN 標本リスト設定＿抽出条件 SC")
            sbFromWhere.Append(" ON SC.ID = S.ID")
            sbFromWhere.Append(" AND SC.規模番号 = @規模番号")
            paraFromWhere.Add(db.CreateParameter("@規模番号", SqlDbType.Int, kiboKaiso))
        End If

        '利用者の所属による暗黙の条件
        If CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Honsyo Then
            '本省はないため、DUMMYの条件を設定
            sbFromWhere.Append(" WHERE 1 = 1")
        ElseIf CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Kyoku Then
            '局
            sbFromWhere.Append(" WHERE HL.農政局 = @利用者の農政局")
            paraFromWhere.Add(db.CreateParameter("@利用者の農政局", SqlDbType.Int, CommonInfo.Kyoku))
        Else
            '実査設置拠点
            sbFromWhere.Append(" WHERE HL.農政局 = @利用者の農政局")
            paraFromWhere.Add(db.CreateParameter("@利用者の農政局", SqlDbType.Int, CommonInfo.Kyoku))
            sbFromWhere.Append(" AND HL.都道府県 = @利用者の都道府県")
            paraFromWhere.Add(db.CreateParameter("@利用者の都道府県", SqlDbType.Int, IIf(CommonInfo.Jimusyo = "51", "1", CommonInfo.Jimusyo)))
            sbFromWhere.Append(" AND HL.実査設置拠点 = @利用者の実査設置拠点")
            paraFromWhere.Add(db.CreateParameter("@利用者の実査設置拠点", SqlDbType.Int, CommonInfo.Center))
        End If
        '農政局
        Dim kyoku = HyohonUtil.GetComboValue(CboKyoku)
        If kyoku <> HyohonConst.識別対象外値 Then
            sbFromWhere.Append(" AND HL.農政局 = @農政局")
            paraFromWhere.Add(db.CreateParameter("@農政局", SqlDbType.Int, kyoku))
        End If
        '実査設置拠点
        Dim kyoten = HyohonUtil.GetComboValue(CboKyoten)
        If kyoten <> HyohonConst.識別対象外値 Then
            sbFromWhere.Append(" AND HL.都道府県 = @都道府県")
            Dim jimusho = CInt(CType(CboKyoten.SelectedItem, DataRowView)("事務所番号"))
            paraFromWhere.Add(db.CreateParameter("@都道府県", SqlDbType.Int, IIf(jimusho = 51, 1, jimusho)))
            sbFromWhere.Append(" AND HL.実査設置拠点 = @実査設置拠点")
            paraFromWhere.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kyoten))
        End If
        '経営形態区分
        Dim keieiKeitai = HyohonUtil.GetComboValue(CboKeieiKeitai)
        If keieiKeitai <> HyohonConst.識別対象外値 Then
            sbFromWhere.Append(" AND HL.経営形態区分 = @経営形態区分")
            paraFromWhere.Add(db.CreateParameter("@経営形態区分", SqlDbType.Int, keieiKeitai))
        End If
        '営農類型区分
        Dim einoRuikei = HyohonUtil.GetComboValue(CboEinoRuikei)
        If einoRuikei <> HyohonConst.識別対象外値 Then
            sbFromWhere.Append(" AND HL.営農類型区分 = @営農類型区分")
            paraFromWhere.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoRuikei))
        End If
        '生産費区分
        Dim seisanhi = HyohonUtil.GetComboValue(CboSeisanhi)
        If seisanhi <> HyohonConst.識別対象外値 Then
            sbFromWhere.Append(" AND HL.生産費区分 = @生産費区分")
            paraFromWhere.Add(db.CreateParameter("@生産費区分", SqlDbType.Int, seisanhi))
        End If
        '田畑区分
        If tahata <> HyohonConst.識別対象外値 Then
            sbFromWhere.Append(" AND HL.田畑区分 = @田畑区分")
            paraFromWhere.Add(db.CreateParameter("@田畑区分", SqlDbType.Int, tahata))
        End If
        '都道府県CD
        AppendSearchValue(db, sbFromWhere, paraFromWhere, "NO2", TxtTodofuken.Text)
        '市区町村CD
        AppendSearchValue(db, sbFromWhere, paraFromWhere, "NO4", TxtShikuchoson.Text)
        '標本区分
        Dim hyohon = HyohonUtil.GetComboValue(CboHyohon)
        If hyohon <> HyohonConst.識別対象外値 Then
            sbFromWhere.Append(String.Format(" AND HL.NO{0} = @標本区分", HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.標本区分)))
            paraFromWhere.Add(db.CreateParameter("@標本区分", SqlDbType.Int, hyohon))
        End If
        '標本候補区分
        Dim hyohonKoho = HyohonUtil.GetComboValue(CboHyohonKoho)
        If hyohonKoho <> HyohonConst.識別対象外値 Then
            sbFromWhere.Append(String.Format(" AND HL.NO{0} = @標本候補区分", HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.標本候補区分)))
            paraFromWhere.Add(db.CreateParameter("@標本候補区分", SqlDbType.Int, hyohonKoho))
        End If
        '一連番号
        If Not String.IsNullOrEmpty(TxtIchirenNo.Text) Then
            sbFromWhere.Append(" AND (")
            Dim appendOR = False
            Dim list = Split(TxtIchirenNo.Text, vbCrLf)
            For i = 0 To UBound(list)
                If String.IsNullOrEmpty(list(i)) Then
                    Continue For
                End If
                If appendOR Then
                    sbFromWhere.Append(" OR ")
                Else
                    appendOR = True
                End If
                sbFromWhere.Append(String.Format(" HL.NO1 = @一連番号{0}", i))
                paraFromWhere.Add(db.CreateParameter(String.Format("@一連番号{0}", i), SqlDbType.VarChar, list(i)))
            Next
            sbFromWhere.Append(")")
        End If
        '規模階層
        If kiboKaiso <> HyohonConst.識別対象外値 Then
            sbFromWhere.Append(" AND (")
            Dim strOr = ""
            For Each kv In HyohonConst.標本リスト比較列番号(CensusNen)
                sbFromWhere.Append(String.Format(" {0} HL.経営形態区分 = {1}", strOr, kv.Key.経営形態区分.ToString("d")))
                sbFromWhere.Append(String.Format(" AND HL.営農類型区分 = {0}", kv.Key.営農類型区分.ToString("d")))
                sbFromWhere.Append(String.Format(" AND ISNULL(HL.NO{0}, -1) = IIF({1} = {2}, ISNULL(HL.NO{0}, -1), {1})", HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.集落営農区分), kv.Key.集落営農区分.ToString("d"), HyohonConst.識別対象外値))
                sbFromWhere.Append(String.Format(" AND HL.生産費区分 = {0}", kv.Key.生産費区分.ToString("d")))
                If tahata <> HyohonConst.識別対象外値 Then
                    sbFromWhere.Append(String.Format(" AND HL.田畑区分 = {0}", kv.Key.田畑区分.ToString("d")))
                End If
                sbFromWhere.Append(String.Format(" AND HL.NO{0} < IIF(SC.下限 IS NOT NULL, ISNULL(SC.上限, {1}), ISNULL(SC.上限, {2}))", kv.Value, HyohonConst.標本リスト設定＿抽出上限値, -HyohonConst.標本リスト設定＿抽出上限値))
                sbFromWhere.Append(String.Format(" AND HL.NO{0} >= ISNULL(SC.下限,{1})", kv.Value, -HyohonConst.標本リスト設定＿抽出上限値))
                strOr = "OR"
            Next
            sbFromWhere.Append(" )")
        End If

        '件数検索
        If maxResult > 0 Then
            resultCount = CInt(db.ExecuteScalar(String.Format("SELECT COUNT(*) FROM (SELECT TOP({0}) NO1", maxResult + 1) & sbFromWhere.ToString() & ") WRK", paraFromWhere))
        Else
            resultCount = CInt(db.ExecuteScalar("SELECT COUNT(*)" & sbFromWhere.ToString(), paraFromWhere))
        End If
        progressDialog.Value = 50

        'ORDER BY
        sbFromWhere.Append(" ORDER BY")
        If maxResult = 0 Then
            'CSV出力時
            sbFromWhere.Append(" HL.経営形態区分,")
            sbFromWhere.Append(" HL.営農類型区分,")
            sbFromWhere.Append(" HL.生産費区分,")
            sbFromWhere.Append(" HL.田畑区分,")
            sbFromWhere.Append(" HL.農政局,")
            sbFromWhere.Append(" HL.都道府県,")
            sbFromWhere.Append(" HL.実査設置拠点,")
            sbFromWhere.Append(" HL.NO1")
        Else
            '表示時
            sbFromWhere.Append(" HL.NO1,")
            sbFromWhere.Append(" HL.経営形態区分,")
            sbFromWhere.Append(" HL.営農類型区分,")
            sbFromWhere.Append(" HL.生産費区分,")
            sbFromWhere.Append(" HL.田畑区分")
        End If

        'DataReader取得
        paraSelect.AddRange(paraFromWhere)
        Dim dr = db.ExecuteReader(sbSelect.ToString() & sbFromWhere.ToString(), paraSelect)
        progressDialog.Value = 100
        Return dr

    End Function

    ''' <summary>
    ''' CSV出力処理
    ''' </summary>
    ''' <param name="csvOutPattern"></param>
    ''' <param name="dr"></param>
    ''' <param name="resultCount"></param>
    ''' <param name="folderPath"></param>
    ''' <param name="progressDialog"></param>
    ''' <returns></returns>
    Private Sub CsvOutput(csvOutPattern As HyohonConst.CSV出力パターン, dr As SqlDataReader, resultCount As Integer, folderPath As String, progressDialog As ProgressDialog)

        Dim helper = New CSVOutputHelper(csvOutPattern)
        Dim readCount = 0
        Dim eof = Not dr.Read()
        While Not eof
            helper.SetKey(dr)
            Dim csvFilePath = folderPath & Path.DirectorySeparatorChar & helper.GetCsvFileName()
            Using sw As New StreamWriter(csvFilePath, False, Encoding.GetEncoding("shift_jis"))
                'ヘッダ出力
                sw.WriteLine(String.Join(",", HyohonHeaders.Where(Function(x, index) index > 0)))
                'データ出力
                While Not eof AndAlso Not helper.SetKey(dr)
                    readCount += 1
                    Dim row = New List(Of Object)
                    For i = 0 To HyohonColCount - 1
                        row.Add(dr(i))
                    Next
                    sw.WriteLine(String.Join(",", row))
                    eof = Not dr.Read()
                End While
            End Using
            '進捗を進める
            progressDialog.Value = CInt(Math.Truncate(readCount / resultCount * 100)) + 100
        End While

    End Sub

    ''' <summary>
    ''' CSV出力ヘルパー
    ''' </summary>
    Private Class CSVOutputHelper
        Private ReadOnly CsvOutPattern As HyohonConst.CSV出力パターン
        Private ReadOnly TimeStamp As String = Now.ToString("yyyyMMdd_HHmmss")
        'NewKeys
        Private NewKeieiKeitai As HyohonConst.経営形態区分
        Private NewEinoRuikei As HyohonConst.営農類型区分
        Private NewSeisanhi As HyohonConst.生産費区分
        Private NewTahata As HyohonConst.田畑区分
        Private NewKyoku As Integer
        Private NewPref As Integer
        Private NewKyoten As Integer
        'OldKeys
        Private OldKeieiKeitai As HyohonConst.経営形態区分 = HyohonConst.経営形態区分.識別対象外
        Private OldEinoRuikei As HyohonConst.営農類型区分
        Private OldSeisanhi As HyohonConst.生産費区分
        Private OldTahata As HyohonConst.田畑区分
        Private OldKyoku As Integer
        Private OldPref As Integer
        Private OldKyoten As Integer

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="csvOutPattern"></param>
        Sub New(csvOutPattern As HyohonConst.CSV出力パターン)
            Me.CsvOutPattern = csvOutPattern
        End Sub

        ''' <summary>
        ''' キーをセット
        ''' </summary>
        ''' <param name="dr"></param>
        Function SetKey(dr As SqlDataReader) As Boolean
            NewKeieiKeitai = DirectCast(dr("経営形態区分"), HyohonConst.経営形態区分)
            NewEinoRuikei = DirectCast(dr("営農類型区分"), HyohonConst.営農類型区分)
            NewSeisanhi = DirectCast(dr("生産費区分"), HyohonConst.生産費区分)
            NewTahata = DirectCast(dr("田畑区分"), HyohonConst.田畑区分)
            NewKyoku = CInt(dr("農政局"))
            NewPref = CInt(dr("都道府県"))
            NewKyoten = CInt(dr("実査設置拠点"))
            Dim isBreak = Me.IsBreak()
            If isBreak Then
                ShiftKey()
            End If
            Return isBreak
        End Function

        ''' <summary>
        ''' キーをシフト
        ''' </summary>
        Private Sub ShiftKey()
            OldKeieiKeitai = NewKeieiKeitai
            OldEinoRuikei = NewEinoRuikei
            OldSeisanhi = NewSeisanhi
            OldTahata = NewTahata
            OldKyoku = NewKyoku
            OldPref = NewPref
            OldKyoten = NewKyoten
        End Sub

        ''' <summary>
        ''' キーブレイク判定
        ''' </summary>
        ''' <returns></returns>
        Private Function IsBreak() As Boolean
            If NewKeieiKeitai <> OldKeieiKeitai _
               OrElse NewEinoRuikei <> OldEinoRuikei _
               OrElse NewSeisanhi <> OldSeisanhi _
               OrElse NewTahata <> OldTahata Then
                Return True
            ElseIf CsvOutPattern = HyohonConst.CSV出力パターン.農政局 Then
                If NewKyoku <> OldKyoku Then
                    Return True
                End If
            ElseIf CsvOutPattern = HyohonConst.CSV出力パターン.実査設置拠点 Then
                If NewPref <> OldPref _
                    OrElse NewKyoten <> OldKyoten Then
                    Return True
                End If
            End If
            Return False
        End Function

        ''' <summary>
        ''' CSVファイル名を取得
        ''' </summary>
        ''' <returns></returns>
        Function GetCsvFileName() As String
            Dim csvFileName = New StringBuilder()
            Select Case CsvOutPattern
                Case HyohonConst.CSV出力パターン.全国
                    csvFileName.Append("全国")
                Case HyohonConst.CSV出力パターン.農政局
                    csvFileName.Append(NewKyoku)
                    csvFileName.Append("_").Append(MasterDao.GetKyokuName(NewKyoku.ToString()).Replace("　", "").Replace(" ", ""))
                Case HyohonConst.CSV出力パターン.実査設置拠点
                    csvFileName.Append(If(NewPref = 1, 51, NewPref))
                    csvFileName.Append("_").Append(NewKyoten)
                    csvFileName.Append("_").Append(MasterDao.GetCenterName(If(NewPref = 1, 51, NewPref).ToString(), NewKyoten.ToString()).Replace("　", "").Replace(" ", ""))
            End Select
            Select Case NewKeieiKeitai
                Case HyohonConst.経営形態区分.個人経営体, HyohonConst.経営形態区分.法人経営体
                    csvFileName.Append("_").Append(HyohonConst.経営形態区分辞書(NewKeieiKeitai).略称)
                    csvFileName.Append("_").Append(NewEinoRuikei)
                    csvFileName.Append("_").Append(HyohonUtil.Getリスト名称(HyohonConst.営農類型区分リスト, NewEinoRuikei))
                Case HyohonConst.経営形態区分.個別経営体, HyohonConst.経営形態区分.組織法人経営体
                    If HyohonConst.生産費区分＿農畜生辞書(NewSeisanhi) = HyohonConst.農畜生区分.農生 Then
                        csvFileName.Append("_農生")
                    Else
                        csvFileName.Append("_畜生")
                    End If
                    csvFileName.Append(HyohonConst.経営形態区分辞書(NewKeieiKeitai).略称)
                    csvFileName.Append(HyohonUtil.Getリスト名称(HyohonUtil.Get生産費区分リスト(NewKeieiKeitai), NewSeisanhi))
                    csvFileName.Append("生産費")
                    If NewTahata <> HyohonConst.田畑区分.識別対象外 Then
                        csvFileName.Append("_").Append(HyohonUtil.Getリスト名称(HyohonConst.田畑区分リスト, NewTahata)).Append("作")
                    End If
            End Select
            csvFileName.Append("_").Append(TimeStamp).Append(".csv")
            Return csvFileName.ToString()
        End Function
    End Class

    ''' <summary>
    ''' 保存時チェック
    ''' </summary>
    ''' <returns>エラーあり:エラーセル、エラーなし:Nothing</returns>
    Private Function CheckForSave() As DataGridViewCell
        '追加行のチェック
        Dim addRows = dgvList.Rows.Cast(Of DataGridViewRow) _
                .Where(Function(x) CInt(x.Cells(変更区分列).Value) = 変更区分.追加).ToList()
        '追加行内の重複チェック
        Dim duplicate = addRows.Cast(Of DataGridViewRow).GroupBy(Function(x) x.Cells(一連番号列).Value.ToString()).Where(Function(x) x.Count() > 1).FirstOrDefault()
        If duplicate IsNot Nothing Then
            Message.ShowMsgBox(MessageID.MSG_E_129, {"[" & duplicate(0).Cells(一連番号列).Value.ToString() & "]"}, MsgBoxStyle.OkOnly)
            Return duplicate(0).Cells(一連番号列)
        End If
        '追加行の登録済チェック
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '1000件ずつにグループ化（IN句の上限1000のため）
            Dim grooped = addRows.Select(Function(x, index) _
                                             (key:=(ichirenNo:=x.Cells(一連番号列).Value.ToString(),
                                             keieiKeitai:=CInt(x.Cells(経営形態区分列).Value),
                                             einoRuikei:=CInt(x.Cells(営農類型区分列).Value),
                                             seisanhi:=CInt(x.Cells(生産費区分列).Value),
                                             tahata:=CInt(x.Cells(田畑区分列).Value)), index)) _
                .Where(Function(x) Not DelInfo.Contains(x.key)) _
                .GroupBy(Function(x) CInt(Math.Floor(x.index / 1000))).ToList()
            For Each group In grooped
                Dim sbSql = New StringBuilder()
                Dim para = New List(Of DBAccess.Parameter)
                sbSql.Append("SELECT TOP(1) NO1 FROM 標本リスト{0} WHERE")
                For i = 0 To group.Count - 1
                    If para.Count > 0 Then
                        sbSql.Append(" OR")
                    End If
                    sbSql.Append(String.Format(" NO1 = @NO1_{0}", i))
                    para.Add(db.CreateParameter(String.Format("@NO1_{0}", i), SqlDbType.VarChar, group(i).key.ichirenNo))
                    sbSql.Append(String.Format(" AND 経営形態区分 = @経営形態区分_{0}", i))
                    para.Add(db.CreateParameter(String.Format("@経営形態区分_{0}", i), SqlDbType.Int, group(i).key.keieiKeitai))
                    sbSql.Append(String.Format(" AND 営農類型区分 = @営農類型区分_{0}", i))
                    para.Add(db.CreateParameter(String.Format("@営農類型区分_{0}", i), SqlDbType.Int, group(i).key.einoRuikei))
                    sbSql.Append(String.Format(" AND 生産費区分 = @生産費区分_{0}", i))
                    para.Add(db.CreateParameter(String.Format("@生産費区分_{0}", i), SqlDbType.Int, group(i).key.seisanhi))
                    sbSql.Append(String.Format(" AND 田畑区分 = @田畑区分_{0}", i))
                    para.Add(db.CreateParameter(String.Format("@田畑区分_{0}", i), SqlDbType.Int, group(i).key.tahata))
                Next
                sbSql.Append(" ORDER BY NO1")
                Dim dt = db.GetDataTable(String.Format(sbSql.ToString(), CensusNen), para)
                If dt.Rows.Count() > 0 Then
                    Dim no1 = dt.Rows(0)("NO1").ToString()
                    Message.ShowMsgBox(MessageID.MSG_E_130, {"[" & no1 & "]"}, MsgBoxStyle.OkOnly)
                    Return dgvList.Rows.Cast(Of DataGridViewRow).Where(Function(x) x.Cells(変更区分列).Value?.ToString() = 変更区分.追加.ToString() AndAlso x.Cells(一連番号列).Value.ToString() = no1).Select(Function(x) x.Cells(一連番号列)).FirstOrDefault
                End If
            Next
        End Using

        '変更、追加行のチェック
        Dim modifiedRows = dgvList.Rows.Cast(Of DataGridViewRow) _
            .Where(Function(x) CInt(x.Cells(変更区分列).Value) <> 変更区分.変更なし).ToList()
        Dim cell As DataGridViewCell
        Dim cellValue = ""
        Dim jimusho = CType(CboKyoten.SelectedItem, DataRowView)("事務所番号").ToString()
        Dim kubunInfo = HyohonConst.標本リスト区分チェック情報(CensusNen)
        kubunInfo.AddRange(HyohonConst.標本リスト区分クリア情報(CensusNen))
        '負数あり列
        Dim kiboNo = HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.営農類型区分規模)
        For Each row In modifiedRows
            '列情報に基づくチェック
            For i = 1 To HyohonColCount
                cell = row.Cells(i)
                cellValue = If(cell.Value?.ToString(), "")
                '必須
                If ColInfo(i - 1).IS_NULLABLE Then
                    If String.IsNullOrEmpty(cellValue) Then
                        Continue For
                    End If
                Else
                    If String.IsNullOrEmpty(cellValue) Then
                        Message.ShowMsgBox(MessageID.MSG_E_127, {i.ToString() & "．" & HyohonHeaders(i)}, MsgBoxStyle.OkOnly)
                        Return cell
                    End If
                End If
                '型、桁数
                Select Case ColInfo(i - 1).TYPE
                    Case "CHAR", "VARCHAR", "NVARCHAR"
                        If cellValue.Length > ColInfo(i - 1).MAX_LENGTH Then
                            Message.ShowMsgBox(MessageID.MSG_E_128, {i.ToString() & "．" & HyohonHeaders(i), ColInfo(i - 1).MAX_LENGTH.ToString() & "文字以内"}, MsgBoxStyle.OkOnly)
                            Return cell
                        End If
                    Case "NUMERIC"
                        If ColInfo(i - 1).SCALE > 0 Then
                            Dim decVal As Decimal
                            If Not Decimal.TryParse(cellValue, decVal) _
                                OrElse Math.Truncate(decVal).ToString().Replace("-", "").Length > ColInfo(i - 1).PRECISION _
                                OrElse (Decimal.GetBits(decVal)(3) >> 16 And &HFF) > ColInfo(i - 1).SCALE _
                                OrElse kiboNo <> i AndAlso decVal < 0 Then
                                Message.ShowMsgBox(MessageID.MSG_E_128, {i.ToString() & "．" & HyohonHeaders(i), String.Format("{0}整数:{1}桁、小数:{2}桁以内", If(i = kiboNo, "", "正の"), ColInfo(i - 1).PRECISION, ColInfo(i - 1).SCALE)}, MsgBoxStyle.OkOnly)
                                Return cell
                            End If
                        Else
                            Dim decVal As Decimal
                            If Not Decimal.TryParse(cellValue, decVal) _
                                OrElse cellValue.Contains(".") _
                                OrElse decVal.ToString().Replace("-", "").Length > ColInfo(i - 1).PRECISION _
                                OrElse kiboNo <> i AndAlso decVal < 0 Then
                                Message.ShowMsgBox(MessageID.MSG_E_128, {i.ToString() & "．" & HyohonHeaders(i), String.Format("{0}整数:{1}桁以内", If(i = kiboNo, "", "正の"), ColInfo(i - 1).PRECISION)}, MsgBoxStyle.OkOnly)
                                Return cell
                            End If
                        End If
                    Case Else
                        Throw New Exception("CHAR,VARCHAR,NVARCHAR,NUMERIC以外の型は未実装です。")
                End Select
            Next

            '都道府県CDが実査設置拠点の都道府県CD
            If Not String.IsNullOrEmpty(jimusho) Then
                cell = row.Cells(HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.都道府県CD))
                cellValue = If(cell.Value?.ToString(), "0")
                Dim pref = CInt(If(jimusho = "51", "1", jimusho))
                If CInt(cellValue) <> pref Then
                    Message.ShowMsgBox(MessageID.MSG_E_128, {"2．" & HyohonHeaders(2), "[" & pref & "]"}, MsgBoxStyle.OkOnly)
                    Return cell
                End If
            End If

            '区分情報に基づくチェック
            For Each info In kubunInfo
                cell = row.Cells(info.NO)
                cellValue = If(cell.Value?.ToString(), "")
                If ColInfo(info.NO).IS_NULLABLE AndAlso String.IsNullOrEmpty(cellValue) Then
                    Continue For
                End If
                If info.RANGE.Contains("～") Then
                    Dim vals = Split(info.RANGE, "～")
                    If CDec(cellValue) < CDec(vals(0)) OrElse CDec(cellValue) > CDec(vals(1)) Then
                        Message.ShowMsgBox(MessageID.MSG_E_128, {info.NO.ToString() & "．" & HyohonHeaders(info.NO), String.Format("[{0}]", info.RANGE)}, MsgBoxStyle.OkOnly)
                        Return cell
                    End If
                Else
                    If Not info.RANGE.Split(","c).Contains(cellValue) Then
                        Message.ShowMsgBox(MessageID.MSG_E_128, {info.NO.ToString() & "．" & HyohonHeaders(info.NO), String.Format("[{0}]", info.RANGE)}, MsgBoxStyle.OkOnly)
                        Return cell
                    End If
                End If
            Next

            '営農類型区分
            Dim einoNo = HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.営農類型区分)
            cell = row.Cells(einoNo)
            cellValue = If(cell.Value?.ToString(), "")
            If CboEinoRuikei.Enabled AndAlso Not String.IsNullOrEmpty(cellValue) Then
                If EinoRuikei.ToString() <> cellValue Then
                    Message.ShowMsgBox(MessageID.MSG_E_128, {einoNo.ToString() & "．" & HyohonHeaders(einoNo), "[" & EinoRuikei.ToString() & "]"}, MsgBoxStyle.OkOnly)
                    Return cell
                End If
            End If
        Next

        'エラー無し
        Return Nothing
    End Function

    ''' <summary>
    ''' 標本リスト保存
    ''' </summary>
    Private Sub Save標本リスト()
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Try
                db.BeginTrans()

                '削除処理
                Delete標本リスト(db, CensusNen, DelInfo)

                '登録処理
                Dim addRows = dgvList.Rows.Cast(Of DataGridViewRow).Where(Function(x) CInt(x.Cells(変更区分列).Value) = 変更区分.追加).ToList()
                Insert標本リスト(db, CensusNen, addRows, Kyoku, Pref, Kyoten, KeieiKeitai, EinoRuikei, Seisanhi, Tahata, ColInfo)

                '更新処理
                Dim updRows = dgvList.Rows.Cast(Of DataGridViewRow).Where(Function(x) CInt(x.Cells(変更区分列).Value) = 変更区分.変更あり).ToList()
                Update標本リスト(db, CensusNen, 一連番号列, 経営形態区分列, 営農類型区分列, 生産費区分列, 田畑区分列, updRows, ColInfo)

                db.CommitTrans()
            Catch ex As Exception
                db.RollBackTrans()
                Throw ex
            End Try
        End Using
    End Sub

    ''' <summary>
    ''' 標本リスト削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNen"></param>
    ''' <param name="delInfo"></param>
    ''' <returns></returns>
    Private Shared Function Delete標本リスト(db As DBAccess, censusNen As Integer, delInfo As HashSet(Of (ichirenNo As String, keieiKeitai As Integer, einoRuikei As Integer, seisanhi As Integer, tahata As Integer))) As Integer
        If delInfo.Count <= 0 Then
            Return 0
        End If

        Dim count = 0
        '400件ずつにグループ化（パラメータ数の上限が2100のため）
        Dim grouped = delInfo.ToList().Select(Function(x, index) _
            (key:=(x.ichirenNo, x.keieiKeitai, x.einoRuikei, x.seisanhi, x.tahata), index)) _
            .GroupBy(Function(x) CInt(Math.Floor(x.index / 400)))
        For Each group In grouped
            Dim sbSql = New StringBuilder()
            Dim para = New List(Of DBAccess.Parameter)
            sbSql.Append("DELETE FROM 標本リスト{0} WHERE")
            For i = 0 To group.Count - 1
                If para.Count > 0 Then
                    sbSql.Append(" OR")
                End If
                sbSql.Append(String.Format(" NO1 = @NO1_{0}", i))
                para.Add(db.CreateParameter(String.Format("@NO1_{0}", i), SqlDbType.VarChar, group(i).key.ichirenNo))
                sbSql.Append(String.Format(" AND 経営形態区分 = @経営形態区分_{0}", i))
                para.Add(db.CreateParameter(String.Format("@経営形態区分_{0}", i), SqlDbType.Int, group(i).key.keieiKeitai))
                sbSql.Append(String.Format(" AND 営農類型区分 = @営農類型区分_{0}", i))
                para.Add(db.CreateParameter(String.Format("@営農類型区分_{0}", i), SqlDbType.Int, group(i).key.einoRuikei))
                sbSql.Append(String.Format(" AND 生産費区分 = @生産費区分_{0}", i))
                para.Add(db.CreateParameter(String.Format("@生産費区分_{0}", i), SqlDbType.Int, group(i).key.seisanhi))
                sbSql.Append(String.Format(" AND 田畑区分 = @田畑区分_{0}", i))
                para.Add(db.CreateParameter(String.Format("@田畑区分_{0}", i), SqlDbType.Int, group(i).key.tahata))
            Next
            count += db.ExecuteNonQuery(String.Format(sbSql.ToString(), censusNen), para)
        Next

        Return count
    End Function

    ''' <summary>
    ''' 標本リスト登録
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNen"></param>
    ''' <param name="addRows"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="pref"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="keieiKeitai"></param>
    ''' <param name="einoRuikei"></param>
    ''' <param name="seisanhi"></param>
    ''' <param name="tahata"></param>
    ''' <param name="colInfo"></param>
    ''' <returns></returns>
    Private Shared Function Insert標本リスト(db As DBAccess, censusNen As Integer, addRows As List(Of DataGridViewRow),
                                        kyoku As Integer, pref As Integer, kyoten As Integer,
                                        keieiKeitai As Integer, einoRuikei As Integer, seisanhi As Integer, tahata As Integer,
                                        colInfo As (COLUMN_ID As Integer, NAME As String, IS_NULLABLE As Boolean, TYPE As String, MAX_LENGTH As Integer, PRECISION As Integer, SCALE As Integer)()) As Integer
        If addRows.Count <= 0 Then
            Return 0
        End If

        Dim count = 0
        Dim sbSql = New StringBuilder()
        sbSql.Append(String.Format("INSERT INTO 標本リスト{0} VALUES(", censusNen))
        For i = 0 To UBound(colInfo)
            sbSql.Append(String.Format(" @{0},", colInfo(i).NAME))
        Next
        sbSql.Append("@農政局,")
        sbSql.Append("@都道府県,")
        sbSql.Append("@実査設置拠点,")
        sbSql.Append("@経営形態区分,")
        sbSql.Append("@営農類型区分,")
        sbSql.Append("@生産費区分,")
        sbSql.Append("@田畑区分,")
        sbSql.Append("SYSDATETIME(),")
        sbSql.Append("@更新者ID")
        sbSql.Append(")")
        For Each row In addRows
            Dim para = New List(Of DBAccess.Parameter)
            For i = 0 To UBound(colInfo)
                Dim sqlDbType As SqlDbType
                Select Case colInfo(i).TYPE
                    Case "CHAR", "VARCHAR", "NVARCHAR"
                        sqlDbType = SqlDbType.VarChar
                    Case Else
                        sqlDbType = SqlDbType.Decimal
                End Select
                para.Add(db.CreateParameter(String.Format("@{0}", colInfo(i).NAME), sqlDbType, row.Cells(i + 1).Value))
            Next
            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kyoku))
            para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, pref))
            para.Add(db.CreateParameter("@実査設置拠点", SqlDbType.Int, kyoten))
            para.Add(db.CreateParameter("@経営形態区分", SqlDbType.Int, keieiKeitai))
            para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoRuikei))
            para.Add(db.CreateParameter("@生産費区分", SqlDbType.Int, seisanhi))
            para.Add(db.CreateParameter("@田畑区分", SqlDbType.Int, tahata))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))
            count += db.ExecuteNonQuery(sbSql.ToString, para)
        Next

        Return count
    End Function

    ''' <summary>
    ''' 標本リスト更新
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="censusNen"></param>
    ''' <param name="一連番号列"></param>
    ''' <param name="経営形態区分列"></param>
    ''' <param name="営農類型区分列"></param>
    ''' <param name="生産費区分列"></param>
    ''' <param name="田畑区分列"></param>
    ''' <param name="updRows"></param>
    ''' <param name="colInfo"></param>
    ''' <returns></returns>
    Private Shared Function Update標本リスト(db As DBAccess, censusNen As Integer, 一連番号列 As Integer,
                                        経営形態区分列 As Integer, 営農類型区分列 As Integer, 生産費区分列 As Integer, 田畑区分列 As Integer,
                                        updRows As List(Of DataGridViewRow),
                                        colInfo As (COLUMN_ID As Integer, NAME As String, IS_NULLABLE As Boolean, TYPE As String, MAX_LENGTH As Integer, PRECISION As Integer, SCALE As Integer)()) As Integer
        If updRows.Count <= 0 Then
            Return 0
        End If

        Dim count = 0
        Dim sbSql = New StringBuilder()
        sbSql.Append(String.Format("UPDATE 標本リスト{0} SET", censusNen))
        For i = 0 To UBound(colInfo)
            sbSql.Append(String.Format(" {0} = @{0},", colInfo(i).NAME))
        Next
        sbSql.Append(" 更新日付 = SYSDATETIME(),")
        sbSql.Append(" 更新者ID = @更新者ID")
        sbSql.Append(" WHERE NO1 = @一連番号")
        sbSql.Append(" AND 経営形態区分 = @経営形態区分")
        sbSql.Append(" AND 営農類型区分 = @営農類型区分")
        sbSql.Append(" AND 生産費区分 = @生産費区分")
        sbSql.Append(" AND 田畑区分 = @田畑区分")
        For Each row In updRows
            Dim para = New List(Of DBAccess.Parameter)
            For i = 0 To UBound(colInfo)
                Dim sqlDbType As SqlDbType
                Select Case colInfo(i).TYPE
                    Case "CHAR", "VARCHAR", "NVARCHAR"
                        sqlDbType = SqlDbType.VarChar
                    Case Else
                        sqlDbType = SqlDbType.Decimal
                End Select
                para.Add(db.CreateParameter(String.Format("@{0}", colInfo(i).NAME), sqlDbType, row.Cells(i + 1).Value))
            Next
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))
            para.Add(db.CreateParameter("@一連番号", SqlDbType.VarChar, row.Cells(一連番号列).Value))
            para.Add(db.CreateParameter("@経営形態区分", SqlDbType.Int, row.Cells(経営形態区分列).Value))
            para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, row.Cells(営農類型区分列).Value))
            para.Add(db.CreateParameter("@生産費区分", SqlDbType.Int, row.Cells(生産費区分列).Value))
            para.Add(db.CreateParameter("@田畑区分", SqlDbType.Int, row.Cells(田畑区分列).Value))
            count += db.ExecuteNonQuery(sbSql.ToString, para)
        Next

        Return count
    End Function

End Class
