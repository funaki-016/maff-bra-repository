'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_000   | 2023.01.13 |大興電子通信        | 要件No.5 新規作成
'//            |            |                    |
'//*************************************************************************************************
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Text

''' <summary>
''' 抽出条件設定画面
''' </summary>
''' <remarks></remarks>
Public Class BRA10310F

    ''' <summary>
    ''' 標本の大きさ最大値
    ''' </summary>
    Private Const MAX_SIZE = 999999999

    ''' <summary>
    ''' 一覧列数(0～地方農政局列＋横計列＋規模階層区分列＋農政局列(非表示))
    ''' </summary>
    Private Shared ReadOnly COLUMN_SIZE As Integer = HyohonConst.規模階層区分リスト.Count + 1
    ''' <summary>
    ''' 一覧列番号
    ''' </summary>
    Private Enum DGV_COL
        地方農政局 = 0
        横計 = 1
        I = 2
        XIII = 14
        農政局 = 15
    End Enum
    ''' <summary>
    ''' 一覧行番号
    ''' </summary>
    Private Enum DGV_ROW
        上限 = 1
        下限 = 2
        縦計 = 3
        局別 = 4
    End Enum

    ''' <summary>
    ''' 経営形態区分（登録時のキー項目）
    ''' </summary>
    Private KeieiKeitai As Integer
    ''' <summary>
    ''' 営農類型区分（登録時のキー項目）
    ''' </summary>
    Private EinoRuikei As Integer
    ''' <summary>
    ''' 主副業区分（登録時のキー項目）
    ''' </summary>
    Private Shufukugyo As (選択肢 As Integer, 値リスト As List(Of Integer))
    ''' <summary>
    ''' 青色申告区分（登録時のキー項目）
    ''' </summary>
    Private Aoiro As Integer
    ''' <summary>
    ''' 集落営農区分（登録時のキー項目）
    ''' </summary>
    Private ShurakuEino As Integer
    ''' <summary>
    ''' 生産費区分（登録時のキー項目）
    ''' </summary>
    Private Seisanhi As Integer
    ''' <summary>
    ''' 田畑区分（登録時のキー項目）
    ''' </summary>
    Private Tahata As Integer

    ''' <summary>
    ''' 取得ID（更新時のキー項目）
    ''' </summary>
    Private 取得ID As Integer

    ''' <summary>
    ''' 初期表示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA10310F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'DGVちらつき防止
            dgvList.GetType().InvokeMember("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, Nothing, dgvList, New Object() {True})

            '経営形態区分プルダウン設定
            HyohonUtil.SetCombobox(CboKeieiKeitai, HyohonUtil.Get経営形態区分リスト())
            '営農類型区分プルダウン設定
            HyohonUtil.SetCombobox(CboEinoruikei, HyohonConst.営農類型区分リスト)
            '主副業区分プルダウン設定
            HyohonUtil.SetCombobox(CboShufukugyo, HyohonConst.主副業区分リスト)
            '青色申告区分プルダウン設定
            HyohonUtil.SetCombobox(CboAoiro, HyohonConst.青色申告区分リスト)
            '集落営農区分プルダウン設定
            HyohonUtil.SetCombobox(CboShurakuEino, HyohonConst.集落営農区分リスト)
            '生産費区分プルダウン設定
            HyohonUtil.SetCombobox(CboSeisanhi, HyohonUtil.Get生産費区分リスト(HyohonConst.経営形態区分.識別対象外))
            '田畑区分プルダウン設定
            HyohonUtil.SetCombobox(CboTahata, HyohonConst.田畑区分リスト)
            '経営形態区分の空を選択
            CboKeieiKeitai.SelectedValue = HyohonConst.経営形態区分.識別対象外

            '設定対象を空にする
            LblTarget.Text = ""

            '一覧の列設定
            Dim col As DataGridViewColumn
            For i = 0 To COLUMN_SIZE
                col = New DataGridViewTextBoxColumn
                If i = DGV_COL.地方農政局 Then
                    '1列目（地方農政局等列）
                    col.ReadOnly = True
                    col.Frozen = True
                    col.Width = 120
                    col.CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                ElseIf i = DGV_COL.横計 Then
                    '2列目（横計列）
                    col.ReadOnly = True
                    col.Frozen = True
                    col.Width = 80
                    col.CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                ElseIf i = DGV_COL.農政局 Then
                    col.ReadOnly = True
                    col.Visible = False
                Else
                    '3～15列目
                    col.ReadOnly = False
                    col.Width = 80
                    col.CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                End If
                dgvList.Columns.Add(col)
            Next

            'DGV共通設定
            ComUtil.ConfigDgvEditable(dgvList, 0)

            'セル選択で編集モードにする
            dgvList.EditMode = DataGridViewEditMode.EditOnEnter

            'クリア、削除ボタン非活性化
            BtnClear.Enabled = False
            BtnSave.Enabled = False

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
                    CboEinoruikei.SelectedValue = HyohonConst.営農類型区分.識別対象外
                    CboEinoruikei.Enabled = False
                    CboShufukugyo.SelectedValue = HyohonConst.主副業区分リスト(0).Key
                    CboShufukugyo.Enabled = False
                    CboSeisanhi.SelectedValue = HyohonConst.生産費区分.識別対象外
                    CboSeisanhi.Enabled = False
                Case HyohonConst.経営形態区分.個人経営体, HyohonConst.経営形態区分.法人経営体
                    '個人経営体、法人経営体
                    CboEinoruikei.Enabled = True
                    CboEinoruikei.SelectedValue = HyohonConst.営農類型区分.識別対象外
                    If keieiKeitai = HyohonConst.経営形態区分.個人経営体 Then
                        CboShufukugyo.Enabled = True
                    Else
                        CboShufukugyo.SelectedValue = HyohonConst.主副業区分リスト(0).Key
                        CboShufukugyo.Enabled = False
                    End If
                    CboSeisanhi.SelectedValue = HyohonConst.生産費区分.識別対象外
                    CboSeisanhi.Enabled = False
                Case Else
                    '個別経営体、組織法人経営体
                    CboEinoruikei.SelectedValue = HyohonConst.営農類型区分.識別対象外
                    CboEinoruikei.Enabled = False
                    CboShufukugyo.SelectedValue = HyohonConst.主副業区分リスト(0).Key
                    CboShufukugyo.Enabled = False
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
    ''' 営農類型区分プルダウン選択
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CboEinoruikei_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboEinoruikei.SelectedIndexChanged
        Try
            Dim keieiKeitai = DirectCast(HyohonUtil.GetComboValue(CboKeieiKeitai), HyohonConst.経営形態区分)
            Dim einouruikei = DirectCast(HyohonUtil.GetComboValue(CboEinoruikei), HyohonConst.営農類型区分)
            If keieiKeitai = HyohonConst.経営形態区分.法人経営体 _
                AndAlso einouruikei = HyohonConst.営農類型区分.水田作 Then
                '法人経営体かつ水田作
                CboShurakuEino.Enabled = True
            Else
                '法人経営体かつ水田作以外
                CboShurakuEino.SelectedValue = HyohonConst.集落営農区分.識別対象外
                CboShurakuEino.Enabled = False
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 主副業区分プルダウン選択
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CboShufukugyo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboShufukugyo.SelectedIndexChanged
        Try
            Dim shufukugyo = HyohonUtil.GetComboValue(CboShufukugyo, HyohonConst.主副業区分リスト(0).Key)
            If shufukugyo.選択肢 = HyohonConst.主副業区分選択肢.副業 Then
                '副業
                CboAoiro.Enabled = True
            Else
                '副業以外
                CboAoiro.SelectedValue = HyohonConst.青色申告区分.識別対象外
                CboAoiro.Enabled = False
            End If

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
                AndAlso (seisanhi = HyohonConst.生産費区分.小麦 _
                OrElse seisanhi = HyohonConst.生産費区分.大豆) Then
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
    ''' 表示ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnShow_Click(sender As Object, e As EventArgs) Handles BtnShow.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            '入力チェック
            '経営形態区分（必須）
            Dim cbo = CboKeieiKeitai
            If HyohonUtil.IsEmpty(cbo) Then
                cbo.Focus()
                Message.ShowMsgBox(MessageID.MSG_E_119, MsgBoxStyle.OkOnly)
                Return
            End If
            '営農類型区分（活性の場合、必須）
            cbo = CboEinoruikei
            If cbo.Enabled AndAlso HyohonUtil.IsEmpty(cbo) Then
                cbo.Focus()
                Message.ShowMsgBox(MessageID.MSG_E_120, MsgBoxStyle.OkOnly)
                Return
            End If
            '主副業区分（活性の場合、必須）
            cbo = CboShufukugyo
            If cbo.Enabled AndAlso HyohonUtil.IsEmpty(cbo, HyohonConst.主副業区分リスト(0).Key) Then
                cbo.Focus()
                Message.ShowMsgBox(MessageID.MSG_E_131, MsgBoxStyle.OkOnly)
                Return
            End If
            '青色申告区分（活性の場合、必須）
            cbo = CboAoiro
            If cbo.Enabled AndAlso HyohonUtil.IsEmpty(cbo) Then
                cbo.Focus()
                Message.ShowMsgBox(MessageID.MSG_E_132, MsgBoxStyle.OkOnly)
                Return
            End If
            '集落営農区分（活性の場合、必須）
            cbo = CboShurakuEino
            If cbo.Enabled AndAlso HyohonUtil.IsEmpty(cbo) Then
                cbo.Focus()
                Message.ShowMsgBox(MessageID.MSG_E_133, MsgBoxStyle.OkOnly)
                Return
            End If
            '生産費区分（活性の場合、必須）
            cbo = CboSeisanhi
            If cbo.Enabled AndAlso HyohonUtil.IsEmpty(cbo) Then
                cbo.Focus()
                Message.ShowMsgBox(MessageID.MSG_E_121, MsgBoxStyle.OkOnly)
                Return
            End If
            '田畑区分（活性の場合、要否確認）
            cbo = CboTahata
            If cbo.Enabled AndAlso HyohonUtil.IsEmpty(cbo) Then
                If Message.ShowMsgBox(MessageID.MSG_Q_057, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    cbo.Focus()
                    Return
                End If
            End If

            'キー項目を退避
            KeieiKeitai = HyohonUtil.GetComboValue(CboKeieiKeitai)
            EinoRuikei = HyohonUtil.GetComboValue(CboEinoruikei)
            Shufukugyo = HyohonUtil.GetComboValue(CboShufukugyo, HyohonConst.主副業区分リスト(0).Key)
            Aoiro = HyohonUtil.GetComboValue(CboAoiro)
            ShurakuEino = HyohonUtil.GetComboValue(CboShurakuEino)
            Seisanhi = HyohonUtil.GetComboValue(CboSeisanhi)
            Tahata = HyohonUtil.GetComboValue(CboTahata)

            '設定対象に入力項目名を設定
            Dim target = HyohonConst.経営形態区分辞書(DirectCast(KeieiKeitai, HyohonConst.経営形態区分)).名称
            If EinoRuikei <> HyohonConst.識別対象外値 Then
                target &= "　" & HyohonUtil.Getリスト名称(HyohonConst.営農類型区分リスト, EinoRuikei)
            End If
            If Shufukugyo.選択肢 <> HyohonConst.識別対象外値 Then
                target &= "　" & HyohonUtil.Getリスト名称(HyohonConst.主副業区分リスト, Shufukugyo)
            End If
            If Aoiro <> HyohonConst.識別対象外値 Then
                target &= "　" & HyohonUtil.Getリスト名称(HyohonConst.青色申告区分リスト, Aoiro)
            End If
            If ShurakuEino <> HyohonConst.識別対象外値 Then
                target &= "　" & HyohonUtil.Getリスト名称(HyohonConst.集落営農区分リスト, ShurakuEino)
            End If
            If Seisanhi <> HyohonConst.識別対象外値 Then
                target &= "　" & HyohonUtil.Getリスト名称(HyohonUtil.Get生産費区分リスト(DirectCast(KeieiKeitai, HyohonConst.経営形態区分)), Seisanhi)
            End If
            If Tahata <> HyohonConst.識別対象外値 Then
                target &= "　" & HyohonUtil.Getリスト名称(HyohonConst.田畑区分リスト, Tahata)
            End If
            LblTarget.Text = target

            '一覧クリア
            dgvList.Rows.Clear()

            '１行目を追加
            HyohonUtil.AddDgvRow(dgvList, {"", "", "Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ", "Ⅺ", "Ⅻ", "XIII", ""}).ReadOnly = True

            Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '標本リスト設定テーブル検索
                取得ID = HyohonUtil.Get標本リスト設定ID(db, KeieiKeitai, EinoRuikei, Shufukugyo.選択肢, Aoiro, ShurakuEino, Seisanhi, Tahata)

                '標本リスト設定＿抽出条件テーブル検索
                Dim dt = HyohonUtil.Select標本リスト設定＿抽出条件(db, 取得ID)
                '2行目
                Dim row2 = New List(Of Object)
                row2.Add("規模番号")
                row2.Add("上限（未満）")
                '3行目
                Dim row3 = New List(Of Object)
                row3.Add("")
                row3.Add("下限（以上）")
                Dim decVal As Decimal
                For Each row As DataRow In dt.Rows
                    If Decimal.TryParse(If(row("上限"), "").ToString(), decVal) Then
                        row2.Add(decVal.ToString("#,0.0"))
                    Else
                        row2.Add("")
                    End If
                    If Decimal.TryParse(If(row("下限"), "").ToString(), decVal) Then
                        row3.Add(decVal.ToString("#,0.0"))
                    Else
                        row3.Add("")
                    End If
                Next
                '2行目を追加
                HyohonUtil.AddDgvRow(dgvList, row2.ToArray())
                '3行目を追加
                HyohonUtil.AddDgvRow(dgvList, row3.ToArray())

                '地方農政局マスタ、標本リスト設定＿期待値テーブル検索
                dt = HyohonUtil.Select標本リスト設定＿期待値(db, 取得ID)
                '農政局でグループ化する
                Dim grouped = dt.Rows.Cast(Of DataRow).GroupBy(Function(x) CInt(x("農政局"))).ToList()
                '4行目以降
                Dim rowx As List(Of Object)
                Dim intVal As Integer
                Dim colIndex As Integer
                For Each group In grouped
                    colIndex = 0
                    rowx = New List(Of Object)
                    For i = 0 To COLUMN_SIZE
                        If i = DGV_COL.地方農政局 Then
                            rowx.Add(group(0)("名称"))
                        ElseIf i = DGV_COL.横計 Then
                            rowx.Add(CInt(group(0)("横計")).ToString("#,#"))
                        ElseIf i = DGV_COL.農政局 Then
                            rowx.Add(group(0)("農政局"))
                        Else
                            If Integer.TryParse(If(group(colIndex)("標本サイズ"), "").ToString(), intVal) Then
                                rowx.Add(intVal.ToString("#,#"))
                            Else
                                rowx.Add("")
                            End If
                            colIndex += 1
                        End If
                    Next
                    If group.Key = HyohonConst.農政局＿全国 Then
                        HyohonUtil.AddDgvRow(dgvList, rowx.ToArray()).ReadOnly = True
                    Else
                        HyohonUtil.AddDgvRow(dgvList, rowx.ToArray())
                    End If
                Next
            End Using

            dgvList.ClearSelection()
            'クリア、削除ボタン活性化
            BtnClear.Enabled = True
            BtnSave.Enabled = True

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 一覧セル描画
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DgvList_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles dgvList.CellPainting
        Try
            If e.RowIndex < 0 Then
                '負数の場合は何もしない
                Return
            End If

            Dim cell = dgvList(e.ColumnIndex, e.RowIndex)
            If e.RowIndex = 0 Then
                '1行目
                cell.ReadOnly = True
                cell.Style.BackColor = Color.Silver
                cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                If e.ColumnIndex = 0 Then
                    '1列目
                    e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None
                End If
            ElseIf e.RowIndex = 1 Then
                '2行目
                If e.ColumnIndex <= 1 Then
                    '1,2列目
                    cell.Style.BackColor = Color.Silver
                    If e.ColumnIndex = 0 Then
                        '1列目
                        e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None
                        e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None
                    Else
                        '2列目
                        cell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    End If
                End If
            ElseIf e.RowIndex = 2 Then
                '3行目
                If e.ColumnIndex <= 1 Then
                    '1,2列目
                    cell.Style.BackColor = Color.Silver
                    If e.ColumnIndex = 0 Then
                        '1列目
                        e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None
                    Else
                        '2列目
                        cell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    End If
                End If
            ElseIf e.RowIndex = 3 Then
                '4行目
                cell.ReadOnly = True
                cell.Style.BackColor = Color.WhiteSmoke
            Else
                '5行目以降
                If e.ColumnIndex <= 1 Then
                    '1,2列目
                    cell.Style.BackColor = Color.WhiteSmoke
                End If
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 一覧セル入力開始
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DgvList_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvList.CellBeginEdit
        Try
            Dim tbc = DirectCast(dgvList.CurrentCell, DataGridViewTextBoxCell)
            'カンマ除去
            tbc.Value = tbc.Value?.ToString()?.Replace(",", "")
            '最大桁数設定
            If e.RowIndex <= DGV_ROW.下限 Then
                '上限、下限
                tbc.MaxInputLength = HyohonConst.標本リスト設定＿抽出上限値.ToString().Length + 1
            Else
                '標本の大きさ
                tbc.MaxInputLength = MAX_SIZE.ToString().Length
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
            Dim cell = dgvList(e.ColumnIndex, e.RowIndex)
            If cell.ReadOnly Then
                '入力不可は何もしない
                Return
            End If
            If e.RowIndex <= DGV_ROW.下限 Then
                '上限、下限
                Dim strVal = If(cell.Value?.ToString(), "")
                If String.IsNullOrEmpty(strVal) Then
                    Return
                End If
                Dim val As Decimal
                If Not Decimal.TryParse(strVal, val) Then
                    '数値以外は0にする
                    val = 0
                End If
                If val > HyohonConst.標本リスト設定＿抽出上限値 Then
                    '最大値
                    val = HyohonConst.標本リスト設定＿抽出上限値
                ElseIf val < -HyohonConst.標本リスト設定＿抽出上限値 Then
                    '最大値の負数
                    val = -HyohonConst.標本リスト設定＿抽出上限値
                End If
                'カンマ編集
                cell.Value = val.ToString("#,0.0")
            Else
                '標本の大きさ
                Dim val As Integer
                If Not Integer.TryParse(cell.Value?.ToString(), val) OrElse val < 0 Then
                    '整数以外、負数は0にする
                    val = 0
                End If
                If val > MAX_SIZE Then
                    '最大値
                    val = MAX_SIZE
                End If
                '0は空欄にする
                cell.Value = If(val = 0, "", val.ToString("#,0"))
                '横計
                dgvList(1, e.RowIndex).Value = dgvList.Rows(e.RowIndex).Cells().Cast(Of DataGridViewCell) _
                    .Where(Function(x) 2 <= x.ColumnIndex AndAlso x.ColumnIndex <= 14) _
                    .Where(Function(x) Decimal.TryParse(x.Value?.ToString(), Nothing)) _
                    .Select(Function(x) CDec(x.Value)) _
                    .Sum().ToString("#,#")
                '縦計
                dgvList(e.ColumnIndex, 3).Value = dgvList.Rows.Cast(Of DataGridViewRow) _
                    .Select(Function(x) x.Cells(e.ColumnIndex)) _
                    .Where(Function(x) 4 <= x.RowIndex) _
                    .Where(Function(x) Decimal.TryParse(x.Value?.ToString(), Nothing)) _
                    .Select(Function(x) CDec(x.Value)) _
                    .Sum().ToString("#,#")
                '総計
                dgvList(1, 3).Value = dgvList.Rows.Cast(Of DataGridViewRow) _
                    .Select(Function(x) x.Cells(1)) _
                    .Where(Function(x) 4 <= x.RowIndex) _
                    .Where(Function(x) Decimal.TryParse(x.Value?.ToString(), Nothing)) _
                    .Select(Function(x) CDec(x.Value)) _
                    .Sum().ToString("#,#")
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
            If Message.ShowMsgBox(MessageID.MSG_Q_062, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If

            '入力セル(not readonly)
            Dim cells = dgvList.Rows.Cast(Of DataGridViewRow).SelectMany(Function(x) x.Cells().Cast(Of DataGridViewCell) _
                .Where(Function(y) Not y.ReadOnly)).ToList()
            '縦計セル
            cells.AddRange(dgvList.Rows(DGV_ROW.縦計).Cells().Cast(Of DataGridViewCell).Where(Function(x) DGV_COL.横計 <= x.ColumnIndex AndAlso x.ColumnIndex <= DGV_COL.XIII))
            '横計セル
            cells.AddRange(dgvList.Rows().Cast(Of DataGridViewRow).SelectMany(Function(x) x.Cells().Cast(Of DataGridViewCell) _
                .Where(Function(y) y.RowIndex >= DGV_ROW.局別 AndAlso y.ColumnIndex = DGV_COL.横計)))
            For Each cell In cells
                cell.Value = ""
            Next

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

            If Message.ShowMsgBox(MessageID.MSG_Q_001, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If

            Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try
                    db.BeginTrans()

                    Dim sbSql As StringBuilder
                    Dim para As List(Of DBAccess.Parameter)

                    If 取得ID = 0 Then
                        'ID登録
                        Try
                            '標本リスト設定テーブル登録
                            sbSql = New StringBuilder()
                            para = New List(Of DBAccess.Parameter)
                            sbSql.Append("INSERT INTO 標本リスト設定 VALUES(")
                            sbSql.Append(" @経営形態区分,")
                            para.Add(db.CreateParameter("@経営形態区分", SqlDbType.Int, KeieiKeitai))
                            sbSql.Append(" @営農類型区分")
                            para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, EinoRuikei))
                            sbSql.Append(" @主副業区分")
                            para.Add(db.CreateParameter("@主副業区分", SqlDbType.Int, Shufukugyo))
                            sbSql.Append(" @青色申告区分")
                            para.Add(db.CreateParameter("@青色申告区分", SqlDbType.Int, Aoiro))
                            sbSql.Append(" @集落営農区分")
                            para.Add(db.CreateParameter("@集落営農区分", SqlDbType.Int, ShurakuEino))
                            sbSql.Append(" @生産費区分")
                            para.Add(db.CreateParameter("@生産費区分", SqlDbType.Int, Seisanhi))
                            sbSql.Append(" @田畑区分")
                            para.Add(db.CreateParameter("@田畑区分", SqlDbType.Int, Tahata))
                            sbSql.Append(" SYSDATETIME(),")
                            sbSql.Append(" @更新者ID)")
                            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))
                            db.ExecuteNonQuery(sbSql.ToString(), para)
                        Catch se As SqlException
                            If se.Number <> 2627 Then
                                '一意制約エラー以外はスロー
                                Throw se
                            End If
                            '一意制約エラーの場合はID取得
                            取得ID = HyohonUtil.Get標本リスト設定ID(db, KeieiKeitai, EinoRuikei, Shufukugyo.選択肢, Aoiro, ShurakuEino, Seisanhi, Tahata)
                        End Try
                    End If

                    If 取得ID <> 0 Then
                        '更新処理
                        '標本リスト設定テーブル更新
                        sbSql = New StringBuilder()
                        para = New List(Of DBAccess.Parameter)
                        sbSql.Append("UPDATE 標本リスト設定 SET")
                        sbSql.Append(" 更新日付 = SYSDATETIME(),")
                        sbSql.Append(" 更新者ID = @更新者ID")
                        para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))
                        sbSql.Append(" WHERE ID = @ID")
                        para.Add(db.CreateParameter("@ID", SqlDbType.Int, 取得ID))
                        db.ExecuteNonQuery(sbSql.ToString(), para)
                    Else
                        '登録したIDを取得
                        取得ID = HyohonUtil.Get標本リスト設定ID(db, KeieiKeitai, EinoRuikei, Shufukugyo.選択肢, Aoiro, ShurakuEino, Seisanhi, Tahata)
                    End If

                    '既存削除
                    '標本リスト設定＿抽出条件テーブル削除
                    sbSql = New StringBuilder()
                    para = New List(Of DBAccess.Parameter)
                    sbSql.Append("DELETE FROM 標本リスト設定＿抽出条件")
                    sbSql.Append(" WHERE ID = @ID")
                    para.Add(db.CreateParameter("@ID", SqlDbType.Int, 取得ID))
                    db.ExecuteNonQuery(sbSql.ToString(), para)

                    '標本リスト設定＿期待値テーブル削除
                    sbSql = New StringBuilder()
                    sbSql.Append("DELETE FROM 標本リスト設定＿期待値")
                    sbSql.Append(" WHERE ID = @ID")
                    para = New List(Of DBAccess.Parameter)
                    para.Add(db.CreateParameter("@ID", SqlDbType.Int, 取得ID))
                    db.ExecuteNonQuery(sbSql.ToString(), para)

                    '登録処理
                    '標本リスト設定＿抽出条件テーブル登録
                    sbSql = New StringBuilder()
                    sbSql.Append("INSERT INTO 標本リスト設定＿抽出条件 VALUES(")
                    sbSql.Append(" @ID,")
                    sbSql.Append(" @規模番号,")
                    sbSql.Append(" @上限,")
                    sbSql.Append(" @下限)")
                    Dim uppers = dgvList.Rows(DGV_ROW.上限).Cells().Cast(Of DataGridViewCell).OrderBy(Function(x) x.ColumnIndex).Select(Function(x) If(x.Value?.ToString(), "")).ToList()
                    Dim lowers = dgvList.Rows(DGV_ROW.下限).Cells().Cast(Of DataGridViewCell).OrderBy(Function(x) x.ColumnIndex).Select(Function(x) If(x.Value?.ToString(), "")).ToList()
                    Dim kibo = 0
                    For i = DGV_COL.I To DGV_COL.XIII
                        kibo += 1
                        para = New List(Of DBAccess.Parameter)
                        para.Add(db.CreateParameter("@ID", SqlDbType.Int, 取得ID))
                        para.Add(db.CreateParameter("@規模番号", SqlDbType.Int, kibo))
                        para.Add(db.CreateParameter("@上限", SqlDbType.Decimal, uppers(i)))
                        para.Add(db.CreateParameter("@下限", SqlDbType.Decimal, lowers(i)))
                        db.ExecuteNonQuery(sbSql.ToString(), para)
                    Next

                    '標本リスト設定＿期待値テーブル登録
                    sbSql = New StringBuilder()
                    sbSql.Append("INSERT INTO 標本リスト設定＿期待値 VALUES(")
                    sbSql.Append(" @ID,")
                    sbSql.Append(" @農政局,")
                    sbSql.Append(" @規模番号,")
                    sbSql.Append(" @標本サイズ)")
                    Dim rows = dgvList.Rows().Cast(Of DataGridViewRow).Where(Function(x) x.Cells(0).RowIndex >= DGV_ROW.縦計).OrderBy(Function(x) x.Cells(0).RowIndex).ToList()
                    For Each row In rows
                        kibo = 0
                        For i = DGV_COL.I To DGV_COL.XIII
                            kibo += 1
                            para = New List(Of DBAccess.Parameter)
                            para.Add(db.CreateParameter("@ID", SqlDbType.Int, 取得ID))
                            para.Add(db.CreateParameter("@農政局", SqlDbType.Int, row.Cells(DGV_COL.農政局).Value))
                            para.Add(db.CreateParameter("@規模番号", SqlDbType.Int, kibo))
                            para.Add(db.CreateParameter("@標本サイズ", SqlDbType.Int, If(row.Cells(i).Value?.ToString().Replace(",", ""), "")))
                            db.ExecuteNonQuery(sbSql.ToString(), para)
                        Next
                    Next

                    db.CommitTrans()
                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using

            '一覧クリア
            LblTarget.Text = ""
            dgvList.Rows.Clear()
            'ボタンを非活性化
            BtnClear.Enabled = False
            BtnSave.Enabled = False

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

End Class
