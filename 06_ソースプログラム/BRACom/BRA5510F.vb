Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Text.RegularExpressions

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.12.15 |Daiko               | 要件No4 バージョン区分追加
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 任意帳票出力画面
''' </summary>
''' <remarks></remarks>
Public Class BRA5510F

    ''' <summary>調査年</summary>
    Private _chosaNen As String

    ''' <summary>集計キー（本年）</summary>
    Private _skeyThis As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey)
    ''' <summary>集計キー（比較）</summary>
    Private _sKeyComp As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey)

    ''' <summary>営農経営体区分</summary>
    Private _einouKeieitai As String

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA5510F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '営農経営体区分コンボボックス設定
            ComUtil.SetEinouKeieitaiComboBox(lblEinouKeieitai, cboEinouKeieitai)

            '調査区分取得
            Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

            If Not (CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計) Then
                '調査年コンボボックス設定
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    ComUtil.SyukeiKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, chosakubun, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
                End Using
            End If

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

            cboEinouKeieitai.Enabled = False

            _chosaNen = cboChosaNen.SelectedValue.ToString
            _einouKeieitai = If(cboEinouKeieitai.SelectedValue Is Nothing, Nothing, cboEinouKeieitai.SelectedValue.ToString)

            '一覧表示
            Me.ShowList(_chosaNen, _einouKeieitai, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 出力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Dim parametersPath As String = txtParameters.Text
            Dim layoutPath As String = txtLayout.Text

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(_skeyThis, parametersPath, layoutPath, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim parameters As Dictionary(Of String, ComUtil.NiniChohyo.編成パラメータ)
            Dim sheetName As List(Of String)
            '任意帳票入力ファイルクラス生成
            Using ExceNiniChohyoInFile = New NiniChohyoInFile()
                'パラメータ取得処理実行
                parameters = ExceNiniChohyoInFile.ExecuteGetParameter(parametersPath)
                sheetName = ExceNiniChohyoInFile.ExecuteGetSheetName(layoutPath)
            End Using

            Dim chosakubun As String = ComUtil.GetChosakubun(_einouKeieitai)

            'エラーチェック
            Dim details As New List(Of String)
            ' REV_001↓
            'If Not Me.CheckError(parameters, sheetName, _sKeyComp, chosakubun, details) Then
            If Not Me.CheckError(parameters, sheetName, _sKeyComp, chosakubun, _skeyThis.Item1.chosaNen, details) Then
                ' REV_001↑
                'エラーメッセージ
                Message.ShowMsgForm(Me, MessageID.MSG_E_010, {String.Join(vbCrLf, details)})
                Exit Sub
            End If

            Dim fileName As String = ComConst.任意帳票.出力用ファイル名称.reportName & ".xlsx"

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Dim replace As Boolean = chkReplace.Checked

            Try
                Dim pgdCount As Integer
                For Each kv As KeyValuePair(Of String, ComUtil.NiniChohyo.編成パラメータ) In parameters
                    pgdCount = pgdCount + kv.Value.指標部.Count
                Next
                pgdCount = pgdCount + parameters.Count

                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New NiniChohyoOut(filePath, _skeyThis, _sKeyComp, parameters, replace, chosakubun, layoutPath)
                    ret = ExcelOutput.Execute(Me, pgdCount, MessageID.MSG_Q_004)
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

    Private Sub btnSetThis_Click(sender As Object, e As EventArgs) Handles btnSetThis.Click
        Try
            _skeyThis = Nothing

            '集計結果表主キー取得
            _skeyThis = Me.GetSyukeiKekkahyoSelectKey(_chosaNen)

            'エラーチェック
            If _skeyThis.Item1 Is Nothing Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_030, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim chosakubun As String = ComUtil.GetChosakubun(_einouKeieitai)
            Dim syukeiName As String
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '集計名称取得
                syukeiName = DAOSyukeiKekkahyo.GetSyukeiName(db, chosakubun, _skeyThis.Item1, _skeyThis.Item2)
            End Using

            txtYearThis.Text = _skeyThis.Item1.chosaNen
            txtSyukeiNoThis.Text = _skeyThis.Item1.syukeiNo
            txtSyukeiNameThis.Text = syukeiName
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnSetComp_Click(sender As Object, e As EventArgs) Handles btnSetComp.Click
        Try
            _sKeyComp = Nothing

            '集計結果表主キー取得
            _sKeyComp = Me.GetSyukeiKekkahyoSelectKey(_chosaNen)

            'エラーチェック
            If _sKeyComp.Item1 Is Nothing Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_030, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim chosakubun As String = ComUtil.GetChosakubun(_einouKeieitai)
            Dim syukeiName As String
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '集計名称取得
                syukeiName = DAOSyukeiKekkahyo.GetSyukeiName(db, chosakubun, _sKeyComp.Item1, _sKeyComp.Item2)
            End Using

            txtYearComp.Text = _sKeyComp.Item1.chosaNen
            txtSyukeiNoComp.Text = _sKeyComp.Item1.syukeiNo
            txtSyukeiNameComp.Text = syukeiName
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnParameters_Click(sender As Object, e As EventArgs) Handles btnParameters.Click
        Try
            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of OpenFileDialog)(Me, IniFileInfo.ExcelInPath)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            txtParameters.Text = filePath
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnLayout_Click(sender As Object, e As EventArgs) Handles btnLayout.Click
        Try
            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of OpenFileDialog)(Me, IniFileInfo.ExcelInPath)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            txtLayout.Text = filePath
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub dgvList_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvList.CellValueChanged
        If e.ColumnIndex = 0 Then
            If CType(dgvList(e.ColumnIndex, e.RowIndex).Value, Boolean) = True Then
                For rowIndex = 0 To dgvList.Rows.Count - 1
                    If rowIndex <> e.RowIndex Then
                        dgvList(0, rowIndex).Value = False
                        dgvList(0, rowIndex).ReadOnly = False
                    End If
                Next
                dgvList(e.ColumnIndex, e.RowIndex).ReadOnly = True
            End If
        End If
    End Sub

    Private Sub dgvList_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvList.CurrentCellDirtyStateChanged
        If dgvList.CurrentCellAddress.X = 0 AndAlso dgvList.IsCurrentCellDirty Then
            dgvList.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub cboEinouKeieitai_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboEinouKeieitai.SelectedIndexChanged
        dgvList.Rows.Clear()

        '調査区分取得
        Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

        '調査年コンボボックス設定
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            ComUtil.SyukeiKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, chosakubun, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
        End Using
    End Sub

    ''' <summary>
    ''' 集計結果表主キー取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSyukeiKekkahyoSelectKey(chosaNen As String) As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey)
        Dim ret As New ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey)

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                Dim pkey As DAOSyukeiKekkahyo.PrimaryKey = New DAOSyukeiKekkahyo.PrimaryKey(chosaNen, dgvList.Rows(i).Cells(1).Value.ToString)
                Dim kkey As DAOSyukeiKekkahyo.KomokuKey = New DAOSyukeiKekkahyo.KomokuKey(dgvList.Rows(i).Cells(5).Value.ToString, dgvList.Rows(i).Cells(6).Value.ToString, dgvList.Rows(i).Cells(7).Value.ToString)
                ret = ValueTuple.Create(pkey, kkey)

                Exit For
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <param name="chosakubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <remarks></remarks>
    Public Sub ShowList(chosaNen As String, einouKeieitai As String, kyoku As String, jimusho As String, kyoten As String)
        Dim dt As DataTable = Nothing

        Dim chosakubun As String = ComUtil.GetChosakubun(einouKeieitai)

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Select Case CommonInfo.Koutei
                Case CommonInfo.KouteiKubun.Code.Center
                    dt = DAOSyukeiKekkahyo.GetList(db, chosakubun, chosaNen, kyoku, jimusho, kyoten)
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    dt = DAOSyukeiKekkahyo.GetList(db, chosakubun, chosaNen, kyoku, Nothing, Nothing)
                Case CommonInfo.KouteiKubun.Code.Honsyo
                    dt = DAOSyukeiKekkahyo.GetList(db, chosakubun, chosaNen, Nothing, Nothing, Nothing)
            End Select
        End Using

        dgvList.Rows.Clear()

        For Each row As DataRow In dt.Rows
            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(1).Value = row("集計番号").ToString
            dgvList.Rows(i).Cells(2).Value = row("集計名称").ToString
            dgvList.Rows(i).Cells(3).Value = DateTime.Parse(row("更新日付").ToString).ToString(ComConst.DATETIME_FORMAT)
            dgvList.Rows(i).Cells(4).Value = row("集計条件").ToString
            dgvList.Rows(i).Cells(5).Value = row("農政局").ToString
            dgvList.Rows(i).Cells(6).Value = row("都道府県").ToString
            dgvList.Rows(i).Cells(7).Value = row("実査設置拠点").ToString

            dgvList.Rows(i).Cells(4).Style.Alignment = DataGridViewContentAlignment.MiddleLeft
        Next
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="sKey"></param>
    ''' <param name="parameters"></param>
    ''' <param name="layout"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(sKey As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey), parameters As String, layout As String, ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        '集計番号設定チェック
        If sKey.Item1 Is Nothing Then
            msgId = MessageID.MSG_E_042
            Return ret
        End If

        '再編成パラメータチェック
        If parameters.Equals(String.Empty) Then
            msgId = MessageID.MSG_E_043
            Return ret
        End If

        '帳票レイアウトチェック
        If layout.Equals(String.Empty) Then
            msgId = MessageID.MSG_E_044
            Return ret
        End If

        ret = True

        Return ret
    End Function

    ' REV_001↓
    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="parameters"></param>
    ''' <param name="sheetName"></param>
    ''' <param name="sKeyComp"></param>
    ''' <param name="chosaKubun"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="details"></param>
    ''' <returns></returns>
    'Private Function CheckError(parameters As Dictionary(Of String, ComUtil.NiniChohyo.編成パラメータ), sheetName As List(Of String), sKeyComp As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey), chosaKubun As String, ByRef details As List(Of String)) As Boolean
    Private Function CheckError(parameters As Dictionary(Of String, ComUtil.NiniChohyo.編成パラメータ), sheetName As List(Of String), sKeyComp As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey), chosaKubun As String, chosaNen As String, ByRef details As List(Of String)) As Boolean
        ' REV_001↑
        Dim ret As Boolean = True
        Dim dtItem As DataTable

        Const max As Integer = ComConst.ERR_MESSAGE_MAX

        Dim msg As String() = {"" _
                             , "{0}件目：再編成パラメータのタイトルが不正です。" _
                             , "{0}件目：シート名（{1}）で指定したシートが帳票レイアウトに存在しません。" _
                             , "{0}件目：シート名（{1}）で指定した開始位置セルはA1形式で入力してください。" _
                             , "{0}件目：シート名（{1}）で指定した開始位置セルは1000行、1000列以内で入力してください。" _
                             , "{0}件目：シート名（{1}）で指定した生産費平均値種類指定は整数一桁で入力してください。" _
                             , "{0}件目：シート名（{1}）で指定した田畑区分指定は整数一桁で入力してください。" _
                             , "{0}件目：シート名（{1}）で指定したビール麦販売区分指定は整数一桁で入力してください。" _
                             , "{0}件目：シート名（{1}）で指定したてんさい栽培区分指定は整数一桁で入力してください。" _
                             , "{0}件目：シート名（{1}）で指定した表頭・表側指定は「表頭」または「表側」のどちらかを入力してください。" _
                             , "{0}件目：シート名（{1}）行数（{2}）に存在しない項番が入力されております。 " _
                             , "{0}件目：シート名（{1}）行数（{2}）の項目番号で「前」を指定した場合、画面上で比較の集計番号を設定してください。" _
                             , "{0}件目：シート名（{1}）行数（{2}）の地域は整数で入力してください。" _
                             , "{0}件目：シート名（{1}）行数（{2}）の規模階層は整数で入力してください。" _
                             , "{0}件目：再編成パラメータは20シート以内で設定してください。"
        }

        Dim cnt As Integer = 0

        Dim row As Integer = 0

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '集計結果表項目マスタ取得
            ' REV_001↓
            'dtItem = DAOOther.GetSyukeiItemMaster(db, chosaKubun)
            dtItem = DAOOther.GetSyukeiItemMaster(db, chosaKubun, ComUtil.getVersionKubunTaikei(chosaNen, chosaKubun))
            ' REV_001↑
        End Using

        For Each kv As KeyValuePair(Of String, ComUtil.NiniChohyo.編成パラメータ) In parameters

            '1）再編成パラメータのタイトル（A1セル）が「任意帳票出力再編成パラメータ」であること。
            If Not kv.Value.タイトル.Equals(ComConst.任意帳票.TITLE) Then
                cnt = cnt + 1
                details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2)))
                ret = False
                If cnt = max Then Return ret
            End If

            '2）再編成パラメータの帳票レイアウト（B4セル）が、画面で設定した帳票レイアウトエクセルに存在するシートであること。
            If Not sheetName.Contains(kv.Value.帳票レイアウト) Then
                cnt = cnt + 1
                details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), kv.Key))
                ret = False
                If cnt = max Then Return ret
            End If

            '3）再編成パラメータの開始位置セルを指定（B7セル）が、EXCELで設定できるセル位置（A1形式）であること。
            If Not Regex.IsMatch(kv.Value.開始位置セル, "^[A-Z]+[0-9]+$") Then
                cnt = cnt + 1
                details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), kv.Key))
                ret = False
                If cnt = max Then Return ret
            End If

            '4）再編成パラメータの開始位置セルを指定（B7セル）が、A1～ALL1000の範囲以内に収まること。
            If Regex.IsMatch(kv.Value.開始位置セル, "^[A-Z]+[0-9]+$") Then
                Dim val1 As Integer = ConvertLetterToColNo(Regex.Replace(kv.Value.開始位置セル, "[0-9]", String.Empty))
                Dim val2 As Integer = Integer.Parse(Regex.Replace(kv.Value.開始位置セル, "[A-Z]", String.Empty))

                If Not (val1 >= 1 And val1 <= 1000) Or Not (val2 >= 1 And val2 <= 1000) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), kv.Key))
                    ret = False
                    If cnt = max Then Return ret
                End If
            End If

            '5）再編成パラメータの生産費平均値種類指定（B10セル）が、整数一桁であるか。
            If Not Regex.IsMatch(kv.Value.生産費平均値種類, "^[0-9]+$") Then
                cnt = cnt + 1
                details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), kv.Key))
                ret = False
                If cnt = max Then Return ret
            Else
                Dim val As Decimal
                If Decimal.TryParse(kv.Value.生産費平均値種類, val) Then
                    If Not (val >= 0D And val <= 9D) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), kv.Key))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
            End If

            '6）再編成パラメータの田畑区分指定（B13セル）が、整数一桁であるか。
            If Not Regex.IsMatch(kv.Value.田畑区分指定, "^[0-9]+$") Then
                cnt = cnt + 1
                details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), kv.Key))
                ret = False
                If cnt = max Then Return ret
            Else
                Dim val As Decimal
                If Decimal.TryParse(kv.Value.田畑区分指定, val) Then
                    If Not (val >= 0D And val <= 9D) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), kv.Key))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
            End If

            '7）再編成パラメータのビール麦販売区分指定（B16セル）が、整数一桁であるか。
            If Not Regex.IsMatch(kv.Value.ビール麦販売区分, "^[0-9]+$") Then
                cnt = cnt + 1
                details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), kv.Key))
                ret = False
                If cnt = max Then Return ret
            Else
                Dim val As Decimal
                If Decimal.TryParse(kv.Value.ビール麦販売区分, val) Then
                    If Not (val >= 0D And val <= 9D) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), kv.Key))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
            End If

            '8）再編成パラメータのてんさい栽培区分指定（B19セル）が、整数一桁であるか。
            If Not Regex.IsMatch(kv.Value.てんさい栽培区分, "^[0-9]+$") Then
                cnt = cnt + 1
                details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2), kv.Key))
                ret = False
                If cnt = max Then Return ret
            Else
                Dim val As Decimal
                If Decimal.TryParse(kv.Value.てんさい栽培区分, val) Then
                    If Not (val >= 0D And val <= 9D) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2), kv.Key))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
            End If

            '9）再編成パラメータの表頭・表側指定（B22セル）が、「表頭」または「表側」であるか。
            If Not ComConst.任意帳票.表頭_表側.リスト.ContainsValue(kv.Value.表頭_表側) Then
                cnt = cnt + 1
                details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2), kv.Key))
                ret = False
                If cnt = max Then Return ret
            End If

            '10）再編成パラメータの項目番号（C27～C226セル）が、存在する項目番号か、存在する項目番号に「前」が付与されているか、「SPACE」のいづれかであること。
            row = 0
            For Each itemNo As String In kv.Value.項目番号
                row = row + 1
                If itemNo.Equals(ComConst.任意帳票.項目番号.SPACE) Then
                    Continue For
                End If

                Dim query = From dr In dtItem Where dr("項目番号").ToString = itemNo.TrimStart(ComConst.任意帳票.項目番号.前) Select dr
                If Not query.Count > 0 Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(10), cnt.ToString.PadLeft(2), kv.Key, row.ToString.PadLeft(3)))
                    ret = False
                    If cnt = max Then Return ret
                End If
            Next

            '11）再編成パラメータの項目番号（C27～C226セル）に、「前」が入力されている項目番号がある場合、画面上の比較に集計番号が設定されていること。
            row = 0
            For Each itemNo As String In kv.Value.項目番号
                row = row + 1
                If itemNo Like ComConst.任意帳票.項目番号.前 & "*" Then
                    If sKeyComp.Item1 Is Nothing Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(11), cnt.ToString.PadLeft(2), kv.Key, row.ToString.PadLeft(3)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
            Next

            row = 0
            For Each shihyobu As ComUtil.NiniChohyo.指標部 In kv.Value.指標部
                row = row + 1

                '12）再編成パラメータの地域（D27～D226セル）が、整数であるか。
                If Not Regex.IsMatch(shihyobu.地域, "^[0-9]+$") Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(12), cnt.ToString.PadLeft(2), kv.Key, row.ToString.PadLeft(3)))
                    ret = False
                    If cnt = max Then Return ret
                Else
                    Dim val As Decimal
                    If Decimal.TryParse(shihyobu.地域, val) Then
                        If Not val >= 0D Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(12), cnt.ToString.PadLeft(2), kv.Key, row.ToString.PadLeft(3)))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If
                End If

                '13）再編成パラメータの規模階層（E27～E226セル）が、整数であるか。
                If Not Regex.IsMatch(shihyobu.規模階層, "^[0-9]+$") Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(13), cnt.ToString.PadLeft(2), kv.Key, row.ToString.PadLeft(3)))
                    ret = False
                    If cnt = max Then Return ret
                Else
                    Dim val As Decimal
                    If Decimal.TryParse(shihyobu.規模階層, val) Then
                        If Not val >= 0D Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(13), cnt.ToString.PadLeft(2), kv.Key, row.ToString.PadLeft(3)))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If
                End If
            Next
        Next

        '14）再編成パラメータのシート数が20以下であるか。
        If Not parameters.Count <= 20 Then
            cnt = cnt + 1
            details.Add(String.Format(msg(14), cnt.ToString.PadLeft(2)))
            ret = False
            If cnt = max Then Return ret
        End If

        Return ret
    End Function

    ''' <summary>
    ''' Excel列英字列番号変換
    ''' </summary>
    ''' <param name="colStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConvertLetterToColNo(ByVal colStr As String) As Integer
        Dim ret As Integer = 0
        For i As Integer = 1 To Len(colStr)
            ret = ret * 26 + (Asc(UCase(Mid(colStr, i, 1))) - 64)
        Next
        Return ret
    End Function

    ''' <summary>
    ''' 任意帳票入力ファイルクラス
    ''' </summary>
    ''' <remarks></remarks>
    Private Class NiniChohyoInFile
        Inherits ExcelProcess

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            MyBase.New(True)
        End Sub

        ''' <summary>
        ''' パラメータ取得処理実行
        ''' </summary>
        ''' <param name="filePath"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ExecuteGetParameter(filePath As String) As Dictionary(Of String, ComUtil.NiniChohyo.編成パラメータ)
            Dim ret As Dictionary(Of String, ComUtil.NiniChohyo.編成パラメータ)

            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Sheets = Nothing

            'Excelアプリ無効
            Me.DisableExcelApp()

            Try
                'Workbookを開く
                xlBook = xlBooks.Open(filePath)
                Try
                    xlSheets = xlBook.Worksheets

                    'パラメータ取得処理
                    ret = Me.GetParameter(xlSheets)
                Catch ex As Exception
                    Throw ex
                Finally
                    'Sheetsの解放
                    ReleaseComObject(xlSheets)
                End Try
            Catch ex As Exception
                Throw ex
            Finally
                'Workbookを閉じる
                If xlBook IsNot Nothing Then
                    xlBook.Saved = True
                    xlBook.Close()
                End If
                'Workbookの解放
                ReleaseComObject(xlBook)
            End Try

            'Excelアプリ有効
            Me.EnableExcelApp()

            Return ret
        End Function

        ''' <summary>
        ''' パラメータ取得処理
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetParameter(xlSheets As Excel.Sheets) As Dictionary(Of String, ComUtil.NiniChohyo.編成パラメータ)
            Dim ret As New Dictionary(Of String, ComUtil.NiniChohyo.編成パラメータ)

            For i As Integer = 1 To xlSheets.Count
                Dim xlSheet As Excel.Worksheet = Nothing

                Try
                    'シートの設定
                    xlSheet = DirectCast(xlSheets.Item(i), Excel.Worksheet)

                    Dim prm As New ComUtil.NiniChohyo.編成パラメータ

                    Dim rng As Excel.Range = Nothing
                    Dim arrData(,) As Object

                    Try
                        'タイトル欄
                        rng = xlSheet.Range(ComConst.任意帳票.編成パラメータ.title)
                        prm.タイトル = rng.Formula.ToString
                        ReleaseComObject(rng)

                        '帳票レイアウト欄
                        rng = xlSheet.Range(ComConst.任意帳票.編成パラメータ.layout)
                        prm.帳票レイアウト = rng.Formula.ToString
                        ReleaseComObject(rng)

                        '開始位置セル欄
                        rng = xlSheet.Range(ComConst.任意帳票.編成パラメータ.startCell)
                        prm.開始位置セル = rng.Formula.ToString
                        ReleaseComObject(rng)

                        '生産費平均値種類欄
                        rng = xlSheet.Range(ComConst.任意帳票.編成パラメータ.seisanhiHeikinShurui)
                        prm.生産費平均値種類 = rng.Formula.ToString
                        ReleaseComObject(rng)

                        '田畑区分欄
                        rng = xlSheet.Range(ComConst.任意帳票.編成パラメータ.tahata)
                        prm.田畑区分指定 = rng.Formula.ToString
                        ReleaseComObject(rng)

                        'ビール麦販売区分欄
                        rng = xlSheet.Range(ComConst.任意帳票.編成パラメータ.beerMugiHanbai)
                        prm.ビール麦販売区分 = rng.Formula.ToString
                        ReleaseComObject(rng)

                        'てんさい栽培区分欄
                        rng = xlSheet.Range(ComConst.任意帳票.編成パラメータ.tensaiSaibai)
                        prm.てんさい栽培区分 = rng.Formula.ToString
                        ReleaseComObject(rng)

                        '表頭・表側欄
                        rng = xlSheet.Range(ComConst.任意帳票.編成パラメータ.hyotoHyosoku)
                        prm.表頭_表側 = rng.Formula.ToString
                        ReleaseComObject(rng)

                        '項目番号欄
                        Dim lstItemNo As New List(Of String)
                        rng = xlSheet.Range(ComConst.任意帳票.編成パラメータ.itemNo)
                        arrData = DirectCast(rng.Formula, Object(,))
                        For j As Integer = LBound(arrData) To UBound(arrData)
                            If arrData(j, 1).ToString.Equals(String.Empty) Then Exit For
                            lstItemNo.Add(arrData(j, 1).ToString)
                        Next
                        prm.項目番号 = lstItemNo
                        ReleaseComObject(rng)

                        '指標部欄
                        Dim lstShihyobu As New List(Of ComUtil.NiniChohyo.指標部)
                        rng = xlSheet.Range(ComConst.任意帳票.編成パラメータ.shihyobu)
                        arrData = DirectCast(rng.Formula, Object(,))
                        For j As Integer = LBound(arrData) To UBound(arrData)
                            If arrData(j, 1).ToString.Equals(String.Empty) Then Exit For
                            Dim shihyobu As New ComUtil.NiniChohyo.指標部
                            shihyobu.地域 = arrData(j, 1).ToString
                            shihyobu.規模階層 = arrData(j, 2).ToString
                            lstShihyobu.Add(shihyobu)
                        Next
                        prm.指標部 = lstShihyobu

                    Finally
                        ReleaseComObject(rng)
                    End Try

                    ret.Add(xlSheet.Name, prm)
                Finally
                    ReleaseComObject(xlSheet)
                End Try
            Next

            Return ret
        End Function

        ''' <summary>
        ''' シート名取得処理実行
        ''' </summary>
        ''' <param name="filePath"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ExecuteGetSheetName(filePath As String) As List(Of String)
            Dim ret As List(Of String)

            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Sheets = Nothing

            'Excelアプリ無効
            Me.DisableExcelApp()

            Try
                'Workbookを開く
                xlBook = xlBooks.Open(filePath)
                Try
                    xlSheets = xlBook.Worksheets

                    '取込処理
                    ret = GetGetSheetName(xlSheets)
                Catch ex As Exception
                    Throw ex
                Finally
                    'Sheetsの解放
                    ReleaseComObject(xlSheets)
                End Try
            Catch ex As Exception
                Throw ex
            Finally
                'Workbookを閉じる
                If xlBook IsNot Nothing Then
                    xlBook.Saved = True
                    xlBook.Close()
                End If
                'Workbookの解放
                ReleaseComObject(xlBook)
            End Try

            'Excelアプリ有効
            Me.EnableExcelApp()

            Return ret
        End Function

        ''' <summary>
        ''' シート名取得処理
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetGetSheetName(xlSheets As Excel.Sheets) As List(Of String)
            Dim ret As New List(Of String)

            For i As Integer = 1 To xlSheets.Count
                Dim xlSheet As Excel.Worksheet = Nothing

                Dim sheetName As String

                Try
                    'シートの設定
                    xlSheet = DirectCast(xlSheets.Item(i), Excel.Worksheet)

                    sheetName = xlSheet.Name
                Finally
                    ReleaseComObject(xlSheet)
                End Try

                ret.Add(sheetName)
            Next

            Return ret
        End Function
    End Class

    ''' <summary>
    ''' 任意帳票出力基底クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class NiniChohyoOutBaseClass
        Inherits ExcelOutputSingleBaseClass
        Implements IDisposable

        ''' <summary>帳票テンプレートファイル名</summary>
        Private _templateFile As String

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Protected Sub New()

        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Protected Sub New(templateFile As String, excelCheck As Boolean, reportCheck As Boolean, ByVal title As String, ByVal outPath As String, ByVal removalMacro As Boolean)
            MyBase.New(templateFile, excelCheck, reportCheck, title, outPath, removalMacro, False)

            _templateFile = templateFile

            Me.OpenExcel()
        End Sub

        ''' <summary>
        ''' 解放処理
        ''' </summary>
        ''' <param name="disposing"></param>
        ''' <remarks></remarks>
        Protected Overrides Sub Dispose(disposing As Boolean)
            If Not isDisposed Then
                If xlApp IsNot Nothing Then
                    'Sheetsの解放
                    ReleaseComObject(CObj(xlSheets))
                    'Workbookを閉じる
                    If xlBook IsNot Nothing Then
                        xlBook.Saved = True
                        xlBook.Close()
                    End If
                    'Workbooksの解放
                    ReleaseComObject(xlBooks)
                    xlApp.DisplayAlerts = True
                    'ガベージコレクト強制
                    GCCollect()
                    'Excelの終了
                    xlApp.Quit()
                    'ExcelApplicationの解放
                    ReleaseComObject(xlApp)
                    'ガベージコレクト強制
                    GCCollect()

                    'ストップウォッチを止める
                    sw.Stop()

                    'ユーザログ出力
                    OutputLog.WriteUserLog(OutputLog.LogLevel.Info, processName, OutputLog.LogType.SyoriEnd, sw.Elapsed.TotalSeconds.ToString("#0"))
                End If
            End If
            isDisposed = True
        End Sub

        ''' <summary>
        ''' Excelアプリケーションオープン
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub OpenExcel()
            Try
                'ストップウォッチを開始する
                sw.Start()

                'ユーザログ出力
                OutputLog.WriteUserLog(OutputLog.LogLevel.Info, processName, OutputLog.LogType.SyoriStart)

                'Excelアプリケーションの作成
                If xlApp Is Nothing Then
                    xlApp = New Excel.Application
                End If

                xlApp.DisplayAlerts = False
                'Excelオブジェクトの設定
                xlBooks = xlApp.Workbooks
                xlBook = xlBooks.Open(_templateFile)
                xlSheets = xlBook.Worksheets

            Catch ex As Exception
                MyBase.Dispose()
                Throw
            End Try
        End Sub
    End Class

    ''' <summary>
    ''' 任意帳票出力クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class NiniChohyoOut
        Inherits NiniChohyoOutBaseClass

        ''' <summary>主キー（本年）</summary>
        Private _pkeyThis As DAOSyukeiKekkahyo.PrimaryKey
        ''' <summary>項目キー（本年）</summary>
        Private _kkeyThis As DAOSyukeiKekkahyo.KomokuKey

        ''' <summary>主キー（比較）</summary>
        Private _pkeyComp As DAOSyukeiKekkahyo.PrimaryKey
        ''' <summary>項目キー（比較）</summary>
        Private _kkeyComp As DAOSyukeiKekkahyo.KomokuKey

        ''' <summary>編成パラメータ</summary>
        Private _parameters As Dictionary(Of String, ComUtil.NiniChohyo.編成パラメータ)

        ''' <summary>３経営体未満"x"置換</summary>
        Private _replace As Boolean

        ''' <summary>調査区分</summary>
        Private _chosakubun As String

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="outPath"></param>
        ''' <param name="skeyThis"></param>
        ''' <param name="sKeyComp"></param>
        ''' <param name="parameters"></param>
        ''' <param name="replace"></param>
        ''' <param name="chosakubun"></param>
        ''' <param name="layoutPath"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal outPath As String, _
                       skeyThis As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey), _
                       sKeyComp As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey), _
                       parameters As Dictionary(Of String, ComUtil.NiniChohyo.編成パラメータ), _
                       replace As Boolean, _
                       chosakubun As String, _
                       layoutPath As String)
            MyBase.New(layoutPath, True, False, ComConst.任意帳票.出力用ファイル名称.reportName, outPath, True)

            _pkeyThis = skeyThis.Item1
            _kkeyThis = skeyThis.Item2
            _pkeyComp = sKeyComp.Item1
            _kkeyComp = sKeyComp.Item2
            _parameters = parameters
            _replace = replace
            _chosakubun = chosakubun
        End Sub

        ''' <summary>
        ''' 帳票編集
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <remarks></remarks>
        Protected Overrides Sub ReportEdit(xlSheets As Excel.Sheets)
            Dim dtItem As DataTable

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '集計結果表項目マスタ取得
                ' REV_001↓
                'dtItem = DAOOther.GetSyukeiItemMaster(db, _chosakubun)
                dtItem = DAOOther.GetSyukeiItemMaster(db, _chosakubun, ComUtil.getVersionKubunTaikei(_pkeyThis.chosaNen, _chosakubun))
                ' REV_001↑

                For Each kv As KeyValuePair(Of String, ComUtil.NiniChohyo.編成パラメータ) In _parameters
                    Dim shihyobuCnt As Integer = 0
                    For Each shihyobu As ComUtil.NiniChohyo.指標部 In kv.Value.指標部

                        Dim dtThis As Dictionary(Of String, DataTable) = Nothing
                        Dim dcThis As Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目) = Nothing

                        Dim dtComp As Dictionary(Of String, DataTable) = Nothing
                        Dim dcComp As Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目) = Nothing

                        Dim kkeyThis As DAOSyukeiKekkahyo.KomokuKey = New DAOSyukeiKekkahyo.KomokuKey(_kkeyThis.kyoku, _kkeyThis.jimusho, _kkeyThis.kyoten)

                        kkeyThis.seisanhiHeikin = kv.Value.生産費平均値種類
                        kkeyThis.tahataKbn = kv.Value.田畑区分指定
                        kkeyThis.beerKbn = kv.Value.ビール麦販売区分
                        kkeyThis.tensaiKbn = kv.Value.てんさい栽培区分
                        kkeyThis.kiboKaisou = shihyobu.規模階層
                        kkeyThis.chiikiCd = shihyobu.地域

                        '集計結果表テーブル取得（任意帳票）
                        dtThis = DAOSyukeiKekkahyo.GetTableNiniChohyo(db, _chosakubun, _pkeyThis, kkeyThis)

                        '集計結果表項目取得
                        dcThis = ComUtil.SyukeiKekkahyo.GetItem(dtItem, dtThis)

                        Dim kkeyComp As DAOSyukeiKekkahyo.KomokuKey
                        If Not _pkeyComp Is Nothing Then
                            kkeyComp = New DAOSyukeiKekkahyo.KomokuKey(_kkeyComp.kyoku, _kkeyComp.jimusho, _kkeyComp.kyoten)

                            kkeyComp.seisanhiHeikin = kv.Value.生産費平均値種類
                            kkeyComp.tahataKbn = kv.Value.田畑区分指定
                            kkeyComp.beerKbn = kv.Value.ビール麦販売区分
                            kkeyComp.tensaiKbn = kv.Value.てんさい栽培区分
                            kkeyComp.kiboKaisou = shihyobu.規模階層
                            kkeyComp.chiikiCd = shihyobu.地域

                            '集計結果表テーブル取得（任意帳票）
                            dtComp = DAOSyukeiKekkahyo.GetTableNiniChohyo(db, _chosakubun, _pkeyComp, kkeyComp)

                            '集計結果表項目取得
                            dcComp = ComUtil.SyukeiKekkahyo.GetItem(dtItem, dtComp)
                        End If

                        '任意帳票シートデータ設定
                        Me.SetSheetData(dcThis, dcComp, xlSheets, kv.Value.帳票レイアウト, kv.Value.表頭_表側, kv.Value.開始位置セル, kv.Value.項目番号, shihyobuCnt, _replace, _chosakubun)

                        shihyobuCnt = shihyobuCnt + 1

                        '進捗加増
                        Me.ProgressAddValue = 1
                    Next

                    '進捗加増
                    Me.ProgressAddValue = 1
                Next
            End Using
        End Sub

        ''' <summary>
        ''' 任意帳票シートデータ設定
        ''' </summary>
        ''' <param name="dcThis"></param>
        ''' <param name="dcComp"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="sheetName"></param>
        ''' <param name="hyotoHyosoku"></param>
        ''' <param name="startCell"></param>
        ''' <param name="lstItemNo"></param>
        ''' <param name="shihyobuCnt"></param>
        ''' <param name="chosakubun"></param>
        ''' <remarks></remarks>
        Private Sub SetSheetData(dcThis As Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目), dcComp As Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目), xlSheets As Excel.Sheets, sheetName As String, hyotoHyosoku As String, startCell As String, lstItemNo As List(Of String), shihyobuCnt As Integer, replace As Boolean, chosakubun As String)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                xlSheet = DirectCast(xlSheets.Item(sheetName), Excel.Worksheet)

                Dim rngStart As Excel.Range = Nothing
                Try
                    rngStart = xlSheet.Range(startCell)

                    Dim rngResize As Excel.Range = Nothing
                    Dim rngOffset As Excel.Range = Nothing
                    Try
                        If hyotoHyosoku.Equals(ComConst.任意帳票.表頭_表側.リスト(ComConst.任意帳票.表頭_表側.enm.表側)) Then
                            rngOffset = rngStart.Offset(0, shihyobuCnt)
                            rngResize = rngOffset.Resize(If(lstItemNo.Count = 1, lstItemNo.Count + 1, lstItemNo.Count), 1)
                        Else
                            rngOffset = rngStart.Offset(shihyobuCnt, 0)
                            rngResize = rngOffset.Resize(1, If(lstItemNo.Count = 1, lstItemNo.Count + 1, lstItemNo.Count))
                        End If

                        Dim arrData(,) As Object = DirectCast(rngResize.Formula, Object(,))

                        For itemCnt As Integer = 0 To lstItemNo.Count - 1
                            If Not lstItemNo(itemCnt).Equals(ComConst.任意帳票.項目番号.SPACE) Then
                                If lstItemNo(itemCnt) Like ComConst.任意帳票.項目番号.前 & "*" Then
                                    '値設定
                                    Me.SetValue(arrData, dcComp, lstItemNo(itemCnt).TrimStart(ComConst.任意帳票.項目番号.前), rngStart, hyotoHyosoku, itemCnt, replace, chosakubun)
                                Else
                                    '値設定
                                    Me.SetValue(arrData, dcThis, lstItemNo(itemCnt), rngStart, hyotoHyosoku, itemCnt, replace, chosakubun)
                                End If
                            End If
                        Next

                        rngResize.Value = arrData
                        rngResize.Value = rngResize.Formula
                    Finally
                        ReleaseComObject(rngResize)
                        ReleaseComObject(rngOffset)
                    End Try
                Finally
                    ReleaseComObject(rngStart)
                End Try

            Finally
                ReleaseComObject(xlSheet)
            End Try
        End Sub

        ''' <summary>
        ''' 値設定
        ''' </summary>
        ''' <param name="arrData"></param>
        ''' <param name="dc"></param>
        ''' <param name="itemNo"></param>
        ''' <param name="rngStart"></param>
        ''' <param name="hyotoHyosoku"></param>
        ''' <param name="itemCnt"></param>
        ''' <param name="replace"></param>
        ''' <param name="chosakubun"></param>
        ''' <remarks></remarks>
        Private Sub SetValue(arrData(,) As Object, dc As Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目), itemNo As String, rngStart As Excel.Range, hyotoHyosoku As String, itemCnt As Integer, replace As Boolean, chosakubun As String)
            Dim row As Integer
            Dim col As Integer

            If dc.ContainsKey(itemNo) Then
                If hyotoHyosoku.Equals(ComConst.任意帳票.表頭_表側.リスト(ComConst.任意帳票.表頭_表側.enm.表側)) Then
                    row = itemCnt + 1
                    col = 1
                Else
                    row = 1
                    col = itemCnt + 1
                End If

                Dim val As Decimal
                If Decimal.TryParse(dc(ComConst.集計結果表.集計戸数(chosakubun)).値, val) Then
                    If replace And Not itemNo.Equals(ComConst.集計結果表.集計戸数(chosakubun)) And val < 3D Then
                        arrData(row, col) = ComConst.任意帳票.出力用ファイル名称.replaceChr
                    Else
                        arrData(row, col) = If(dc(itemNo).値.Equals(String.Empty), ComConst.任意帳票.出力用ファイル名称.replaceBlank, dc(itemNo).値)
                    End If
                Else
                    arrData(row, col) = If(dc(itemNo).値.Equals(String.Empty), ComConst.任意帳票.出力用ファイル名称.replaceBlank, dc(itemNo).値)
                End If
            End If
        End Sub
    End Class
End Class
