Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.12.19 |Daiko               | 要件No4 バージョン区分追加
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 任意階層修正画面
''' </summary>
''' <remarks></remarks>
Public Class BRA5910F

    ''' <summary>階層規模テキストボックス名称</summary>
    Private Const KAISOU_NAME As String = "txtKaisou"
    ''' <summary>階層規模テキストボックス上限</summary>
    Private Const KAISOU_MAX As String = "Max"
    ''' <summary>階層規模テキストボックス下限</summary>
    Private Const KAISOU_MIN As String = "Min"

    ''' <summary>階層規模件数</summary>
    Private Const KAISOU_COUNT As Integer = 20

    ''' <summary>ファイルタイトル文字列</summary>
    Private Const FILE_TITLE As String = "任意階層"

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA5910F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            ' REV_001↓
            'バージョン区分コンボボックス設定
            ComUtil.SetVersionKbnComboBox(cboVerKubun)
            If ComConst.令和４年体系.対象調査区分2022.IndexOf(CommonInfo.Chosakubun) > -1 _
                Or ComConst.令和４年体系.対象調査区分2023.IndexOf(CommonInfo.Chosakubun) > -1 Then
                cboVerKubun.SelectedValue = ComConst.バージョン区分.結果表等項目2022
            Else
                cboVerKubun.SelectedValue = ComConst.バージョン区分.結果表等項目2021
                cboVerKubun.Visible = False
                lblVerKubun.Visible = False
            End If
            ' REV_001↑

            '営農経営体区分コンボボックス設定
            ComUtil.SetEinouKeieitaiComboBox(lblEinouKeieitai, cboEinouKeieitai)

            ' REV_001↓
            'If cboEinouKeieitai.Visible Then
            '    cboEinouKeieitai.SelectedValue = CommonInfo.Chosakubun
            'Else
            '    'データ表示
            '    Me.SetData(CommonInfo.Chosakubun)
            'End If
            ' REV_001↑
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 設定ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSetting_Click(sender As Object, e As EventArgs) Handles btnSetting.Click
        Try
            '調査区分取得
            Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

            '規模階層判定項番取得
            Dim kiboKaisou As String = txtKiboKaisou.Text

            '規模階層範囲取得
            Dim kaisou As New List(Of ValueTuple(Of String, String))
            For i As Integer = 1 To KAISOU_COUNT
                Dim vtp As New ValueTuple(Of String, String)

                Dim txtMax As String = DirectCast(Me.GroupBox1.Controls(KAISOU_NAME & i & KAISOU_MAX), TextBox).Text
                Dim txtMin As String = DirectCast(Me.GroupBox1.Controls(KAISOU_NAME & i & KAISOU_MIN), TextBox).Text

                vtp = ValueTuple.Create(txtMax, txtMin)

                kaisou.Add(vtp)
            Next

            'エラーチェック
            Dim details As New List(Of String)
            ' REV_001↓
            'If Not Me.CheckError(chosakubun, kiboKaisou, kaisou, details) Then
            If Not Me.CheckError(chosakubun, cboVerKubun.SelectedValue.ToString, kiboKaisou, kaisou, details) Then
                ' REV_001↑
                'エラーメッセージ
                Message.ShowMsgForm(Me, MessageID.MSG_E_010, {String.Join(vbCrLf, details)})
                Exit Sub
            End If

            'データ取得
            Dim lstDc As List(Of Dictionary(Of String, String)) = Me.GetData(True)

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try

                    db.BeginTrans()

                    '任意階層削除
                    ' REV_001↓
                    'DAOOther.DeleteNiniKaisou(db, chosakubun)
                    DAOOther.DeleteNiniKaisou(db, chosakubun, cboVerKubun.SelectedValue.ToString)
                    ' REV_001↑

                    For Each dc As Dictionary(Of String, String) In lstDc
                        '任意階層追加
                        ' REV_001↓
                        'DAOOther.InsertNiniKaisou(db, chosakubun, dc)
                        DAOOther.InsertNiniKaisou(db, chosakubun, cboVerKubun.SelectedValue.ToString, dc)
                        ' REV_001↑
                    Next

                    db.CommitTrans()

                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using

            '完了メッセージ
            Message.ShowMsgBox(MessageID.MSG_I_001, MsgBoxStyle.OkOnly)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub cboEinouKeieitai_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboEinouKeieitai.SelectedIndexChanged
        Try
            '調査区分取得
            Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

            'データ表示
            ' REV_001↓
            'Me.SetData(chosakubun)
            Me.SetData(chosakubun, cboVerKubun)
            ' REV_001↑
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ' REV_001↓
    Private Sub cboVerKubun_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboVerKubun.SelectedIndexChanged
        Try
            '調査区分取得
            Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

            'データ表示
            Me.SetData(chosakubun, cboVerKubun)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
    ' REV_001↑

    Private Sub btnFileInput_Click(sender As Object, e As EventArgs) Handles btnFileInput.Click
        Try
            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of OpenFileDialog)(Me, IniFileInfo.ExcelInPath, , ComConst.CSVファイル.FILEDIALOG_CSV_FILEFILTER)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '入力ファイル取得
            Dim lstArr As List(Of String()) = Me.GetInputFile(filePath)

            'ファイル形式チェック
            Dim msgId As String = String.Empty
            If Not Me.CheckFileFormat(lstArr, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '調査区分取得
            Dim chosakubun As String
            If {ComConst.調査区分.営農類型別経営統計_個人, ComConst.調査区分.営農類型別経営統計_法人}.Contains(CommonInfo.Chosakubun) Then
                chosakubun = (From dr In ComConst.営農経営体区分.リスト Where dr.Value.名称２ = lstArr(0)(1) Select dr.Key).ToArray(0)
            Else
                chosakubun = (From dr In ComConst.調査区分.リスト Where dr.Value.名称 = lstArr(0)(1) Select dr.Key).ToArray(0)
            End If

            '規模階層判定項番取得
            Dim kiboKaisou As String = lstArr(0)(2)

            '規模階層範囲取得
            Dim kaisou As New List(Of ValueTuple(Of String, String))
            For i As Integer = 1 To KAISOU_COUNT
                Dim vtp As New ValueTuple(Of String, String)

                Dim txtMax As String = lstArr(i)(1)
                Dim txtMin As String = lstArr(i)(2)

                vtp = ValueTuple.Create(txtMax, txtMin)

                kaisou.Add(vtp)
            Next

            'エラーチェック
            Dim details As New List(Of String)
            ' REV_001↓
            'If Not Me.CheckError(chosakubun, kiboKaisou, kaisou, details) Then
            If Not Me.CheckError(chosakubun, cboVerKubun.SelectedValue.ToString, kiboKaisou, kaisou, details) Then
                ' REV_001↑
                'エラーメッセージ
                Message.ShowMsgForm(Me, MessageID.MSG_E_010, {String.Join(vbCrLf, details)})
                Exit Sub
            End If

            '入力ファイル設定
            Me.SetInputFile(lstArr)

            '完了メッセージ
            Message.ShowMsgBox(MessageID.MSG_I_025, MsgBoxStyle.OkOnly)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnFileOutput_Click(sender As Object, e As EventArgs) Handles btnFileOutput.Click
        Try
            '調査区分名取得
            Dim chosakubunName As String = ComUtil.GetChosakubunName(ComUtil.GetChosakubun(cboEinouKeieitai))

            Dim fileName As String = FILE_TITLE & "_" _
                                    & chosakubunName & "_" _
                                    & CommonInfo.Kyoku & ".csv"

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName, ComConst.CSVファイル.FILEDIALOG_CSV_FILEFILTER)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'データ取得
            Dim lstDc As List(Of Dictionary(Of String, String)) = Me.GetData(False)

            '出力ファイル出力
            Dim ret As ComConst.CSVファイル.enmOutputReturn = Me.PutOutputFile(filePath, lstDc, chosakubunName)

            Select Case ret
                Case ComConst.CSVファイル.enmOutputReturn.OK
                    '完了メッセージ
                    Message.ShowMsgBox(MessageID.MSG_I_026, MsgBoxStyle.OkOnly)
                Case ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS
                    'エラーメッセージ
                    Message.ShowMsgBox(MessageID.MSG_E_058, MsgBoxStyle.OkOnly)
            End Select
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' データクリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearData()
        'テキストクリア
        txtKiboKaisou.Clear()

        For Each ctl In Me.GroupBox1.Controls
            If TypeOf ctl Is TextBox Then
                DirectCast(ctl, TextBox).Clear()
            End If
        Next
    End Sub

    ' REV_001↓
    ''' <summary>
    ''' データ設定
    ''' </summary>
    ''' <param name="chosakubun"></param>
    ''' <param name="cbo"></param>
    'Private Sub SetData(chosakubun As String)
    Private Sub SetData(chosakubun As String, cbo As ComboBox)
        ' REV_001↑
        'データクリア
        Me.ClearData()

        Dim dt As DataTable

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '任意階層項目番号取得
            ' REV_001↓
            'dt = DAOOther.GetNiniKaisouItemNo(db, chosakubun)
            dt = DAOOther.GetNiniKaisouItemNo(db, chosakubun, cbo.SelectedValue.ToString)
            ' REV_001↑

            If dt.Rows.Count > 0 Then
                txtKiboKaisou.Text = dt.Rows(0)("項目番号").ToString
            End If
            dt.Clear()

            '任意階層取得
            ' REV_001↓
            'dt = DAOOther.GetNiniKaisou(db, chosakubun)
            dt = DAOOther.GetNiniKaisou(db, chosakubun, cbo.SelectedValue.ToString)
            ' REV_001↑
            For Each row As DataRow In dt.Rows
                DirectCast(Me.GroupBox1.Controls(KAISOU_NAME & row("規模階層").ToString & KAISOU_MAX), TextBox).Text = row("上限").ToString
                DirectCast(Me.GroupBox1.Controls(KAISOU_NAME & row("規模階層").ToString & KAISOU_MIN), TextBox).Text = row("下限").ToString
            Next
        End Using
    End Sub

    ''' <summary>
    ''' データ取得
    ''' </summary>
    ''' <param name="condition"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetData(condition As Boolean) As List(Of Dictionary(Of String, String))
        Dim ret As New List(Of Dictionary(Of String, String))

        Dim kiboKaisou As String = txtKiboKaisou.Text

        For i As Integer = 1 To KAISOU_COUNT
            Dim txtMax As String = DirectCast(Me.GroupBox1.Controls(KAISOU_NAME & i & KAISOU_MAX), TextBox).Text
            Dim txtMin As String = DirectCast(Me.GroupBox1.Controls(KAISOU_NAME & i & KAISOU_MIN), TextBox).Text

            If condition Then
                If Not txtMax.Equals(String.Empty) And Not txtMin.Equals(String.Empty) Then
                    Dim dc As Dictionary(Of String, String) = Me.SetItem(kiboKaisou, i, txtMax, txtMin)
                    ret.Add(dc)
                End If
            Else
                Dim dc As Dictionary(Of String, String) = Me.SetItem(kiboKaisou, i, txtMax, txtMin)
                ret.Add(dc)
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 項目設定
    ''' </summary>
    ''' <param name="i"></param>
    ''' <param name="max"></param>
    ''' <param name="min"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetItem(kiboKaisou As String, i As Integer, max As String, min As String) As Dictionary(Of String, String)
        Dim ret As New Dictionary(Of String, String)

        ret("項目番号") = txtKiboKaisou.Text
        ret("規模階層") = i.ToString
        ret("上限") = max
        ret("下限") = min

        Return ret
    End Function

    ' REV_001↓
    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="chosakubun"></param>
    ''' <param name="kiboKaisou"></param>
    ''' <param name="kaisou"></param>
    ''' <param name="details"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    'Private Function CheckError(chosakubun As String, kiboKaisou As String, kaisou As List(Of ValueTuple(Of String, String)), ByRef details As List(Of String)) As Boolean
    Private Function CheckError(chosakubun As String, versionKbn As String, kiboKaisou As String, kaisou As List(Of ValueTuple(Of String, String)), ByRef details As List(Of String)) As Boolean
        ' REV_001↑
        Dim ret As Boolean = True
        Dim dtKobetsuItemMst As DataTable

        Const max As Integer = ComConst.ERR_MESSAGE_MAX

        Dim msg As String() = {"" _
                             , "{0}件目：「規模階層判定項番」を入力してください。" _
                             , "{0}件目：「規模階層判定項番」に存在しない項番が入力されております。" _
                             , "{0}件目：「データ範囲下限」は正の整数12桁、小数3桁までで入力してください。" _
                             , "{0}件目：「データ範囲上限」は正の整数12桁、小数3桁までで入力してください。" _
                             , "{0}件目：規模階層（{1}）の「データ範囲上限」と「データ範囲下限」は片方だけの入力は出来ません。両方入力してください。" _
                             , "{0}件目：規模階層（{1}）の「データ範囲上限」は「データ範囲下限」より大きな値で入力してください。" _
                             , "{0}件目：規模階層（{1}）は階層を詰めて入力してください。"
        }

        Dim cnt As Integer = 0

        '1）規模階層判定項番が入力されているか。
        If kiboKaisou.Equals(String.Empty) Then
            cnt = cnt + 1
            details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2)))
            ret = False
            If cnt = max Then Return ret
        End If

        '2）規模階層判定項番が存在するか。
        If Not kiboKaisou.Equals(String.Empty) Then
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '個別結果表項目マスタ取得
                If chosakubun.Equals(ComConst.営農経営体区分.農業経営体) Then
                    ' REV_001↓
                    'dtKobetsuItemMst = DAOOther.GetNougyouKeieitaiKobetsuKekkahyoItemMaster(db)
                    dtKobetsuItemMst = DAOOther.GetNougyouKeieitaiKobetsuKekkahyoItemMaster(db, versionKbn)
                    ' REV_001↑
                Else
                    ' REV_001↓
                    'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun)
                    dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, versionKbn)
                    ' REV_001↑
                End If
            End Using

            If chosakubun.Equals(ComConst.営農経営体区分.農業経営体) And Regex.IsMatch(kiboKaisou, "[【]+.+[】]") Then
                '判定項番が固定文字列の場合のチェック
                ' REV_001↓
                'If (Not ComUtil.CheckConstJouken(kiboKaisou, ComConst.調査区分.営農類型別経営統計_個人)) Or (Not ComUtil.CheckConstJouken(kiboKaisou, ComConst.調査区分.営農類型別経営統計_法人)) Then
                If (Not ComUtil.CheckConstJouken(kiboKaisou, ComConst.調査区分.営農類型別経営統計_個人, versionKbn)) Or (Not ComUtil.CheckConstJouken(kiboKaisou, ComConst.調査区分.営農類型別経営統計_法人, versionKbn)) Then
                    ' REV_001↑
                    cnt = cnt + 1
                    details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2)))
                    ret = False
                    If cnt = max Then Return ret
                End If
            Else
                '項番の場合のチェック
                Dim query1 = From dr In dtKobetsuItemMst Where dr("項目番号").ToString = kiboKaisou Select dr
                If Not query1.Count > 0 Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2)))
                    ret = False
                    If cnt = max Then Return ret
                End If
            End If
        End If

        '3）データ範囲下限が正の整数12桁まで、小数3桁までであるか。
        For i As Integer = 1 To KAISOU_COUNT
            Dim txt As String = kaisou(i - 1).Item2
            If Not txt.Equals(String.Empty) AndAlso Not Regex.IsMatch(txt, "^[0-9]{1,12}(\.[0-9]{1,3})?$") Then
                cnt = cnt + 1
                details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2)))
                ret = False
                If cnt = max Then Return ret
                Exit For
            End If
        Next

        '4）データ範囲上限が正の整数12桁まで、小数3桁までであるか。
        For i As Integer = 1 To KAISOU_COUNT
            Dim txt As String = kaisou(i - 1).Item1
            If Not txt.Equals(String.Empty) AndAlso Not Regex.IsMatch(txt, "^[0-9]{1,12}(\.[0-9]{1,3})?$") Then
                cnt = cnt + 1
                details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2)))
                ret = False
                If cnt = max Then Return ret
                Exit For
            End If
        Next

        '5）規模階層1つに対して入力がある場合、データ範囲上限とデータ範囲下限両方とも入力されていること。
        For i As Integer = 1 To KAISOU_COUNT
            Dim txtMax As String = kaisou(i - 1).Item1
            Dim txtMin As String = kaisou(i - 1).Item2
            If txtMax.Equals(String.Empty) Xor txtMin.Equals(String.Empty) Then
                cnt = cnt + 1
                details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), i.ToString))
                ret = False
                If cnt = max Then Return ret
            End If
        Next

        '6）データ範囲上限 ≧ データ範囲下限でがあるか。
        For i As Integer = 1 To KAISOU_COUNT
            Dim txtMax As String = kaisou(i - 1).Item1
            Dim txtMin As String = kaisou(i - 1).Item2
            If Not txtMax.Equals(String.Empty) And Not txtMin.Equals(String.Empty) Then
                Dim valMax As Decimal
                Dim valMin As Decimal
                If Decimal.TryParse(txtMax, valMax) And Decimal.TryParse(txtMin, valMin) Then
                    If Not valMax >= valMin Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), i.ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
            End If
        Next

        '7）規模階層に連続して入力がされているか。
        For i As Integer = 2 To KAISOU_COUNT
            Dim txtMax As String = kaisou(i - 1).Item1
            Dim txtMin As String = kaisou(i - 1).Item2
            If Not txtMax.Equals(String.Empty) And Not txtMin.Equals(String.Empty) Then
                txtMax = kaisou(i - 2).Item1
                txtMin = kaisou(i - 2).Item2
                If txtMax.Equals(String.Empty) Or txtMin.Equals(String.Empty) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), i.ToString))
                    ret = False
                    If cnt = max Then Return ret
                End If
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 出力ファイル出力
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <param name="lstDc"></param>
    ''' <param name="chosakubunName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PutOutputFile(filePath As String, lstDc As List(Of Dictionary(Of String, String)), chosakubunName As String) As ComConst.CSVファイル.enmOutputReturn
        Dim ret As ComConst.CSVファイル.enmOutputReturn = ComConst.CSVファイル.enmOutputReturn.CANCEL

        Dim sjisEnc As Encoding = Encoding.GetEncoding(ComConst.CSVファイル.CODEPAGE_SHIFT_JIS)

        'ディレクトリ作成
        Dim outDir As String = Path.GetDirectoryName(filePath)
        If Not System.IO.Directory.Exists(outDir) Then
            Directory.CreateDirectory(outDir)
        End If

        'ファイル存在チェック
        If IO.File.Exists(filePath) Then
            If Message.ShowMsgBox(MessageID.MSG_Q_004, {IO.Path.GetFileName(filePath)}, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                Return ret
            End If
        End If

        Try
            'ファイル出力
            Using sw As New System.IO.StreamWriter(filePath, False, sjisEnc)
                sw.WriteLine(ComConst.CSVファイル.START_ADDITION & FILE_TITLE & ComConst.CSVファイル.END_ADDITION & ComConst.CSVファイル.CSV_DELIMITER _
                             & ComConst.CSVファイル.START_ADDITION & chosakubunName & ComConst.CSVファイル.END_ADDITION & ComConst.CSVファイル.CSV_DELIMITER _
                             & ComConst.CSVファイル.START_ADDITION & lstDc(0)("項目番号") & ComConst.CSVファイル.END_ADDITION)

                For Each dc As Dictionary(Of String, String) In lstDc
                    Dim arr As Object() = {dc("規模階層"), dc("上限"), dc("下限")}
                    sw.WriteLine(ComConst.CSVファイル.START_ADDITION _
                                 & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, arr) _
                                 & ComConst.CSVファイル.END_ADDITION)
                Next
            End Using
            ret = ComConst.CSVファイル.enmOutputReturn.OK
        Catch ex As Exception
            ret = ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 入力ファイル取得
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetInputFile(filePath As String) As List(Of String())
        Dim ret As List(Of String()) = Nothing

        Dim sjisEnc As Encoding = Encoding.GetEncoding(ComConst.CSVファイル.CODEPAGE_SHIFT_JIS)

        'ファイル入力
        If System.IO.File.Exists(filePath) Then
            ret = New List(Of String())

            Using parser As New TextFieldParser(filePath, sjisEnc)

                parser.TextFieldType = FieldType.Delimited
                parser.SetDelimiters(ComConst.CSVファイル.CSV_DELIMITER)

                While Not parser.EndOfData
                    Dim arr As String() = parser.ReadFields()
                    ret.Add(arr)
                End While
            End Using
        End If

        Return ret
    End Function

    ''' <summary>
    ''' ファイル形式チェック
    ''' </summary>
    ''' <param name="lstArr"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckFileFormat(lstArr As List(Of String()), ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        '１行目、１項目目チェック
        If Not lstArr(0)(0).Equals(FILE_TITLE) Then
            msgId = MessageID.MSG_E_019
            Return ret
        End If

        '１行目、２項目目チェック
        If Not lstArr(0)(1).Equals(CommonInfo.ChosakubunName) Then
            If Not ({ComConst.調査区分.営農類型別経営統計_個人, ComConst.調査区分.営農類型別経営統計_法人}.Contains(CommonInfo.Chosakubun) _
                    And lstArr(0)(1).Equals(ComConst.営農経営体区分.リスト(ComConst.営農経営体区分.農業経営体).名称２)) Then
                msgId = MessageID.MSG_E_019
                Return ret
            End If
        End If

        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' 入力ファイル設定
    ''' </summary>
    ''' <param name="lstArr"></param>
    ''' <remarks></remarks>
    Private Sub SetInputFile(lstArr As List(Of String()))
        'データクリア
        Me.ClearData()

        If lstArr.Count > 0 Then
            txtKiboKaisou.Text = lstArr(0)(2)

            If {ComConst.調査区分.営農類型別経営統計_個人, ComConst.調査区分.営農類型別経営統計_法人}.Contains(CommonInfo.Chosakubun) Then
                RemoveHandler cboEinouKeieitai.SelectedIndexChanged, AddressOf Me.cboEinouKeieitai_SelectedIndexChanged

                cboEinouKeieitai.SelectedValue = (From dr In ComConst.営農経営体区分.リスト Where dr.Value.名称２ = lstArr(0)(1) Select dr.Key).ToArray(0)

                AddHandler cboEinouKeieitai.SelectedIndexChanged, AddressOf Me.cboEinouKeieitai_SelectedIndexChanged
            End If

            Dim cnt As Integer = If(lstArr.Count - 1 > KAISOU_COUNT, KAISOU_COUNT, lstArr.Count - 1)

            For i As Integer = 1 To cnt
                DirectCast(Me.GroupBox1.Controls(KAISOU_NAME & i & KAISOU_MAX), TextBox).Text = lstArr(i)(1)
                DirectCast(Me.GroupBox1.Controls(KAISOU_NAME & i & KAISOU_MIN), TextBox).Text = lstArr(i)(2)
            Next i
        End If
    End Sub
End Class
