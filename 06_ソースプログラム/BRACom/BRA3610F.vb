''' <summary>
''' 個別結果表チェックリスト出力画面
''' </summary>
''' <remarks></remarks>
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.11.24 |TSP)                | フェーズ3 追加要件No.14,15 修正
'//  REV_002   | 2022.10.11 |Daiko               | 要件No1 バージョン区分追加
'//  REV_003   | 2023.08.08 |Daiko               | 要件No.8 "前"付項番は前年を参照する
'//  REV_004   | 2023.11.27 |Daiko               | 要件No.17 1客体の場合は出力ファイル名にセンサス番号を付与
'//  REV_005   | 2023.11.27 |GCU                 | 要件No.6 チェックリストファイル名への調査年産等挿入
'//            |            |                    |
'//*************************************************************************************************
Public Class BRA3610F

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
    Private Sub BRA3610F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '調査年コンボボックス設定
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.KobetsuKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
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
    ''' 農政局コンボボックス変更時
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
    ''' 基本チェックボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBaseCheck_Click(sender As Object, e As EventArgs) Handles btnBaseCheck.Click
        Dim errCheckType As ComConst.エラーチェック種別.enm
        Dim fileName As String
        Dim ProgressDialog As New ProgressDialog()

        Try
            errCheckType = ComConst.エラーチェック種別.enm.基本

            Dim pKeys As List(Of DAOKobetsuKekkahyo.PrimaryKey) = Me.GetKobetsuKekkahyoPrimaryKey(_chosaNen)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(pKeys, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim dt As DataTable

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '個別結果表審査論理取得
                ' REV_002↓
                'dt = DAOOther.GetKobetsuKekkahyoShinsaRonri(db, errCheckType)
                dt = DAOOther.GetKobetsuKekkahyoShinsaRonri(db, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun), errCheckType)
                ' REV_002↑
            End Using

            If dt.Rows.Count = 0 Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_025, MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            'REV_005
            If CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計 Then
                '対象客体が1件の場合はファイル名にセンサス番号を付与
                If pKeys.Count = 1 Then
                    fileName = String.Format(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.reportName, ComConst.エラーチェック種別.リスト(errCheckType)) & "_" _
                                    & CommonInfo.ChosakubunName & "_" _
                                    & CommonInfo.Kyoku & "_" _
                                    & CommonInfo.Jimusyo & "_" _
                                    & CommonInfo.Center & "_" _
                                    & pKeys(0).censusNo & ".xlsx"
                Else
                    fileName = String.Format(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.reportName, ComConst.エラーチェック種別.リスト(errCheckType)) & "_" _
                                    & CommonInfo.ChosakubunName & "_" _
                                    & CommonInfo.Kyoku & "_" _
                                    & CommonInfo.Jimusyo & "_" _
                                    & CommonInfo.Center & ".xlsx"
                End If
            Else
                '対象客体が1件の場合はファイル名にセンサス番号を付与
                If pKeys.Count = 1 Then
                    fileName = String.Format(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.reportName, ComConst.エラーチェック種別.リスト(errCheckType)) & "_" _
                                    & CommonInfo.ChosakubunName & "_" _
                                    & CommonInfo.Kyoku & "_" _
                                    & CommonInfo.Jimusyo & "_" _
                                    & CommonInfo.Center & "_" _
                                    & pKeys(0).censusNo & "_" & _chosaNen & "_" & Now.ToString("yyyyMMddHHmm") & ".xlsx"
                Else
                    fileName = String.Format(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.reportName, ComConst.エラーチェック種別.リスト(errCheckType)) & "_" _
                                    & CommonInfo.ChosakubunName & "_" _
                                    & CommonInfo.Kyoku & "_" _
                                    & CommonInfo.Jimusyo & "_" _
                                    & CommonInfo.Center & "_" & _chosaNen & "_" & Now.ToString("yyyyMMddHHmm") & ".xlsx"
                End If
            End If
            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath), fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '進捗ダイアログを表示する
            Dim adjust As Integer = CType(Math.Ceiling(dt.Rows.Count * pKeys.Count * 0.05), Integer)
            ProgressDialog.Maximum = dt.Rows.Count * pKeys.Count + adjust
            ProgressDialog.Show(Me)

            '個別結果表チェックリスト(基本・追加)取得
            Dim lst As List(Of Dictionary(Of String, String)) = Me.GetKobetsuKekkahyoCheckList(errCheckType, ProgressDialog)

            If lst.Count = 0 Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_I_007, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '個別結果表チェックリスト出力を行う
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New BRA3610R(errCheckType, _chosaNen, lst, filePath)
                    ret = ExcelOutput.Execute(MessageID.MSG_Q_004, ProgressDialog)
                End Using

                '進捗を進める
                ProgressDialog.AddValue = adjust

                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()

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

        Catch ex As ShinsaException
            '進捗ダイアログを閉じる
            ProgressDialog.endDispose()
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_056, {ex.CensusNo, ex.ErrSign}, MsgBoxStyle.OkOnly)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            If Not ProgressDialog Is Nothing Then
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                ProgressDialog = Nothing
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 範囲チェックボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRangeCheck_Click(sender As Object, e As EventArgs) Handles btnRangeCheck.Click
        Dim ProgressDialog As New ProgressDialog()

        Try
            Dim pKeys As List(Of DAOKobetsuKekkahyo.PrimaryKey) = Me.GetKobetsuKekkahyoPrimaryKey(_chosaNen)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(pKeys, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim dt As DataTable

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '個別結果表審査論理範囲取得
                ' REV_002↓
                'dt = DAOOther.GetKobetsuKekkahyoShinsaRonriRange(db)
                dt = DAOOther.GetKobetsuKekkahyoShinsaRonriRange(db, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
                ' REV_002↑
            End Using

            If dt.Rows.Count = 0 Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_029, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim fileName As String
            'REV_005
            If CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計 Then
                '対象客体が1件の場合はファイル名にセンサス番号を付与
                If pKeys.Count = 1 Then
                    fileName = ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.reportName & "_" _
                                    & CommonInfo.ChosakubunName & "_" _
                                    & CommonInfo.Kyoku & "_" _
                                    & CommonInfo.Jimusyo & "_" _
                                    & CommonInfo.Center & "_" _
                                    & pKeys(0).censusNo & ".xlsx"
                Else
                    fileName = ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.reportName & "_" _
                                    & CommonInfo.ChosakubunName & "_" _
                                    & CommonInfo.Kyoku & "_" _
                                    & CommonInfo.Jimusyo & "_" _
                                    & CommonInfo.Center & ".xlsx"
                End If
            Else
                '対象客体が1件の場合はファイル名にセンサス番号を付与
                If pKeys.Count = 1 Then
                    fileName = ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.reportName & "_" _
                                    & CommonInfo.ChosakubunName & "_" _
                                    & CommonInfo.Kyoku & "_" _
                                    & CommonInfo.Jimusyo & "_" _
                                    & CommonInfo.Center & "_" _
                                    & pKeys(0).censusNo & "_" & _chosaNen & "_" & Now.ToString("yyyyMMddHHmm") & ".xlsx"
                Else
                    fileName = ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.reportName & "_" _
                                    & CommonInfo.ChosakubunName & "_" _
                                    & CommonInfo.Kyoku & "_" _
                                    & CommonInfo.Jimusyo & "_" _
                                    & CommonInfo.Center & "_" & _chosaNen & "_" & Now.ToString("yyyyMMddHHmm") & ".xlsx"
                End If
            End If

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath), fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '進捗ダイアログを表示する
            Dim adjust As Integer = CType(Math.Ceiling(dt.Rows.Count * pKeys.Count * 0.05), Integer)
            ProgressDialog.Maximum = dt.Rows.Count * pKeys.Count + adjust
            ProgressDialog.Show(Me)

            '個別結果表範囲チェックリスト取得
            Dim lst As List(Of Dictionary(Of String, String)) = Me.GetKobetsuKekkahyoCheckListRange(pKeys, dt, ProgressDialog)

            If lst.Count = 0 Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_I_007, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '個別結果表範囲チェックリスト出力
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New BRA3620R(_chosaNen, lst, filePath)
                    ret = ExcelOutput.Execute(MessageID.MSG_Q_004, ProgressDialog)
                End Using

                '進捗を進める
                ProgressDialog.AddValue = adjust

                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()

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
            If Not ProgressDialog Is Nothing Then
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                ProgressDialog = Nothing
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 追加チェックボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAddCheck_Click(sender As Object, e As EventArgs) Handles btnAddCheck.Click
        Dim errCheckType As ComConst.エラーチェック種別.enm
        Dim fileName As String
        Dim ProgressDialog As New ProgressDialog()

        Try
            errCheckType = ComConst.エラーチェック種別.enm.追加

            Dim pKeys As List(Of DAOKobetsuKekkahyo.PrimaryKey) = Me.GetKobetsuKekkahyoPrimaryKey(_chosaNen)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(pKeys, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim dt As DataTable

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '個別結果表審査論理取得
                ' REV_002↓
                'dt = DAOOther.GetKobetsuKekkahyoShinsaRonri(db, errCheckType)
                dt = DAOOther.GetKobetsuKekkahyoShinsaRonri(db, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun), errCheckType)
                ' REV_002↑
            End Using

            If dt.Rows.Count = 0 Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_026, MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            'REV_005
            If CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計 Then
                '対象客体が1件の場合はファイル名にセンサス番号を付与
                If pKeys.Count = 1 Then
                    fileName = String.Format(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.reportName, ComConst.エラーチェック種別.リスト(errCheckType)) & "_" _
                                & CommonInfo.ChosakubunName & "_" _
                                & CommonInfo.Kyoku & "_" _
                                & CommonInfo.Jimusyo & "_" _
                                & CommonInfo.Center & "_" _
                                & pKeys(0).censusNo & ".xlsx"
                Else
                    fileName = String.Format(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.reportName, ComConst.エラーチェック種別.リスト(errCheckType)) & "_" _
                                & CommonInfo.ChosakubunName & "_" _
                                & CommonInfo.Kyoku & "_" _
                                & CommonInfo.Jimusyo & "_" _
                                & CommonInfo.Center & ".xlsx"
                End If
            Else
                '対象客体が1件の場合はファイル名にセンサス番号を付与
                If pKeys.Count = 1 Then
                    fileName = String.Format(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.reportName, ComConst.エラーチェック種別.リスト(errCheckType)) & "_" _
                                & CommonInfo.ChosakubunName & "_" _
                                & CommonInfo.Kyoku & "_" _
                                & CommonInfo.Jimusyo & "_" _
                                & CommonInfo.Center & "_" _
                                & pKeys(0).censusNo & "_" & _chosaNen & "_" & Now.ToString("yyyyMMddHHmm") & ".xlsx"
                Else
                    fileName = String.Format(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.reportName, ComConst.エラーチェック種別.リスト(errCheckType)) & "_" _
                                & CommonInfo.ChosakubunName & "_" _
                                & CommonInfo.Kyoku & "_" _
                                & CommonInfo.Jimusyo & "_" _
                                & CommonInfo.Center & "_" & _chosaNen & "_" & Now.ToString("yyyyMMddHHmm") & ".xlsx"
                End If
            End If
            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath), fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '進捗ダイアログを表示する
            Dim adjust As Integer = CType(Math.Ceiling(dt.Rows.Count * pKeys.Count * 0.05), Integer)
            ProgressDialog.Maximum = dt.Rows.Count * pKeys.Count + adjust
            ProgressDialog.Show(Me)

            '個別結果表チェックリスト(基本・追加)取得
            Dim lst As List(Of Dictionary(Of String, String)) = Me.GetKobetsuKekkahyoCheckList(errCheckType, ProgressDialog)

            If lst.Count = 0 Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_I_007, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '個別結果表チェックリスト出力を行う
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New BRA3610R(errCheckType, _chosaNen, lst, filePath)
                    ret = ExcelOutput.Execute(MessageID.MSG_Q_004, ProgressDialog)
                End Using

                '進捗を進める
                ProgressDialog.AddValue = adjust

                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()

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

        Catch ex As ShinsaException
            '進捗ダイアログを閉じる
            ProgressDialog.endDispose()
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_056, {ex.CensusNo, ex.ErrSign}, MsgBoxStyle.OkOnly)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            If Not ProgressDialog Is Nothing Then
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                ProgressDialog = Nothing
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 個別結果表主キー取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetKobetsuKekkahyoPrimaryKey(chosaNen As String) As List(Of DAOKobetsuKekkahyo.PrimaryKey)
        Dim ret As New List(Of DAOKobetsuKekkahyo.PrimaryKey)

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                Dim pkey As DAOKobetsuKekkahyo.PrimaryKey = New DAOKobetsuKekkahyo.PrimaryKey(chosaNen, dgvList.Rows(i).Cells(7).Value.ToString)
                ret.Add(pkey)
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
            dt = DAOKobetsuKekkahyo.GetList(db, chosaNen, kyoku, jimusho, kyoten, einouRuikei)
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
            dgvList.Rows(i).Cells(8).Value = row("農政局").ToString
            dgvList.Rows(i).Cells(9).Value = row("都道府県").ToString
            dgvList.Rows(i).Cells(10).Value = row("実査設置拠点").ToString
        Next
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="pKey"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(pKeys As List(Of DAOKobetsuKekkahyo.PrimaryKey), ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'センサス番号選択チェック
        If pKeys.Count = 0 Then
            msgId = MessageID.MSG_E_003
            Return ret
        End If

        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' 個別結果表範囲チェックリスト取得
    ''' </summary>
    ''' <param name="pKeys"></param>
    ''' <param name="dtRonri"></param>
    ''' <param name="progressDialog"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetKobetsuKekkahyoCheckListRange(pKeys As List(Of DAOKobetsuKekkahyo.PrimaryKey), dtRonri As DataTable, ByVal progressDialog As ProgressDialog) As List(Of Dictionary(Of String, String))
        Dim ret As New List(Of Dictionary(Of String, String))

        Dim cnt As Integer

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '個別結果表項目マスタ取得
            ' REV_002↓
            'Dim dtItem As DataTable = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun)
            Dim dtItem As DataTable = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
            ' REV_002↑

            For Each pKey As DAOKobetsuKekkahyo.PrimaryKey In pKeys
                '個別結果表テーブル取得
                Dim dtKobetsu As Dictionary(Of String, DataTable) = DAOKobetsuKekkahyo.GetTable(db, pKey)
                'REV_003↓
                Dim pKeyPrev = New DAOKobetsuKekkahyo.PrimaryKey((CInt(pKey.chosaNen) - 1).ToString, pKey.censusNo)
                Dim dtKobetsuPrev As Dictionary(Of String, DataTable) = DAOKobetsuKekkahyo.GetTable(db, pKeyPrev)
                'REV_003↑

                '個別結果表項目取得
                Dim dcKobetsu As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目) = ComUtil.KobetsuKekkahyo.GetItem(dtItem, dtKobetsu)
                'REV_003↓
                Dim dcKobetsuPrev As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目) = ComUtil.KobetsuKekkahyo.GetItem(dtItem, dtKobetsuPrev)
                'REV_003↑

                For Each row As DataRow In dtRonri.Rows

                    Dim val1 As Decimal
                    Dim val2 As Decimal
                    Dim val3 As Decimal

                    Dim tanshu As Decimal

                    'REV_003↓
                    'If Decimal.TryParse(dcKobetsu(row("項目番号１").ToString).値, val1) _
                    '    AndAlso Decimal.TryParse(dcKobetsu(row("項目番号２").ToString).値, val2) _
                    '    AndAlso Decimal.TryParse(row("値").ToString, val3) Then
                    Dim itemNo1 = row("項目番号１").ToString
                    Dim itemNo2 = row("項目番号２").ToString
                    '前付項番で前年データがない場合はスキップ
                    If (itemNo1.StartsWith("前") OrElse itemNo2.StartsWith("前")) _
                        AndAlso dcKobetsuPrev.Count = 0 Then
                        '進捗を進める
                        progressDialog.AddValue = 1
                        Continue For
                    End If
                    Dim strVal1 = If(itemNo1.StartsWith("前"), dcKobetsuPrev, dcKobetsu)(itemNo1.Replace("前", "")).値
                    Dim strVal2 = If(itemNo2.StartsWith("前"), dcKobetsuPrev, dcKobetsu)(itemNo2.Replace("前", "")).値
                    If Decimal.TryParse(strVal1, val1) _
                        AndAlso Decimal.TryParse(strVal2, val2) _
                        AndAlso Decimal.TryParse(row("値").ToString, val3) Then
                        'REV_003↑

                        ' --- REV_001 MOD START
                        'If val1 > 0 And val2 > 0 Then   
                        If val1 <> 0 Then
                            ' --- REV_001 MOD END

                            tanshu = val2 / val1 * val3

                            Dim low As Decimal = Decimal.Parse(row("下限").ToString)
                            Dim up As Decimal = Decimal.Parse(row("上限").ToString)

                            If Not (low <= tanshu And tanshu <= up) Then
                                Dim dc As New Dictionary(Of String, String)

                                cnt = cnt + 1

                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(1)) = cnt.ToString
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(2)) = ComUtil.GetTodofuken(pKey.censusNo)
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(3)) = ComUtil.GetShikuchoson(pKey.censusNo)
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(4)) = ComUtil.GetKyuShikuchoson(pKey.censusNo)
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(5)) = ComUtil.GetNogyoShuraku(pKey.censusNo)
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(6)) = ComUtil.GetChosaku(pKey.censusNo)
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(7)) = ComUtil.GetKyakutaiNo(pKey.censusNo)
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(8)) = row("チェック項目名").ToString
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(9)) = row("下限").ToString
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(10)) = row("上限").ToString
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(11)) = row("値").ToString
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(12)) = row("項目番号１").ToString
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(13)) = row("項目番号２").ToString
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(14)) = val1.ToString
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(15)) = val2.ToString
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(16)) = tanshu.ToString
                                '--- REV.001 ADD START
                                dc(ComConst.個別結果表範囲チェックリスト.出力用ファイル名称.Field(17)) = String.Empty
                                '--- REV.001 ADD END
                                ret.Add(dc)
                            End If
                        End If
                    End If

                    '進捗を進める
                    progressDialog.AddValue = 1
                Next
            Next
        End Using

        Return ret
    End Function

    ''' <summary>
    ''' 個別結果表チェックリスト(基本・追加)取得
    ''' </summary>
    ''' <param name="pErrCheckType"></param>
    ''' <param name="progressDialog"></param>
    ''' <remarks></remarks>
    Private Function GetKobetsuKekkahyoCheckList(ByVal pErrCheckType As ComConst.エラーチェック種別.enm, ByVal progressDialog As ProgressDialog) As List(Of Dictionary(Of String, String))
        Dim censusNoList As List(Of String)
        Dim ret As List(Of Dictionary(Of String, String))

        Try

            'センサス番号リスト作成
            censusNoList = New List(Of String)
            For i As Integer = 0 To dgvList.Rows.Count - 1
                If Not Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                    Continue For
                End If
                censusNoList.Add(dgvList.Rows(i).Cells(7).Value.ToString)
            Next
            Dim shinsa As Shinsa = New Shinsa(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei),
                                              CommonInfo.Chosakubun,
                                              ComConst.審査論理データ種別.個別結果表,
                                              If(pErrCheckType = ComConst.エラーチェック種別.enm.基本, ComConst.審査論理種別.基本チェック, ComConst.審査論理種別.追加チェック),
                                              _chosaNen,
                                              censusNoList,
                                              progressDialog)

            '審査実行
            Dim shinsaKekkaList As List(Of Shinsa.ShinsaKekka) = shinsa.Execute()

            ret = New List(Of Dictionary(Of String, String))
            Dim cnt As Integer = 0
            For Each kekka In shinsaKekkaList
                'エラーが存在しない場合
                If Not (From n In kekka.ShinsaInfoList Where n.Result = False).Any Then
                    Continue For
                End If

                For Each info In kekka.ShinsaInfoList
                    'エラーなしの場合
                    If info.Result Then
                        Continue For
                    End If

                    Dim dc As New Dictionary(Of String, String)

                    cnt = cnt + 1

                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(1)) = cnt.ToString
                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(2)) = ComUtil.GetTodofuken(kekka.CensusNo)
                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(3)) = ComUtil.GetShikuchoson(kekka.CensusNo)
                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(4)) = ComUtil.GetKyuShikuchoson(kekka.CensusNo)
                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(5)) = ComUtil.GetNogyoShuraku(kekka.CensusNo)
                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(6)) = ComUtil.GetChosaku(kekka.CensusNo)
                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(7)) = ComUtil.GetKyakutaiNo(kekka.CensusNo)
                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(8)) = info.ErrSign
                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(9)) = info.CheckItemName
                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(10)) = info.ErrNaiyo
                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(11)) = info.ErrJoken
                    '--- REV.001 MOD START
                    'dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(12)) = info.ErrKubun
                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(12)) = String.Empty
                    dc(ComConst.個別結果表基本追加チェックリスト.出力用ファイル名称.Field(13)) = info.ErrKubun
                    '--- REV.001 MOD END
                    ret.Add(dc)
                Next
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret
    End Function
End Class
