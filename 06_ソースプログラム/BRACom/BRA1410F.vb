'//*************************************************************************************************
'//  修正履歴
'// ------------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前           |                修  正  内  容
'// -----------+------------+------------------------+----------------------------------------------
'//  REV_001   | 2020.11.24 |TSP)                    | フェーズ3 追加要件No.12修正
'//  REV_002   | 2021.12.07 |日本コンピュータシステム| フェーズ3 追加要件No.12修正
'//  REV_003   | 2021.12.23 |日本コンピュータシステム| 要件11修正
'//  REV_004   | 2023.11.27 |Daiko                   | 要件No.20 エラーリスト出力処理追加
'//  REV_005   | 2025.09.10 |GCU                     | 要件No.5 エラーチェックへの調査年産等追加
'//            |            |                        |
'//*************************************************************************************************
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Threading
Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop.Forms
''' <summary>
''' 電子調査票入力・修正画面
''' </summary>
''' <remarks></remarks>
Public Class BRA1410F

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
    ''' <summary>ロックファイル読取</summary>
    Private _lockFileStream As System.IO.FileStream 'REV_003 ADD
    ''' <summary>シート</summary>
    Private _sheet As List(Of String)
    ''' <summary>エラー時コメント改行文字数</summary>
    Private Const エラー時コメント改行文字数 As Integer = 30

    Private Const COLSOSHIKI As String = "dgcChosaKeieitai"
    Private Const COLTANTOU As String = "dgcTantoMeisho"

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA1410F_Load(sender As Object, e As EventArgs) Handles Me.Load
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
            '一部例外を調整
            For Each dgvCol As DataGridViewColumn In Me.dgvList.Columns
                If (dgvCol.Name = COLSOSHIKI Or dgvCol.Name = COLTANTOU) And Not CommonInfo.SenmonChosain Then
                    dgvCol.ReadOnly = False
                    dgvCol.SortMode = DataGridViewColumnSortMode.Automatic
                    'REV START 2022/03/07 連絡票No293　専門調査員の場合調査経営体のみ書き換え可能にする
                ElseIf (dgvCol.Name = COLSOSHIKI) And CommonInfo.SenmonChosain Then
                    dgvCol.ReadOnly = False
                    dgvCol.SortMode = DataGridViewColumnSortMode.Automatic
                    'REV END 2022/03/07 連絡票No293
                End If
            Next

            '新規作成ボタン非活性設定（農政局工程、本省工程の場合、非活性とする）
            If CommonInfo.Koutei <> CommonInfo.KouteiKubun.Code.Center Then
                btnRegister.Enabled = False
                '--- REV.001 ADD START
                '労賃単価反映ボタン非活性設定（農政局工程、本省工程の場合、非活性とする）
                BtnRouchin.Enabled = False
            Else
                '労賃単価反映ボタン非活性設定（生産費以外の場合非活性とする）
                If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                    BtnRouchin.Enabled = False
                End If

                '労賃単価反映ボタン非活性設定（専門調査員の場合非活性とする）
                If CommonInfo.SenmonChosain Then
                    BtnRouchin.Enabled = False
                    '2022/1/30 ADD START 専門調査員の場合は、検索条件の「担当名称」フィールドと「経営体・担当名称」ボタンを非活性とする
                    txtTantoshaMei.Enabled = False
                    'btnKeieiTanto.Enabled = False   'REV ADD 2022/03/07 「経営体・担当名称」ボタンは活性状態
                    '2022/1/30 ADD END
                End If
                '--- REV.001 ADD END
            End If

            'エラーリスト出力関連ボタン非活性
            btnSheet.Enabled = False
            btnBaseCheck.Enabled = False
            btnAddCheck.Enabled = False
            btnRangeCheck.Enabled = False

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 新規作成ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRegister_Click(sender As Object, e As EventArgs) Handles btnRegister.Click
        Dim xlForm As BRA1510X = Nothing

        Try
            '↓2022/01/26 pKeyの設定
            Dim pKey As DAOChosahyo.PrimaryKey
            Dim kKey As New DAOChosahyo.KotenKey
            Dim msgId As String = String.Empty

            If _chosaNen = Nothing Then
                If cboChosaNen.SelectedValue Is Nothing Then
                    _chosaNen = Date.Today.Year.ToString
                Else
                    _chosaNen = cboChosaNen.SelectedValue.ToString
                End If
            End If

            pKey = Me.GetChosahyoPrimaryKey(_chosaNen, kKey)

            If Me.CheckError(pKey, msgId) Then
                ' OK
            ElseIf Not String.IsNullOrEmpty(cboChosaNen.Text) Then
                ' 調査年コンボボックスに選択値有り
                pKey = New DAOChosahyo.PrimaryKey(cboChosaNen.Text, Nothing)
            Else
                pKey = New DAOChosahyo.PrimaryKey(Date.Today.Year.ToString, Nothing)
            End If
            '↑2022/01/26 pKeyの設定

            '↓2022/03/24　2人以上が新規作成したときのエラー回避
            Dim rnd As New Random()    ' Randomオブジェクトの作成
            pKey.censusNo = Format(Now, "yyyyMMddhhmmss") & rnd.Next(10, 99).ToString
            '↑2022/03/24　2人以上が新規作成したときのエラー回避

            '↓2022/03/24　新規作成時のロックファイル対応をコメントアウト
            '上位フォルダが存在しない場合、処理終了  'REV_003 ADD
            'Dim lockFilePath As String = ComUtil.GetLockFilePath(pKey.chosaNen, pKey.censusNo)
            'If Not System.IO.Directory.Exists(IniFileInfo.LockParentPath) Then
            '    Message.ShowMsgBox(MessageID.MSG_E_115, {CommonInfo.Koutei}, MsgBoxStyle.OkOnly)
            '    Exit Sub
            'End If

            'ロックファイルが存在する場合、処理終了  'REV_003 ADD
            'If Me.LockFileExist(pKey.chosaNen, pKey.censusNo) Then
            '    Message.ShowMsgBox(MessageID.MSG_E_105, MsgBoxStyle.OkOnly)
            '    Exit Sub
            'End If

            ' ファイルをオープン状態にする  'REV_003 ADD
            'System.IO.File.Create(lockFilePath).Close()
            '_lockFileStream = New System.IO.FileStream(lockFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Write)

            FileCopy(System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun))).tempFileName), ComUtil.GetTemplateCopyFilePath(pKey.chosaNen, pKey.censusNo))

            '↓2022/01/26 引数変更
            'REV002 START --------
            'xlForm = New BRA1510X(Me, BRA1510X.編集モード種別.新規, _chosaNen)
            xlForm = New BRA1510X(Me, BRA1510X.編集モード種別.新規, _chosaNen, pKey, _lockFileStream)
            'REV002 END --------
            '↑2022/01/26 引数変更
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)

            'ロックファイル削除・ストリーム解放  'REV_003 ADD
            'If Not _lockFileStream Is Nothing Then
            '    ComUtil.DeleteLockFile(_lockFileStream)
            'End If

            '↑2022/03/24　新規作成時のロックファイル対応をコメントアウト

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
            _einouRuikei = If(cboEinouRuikei.SelectedValue Is Nothing, Nothing, cboEinouRuikei.SelectedValue.ToString)

            '一覧表示
            Me.ShowList(_chosaNen, _kyoku, _jimusho, _kyoten, _einouRuikei)

            'エラーリスト出力関連初期化
            Dim sheetList As New List(Of String)
            If CommonInfo.Kubun1 = ComConst.区分１.農業経営統計調査 AndAlso ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2020 Then
                btnSheet.Enabled = True
                For Each kv As KeyValuePair(Of String, String) In ComConst.調査票.シートデータ範囲(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun)))
                    sheetList.Add(kv.Key)
                Next
            Else
                btnSheet.Enabled = False
            End If

            SetSheet(sheetList)

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

    Private Sub btnModify_Click(sender As Object, e As EventArgs) Handles btnModify.Click
        Dim xlForm As BRA1510X = Nothing

        '>>>2022/01/25 TryCatch内へ移動
        ''FileCopy(System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun))).tempFileName), System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), "C" & ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun))).tempFileName))
        '<<<

        Try
            Dim pKey As DAOChosahyo.PrimaryKey
            Dim kKey As New DAOChosahyo.KotenKey

            pKey = Me.GetChosahyoPrimaryKey(_chosaNen, kKey)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(pKey, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '上位フォルダが存在しない場合、処理終了  'REV_003 ADD
            If Not System.IO.Directory.Exists(IniFileInfo.LockParentPath) Then
                Message.ShowMsgBox(MessageID.MSG_E_115, {CommonInfo.Koutei}, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            'ロックファイルが存在する場合、処理終了  'REV_003 ADD
            Dim lockFilePath As String = ComUtil.GetLockFilePath(pKey.chosaNen, pKey.censusNo)
            If Me.LockFileExist(pKey.chosaNen, pKey.censusNo) Then
                Message.ShowMsgBox(MessageID.MSG_E_105, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            ' ファイルをオープン状態にする  'REV_003 ADD
            System.IO.File.Create(lockFilePath).Close()
            _lockFileStream = New System.IO.FileStream(lockFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Write)

            '>>>2022/01/25
            '調査票テンプレートからワークへコピー
            FileCopy(System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun))).tempFileName), ComUtil.GetTemplateCopyFilePath(pKey.chosaNen, pKey.censusNo))
            '<<<2022/01/25

            xlForm = New BRA1510X(Me, BRA1510X.編集モード種別.修正, pKey, kKey, _lockFileStream)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)

            'ロックファイル削除・ストリーム解放  'REV_003 ADD
            If Not _lockFileStream Is Nothing Then
                ComUtil.DeleteLockFile(_lockFileStream)
            End If

            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub dgvList_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvList.CellValueChanged
        Dim checkCount As Integer = 0
        For rowIndex = 0 To dgvList.Rows.Count - 1
            If CType(dgvList(0, rowIndex).Value, Boolean) = True Then
                checkCount = checkCount + 1
            End If
        Next
        If checkCount > 1 Then
            btnModify.Enabled = False
            btnDelete.Enabled = False
        Else
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If
    End Sub

    Private Sub dgvList_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvList.CurrentCellDirtyStateChanged
        If dgvList.CurrentCellAddress.X = 0 AndAlso dgvList.IsCurrentCellDirty Then
            dgvList.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    ''' <summary>
    ''' 全選択ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAllSelect_Click(sender As Object, e As EventArgs) Handles btnAllSelect.Click
        Try
            For i As Integer = 0 To dgvList.Rows.Count - 1
                If dgvList(0, i).ReadOnly = False Then
                    dgvList(0, i).Value = True
                End If
            Next
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
            For i As Integer = 0 To dgvList.Rows.Count - 1
                If dgvList(0, i).ReadOnly = False Then
                    dgvList(0, i).Value = False
                End If
            Next
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            Dim pKey As DAOChosahyo.PrimaryKey
            Dim kKey As New DAOChosahyo.KotenKey

            pKey = Me.GetChosahyoPrimaryKey(_chosaNen, kKey)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(pKey, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '上位フォルダが存在しない場合、処理終了  'REV_003 ADD
            If Not System.IO.Directory.Exists(IniFileInfo.LockParentPath) Then
                Message.ShowMsgBox(MessageID.MSG_E_115, {CommonInfo.Koutei}, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            'ロックファイルが存在する場合、処理終了  'REV_003 ADD
            If Me.LockFileExist(pKey.chosaNen, pKey.censusNo) Then
                Message.ShowMsgBox(MessageID.MSG_E_112, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_003, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                    '個別結果表存在チェック
                    Dim pKobet As DAOKobetsuKekkahyo.PrimaryKey = New DAOKobetsuKekkahyo.PrimaryKey(pKey.chosaNen, pKey.censusNo)
                    Dim kKobet As DAOKobetsuKekkahyo.KyotenKey = New DAOKobetsuKekkahyo.KyotenKey(kKey.kyoku, kKey.jimusho, kKey.kyoten)
                    Dim bln As Boolean = DAOKobetsuKekkahyo.CheckExist(db, pKobet, kKobet)

                    If bln Then
                        '確認メッセージ
                        If Message.ShowMsgBox(MessageID.MSG_Q_006, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If

                    Try

                        db.BeginTrans()

                        '調査票データ削除
                        DAOChosahyo.DeleteChosahyoTable(db, pKey, kKey)

                        If bln Then
                            '個別結果表データ削除
                            DAOKobetsuKekkahyo.DeleteTable(db, pKobet, kKobet)
                        End If

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
    ''' 調査票主キー取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="kKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChosahyoPrimaryKey(chosaNen As String, Optional ByRef kKey As DAOChosahyo.KotenKey = Nothing) As DAOChosahyo.PrimaryKey
        Dim ret As DAOChosahyo.PrimaryKey = Nothing

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                ret = New DAOChosahyo.PrimaryKey(chosaNen, dgvList.Rows(i).Cells(7).Value.ToString)
                If Not kKey Is Nothing Then
                    kKey.kyoku = dgvList.Rows(i).Cells(9).Value.ToString
                    kKey.jimusho = dgvList.Rows(i).Cells(10).Value.ToString
                    kKey.kyoten = dgvList.Rows(i).Cells(11).Value.ToString
                End If
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 調査票主キー取得（複数）
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChosahyoPrimaryKeys(chosaNen As String) As List(Of DAOChosahyo.PrimaryKey)
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
        Dim senmonName As String = ""

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dt = DAOChosahyo.GetChosahyoList(db, chosaNen, kyoku, jimusho, kyoten, einouRuikei, txtChosaKeieitai.Text, txtTantoshaMei.Text)
            senmonName = DAOChosahyo.GetSenmonChosainName(db)
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
    ''' <param name="pKey"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(pKey As DAOChosahyo.PrimaryKey, ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'センサス番号選択チェック
        If pKey Is Nothing Then
            msgId = MessageID.MSG_E_003
            Return ret
        End If

        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="pKeys"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(pKeys As List(Of DAOChosahyo.PrimaryKey), ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'センサス番号選択チェック
        If pKeys.Count = 0 Then
            msgId = MessageID.MSG_E_003
            Return ret
        End If

        ret = True

        Return ret
    End Function

    '--- REV.001 ADD START
    Private Sub BtnRouchin_Click(sender As Object, e As EventArgs) Handles BtnRouchin.Click
        Dim dt As DataTable
        Dim seisanhiKbn As String

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


        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            dt = DAOChosahyo.GetChosahyoList(db, _chosaNen, _kyoku, _jimusho, _kyoten, Nothing, txtChosaKeieitai.Text, txtTantoshaMei.Text)
        End Using

        'エラーチェック
        '対象客体チェック
        If dt.Rows.Count = 0 Then
            Message.ShowMsgBox(MessageID.MSG_E_068, MsgBoxStyle.OkOnly)
            Exit Sub
        Else
            '処理続行確認
            If Message.ShowMsgBox(MessageID.MSG_Q_036, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If


        seisanhiKbn = ComUtil.RouchinTanka.GetSeisanhiKbn(CommonInfo.Chosakubun)

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            DAOChosahyo.setRouchinTanka(db, seisanhiKbn, _chosaNen, _kyoku, _jimusho, _kyoten)
        End Using

        '完了メッセージ
        Message.ShowMsgBox(MessageID.MSG_I_045, MsgBoxStyle.OkOnly)
    End Sub

    '--- REV.001 ADD END

    ''' <summary>
    ''' ロックファイル存在チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>REV_003 ADD</remarks>
    Private Function LockFileExist(chosaNen As String, censusNo As String) As Boolean
        Dim ret As Boolean = False
        Dim lockDirectoryPath = IniFileInfo.LockPath

        'ロックファイル用フォルダが存在しない場合、作成する
        If Not System.IO.Directory.Exists(lockDirectoryPath) Then
            System.IO.Directory.CreateDirectory(lockDirectoryPath)
        End If

        'ロックファイルが存在する場合、メッセージ表示→処理終了
        If System.IO.File.Exists(ComUtil.GetLockFilePath(chosaNen, censusNo)) Then
            ret = True
        End If

        Return ret
    End Function

    Private Sub btnKeieiTanto_Click(sender As Object, e As EventArgs) Handles btnKeieiTanto.Click
        Try
            Dim pKey As DAOChosahyo.PrimaryKey = Nothing
            Dim kKey As DAOChosahyo.KotenKey = Nothing

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_001, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    For i As Integer = 0 To dgvList.Rows.Count - 1
                        pKey = New DAOChosahyo.PrimaryKey(_chosaNen, dgvList.Rows(i).Cells(7).Value.ToString)
                        If Not kKey Is Nothing Then
                            kKey.kyoku = dgvList.Rows(i).Cells(9).Value.ToString
                            kKey.jimusho = dgvList.Rows(i).Cells(10).Value.ToString
                            kKey.kyoten = dgvList.Rows(i).Cells(11).Value.ToString
                        End If
                        Try
                            db.BeginTrans()

                            Dim chosaSoshiki As String
                            If dgvList.Rows(i).Cells(12).Value Is Nothing Then
                                chosaSoshiki = ""
                            Else
                                chosaSoshiki = dgvList.Rows(i).Cells(12).Value.ToString
                            End If

                            Dim tantoMeisho As String
                            If dgvList.Rows(i).Cells(13).Value Is Nothing Then
                                tantoMeisho = ""
                            Else
                                tantoMeisho = dgvList.Rows(i).Cells(13).Value.ToString
                            End If

                            '調査票データ更新
                            DAOChosahyo.UpdateChosahyoTable_SoshikiTanto(db, pKey, kKey, chosaSoshiki, tantoMeisho)

                            db.CommitTrans()
                        Catch ex As Exception
                            db.RollBackTrans()
                            Throw ex
                        End Try
                    Next
                End Using

                '完了メッセージ
                Message.ShowMsgBox(MessageID.MSG_I_001, MsgBoxStyle.OkOnly)

                '一覧表示
                Me.ShowList(_chosaNen, _kyoku, _jimusho, _kyoten, _einouRuikei)

            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub txtChosaKeieitai_Click(sender As Object, e As EventArgs) Handles txtChosaKeieitai.Click
        txtChosaKeieitai.ReadOnly = False
    End Sub

    Private Sub txtTantoshaMei_Click(sender As Object, e As EventArgs) Handles txtTantoshaMei.Click
        '2022/1/30 ADD START 専門調査員の場合は、検索条件の「担当名称」フィールドをクリックしても入力できないようにする IF文追加
        If CommonInfo.SenmonChosain = False Then
            '2022/1/30 ADD END
            txtTantoshaMei.ReadOnly = False
            '2022/1/30 ADD START 専門調査員の場合は、検索条件の「担当名称」フィールドをクリックしても入力できないようにする IFと対となるENDを追加
        End If
        '2022/1/30 ADD END
    End Sub

    Private Sub btnSheet_Click(sender As Object, e As EventArgs) Handles btnSheet.Click
        Try
            Dim frm As New BRA1420F(_chosaNen)
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' シート設定
    ''' </summary>
    ''' <param name="sheetList"></param>
    ''' <remarks></remarks>
    Public Sub SetSheet(sheetList As List(Of String))
        _sheet = sheetList
        txtSheet.Text = String.Empty

        For Each sheet As String In sheetList
            If txtSheet.Text = String.Empty Then
                txtSheet.Text = sheet
            Else
                txtSheet.Text &= Environment.NewLine
                txtSheet.Text &= sheet
            End If
        Next

        If _sheet.Count > 0 Then
            btnBaseCheck.Enabled = True
            btnAddCheck.Enabled = True
            btnRangeCheck.Enabled = True
        Else
            btnBaseCheck.Enabled = False
            btnAddCheck.Enabled = False
            btnRangeCheck.Enabled = False
        End If

    End Sub

    ''' <summary>
    ''' チェックボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBaseAddCheck_Click(sender As Object, e As EventArgs) Handles btnBaseCheck.Click, btnAddCheck.Click, btnRangeCheck.Click
        Dim errCheckType As ComConst.エラーチェック種別.enm
        Dim dtCheckRonri As DataTable
        Dim fileName As String
        Dim ProgressDialog As New ProgressDialog()

        Try
            Dim button = CType(sender, Button)
            If button.Name = btnBaseCheck.Name Then
                errCheckType = ComConst.エラーチェック種別.enm.基本
            ElseIf button.Name = btnAddCheck.Name Then
                errCheckType = ComConst.エラーチェック種別.enm.追加
            Else
                errCheckType = ComConst.エラーチェック種別.enm.範囲
            End If

            Dim pKeys As List(Of DAOChosahyo.PrimaryKey) = GetChosahyoPrimaryKeys(_chosaNen)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not CheckError(pKeys, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '審査論理取得
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                If errCheckType = ComConst.エラーチェック種別.enm.範囲 Then
                    dtCheckRonri = DAOOther.GetChosahyoShinsaRonriRange(db)
                Else
                    dtCheckRonri = DAOOther.GetChosahyoShinsaRonri(db, errCheckType)
                End If
                If dtCheckRonri.Rows.Count = 0 Then
                    '審査論理データ無し
                    If errCheckType = ComConst.エラーチェック種別.enm.基本 Then
                        Message.ShowMsgBox(MessageID.MSG_E_025, MsgBoxStyle.OkOnly)
                    ElseIf errCheckType = ComConst.エラーチェック種別.enm.追加 Then
                        Message.ShowMsgBox(MessageID.MSG_E_026, MsgBoxStyle.OkOnly)
                    Else
                        Message.ShowMsgBox(MessageID.MSG_E_102, MsgBoxStyle.OkOnly)
                    End If
                    Exit Sub
                End If
            End Using
            'REV_005↓
            If CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計 Then
                If pKeys.Count = 1 Then
                    fileName = ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.reportName(errCheckType) & "_" _
                                    & CommonInfo.ChosakubunName & "_" & CommonInfo.Kyoku & "_" & CommonInfo.Jimusyo & "_" & CommonInfo.Center & "_" & pKeys(0).censusNo & ".xlsx"
                Else
                    fileName = ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.reportName(errCheckType) & "_" _
                                    & CommonInfo.ChosakubunName & "_" & CommonInfo.Kyoku & "_" & CommonInfo.Jimusyo & "_" & CommonInfo.Center & ".xlsx"
                End If
            Else
                If pKeys.Count = 1 Then
                    fileName = ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.reportName(errCheckType) & "_" _
                & CommonInfo.ChosakubunName & "_" & CommonInfo.Kyoku & "_" & CommonInfo.Jimusyo & "_" & CommonInfo.Center & "_" & pKeys(0).censusNo _
                & "_" & _chosaNen & "_" & Now.ToString("yyyyMMddHHmm") & ".xlsx"
                Else
                    fileName = ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Report.reportName(errCheckType) & "_" _
                & CommonInfo.ChosakubunName & "_" & CommonInfo.Kyoku & "_" & CommonInfo.Jimusyo & "_" & CommonInfo.Center _
                & "_" & _chosaNen & "_" & Now.ToString("yyyyMMddHHmm") & ".xlsx"
                End If

            End If
            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath), fileName)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            Cursor.Current = Cursors.WaitCursor

            '進捗ダイアログを表示する
            Dim adjust As Integer = CType(Math.Ceiling(dtCheckRonri.Rows.Count * pKeys.Count * 0.05), Integer)
            ProgressDialog.Maximum = dtCheckRonri.Rows.Count * pKeys.Count + adjust
            ProgressDialog.Show(Me)

            Dim lst As List(Of Dictionary(Of String, String))
            '調査票チェックリスト(基本・追加)取得
            If errCheckType = ComConst.エラーチェック種別.enm.範囲 Then
                lst = GetChosahyoCheckListRange(pKeys, dtCheckRonri, ProgressDialog)
            Else
                lst = GetChosahyoCheckList(errCheckType, ProgressDialog)
            End If

            If lst.Count = 0 Then
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                Message.ShowMsgBox(MessageID.MSG_I_056, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '調査票チェックリスト出力を行う
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New BRA1410R(errCheckType, _chosaNen, lst, filePath, _sheet)
                    ret = ExcelOutput.Execute(MessageID.MSG_Q_004, ProgressDialog)
                End Using

                '進捗を進める
                ProgressDialog.AddValue = adjust

                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()

                If ret = ExcelOutputBaseClass.enmOutputReturn.OK Then
                    '完了メッセージ
                    If lst.Count > ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Row.Max Then
                        Message.ShowMsgBox(MessageID.MSG_W_002, {CStr(lst.Count)}, MsgBoxStyle.OkOnly)
                    Else
                        Message.ShowMsgBox(MessageID.MSG_I_002, MsgBoxStyle.OkOnly)
                    End If
                End If
            Catch ex As ExcelOutputBaseClass.SaveAsException
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_006, MsgBoxStyle.OkOnly)
            Catch ex As Exception
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                Throw
            End Try

        Catch ex As ShinsaException
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            '進捗ダイアログを閉じる
            ProgressDialog.endDispose()
            Message.ShowMsgBox(MessageID.MSG_E_160, {ex.CensusNo, ex.ErrSign}, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            '進捗ダイアログを閉じる
            ProgressDialog.endDispose()
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)

        Finally
            If Not ProgressDialog Is Nothing Then
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                ProgressDialog = Nothing
            End If

            Cursor.Current = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 調査票チェックリスト(基本・追加)取得
    ''' </summary>
    ''' <param name="pErrCheckType"></param>
    ''' <param name="progressDialog"></param>
    ''' <remarks></remarks>
    Private Function GetChosahyoCheckList(pErrCheckType As ComConst.エラーチェック種別.enm, progressDialog As ProgressDialog) As List(Of Dictionary(Of String, String))
        Dim dtChoItemMst As DataTable
        Dim censusNoList As List(Of String)
        Dim errCheckList As List(Of Dictionary(Of String, String)) = New List(Of Dictionary(Of String, String))

        Try
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '調査票項目マスタ取得
                dtChoItemMst = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)
            End Using

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
                                              ComConst.審査論理データ種別.調査票,
                                              If(pErrCheckType = ComConst.エラーチェック種別.enm.基本, ComConst.審査論理種別.基本チェック, ComConst.審査論理種別.追加チェック),
                                              _chosaNen,
                                              censusNoList,
                                              progressDialog,
                                              True)

            '審査実行
            Dim shinsaKekkaList As List(Of Shinsa.ShinsaKekka) = shinsa.Execute()

            Dim cnt As Integer = 0
            For Each shinsaKekka In shinsaKekkaList
                For Each shinsaInfo In From n In shinsaKekka.ShinsaInfoList Where n.Result = False
                    Dim itemNoList As New List(Of String)
                    '「エラーとなる条件(繰り返し項目増加したもの)」から『[』『]』で囲まれた文字列検索する
                    Dim matchList As MatchCollection = Regex.Matches(shinsaInfo.ErrJokenConv, Shinsa.C_SearchKakko)
                    For Each mat As Match In matchList
                        Dim itemNo As String = mat.Value.Trim("["c, "]"c)
                        '括弧なし項目番号の重複しないリストを作成する
                        If Not itemNoList.Contains(itemNo) Then
                            itemNoList.Add(itemNo)
                        End If
                    Next

                    'コメントリスト・エラーチェック一覧出力用データを作成
                    Dim errSheetList As List(Of String) = New List(Of String)
                    Dim changeableRowList As List(Of String) = New List(Of String)
                    For Each itemNo In itemNoList
                        Dim meisaiNo As Integer = 0
                        Dim drChoItemMst As DataRow
                        If itemNo.Contains(ComConst.ITEM_NO_DELIMITER) Then
                            drChoItemMst = (From n In dtChoItemMst.AsEnumerable Where n("項目番号").ToString = itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(0) Select n).FirstOrDefault
                            meisaiNo = CType(itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(1), Integer)
                        Else
                            drChoItemMst = (From n In dtChoItemMst.AsEnumerable Where n("項目番号").ToString = itemNo Select n).FirstOrDefault
                        End If
                        If drChoItemMst Is Nothing Then
                            Continue For
                        End If

                        errSheetList.Add(drChoItemMst("シート名").ToString)

                        Dim row As Integer
                        Dim Col As Integer
                        If drChoItemMst("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                            row = CInt(drChoItemMst("行位置"))
                            Col = CInt(drChoItemMst("列位置"))
                            changeableRowList.Add("")
                        Else
                            'count、sumの場合
                            If meisaiNo = 0 Then
                                row = CInt(drChoItemMst("行位置"))
                                Col = CInt(drChoItemMst("列位置"))
                            Else
                                row = CInt(drChoItemMst("行位置")) + If(drChoItemMst("可変方向").ToString = ComConst.可変方向.縦, (meisaiNo - 1) * CInt(drChoItemMst("可変増量")), 0)
                                Col = CInt(drChoItemMst("列位置")) + If(drChoItemMst("可変方向").ToString = ComConst.可変方向.横, (meisaiNo - 1) * CInt(drChoItemMst("可変増量")), 0)
                            End If
                            changeableRowList.Add(row.ToString())
                        End If
                    Next

                    'エラー出力対象シートか判定
                    Dim blnOut As Boolean = False
                    For Each errSheet In errSheetList
                        If _sheet.Contains(errSheet) Then
                            blnOut = True
                            Exit For
                        End If
                    Next
                    If blnOut = False Then
                        Continue For
                    End If

                    '出力データ設定
                    Dim errCheckDictionary As Dictionary(Of String, String) = New Dictionary(Of String, String)
                    cnt = cnt + 1
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.No), cnt.ToString)
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.都道府県), ComUtil.GetTodofuken(shinsaKekka.CensusNo))
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.市区町村), ComUtil.GetShikuchoson(shinsaKekka.CensusNo))
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.旧市区町村), ComUtil.GetKyuShikuchoson(shinsaKekka.CensusNo))
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.農業集落), ComUtil.GetNogyoShuraku(shinsaKekka.CensusNo))
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.調査区), ComUtil.GetChosaku(shinsaKekka.CensusNo))
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.客体番号), ComUtil.GetKyakutaiNo(shinsaKekka.CensusNo))
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーシート名), String.Join(ComConst.調査票エラーチェックリスト一覧複数客体.delimiter, errSheetList.ToArray()))
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.項目番号), String.Join(ComConst.調査票エラーチェックリスト一覧複数客体.delimiter, itemNoList.ToArray()))
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.可変行番号),
                                           If(changeableRowList.Count = changeableRowList.Where(Function(changeablerow) changeablerow.Equals("")).Count, "", String.Join(ComConst.調査票エラーチェックリスト一覧複数客体.delimiter, changeableRowList.ToArray())))
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーサイン), shinsaInfo.ErrSign)
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラー内容), AddStrCRLF(shinsaInfo.ErrNaiyo, エラー時コメント改行文字数))
                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーチェック種別), ComConst.エラーチェック種別.リスト(pErrCheckType) & shinsaInfo.ErrKubun)
                    errCheckList.Add(errCheckDictionary)
                Next
            Next

        Catch ex As Exception
            Throw ex
        End Try

        Return errCheckList
    End Function

    ''' <summary>
    ''' 文字列に指定された文字数ごとに改行コードを挿入する
    ''' </summary>
    ''' <param name="myString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AddStrCRLF(myString As String, myLength As Integer) As String
        Dim ret As Text.StringBuilder

        Try

            ret = New Text.StringBuilder

            If myLength = 0 Then
                ret.Append(myString)
                Exit Try
            End If

            For i As Integer = 0 To myString.Length - 1
                If i <> 0 AndAlso (i Mod myLength) = 0 Then
                    ret.Append(vbLf)
                End If

                ret.Append(myString(i))
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret.ToString
    End Function

    ''' <summary>
    ''' 調査票範囲チェックリスト取得
    ''' </summary>
    ''' <param name="pErrCheckType"></param>
    ''' <param name="progressDialog"></param>
    ''' <remarks></remarks>
    Private Function GetChosahyoCheckListRange(pKeys As List(Of DAOChosahyo.PrimaryKey), dtRonri As DataTable, ByVal progressDialog As ProgressDialog) As List(Of Dictionary(Of String, String))
        Dim dtChoItemMst As DataTable
        Dim dtChosahyo As Dictionary(Of String, DataTable)
        Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)
        Dim errCheckList As List(Of Dictionary(Of String, String)) = New List(Of Dictionary(Of String, String))
        Dim koumoku_mojisuu As Integer = 9
        Dim cnt As Integer = 0

        Try
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '調査票項目マスタ取得
                dtChoItemMst = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _chosaNen)

                For Each pKey As DAOChosahyo.PrimaryKey In pKeys
                    '調査票テーブル取得
                    dtChosahyo = DAOChosahyo.GetChosahyoTable(db, pKey)
                    '調査票項目取得
                    dcChosahyo = ComUtil.Chosahyo.GetItem(dtChoItemMst, dtChosahyo)

                    Dim targetDataList As New List(Of Shinsa.TargetData)
                    For Each dc In dcChosahyo
                        Dim meisaiNo As Integer
                        If dc.Key.Contains(ComConst.ITEM_NO_DELIMITER) Then
                            meisaiNo = CInt(dc.Key.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(1))
                        Else
                            meisaiNo = 0
                        End If
                        targetDataList.Add(New Shinsa.TargetData With {.ItemNo = dc.Key, .Value = dc.Value.値, .MeisaiNo = meisaiNo})
                    Next

                    For Each row As DataRow In dtRonri.Rows
                        Dim val1 As Decimal
                        Dim val2 As Decimal
                        Dim val3 As Decimal
                        Dim tanshu As Decimal

                        '可変項目ならば
                        If row("繰り返し").ToString = "○" Then

                            '項目番号1、項目番号2が可変項目か判定するためのフラグ
                            '可変項目の時はfalse
                            Dim koumoku1_flag As Boolean = False
                            Dim koumoku2_flag As Boolean = False

                            '項目番号１が可変項目かチェック
                            Dim query = From dr In dtChoItemMst Where dr("項目番号").ToString = row("項目番号１").ToString
                            If query.Any() Then
                                If Integer.Parse(query(0)("可変区分").ToString) = 0 Then
                                    koumoku1_flag = True
                                End If
                            End If
                            '項目番号２が可変項目かチェック
                            Dim query2 = From dr In dtChoItemMst Where dr("項目番号").ToString = row("項目番号２").ToString
                            If query2.Any() Then
                                If Integer.Parse(query2(0)("可変区分").ToString) = 0 Then
                                    koumoku2_flag = True
                                End If
                            End If

                            '可変区分のデータ数分ループ
                            For Each targetdata In targetDataList
                                'どちらも可変項目の場合
                                If koumoku1_flag = False And koumoku2_flag = False Then
                                    If row("項目番号２").ToString = Strings.Left(targetdata.ItemNo, koumoku_mojisuu) Then
                                        If Decimal.TryParse(dcChosahyo(row("項目番号１").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).値, val1) _
                                            AndAlso Decimal.TryParse(dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).値, val2) _
                                            AndAlso Decimal.TryParse(row("値").ToString, val3) Then

                                            tanshu = val2 / val1 * val3

                                            Dim low As Decimal = Decimal.Parse(row("下限").ToString)
                                            Dim up As Decimal = Decimal.Parse(row("上限").ToString)

                                            If Not (low <= tanshu And tanshu <= up) Then
                                                'エラー出力対象シートか判定
                                                If Not _sheet.Contains(dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).シート名) Then
                                                    Continue For
                                                End If
                                                ' エラー帳票出力用データを設定
                                                Dim errCheckDictionary As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                                cnt = cnt + 1
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.No), cnt.ToString)
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.都道府県), ComUtil.GetTodofuken(pKey.censusNo))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.市区町村), ComUtil.GetShikuchoson(pKey.censusNo))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.旧市区町村), ComUtil.GetKyuShikuchoson(pKey.censusNo))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.農業集落), ComUtil.GetNogyoShuraku(pKey.censusNo))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.調査区), ComUtil.GetChosaku(pKey.censusNo))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.客体番号), ComUtil.GetKyakutaiNo(pKey.censusNo))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーシート名), dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).シート名)
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.項目番号), row("項目番号２").ToString)
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.可変行番号), targetdata.MeisaiNo.ToString)
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーサイン), "")
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラー内容),
                                                                       AddStrCRLF(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo & ":" & row("チェック項目名").ToString & "が下限「" & row("下限").ToString & "」から上限「" & row("上限").ToString & "」の範囲ではありません。", エラー時コメント改行文字数))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーチェック種別), ComConst.エラーチェック種別.リスト(ComConst.エラーチェック種別.enm.範囲))
                                                errCheckList.Add(errCheckDictionary)
                                            End If
                                        End If
                                    End If

                                    '項目番号1のみが可変項目の場合
                                ElseIf koumoku1_flag = False Then
                                    If row("項目番号２").ToString = Strings.Left(targetdata.ItemNo, koumoku_mojisuu) Then
                                        '項目番号1と項目番号が同じtargetdataListの可変の数分ループ
                                        Dim query3 = From dr In targetDataList Where Strings.Left(dr.ItemNo, koumoku_mojisuu) = row("項目番号１").ToString
                                        For Each kahen1_list In query3
                                            If Decimal.TryParse(dcChosahyo(row("項目番号１").ToString & ComConst.ITEM_NO_DELIMITER & kahen1_list.MeisaiNo).値, val1) _
                                                AndAlso Decimal.TryParse(dcChosahyo(row("項目番号２").ToString).値, val2) _
                                                AndAlso Decimal.TryParse(row("値").ToString, val3) Then

                                                tanshu = val2 / val1 * val3

                                                Dim low As Decimal = Decimal.Parse(row("下限").ToString)
                                                Dim up As Decimal = Decimal.Parse(row("上限").ToString)

                                                If Not (low <= tanshu And tanshu <= up) Then
                                                    'エラー出力対象シートか判定
                                                    If Not _sheet.Contains(dcChosahyo(row("項目番号２").ToString).シート名) Then
                                                        Continue For
                                                    End If
                                                    ' エラー帳票出力用データを設定
                                                    Dim errCheckDictionary As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                                    cnt = cnt + 1
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.No), cnt.ToString)
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.都道府県), ComUtil.GetTodofuken(pKey.censusNo))
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.市区町村), ComUtil.GetShikuchoson(pKey.censusNo))
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.旧市区町村), ComUtil.GetKyuShikuchoson(pKey.censusNo))
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.農業集落), ComUtil.GetNogyoShuraku(pKey.censusNo))
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.調査区), ComUtil.GetChosaku(pKey.censusNo))
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.客体番号), ComUtil.GetKyakutaiNo(pKey.censusNo))
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーシート名), dcChosahyo(row("項目番号２").ToString).シート名)
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.項目番号), row("項目番号２").ToString)
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.可変行番号), If(CInt(dcChosahyo(row("項目番号２").ToString).可変範囲) > 0, row.ToString, ""))
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーサイン), "")
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラー内容),
                                                                       AddStrCRLF(row("項目番号２").ToString & ":" & row("チェック項目名").ToString & "が下限「" & row("下限").ToString & "」から上限「" & row("上限").ToString & "」の範囲ではありません。", エラー時コメント改行文字数))
                                                    errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーチェック種別), ComConst.エラーチェック種別.リスト(ComConst.エラーチェック種別.enm.範囲))
                                                    errCheckList.Add(errCheckDictionary)
                                                End If
                                            End If
                                        Next
                                    End If

                                    '項目番号2のみが可変項目の場合
                                ElseIf koumoku2_flag = False Then
                                    If row("項目番号２").ToString = Strings.Left(targetdata.ItemNo, koumoku_mojisuu) Then
                                        If Decimal.TryParse(dcChosahyo(row("項目番号１").ToString).値, val1) _
                                            AndAlso Decimal.TryParse(dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).値, val2) _
                                            AndAlso Decimal.TryParse(row("値").ToString, val3) Then

                                            tanshu = val2 / val1 * val3

                                            Dim low As Decimal = Decimal.Parse(row("下限").ToString)
                                            Dim up As Decimal = Decimal.Parse(row("上限").ToString)

                                            If Not (low <= tanshu And tanshu <= up) Then
                                                'エラー出力対象シートか判定
                                                If Not _sheet.Contains(dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).シート名) Then
                                                    Continue For
                                                End If
                                                ' エラー帳票出力用データを設定
                                                Dim errCheckDictionary As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                                cnt = cnt + 1
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.No), cnt.ToString)
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.都道府県), ComUtil.GetTodofuken(pKey.censusNo))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.市区町村), ComUtil.GetShikuchoson(pKey.censusNo))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.旧市区町村), ComUtil.GetKyuShikuchoson(pKey.censusNo))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.農業集落), ComUtil.GetNogyoShuraku(pKey.censusNo))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.調査区), ComUtil.GetChosaku(pKey.censusNo))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.客体番号), ComUtil.GetKyakutaiNo(pKey.censusNo))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーシート名), dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).シート名)
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.項目番号), row("項目番号２").ToString)
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.可変行番号), targetdata.MeisaiNo.ToString)
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーサイン), "")
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラー内容),
                                                                       AddStrCRLF(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo & ":" & row("チェック項目名").ToString & "が下限「" & row("下限").ToString & "」から上限「" & row("上限").ToString & "」の範囲ではありません。", エラー時コメント改行文字数))
                                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーチェック種別), ComConst.エラーチェック種別.リスト(ComConst.エラーチェック種別.enm.範囲))
                                                errCheckList.Add(errCheckDictionary)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            If Decimal.TryParse(dcChosahyo(row("項目番号１").ToString).値, val1) _
                                AndAlso Decimal.TryParse(dcChosahyo(row("項目番号２").ToString).値, val2) _
                                AndAlso Decimal.TryParse(row("値").ToString, val3) Then

                                If val1 <> 0 Then

                                    tanshu = val2 / val1 * val3

                                    Dim low As Decimal = Decimal.Parse(row("下限").ToString)
                                    Dim up As Decimal = Decimal.Parse(row("上限").ToString)

                                    If Not (low <= tanshu And tanshu <= up) Then
                                        'エラー出力対象シートか判定
                                        If Not _sheet.Contains(dcChosahyo(row("項目番号２").ToString).シート名) Then
                                            Continue For
                                        End If
                                        'エラー帳票出力用データを設定
                                        Dim errCheckDictionary As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                        cnt = cnt + 1
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.No), cnt.ToString)
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.都道府県), ComUtil.GetTodofuken(pKey.censusNo))
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.市区町村), ComUtil.GetShikuchoson(pKey.censusNo))
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.旧市区町村), ComUtil.GetKyuShikuchoson(pKey.censusNo))
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.農業集落), ComUtil.GetNogyoShuraku(pKey.censusNo))
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.調査区), ComUtil.GetChosaku(pKey.censusNo))
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.客体番号), ComUtil.GetKyakutaiNo(pKey.censusNo))
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーシート名), dcChosahyo(row("項目番号２").ToString).シート名)
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.項目番号), row("項目番号２").ToString)
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.可変行番号), If(CInt(dcChosahyo(row("項目番号２").ToString).可変範囲) > 0, row.ToString, ""))
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーサイン), "")
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラー内容),
                                                               AddStrCRLF(row("項目番号２").ToString & ":" & row("チェック項目名").ToString & "が下限「" & row("下限").ToString & "」から上限「" & row("上限").ToString & "」の範囲ではありません。", エラー時コメント改行文字数))
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧複数客体.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧複数客体.enm.エラーチェック種別), ComConst.エラーチェック種別.リスト(ComConst.エラーチェック種別.enm.範囲))
                                        errCheckList.Add(errCheckDictionary)
                                    End If
                                End If
                            End If

                            '進捗を進める
                            progressDialog.AddValue = 1
                        End If
                    Next
                Next
            End Using

        Catch ex As Exception
            Throw ex
        End Try

        Return errCheckList
    End Function

End Class
