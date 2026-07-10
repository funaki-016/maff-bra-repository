Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.10.11 |Daiko               | 要件No1 バージョン区分追加
'//  REV_002   | 2022.12.14 |Daiko               | 要件No4 バージョン区分追加
'//  REV_003   | 2023.02.08 |Daiko               | 要件No12 生産費の地域集計時のInsert対象の集計結果表レコードに裏項番を含める
'//  REV_004   | 2022.01.04 |Daiko               | 要件No8 バージョン区分追加
'//  REV_005   | 2025.09.18 |GCU                 | 要件No13 集計条件不具合の修正
'//  REV_006   | 2025.08.28 |GCU                 | 要件No2 継続区分追加
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 集計結果表作成画面
''' </summary>
''' <remarks></remarks>
Public Class BRA5110F

    ''' <summary>地域</summary>
    Private _chiiki As Dictionary(Of String, String)
    ''' <summary>部門</summary>
    Private _bumon As Dictionary(Of String, String)
    ''' <summary>Insert対象の集計結果表レコード</summary>
    Private _syukeiRecordList As New List(Of ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)))


    ''' <summary>規模階層表示フラグ</summary>
    Private _dispKiboKaisouFlg As Boolean

    ''' <summary>DataGridView行数</summary>
    Private Const DATAGRIDVIEW_ROW As Integer = 20
    ''' <summary>条件文字列(and)</summary>
    Private Const STR_AND As String = "AND"
    ''' <summary>条件文字列(or)</summary>
    Private Const STR_OR As String = "OR"
    ''' <summary>区切り文字</summary>
    Private Const CHR_SEPARATOR As Char = ";"c

    ''' <summary>集計パターン：定型集計</summary>
    Private Const PATTERN_FIX As String = "定型"
    ''' <summary>集計パターン：任意集計</summary>
    Private Const PATTERN_ANY As String = "任意"
    ''' <summary>生産費平均値種類の数</summary>
    Private Const SEISANHIHEIKIN_MAX As Integer = 4
    ''' <summary>生産費平均値種類の数_牛乳</summary>
    Private Const GYUNYU_SEISANHIHEIKIN_MAX As Integer = 5
    ''' <summary>集計結果表作成論理：生産費平均値種類 総数</summary>
    Private Const KBN_SOUSU As String = "1"
    ''' <summary>集計結果表作成論理：生産費平均値種類 総数以外</summary>
    Private Const KBN_EXCEPTSOUSU As String = "9"
    ''' <summary>ドロップダウンリスト未選択</summary>
    Private Const CODE_UNSELECTED As String = "0"

    ''' <summary>選択項目区切文字</summary>
    Private Const SELECT_ITEM_DELIMITER As String = "、"

    ''' <summary>ファイルタイトル文字列</summary>
    Private Const SYUKEI_JOUKEN_FILE_TITLE As String = "集計条件"

    ''' <summary>プログレスバー</summary>
    Private _progressDialog As ProgressDialog

    ''' <summary>集計クラス</summary>
    Private _createSyukei As CreateSyukei


    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA5110F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '営農経営体区分コンボボックス設定
            ComUtil.SetEinouKeieitaiComboBox(lblEinouKeieitai, cboEinouKeieitai)

            '調査区分取得
            Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

            If Not ComUtil.IsEinou Then
                '調査年コンボボックス設定
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    ComUtil.KobetsuKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, chosakubun, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center, ComConst.欠測値補完.有)
                End Using
            End If

            '平均種類コンボボックス設定
            ComUtil.SetHeikinSyuruiComboBox(cboHeikinSyurui)

            '任意階層利用コンボボックス設定
            ComUtil.SetNiniKaisouComboBox(cboNiniKaisou)

            '規模階層コンボボックス設定
            ComUtil.SetKiboKaisouComboBox(cboKiboKaisou)

            'REV_006↓
            '継続区分コンボボックス設定
            ComUtil.SetKeizokuKubunComboBox(lblKeizokuKubun, cboKeizokuKubun, chkZenkaiCensus, cboEinouKeieitai)

            '営農と生産費で設定するコンボボックスを切り替える
            If ComUtil.IsEinou Then
                lblSyukei1.Text = "集計１"
                lblSyukei2.Text = "集計２"
                lblSyukei3.Text = "集計３"
                lblSyukei4.Text = "集計４"
                ' REV_004↓
                ComUtil.SetSyukei1ComboBox(cboEinouKeieitai, cboSyukei1, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), ComUtil.GetChosakubun(cboEinouKeieitai)))
                ' REV_004↑
            Else
                lblEinouKeieitai.Visible = False
                cboEinouKeieitai.Visible = False
                ComUtil.SetSyukeiKubunComboBox(cboSyukei1)
                ComUtil.SetTahataKubunComboBox(cboSyukei2)
                ComUtil.SetBeerMugiKubunComboBox(cboSyukei3)
                ComUtil.SetTensaiKubunComboBox(cboSyukei4)
            End If

            '営農個人以外では部門選択は非活性
            If Not CommonInfo.Chosakubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) Then
                btnBumon.Enabled = False
                txtBumon.Enabled = False
            End If

            'DataGridView設定
            ComUtil.ConfigDgvEditable(Me.dgvList, DATAGRIDVIEW_ROW)

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try

                    db.BeginTrans()

                    '集計用センサス番号管理削除
                    DAOOther.DeleteSyukeiCensusNo(db)

                    db.CommitTrans()

                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ' REV_004↓
    Private Sub cboChosaNen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboChosaNen.SelectedIndexChanged
        If ComUtil.IsEinou Then
            ComUtil.SetSyukei1ComboBox(cboEinouKeieitai, cboSyukei1, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), ComUtil.GetChosakubun(cboEinouKeieitai)))
        End If
    End Sub
    ' REV_004↑

    Private Sub cboNiniKaisou_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNiniKaisou.SelectedIndexChanged
        '任意階層利用有りの場合、集計１～集計４と部門を非活性とする
        If ComUtil.IsEinou Then
            Try
                If cboNiniKaisou.SelectedValue.Equals(ComConst.任意階層利用.有) Then
                    ' REV_004↓
                    ComUtil.SetSyukei1ComboBox(cboEinouKeieitai, cboSyukei1, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), ComUtil.GetChosakubun(cboEinouKeieitai)))
                    ' REV_004↑
                    cboSyukei1.Enabled = False
                    cboSyukei2.Enabled = False
                    cboSyukei3.Enabled = False
                    cboSyukei4.Enabled = False
                    btnBumon.Enabled = False
                    txtBumon.Text = String.Empty
                    _bumon = Nothing
                Else
                    cboSyukei1.Enabled = True
                    cboSyukei2.Enabled = True
                    cboSyukei3.Enabled = True
                    cboSyukei4.Enabled = True
                    If CommonInfo.Chosakubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) Then
                        btnBumon.Enabled = True
                    End If
                End If
            Catch ex As Exception
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
            End Try
        End If
    End Sub

    Private Sub cboEinouKeieitai_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboEinouKeieitai.SelectedIndexChanged
        If ComUtil.IsEinou Then
            Try
                '法人経営体又は農業経営体が選択されたら部門を削除してボタンを非活性にする
                If Not cboEinouKeieitai.SelectedValue.Equals(ComConst.営農経営体区分.個人経営体) Then
                    txtBumon.Text = String.Empty
                    _bumon = Nothing
                    btnBumon.Enabled = False
                    txtBumon.Enabled = False
                Else
                    btnBumon.Enabled = True
                    txtBumon.Enabled = True
                End If

                '調査区分取得
                Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

                '調査年コンボボックス設定
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    ComUtil.KobetsuKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, chosakubun, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center, ComConst.欠測値補完.有)
                End Using

                ' REV_004↓
                '集計１コンボボックス設定
                ComUtil.SetSyukei1ComboBox(cboEinouKeieitai, cboSyukei1, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), ComUtil.GetChosakubun(cboEinouKeieitai)))
                ' REV_004↑
            Catch ex As Exception
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
            End Try
        End If
    End Sub

    Private Sub cboSyukei1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSyukei1.SelectedIndexChanged
        If ComUtil.IsEinou Then
            Try
                '集計２コンボボックス設定
                ' REV_004↓
                ComUtil.SetSyukei2ComboBox(cboEinouKeieitai, cboSyukei1, cboSyukei2, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), ComUtil.GetChosakubun(cboEinouKeieitai)))
                ' REV_004↑
            Catch ex As Exception
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
            End Try
        End If
    End Sub

    Private Sub cboSyukei2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSyukei2.SelectedIndexChanged
        If ComUtil.IsEinou Then
            Try
                '集計３コンボボックス設定
                ' REV_004↓
                ComUtil.SetSyukei3ComboBox(cboEinouKeieitai, cboSyukei1, cboSyukei2, cboSyukei3, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), ComUtil.GetChosakubun(cboEinouKeieitai)))
                ' REV_004↑
            Catch ex As Exception
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
            End Try
        End If
    End Sub

    Private Sub cboSyukei3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSyukei3.SelectedIndexChanged
        If ComUtil.IsEinou Then
            Try
                '集計４コンボボックス設定
                ' REV_004↓
                ComUtil.SetSyukei4ComboBox(cboEinouKeieitai, cboSyukei1, cboSyukei2, cboSyukei3, cboSyukei4, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), ComUtil.GetChosakubun(cboEinouKeieitai)))
                ' REV_004↑
            Catch ex As Exception
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
            End Try
        End If
    End Sub

    Private Sub cboSyukei4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSyukei4.SelectedIndexChanged
        If ComUtil.IsEinou Then
            Try
                If Not cboSyukei4.SelectedValue Is Nothing Then
                    '集計４まで選択されたら部門を削除する
                    txtBumon.Text = String.Empty
                    _bumon = Nothing
                End If
            Catch ex As Exception
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
            End Try
        End If
    End Sub

    'REV_006↓
    Private Sub cboKeizokuKubun_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKeizokuKubun.SelectedIndexChanged
        Try
            '継続区分の選択に応じて前回センサス番号使用チェックボックスの活性状態を制御
            ComUtil.SetZenkaiCensusEnabled(chkZenkaiCensus, cboKeizokuKubun)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
    'REV_006↑

    ''' <summary>
    ''' 集計ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAggregate_Click(sender As Object, e As EventArgs) Handles btnAggregate.Click

        '個別結果表項目マスタ
        Dim dtKobetsuItemMst As DataTable
        '集計結果表項目マスタ
        Dim dtSyukeiItemMst As DataTable
        '集計結果表作成論理
        Dim dtCreateRonri As DataTable

        _syukeiRecordList = New List(Of ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)))

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Try
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                '個別結果表項目マスタ取得(裏項番含める)
                ' REV_001↓
                Dim chosaNen As String = cboChosaNen.SelectedValue.ToString
                If ComUtil.GetChosakubun(cboEinouKeieitai).Equals(ComConst.営農経営体区分.農業経営体) Then
                    'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, ComConst.調査区分.営農類型別経営統計_個人, True)
                    dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, ComConst.調査区分.営農類型別経営統計_個人, ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun), True)
                    'dtKobetsuItemMst.Merge(DAOOther.GetKobetsuKekkahyoItemMaster(db, ComConst.調査区分.営農類型別経営統計_法人, True))
                    dtKobetsuItemMst.Merge(DAOOther.GetKobetsuKekkahyoItemMaster(db, ComConst.調査区分.営農類型別経営統計_法人, ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun), True))
                Else
                    'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, True)
                    dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun), True)
                End If
                ' REV_001↑

                If ComUtil.IsEinou Then
                    '集計結果表作成論理取得
                    ' REV_002↓
                    'dtCreateRonri = DAOOther.GetSyukeiKekkahyoSakuseiRonri(db, ComUtil.GetChosakubun(cboEinouKeieitai))
                    dtCreateRonri = DAOOther.GetSyukeiKekkahyoSakuseiRonri(db, ComUtil.GetChosakubun(cboEinouKeieitai), ComUtil.getVersionKubunTaikei(chosaNen, ComUtil.GetChosakubun(cboEinouKeieitai)))
                    '集計結果表項目マスタ取得(裏項番含める)
                    'dtSyukeiItemMst = DAOOther.GetSyukeiItemMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai))
                    dtSyukeiItemMst = DAOOther.GetSyukeiItemMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai), ComUtil.getVersionKubunTaikei(chosaNen, ComUtil.GetChosakubun(cboEinouKeieitai)))
                    ' REV_002↑
                Else
                    ' REV_002↓
                    'dtCreateRonri = DAOOther.GetSyukeiKekkahyoSakuseiRonri(db, CommonInfo.Chosakubun, KBN_SOUSU)
                    dtCreateRonri = DAOOther.GetSyukeiKekkahyoSakuseiRonri(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun), KBN_SOUSU)
                    'dtSyukeiItemMst = DAOOther.GetSyukeiItemMaster(db, CommonInfo.Chosakubun)
                    dtSyukeiItemMst = DAOOther.GetSyukeiItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun))
                    ' REV_002↑
                End If

                '-----------------------------------------------------------------------------------------------------
                'エラーチェック
                '-----------------------------------------------------------------------------------------------------
                '調査年選択チェック
                If cboChosaNen.SelectedValue Is Nothing Then
                    Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                '集計番号・集計名称入力チェック
                If String.IsNullOrWhiteSpace(txtSyukeiNo.Text) Or String.IsNullOrWhiteSpace(txtSyukeiName.Text) Then
                    Message.ShowMsgBox(MessageID.MSG_E_046, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                '集計番号入力値チェック
                If Not validateSyukeiNo(txtSyukeiNo.Text) Then
                    Message.ShowMsgBox(MessageID.MSG_E_057, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                '地域入力チェック
                If _chiiki Is Nothing OrElse _chiiki.Count = 0 Then
                    Message.ShowMsgBox(MessageID.MSG_E_059, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                '集計番号存在チェック
                If isExistSyukeiNo(db) Then
                    Message.ShowMsgBox(MessageID.MSG_E_047, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                '任意階層マスタ存在チェック
                If cboNiniKaisou.SelectedValue.Equals(ComConst.任意階層利用.有) Then
                    If Not isExistNiniKaisou(db) Then
                        Message.ShowMsgBox(MessageID.MSG_E_051, MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If
                End If

                '集計条件１～４妥当性チェック
                Dim errLst As List(Of String) = isExecutableNiniSyukei(db)
                If errLst.Count > 0 Then
                    Message.ShowMsgForm(Me, MessageID.MSG_E_010, {String.Join(vbCrLf, errLst)})
                    Exit Sub
                End If

                '営農のみ：集計１～４コンボ入力チェック
                If ComUtil.IsEinou Then
                    If Not isSelectedEinouChusyutsuJoken() Then
                        Message.ShowMsgBox(MessageID.MSG_E_053, MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If
                End If

                '集計結果表作成論理のデータが１件も存在しない場合
                If dtCreateRonri.Rows.Count = 0 Then
                    'エラーメッセージ
                    Message.ShowMsgBox(MessageID.MSG_E_049, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                '集計実行
                Me.ExecuteSyukei(db, dtKobetsuItemMst, dtSyukeiItemMst, dtCreateRonri)

            Catch ex As CreateSyukeiException
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Message.ShowMsgBox(MessageID.MSG_E_050, {ex.SyukeiNo, ex.ItemNo}, MsgBoxStyle.OkOnly)

            Catch ex As Exception
                'システムログ出力
                db.RollBackTrans()

                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
            Finally
                Me.closeProgressDialog()
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            End Try
        End Using
    End Sub

    ''' <summary>
    ''' 集計処理（集計条件の設定と集計パターンの取得）
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ExecuteSyukei(db As DBAccess, dtKobetsuItemMst As DataTable, dtSyukeiItemMst As DataTable, dtCreateRonri As DataTable)
        '採番開始集計番号
        Dim startSyukeiNo As String
        '採番終了集計番号
        Dim endSyukeiNo As String

        '採番開始集計番号を設定
        startSyukeiNo = txtSyukeiNo.Text
        '-----------------------------------------------------------------------------------------------------
        '集計条件を設定
        '-----------------------------------------------------------------------------------------------------
        '集計条件クラス
        Dim syukeiInfo As New DAOKobetsuKekkahyo.SyukeiInfo

        '集計番号
        syukeiInfo.syukeiNo = CommonInfo.Koutei & CInt(CommonInfo.Kyoku) & txtSyukeiNo.Text
        '連番
        syukeiInfo.groupKey = 1
        '集計名称
        syukeiInfo.syukeiName = txtSyukeiName.Text
        '集計１
        syukeiInfo.syukei1 = If(cboSyukei1.SelectedValue Is Nothing, Nothing, cboSyukei1.SelectedValue.ToString)
        '集計２
        syukeiInfo.syukei2 = If(cboSyukei2.SelectedValue Is Nothing, Nothing, cboSyukei2.SelectedValue.ToString)
        '集計３
        syukeiInfo.syukei3 = If(cboSyukei3.SelectedValue Is Nothing, Nothing, cboSyukei3.SelectedValue.ToString)
        '集計４
        syukeiInfo.syukei4 = If(cboSyukei4.SelectedValue Is Nothing, Nothing, cboSyukei4.SelectedValue.ToString)
        'REV_004↓
        '基本・詳細項目集計
        If Not syukeiInfo.syukei1 Is Nothing Then
            Dim dt As New DataTable
            dt = DAOOther.getEinouSyukeiJoukenMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai), syukeiInfo.syukei1, syukeiInfo.syukei2, syukeiInfo.syukei3, syukeiInfo.syukei4, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))
            If dt.Rows.Count > 0 Then
                syukeiInfo.kihonSyosaiSyukei = ComConst.基本詳細項目集計.getValue(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun), dt.Rows(0).Item("集計コード").ToString)
            End If
        End If
        If String.IsNullOrEmpty(syukeiInfo.kihonSyosaiSyukei) Then
            syukeiInfo.kihonSyosaiSyukei = ComConst.基本詳細項目集計.基本詳細項目集計
        End If
        'REV_004↑

        '任意集計
        syukeiInfo.niniSyukei = getNiniSyukei()
        '任意集計（指標部用）
        syukeiInfo.niniSyukeiShihyou = getNiniSyukeiUseShihyou()
        '部門
        syukeiInfo.bumon = CODE_UNSELECTED

        '集計条件に規模階層を表示させるかのフラグ
        _dispKiboKaisouFlg = True

        '進捗ダイアログを表示する
        Me.openProgressDialog(db, syukeiInfo, dtCreateRonri)

        '集計結果表作成クラス
        _createSyukei = New CreateSyukei(db, ComUtil.GetChosakubun(cboEinouKeieitai),
                                                         cboChosaNen.SelectedValue.ToString,
                                                         dtKobetsuItemMst,
                                                         dtSyukeiItemMst,
                                                         dtCreateRonri,
                                                         False, 'REV_006←↓
                                                         If(cboKeizokuKubun.Visible AndAlso cboKeizokuKubun.SelectedValue IsNot Nothing, cboKeizokuKubun.SelectedValue.ToString, Nothing),
                                                         chkZenkaiCensus.Checked)

        '-----------------------------------------------------------------------------------------------------
        '集計パターンの取得
        '-----------------------------------------------------------------------------------------------------
        If ComUtil.IsEinou Then
            '【営農】

            '集計パターン(定型/任意)
            syukeiInfo.syukeipattern = If(cboSyukei1.SelectedValue Is Nothing And (_bumon Is Nothing OrElse _bumon.Count = 0), PATTERN_ANY, PATTERN_FIX)

            '個別結果表抽出条件を取得
            Dim dtEinouChusyutsuJouken As New DataTable
            If Not syukeiInfo.syukei1 Is Nothing Then
                ' REV_004↓
                dtEinouChusyutsuJouken = DAOOther.getEinouSyukeiJoukenMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai), syukeiInfo.syukei1, syukeiInfo.syukei2, syukeiInfo.syukei3, syukeiInfo.syukei4, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))
                ' REV_004↑
            End If

            Dim einouChusyutsuJoukenList As New List(Of DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu)

            '営農経営体抽出条件がある場合：抽出条件とそれに紐づく階層コード。 ない場合：部門コード
            If dtEinouChusyutsuJouken.Rows.Count > 0 Then

                '親子関係2段階まで対応
                For Each row As DataRow In dtEinouChusyutsuJouken.Rows
                    '『a-xxx ～ a-xxx までの全ての集計を行う』
                    If row.Item("個別結果表抽出条件").ToString.Contains(CHR_SEPARATOR) Then

                        '複数条件の集計を行う場合は集計条件文字列に規模階層を表示させない
                        _dispKiboKaisouFlg = False

                        For Each str As String In row.Item("個別結果表抽出条件").ToString.Split(CHR_SEPARATOR)
                            '抽出条件を分割して取得した集計コードからさらに集計条件マスタを取得
                            ' REV_004↓
                            Dim einouChusyutsuJouken = DAOOther.getEinouSyukeiJoukenMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai), str, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun)).Rows(0)
                            ' REV_004↑
                            Dim syukeiCd As String = einouChusyutsuJouken.Item("集計コード").ToString
                            Dim syukei1 As String = einouChusyutsuJouken.Item("集計１").ToString
                            Dim syukei2 As String = einouChusyutsuJouken.Item("集計２").ToString
                            Dim syukei3 As String = einouChusyutsuJouken.Item("集計３").ToString
                            Dim syukei4 As String = einouChusyutsuJouken.Item("集計４").ToString
                            ' REV_004↓
                            Dim jouken As String = einouChusyutsuJouken.Item("個別結果表抽出条件").ToString
                            If ComConst.基本詳細項目集計.IsShosaiOnly(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun), einouChusyutsuJouken.Item("集計コード").ToString) Then
                                jouken = "K000026=2 and " + jouken
                            End If
                            ' REV_004↑
                            Dim kaisouCd As String = einouChusyutsuJouken.Item("階層コード").ToString
                            Dim bumonCd As String = If(IsDBNull(einouChusyutsuJouken.Item("部門コード")), Nothing, einouChusyutsuJouken.Item("部門コード").ToString)
                            Dim einouruiKei As String = einouChusyutsuJouken.Item("営農類型").ToString
                            einouChusyutsuJoukenList.Add(New DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu(syukeiCd, syukei1, syukei2, syukei3, syukei4, jouken, kaisouCd, bumonCd, einouruiKei))
                        Next
                    Else
                        'そのままリストを作成
                        Dim syukeiCd As String = row.Item("集計コード").ToString
                        Dim syukei1 As String = row.Item("集計１").ToString
                        Dim syukei2 As String = row.Item("集計２").ToString
                        Dim syukei3 As String = row.Item("集計３").ToString
                        Dim syukei4 As String = row.Item("集計４").ToString
                        Dim jouken As String = row.Item("個別結果表抽出条件").ToString
                        ' REV_004↓
                        If ComConst.基本詳細項目集計.IsShosaiOnly(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun), row.Item("集計コード").ToString) Then
                            jouken = "K000026=2 and " + jouken
                        End If
                        ' REV_004↑
                        Dim kaisouCd As String = row.Item("階層コード").ToString
                        Dim bumonCd As String = If(IsDBNull(row.Item("部門コード")), Nothing, row.Item("部門コード").ToString)
                        Dim einouruiKei As String = row.Item("営農類型").ToString
                        einouChusyutsuJoukenList.Add(New DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu(syukeiCd, syukei1, syukei2, syukei3, syukei4, jouken, kaisouCd, bumonCd, einouruiKei))
                    End If
                Next
            Else
                If _bumon Is Nothing OrElse _bumon.Keys.Count = 0 Then
                    '集計１～４なし、部門もなし
                    einouChusyutsuJoukenList.Add(New DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing))

                Else
                    '部門別集計の場合
                    For Each bumonCd As String In _bumon.Keys
                        '部門コードから営農抽出条件マスタを取得
                        ' REV_004↓
                        Dim einouChusyutsuJouken = DAOOther.getEinouSyukeiJoukenMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai), bumonCd, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun), True).Rows(0)
                        ' REV_004↑
                        Dim kaisouCd As String = einouChusyutsuJouken.Item("階層コード").ToString
                        Dim jouken As String = einouChusyutsuJouken.Item("個別結果表抽出条件").ToString
                        ' REV_004↓
                        If ComConst.基本詳細項目集計.IsShosaiOnly(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun), einouChusyutsuJouken.Item("集計コード").ToString) Then
                            jouken = "K000026=2 and " + jouken
                        End If
                        Dim einouruiKei As String = einouChusyutsuJouken.Item("営農類型").ToString
                        einouChusyutsuJoukenList.Add(New DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu(einouChusyutsuJouken.Item("集計コード").ToString, Nothing, Nothing, Nothing, Nothing, jouken, kaisouCd, bumonCd, einouruiKei))
                        ' REV_004↑
                    Next
                End If
            End If

            '採番終了集計番号を設定
            endSyukeiNo = String.Format("{0:D6}", CInt(syukeiInfo.syukeiNo.Substring(2)) + einouChusyutsuJoukenList.Count - 1)

            '集計番号に空きがあるか調べる
            If isExistSyukeiNoRange(db, startSyukeiNo, endSyukeiNo) Then
                Me.closeProgressDialog()
                Message.ShowMsgBox(MessageID.MSG_E_052, {startSyukeiNo, endSyukeiNo}, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '個別結果表抽出条件毎の集計を実行
            For Each einouChusyutsuJouken As DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu In einouChusyutsuJoukenList
                '部門
                syukeiInfo.bumon = If(einouChusyutsuJouken.bumonCd Is Nothing, CODE_UNSELECTED, einouChusyutsuJouken.bumonCd)

                ' REV_004↓
                If einouChusyutsuJouken.syukeiCd IsNot Nothing Then
                    '基本・詳細項目集計
                    syukeiInfo.kihonSyosaiSyukei = ComConst.基本詳細項目集計.getValue(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun), einouChusyutsuJouken.syukeiCd)
                End If
                ' REV_004↑

                Dim beforeSyukeiRecordCount As Integer = _syukeiRecordList.Count
                ExecuteSyukei_kiboKaisou(db, syukeiInfo, einouChusyutsuJouken)
                Dim afterSyukeiRecordCount As Integer = _syukeiRecordList.Count

                If beforeSyukeiRecordCount <> afterSyukeiRecordCount Then
                    '1定型集計で1つでも集計が行われたら定型集計条件毎に集計番号をインクリメント
                    syukeiInfo.syukeiNo = CommonInfo.Koutei & CInt(CommonInfo.Kyoku) & String.Format("{0:D6}", CInt(syukeiInfo.syukeiNo.Substring(2)) + 1)
                End If
                '連番をリセット
                syukeiInfo.groupKey = 1
            Next

            '集計結果表レコードをInsertする
            db.BeginTrans()

            Me.insertSyukeiKekkahyou(db)

            db.CommitTrans()

            Me.closeProgressDialog()

            '完了メッセージ
            If _syukeiRecordList.Any Then
                '今回集計した集計番号の最後を取得する
                Dim unFindFlg As Boolean = True
                Do While unFindFlg
                    If String.IsNullOrEmpty(DAOSyukeiKekkahyo.getSyukeiNo(db, ComUtil.GetChosakubun(cboEinouKeieitai), cboChosaNen.SelectedValue.ToString, syukeiInfo.syukeiNo)) Then
                        syukeiInfo.syukeiNo = CommonInfo.Koutei & CInt(CommonInfo.Kyoku) & String.Format("{0:D6}", CInt(syukeiInfo.syukeiNo.Substring(2)) - 1)
                    Else
                        unFindFlg = False
                    End If
                Loop
                Message.ShowMsgBox(MessageID.MSG_I_021, {startSyukeiNo, syukeiInfo.syukeiNo.Substring(2)}, MsgBoxStyle.OkOnly)
            Else
                '集計対象の個別結果表が存在しなければその旨表示し処理を終了する
                Message.ShowMsgBox(MessageID.MSG_I_031, MsgBoxStyle.OkOnly)
            End If

        Else
            '【生産費】
            '画面で選択したコンボボックスから集計パターンを取得
            Dim syukeiPetternList As List(Of DAOKobetsuKekkahyo.SeisanhiChusyutsu)
            syukeiPetternList = setSeisanhiChusyutsuJouken(syukeiInfo)

            '規模階層取得
            Dim kibokaisouList As New List(Of DAOKobetsuKekkahyo.Kibokaisou)
            '平均値を求める際に必要な情報
            kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(CODE_UNSELECTED, Nothing, Nothing, Nothing))
            If cboKiboKaisou.SelectedValue.Equals(ComConst.規模階層.規模階層含) Then

                If cboNiniKaisou.SelectedValue.Equals(ComConst.任意階層利用.無) Then
                    '通常の規模階層を使用する場合
                    Dim dtKiboKaisou As DataTable = DAOOther.getKiboKaisouMaster(db, CommonInfo.Chosakubun)

                    For Each row As DataRow In dtKiboKaisou.Rows
                        Dim kibokaisou As String = row.Item("規模階層").ToString
                        Dim itemNo As String = row.Item("項目番号").ToString
                        Dim max As String = If(IsDBNull(row.Item("上限")), Nothing, row.Item("上限").ToString)
                        Dim min As String = If(IsDBNull(row.Item("下限")), Nothing, row.Item("下限").ToString)
                        kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(kibokaisou, itemNo, max, min))
                    Next

                Else
                    '任意階層を使用する場合
                    ' REV_002↓
                    'Dim dtKiboKaisou As DataTable = DAOOther.GetNiniKaisou(db, CommonInfo.Chosakubun)
                    Dim dtKiboKaisou As DataTable = DAOOther.GetNiniKaisou(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, CommonInfo.Chosakubun))
                    ' REV_002↑

                    For Each row As DataRow In dtKiboKaisou.Rows
                        kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(row.Item("規模階層").ToString, row.Item("項目番号").ToString, row.Item("上限").ToString, row.Item("下限").ToString))
                    Next
                End If
            End If

            '生産費平均値種類ごとのループ(牛乳は平均値種類「5」まで、その他生産費は「4」まで存在する)
            For seisanhiHeikin As Integer = 1 To If(CommonInfo.Chosakubun.Equals(ComConst.調査区分.牛乳生産費統計_個別) Or CommonInfo.Chosakubun.Equals(ComConst.調査区分.経営分析調査_牛乳生産費), GYUNYU_SEISANHIHEIKIN_MAX, SEISANHIHEIKIN_MAX)

                syukeiInfo.seisanhiHeikin = seisanhiHeikin.ToString
                '作成論理の生産費平均値種類：生産費の場合、総数は「1」、総数以外は「9」
                If seisanhiHeikin = 2 Then
                    ' REV_002↓
                    '_createSyukei.addCreateRonriInfo(DAOOther.GetSyukeiKekkahyoSakuseiRonri(db, CommonInfo.Chosakubun, KBN_EXCEPTSOUSU))
                    _createSyukei.addCreateRonriInfo(DAOOther.GetSyukeiKekkahyoSakuseiRonri(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, CommonInfo.Chosakubun), KBN_EXCEPTSOUSU))
                    ' REV_002↑
                End If

                '集計パターンごとのループ
                For Each syukeiPettern As DAOKobetsuKekkahyo.SeisanhiChusyutsu In syukeiPetternList
                    '規模階層毎の集計
                    For Each kaisouInfo As DAOKobetsuKekkahyo.Kibokaisou In kibokaisouList
                        ExecuteSyukei_Chiiki(db, syukeiInfo, syukeiPettern, kaisouInfo)
                    Next
                Next
            Next

            syukeiInfo.seisanhiHeikin = "0"
            '集計パターンごとのループ
            For Each syukeiPettern As DAOKobetsuKekkahyo.SeisanhiChusyutsu In syukeiPetternList
                '規模階層毎の集計
                For Each kaisouInfo As DAOKobetsuKekkahyo.Kibokaisou In kibokaisouList
                    ExecuteSyukei_Merge(db, syukeiInfo, dtKobetsuItemMst, dtSyukeiItemMst, syukeiPettern, kaisouInfo)
                Next
            Next

            '集計結果表レコードをInsertする
            db.BeginTrans()

            Me.insertSyukeiKekkahyou(db)

            db.CommitTrans()

            Me.closeProgressDialog()

            '完了メッセージ
            If _syukeiRecordList.Any Then
                Message.ShowMsgBox(MessageID.MSG_I_020, MsgBoxStyle.OkOnly)
            Else
                '集計対象の個別結果表が存在しなければその旨表示し処理を終了する
                Message.ShowMsgBox(MessageID.MSG_I_031, MsgBoxStyle.OkOnly)
            End If

        End If

    End Sub

    ''' <summary>
    ''' 【営農】集計処理：規模階層ループ
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="syukeiInfo"></param>
    ''' <param name="dtKobetsuItemMst"></param>
    ''' <param name="dtSyukeiItemMst"></param>
    ''' <param name="dtCreateRonri"></param>
    ''' <param name="einouChusyutsuJouken"></param>
    ''' <remarks></remarks>
    Private Overloads Sub ExecuteSyukei_kiboKaisou(db As DBAccess, syukeiInfo As DAOKobetsuKekkahyo.SyukeiInfo, einouChusyutsuJouken As DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu)
        '階層コード
        Dim kibokaisouList As New List(Of DAOKobetsuKekkahyo.Kibokaisou)

        '平均値を求める際に必要な情報
        kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(CODE_UNSELECTED, Nothing, Nothing, Nothing))
        If cboKiboKaisou.SelectedValue.Equals(ComConst.規模階層.規模階層含) Then

            If cboNiniKaisou.SelectedValue.Equals(ComConst.任意階層利用.無) Then
                '通常の規模階層を使用する場合
                Dim dtKiboKaisou As DataTable = DAOOther.getKiboKaisouMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai), einouChusyutsuJouken.kaisouCd)

                For Each row As DataRow In dtKiboKaisou.Rows
                    Dim kibokaisou As String = row.Item("規模階層").ToString
                    Dim itemNo As String = row.Item("項目番号").ToString
                    Dim max As String = If(IsDBNull(row.Item("上限")), Nothing, row.Item("上限").ToString)
                    Dim min As String = If(IsDBNull(row.Item("下限")), Nothing, row.Item("下限").ToString)
                    kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(kibokaisou, itemNo, max, min))
                Next

                '規模階層内に集計コードが指定してあるパターンの集計処理
                If kibokaisouList.Count > 1 AndAlso kibokaisouList(1).jouken.Contains("-") Then
                    For i As Integer = 1 To kibokaisouList.Count - 1
                        '規模階層の項目番号に指定されている集計コードから集計条件を取得する
                        ' REV_004↓
                        kibokaisouList(i).syukeiCd = kibokaisouList(i).jouken
                        Dim dt As DataTable = DAOOther.getEinouSyukeiJoukenMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai), kibokaisouList(i).jouken, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))
                        If dt.Rows.Count <> 0 Then
                            kibokaisouList(i).jouken = dt.Rows(0).Item("個別結果表抽出条件").ToString
                            If ComConst.基本詳細項目集計.IsShosaiOnly(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun), dt.Rows(0).Item("集計コード").ToString) Then
                                kibokaisouList(i).jouken = "K000026=2 and " + kibokaisouList(i).jouken
                            End If
                        Else
                            kibokaisouList(i).jouken = Nothing
                        End If
                        ' REV_004↑
                    Next
                End If
                ' REV_004↓
                kibokaisouList.RemoveAll(Function(n) n.kaisouNo <> "0" And n.jouken Is Nothing)
                ' REV_004↑
            Else
                '任意階層を使用する場合
                ' REV_002↓
                'Dim dtKiboKaisou As DataTable = DAOOther.GetNiniKaisou(db, ComUtil.GetChosakubun(cboEinouKeieitai))
                Dim dtKiboKaisou As DataTable = DAOOther.GetNiniKaisou(db, ComUtil.GetChosakubun(cboEinouKeieitai), ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, ComUtil.GetChosakubun(cboEinouKeieitai)))
                ' REV_002↑

                For Each row As DataRow In dtKiboKaisou.Rows
                    kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(row.Item("規模階層").ToString, row.Item("項目番号").ToString, row.Item("上限").ToString, row.Item("下限").ToString))
                Next
            End If

        End If

        '規模階層毎の集計
        For Each kaisouInfo As DAOKobetsuKekkahyo.Kibokaisou In kibokaisouList
            ExecuteSyukei_Chiiki(db, syukeiInfo, einouChusyutsuJouken, kaisouInfo)
        Next

    End Sub

    ''' <summary>
    ''' 【営農】集計処理：地域ループ
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="syukeiInfo"></param>
    ''' <param name="dtKobetsuItemMst"></param>
    ''' <param name="dtSyukeiItemMst"></param>
    ''' <param name="dtCreateRonri"></param>
    ''' <param name="einouChusyutsuJouken"></param>
    ''' <param name="kaisouInfo"></param>
    ''' <remarks></remarks>
    Private Overloads Sub ExecuteSyukei_Chiiki(db As DBAccess, syukeiInfo As DAOKobetsuKekkahyo.SyukeiInfo, einouChusyutsuJouken As DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu, kaisouInfo As DAOKobetsuKekkahyo.Kibokaisou)
        Dim pKey As DAOSyukeiKekkahyo.PrimaryKey
        Dim kKey As DAOSyukeiKekkahyo.KomokuKey
        '区分関係
        Dim tahataKbn As String
        Dim beerKbn As String
        Dim tensaiKbn As String

        '地域毎の集計
        For Each chiikiKbn As String In _chiiki.Keys
            '地域
            syukeiInfo.chiiki = ComConst.地域.リスト(chiikiKbn).都道府県
            '平均種類
            syukeiInfo.heikinSyurui = cboHeikinSyurui.SelectedValue.ToString
            '規模階層
            syukeiInfo.kiboKaisou = kaisouInfo.kaisouNo
            '集計名称
            syukeiInfo.syukeiName = getSyukeiName(syukeiInfo, einouChusyutsuJouken)
            ' REV_004↓
            If kaisouInfo.syukeiCd IsNot Nothing AndAlso kaisouInfo.syukeiCd.Contains("-") Then
                '基本・詳細項目集計
                syukeiInfo.kihonSyosaiSyukei = ComConst.基本詳細項目集計.getValue(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun), kaisouInfo.syukeiCd.ToString)
            End If
            ' REV_004↑


            '集計結果表テーブル主キー設定
            '営農は生産費平均値種類・田畑区分・ビール麦販売区分・てんさい区分は"0"固定
            syukeiInfo.seisanhiHeikin = CODE_UNSELECTED
            tahataKbn = CODE_UNSELECTED
            beerKbn = CODE_UNSELECTED
            tensaiKbn = CODE_UNSELECTED

            pKey = New DAOSyukeiKekkahyo.PrimaryKey(cboChosaNen.SelectedValue.ToString, syukeiInfo.syukeiNo, CStr(syukeiInfo.groupKey))

            kKey = New DAOSyukeiKekkahyo.KomokuKey(CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center, cboHeikinSyurui.SelectedValue.ToString, syukeiInfo.kiboKaisou,
                                                   syukeiInfo.seisanhiHeikin, tahataKbn, beerKbn, tensaiKbn, syukeiInfo.syukeiName, getSyukeiJouken(kaisouInfo.kaisouNo), chiikiKbn, syukeiInfo.bumon)



            '集計結果表作成実行
            Dim itemInfoDic As Dictionary(Of String, CreateSyukei.ItemInfo) = _createSyukei.Execute(_progressDialog, chiikiKbn, kaisouInfo, einouChusyutsuJouken, Nothing, syukeiInfo)

            If Not itemInfoDic Is Nothing Then
                '集計結果表で抽出
                Dim ItemInfoListSyukei As List(Of CreateSyukei.ItemInfo) = (From n In itemInfoDic.Values Where n.ItemType = CreateSyukei.enmItemType.集計結果表 And Not n.ItemNo.Contains("前") And n.IsHidden = False).ToList

                '集計結果表レコードを作成
                createSyukeiKekkahyouRecord(db, pKey, kKey, ItemInfoListSyukei)

                '連番をインクリメントする
                syukeiInfo.groupKey += 1
            End If

        Next

    End Sub

    ''' <summary>
    ''' 【生産費】集計処理：地域ループ
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="syukeiInfo"></param>
    ''' <param name="dtKobetsuItemMst"></param>
    ''' <param name="dtSyukeiItemMst"></param>
    ''' <param name="dtCreateRonriHeikinSyurui"></param>
    ''' <param name="seisanhiChusyutsuJouken"></param>
    ''' <param name="kaisouInfo"></param>
    ''' <remarks></remarks>
    Private Overloads Sub ExecuteSyukei_Chiiki(db As DBAccess, syukeiInfo As DAOKobetsuKekkahyo.SyukeiInfo, seisanhiChusyutsuJouken As DAOKobetsuKekkahyo.SeisanhiChusyutsu, kaisouInfo As DAOKobetsuKekkahyo.Kibokaisou)
        Dim pKey As DAOSyukeiKekkahyo.PrimaryKey
        Dim kKey As DAOSyukeiKekkahyo.KomokuKey
        '区分関係
        Dim tahataKbn As String
        Dim beerKbn As String
        Dim tensaiKbn As String

        '地域毎の集計
        For Each chiikiKbn As String In _chiiki.Keys
            '地域
            syukeiInfo.chiiki = ComConst.地域.リスト(chiikiKbn).都道府県
            '平均種類
            syukeiInfo.heikinSyurui = cboHeikinSyurui.SelectedValue.ToString
            '規模階層
            syukeiInfo.kiboKaisou = kaisouInfo.kaisouNo
            '集計名称
            syukeiInfo.syukeiName = getSyukeiName(syukeiInfo, Nothing)
            tahataKbn = If(seisanhiChusyutsuJouken Is Nothing, CODE_UNSELECTED, seisanhiChusyutsuJouken.tahataKbn)
            beerKbn = If(seisanhiChusyutsuJouken Is Nothing, CODE_UNSELECTED, seisanhiChusyutsuJouken.beerKbn)
            tensaiKbn = If(seisanhiChusyutsuJouken Is Nothing, CODE_UNSELECTED, seisanhiChusyutsuJouken.tensaiKbn)


            '集計結果表テーブル主キー設定
            pKey = New DAOSyukeiKekkahyo.PrimaryKey(cboChosaNen.SelectedValue.ToString, syukeiInfo.syukeiNo, CStr(syukeiInfo.groupKey))

            kKey = New DAOSyukeiKekkahyo.KomokuKey(CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center, cboHeikinSyurui.SelectedValue.ToString, syukeiInfo.kiboKaisou,
                                                   syukeiInfo.seisanhiHeikin, tahataKbn, beerKbn, tensaiKbn, syukeiInfo.syukeiName,
                                                   getSyukeiJouken(kaisouInfo.kaisouNo), chiikiKbn, syukeiInfo.bumon)


            '生産費平均値種類２～５のとき、総数のレコードを取得
            Dim sousuRecord As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)) = Nothing
            If Not syukeiInfo.seisanhiHeikin.Equals(KBN_SOUSU) Then
                sousuRecord = (From n In _syukeiRecordList Where n.Item1.chosaNen = cboChosaNen.SelectedValue.ToString _
                                                      And n.Item1.syukeiNo = syukeiInfo.syukeiNo _
                                                      And n.Item2.heikinSyurui = syukeiInfo.heikinSyurui _
                                                      And n.Item2.kiboKaisou = syukeiInfo.kiboKaisou _
                                                      And n.Item2.seisanhiHeikin = KBN_SOUSU _
                                                      And n.Item2.tahataKbn = tahataKbn _
                                                      And n.Item2.beerKbn = beerKbn _
                                                      And n.Item2.tensaiKbn = tensaiKbn _
                                                      And n.Item2.chiikiCd = chiikiKbn).FirstOrDefault
            End If

            '集計結果表作成実行
            Dim itemInfoDic As Dictionary(Of String, CreateSyukei.ItemInfo) = _createSyukei.Execute(_progressDialog, chiikiKbn, kaisouInfo, Nothing, seisanhiChusyutsuJouken, syukeiInfo, sousuRecord.Item3)


            If Not itemInfoDic Is Nothing Then
                '集計結果表(前年値、裏項番を含まない）で抽出
                'REV_003↓ 出力用編集で使用するため、裏項番を含めるように修正
                'Dim ItemInfoListSyukei As List(Of CreateSyukei.ItemInfo) = (From n In itemInfoDic.Values Where n.ItemType = CreateSyukei.enmItemType.集計結果表 And Not n.ItemNo.Contains("前") And n.IsHidden = False).ToList
                Dim ItemInfoListSyukei As List(Of CreateSyukei.ItemInfo) = (From n In itemInfoDic.Values Where n.ItemType = CreateSyukei.enmItemType.集計結果表 And Not n.ItemNo.Contains("前")).ToList
                'REV_003↑
                '集計結果表レコードを作成
                createSyukeiKekkahyouRecord(db, pKey, kKey, ItemInfoListSyukei)

                '連番をインクリメントする
                syukeiInfo.groupKey += 1
            End If
        Next

    End Sub

    ''' <summary>
    ''' 【生産費】集計処理：生産費平均値種類マージ版作成
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="syukeiInfo"></param>
    ''' <param name="dtKobetsuItemMst"></param>
    ''' <param name="dtSyukeiItemMst"></param>
    ''' <param name="seisanhiChusyutsuJouken"></param>
    ''' <remarks></remarks>
    Private Sub ExecuteSyukei_Merge(db As DBAccess, syukeiInfo As DAOKobetsuKekkahyo.SyukeiInfo, dtKobetsuItemMst As DataTable, dtSyukeiItemMst As DataTable, seisanhiChusyutsuJouken As DAOKobetsuKekkahyo.SeisanhiChusyutsu, kaisouInfo As DAOKobetsuKekkahyo.Kibokaisou)
        Dim pKey As DAOSyukeiKekkahyo.PrimaryKey
        Dim kKey As DAOSyukeiKekkahyo.KomokuKey
        '区分関係
        Dim tahataKbn As String
        Dim beerKbn As String
        Dim tensaiKbn As String
        '地域毎の集計
        For Each chiikiKbn As String In _chiiki.Keys
            '地域
            syukeiInfo.chiiki = ComConst.地域.リスト(chiikiKbn).都道府県
            '平均種類
            syukeiInfo.heikinSyurui = cboHeikinSyurui.SelectedValue.ToString
            '規模階層
            syukeiInfo.kiboKaisou = kaisouInfo.kaisouNo
            '集計名称
            syukeiInfo.syukeiName = getSyukeiName(syukeiInfo, Nothing)
            tahataKbn = If(seisanhiChusyutsuJouken Is Nothing, CODE_UNSELECTED, seisanhiChusyutsuJouken.tahataKbn)
            beerKbn = If(seisanhiChusyutsuJouken Is Nothing, CODE_UNSELECTED, seisanhiChusyutsuJouken.beerKbn)
            tensaiKbn = If(seisanhiChusyutsuJouken Is Nothing, CODE_UNSELECTED, seisanhiChusyutsuJouken.tensaiKbn)


            '集計結果表テーブル主キー設定
            pKey = New DAOSyukeiKekkahyo.PrimaryKey(cboChosaNen.SelectedValue.ToString, syukeiInfo.syukeiNo, CStr(syukeiInfo.groupKey))

            kKey = New DAOSyukeiKekkahyo.KomokuKey(CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center, cboHeikinSyurui.SelectedValue.ToString, syukeiInfo.kiboKaisou,
                                                   syukeiInfo.seisanhiHeikin, tahataKbn, beerKbn, tensaiKbn, syukeiInfo.syukeiName,
                                                   getSyukeiJouken(kaisouInfo.kaisouNo), chiikiKbn, syukeiInfo.bumon)

            '生産費平均値種類１～５のレコードを取得
            Dim mergeTargetRecordList As List(Of ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)))
            mergeTargetRecordList = (From n In _syukeiRecordList Where n.Item1.chosaNen = cboChosaNen.SelectedValue.ToString _
                                                  And n.Item1.syukeiNo = syukeiInfo.syukeiNo _
                                                  And n.Item2.heikinSyurui = syukeiInfo.heikinSyurui _
                                                  And n.Item2.kiboKaisou = syukeiInfo.kiboKaisou _
                                                  And n.Item2.tahataKbn = tahataKbn _
                                                  And n.Item2.beerKbn = beerKbn _
                                                  And n.Item2.tensaiKbn = tensaiKbn _
                                                  And n.Item2.chiikiCd = chiikiKbn).ToList

            '集計結果表作成実行（マージデータ）
            Dim itemInfoDic As Dictionary(Of String, CreateSyukei.ItemInfo) = _createSyukei.ExecuteMerge(_progressDialog, mergeTargetRecordList)

            If Not itemInfoDic Is Nothing Then
                '集計結果表(前年値、裏項番を含まない）で抽出
                Dim ItemInfoListSyukei As List(Of CreateSyukei.ItemInfo) = (From n In itemInfoDic.Values Where n.ItemType = CreateSyukei.enmItemType.集計結果表 And Not n.ItemNo.Contains("前") And n.IsHidden = False).ToList

                '集計結果表レコードを作成
                createSyukeiKekkahyouRecord(db, pKey, kKey, ItemInfoListSyukei)

                '連番をインクリメントする
                syukeiInfo.groupKey += 1
            End If
        Next
    End Sub


    ''' <summary>
    '''集計結果表レコードを作成する
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="pkey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="iteminfoListSyukei"></param>
    ''' <remarks></remarks>
    Private Sub createSyukeiKekkahyouRecord(db As DBAccess, pkey As DAOSyukeiKekkahyo.PrimaryKey, kKey As DAOSyukeiKekkahyo.KomokuKey, iteminfoListSyukei As List(Of CreateSyukei.ItemInfo))
        Dim dcSyukei As Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)

        dcSyukei = New Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)
        For Each info In iteminfoListSyukei
            Dim item As New DAOSyukeiKekkahyo.集計結果表項目
            With item
                .値 = If(info.Value Is Nothing OrElse info.Value.Count = 0 OrElse info.Value(0) Is Nothing OrElse IsDBNull(info.Value(0)), Nothing, info.Value(0).ToString)
                .型区分 = info.ValueType
            End With
            dcSyukei.Add(info.ItemNo, item)
        Next
        _syukeiRecordList.Add(ValueTuple.Create(pkey, kKey, dcSyukei))

    End Sub

    ''' <summary>
    ''' 集計結果表レコードをInsertする
    ''' </summary>
    ''' <param name="db"></param>
    ''' <remarks></remarks>
    Private Sub insertSyukeiKekkahyou(db As DBAccess)
        '調査区分取得
        Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

        Try
            DAOSyukeiKekkahyo.InsertTable(db, chosakubun, _syukeiRecordList)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub btnChiiki_Click(sender As Object, e As EventArgs) Handles btnChiiki.Click
        Try
            Dim frm As New BRA5120F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnBumon_Click(sender As Object, e As EventArgs) Handles btnBumon.Click
        Try
            Dim frm As New BRA5130F()
            frm.Show(Me)
            Me.Hide()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 地域設定
    ''' </summary>
    ''' <param name="dc"></param>
    ''' <remarks></remarks>
    Public Sub SetChiiki(dc As Dictionary(Of String, String))
        _chiiki = dc
        txtChiiki.Text = String.Empty

        For Each kv As KeyValuePair(Of String, String) In dc
            If txtChiiki.Text = String.Empty Then
                txtChiiki.Text = kv.Value
            Else
                txtChiiki.Text &= SELECT_ITEM_DELIMITER
                txtChiiki.Text &= kv.Value
            End If
        Next
    End Sub

    ''' <summary>
    ''' 部門設定
    ''' </summary>
    ''' <param name="dc"></param>
    ''' <remarks></remarks>
    Public Sub SetBumon(dc As Dictionary(Of String, String))
        _bumon = dc
        txtBumon.Text = String.Empty

        For Each kv As KeyValuePair(Of String, String) In dc
            If txtBumon.Text = String.Empty Then
                txtBumon.Text = kv.Value
            Else
                txtBumon.Text &= SELECT_ITEM_DELIMITER
                txtBumon.Text &= kv.Value
            End If
        Next

        If txtBumon.Text.Length > 0 Then
            ' REV_004↓
            ComUtil.SetSyukei1ComboBox(cboEinouKeieitai, cboSyukei1, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), ComUtil.GetChosakubun(cboEinouKeieitai)))
            ' REV_004↑
        End If

    End Sub

    ''' <summary>
    ''' 集計条件文字列を作成する
    ''' </summary>
    ''' <param name="pKaisou"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getSyukeiJouken(pKaisou As String) As String
        Dim joken As New StringBuilder

        With joken
            .Append("平均種類：")
            .Append(ComConst.平均種類.リスト(cboHeikinSyurui.SelectedValue.ToString).名称)
            If _dispKiboKaisouFlg Then
                .Append(", ")
                .Append("規模階層：")
                .Append(ComConst.規模階層.リスト(cboKiboKaisou.SelectedValue.ToString))
            End If
            If ComUtil.IsEinou Then

                If Not cboSyukei1.SelectedValue Is Nothing Then
                    ' REV_004↓
                    .Append(", ")
                    .Append("集計１：")
                    .Append(ComConst.集計１.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))(cboSyukei1.SelectedValue.ToString).名称)
                    .Append(", ")
                    .Append("集計２：")
                    .Append(ComConst.集計２.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))(cboSyukei2.SelectedValue.ToString).名称)
                    .Append(", ")
                    .Append("集計３：")
                    .Append(ComConst.集計３.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))(cboSyukei3.SelectedValue.ToString).名称)
                    .Append(", ")
                    .Append("集計４：")
                    .Append(ComConst.集計４.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))(cboSyukei4.SelectedValue.ToString).名称)
                    ' REV_004↑
                End If
            Else
                If Not cboSyukei1.SelectedValue Is Nothing Then
                    .Append(", ")
                    .Append("集計区分：")
                    .Append(If(cboSyukei1.SelectedValue Is Nothing, String.Empty, ComConst.集計区分.リスト(cboSyukei1.SelectedValue.ToString)))
                End If
                If Not cboSyukei2.SelectedValue Is Nothing Then
                    .Append(", ")
                    .Append("田畑区分：")
                    .Append(If(cboSyukei2.SelectedValue Is Nothing, String.Empty, ComConst.田畑区分.リスト(cboSyukei2.SelectedValue.ToString)))
                End If
                If Not cboSyukei3.SelectedValue Is Nothing Then
                    .Append(", ")
                    .Append("ビール麦販売区分：")
                    .Append(If(cboSyukei3.SelectedValue Is Nothing, String.Empty, ComConst.ビール麦販売区分.リスト(cboSyukei3.SelectedValue.ToString)))
                End If
                If Not cboSyukei4.SelectedValue Is Nothing Then
                    .Append(", ")
                    .Append("てんさい栽培区分：")
                    .Append(If(cboSyukei4.SelectedValue Is Nothing, String.Empty, ComConst.てんさい栽培区分.リスト(cboSyukei4.SelectedValue.ToString)))
                End If

            End If

            'REV_006↓
            ' 継続区分を追加
            If cboKeizokuKubun.Visible AndAlso cboKeizokuKubun.SelectedValue IsNot Nothing Then
                .Append(", ")
                .Append("継続区分：")
                .Append(ComConst.継続区分.リスト(cboKeizokuKubun.SelectedValue.ToString))
            End If
            'REV_006↑
        End With

        Return joken.ToString
    End Function

    ''' <summary>
    ''' 【営農】集計名称の末尾に集計条件を付与する
    ''' </summary>
    ''' <param name="syukeiInfo"></param>
    ''' <param name="einouChusyutsuJouken"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getSyukeiName(syukeiInfo As DAOKobetsuKekkahyo.SyukeiInfo, einouChusyutsuJouken As DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu) As String
        Dim syukeiName As New StringBuilder
        Const MAX_INDEX As Integer = 69

        With syukeiName
            .Append(txtSyukeiName.Text)

            If Not einouChusyutsuJouken Is Nothing Then
                If Not einouChusyutsuJouken.syukei1 Is Nothing Then
                    ' REV_004↓
                    .Append(ComConst.集計２.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))(einouChusyutsuJouken.syukei2).名称)
                    .Append(ComConst.集計３.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))(einouChusyutsuJouken.syukei3).名称)
                    .Append(ComConst.集計４.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))(einouChusyutsuJouken.syukei4).名称)
                    ' REV_004↑
                ElseIf Not einouChusyutsuJouken.bumonCd Is Nothing Then
                    .Append(ComConst.部門.リスト(einouChusyutsuJouken.bumonCd).名称)
                End If

            End If

        End With

        '連結結果が70文字以上の場合切り捨てる
        Return If(syukeiName.Length > MAX_INDEX + 1, syukeiName.ToString.Remove(MAX_INDEX), syukeiName.ToString)

    End Function


    ''' <summary>
    ''' 集計条件１～４(任意集計）文字列を繋げて一つの条件にする
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getNiniSyukei() As String
        Dim niniSyukei As New StringBuilder

        If dgvList.Rows.Count > 0 Then
            With niniSyukei
                For i As Integer = 0 To dgvList.Columns.Count - 1
                    If dgvList.Rows(0).Cells(i).Value Is Nothing Then
                        Exit For
                    End If
                    If i = 0 Then
                        .Append("(")
                    Else
                        .Append(" ")
                        .Append(STR_AND)
                        .Append(" (")
                    End If
                    For j As Integer = 0 To dgvList.Rows.Count - 1
                        If dgvList.Rows(j).Cells(i).Value Is Nothing Then
                            Exit For
                        End If
                        If j <> 0 Then
                            .Append(" ")
                            .Append(STR_OR)
                            .Append(" ")
                        End If
                        .Append(dgvList.Rows(j).Cells(i).Value)
                    Next
                    .Append(")")
                Next
            End With
            '実行可能文字列に置換

            '全角不等号『＜』『＞』 ⇒ 半角不等号『<』『>』
            niniSyukei = niniSyukei.Replace("＜", "<").Replace("＞", ">")
            '全角等号付き不等号『≦』『≧』 ⇒ 『<=』『>=』
            niniSyukei = niniSyukei.Replace("≦", "<=").Replace("≧", ">=")
            '全角等号『＝』 ⇒ 半角等号『=』
            niniSyukei = niniSyukei.Replace("＝", "=")
            '全角等号否定『≠』 ⇒ 半角山括弧『<>』
            niniSyukei = niniSyukei.Replace("≠", "<>")
            '全角記号『＋－÷×』 ⇒ 半角記号『+-/*』
            niniSyukei = niniSyukei.Replace("＋", "+").Replace("－", "- ").Replace("-", "- ").Replace("÷", "/").Replace("×", "*")

            '項番にISNULLを付与
            niniSyukei = niniSyukei.Replace("[", "ISNULL([")
            niniSyukei = niniSyukei.Replace("]", "], 0)")
        End If

        Return niniSyukei.ToString
    End Function

    ''' <summary>
    ''' 集計条件１～４(任意集計）文字列を繋げる（指標部への表示用）
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getNiniSyukeiUseShihyou() As String
        Dim niniSyukei As New StringBuilder

        If dgvList.Rows.Count > 0 Then
            With niniSyukei
                For i As Integer = 0 To dgvList.Columns.Count - 1
                    If dgvList.Rows(0).Cells(i).Value Is Nothing Then
                        Exit For
                    End If
                    If i <> 0 Then
                        .Append("：")
                    End If
                    For j As Integer = 0 To dgvList.Rows.Count - 1
                        If dgvList.Rows(j).Cells(i).Value Is Nothing Then
                            Exit For
                        End If
                        If j <> 0 Then
                            .Append(" ")
                        End If
                        .Append(dgvList.Rows(j).Cells(i).Value)
                    Next
                Next
            End With
        End If

        Return niniSyukei.ToString
    End Function


    Private Sub dgvList_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvList.CurrentCellDirtyStateChanged
        If dgvList.CurrentCellAddress.X = 0 AndAlso dgvList.IsCurrentCellDirty Then
            dgvList.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    ''' <summary>
    ''' 集計条件読込ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRead_Click(sender As Object, e As EventArgs) Handles btnRead.Click
        Try
            '調査年選択チェック
            If cboChosaNen.SelectedValue Is Nothing Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of OpenFileDialog)(Me, IniFileInfo.ExcelInPath, , ComConst.CSVファイル.FILEDIALOG_CSV_FILEFILTER)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '入力ファイル取得
            Dim lstArr As List(Of String()) = Me.GetInputFile(filePath)

            'エラーチェック
            'ファイル形式チェック
            If lstArr.Count <> 2 OrElse lstArr(0).Length <> 2 OrElse lstArr(1).Length <> 13 Then
                Message.ShowMsgBox(MessageID.MSG_E_019, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            Dim details As New List(Of String)
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                If Not Me.CheckFileFormatSyukeiJouken(db, lstArr, details) Then
                    'エラーメッセージ
                    Message.ShowMsgForm(Me, MessageID.MSG_E_010, {String.Join(vbCrLf, details)})
                    Exit Sub
                End If
            End Using

            'ファイルの内容を画面にセット
            Me.setSyukeiJouken(lstArr)

            '完了メッセージ
            Message.ShowMsgBox(MessageID.MSG_I_028, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 集計条件出力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSyukeiJouken_Click(sender As Object, e As EventArgs) Handles btnSyukeiJouken.Click
        Try
            '調査年選択チェック
            If cboChosaNen.SelectedValue Is Nothing Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '調査区分名取得
            Dim chosakubunName As String = ComUtil.GetChosakubunName(ComUtil.GetChosakubun(cboEinouKeieitai))

            'フォルダパス取得
            Dim fileName As String = SYUKEI_JOUKEN_FILE_TITLE & "_" & chosakubunName & "_" & CommonInfo.Kyoku & ".csv"
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName, ComConst.CSVファイル.FILEDIALOG_CSV_FILEFILTER)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor


            '集計条件１～４妥当性チェック
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Dim errLst As List(Of String) = isExecutableNiniSyukei(db)

                If errLst.Count > 0 Then
                    Message.ShowMsgForm(Me, MessageID.MSG_E_010, {String.Join(vbCrLf, errLst)})
                    Exit Sub
                End If
            End Using

            'データ取得
            Dim dc As Dictionary(Of String, String) = Me.getSyukeiJoukenData()

            'CSVファイル出力
            Dim ret As ComConst.CSVファイル.enmOutputReturn = Me.PutOutputFile(filePath, dc)

            Select Case ret
                Case ComConst.CSVファイル.enmOutputReturn.OK
                    '完了メッセージ
                    Message.ShowMsgBox(MessageID.MSG_I_029, MsgBoxStyle.OkOnly)
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
    ''' センサス番号取込ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnInputCensus_Click(sender As Object, e As EventArgs) Handles btnInputCensus.Click
        Try
            '調査年選択チェック
            If cboChosaNen.SelectedValue Is Nothing Then
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_002, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            'ファイルパス取得
            Dim filePath As String = ComUtil.GetFilePath(Of OpenFileDialog)(Me, IniFileInfo.ExcelInPath, , ComConst.CSVファイル.FILEDIALOG_CSV_FILEFILTER)

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '入力ファイル取得
            Dim lstArr As List(Of String()) = Me.GetInputFile(filePath)

            Dim chosaNen As String = cboChosaNen.SelectedValue.ToString
            Dim einouKeieitai As String = If(cboEinouKeieitai.SelectedValue Is Nothing, Nothing, cboEinouKeieitai.SelectedValue.ToString)

            Dim lst As List(Of String) = Nothing

            'エラーチェック
            Dim details As New List(Of String)
            If Not Me.CheckFileFormat(lstArr, chosaNen, einouKeieitai, lst, details) Then
                'エラーメッセージ
                Message.ShowMsgForm(Me, MessageID.MSG_E_010, {String.Join(vbCrLf, details)})
                Exit Sub
            End If

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try

                    db.BeginTrans()

                    '集計用センサス番号管理削除
                    DAOOther.DeleteSyukeiCensusNo(db)

                    For Each val As String In lst
                        '集計用センサス番号管理追加
                        DAOOther.InsertSyukeiCensusNo(db, chosaNen, val)
                    Next

                    db.CommitTrans()

                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using

            'DataGridView初期化
            dgvList.Rows.Clear()
            dgvList.Enabled = False

            '調査年産コンボボックス非活性
            cboChosaNen.Enabled = False

            '集計条件読み込みボタン非活性
            btnRead.Enabled = False

            '集計条件登録ボタン非活性
            btnSyukeiJouken.Enabled = False

            'センサス番号取込ボタン非活性
            btnInputCensus.Enabled = False

            '「センサス番号取込済」ラベル表示
            lblCensusShitei.Text = "センサス番号取込済"

            '営農経営体区分コンボボックス非活性
            cboEinouKeieitai.Enabled = False

            '完了メッセージ
            Message.ShowMsgBox(MessageID.MSG_I_030, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 集計番号存在チェック
    ''' </summary>
    ''' <param name="syukeiNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isExistSyukeiNo(db As DBAccess) As Boolean
        Dim ret As Boolean = False
        Dim syukeiNo As String = CommonInfo.Koutei & CInt(CommonInfo.Kyoku) & txtSyukeiNo.Text

        '調査区分取得
        Dim chosakubun As String
        chosakubun = ComUtil.GetChosakubun(cboEinouKeieitai)

        If Not String.IsNullOrEmpty(DAOSyukeiKekkahyo.getSyukeiNo(db, chosakubun, cboChosaNen.SelectedValue.ToString, syukeiNo)) Then
            ret = True
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 集計番号存在チェック（範囲）
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="startSyukeiNo"></param>
    ''' <param name="endSyukeiNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isExistSyukeiNoRange(db As DBAccess, startSyukeiNo As String, endSyukeiNo As String) As Boolean
        Dim ret As Boolean = False
        Dim range As Integer = CInt(endSyukeiNo) - CInt(startSyukeiNo)

        '調査区分取得
        Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

        '集計番号桁あふれのチェック
        If endSyukeiNo.Length > 6 Then
            Return True
        End If

        For i As Integer = 0 To range
            Dim syukeiNo As String = CommonInfo.Koutei & CInt(CommonInfo.Kyoku) & String.Format("{0:D6}", CInt(startSyukeiNo) + i)

            If Not String.IsNullOrEmpty(DAOSyukeiKekkahyo.getSyukeiNo(db, chosakubun, cboChosaNen.SelectedValue.ToString, syukeiNo)) Then
                ret = True
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 集計１～４選択チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isSelectedEinouChusyutsuJoken() As Boolean
        Dim ret As Boolean = False

        '集計１～４が全て選択されている または 任意階層利用有 または 部門が存在する
        If (Not cboSyukei1.SelectedValue Is Nothing) And (Not cboSyukei2.SelectedValue Is Nothing) And (Not cboSyukei3.SelectedValue Is Nothing) And (Not cboSyukei4.SelectedValue Is Nothing) _
            Or cboNiniKaisou.SelectedValue.Equals(ComConst.任意階層利用.有) _
            Or ((Not _bumon Is Nothing AndAlso _bumon.Count > 0) And cboSyukei1.SelectedValue Is Nothing) Then
            ret = True
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 集計条件１～４(任意集計）実行できる条件であるか
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isExecutableNiniSyukei(db As DBAccess) As List(Of String)
        Const kobetsuItem As String = "\[[K]+[0-9_]+\]"


        Dim msg As String() = {"" _
                             , "{0}件目：　{1}行目　集計条件（{2}）　数値、記号、括弧以外は入力できません。" _
                             , "{0}件目：　{1}行目　集計条件（{2}）　存在しない項番が入力されています。" _
                             , "{0}件目：　{1}行目　集計条件（{2}）　条件式が不正です。" _
                             , "{0}件目：　集計条件１に連続して入力がされていません。上から連続して入力してください。" _
                             , "{0}件目：　集計条件２に連続して入力がされていません。上から連続して入力してください。" _
                             , "{0}件目：　集計条件３に連続して入力がされていません。上から連続して入力してください。" _
                             , "{0}件目：　集計条件４に連続して入力がされていません。上から連続して入力してください。" _
                             , "{0}件目：　集計条件１～４に連続して入力がされていません。小さい数字から入力してください。"
        }

        'エラーメッセージリスト
        Dim details As New List(Of String)

        If dgvList.Rows.Count > 0 Then
            Dim errCnt As Integer = 0

            '列方向連続入力チェック(先頭行のみ)
            Dim emptyFlg As Boolean = False
            For i As Integer = 0 To dgvList.Columns.Count - 1
                Dim jouken As Object = dgvList.Rows(0).Cells(i).Value
                If jouken Is Nothing Then
                    emptyFlg = True
                ElseIf emptyFlg And (Not jouken Is Nothing) Then
                    '空の列があった後に条件式が入っている列が存在したらエラー
                    errCnt += 1
                    details.Add(String.Format(msg(8), errCnt))
                End If
            Next

            '行方向連続入力チェック
            For i As Integer = 0 To dgvList.Columns.Count - 1
                emptyFlg = False
                For j As Integer = 0 To dgvList.Rows.Count - 1
                    Dim jouken As Object = dgvList.Rows(j).Cells(i).Value
                    If jouken Is Nothing Then
                        emptyFlg = True
                    ElseIf emptyFlg And (Not jouken Is Nothing) Then
                        '空の行があった後に条件式が入っている行が存在したらエラー
                        errCnt += 1
                        details.Add(String.Format(msg(i + 4), errCnt))
                        Exit For
                    End If
                Next
            Next

            '調査区分取得
            Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

            '条件に使用することのできる個別結果表項目マスタを取得する
            Dim dtJoukenKobetsItemMst As DataTable
            If chosakubun.Equals(ComConst.営農経営体区分.農業経営体) Then
                ' REV_001↓
                'dtJoukenKobetsItemMst = DAOOther.GetNougyouKeieitaiKobetsuKekkahyoItemMaster(db)
                dtJoukenKobetsItemMst = DAOOther.GetNougyouKeieitaiKobetsuKekkahyoItemMaster(db, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, CommonInfo.Chosakubun))
                ' REV_001↑
            Else
                ' REV_001↓
                'dtJoukenKobetsItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, chosakubun)
                dtJoukenKobetsItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, chosakubun, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, CommonInfo.Chosakubun))
                ' REV_001↑
            End If


            '条件式妥当性チェック
            For i As Integer = 0 To dgvList.Columns.Count - 1
                For j As Integer = 0 To dgvList.Rows.Count - 1

                    Dim jouken As Object = dgvList.Rows(j).Cells(i).Value
                    If Not jouken Is Nothing Then

                        '記号をSQL実行できるよう置換
                        Dim replacedJouken As String = jouken.ToString

                        '全角不等号『＜』『＞』 ⇒ 半角不等号『<』『>』
                        replacedJouken = replacedJouken.Replace("＜", "<").Replace("＞", ">")
                        '全角等号付き不等号『≦』『≧』 ⇒ 『<=』『>=』
                        replacedJouken = replacedJouken.Replace("≦", "<=").Replace("≧", ">=")
                        '全角等号『＝』 ⇒ 半角等号『=』
                        replacedJouken = replacedJouken.Replace("＝", "=")
                        '全角等号否定『≠』 ⇒ 半角山括弧『<>』
                        replacedJouken = replacedJouken.Replace("≠", "<>")
                        '全角記号『＋－÷×』 ⇒ 半角記号『+-/*』
                        replacedJouken = replacedJouken.Replace("＋", "+").Replace("－", "- ").Replace("-", "- ").Replace("÷", "/").Replace("×", "*")

                        '①数値・記号・括弧以外の入力がないか
                        If Not validateSyukeiJouken(replacedJouken) Then
                            errCnt += 1
                            details.Add(String.Format(msg(1), errCnt, j + 1, i + 1))
                            'エラー最大件数５０件
                            If errCnt = ComConst.ERR_MESSAGE_MAX Then Return details
                        End If



                        '②存在しない項番が入力されていないか
                        Dim matchList As MatchCollection = Nothing

                        '括弧付き項目番号『[010101]』『[010101_1]』を検索
                        matchList = Regex.Matches(replacedJouken, kobetsuItem)
                        For Each mat As Match In matchList
                            '『[010101]』『[010101_1]』から項目番号『010101』『010101_1』を抽出する
                            Dim itemNo As String = Regex.Match(mat.Value, kobetsuItem).Value.Trim("["c, "]"c)

                            Dim query1 = From dr In dtJoukenKobetsItemMst Where dr("項目番号").ToString = itemNo
                            If Not query1.Count > 0 Then
                                errCnt += 1
                                details.Add(String.Format(msg(2), errCnt, j + 1, i + 1))
                                'エラー最大件数５０件
                                If errCnt = ComConst.ERR_MESSAGE_MAX Then Return details
                            End If
                        Next

                        '③条件式が正しいか(SQLでエラーにならないこと)
                        ' REV_002↓
                        'If Not DAOKobetsuKekkahyo.isExecutableSQL(db, replacedJouken, chosakubun) Then
                        If Not DAOKobetsuKekkahyo.isExecutableSQL(db, replacedJouken, chosakubun, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, chosakubun)) Then
                            ' REV_002↑
                            errCnt += 1
                            details.Add(String.Format(msg(3), errCnt, j + 1, i + 1))
                            'エラー最大件数５０件
                            If errCnt = ComConst.ERR_MESSAGE_MAX Then Return details
                        End If
                    End If
                Next
            Next
        End If

        Return details

    End Function

    ''' <summary>
    ''' 任意階層が存在するか
    ''' </summary>
    ''' <param name="db"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isExistNiniKaisou(db As DBAccess) As Boolean
        Dim ret As Boolean = False
        Dim dt As DataTable = Nothing

        ' REV_002↓
        'dt = DAOOther.GetNiniKaisou(db, ComUtil.GetChosakubun(cboEinouKeieitai))
        dt = DAOOther.GetNiniKaisou(db, ComUtil.GetChosakubun(cboEinouKeieitai), ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, ComUtil.GetChosakubun(cboEinouKeieitai)))
        ' REV_002↑

        If dt.Rows.Count > 0 Then
            ret = True
        End If

        Return ret

    End Function

    ''' <summary>
    ''' 集計番号が６桁の数値か判定する
    ''' </summary>
    ''' <param name="syukeiNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function validateSyukeiNo(syukeiNo As String) As Boolean

        Dim ret As Boolean = Regex.IsMatch(syukeiNo, "^[0-9]{6}$")

        Return ret
    End Function

    ''' <summary>
    ''' 任意集計条件（集計条件１～４）に使用不可能な文字列が含まれているか
    ''' </summary>
    ''' <param name="jouken"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function validateSyukeiJouken(ByVal jouken As String) As Boolean
        Const kobetsuItem As String = "\[[K]+[0-9_]+\]"
        Const constRonri As String = "[【]+.+[】]"
        Dim ret As Boolean = True
        'REV_005 START--------
        '集計条件不具合の修正 (要件No.13)
        Dim checkChar As Char() = {"＋"c, "+"c, "－"c, "-"c, "×"c, "*"c, "÷"c, "/"c, "＝"c, "="c, "≠"c, "＞"c, ">"c, "＜"c, "<"c, "≧"c, "≦"c, "("c, ")"c}
        'REV_005 END----------
        Dim matchList As MatchCollection = Nothing

        '条件文字列から項番を削除
        matchList = Regex.Matches(jouken, kobetsuItem)
        For Each mat As Match In matchList
            jouken = jouken.Replace(mat.ToString, "")
        Next

        '農業経営体のみ：条件文字列から【　】で囲われた固定文字列論理を削除(論理の妥当性チェックはここでは行わない)
        If ComUtil.GetChosakubun(cboEinouKeieitai).Equals(ComConst.営農経営体区分.農業経営体) Then
            matchList = Regex.Matches(jouken, constRonri)
            For Each mat As Match In matchList
                jouken = jouken.Replace(mat.ToString, "")
            Next
        End If

        '条件文字列から使用可能文字列を削除
        For Each chr As Char In checkChar
            If jouken.Contains(chr) Then
                jouken = jouken.Replace(chr, "")
            End If
        Next


        Dim splitedJouken As String() = jouken.Split()
        For Each str As String In splitedJouken
            '数字でなければ条件式内に使用不可の文字が含まれている
            If (Not String.IsNullOrWhiteSpace(str)) And (Not IsNumeric(str)) Then
                ret = False
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 生産費の個別結果表抽出条件（集計パターン）をListにして返す
    ''' </summary>
    ''' <param name="syukeiInfo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function setSeisanhiChusyutsuJouken(syukeiInfo As DAOKobetsuKekkahyo.SyukeiInfo) As List(Of DAOKobetsuKekkahyo.SeisanhiChusyutsu)

        Dim joukenList As New List(Of DAOKobetsuKekkahyo.SeisanhiChusyutsu)
        Dim syukeiKbn As String = Nothing

        '集計区分が「Nothing」 ＝ 畜産
        If Not syukeiInfo.syukei1 Is Nothing Then
            If syukeiInfo.syukei1.Equals(ComConst.集計区分.集計対象) Then
                syukeiKbn = If(CommonInfo.Chosakubun.Equals(ComConst.調査区分.米生産費統計_個別) Or CommonInfo.Chosakubun.Equals(ComConst.調査区分.米生産費統計_組織法人), "K000005 = 5", "K000006 = 6")

            ElseIf syukeiInfo.syukei1.Equals(ComConst.集計区分.全調査対象) Then
                syukeiKbn = If(CommonInfo.Chosakubun.Equals(ComConst.調査区分.米生産費統計_個別) Or CommonInfo.Chosakubun.Equals(ComConst.調査区分.米生産費統計_組織法人), "K000005 in(2, 3, 4, 5)", "K000006 in(3, 4, 5, 6)")

            End If

            '調査区分が二条大麦の場合(集計条件が田畑区分とビール麦販売区分の混合)
            If CommonInfo.Chosakubun.Equals(ComConst.調査区分.二条大麦生産費統計_個別) Or
                CommonInfo.Chosakubun.Equals(ComConst.調査区分.経営分析調査_二条大麦生産費) Then

                Dim tahataKbnJouken As New Dictionary(Of String, String)
                If syukeiInfo.syukei2.Equals(ComConst.田畑区分.田畑計) Then
                    tahataKbnJouken.Add(ComConst.田畑区分.田畑計, syukeiKbn)

                ElseIf syukeiInfo.syukei2.Equals(ComConst.田畑区分.田) Then
                    tahataKbnJouken.Add(ComConst.田畑区分.田, syukeiKbn & " AND K000005 = 1")

                ElseIf syukeiInfo.syukei2.Equals(ComConst.田畑区分.畑) Then
                    tahataKbnJouken.Add(ComConst.田畑区分.畑, syukeiKbn & " AND K000005 = 2")

                ElseIf syukeiInfo.syukei2.Equals(ComConst.田畑区分.全て) Then
                    tahataKbnJouken.Add(ComConst.田畑区分.田畑計, syukeiKbn)
                    tahataKbnJouken.Add(ComConst.田畑区分.田, syukeiKbn & " AND K000005 = 1")
                    tahataKbnJouken.Add(ComConst.田畑区分.畑, syukeiKbn & " AND K000005 = 2")
                End If

                Dim beerKbnJouken As New Dictionary(Of String, String)
                If syukeiInfo.syukei3.Equals(ComConst.ビール麦販売区分.ビール麦100) Then
                    beerKbnJouken.Add(ComConst.ビール麦販売区分.ビール麦100, " AND K000013 = 1")

                ElseIf syukeiInfo.syukei3.Equals(ComConst.ビール麦販売区分.ビール麦100未満) Then
                    beerKbnJouken.Add(ComConst.ビール麦販売区分.ビール麦100未満, " AND K000013 = 2")

                ElseIf syukeiInfo.syukei3.Equals(ComConst.ビール麦販売区分.ビール麦無) Then
                    beerKbnJouken.Add(ComConst.ビール麦販売区分.ビール麦無, " AND K000013 = 3")

                ElseIf syukeiInfo.syukei3.Equals(ComConst.ビール麦販売区分.指定しない) Then
                    beerKbnJouken.Add(ComConst.ビール麦販売区分.指定しない, String.Empty)

                ElseIf syukeiInfo.syukei3.Equals(ComConst.ビール麦販売区分.全て) Then
                    beerKbnJouken.Add(ComConst.ビール麦販売区分.ビール麦100, " AND K000013 = 1")
                    beerKbnJouken.Add(ComConst.ビール麦販売区分.ビール麦100未満, " AND K000013 = 2")
                    beerKbnJouken.Add(ComConst.ビール麦販売区分.ビール麦無, " AND K000013 = 3")
                    beerKbnJouken.Add(ComConst.ビール麦販売区分.指定しない, String.Empty)
                End If

                For Each tahataKbn As KeyValuePair(Of String, String) In tahataKbnJouken
                    For Each beerKbn As KeyValuePair(Of String, String) In beerKbnJouken
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, tahataKbn.Key, beerKbn.Key, CODE_UNSELECTED, tahataKbn.Value & beerKbn.Value))
                    Next
                Next

            Else
                If Not syukeiInfo.syukei2 Is Nothing Then

                    If syukeiInfo.syukei2.Equals(ComConst.田畑区分.田畑計) Then
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, syukeiInfo.syukei2, CODE_UNSELECTED, CODE_UNSELECTED, syukeiKbn))

                    ElseIf syukeiInfo.syukei2.Equals(ComConst.田畑区分.田) Then
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, syukeiInfo.syukei2, CODE_UNSELECTED, CODE_UNSELECTED, syukeiKbn & " AND K000005 = 1"))

                    ElseIf syukeiInfo.syukei2.Equals(ComConst.田畑区分.畑) Then
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, syukeiInfo.syukei2, CODE_UNSELECTED, CODE_UNSELECTED, syukeiKbn & " AND K000005 = 2"))

                    ElseIf syukeiInfo.syukei2.Equals(ComConst.田畑区分.全て) Then
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, ComConst.田畑区分.田畑計, CODE_UNSELECTED, CODE_UNSELECTED, syukeiKbn))
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, ComConst.田畑区分.田, CODE_UNSELECTED, CODE_UNSELECTED, syukeiKbn & " AND K000005 = 1"))
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, ComConst.田畑区分.畑, CODE_UNSELECTED, CODE_UNSELECTED, syukeiKbn & " AND K000005 = 2"))
                    End If
                End If

                If Not syukeiInfo.syukei4 Is Nothing Then
                    If syukeiInfo.syukei4.Equals(ComConst.てんさい栽培区分.移植有) Then
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, CODE_UNSELECTED, CODE_UNSELECTED, syukeiInfo.syukei4, syukeiKbn & " AND K000005 = 1"))

                    ElseIf syukeiInfo.syukei4.Equals(ComConst.てんさい栽培区分.移植無) Then
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, CODE_UNSELECTED, CODE_UNSELECTED, syukeiInfo.syukei4, syukeiKbn & " AND K000005 = 2"))

                    ElseIf syukeiInfo.syukei4.Equals(ComConst.てんさい栽培区分.指定しない) Then
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, CODE_UNSELECTED, CODE_UNSELECTED, syukeiInfo.syukei4, syukeiKbn))

                    ElseIf syukeiInfo.syukei4.Equals(ComConst.てんさい栽培区分.全て) Then
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, CODE_UNSELECTED, CODE_UNSELECTED, ComConst.てんさい栽培区分.移植有, syukeiKbn & " AND K000005 = 1"))
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, CODE_UNSELECTED, CODE_UNSELECTED, ComConst.てんさい栽培区分.移植無, syukeiKbn & " AND K000005 = 2"))
                        joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, CODE_UNSELECTED, CODE_UNSELECTED, ComConst.てんさい栽培区分.指定しない, syukeiKbn))
                    End If
                End If

                '田畑区分、ビール麦販売区分、てんさい栽培区分が全て選択されていない場合
                If syukeiInfo.syukei2 Is Nothing And syukeiInfo.syukei3 Is Nothing And syukeiInfo.syukei4 Is Nothing Then
                    joukenList.Add(New DAOKobetsuKekkahyo.SeisanhiChusyutsu(syukeiInfo.syukei1, CODE_UNSELECTED, CODE_UNSELECTED, CODE_UNSELECTED, syukeiKbn))
                End If
            End If
        Else
            '畜産の場合は条件なし
            joukenList.Add(Nothing)
        End If

        Return joukenList

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
    ''' <param name="chosaNen"></param>
    ''' <param name="einouKeieitai"></param>
    ''' <param name="lst"></param>
    ''' <param name="details"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckFileFormat(lstArr As List(Of String()), chosaNen As String, einouKeieitai As String, ByRef lst As List(Of String), ByRef details As List(Of String)) As Boolean
        Dim ret As Boolean = True
        Dim dt As DataTable = Nothing

        Const max As Integer = ComConst.ERR_MESSAGE_MAX

        Dim msg As String() = {"" _
                             , "{0}件目：{1}行目　16桁の半角数値で入力してください。" _
                             , "{0}件目：{1}行目　センサス番号が個別結果表に登録されていません。" _
                             , "{0}件目：{1}行目　センサス番号がファイル内で重複しております。"
        }

        Dim cnt As Integer = 0

        Dim row As Integer = 0

        Dim chosakubun As String = ComUtil.GetChosakubun(einouKeieitai)

        Dim lstCensusNo As List(Of String)

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            If Not chosakubun.Equals(ComConst.営農経営体区分.農業経営体) Then
                'センサス番号取得
                Select Case CommonInfo.Koutei
                    Case CommonInfo.KouteiKubun.Code.Center
                        dt = DAOKobetsuKekkahyo.GetCensusNo(db, chosakubun, chosaNen, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
                    Case CommonInfo.KouteiKubun.Code.Kyoku
                        dt = DAOKobetsuKekkahyo.GetCensusNo(db, chosakubun, chosaNen, CommonInfo.Kyoku, Nothing, Nothing)
                    Case CommonInfo.KouteiKubun.Code.Honsyo
                        dt = DAOKobetsuKekkahyo.GetCensusNo(db, chosakubun, chosaNen, Nothing, Nothing, Nothing)
                End Select
                lstCensusNo = (From dr In dt Select CStr(dr("センサス番号"))).ToList
            Else
                lstCensusNo = New List(Of String)

                For Each arr As String In {ComConst.営農経営体区分.個人経営体, ComConst.営農経営体区分.法人経営体}
                    'センサス番号取得
                    Select Case CommonInfo.Koutei
                        Case CommonInfo.KouteiKubun.Code.Center
                            dt = DAOKobetsuKekkahyo.GetCensusNo(db, arr, chosaNen, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
                        Case CommonInfo.KouteiKubun.Code.Kyoku
                            dt = DAOKobetsuKekkahyo.GetCensusNo(db, arr, chosaNen, CommonInfo.Kyoku, Nothing, Nothing)
                        Case CommonInfo.KouteiKubun.Code.Honsyo
                            dt = DAOKobetsuKekkahyo.GetCensusNo(db, arr, chosaNen, Nothing, Nothing, Nothing)
                    End Select
                    lstCensusNo.AddRange((From dr In dt Select CStr(dr("センサス番号"))).ToList)
                Next
            End If
        End Using

        lst = (From arr In lstArr Select arr(0)).ToList

        For Each val As String In lst
            row = row + 1

            '1）16桁の半角数値であるか。
            If Not Regex.IsMatch(val, "^[0-9]{16}$") Then
                cnt = cnt + 1
                details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4)))
                ret = False
                If cnt = max Then Return ret
            End If

            '2）入力されているセンサス番号が対象調査年の個別結果表に存在するか。
            If Not lstCensusNo.Contains(val) Then
                cnt = cnt + 1
                details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4)))
                ret = False
                If cnt = max Then Return ret
            End If

            '3）ファイル内でセンサス番号が重複していないか。
            Dim query = From vl In lst Where vl = val
            If query.Count > 1 Then
                cnt = cnt + 1
                details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4)))
                ret = False
                If cnt = max Then Return ret
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 【CSV出力】集計条件データ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getSyukeiJoukenData() As Dictionary(Of String, String)
        Dim ret As New Dictionary(Of String, String)

        ret("平均種類") = If(cboHeikinSyurui.SelectedValue Is Nothing, String.Empty, cboHeikinSyurui.SelectedValue.ToString)
        ret("任意階層利用") = If(cboNiniKaisou.SelectedValue Is Nothing, String.Empty, cboNiniKaisou.SelectedValue.ToString)
        ret("規模階層") = If(cboKiboKaisou.SelectedValue Is Nothing, String.Empty, cboKiboKaisou.SelectedValue.ToString)
        ret("集計１") = If(cboSyukei1.SelectedValue Is Nothing, String.Empty, cboSyukei1.SelectedValue.ToString)
        ret("集計２") = If(cboSyukei2.SelectedValue Is Nothing, String.Empty, cboSyukei2.SelectedValue.ToString)
        ret("集計３") = If(cboSyukei3.SelectedValue Is Nothing, String.Empty, cboSyukei3.SelectedValue.ToString)
        ret("集計４") = If(cboSyukei4.SelectedValue Is Nothing, String.Empty, cboSyukei4.SelectedValue.ToString)

        Dim sb As New StringBuilder
        With sb
            '地域コードをセミコロンで繋げる
            If Not _chiiki Is Nothing Then
                For i As Integer = 0 To _chiiki.Keys.Count - 1
                    If i = 0 Then
                        .Append(_chiiki.Keys(i))
                    Else
                        .Append(String.Format(";{0}", _chiiki.Keys(i)))
                    End If
                Next
            End If

            ret("地域") = .ToString
            .Clear()

            '部門コードをセミコロンで繋げる
            If Not _bumon Is Nothing Then
                For i As Integer = 0 To _bumon.Keys.Count - 1
                    If i = 0 Then
                        .Append(_bumon.Keys(i))
                    Else
                        .Append(String.Format(";{0}", _bumon.Keys(i)))
                    End If
                Next
            End If

            ret("部門") = .ToString

            '任意集計条件１～４
            If dgvList.Rows.Count > 0 Then
                For i As Integer = 0 To dgvList.Columns.Count - 1
                    .Clear()
                    '集計条件ごとにセミコロンで繋げる
                    For j As Integer = 0 To dgvList.Rows.Count - 1
                        Dim value As Object = dgvList.Rows(j).Cells(i).Value
                        If j = 0 Then
                            .Append(If(value Is Nothing, String.Empty, value.ToString()))
                        Else
                            .Append(String.Format(";{0}", If(value Is Nothing, String.Empty, value.ToString())))
                        End If
                    Next
                    ret("集計条件" & i + 1) = .ToString
                Next
            Else
                For i As Integer = 0 To dgvList.Columns.Count - 1
                    .Clear()
                    '集計条件ごとにセミコロンで繋げる
                    For j As Integer = 0 To DATAGRIDVIEW_ROW - 1
                        If j = 0 Then
                            .Append(String.Empty)
                        Else
                            .Append(String.Format(";{0}", String.Empty))
                        End If
                    Next
                    ret("集計条件" & i + 1) = .ToString
                Next
            End If

        End With


        Return ret
    End Function

    ''' <summary>
    ''' 出力ファイル作成
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <param name="lstDc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PutOutputFile(filePath As String, dc As Dictionary(Of String, String)) As ComConst.CSVファイル.enmOutputReturn
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

            '調査区分名取得
            Dim chosakubunName As String = ComUtil.GetChosakubunName(ComUtil.GetChosakubun(cboEinouKeieitai))

            'ファイル出力
            Using sw As New System.IO.StreamWriter(filePath, False, sjisEnc)
                sw.WriteLine(ComConst.CSVファイル.START_ADDITION & SYUKEI_JOUKEN_FILE_TITLE & ComConst.CSVファイル.END_ADDITION & ComConst.CSVファイル.CSV_DELIMITER _
                             & ComConst.CSVファイル.START_ADDITION & chosakubunName & ComConst.CSVファイル.END_ADDITION)

                Dim arr As Object() = {dc("平均種類"), dc("任意階層利用"), dc("規模階層"), dc("集計１"), dc("集計２"), dc("集計３"), dc("集計４"), dc("地域"), dc("部門"), dc("集計条件1"), dc("集計条件2"), dc("集計条件3"), dc("集計条件4")}
                sw.WriteLine(ComConst.CSVファイル.START_ADDITION _
                             & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, arr) _
                             & ComConst.CSVファイル.END_ADDITION)
            End Using
            ret = ComConst.CSVファイル.enmOutputReturn.OK
        Catch ex As Exception
            ret = ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 集計条件ファイルチェック
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="lstArr"></param>
    ''' <param name="details"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckFileFormatSyukeiJouken(db As DBAccess, lstArr As List(Of String()), ByRef details As List(Of String)) As Boolean
        Dim ret As Boolean = True
        Dim cnt As Integer = 0

        Dim msg As String() = {"" _
                             , "{0}件目：　ファイルタイトルが正しくありません。" _
                             , "{0}件目：　調査区分が正しくありません。" _
                             , "{0}件目：　平均種類が正しくありません。" _
                             , "{0}件目：　任意階層有無が正しくありません。" _
                             , "{0}件目：　規模階層が正しくありません。" _
                             , "{0}件目：　集計１ または、集計区分が正しくありません。" _
                             , "{0}件目：　集計２ または、田畑区分が正しくありません。" _
                             , "{0}件目：　集計３ または、ビール麦販売区分が正しくありません。" _
                             , "{0}件目：　集計４ または、てんさい栽培区分が正しくありません。" _
                             , "{0}件目：　集計１～４の組み合わせが正しくありません。" _
                             , "{0}件目：　地域が正しくありません。" _
                             , "{0}件目：　部門が正しくありません。" _
                             , "{0}件目：　集計条件１が正しくありません。" _
                             , "{0}件目：　集計条件２が正しくありません。" _
                             , "{0}件目：　集計条件３が正しくありません。" _
                             , "{0}件目：　集計条件４が正しくありません。"
        }

        '1行目：1項目(ファイルタイトル)チェック
        If Not lstArr(0)(0).Equals(SYUKEI_JOUKEN_FILE_TITLE) Then
            cnt = cnt + 1
            details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2)))
            ret = False
        End If

        '1行目：2項目(調査区分)チェック
        '調査区分名取得
        Dim chosakubunName As String = ComUtil.GetChosakubunName(ComUtil.GetChosakubun(cboEinouKeieitai))

        If Not lstArr(0)(1).Equals(chosakubunName) Then
            cnt = cnt + 1
            details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2)))
            ret = False
        End If

        Dim dcEntry As DictionaryEntry

        '2行目：1項目(平均種類)チェック
        If ComConst.平均種類.リスト.ContainsKey(lstArr(1)(0)) Then
            dcEntry.Key = lstArr(1)(0)
            dcEntry.Value = ComConst.平均種類.リスト(lstArr(1)(0)).名称

            If cboHeikinSyurui.Items.IndexOf(dcEntry) = -1 Then
                cnt = cnt + 1
                details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2)))
                ret = False
            End If
        Else
            cnt = cnt + 1
            details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2)))
            ret = False
        End If

        '2行目：2項目(任意階層利用)チェック
        If ComConst.任意階層利用.リスト.ContainsKey(lstArr(1)(1)) Then
            dcEntry.Key = lstArr(1)(1)
            dcEntry.Value = ComConst.任意階層利用.リスト(lstArr(1)(1))

            If cboNiniKaisou.Items.IndexOf(dcEntry) = -1 Then
                cnt = cnt + 1
                details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2)))
                ret = False
            End If
        Else
            cnt = cnt + 1
            details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2)))
            ret = False
        End If

        '2行目：3項目(規模階層)チェック
        If ComConst.規模階層.リスト.ContainsKey(lstArr(1)(2)) Then
            dcEntry.Key = lstArr(1)(2)
            dcEntry.Value = ComConst.規模階層.リスト(lstArr(1)(2))

            If cboKiboKaisou.Items.IndexOf(dcEntry) = -1 Then
                cnt = cnt + 1
                details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2)))
                ret = False
            End If
        Else
            cnt = cnt + 1
            details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2)))
            ret = False
        End If

        If ComUtil.IsEinou Then
            Dim cboItem1 As Object = cboSyukei1.SelectedItem
            Dim cboItem2 As Object = cboSyukei2.SelectedItem
            Dim cboItem3 As Object = cboSyukei3.SelectedItem
            Dim cboItem4 As Object = cboSyukei4.SelectedItem

            '2行目：3項目(集計１)チェック
            ' REV_004↓
            If ComConst.集計１.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun)).ContainsKey(lstArr(1)(3)) Then
                ' REV_004↑
                dcEntry.Key = lstArr(1)(3)
                ' REV_004↓
                dcEntry.Value = ComConst.集計１.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))(lstArr(1)(3)).名称
                ' REV_004↑

                If cboSyukei1.Items.IndexOf(dcEntry) = -1 Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2)))
                    ret = False
                Else
                    cboSyukei1.SelectedItem = dcEntry
                End If
            Else
                If Not lstArr(1)(3).Equals(String.Empty) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            End If

            '2行目：4項目(集計２)チェック
            ' REV_004↓
            If ComConst.集計２.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun)).ContainsKey(lstArr(1)(4)) Then
                ' REV_004↑
                dcEntry.Key = lstArr(1)(4)
                ' REV_004↓
                dcEntry.Value = ComConst.集計２.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))(lstArr(1)(4)).名称
                ' REV_004↑

                If cboSyukei2.Items.IndexOf(dcEntry) = -1 Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2)))
                    ret = False
                Else
                    cboSyukei2.SelectedItem = dcEntry
                End If
            Else
                If Not lstArr(1)(4).Equals(String.Empty) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            End If

            '2行目：5項目(集計３)チェック
            ' REV_004↓
            If ComConst.集計３.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun)).ContainsKey(lstArr(1)(5)) Then
                dcEntry.Key = lstArr(1)(5)
                ' REV_004↓
                dcEntry.Value = ComConst.集計３.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))(lstArr(1)(5)).名称

                If cboSyukei3.Items.IndexOf(dcEntry) = -1 Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2)))
                    ret = False
                Else
                    cboSyukei3.SelectedItem = dcEntry
                End If
            Else
                If Not lstArr(1)(5).Equals(String.Empty) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            End If

            '2行目：6項目(集計４)チェック
            ' REV_004↓
            If ComConst.集計４.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun)).ContainsKey(lstArr(1)(6)) Then
                ' REV_004↑
                dcEntry.Key = lstArr(1)(6)
                ' REV_004↓
                dcEntry.Value = ComConst.集計４.リスト(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))(lstArr(1)(6)).名称
                ' REV_004↑

                If cboSyukei4.Items.IndexOf(dcEntry) = -1 Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2)))
                    ret = False
                Else
                    cboSyukei4.SelectedItem = dcEntry
                End If
            Else
                If Not lstArr(1)(6).Equals(String.Empty) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            End If

            '集計１～４の組み合わせチェック
            Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

            If (Not String.IsNullOrEmpty(lstArr(1)(3))) And (Not String.IsNullOrEmpty(lstArr(1)(4))) And (Not String.IsNullOrEmpty(lstArr(1)(5))) And (Not String.IsNullOrEmpty(lstArr(1)(6))) Then
                ' REV_004↓
                If DAOOther.getEinouSyukeiJoukenMaster(db, chosakubun, lstArr(1)(3), lstArr(1)(4), lstArr(1)(5), lstArr(1)(6), ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun)).Rows.Count = 0 Then
                    ' REV_004↑
                    cnt = cnt + 1
                    details.Add(String.Format(msg(10), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            End If


            cboSyukei1.SelectedItem = cboItem1
            cboSyukei2.SelectedItem = cboItem2
            cboSyukei3.SelectedItem = cboItem3
            cboSyukei4.SelectedItem = cboItem4

        Else
            '2行目：3項目(集計区分)チェック
            If ComConst.集計区分.リスト.ContainsKey(lstArr(1)(3)) Then
                dcEntry.Key = lstArr(1)(3)
                dcEntry.Value = ComConst.集計区分.リスト(lstArr(1)(3))
                If cboSyukei1.Items.IndexOf(dcEntry) = -1 Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            Else
                If lstArr(1)(3).Equals(String.Empty) Then
                    If cboSyukei1.Items.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2)))
                        ret = False
                    End If
                Else
                    cnt = cnt + 1
                    details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            End If

            '2行目：4項目(田畑区分)チェック
            If ComConst.田畑区分.リスト.ContainsKey(lstArr(1)(4)) Then
                dcEntry.Key = lstArr(1)(4)
                dcEntry.Value = ComConst.田畑区分.リスト(lstArr(1)(4))
                If cboSyukei2.Items.IndexOf(dcEntry) = -1 Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            Else
                If lstArr(1)(4).Equals(String.Empty) Then
                    If cboSyukei2.Items.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2)))
                        ret = False
                    End If
                Else
                    cnt = cnt + 1
                    details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            End If

            '2行目：5項目(ビール麦販売区分)チェック
            If ComConst.ビール麦販売区分.リスト.ContainsKey(lstArr(1)(5)) Then
                dcEntry.Key = lstArr(1)(5)
                dcEntry.Value = ComConst.ビール麦販売区分.リスト(lstArr(1)(5))
                If cboSyukei3.Items.IndexOf(dcEntry) = -1 Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            Else
                If lstArr(1)(5).Equals(String.Empty) Then
                    If cboSyukei3.Items.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2)))
                        ret = False
                    End If
                Else
                    cnt = cnt + 1
                    details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            End If

            '2行目：6項目(てんさい栽培区分)チェック
            If ComConst.てんさい栽培区分.リスト.ContainsKey(lstArr(1)(6)) Then
                dcEntry.Key = lstArr(1)(6)
                dcEntry.Value = ComConst.てんさい栽培区分.リスト(lstArr(1)(6))
                If cboSyukei4.Items.IndexOf(dcEntry) = -1 Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            Else
                If lstArr(1)(6).Equals(String.Empty) Then
                    If cboSyukei4.Items.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2)))
                        ret = False
                    End If
                Else
                    cnt = cnt + 1
                    details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2)))
                    ret = False
                End If
            End If
        End If

        '2行目：7項目(地域)チェック
        If Not String.IsNullOrEmpty(lstArr(1)(7)) Then
            Dim chiikiList As String() = lstArr(1)(7).Split(CHR_SEPARATOR)

            For Each chiiki As String In chiikiList
                '存在しない地域コードが設定されている場合エラーとする
                If Not ComConst.地域.リスト.ContainsKey(chiiki) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(11), cnt.ToString.PadLeft(2)))
                    ret = False
                    Exit For
                End If
                '自拠点に紐づかない地域が設定されている場合エラーとする
                If CommonInfo.Koutei.Equals(CommonInfo.KouteiKubun.Code.Kyoku) Then
                    If (Not ComConst.地域.リスト(chiiki).局.Equals(Integer.Parse(CommonInfo.Kyoku).ToString)) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(11), cnt.ToString.PadLeft(2)))
                        ret = False
                        Exit For
                    End If
                End If
            Next
        End If

        '2行目：8項目(部門)チェック
        '営農個別以外なら情報を破棄する
        If CommonInfo.Chosakubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) AndAlso ComUtil.GetChosakubun(cboEinouKeieitai).Equals(ComConst.営農経営体区分.個人経営体) Then
            Dim bumonList As String() = lstArr(1)(8).Split(CHR_SEPARATOR)
            For Each bumon As String In bumonList
                If (Not String.IsNullOrEmpty(bumon)) And (Not ComConst.部門.リスト.ContainsKey(bumon)) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(12), cnt.ToString.PadLeft(2)))
                    ret = False
                    Exit For
                End If
            Next
        Else
            lstArr(1)(8) = String.Empty
        End If

        '2行目：9項目(集計条件１)チェック
        If Not lstArr(1)(9).Split(CHR_SEPARATOR).Length = DATAGRIDVIEW_ROW Then
            cnt = cnt + 1
            details.Add(String.Format(msg(13), cnt.ToString.PadLeft(2)))
            ret = False
        End If

        '2行目：10項目(集計条件２)チェック
        If Not lstArr(1)(10).Split(CHR_SEPARATOR).Length = DATAGRIDVIEW_ROW Then
            cnt = cnt + 1
            details.Add(String.Format(msg(14), cnt.ToString.PadLeft(2)))
            ret = False
        End If

        '2行目：11項目(集計条件３)チェック
        If Not lstArr(1)(11).Split(CHR_SEPARATOR).Length = DATAGRIDVIEW_ROW Then
            cnt = cnt + 1
            details.Add(String.Format(msg(15), cnt.ToString.PadLeft(2)))
            ret = False
        End If

        '2行目：12項目(集計条件４)チェック
        If Not lstArr(1)(12).Split(CHR_SEPARATOR).Length = DATAGRIDVIEW_ROW Then
            cnt = cnt + 1
            details.Add(String.Format(msg(16), cnt.ToString.PadLeft(2)))
            ret = False
        End If

        Return ret
    End Function

    Private Sub setSyukeiJouken(lstArr As List(Of String()))

        '平均種類
        cboHeikinSyurui.SelectedValue = lstArr(1)(0)

        '規模階層
        cboKiboKaisou.SelectedValue = lstArr(1)(2)

        '集計１～集計４ または 平均種類～てんさい栽培区分
        If ComUtil.IsEinou Then
            cboSyukei1.SelectedValue = lstArr(1)(3)
            cboSyukei2.SelectedValue = lstArr(1)(4)
            cboSyukei3.SelectedValue = lstArr(1)(5)
            cboSyukei4.SelectedValue = lstArr(1)(6)
        Else
            If cboSyukei1.Items.Count > 0 Then
                cboSyukei1.SelectedValue = lstArr(1)(3)
            End If
            If cboSyukei2.Items.Count > 0 Then
                cboSyukei2.SelectedValue = lstArr(1)(4)
            End If
            If cboSyukei3.Items.Count > 0 Then
                cboSyukei3.SelectedValue = lstArr(1)(5)
            End If
            If cboSyukei4.Items.Count > 0 Then
                cboSyukei4.SelectedValue = lstArr(1)(6)
            End If
        End If

        '地域
        Dim chiikiList As String() = lstArr(1)(7).Split(CHR_SEPARATOR)
        Dim dicChiiki As New Dictionary(Of String, String)
        For Each chiiki As String In chiikiList
            If Not String.IsNullOrEmpty(chiiki) Then
                dicChiiki.Add(chiiki, ComConst.地域.リスト(chiiki).名称)
            End If
        Next
        Me.SetChiiki(dicChiiki)

        '部門
        If CommonInfo.Chosakubun.Equals(ComConst.調査区分.営農類型別経営統計_個人) Then
            Dim bumonList As String() = lstArr(1)(8).Split(CHR_SEPARATOR)
            Dim dicBumon As New Dictionary(Of String, String)
            For Each bumon As String In bumonList
                If Not String.IsNullOrEmpty(bumon) Then
                    dicBumon.Add(bumon, ComConst.部門.リスト(bumon).名称)
                End If
            Next
            Me.SetBumon(dicBumon)
        End If

        '任意集計条件１～４
        If dgvList.Rows.Count > 0 Then
            Dim jouken1List As String() = lstArr(1)(9).Split(CHR_SEPARATOR)
            For i As Integer = 0 To dgvList.Rows.Count - 1
                dgvList.Rows(i).Cells(0).Value = If(String.IsNullOrEmpty(jouken1List(i)), Nothing, jouken1List(i))
            Next
            Dim jouken2List As String() = lstArr(1)(10).Split(CHR_SEPARATOR)
            For i As Integer = 0 To dgvList.Rows.Count - 1
                dgvList.Rows(i).Cells(1).Value = If(String.IsNullOrEmpty(jouken2List(i)), Nothing, jouken2List(i))
            Next
            Dim jouken3List As String() = lstArr(1)(11).Split(CHR_SEPARATOR)
            For i As Integer = 0 To dgvList.Rows.Count - 1
                dgvList.Rows(i).Cells(2).Value = If(String.IsNullOrEmpty(jouken3List(i)), Nothing, jouken3List(i))
            Next
            Dim jouken4List As String() = lstArr(1)(12).Split(CHR_SEPARATOR)
            For i As Integer = 0 To dgvList.Rows.Count - 1
                dgvList.Rows(i).Cells(3).Value = If(String.IsNullOrEmpty(jouken4List(i)), Nothing, jouken4List(i))
            Next

        End If

        '任意階層利用(集計１～４や部門の活性非活性に関わる為最後)
        cboNiniKaisou.SelectedValue = lstArr(1)(1)
    End Sub

    ''' <summary>
    ''' 集計結果表の作成レコード数を求める
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="syukeiInfo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getSyukeiCount(db As DBAccess, syukeiInfo As DAOKobetsuKekkahyo.SyukeiInfo, dtCreateRonri As DataTable) As Integer
        Dim count As Integer = 0

        If ComUtil.IsEinou Then
            '【営農】

            '個別結果表抽出条件を取得
            Dim dtEinouChusyutsuJouken As New DataTable
            If Not syukeiInfo.syukei1 Is Nothing Then
                ' REV_004↓
                dtEinouChusyutsuJouken = DAOOther.getEinouSyukeiJoukenMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai), syukeiInfo.syukei1, syukeiInfo.syukei2, syukeiInfo.syukei3, syukeiInfo.syukei4, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))
                ' REV_004↑
            End If

            Dim einouChusyutsuJoukenList As New List(Of DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu)

            '営農経営体抽出条件がある場合：抽出条件とそれに紐づく階層コード。 ない場合：部門コード
            If dtEinouChusyutsuJouken.Rows.Count > 0 Then

                '親子関係2段階まで対応
                For Each row As DataRow In dtEinouChusyutsuJouken.Rows
                    '『a-xxx ～ a-xxx までの全ての集計を行う』
                    If row.Item("個別結果表抽出条件").ToString.Contains(CHR_SEPARATOR) Then

                        For Each str As String In row.Item("個別結果表抽出条件").ToString.Split(CHR_SEPARATOR)
                            '抽出条件を分割して取得した集計コードからさらに集計条件マスタを取得
                            ' REV_004↓
                            Dim einouChusyutsuJouken = DAOOther.getEinouSyukeiJoukenMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai), str, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun)).Rows(0)
                            ' REV_004↑
                            Dim syukeiCd As String = einouChusyutsuJouken.Item("集計コード").ToString
                            Dim syukei1 As String = einouChusyutsuJouken.Item("集計１").ToString
                            Dim syukei2 As String = einouChusyutsuJouken.Item("集計２").ToString
                            Dim syukei3 As String = einouChusyutsuJouken.Item("集計３").ToString
                            Dim syukei4 As String = einouChusyutsuJouken.Item("集計４").ToString
                            Dim jouken As String = einouChusyutsuJouken.Item("個別結果表抽出条件").ToString
                            ' REV_004↓
                            If ComConst.基本詳細項目集計.IsShosaiOnly(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun), einouChusyutsuJouken.Item("集計コード").ToString) Then
                                jouken = "K000026=2 and " + jouken
                            End If
                            ' REV_004↑
                            Dim kaisouCd As String = einouChusyutsuJouken.Item("階層コード").ToString
                            Dim einouruiKei As String = einouChusyutsuJouken.Item("営農類型").ToString
                            einouChusyutsuJoukenList.Add(New DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu(syukeiCd, syukei1, syukei2, syukei3, syukei4, jouken, kaisouCd, Nothing, einouruiKei))
                        Next
                    Else
                        'そのままリストを作成
                        Dim syukeiCd As String = row.Item("集計コード").ToString
                        Dim syukei1 As String = row.Item("集計１").ToString
                        Dim syukei2 As String = row.Item("集計２").ToString
                        Dim syukei3 As String = row.Item("集計３").ToString
                        Dim syukei4 As String = row.Item("集計４").ToString
                        Dim jouken As String = row.Item("個別結果表抽出条件").ToString
                        Dim kaisouCd As String = row.Item("階層コード").ToString
                        Dim einouruiKei As String = row.Item("営農類型").ToString
                        einouChusyutsuJoukenList.Add(New DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu(syukeiCd, syukei1, syukei2, syukei3, syukei4, jouken, kaisouCd, Nothing, einouruiKei))
                    End If
                Next
            Else
                If _bumon Is Nothing OrElse _bumon.Keys.Count = 0 Then
                    '集計１～４なし、部門もなし
                    einouChusyutsuJoukenList.Add(New DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing))

                Else
                    '部門別集計の場合
                    For Each bumonCd As String In _bumon.Keys
                        '部門コードから営農抽出条件マスタを取得
                        ' REV_004↓
                        Dim einouChusyutsuJouken = DAOOther.getEinouSyukeiJoukenMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai), bumonCd, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun), True).Rows(0)
                        ' REV_004↑
                        Dim kaisouCd As String = einouChusyutsuJouken.Item("階層コード").ToString
                        Dim jouken As String = einouChusyutsuJouken.Item("個別結果表抽出条件").ToString
                        ' REV_004↓
                        If ComConst.基本詳細項目集計.IsShosaiOnly(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun), einouChusyutsuJouken.Item("集計コード").ToString) Then
                            jouken = "K000026=2 and " + jouken
                        End If
                        ' REV_004↑
                        Dim einouruiKei As String = einouChusyutsuJouken.Item("営農類型").ToString
                        einouChusyutsuJoukenList.Add(New DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu(Nothing, Nothing, Nothing, Nothing, Nothing, jouken, kaisouCd, bumonCd, einouruiKei))
                    Next
                End If
            End If

            '個別結果表抽出条件毎
            For Each einouChusyutsuJouken As DAOKobetsuKekkahyo.EinouKeieitaiChusyutsu In einouChusyutsuJoukenList
                '階層コード
                Dim kibokaisouList As New List(Of DAOKobetsuKekkahyo.Kibokaisou)

                '平均値を求める際に必要な情報
                kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(CODE_UNSELECTED, Nothing, Nothing, Nothing))
                If cboKiboKaisou.SelectedValue.Equals(ComConst.規模階層.規模階層含) Then

                    If cboNiniKaisou.SelectedValue.Equals(ComConst.任意階層利用.無) Then
                        '通常の規模階層を使用する場合
                        Dim dtKiboKaisou As DataTable = DAOOther.getKiboKaisouMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai), einouChusyutsuJouken.kaisouCd)

                        For Each row As DataRow In dtKiboKaisou.Rows
                            Dim kibokaisou As String = row.Item("規模階層").ToString
                            Dim itemNo As String = row.Item("項目番号").ToString
                            Dim max As String = If(IsDBNull(row.Item("上限")), Nothing, row.Item("上限").ToString)
                            Dim min As String = If(IsDBNull(row.Item("下限")), Nothing, row.Item("下限").ToString)
                            kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(kibokaisou, itemNo, max, min))
                        Next

                        '規模階層内に集計コードが指定してあるパターンの集計処理
                        If kibokaisouList.Count > 1 AndAlso kibokaisouList(1).jouken.Contains("-") Then
                            For i As Integer = 1 To kibokaisouList.Count - 1
                                '規模階層の項目番号に指定されている集計コードから集計条件を取得する
                                ' REV_004↓
                                Dim dt As DataTable = DAOOther.getEinouSyukeiJoukenMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai), kibokaisouList(i).jouken, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun))
                                If dt.Rows.Count <> 0 Then
                                    kibokaisouList(i).jouken = dt.Rows(0).Item("個別結果表抽出条件").ToString
                                    If ComConst.基本詳細項目集計.IsShosaiOnly(ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString(), CommonInfo.Chosakubun), dt.Rows(0).Item("集計コード").ToString) Then
                                        kibokaisouList(i).jouken = "K000026=2 and " + kibokaisouList(i).jouken
                                    End If
                                Else
                                    kibokaisouList(i).jouken = Nothing
                                End If
                                ' REV_004↑
                            Next
                        End If
                        ' REV_004↓
                        kibokaisouList.RemoveAll(Function(n) n.kaisouNo <> "0" And n.jouken Is Nothing)
                        ' REV_004↑
                    Else
                        '任意階層を使用する場合
                        ' REV_002↓
                        'Dim dtKiboKaisou As DataTable = DAOOther.GetNiniKaisou(db, ComUtil.GetChosakubun(cboEinouKeieitai))
                        Dim dtKiboKaisou As DataTable = DAOOther.GetNiniKaisou(db, ComUtil.GetChosakubun(cboEinouKeieitai), ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, ComUtil.GetChosakubun(cboEinouKeieitai)))
                        ' REV_002↑

                        For Each row As DataRow In dtKiboKaisou.Rows
                            kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(row.Item("規模階層").ToString, row.Item("項目番号").ToString, row.Item("上限").ToString, row.Item("下限").ToString))
                        Next
                    End If
                End If

                '規模階層毎の集計
                For Each kaisouInfo As DAOKobetsuKekkahyo.Kibokaisou In kibokaisouList
                    For Each chiikiKbn As String In _chiiki.Keys
                        '一つのレコード
                        count += dtCreateRonri.Rows.Count
                    Next
                Next
            Next
        Else
            '【生産費】

            '画面で選択したコンボボックスから集計パターンを取得
            Dim syukeiPetternList As List(Of DAOKobetsuKekkahyo.SeisanhiChusyutsu)
            syukeiPetternList = setSeisanhiChusyutsuJouken(syukeiInfo)

            '生産費平均値種類ごとのループ(牛乳は平均値種類「5」まで、その他生産費は「4」まで存在する)
            For seisanhiHeikin As Integer = 1 To If(CommonInfo.Chosakubun.Equals(ComConst.調査区分.牛乳生産費統計_個別) Or CommonInfo.Chosakubun.Equals(ComConst.調査区分.経営分析調査_牛乳生産費), GYUNYU_SEISANHIHEIKIN_MAX, SEISANHIHEIKIN_MAX)

                Dim dtCreateRonriHeikinSyurui As DataTable

                '作成論理の生産費平均値種類：生産費の場合、総数は「1」、総数以外は「9」
                ' REV_002↓
                If seisanhiHeikin.ToString.Equals(KBN_SOUSU) Then
                    'dtCreateRonriHeikinSyurui = DAOOther.GetSyukeiKekkahyoSakuseiRonri(db, CommonInfo.Chosakubun, KBN_SOUSU)
                    dtCreateRonriHeikinSyurui = DAOOther.GetSyukeiKekkahyoSakuseiRonri(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, CommonInfo.Chosakubun), KBN_SOUSU)
                Else
                    'dtCreateRonriHeikinSyurui = DAOOther.GetSyukeiKekkahyoSakuseiRonri(db, CommonInfo.Chosakubun, KBN_EXCEPTSOUSU)
                    dtCreateRonriHeikinSyurui = DAOOther.GetSyukeiKekkahyoSakuseiRonri(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, CommonInfo.Chosakubun), KBN_EXCEPTSOUSU)
                End If
                ' REV_002↑

                '集計パターンごとのループ
                For Each syukeiPettern As DAOKobetsuKekkahyo.SeisanhiChusyutsu In syukeiPetternList
                    '階層コード
                    Dim kibokaisouList As New List(Of DAOKobetsuKekkahyo.Kibokaisou)

                    '平均値を求める際に必要な情報
                    kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(CODE_UNSELECTED, Nothing, Nothing, Nothing))
                    If cboKiboKaisou.SelectedValue.Equals(ComConst.規模階層.規模階層含) Then

                        If cboNiniKaisou.SelectedValue.Equals(ComConst.任意階層利用.無) Then
                            '通常の規模階層を使用する場合
                            Dim dtKiboKaisou As DataTable = DAOOther.getKiboKaisouMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai))

                            For Each row As DataRow In dtKiboKaisou.Rows
                                Dim kibokaisou As String = row.Item("規模階層").ToString
                                Dim itemNo As String = row.Item("項目番号").ToString
                                Dim max As String = If(IsDBNull(row.Item("上限")), Nothing, row.Item("上限").ToString)
                                Dim min As String = If(IsDBNull(row.Item("下限")), Nothing, row.Item("下限").ToString)
                                kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(kibokaisou, itemNo, max, min))
                            Next

                        Else
                            '任意階層を使用する場合
                            ' REV_002↓
                            'Dim dtKiboKaisou As DataTable = DAOOther.GetNiniKaisou(db, ComUtil.GetChosakubun(cboEinouKeieitai))
                            Dim dtKiboKaisou As DataTable = DAOOther.GetNiniKaisou(db, ComUtil.GetChosakubun(cboEinouKeieitai), ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, ComUtil.GetChosakubun(cboEinouKeieitai)))
                            ' REV_002↑

                            For Each row As DataRow In dtKiboKaisou.Rows
                                kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(row.Item("規模階層").ToString, row.Item("項目番号").ToString, row.Item("上限").ToString, row.Item("下限").ToString))
                            Next
                        End If
                    End If

                    '規模階層毎の集計
                    For Each kaisouInfo As DAOKobetsuKekkahyo.Kibokaisou In kibokaisouList
                        '地域毎の集計
                        For Each chiikiKbn As String In _chiiki.Keys
                            '一つのレコード
                            count += dtCreateRonriHeikinSyurui.Rows.Count
                        Next
                    Next
                Next
            Next

            '生産費平均値種類０ 集計パターンごとのループ
            ' REV_002↓
            'Dim dtSyutsuryokuRonri As DataTable = DAOOther.GetSeisanhiSyutsuryokuHensyu(db, CommonInfo.Chosakubun)
            Dim dtSyutsuryokuRonri As DataTable = DAOOther.GetSeisanhiSyutsuryokuHensyu(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, CommonInfo.Chosakubun))
            ' REV_002↑
            For Each syukeiPettern As DAOKobetsuKekkahyo.SeisanhiChusyutsu In syukeiPetternList

                '階層コード
                Dim kibokaisouList As New List(Of DAOKobetsuKekkahyo.Kibokaisou)

                '平均値を求める際に必要な情報
                kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(CODE_UNSELECTED, Nothing, Nothing, Nothing))
                If cboKiboKaisou.SelectedValue.Equals(ComConst.規模階層.規模階層含) Then

                    If cboNiniKaisou.SelectedValue.Equals(ComConst.任意階層利用.無) Then
                        '通常の規模階層を使用する場合
                        Dim dtKiboKaisou As DataTable = DAOOther.getKiboKaisouMaster(db, ComUtil.GetChosakubun(cboEinouKeieitai))

                        For Each row As DataRow In dtKiboKaisou.Rows
                            Dim kibokaisou As String = row.Item("規模階層").ToString
                            Dim itemNo As String = row.Item("項目番号").ToString
                            Dim max As String = If(IsDBNull(row.Item("上限")), Nothing, row.Item("上限").ToString)
                            Dim min As String = If(IsDBNull(row.Item("下限")), Nothing, row.Item("下限").ToString)
                            kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(kibokaisou, itemNo, max, min))
                        Next

                    Else
                        '任意階層を使用する場合
                        ' REV_002↓
                        'Dim dtKiboKaisou As DataTable = DAOOther.GetNiniKaisou(db, ComUtil.GetChosakubun(cboEinouKeieitai))
                        Dim dtKiboKaisou As DataTable = DAOOther.GetNiniKaisou(db, ComUtil.GetChosakubun(cboEinouKeieitai), ComUtil.getVersionKubunTaikei(cboChosaNen.SelectedValue.ToString, ComUtil.GetChosakubun(cboEinouKeieitai)))
                        ' REV_002↑

                        For Each row As DataRow In dtKiboKaisou.Rows
                            kibokaisouList.Add(New DAOKobetsuKekkahyo.Kibokaisou(row.Item("規模階層").ToString, row.Item("項目番号").ToString, row.Item("上限").ToString, row.Item("下限").ToString))
                        Next
                    End If

                End If

                '規模階層毎の集計
                For Each kaisouInfo As DAOKobetsuKekkahyo.Kibokaisou In kibokaisouList

                    '地域毎の集計
                    For Each chiikiKbn As String In _chiiki.Keys
                        '一つのレコード
                        count += dtSyutsuryokuRonri.Rows.Count
                    Next
                Next
            Next
        End If

        Return count
    End Function

    ''' <summary>
    ''' 進捗ダイアログを表示する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub openProgressDialog(db As DBAccess, syukeiInfo As DAOKobetsuKekkahyo.SyukeiInfo, dtCreateRonri As DataTable)
        _progressDialog = New ProgressDialog()
        _progressDialog.Maximum = getSyukeiCount(db, syukeiInfo, dtCreateRonri)
        _progressDialog.Show(Me)
    End Sub

    ''' <summary>
    ''' 進捗ダイアログを閉じる
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub closeProgressDialog()
        If Not _progressDialog Is Nothing Then
            _progressDialog.endDispose()
            _progressDialog = Nothing
        End If
    End Sub

End Class
