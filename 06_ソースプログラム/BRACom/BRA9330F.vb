'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.12.15 |Daiko               | 要件No4 バージョン区分追加
'//  REV_002   | 2022.12.21 |Daiko               | 要件No15 集計結果検討表（報告用）追加
'//  REV_003   | 2023.01.13 |Daiko               | 要件No4－⑦ 営農の集計結果検討表
'//  REV_004   | 2023.01.12 |Daiko               | 要件No.4 制度受取金・積立金等項目名称取得処理追加（集計結果検討表）
'//*************************************************************************************************
''' <summary>
''' 集計結果帳票出力画面
''' </summary>
''' <remarks></remarks>
Public Class BRA9330F

    ''' <summary>主キーペアクラス</summary>
    Public Class KeyPair
        Public PKeyThis As DAOSyukeiKekkahyo.PrimaryKey
        Public PKeyLast As DAOSyukeiKekkahyo.PrimaryKey
    End Class

    ''' <summary>帳票種別</summary>
    Public Enum 帳票種別
        集計結果表 = 1
        集計結果検討表
        ' REV_002↓
        集計結果検討表_報告用
        ' REV_002↑
    End Enum

    ''' <summary>
    ''' 帳票種別名称
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared 帳票種別名称 As New Dictionary(Of Integer, String) From {
          {帳票種別.集計結果表, "集計結果表"} _
        , {帳票種別.集計結果検討表, "集計結果検討表"} _
        , {帳票種別.集計結果検討表_報告用, "集計結果検討表（報告用）"}       ' REV_002↑
    }

    ''' <summary>帳票種別</summary>
    Private _printType As 帳票種別

    ''' <summary>主キー（本年）</summary>
    Private _pkeyThis As DAOSyukeiKekkahyo.PrimaryKey
    ''' <summary>項目キー（本年）</summary>
    Private _kkeyThis As DAOSyukeiKekkahyo.KomokuKey

    ''' <summary>主キー（前年）</summary>
    Private _pkeyLast As DAOSyukeiKekkahyo.PrimaryKey
    ''' <summary>項目キー（前年）</summary>
    Private _kkeyLast As DAOSyukeiKekkahyo.KomokuKey

    ''' <summary>営農経営体区分</summary>
    Private _einouKeieitai As String
    ''' <summary>調査区分</summary>
    Private _chosakubun As String

    ''' <summary>規模階層</summary>
    Private _kiboKaisou As String
    ''' <summary>地域</summary>
    Private _chiiki As String

    Public Sub New(chohyoShubetsu As 帳票種別, sKeyThis As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey), einouKeieitai As String)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        _printType = chohyoShubetsu

        txtChohyoMei.Text = 帳票種別名称(chohyoShubetsu)

        _pkeyThis = sKeyThis.Item1
        _kkeyThis = sKeyThis.Item2

        _einouKeieitai = einouKeieitai
        _chosakubun = ComUtil.GetChosakubun(_einouKeieitai)
    End Sub

    Public Sub New(chohyoShubetsu As 帳票種別, sKeyThis As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey), sKeyLast As ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.KomokuKey), einouKeieitai As String)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        _printType = chohyoShubetsu

        txtChohyoMei.Text = 帳票種別名称(chohyoShubetsu)

        _pkeyThis = sKeyThis.Item1
        _kkeyThis = sKeyThis.Item2

        _pkeyLast = sKeyLast.Item1
        _kkeyLast = sKeyLast.Item2

        _einouKeieitai = einouKeieitai
        _chosakubun = ComUtil.GetChosakubun(_einouKeieitai)
    End Sub

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA9330F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            txtChosanen.Text = _pkeyThis.chosaNen

            If CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計 Then
                txtEinouKeieitai.Text = ComConst.営農経営体区分.リスト(_einouKeieitai).名称
            Else
                lblEinouKeieitai.Visible = False
                txtEinouKeieitai.Visible = False
            End If

            '地域区分コンボボックス設定
            ComUtil.SetChiikiKbnComboBox(cboChiiki)

            txtSyukeiNoHonnen.Text = _pkeyThis.syukeiNo

            '集計名称取得
            Dim syukeiName As String
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                syukeiName = DAOSyukeiKekkahyo.GetSyukeiName(db, _chosakubun, _pkeyThis, _kkeyThis)
                txtSyukeiNameHonnen.Text = syukeiName
            End Using

            '集計結果表出力時は前年を非活性にする
            If _printType = 帳票種別.集計結果表 Then
                lblSyukeiNoZennen.Enabled = False
                lblSyukeiNameZennen.Enabled = False
                txtSyukeiNoZennen.Enabled = False
                txtSyukeiNameZennen.Enabled = False
            Else
                txtSyukeiNoZennen.Text = _pkeyLast.syukeiNo
                '集計名称取得
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    syukeiName = DAOSyukeiKekkahyo.GetSyukeiName(db, _chosakubun, _pkeyLast, _kkeyLast)
                    txtSyukeiNameZennen.Text = syukeiName
                End Using
            End If

            Dim dt As DataTable = Nothing

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                dt = DAOSyukeiKekkahyo.GetList(db, _chosakubun, _pkeyThis, _kkeyThis)
            End Using

            '規模階層コンボボックス設定
            Me.SetKiboKaisouComboBox(cboKiboKaisou, dt, True)

            '地域コンボボックス設定
            Me.SetChiikiComboBox(cboChiiki, dt, True)

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
            _kiboKaisou = If(cboKiboKaisou.SelectedValue Is Nothing, Nothing, cboKiboKaisou.SelectedValue.ToString)
            _chiiki = If(cboChiiki.SelectedValue Is Nothing, Nothing, cboChiiki.SelectedValue.ToString)

            '一覧表示
            Me.ShowList(_kiboKaisou, _chiiki)

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
        Dim dtSyukeiItemMst As DataTable
        Dim dtSyukeiKentoItemMst As DataTable
        Dim dtCreateRonri As DataTable

        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Dim pKeys As List(Of ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.PrimaryKey)) = Me.GetSyukeiKekkahyoPrimaryKey(_pkeyThis, _pkeyLast)

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(pKeys, msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            'フォルダパス取得
            Dim folderPath As String = ComUtil.GetFolderPath(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath))

            If folderPath.Equals(String.Empty) Then
                Exit Sub
            End If

            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Select Case _printType
                    Case 帳票種別.集計結果表
                        Dim pKeysThis As List(Of DAOSyukeiKekkahyo.PrimaryKey) = (From t In pKeys Select t.Item1).ToList()
                        ' REV_001↓
                        'Using ExcelOutput = New BRA9330R(folderPath, _chosakubun)
                        Using ExcelOutput = New BRA9330R(folderPath, _chosakubun, _pkeyThis.chosaNen)
                            ' REV_001↑
                            ret = ExcelOutput.Execute(pKeysThis, Me, 3)
                        End Using
                    Case 帳票種別.集計結果検討表
                        Dim seidoUketoriChosaNen As String = Nothing 'REV_004
                        Dim nousanFlg As Boolean = False 'REV_004

                        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                            'REV_004 START---------------
                            '農産物生産費かつR4体系の場合、制度受取金・積立金等項目名を設定する
                            If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 And ComConst.バージョン区分.結果表等項目2022 = ComUtil.getVersionKubunTaikei(_pkeyThis.chosaNen, CommonInfo.Chosakubun) Then
                                seidoUketoriChosaNen = _pkeyThis.chosaNen
                                nousanFlg = True
                            End If
                            'REV_004 END---------------
                            ' REV_001↓
                            '集計結果表項目マスタ取得(裏項番含める)
                            'dtSyukeiItemMst = DAOOther.GetSyukeiItemMaster(db, _chosakubun)
                            dtSyukeiItemMst = DAOOther.GetSyukeiItemMaster(db, _chosakubun, ComUtil.getVersionKubunTaikei(_pkeyThis.chosaNen, _chosakubun))
                            '集計結果検討表項目マスタ取得
                            'dtSyukeiKentoItemMst = DAOOther.GetSyukeiKekkaKentohyoItemMaster(db, _chosakubun)
                            dtSyukeiKentoItemMst = DAOOther.GetSyukeiKekkaKentohyoItemMaster(db, _chosakubun, ComUtil.getVersionKubunTaikei(_pkeyThis.chosaNen, _chosakubun))
                            '集計結果検討表作成論理取得
                            'dtCreateRonri = DAOOther.GetSyukeiKekkaKentohyoSakuseiRonri(db, _chosakubun)
                            dtCreateRonri = DAOOther.GetSyukeiKekkaKentohyoSakuseiRonri(db, _chosakubun, ComUtil.getVersionKubunTaikei(_pkeyThis.chosaNen, _chosakubun), seidoUketoriChosaNen) 'REV_004 
                            ' REV_001↑
                            '集計結果検討表作成論理のデータが１件も存在しない場合
                            If dtCreateRonri.Rows.Count = 0 Then
                                'エラーメッセージ
                                Message.ShowMsgBox(MessageID.MSG_E_048, MsgBoxStyle.OkOnly)
                                Exit Sub
                            End If
                            '集計結果検討表作成クラス

                            Dim kentohyo As CreateSyukeiKentohyo = New CreateSyukeiKentohyo(db,
                                                                                            _chosakubun,
                                                                                            dtSyukeiItemMst,
                                                                                            dtSyukeiKentoItemMst,
                                                                                            dtCreateRonri,
                                                                                            ComUtil.IsChikusan(),
                                                                                            nousanFlg, 'REV_004
                                                                                            ComConst.集計結果検討表作成論理.論理種別.集計結果検討表)　'REV_002

                            Dim pKeysPair As New List(Of KeyPair)
                            For Each pkey In pKeys
                                pKeysPair.Add(New KeyPair With {.PKeyThis = pkey.Item1, .PKeyLast = pkey.Item2})
                            Next
                            ' REV_001↓
                            'Using ExcelOutput = New BRA9340R(folderPath, _chosakubun, kentohyo)
                            Using ExcelOutput = New BRA9340R(folderPath, _chosakubun, kentohyo, _pkeyThis.chosaNen)
                                ' REV_001↑
                                ret = ExcelOutput.Execute(pKeysPair, Me, 3)
                            End Using
                        End Using
                        ' REV_002↓
#If BRAN Then
                    Case 帳票種別.集計結果検討表_報告用
                        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                            '集計結果表項目マスタ取得(裏項番含める)
                            dtSyukeiItemMst = DAOOther.GetSyukeiItemMaster(db, _chosakubun, ComUtil.getVersionKubunTaikei(_pkeyThis.chosaNen, _chosakubun))
                            '集計結果検討表項目マスタ取得
                            dtSyukeiKentoItemMst = DAOOther.GetSyukeiKekkaKentohyoHoukokuyoItemMaster(db, _chosakubun, ComUtil.getVersionKubunTaikei(_pkeyThis.chosaNen, _chosakubun))
                            '集計結果検討表作成論理取得
                            dtCreateRonri = DAOOther.GetSyukeiKekkaKentohyoHoukokuyoSakuseiRonri(db, _chosakubun, ComUtil.getVersionKubunTaikei(_pkeyThis.chosaNen, _chosakubun))
                            '集計結果検討表作成論理のデータが１件も存在しない場合
                            If dtCreateRonri.Rows.Count = 0 Then
                                'エラーメッセージ
                                Message.ShowMsgBox(MessageID.MSG_E_139, MsgBoxStyle.OkOnly)
                                Exit Sub
                            End If
                            '集計結果検討表作成クラス

                            Dim kentohyo As CreateSyukeiKentohyo = New CreateSyukeiKentohyo(db,
                                                                                            _chosakubun,
                                                                                            dtSyukeiItemMst,
                                                                                            dtSyukeiKentoItemMst,
                                                                                            dtCreateRonri,
                                                                                            ComUtil.IsChikusan(),
                                                                                            False,　'REV_004
                                                                                            ComConst.集計結果検討表作成論理.論理種別.集計結果検討表_報告用)　'REV_002

                            Dim pKeysPair As New List(Of KeyPair)
                            For Each pkey In pKeys
                                pKeysPair.Add(New KeyPair With {.PKeyThis = pkey.Item1, .PKeyLast = pkey.Item2})
                            Next
                            Using ExcelOutput = New BRA9350R(folderPath, _chosakubun, kentohyo, _pkeyThis.chosaNen)
                                ret = ExcelOutput.Execute(pKeysPair, Me, 3)
                            End Using
                        End Using
#End If
                        ' REV_002↑
                    Case Else
                        Throw New Exception("帳票種別選択エラー")
                End Select

                If ret = ExcelOutputBaseClass.enmOutputReturn.OK Then
                    '完了メッセージ
                    Message.ShowMsgBox(MessageID.MSG_I_002, MsgBoxStyle.OkOnly)
                End If
            Catch ex As ExcelOutputBaseClass.SaveAsException
                'エラーメッセージ
                Message.ShowMsgBox(MessageID.MSG_E_006, MsgBoxStyle.OkOnly)
            Catch ex As CreateSyukeiKentohyoException
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Message.ShowMsgBox(MessageID.MSG_E_054, {ex.SyukeiNo, ex.ItemNo}, MsgBoxStyle.OkOnly)
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

    Private Sub btnAllSelect_Click(sender As Object, e As EventArgs) Handles btnAllSelect.Click
        Try
            ComUtil.SetDataGridViewAllCheck(dgvList, True)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

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
    ''' 規模階層コンボボックス設定
    ''' </summary>
    ''' <param name="cbo"></param>
    ''' <param name="dt"></param>
    ''' <param name="none"></param>
    ''' <remarks></remarks>
    Private Sub SetKiboKaisouComboBox(cbo As ComboBox, dt As DataTable, Optional none As Boolean = False)
        Dim lst As New ArrayList()

        For Each row As DataRow In dt.Rows
            If CheckListRow(row("生産費平均値種類").ToString) Then Continue For
            row.Delete()
        Next
        dt.AcceptChanges()

        Dim query = (From dr In dt Select dr("規模階層")).Distinct().ToList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each obj In query
            lst.Add(New DictionaryEntry(obj, obj))
        Next

        cbo.DisplayMember = "Value"
        cbo.ValueMember = "Key"
        cbo.DataSource = lst
    End Sub

    ''' <summary>
    ''' 地域コンボボックス設定
    ''' </summary>
    ''' <param name="cbo"></param>
    ''' <param name="dt"></param>
    ''' <param name="none"></param>
    ''' <remarks></remarks>
    Private Sub SetChiikiComboBox(cbo As ComboBox, dt As DataTable, Optional none As Boolean = False)
        Dim lst As New ArrayList()

        For Each row As DataRow In dt.Rows
            If CheckListRow(row("生産費平均値種類").ToString) Then Continue For
            row.Delete()
        Next
        dt.AcceptChanges()

        Dim query = (From dr In dt Select dr("地域コード")).Distinct().ToList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each obj In query
            lst.Add(New DictionaryEntry(obj, ComConst.地域.リスト(obj.ToString).名称))
        Next

        cbo.DisplayMember = "Value"
        cbo.ValueMember = "Key"
        cbo.DataSource = lst
    End Sub

    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <param name="kiboKaisou"></param>
    ''' <param name="chiiki"></param>
    ''' <remarks></remarks>
    Public Sub ShowList(kiboKaisou As String, chiiki As String)
        Dim dt As DataTable = Nothing

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            If _pkeyLast Is Nothing Then
                dt = DAOSyukeiKekkahyo.GetList(db, _chosakubun, _pkeyThis, _kkeyThis, kiboKaisou, chiiki)
            Else
                dt = DAOSyukeiKekkahyo.GetList(db, _chosakubun, _pkeyThis, _kkeyThis, _pkeyLast, _kkeyLast, kiboKaisou, chiiki)
            End If
        End Using

        Dim query = dt.Select(Nothing, "生産費平均値種類, 規模階層, 地域コード, 田畑区分, ビール麦販売区分, てんさい栽培区分")

        dgvList.Rows.Clear()

        For Each row As DataRow In query
            If Not CheckListRow(row("生産費平均値種類").ToString) Then Continue For

            dgvList.Rows.Add()
            Dim i As Integer = dgvList.Rows.Count - 1
            dgvList.Rows(i).Cells(1).Value = row("生産費平均値種類").ToString
            dgvList.Rows(i).Cells(2).Value = row("規模階層").ToString
            dgvList.Rows(i).Cells(3).Value = ComConst.地域.リスト(row("地域コード").ToString).名称
            dgvList.Rows(i).Cells(4).Value = row("本年集計戸数").ToString
            dgvList.Rows(i).Cells(6).Value = row("本年連番").ToString

            If Not _pkeyLast Is Nothing Then
                dgvList.Rows(i).Cells(5).Value = If(IsDBNull(row("前年集計戸数")), Nothing, row("前年集計戸数").ToString)
                dgvList.Rows(i).Cells(7).Value = If(IsDBNull(row("前年連番")), Nothing, row("前年連番").ToString)
            End If

            dgvList.Rows(i).Cells(8).Value = If(ComConst.田畑区分.リスト.ContainsKey(row("田畑区分").ToString), ComConst.田畑区分.リスト(row("田畑区分").ToString), Nothing)
            dgvList.Rows(i).Cells(9).Value = If(ComConst.ビール麦販売区分.リスト.ContainsKey(row("ビール麦販売区分").ToString), ComConst.ビール麦販売区分.リスト(row("ビール麦販売区分").ToString), Nothing)
            dgvList.Rows(i).Cells(10).Value = If(ComConst.てんさい栽培区分.リスト.ContainsKey(row("てんさい栽培区分").ToString), ComConst.てんさい栽培区分.リスト(row("てんさい栽培区分").ToString), Nothing)
        Next
    End Sub

    ''' <summary>
    ''' 集計結果表主キー取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSyukeiKekkahyoPrimaryKey(pkeyThis As DAOSyukeiKekkahyo.PrimaryKey, pkeyLast As DAOSyukeiKekkahyo.PrimaryKey) As List(Of ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.PrimaryKey))
        Dim ret As New List(Of ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.PrimaryKey))

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                Dim pkey1 As DAOSyukeiKekkahyo.PrimaryKey = New DAOSyukeiKekkahyo.PrimaryKey(pkeyThis.chosaNen, pkeyThis.syukeiNo, dgvList.Rows(i).Cells(6).Value.ToString)
                Dim pkey2 As DAOSyukeiKekkahyo.PrimaryKey = New DAOSyukeiKekkahyo.PrimaryKey(If(pkeyLast Is Nothing, Nothing, pkeyLast.chosaNen), If(pkeyLast Is Nothing, Nothing, pkeyLast.syukeiNo), If(dgvList.Rows(i).Cells(7).Value Is Nothing, Nothing, dgvList.Rows(i).Cells(7).Value.ToString))
                ret.Add(ValueTuple.Create(pkey1, pkey2))
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="pKey"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(pKeys As List(Of ValueTuple(Of DAOSyukeiKekkahyo.PrimaryKey, DAOSyukeiKekkahyo.PrimaryKey)), ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'センサス番号選択チェック
        If pKeys.Count = 0 Then
            msgId = MessageID.MSG_E_045
            Return ret
        End If

        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' 一覧行チェック
    ''' </summary>
    ''' <param name="seisanhiHeikin"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckListRow(seisanhiHeikin As String) As Boolean
        Dim ret As Boolean = False

        Select Case _printType
            Case 帳票種別.集計結果表
                If CommonInfo.Koutei.Equals(CommonInfo.KouteiKubun.Code.Honsyo) Then
                    ret = True
                ElseIf CommonInfo.Koutei.Equals(CommonInfo.KouteiKubun.Code.Kyoku) And seisanhiHeikin.Equals("0") Then
                    ret = True
                End If
            Case 帳票種別.集計結果検討表
                If CommonInfo.Kubun2.Equals(ComConst.区分２.農産物生産費) And {"1", "2", "3", "4"}.Contains(seisanhiHeikin) Then
                    ret = True
                ElseIf CommonInfo.Kubun2.Equals(ComConst.区分２.畜産物生産費) And seisanhiHeikin.Equals("0") Then
                    ret = True
                    'REV_003↓
                ElseIf CommonInfo.Kubun2.Equals(ComConst.区分２.営農類型別経営統計) Then
                    ret = True
                    'REV_003↑
                End If
            ' REV_002↓
            Case 帳票種別.集計結果検討表_報告用
                If CommonInfo.Kubun2.Equals(ComConst.区分２.農産物生産費) And seisanhiHeikin.Equals("0") Then
                    ret = True
                ElseIf CommonInfo.Kubun2.Equals(ComConst.区分２.畜産物生産費) And seisanhiHeikin.Equals("0") Then
                    ret = True
                End If
                ' REV_002↑
        End Select

        Return ret
    End Function
End Class
