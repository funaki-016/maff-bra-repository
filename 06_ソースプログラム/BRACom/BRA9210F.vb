Imports System.Text
Imports System.IO

''' <summary>
''' 帳票出力画面
''' </summary>
''' <remarks></remarks>

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.11.16 |TSP)                | フェーズ3 要件No7,No10 対応
'//  REV_002   | 2022.10.11 |Daiko               | 要件No1 バージョン区分追加
'//  REV_003   | 2022.10.13 |大興電子通信        | 要件No.10
'//  REV_004   | 2023.11.27 |大興電子通信        | 要件No.19 営農個人歯抜け解消対応、営農改行除去対応
'//            |            |                    |
'//*************************************************************************************************
Public Class BRA9210F

    ''' <summary>帳票種別</summary>
    Public Enum 帳票種別
        電子調査票 = 1
        個別結果表
        個別結果検討表
    End Enum

    ''' <summary>
    ''' 帳票種別名称
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared 帳票種別名称 As New Dictionary(Of Integer, String) From {
          {帳票種別.電子調査票, "電子調査票"} _
        , {帳票種別.個別結果表, "個別結果表"} _
        , {帳票種別.個別結果検討表, "個別結果検討表"}
    }

    ''' <summary>帳票種別</summary>
    Private _printType As 帳票種別

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
    ''' <summary>欠測値補完</summary>
    Private _kessokuchiHokan As String
    ''' <summary>貸借対照表</summary>
    Private _taishakuTaishohyo As String

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA9210F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '用紙サイズの設定はいったん見えなくする
            GroupBoxYoshi.Visible = False

            '調査年コンボボックス設定
            Select Case _printType
                Case 帳票種別.電子調査票
                    Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                        ComUtil.Chosahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
                    End Using
                Case 帳票種別.個別結果表, 帳票種別.個別結果検討表
                    Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                        ComUtil.KobetsuKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
                    End Using
                Case Else
                    Throw New Exception("帳票種別選択エラー")
            End Select

            '局コンボボックス設定
            ComUtil.SetKyokuComboBox(cboKyoku)

            '営農類型コンボボックス設定
            ComUtil.SetEinouRuikeiComboBox(lblEinouRuikei, cboEinouRuikei)

            'DataGridView設定
            ComUtil.ConfigDgv(Me.dgvList)

            If _printType = 帳票種別.個別結果表 Then
                '欠測値補完コンボボックス設定
                ComUtil.SetKessokuchiHokanComboBox(lblKessokuchiHokan, cboKessokuchiHokan)

                '貸借対照表区分コンボボックス設定
                ComUtil.SetTaishakuTaishohyoComboBox(lblTaishakuTaishohyo, cboTaishakuTaishohyo)
            Else
                '欠測値補完コンボボックス非活性
                lblKessokuchiHokan.Visible = False
                cboKessokuchiHokan.Visible = False

                '貸借対照表区分コンボボックス非活性
                lblTaishakuTaishohyo.Visible = False
                cboTaishakuTaishohyo.Visible = False
            End If

            '--- REV.001 ADD START
            If _printType = 帳票種別.電子調査票 Then
                '電子調査票だけは用紙サイズの設定をできるようにする
                GroupBoxYoshi.Visible = True
                ' REV_003↓（調査年（産）選択変更イベントへ移動）
                '    If CommonInfo.Koutei <> CommonInfo.KouteiKubun.Code.Honsyo Then
                '        '本省工程以外の場合、調査票CSV出力ボタンは表示しない。
                '        btnCSVoutput.Visible = False
                '        btnOutPut.Text = "出力"
                '        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 OrElse CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                '            chkShokuin.Visible = False
                '        Else
                '            '職員記入欄を折りたたむチェックボックスの位置を調整する
                '            chkShokuin.Visible = True
                '            chkShokuin.Location = New Point(461, 551)
                '        End If
                '    Else
                '        btnCSVoutput.Visible = True
                '        btnOutPut.Text = "Excel出力"
                '        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 OrElse CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                '            chkShokuin.Visible = False
                '        Else
                '            '職員記入欄を折りたたむチェックボックスの位置を調整する
                '            chkShokuin.Visible = True
                '            chkShokuin.Location = New Point(365, 551)
                '        End If
                '    End If
                ' REV_003↑
            Else
                '電子調査票以外の場合、調査票CSV出力ボタンは表示しない
                btnCSVoutput.Visible = False
                btnOutPut.Text = "出力"
                '職員記入欄を折りたたむチェックボックスは表示しない
                chkShokuin.Visible = False
            End If
            '--- REV.001 ADD END

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
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="chohyoShubetsu"></param>
    ''' <remarks></remarks>
    Public Sub New(chohyoShubetsu As 帳票種別)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        _printType = chohyoShubetsu

        txtChohyoMei.Text = 帳票種別名称(chohyoShubetsu)
    End Sub

    ''' <summary>
    ''' 出力ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click
        Dim dtKobetsuItemMst As DataTable
        Dim dtKobetsuKentoItemMst As DataTable
        Dim dtCreateRonri As DataTable

        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(msgId) Then
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
                    Case 帳票種別.電子調査票
                        Dim pKeys As List(Of DAOChosahyo.PrimaryKey) = Me.GetChosahyoPrimaryKey(_chosaNen)
                        '--- REV.001 MOD START
                        Dim pShokuin As Boolean = chkShokuin.Checked

                        Dim yoshiSize As String
                        If rbYoshiA3.Checked Then
                            yoshiSize = rbYoshiA3.Text
                        ElseIf rbYoshiA4.Checked Then
                            yoshiSize = rbYoshiA4.Text
                        Else
                            yoshiSize = ""
                        End If

                        Using ExcelOutput = New BRA9210R(folderPath, pShokuin, _chosaNen, yoshiSize)
                            'Using ExcelOutput = New BRA9210R(folderPath)
                            '--- REV.001 MOD END
                            ret = ExcelOutput.Execute(pKeys, Me, 3)
                        End Using
                    Case 帳票種別.個別結果表
                        Dim pKeys As List(Of DAOKobetsuKekkahyo.PrimaryKey) = Me.GetKobetsuKekkahyoPrimaryKey(_chosaNen)
                        ' REV_002↓
                        'Using ExcelOutput = New BRA9220R(folderPath, _kessokuchiHokan)
                        Using ExcelOutput = New BRA9220R(folderPath, _kessokuchiHokan, _chosaNen)
                            ' REV_002↑
                            ret = ExcelOutput.Execute(pKeys, Me, 3)
                        End Using
                    Case 帳票種別.個別結果検討表
                        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                            '個別結果表項目マスタ取得(裏項番含める)
                            ' REV_002↓
                            'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, True)
                            dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun), True)
                            ' REV_002↑
                            '個別結果検討表項目マスタ取得
                            ' REV_002↓
                            'dtKobetsuKentoItemMst = DAOOther.GetKobetsuKekkaKentohyoItemMaster(db, CommonInfo.Chosakubun)
                            dtKobetsuKentoItemMst = DAOOther.GetKobetsuKekkaKentohyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
                            ' REV_002↑
                            '個別結果検討表作成論理取得
                            ' REV_002↓
                            'dtCreateRonri = DAOOther.GetKobetsuKekkaKentohyoSakuseiRonri(db)
                            dtCreateRonri = DAOOther.GetKobetsuKekkaKentohyoSakuseiRonri(db, ComUtil.getVersionKubunTaikei(_chosaNen, CommonInfo.Chosakubun))
                            ' REV_002↑
                            '個別結果検討表作成論理のデータが１件も存在しない場合
                            If dtCreateRonri.Rows.Count = 0 Then
                                'エラーメッセージ
                                Message.ShowMsgBox(MessageID.MSG_E_018, MsgBoxStyle.OkOnly)
                                Exit Sub
                            End If
                            '個別結果表・個別結果検討表作成クラス
                            Dim kobetsu As CreateKobetsu = New CreateKobetsu(db,
                                                                             CommonInfo.Chosakubun,
                                                                             _chosaNen,
                                                                             CreateKobetsu.enmCreateType.個別結果検討表作成,
                                                                             Nothing,
                                                                             dtKobetsuItemMst,
                                                                             dtKobetsuKentoItemMst,
                                                                             dtCreateRonri,
                                                                             Nothing,
                                                                             Nothing)

                            Dim pKeys As List(Of DAOKobetsuKekkahyo.PrimaryKey) = Me.GetKobetsuKekkahyoPrimaryKey(_chosaNen)
                            ' REV_002↓
                            'Using ExcelOutput = New BRA9230R(folderPath, kobetsu)
                            Using ExcelOutput = New BRA9230R(folderPath, kobetsu, _chosaNen)
                                ' REV_002↑
                                ret = ExcelOutput.Execute(pKeys, Me, 3)
                            End Using
                        End Using

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
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 農政局コンボボックス選択時
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
            _kessokuchiHokan = ComUtil.GetkessokuchiHokana(cboKessokuchiHokan)
            _taishakuTaishohyo = If(cboTaishakuTaishohyo.SelectedValue Is Nothing, Nothing, cboTaishakuTaishohyo.SelectedValue.ToString)

            '一覧表示
            Me.ShowList(_chosaNen, _kyoku, _jimusho, _kyoten, _einouRuikei, _taishakuTaishohyo, _kessokuchiHokan)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 欠測値補完コンボボックス選択時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboKessokuchiHokan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKessokuchiHokan.SelectedIndexChanged
        Try
            dgvList.Rows.Clear()

            '欠測値補完取得
            Dim kessokuchiHokan As String = ComUtil.GetkessokuchiHokana(cboKessokuchiHokan)

            '調査年コンボボックス設定
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ComUtil.KobetsuKekkahyo.SetChosaNenComboBox(cboChosaNen, db, CommonInfo.Koutei, CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center, kessokuchiHokan)
            End Using
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 一覧表示
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="einouRuikei"></param>
    ''' <param name="taishakuTaishohyo"></param>
    ''' <param name="kessokuchiHokan"></param>
    ''' <remarks></remarks>
    Private Sub ShowList(chosaNen As String, kyoku As String, jimusho As String, kyoten As String, einouRuikei As String, taishakuTaishohyo As String, kessokuchiHokan As String)
        Dim dt As DataTable

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Select Case _printType
                Case 帳票種別.電子調査票
                    dt = DAOChosahyo.GetChosahyoList(db, chosaNen, kyoku, jimusho, kyoten, einouRuikei, "", "")
                Case 帳票種別.個別結果表, 帳票種別.個別結果検討表
                    dt = DAOKobetsuKekkahyo.GetList(db, chosaNen, kyoku, jimusho, kyoten, einouRuikei, taishakuTaishohyo, kessokuchiHokan)
                Case Else
                    Throw New Exception("帳票種別選択エラー")
            End Select
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
        Next
    End Sub

    ''' <summary>
    ''' 調査票主キー取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChosahyoPrimaryKey(chosaNen As String) As List(Of DAOChosahyo.PrimaryKey)
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
    ''' 個別結果表主キー取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
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
    ''' エラーチェック
    ''' </summary>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        'センサス番号選択チェック
        If dgvList.Rows.Count = 0 Then
            msgId = MessageID.MSG_E_003
            Return ret
        End If

        For i As Integer = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                Exit For
            Else
                If i = dgvList.Rows.Count - 1 Then
                    msgId = MessageID.MSG_E_003
                    Return ret
                End If
            End If
        Next

        ret = True

        Return ret
    End Function

    '--- REV.001 ADD START
    ''' <summary>
    ''' CSV出力ボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnCSVoutput_Click(sender As Object, e As EventArgs) Handles btnCSVoutput.Click
        Try

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            'フォルダパス取得
            Dim folderPath As String = ComUtil.GetFolderPath(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath))

            If folderPath.Equals(String.Empty) Then
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_035, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                Dim dcs As Dictionary(Of String, DataTable) = Nothing
                Dim ColName As DataTable

                '調査票主キー取得
                Dim pKeys As List(Of DAOChosahyo.PrimaryKey) = Me.GetChosahyoPrimaryKey(_chosaNen)

                '調査票項目名マスタ取得
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    ColName = DAOChosahyo.getChosahyoColName(db)
                End Using

                For Each key In pKeys
                    Dim dc As Dictionary(Of String, DataTable)
                    Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                        '調査票テーブル取得
                        dc = DAOChosahyo.GetCsvChosahyoTable(db, key)
                    End Using

                    If dcs Is Nothing Then
                        dcs = dc
                    Else
                        For Each _dc As KeyValuePair(Of String, DataTable) In dc
                            Dim tableName As String = _dc.Key
                            Dim dt As DataTable = _dc.Value
                            For i As Integer = 0 To dt.Rows.Count - 1
                                dcs(tableName).ImportRow(dt(i))
                            Next
                        Next
                    End If
                Next

                'ファイル名取得
                Dim fileNamePattern As String = Me.GetFileNamePattern(_chosaNen)

                'CSV出力
                Dim ret As ComConst.CSVファイル.enmOutputReturn = Me.PutCSV(_chosaNen, folderPath, fileNamePattern, dcs, ColName)

                Select Case ret
                    Case ComConst.CSVファイル.enmOutputReturn.OK
                        '完了メッセージ
                        Message.ShowMsgBox(MessageID.MSG_I_002, MsgBoxStyle.OkOnly)
                    Case ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS
                        'エラーメッセージ
                        Message.ShowMsgBox(MessageID.MSG_E_058, MsgBoxStyle.OkOnly)
                End Select
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
    ''' CSV出力
    ''' </summary>
    ''' <param name="outDir"></param>
    ''' <param name="fileName"></param>
    ''' <param name="dc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PutCSV(year As String, outDir As String, fileName As String, dc As Dictionary(Of String, DataTable), col As DataTable) As ComConst.CSVファイル.enmOutputReturn
        Dim ret As ComConst.CSVファイル.enmOutputReturn = ComConst.CSVファイル.enmOutputReturn.CANCEL

        Dim sjisEnc As Encoding = Encoding.GetEncoding(ComConst.CSVファイル.CODEPAGE_SHIFT_JIS)

        Dim filePathTemp As String = System.IO.Path.Combine(outDir, fileName)

        If Not System.IO.Directory.Exists(outDir) Then
            Directory.CreateDirectory(outDir)
        End If

        For Each kv As KeyValuePair(Of String, DataTable) In dc
            Dim tableName As String = kv.Key
            Dim dt As DataTable = kv.Value

            Try
                Using sw As New System.IO.StreamWriter(String.Format(filePathTemp, tableName & ComUtil.KobetsuKekkahyo.GetSyukeiTableAddName(CommonInfo.Chosakubun, _kessokuchiHokan)), False, sjisEnc)
                    '1行目はテーブル名と調査区分を出力
                    Dim header(1) As Object
                    header(0) = kv.Key & "_" & year
                    header(1) = CommonInfo.Chosakubun
                    sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, header) & ComConst.CSVファイル.END_ADDITION)

                    Dim colArr(dt.Columns.Count - 1) As Object
                    'カラム名の配列作成
                    dt.Columns.CopyTo(colArr, 0)

                    '2行目以降は分岐
                    If tableName.Contains("＿可変") Then

                        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 OrElse CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                            '調査区分が営農個人または営農法人の時のみ処理
                            '各調査区分の可変項目数をチェック
                            Dim dtChoItemMstKahen As DataTable
                            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                                '調査票項目マスタ取得（可変のみ）
                                dtChoItemMstKahen = DAOOther.GetChosahyoItemMasterKahen(db, CommonInfo.Chosakubun, year)
                            End Using

                            'ヘッダー行作成
                            '調査年、センサス番号、農政局、都道府県、実査設置拠点
                            For i As Integer = 0 To 4
                                sw.Write(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, colArr(i)) & ComConst.CSVファイル.END_ADDITION & ",")
                            Next
                            For Each row As DataRow In dtChoItemMstKahen.Rows
                                For i As Integer = 1 To CInt(row("可変最大数"))
                                    '項目番号、明細番号、値
                                    For j = 5 To 7
                                        sw.Write(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, colArr(j)) & ComConst.CSVファイル.END_ADDITION & ",")
                                    Next
                                Next
                            Next
                            '更新日付、更新者ID
                            sw.Write(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, colArr(8)) & ComConst.CSVファイル.END_ADDITION & ",")
                            sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, colArr(9)) & ComConst.CSVファイル.END_ADDITION)

                            '明細作成
                            For Each kv2 As KeyValuePair(Of String, DataTable) In dc
                                If kv2.Key.Contains("＿可変") Then
                                    Continue For
                                End If

                                '可変じゃないテーブルの行（センサス番号）でループ
                                For Each row As DataRow In kv2.Value.Rows
                                    Dim dr As DataRow()
                                    '調査年、センサス番号、農政局、都道府県、実査設置拠点
                                    For i As Integer = 0 To 4
                                        sw.Write(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, row(i)) & ComConst.CSVファイル.END_ADDITION & ",")
                                    Next

                                    '調査票項目マスタ毎
                                    For Each koumokuRow As DataRow In dtChoItemMstKahen.Rows
                                        '可変最大数
                                        For i As Integer = 1 To CInt(koumokuRow("可変最大数"))
                                            '項目番号、明細番号
                                            sw.Write(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, koumokuRow(2)) & ComConst.CSVファイル.END_ADDITION & ",")
                                            sw.Write(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, i) & ComConst.CSVファイル.END_ADDITION & ",")

                                            '値
                                            dr = dt.Select("センサス番号 = '" & row(1).ToString & "' and 項目番号 = '" & koumokuRow(2).ToString & "' and 明細番号 = '" & i.ToString & "'")
                                            If dr.Length = 0 Then
                                                sw.Write(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, "") & ComConst.CSVファイル.END_ADDITION & ",")
                                            Else
                                                sw.Write(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, dr(0)(7).ToString.Replace(Chr(10), "")) & ComConst.CSVファイル.END_ADDITION & ",")
                                            End If

                                        Next
                                    Next

                                    '更新日付、更新者ID
                                    dr = dt.Select("センサス番号 = '" & row(1).ToString & "'")
                                    If dr.Length = 0 Then
                                        sw.Write(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, "") & ComConst.CSVファイル.END_ADDITION & ",")
                                        sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, "") & ComConst.CSVファイル.END_ADDITION)
                                    Else
                                        sw.Write(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, dt(0)(8).ToString) & ComConst.CSVファイル.END_ADDITION & ",")
                                        sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, dt(0)(9).ToString) & ComConst.CSVファイル.END_ADDITION)
                                    End If
                                Next
                            Next

                        Else
                            '調査区分が営農個人／営農法人以外
                            '可変テーブルの場合、カラム名をそのまま出力
                            sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, colArr) & ComConst.CSVファイル.END_ADDITION)
                            For Each row As DataRow In dt.Rows
                                '3行目以降は[項番]列の変換を行うそのまま出力
                                Dim arr As Object() = row.ItemArray().ToArray()
                                Dim d1() As DataRow
                                d1 = col.Select("項番 = '" & arr(5).ToString & "'")

                                If d1.Length <> 0 Then
                                    arr(5) = arr(5).ToString & "(" & d1(0).Item("項目名").ToString & ")"
                                End If

                                sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, arr) & ComConst.CSVファイル.END_ADDITION)
                            Next
                        End If

                    Else
                        '可変テーブル以外の場合、項目名マスタから取得した項目名を出力
                        Dim colNameArr(dt.Columns.Count - 1) As Object

                        For i As Integer = 0 To colArr.Length - 1
                            Dim d1() As DataRow
                            d1 = col.Select("項番 = '" & colArr(i).ToString & "'")

                            If d1.Length <> 0 Then
                                colNameArr(i) = colArr(i).ToString & "(" & d1(0).Item("項目名").ToString & ")"
                            Else
                                colNameArr(i) = colArr(i).ToString
                            End If
                        Next

                        sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, colNameArr) & ComConst.CSVファイル.END_ADDITION)

                        For Each row As DataRow In dt.Rows
                            '3行目以降はそのまま出力
                            Dim arr As Object() = row.ItemArray().ToArray()
                            '営農個人営農法人　改行除去対応
                            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 OrElse CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                                For i = 0 To arr.Length() - 1
                                    arr(i) = arr.GetValue(i).ToString.Replace(Chr(10), "")
                                Next
                            End If
                            sw.WriteLine(ComConst.CSVファイル.START_ADDITION & String.Join(ComConst.CSVファイル.START_ADDITION & ComConst.CSVファイル.CSV_DELIMITER & ComConst.CSVファイル.END_ADDITION, arr) & ComConst.CSVファイル.END_ADDITION)

                        Next

                    End If

                End Using
            Catch ex As Exception
                ret = ComConst.CSVファイル.enmOutputReturn.ERR_SAVEAS
                Return ret
            End Try
        Next

        ret = ComConst.CSVファイル.enmOutputReturn.OK

        Return ret
    End Function

    ''' <summary>
    ''' ファイル名パターン取得
    ''' </summary>
    ''' <param name="year"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetFileNamePattern(year As String) As String
        Return "{0}_" & year & ".csv"
    End Function
    '--- REV.001 ADD END

    ''' <summary>
    ''' 調査年（産）選択変更(REV_003)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cboChosaNen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboChosaNen.SelectedIndexChanged
        If _printType <> 帳票種別.電子調査票 Then
            Return
        End If
        Dim year = CInt(If(cboChosaNen.SelectedValue, 0))
        If year <= 2021 AndAlso CommonInfo.Koutei <> CommonInfo.KouteiKubun.Code.Honsyo Then
            '～2021かつ本省工程以外の場合、調査票CSV出力ボタンは表示しない。
            btnCSVoutput.Visible = False
            btnOutPut.Text = "出力"

            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 OrElse CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                chkShokuin.Visible = False
            Else
                '職員記入欄を折りたたむチェックボックスの位置を調整する
                chkShokuin.Visible = True
                chkShokuin.Location = New Point(461, 551)
            End If
        Else
            btnCSVoutput.Visible = True
            btnOutPut.Text = "Excel出力"
            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 OrElse CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                chkShokuin.Visible = False
            Else
                '職員記入欄を折りたたむチェックボックスの位置を調整する
                chkShokuin.Visible = True
                chkShokuin.Location = New Point(365, 551)
            End If
        End If
    End Sub

End Class
