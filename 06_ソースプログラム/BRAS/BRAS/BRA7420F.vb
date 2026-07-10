'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.12.16 |Daiko               | 要件No1 バージョン区分追加
'//  REV_002   | 2022.12.16 |Daiko               | 要件No4 バージョン区分追加
'//  REV_003   | 2023.12.15 |Daiko               | 変更要件No.5 営農は個別結果表項目をバージョン区分に関わらず取得するように修正
'//*************************************************************************************************
''' <summary>
''' 還元資料印刷内容指定画面
''' </summary>
''' <remarks></remarks>

'還元資料（還元資料印刷内容指定）クラス
Imports System.IO
Imports System.Reflection
Imports System.Data.SqlClient

Public Class BRA7420F
    '定数
    '===============================================
    Private _CensusNo As String
    Private _ChosaNen As String
    Private _Chosakubun As String

    'シート指定表示名称
    Private Const CHK_KOJIN1 As String = "営個還１"
    Private Const CHK_KOJIN2 As String = "営個還２"
    Private Const CHK_KOJIN3 As String = "営個還３"
    Private Const CHK_KOJIN4 As String = "営個還４"
    Private Const CHK_KOJIN5 As String = "営個還５"
    Private Const CHK_KOJIN6 As String = "営個還６"
    Private Const CHK_KOJIN7 As String = "営個還７"
    Private Const CHK_KOJIN8 As String = "営個還８"
    Private Const CHK_KOJIN9 As String = "営個還９"
    Private Const CHK_KOJIN10 As String = "営個還１０"

    Private Const CHK_HOUJIN1 As String = "営法還１"
    Private Const CHK_HOUJIN2 As String = "営法還２"
    Private Const CHK_HOUJIN3 As String = "営法還３"
    Private Const CHK_HOUJIN4 As String = "営法還４"
    Private Const CHK_HOUJIN5 As String = "営法還５"
    Private Const CHK_HOUJIN6 As String = "営法還６"
    Private Const CHK_HOUJIN7 As String = "営法還７"
    Private Const CHK_HOUJIN8 As String = "営法還８"
    Private Const CHK_HOUJIN9 As String = "営法還９"
    Private Const CHK_HOUJIN10 As String = "営法還１０"

    Private Const CHK_SEISAN1 As String = "還元１"
    Private Const CHK_SEISAN2 As String = "還元２"
    Private Const CHK_SEISAN3 As String = "還元３"
    Private Const CHK_SEISAN4 As String = "還元４"
    Private Const CHK_SEISAN5 As String = "還元５"
    Private Const CHK_SEISAN6 As String = "還元６"
    Private Const CHK_SEISAN7 As String = "還元７"
    Private Const CHK_SEISAN8 As String = "還元８"
    Private Const CHK_SEISAN9 As String = "還元９"
    Private Const CHK_SEISAN10 As String = "還元１０"
    Private Const CHK_SEISAN11 As String = "還元１１"

    ' 経年個別結果表のテキストボックスの配列
    Private mKeinenKobetuInfo As ArrayList = New ArrayList
    Private Const INX_YEAR = 0    '経年個別結果表のテキストボックスの配列インデックス:調査年
    Private Const INX_CENSUS = 1    '経年個別結果表のテキストボックスの配列インデックス:センサス番号
    Private Const INX_FIND = 2    '経年個別結果表のテキストボックスの配列インデックス:個別結果表有無
    Private Const TXT_FOUND = "有"
    Private Const TXT_NOT_FOUND = "無"
    Private Enum CellStyle        ' 経年個別結果表のセルタイプ
        FOUND
        NOT_FOUND
        NORMAL
    End Enum

    'テーブル名
    Private Const TBL_KANGENSETTING As String = "還元資料設定値"
    Private Const TBL_KANGENKOUMOKU As String = "還元資料項目名マスタ"
    Private Const TBL_KANGENCHOHYO As String = "還元資料固定項目定義マスタ"
    Private Const TBL_KANGENKANRI As String = "還元資料作成管理"
    Private Const TBL_TAIOU As String = "還元資料個別結果表対応マスタ"

    Private Const DEFAULT_CSNSUS As String = "0000000000000000"
    Private Const FILENAME_EINOU As String = "農業経営の概況_"
    Private Const FILENAME_SEISAN As String = "生産費の概要_"
    Private Const FILE_NAME_TEMP As String = "temp"

    Private Const EINOU_KOJIN As String = "営農個人"
    Private Const EINOU_HOUJIN As String = "営農法人"
    Private Const SEISANHI As String = "生産費"
    Private Const CENSUS As String = "K000003"

    Private Const SHUKEIKEKKA As String = "2"
    Dim startFlg As Boolean = True
    '===============================================

#Region "【処理詳細仕様 1】初期表示"

    Public Sub New(pCensusNo As String, pChosaNen As String)

        InitializeComponent()   '呼び出しの後で初期化

        _CensusNo = pCensusNo
        _ChosaNen = pChosaNen

        _Chosakubun = String.Empty

        If CommonInfo.Kubun1 = ComConst.区分１.農業経営統計調査 Then
            If CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計 Then
                If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                    _Chosakubun = EINOU_KOJIN
                ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                    _Chosakubun = EINOU_HOUJIN
                End If
            End If
        End If

        If _Chosakubun = String.Empty Then
            _Chosakubun = SEISANHI
        End If

        '経年個別結果表テキストボックスのリスト化
        mKeinenKobetuInfo.Add({txtChosaYear1, txtPartCenterBg1, txtIsFind1})
        mKeinenKobetuInfo.Add({txtChosaYear2, txtPartCenterBg2, txtIsFind2})
        mKeinenKobetuInfo.Add({txtChosaYear3, txtPartCenterBg3, txtIsFind3})
        mKeinenKobetuInfo.Add({txtChosaYear4, txtPartCenterBg4, txtIsFind4})
        mKeinenKobetuInfo.Add({txtChosaYear5, txtPartCenterBg5, txtIsFind5})

        startFlg = False
    End Sub

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA7420F_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'シート指定のチェックボックスの名称の編集
            SetCheckBoxControl(_Chosakubun)

            ' 経年個別結果表選択部分の初期設定
            SetKeinenKobetuSentaku()

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try

    End Sub
#End Region

#Region "【処理詳細仕様2】「全選択」ボタンクリック"
    ''' <summary>
    ''' 「全選択」ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAllSelect_Click(sender As System.Object, e As System.EventArgs) Handles btnAllSelect.Click

        If CheckBox1.Visible = True AndAlso CheckBox1.Enabled = True Then
            CheckBox1.Checked = True
        End If

        If CheckBox2.Visible = True AndAlso CheckBox2.Enabled = True Then
            CheckBox2.Checked = True
        End If

        If CheckBox3.Visible = True AndAlso CheckBox3.Enabled = True Then
            CheckBox3.Checked = True
        End If

        If CheckBox4.Visible = True AndAlso CheckBox4.Enabled = True Then
            CheckBox4.Checked = True
        End If

        If CheckBox5.Visible = True AndAlso CheckBox5.Enabled = True Then
            CheckBox5.Checked = True
        End If

        If CheckBox6.Visible = True AndAlso CheckBox6.Enabled = True Then
            CheckBox6.Checked = True
        End If

        If CheckBox7.Visible = True AndAlso CheckBox7.Enabled = True Then
            CheckBox7.Checked = True
        End If

        If CheckBox8.Visible = True AndAlso CheckBox8.Enabled = True Then
            CheckBox8.Checked = True
        End If

        If CheckBox9.Visible = True AndAlso CheckBox9.Enabled = True Then
            CheckBox9.Checked = True
        End If

        If CheckBox10.Visible = True AndAlso CheckBox10.Enabled = True Then
            CheckBox10.Checked = True
        End If

        If CheckBox11.Visible = True AndAlso CheckBox11.Enabled = True Then
            CheckBox11.Checked = True
        End If

    End Sub
#End Region

#Region "【処理詳細仕様3】「全解除」ボタンクリック"
    ''' <summary>
    ''' 「全解除」ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAllCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnAllCancel.Click

        If CheckBox1.Visible = True AndAlso CheckBox1.Enabled = True Then
            CheckBox1.Checked = False
        End If

        If CheckBox2.Visible = True AndAlso CheckBox2.Enabled = True Then
            CheckBox2.Checked = False
        End If

        If CheckBox3.Visible = True AndAlso CheckBox3.Enabled = True Then
            CheckBox3.Checked = False
        End If

        If CheckBox4.Visible = True AndAlso CheckBox4.Enabled = True Then
            CheckBox4.Checked = False
        End If

        If CheckBox5.Visible = True AndAlso CheckBox5.Enabled = True Then
            CheckBox5.Checked = False
        End If

        If CheckBox6.Visible = True AndAlso CheckBox6.Enabled = True Then
            CheckBox6.Checked = False
        End If

        If CheckBox7.Visible = True AndAlso CheckBox7.Enabled = True Then
            CheckBox7.Checked = False
        End If

        If CheckBox8.Visible = True AndAlso CheckBox8.Enabled = True Then
            CheckBox8.Checked = False
        End If

        If CheckBox9.Visible = True AndAlso CheckBox9.Enabled = True Then
            CheckBox9.Checked = False
        End If

        If CheckBox10.Visible = True AndAlso CheckBox10.Enabled = True Then
            CheckBox10.Checked = False
        End If

        If CheckBox11.Visible = True AndAlso CheckBox11.Enabled = True Then
            CheckBox11.Checked = False
        End If

    End Sub
#End Region

#Region "【処理詳細仕様4】「詳細設定」ボタンクリック"
    ''' <summary>
    ''' 「詳細設定」ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSetting_Click(sender As Object, e As EventArgs) Handles btnSetting.Click
        '個別結果表から営農区分を取得
        Dim einou As String
        Dim kibokaiso As Integer

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                einou = GetEinouKubunData(db)

            Else
                einou = "0"
            End If
            kibokaiso = GetKiboKaisoData(db)
        End Using

        '画面遷移
        Dim frm As New BRA7430F(_CensusNo, _ChosaNen, einou, kibokaiso)
        frm.Show(Me)
        Me.Hide()
    End Sub
#End Region

#Region "【処理詳細仕様5】「個別結果表確認」ボタンクリック"
    Private Sub btnCheck_Click(sender As System.Object, e As System.EventArgs) Handles btnCheck.Click
        Try
            Dim txtBoxs As System.Windows.Forms.TextBox()  ' テキストボックスの配列
            Dim year As String = ""
            Dim censusNo As String = ""
            Dim kyoku As String = CommonInfo.Kyoku
            Dim jimusho As String = CommonInfo.Jimusyo
            Dim kyoten As String = CommonInfo.Center

            '調査年の妥当性をチェックする。
            If chkTyosaYearTxtBox() = False Then
                Exit Sub
            End If

            For i As Integer = 0 To mKeinenKobetuInfo.Count - 1
                '一つ目は固定なので飛ばす
                If i = 0 Then
                    Continue For
                End If
                txtBoxs = CType(mKeinenKobetuInfo(i), TextBox())

                year = txtBoxs(INX_YEAR).Text
                censusNo = txtBoxs(INX_CENSUS).Text

                If year = String.Empty OrElse censusNo = String.Empty Then
                    txtBoxs(INX_YEAR).BackColor = Color.White
                    txtBoxs(INX_CENSUS).BackColor = Color.White

                    txtBoxs(INX_FIND).Text = String.Empty

                    Continue For
                End If

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                    '個別結果表存在チェック
                    If DAOKobetsuKekkahyo.CheckExistSenmon(db, New DAOKobetsuKekkahyo.PrimaryKey(year, censusNo),
                                                             New DAOKobetsuKekkahyo.KyotenKey(kyoku, jimusho, kyoten)) Then
                        txtBoxs(INX_YEAR).BackColor = Color.White
                        txtBoxs(INX_CENSUS).BackColor = Color.White
                        txtBoxs(INX_FIND).Text = TXT_FOUND
                    Else
                        '見つからなかった場合
                        txtBoxs(INX_YEAR).BackColor = Color.Red
                        txtBoxs(INX_CENSUS).BackColor = Color.Red

                        txtBoxs(INX_FIND).Text = TXT_NOT_FOUND
                    End If

                End Using
            Next
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region

#Region "【処理詳細仕様6】「印刷」ボタンクリック"
    Private Sub btnPrint_Click(ByVal eventSender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim keinenFlg As Boolean
        Dim txtBoxs() As System.Windows.Forms.TextBox
        Dim keys As New List(Of DAOKobetsuKekkahyo.PrimaryKey)
        Dim checkBox As New ArrayList
        Dim chosaNenKey As New ArrayList

        Dim dtSetting As New DataTable
        Dim dtKobetsu As New Dictionary(Of String, Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目))
        Dim dtShukei As New Dictionary(Of String, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目))
        Dim dtChohyoMaster As New DataTable
        Dim dtTaiouMaster As New DataTable
        Dim dtKangenKomokuMaster As New DataTable
        Dim fileName As String = String.Empty
        Dim filePath As String = String.Empty
        Dim dtSeidoUketoriItem As New DataTable ' REV_002

        Try
            '経年個別結果表欄のチェック
            For i As Integer = 0 To mKeinenKobetuInfo.Count - 1
                txtBoxs = CType(mKeinenKobetuInfo(i), TextBox())
                '経年個別結果表有無の当年以降に空白または「無」がある場合対象外
                If txtBoxs(INX_FIND).Text <> TXT_FOUND Then
                    keinenFlg = True
                    Continue For
                End If
                If Not String.IsNullOrEmpty(txtBoxs(INX_YEAR).Text) And Not String.IsNullOrEmpty(txtBoxs(INX_CENSUS).Text) Then
                    Dim pkey As DAOKobetsuKekkahyo.PrimaryKey = New DAOKobetsuKekkahyo.PrimaryKey(txtBoxs(INX_YEAR).Text, txtBoxs(INX_CENSUS).Text)
                    keys.Add(pkey)
                End If

            Next

            '経年個別結果表有無の当年以降に空白または「無」がある場合,警告を表示
            If keinenFlg Then
                If Message.ShowMsgBox(MessageID.MSG_Q_032, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If

            '処理確認
            If Message.ShowMsgBox(MessageID.MSG_Q_033, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If

            ' REV_002↓
            'Call getData(dtKobetsu, chosaNenKey, dtSetting, dtKangenKomokuMaster, dtChohyoMaster, dtTaiouMaster, dtShukei, keys)
            Call getData(dtKobetsu, chosaNenKey, dtSetting, dtKangenKomokuMaster, dtChohyoMaster, dtTaiouMaster, dtShukei, keys, dtSeidoUketoriItem)
            ' REV_002↑

            'チェックボックスの状態を配列に格納
            checkBox = Me.GetCheckBoxList

            '一時ファイルパス取得
            fileName = FILE_NAME_TEMP & ".xlsx"

            filePath = If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath) & "\" & fileName

            '一時ファイル存在ﾁｪｯｸ
            If File.Exists(filePath) Then
                '存在している場合、ファイルを削除する
                File.Delete(filePath)
            End If

            '還元資料設定値マスタ、個別結果表データ、集計結果表データ、還元資料帳票マスタを使用して、還元資料の作成を行う。
            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                ' REV_002↓
                'Using kangen = New CreateKangenEinou(checkBox, True, False, filePath, dtSetting, dtKobetsu, dtShukei, dtChohyoMaster, dtTaiouMaster, dtKangenKomokuMaster, chosaNenKey)
                Using kangen = New CreateKangenEinou(checkBox, True, False, filePath, dtSetting, dtKobetsu, dtShukei, dtChohyoMaster, dtTaiouMaster, dtKangenKomokuMaster, chosaNenKey, dtSeidoUketoriItem)
                    ' REV_002↑
                    kangen.Execute(MessageID.MSG_Q_004)
                End Using
            Else
                ' REV_002↓
                'Using kangen = New CreateKangenSeisan(checkBox, True, False, filePath, dtSetting, dtKobetsu, dtShukei, dtChohyoMaster, dtTaiouMaster, dtKangenKomokuMaster, chosaNenKey)
                Using kangen = New CreateKangenSeisan(checkBox, True, False, filePath, dtSetting, dtKobetsu, dtShukei, dtChohyoMaster, dtTaiouMaster, dtKangenKomokuMaster, chosaNenKey, dtSeidoUketoriItem)
                    ' REV_002↑
                    kangen.Execute(MessageID.MSG_Q_004)
                End Using
            End If


            '作成履歴を登録する。
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try
                    UpdateKangenRireki(db)
                Catch ex As Exception
                    InsertKangenRireki(db)
                End Try

            End Using


            Message.ShowMsgBox(MessageID.MSG_I_034, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try

    End Sub

#End Region


#Region "【処理詳細仕様7】「excel出力」ボタンクリック"
    Private Sub btnMake_Click(ByVal eventSender As System.Object, ByVal e As System.EventArgs) Handles btnMake.Click
        Dim keinenFlg As Boolean
        Dim txtBoxs() As System.Windows.Forms.TextBox
        Dim keys As New List(Of DAOKobetsuKekkahyo.PrimaryKey)
        Dim checkBox As New ArrayList
        Dim chosaNenKey As New ArrayList

        Dim dtSetting As New DataTable
        Dim dtKobetsu As New Dictionary(Of String, Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目))
        Dim dtShukei As New Dictionary(Of String, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目))
        Dim dtChohyoMaster As New DataTable
        Dim dtTaiouMaster As New DataTable
        Dim dtKangenKomokuMaster As New DataTable
        Dim fileName As String = String.Empty
        Dim filePath As String = String.Empty
        Dim dtSeidoUketoriItem As New DataTable ' REV_002

        Try
            '経年個別結果表欄のチェック
            For i As Integer = 0 To mKeinenKobetuInfo.Count - 1
                txtBoxs = CType(mKeinenKobetuInfo(i), TextBox())
                '経年個別結果表有無の当年以降に空白または「無」がある場合対象外
                If txtBoxs(INX_FIND).Text <> TXT_FOUND Then
                    keinenFlg = True
                    Continue For
                End If
                If Not String.IsNullOrEmpty(txtBoxs(INX_YEAR).Text) And Not String.IsNullOrEmpty(txtBoxs(INX_CENSUS).Text) Then
                    Dim pkey As DAOKobetsuKekkahyo.PrimaryKey = New DAOKobetsuKekkahyo.PrimaryKey(txtBoxs(INX_YEAR).Text, txtBoxs(INX_CENSUS).Text)
                    keys.Add(pkey)
                End If

            Next

            '経年個別結果表有無の当年以降に空白または「無」がある場合,警告を表示
            If keinenFlg Then

                If Message.ShowMsgBox(MessageID.MSG_Q_032, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If

            '処理確認
            If Message.ShowMsgBox(MessageID.MSG_Q_034, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If

            ' REV_002↓
            'Call getData(dtKobetsu, chosaNenKey, dtSetting, dtKangenKomokuMaster, dtChohyoMaster, dtTaiouMaster, dtShukei, keys)
            Call getData(dtKobetsu, chosaNenKey, dtSetting, dtKangenKomokuMaster, dtChohyoMaster, dtTaiouMaster, dtShukei, keys, dtSeidoUketoriItem)
            ' REV_002↑

            'チェックボックスの状態を配列に格納
            checkBox = Me.GetCheckBoxList

            'ファイル名を設定
            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                fileName = CStr(chosaNenKey(0)) & "_農業経営の概況_" & ComUtil.getEinouName(dtKobetsu(CStr(chosaNenKey(0)))) & "_" & dtKobetsu(CStr(chosaNenKey(0))).Item(CENSUS).値
            Else
                fileName = CStr(chosaNenKey(0)) & "_生産費の概要_" & ComConst.還元資料.帳票名(CommonInfo.Chosakubun) & "_" & dtKobetsu(CStr(chosaNenKey(0))).Item(CENSUS).値
            End If

            'フォルダパス取得
            filePath = ComUtil.GetFolderPath(Me, If(CommonInfo.SenmonChosain, IniFileInfo.SenmonchosainOutPath, IniFileInfo.ExcelOutPath))

            If filePath.Equals(String.Empty) Then
                Exit Sub
            End If

            filePath = filePath & "\" & fileName

            '還元資料設定値マスタ、個別結果表データ、集計結果表データ、還元資料帳票マスタを使用して、還元資料の作成を行う。
            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                ' REV_002↓
                'Using kangen = New CreateKangenEinou(checkBox, False, True, filePath, dtSetting, dtKobetsu, dtShukei, dtChohyoMaster, dtTaiouMaster, dtKangenKomokuMaster, chosaNenKey)
                Using kangen = New CreateKangenEinou(checkBox, False, True, filePath, dtSetting, dtKobetsu, dtShukei, dtChohyoMaster, dtTaiouMaster, dtKangenKomokuMaster, chosaNenKey, dtSeidoUketoriItem)
                    ' REV_002↑
                    kangen.Execute(MessageID.MSG_Q_004)
                End Using
            Else
                ' REV_002↓
                'Using kangen = New CreateKangenSeisan(checkBox, False, True, filePath, dtSetting, dtKobetsu, dtShukei, dtChohyoMaster, dtTaiouMaster, dtKangenKomokuMaster, chosaNenKey)
                Using kangen = New CreateKangenSeisan(checkBox, False, True, filePath, dtSetting, dtKobetsu, dtShukei, dtChohyoMaster, dtTaiouMaster, dtKangenKomokuMaster, chosaNenKey, dtSeidoUketoriItem)
                    ' REV_002↑
                    kangen.Execute(MessageID.MSG_Q_004)
                End Using
            End If


            '作成履歴を登録する。
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try
                    UpdateKangenRireki(db)
                Catch ex As Exception
                    InsertKangenRireki(db)
                End Try

            End Using

            Message.ShowMsgBox(MessageID.MSG_I_035, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region

#Region "処理全般"

    ''' <summary>
    ''' チェックボックスの名称変更
    ''' </summary>
    ''' <param name="pKubun1"></param>
    ''' <param name="pKubun2"></param>
    ''' <param name="pChosaKubun"></param>
    ''' <remarks></remarks>
    Private Sub SetCheckBoxControl(ByVal pChosaKubun As String)
        '営農個人の場合
        If (pChosaKubun = EINOU_KOJIN) Then
            SetCheckBoxPro(CheckBox1, CHK_KOJIN1, True, True)
            SetCheckBoxPro(CheckBox2, CHK_KOJIN2, True, True)
            SetCheckBoxPro(CheckBox3, CHK_KOJIN3, True, True)
            SetCheckBoxPro(CheckBox4, CHK_KOJIN4, True, True)
            SetCheckBoxPro(CheckBox5, CHK_KOJIN5, True, True)
            SetCheckBoxPro(CheckBox6, CHK_KOJIN6, True, True)
            SetCheckBoxPro(CheckBox7, CHK_KOJIN7, True, True)
            SetCheckBoxPro(CheckBox8, CHK_KOJIN8, True, True)
            SetCheckBoxPro(CheckBox9, CHK_KOJIN9, True, True)
            SetCheckBoxPro(CheckBox10, CHK_KOJIN10, True, True)
            SetCheckBoxPro(CheckBox11, String.Empty, False, False)
            '営農法人の場合
        ElseIf (pChosaKubun = EINOU_HOUJIN) Then
            SetCheckBoxPro(CheckBox1, CHK_HOUJIN1, True, True)
            SetCheckBoxPro(CheckBox2, CHK_HOUJIN2, True, True)
            SetCheckBoxPro(CheckBox3, CHK_HOUJIN3, True, True)
            SetCheckBoxPro(CheckBox4, CHK_HOUJIN4, True, True)
            SetCheckBoxPro(CheckBox5, CHK_HOUJIN5, True, True)
            SetCheckBoxPro(CheckBox6, CHK_HOUJIN6, True, True)
            SetCheckBoxPro(CheckBox7, CHK_HOUJIN7, True, True)
            SetCheckBoxPro(CheckBox8, CHK_HOUJIN8, True, True)
            SetCheckBoxPro(CheckBox9, CHK_HOUJIN9, True, True)
            SetCheckBoxPro(CheckBox10, CHK_HOUJIN10, True, True)
            SetCheckBoxPro(CheckBox11, String.Empty, False, False)
            '上記以外の場合、下記を行う
        Else
            SetCheckBoxPro(CheckBox1, CHK_SEISAN1, True, True)
            SetCheckBoxPro(CheckBox2, CHK_SEISAN2, True, True)
            SetCheckBoxPro(CheckBox3, CHK_SEISAN3, True, True)
            SetCheckBoxPro(CheckBox4, CHK_SEISAN4, True, True)
            SetCheckBoxPro(CheckBox5, CHK_SEISAN5, True, True)
            SetCheckBoxPro(CheckBox6, CHK_SEISAN6, True, True)
            SetCheckBoxPro(CheckBox7, CHK_SEISAN7, True, True)
            SetCheckBoxPro(CheckBox8, CHK_SEISAN8, True, True)
            SetCheckBoxPro(CheckBox9, CHK_SEISAN9, True, True)
            SetCheckBoxPro(CheckBox10, CHK_SEISAN10, True, True)
            SetCheckBoxPro(CheckBox11, CHK_SEISAN11, True, True)
        End If

    End Sub

    ''' <summary>
    ''' チェックボックスのプロパティ変更
    ''' </summary>
    ''' <param name="chkBox"></param>
    ''' <param name="showName"></param>
    ''' <param name="enabledFlag"></param>
    ''' <remarks></remarks>
    Private Sub SetCheckBoxPro(ByRef chkBox As System.Windows.Forms.CheckBox, ByVal showName As String, ByVal visibleFlag As Boolean, ByVal enabledFlag As Boolean)
        chkBox.Text = showName
        chkBox.Visible = visibleFlag
        chkBox.Enabled = enabledFlag

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        'シート選択のチェックボックスのチェック
        If (Me.CheckAllChkBox = False) Then
            'Print、ファイル出力ボタンが非活性化
            SetButtonControl(False)
        Else
            'Print、ファイル出力ボタンが活性化
            SetButtonControl(True)
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        'シート選択のチェックボックスのチェック
        If (Me.CheckAllChkBox = False) Then
            'Print、ファイル出力ボタンが非活性化
            SetButtonControl(False)
        Else
            'Print、ファイル出力ボタンが活性化
            SetButtonControl(True)
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        'シート選択のチェックボックスのチェック
        If (Me.CheckAllChkBox = False) Then
            'Print、ファイル出力ボタンが非活性化
            SetButtonControl(False)
        Else
            'Print、ファイル出力ボタンが活性化
            SetButtonControl(True)
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        'シート選択のチェックボックスのチェック
        If (Me.CheckAllChkBox = False) Then
            'Print、ファイル出力ボタンが非活性化
            SetButtonControl(False)
        Else
            'Print、ファイル出力ボタンが活性化
            SetButtonControl(True)
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        'シート選択のチェックボックスのチェック
        If (Me.CheckAllChkBox = False) Then
            'Print、ファイル出力ボタンが非活性化
            SetButtonControl(False)
        Else
            'Print、ファイル出力ボタンが活性化
            SetButtonControl(True)
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        'シート選択のチェックボックスのチェック
        If (Me.CheckAllChkBox = False) Then
            'Print、ファイル出力ボタンが非活性化
            SetButtonControl(False)
        Else
            'Print、ファイル出力ボタンが活性化
            SetButtonControl(True)
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        'シート選択のチェックボックスのチェック
        If (Me.CheckAllChkBox = False) Then
            'Print、ファイル出力ボタンが非活性化
            SetButtonControl(False)
        Else
            'Print、ファイル出力ボタンが活性化
            SetButtonControl(True)
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'シート選択のチェックボックスのチェック
        If (Me.CheckAllChkBox = False) Then
            'Print、ファイル出力ボタンが非活性化
            SetButtonControl(False)
        Else
            'Print、ファイル出力ボタンが活性化
            SetButtonControl(True)
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'シート選択のチェックボックスのチェック
        If (Me.CheckAllChkBox = False) Then
            'Print、ファイル出力ボタンが非活性化
            SetButtonControl(False)
        Else
            'Print、ファイル出力ボタンが活性化
            SetButtonControl(True)
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'シート選択のチェックボックスのチェック
        If (Me.CheckAllChkBox = False) Then
            'Print、ファイル出力ボタンが非活性化
            SetButtonControl(False)
        Else
            'Print、ファイル出力ボタンが活性化
            SetButtonControl(True)
        End If
    End Sub

    Private Sub CheckBox11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'シート選択のチェックボックスのチェック
        If Me.CheckAllChkBox = False Then
            'Print、ファイル出力ボタンが非活性化
            SetButtonControl(False)
        Else
            'Print、ファイル出力ボタンが活性化
            SetButtonControl(True)
        End If
    End Sub

    'チェックボックスが選択されているかをチェックする。
    Private Function CheckAllChkBox() As Boolean
        Dim checkResult As Boolean = False
        Dim tempSheelName As String

        tempSheelName = GetCheckBoxTextName(CheckBox1, True)
        If (tempSheelName <> String.Empty) Then
            checkResult = True
        End If

        If (checkResult = False) Then
            tempSheelName = GetCheckBoxTextName(CheckBox2, True)
            If (tempSheelName <> String.Empty) Then
                checkResult = True
            End If
        End If

        If (checkResult = False) Then
            tempSheelName = GetCheckBoxTextName(CheckBox3, True)
            If (tempSheelName <> String.Empty) Then
                checkResult = True
            End If
        End If

        If (checkResult = False) Then
            tempSheelName = GetCheckBoxTextName(CheckBox4, True)
            If (tempSheelName <> String.Empty) Then
                checkResult = True
            End If
        End If

        If (checkResult = False) Then
            tempSheelName = GetCheckBoxTextName(CheckBox5, True)
            If (tempSheelName <> String.Empty) Then
                checkResult = True
            End If
        End If

        If (checkResult = False) Then
            tempSheelName = GetCheckBoxTextName(CheckBox6, True)
            If (tempSheelName <> String.Empty) Then
                checkResult = True
            End If
        End If

        If (checkResult = False) Then
            tempSheelName = GetCheckBoxTextName(CheckBox7, True)
            If (tempSheelName <> String.Empty) Then
                checkResult = True
            End If
        End If

        If (checkResult = False) Then
            tempSheelName = GetCheckBoxTextName(CheckBox8, True)
            If (tempSheelName <> String.Empty) Then
                checkResult = True
            End If
        End If

        If (checkResult = False) Then
            tempSheelName = GetCheckBoxTextName(CheckBox9, True)
            If (tempSheelName <> String.Empty) Then
                checkResult = True
            End If
        End If

        If (checkResult = False) Then
            tempSheelName = GetCheckBoxTextName(CheckBox10, True)
            If (tempSheelName <> String.Empty) Then
                checkResult = True
            End If
        End If

        If (checkResult = False) Then
            tempSheelName = GetCheckBoxTextName(CheckBox11, True)
            If (tempSheelName <> String.Empty) Then
                checkResult = True
            End If
        End If

        Return checkResult
    End Function

    'チェックボックスの状態をリストとして取得する。
    Private Function GetCheckBoxList() As ArrayList
        Dim checkList As ArrayList = New ArrayList()
        Dim checkResult As Boolean = False
        Dim tempSheelName As String

        tempSheelName = GetCheckBoxTextName(CheckBox1, True)
        If (tempSheelName <> String.Empty) Then
            checkList.Add(True)
        Else
            checkList.Add(False)
        End If


        tempSheelName = GetCheckBoxTextName(CheckBox2, True)
        If (tempSheelName <> String.Empty) Then
            checkList.Add(True)
        Else
            checkList.Add(False)
        End If

        tempSheelName = GetCheckBoxTextName(CheckBox3, True)
        If (tempSheelName <> String.Empty) Then
            checkList.Add(True)
        Else
            checkList.Add(False)
        End If



        tempSheelName = GetCheckBoxTextName(CheckBox4, True)
        If (tempSheelName <> String.Empty) Then
            checkList.Add(True)
        Else
            checkList.Add(False)
        End If


        tempSheelName = GetCheckBoxTextName(CheckBox5, True)
        If (tempSheelName <> String.Empty) Then
            checkList.Add(True)
        Else
            checkList.Add(False)
        End If


        tempSheelName = GetCheckBoxTextName(CheckBox6, True)
        If (tempSheelName <> String.Empty) Then
            checkList.Add(True)
        Else
            checkList.Add(False)
        End If



        tempSheelName = GetCheckBoxTextName(CheckBox7, True)
        If (tempSheelName <> String.Empty) Then
            checkList.Add(True)
        Else
            checkList.Add(False)
        End If



        tempSheelName = GetCheckBoxTextName(CheckBox8, True)
        If (tempSheelName <> String.Empty) Then
            checkList.Add(True)
        Else
            checkList.Add(False)
        End If


        tempSheelName = GetCheckBoxTextName(CheckBox9, True)
        If (tempSheelName <> String.Empty) Then
            checkList.Add(True)
        Else
            checkList.Add(False)
        End If



        tempSheelName = GetCheckBoxTextName(CheckBox10, True)
        If (tempSheelName <> String.Empty) Then
            checkList.Add(True)
        Else
            checkList.Add(False)
        End If


        tempSheelName = GetCheckBoxTextName(CheckBox11, True)
        If (tempSheelName <> String.Empty) Then
            checkList.Add(True)
        Else
            checkList.Add(False)
        End If


        Return checkList
    End Function

    ''' <summary>
    ''' 経年個別結果表の初期設定
    ''' </summary>
    ''' <param name="chkBox"></param>
    ''' <param name="showName"></param>
    ''' <param name="enabledFlag"></param>
    ''' <remarks></remarks>
    Private Sub SetKeinenKobetuSentaku()


        Dim txtBoxs() As System.Windows.Forms.TextBox
        Dim todofukenCenterBg = ""
        Dim notification As String = ""

        If mKeinenKobetuInfo.Count > 0 Then
            For i As Integer = 0 To mKeinenKobetuInfo.Count - 1
                txtBoxs = CType(mKeinenKobetuInfo(i), TextBox())

                txtBoxs(INX_YEAR).Text = CStr(CInt(_ChosaNen) - i)
                txtBoxs(INX_CENSUS).Text = _CensusNo
                txtBoxs(INX_FIND).Text = TXT_FOUND
                If i = 0 Then
                    txtBoxs(INX_FIND).Text = TXT_FOUND

                    txtBoxs(INX_YEAR).Enabled = False
                    txtBoxs(INX_CENSUS).Enabled = False
                Else
                    txtBoxs(INX_FIND).Text = String.Empty
                End If

            Next
        End If

    End Sub

    Private Function chkTyosaYearTxtBox() As Boolean
        Dim txtBoxs As System.Windows.Forms.TextBox()  ' テキストボックスの配列
        Dim preYear As String = ""      ' 比較元の調査年
        Dim nextYear As String = ""     ' 比較対象の調査年
        Dim i As Integer = 0

        chkTyosaYearTxtBox = False

        For i = 0 To mKeinenKobetuInfo.Count - 1
            txtBoxs = CType(mKeinenKobetuInfo(i), TextBox())

            '一番上の場合、調査年を取得
            If i = 0 Then
                preYear = txtBoxs(INX_YEAR).Text
            Else
                nextYear = txtBoxs(INX_YEAR).Text
                '2年目以降で空欄の項目は無視する
                If (nextYear = String.Empty) Then
                    Continue For
                End If

                ' 比較対象の調査年が、比較元の調査年よりも等しいか大きかったらNG
                If CInt(preYear) <= CInt(nextYear) Then
                    chkTyosaYearTxtBox = False
                    Exit Function
                End If

                preYear = nextYear
            End If
        Next

        ' ここまできたらすべてOK
        chkTyosaYearTxtBox = True
    End Function

    ''' <summary>
    ''' 経年個別結果表有無ﾁｪｯｸ
    ''' </summary>
    ''' <param name="chkBox"></param>
    ''' <param name="showName"></param>
    ''' <param name="enabledFlag"></param>
    ''' <remarks></remarks>
    Private Function chkIsFindTxtBox() As Boolean
        Dim i As Integer = 0
        Dim notFoundFlg As Boolean = False
        Dim blankYearOrCenter As Boolean = False
        Dim txtBoxs() As System.Windows.Forms.TextBox
        chkIsFindTxtBox = False


        ' １つでも個別結果表有無が「無」であるかチェック
        For i = 0 To mKeinenKobetuInfo.Count - 1
            txtBoxs = CType(mKeinenKobetuInfo(i), TextBox())

            If txtBoxs(INX_FIND).Text = TXT_NOT_FOUND Then
                notFoundFlg = True
                Exit For
            End If
        Next

        '個別結果表有無が「有」以外で、調査年、経営体番号のいずれかが入力されているかのチェック
        For i = 0 To mKeinenKobetuInfo.Count - 1
            txtBoxs = CType(mKeinenKobetuInfo(i), TextBox())

            '個別結果表有無が「有」以外 
            If txtBoxs(INX_FIND).Text <> TXT_FOUND Then

                ' 調査年、経営体番号がどちらか一方でも空白ではない
                If Not (String.IsNullOrEmpty(txtBoxs(INX_YEAR).Text) And String.IsNullOrEmpty(txtBoxs(INX_CENSUS).Text)) Then
                    blankYearOrCenter = True
                End If
            End If
        Next

        If notFoundFlg Or blankYearOrCenter Then
            chkIsFindTxtBox = False
        Else
            chkIsFindTxtBox = True
        End If

    End Function

    '各データの取得
    Private Sub getData(ByRef pKobetsu As Dictionary(Of String, Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)), ByRef pChosanenKey As ArrayList,
                        ByRef pSetting As DataTable, ByRef pKoumokuMaster As DataTable, ByRef pChohyoMaster As DataTable, ByRef pTaiouMaster As DataTable,
                        ByRef pShukei As Dictionary(Of String, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)), pKobestuKey As List(Of DAOKobetsuKekkahyo.PrimaryKey),
                        ByRef pSeidoUketoriItem As DataTable) ' REV_002

        Dim dtItem As DataTable
        Dim setCensusNo As String
        Dim einou As String
        Dim chosakubun As Integer
        Dim setChosakubun As String

        '6－1 還元資料の作成
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '個別結果表項目マスタ取得
            ' REV_001↓
            'dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun)
            ' REV_003↓営農はバージョン区分に関わらず取得する
            'dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_ChosaNen, CommonInfo.Chosakubun))
            If _Chosakubun = EINOU_KOJIN OrElse _Chosakubun = EINOU_HOUJIN Then
                dtItem = New DataTable()
                '調査年のバージョン区分を優先してマージする
                Dim current = ComUtil.getVersionKubunTaikei(_ChosaNen, CommonInfo.Chosakubun)
                dtItem.Columns.Add("項目番号")
                dtItem.PrimaryKey = New DataColumn() {dtItem.Columns("項目番号")}
                For Each versionKbn In ComConst.バージョン区分.体系リスト.Keys
                    If versionKbn = current Then
                        Continue For
                    End If
                    dtItem.Merge(DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, versionKbn))
                Next
                dtItem.Merge(DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, current))
            Else
                dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_ChosaNen, CommonInfo.Chosakubun))
            End If
            ' REV_003↑
            ' REV_001↑

            '個別結果表からテーブルデータを取得
            For Each key As DAOKobetsuKekkahyo.PrimaryKey In pKobestuKey
                Dim dc As Dictionary(Of String, DataTable)
                Dim dc2 As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)
                dc = DAOKobetsuKekkahyo.GetTable(db, key)

                dc2 = ComUtil.KobetsuKekkahyo.GetItem(dtItem, dc)

                pKobetsu.Add(key.chosaNen, dc2)
                pChosanenKey.Add(key.chosaNen)
            Next
        End Using

        '還元資料設定テーブルからデータを取得
        'デフォルト値判定
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            If Me.CheckCencusExist(db, pKobestuKey(0).censusNo) Then
                setCensusNo = pKobestuKey(0).censusNo
            Else
                setCensusNo = DEFAULT_CSNSUS
            End If
        End Using

        If _Chosakubun = EINOU_KOJIN OrElse _Chosakubun = EINOU_HOUJIN Then
            einou = pKobetsu(CStr(pChosanenKey(0))).Item(ComConst.個別結果表.営農類型(CommonInfo.Chosakubun)).値
            If setCensusNo = DEFAULT_CSNSUS Then
                chosakubun = CInt(einou) * 1000 + CInt(CommonInfo.Chosakubun)
            Else
                chosakubun = CInt(CommonInfo.Chosakubun)
            End If

            setChosakubun = CStr(chosakubun)
        Else
            einou = "0"
            setChosakubun = CommonInfo.Chosakubun
        End If

        '還元資料設定値取得
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            pSetting = Me.GetKangenSetting(db, setChosakubun, setCensusNo)
        End Using

        '還元資料項目名マスタからテーブルデータを取得()
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            ' REV_002↓
            'pKoumokuMaster = Me.GetKangenKomokuMaster(db, CommonInfo.Chosakubun, einou)
            pKoumokuMaster = Me.GetKangenKomokuMaster(db, CommonInfo.Chosakubun, einou, ComUtil.getVersionKubunTaikei(_ChosaNen, CommonInfo.Chosakubun))
            ' REV_002↑
        End Using

        '還元資料固定項目定義マスタからテーブルデータを取得()
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            pChohyoMaster = Me.GetKangenChosaMaster(db, CommonInfo.Chosakubun, einou)
        End Using

        '還元資料個別結果表対応マスタからテーブルデータを取得
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            pTaiouMaster = Me.GetTaiouMaster(db, CommonInfo.Chosakubun)
        End Using

        '集計結果表からテーブルデータを取得
        '還元資料設定テーブルおよび還元資料項目名マスタテーブルの集計結果表データを取得
        Dim d1 As DataRow()
        d1 = pSetting.Select("項目番号 = '" & SHUKEIKEKKA & "'")

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

            '集計結果表項目マスタ取得
            ' REV_002↓
            'dtItem = DAOOther.GetSyukeiKekkahyoItemMaster(db, CommonInfo.Chosakubun, "0")
            dtItem = DAOOther.GetSyukeiKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_ChosaNen, CommonInfo.Chosakubun), "0")
            ' REV_002↑

            For Each d2 As DataRow In d1

                Dim values() As String = d2.Item("設定値").ToString.Split("_"c)
                If values.Count = 4 Then
                    '還元資料設定値より、使用する集計結果表のキーを取得
                    Dim dtShukeiKey As DataTable = DAOSyukeiKekkahyo.GetListKangen(db, CommonInfo.Chosakubun, values(0), values(1), values(2), values(3))

                    ' 設定値に該当する集計結果表を取得する
                    For Each drKey As DataRow In dtShukeiKey.Rows
                        Dim key As DAOSyukeiKekkahyo.PrimaryKey
                        Dim dc As Dictionary(Of String, DataTable)
                        Dim dc2 As Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)

                        'キー作成
                        key = New DAOSyukeiKekkahyo.PrimaryKey(
                            drKey.Item("調査年").ToString, drKey.Item("集計番号").ToString, drKey.Item("連番").ToString)

                        '集計結果表取得
                        dc = DAOSyukeiKekkahyo.GetTable(db, CommonInfo.Chosakubun, key)
                        dc2 = ComUtil.SyukeiKekkahyo.GetItem(dtItem, dc)

                        ' 集計データを明細番号(1～4)＋生産費平均値種類で管理
                        pShukei.Add(d2.Item("明細番号").ToString + "_" + dc(ComConst.集計結果表.テーブル名称(CommonInfo.Chosakubun)(0)).Rows(0).Item("生産費平均値種類").ToString, dc2)

                    Next
                End If

            Next

        End Using

        ' REV_002↓
        '還元資料項目名マスタからテーブルデータを取得()
        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            pSeidoUketoriItem = Me.GetSeidoUketoriItem(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_ChosaNen, CommonInfo.Chosakubun), _ChosaNen)
        End Using
        ' REV_002↑

    End Sub


    Private Function CheckCencusExist(db As DBAccess, pcensusNo As String) As Boolean
        Dim ret As Boolean
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim dr As SqlDataReader = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT *")
                .AppendLine(" FROM " & TBL_KANGENSETTING)
                .AppendLine(" WHERE センサス番号 = @センサス番号 ")
                .AppendLine("AND 調査区分 = @調査区分")
            End With
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pcensusNo))
            para.Add(db.CreateParameter("@調査区分", SqlDbType.VarChar, CommonInfo.Chosakubun))
            dr = db.ExecuteReader(sb.ToString, para)

            ret = dr.HasRows

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    Private Function GetKangenSetting(db As DBAccess, pChosakubun As String, pCensusNo As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT 調査区分,")
                .AppendLine(" 項目番号,")
                .AppendLine(" 明細番号,")
                .AppendLine(" 設定値")
                .AppendLine("FROM " & TBL_KANGENSETTING)
                .AppendLine(" WHERE 調査区分 = @調査区分 ")
                .AppendLine("and センサス番号 = @センサス番号 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, pChosakubun))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, pCensusNo))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    Private Function GetKangenChosaMaster(db As DBAccess, pChosakbn As String, einou As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * ")
                .AppendLine("FROM " & TBL_KANGENCHOHYO)
                .AppendLine(" WHERE 調査区分 = @調査区分 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, pChosakbn))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    Private Function GetTaiouMaster(db As DBAccess, pChosakubun As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * ")
                .AppendLine("FROM " & TBL_TAIOU)
                .AppendLine(" WHERE 調査区分 = @調査区分 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, pChosakubun))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_002↓
    'Private Function GetKangenKomokuMaster(db As DBAccess, pChosakubun As String, pEinokubun As String) As DataTable
    Private Function GetKangenKomokuMaster(db As DBAccess, pChosakubun As String, pEinokubun As String, versionKbn As String) As DataTable
        ' REV_002↑
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT * ")
                .AppendLine("FROM " & TBL_KANGENKOUMOKU)
                .AppendLine(" WHERE 調査区分 = @調査区分 ")
                .AppendLine(String.Format(" AND 営農類型 in (0,{0})", pEinokubun))
                .AppendLine(" AND   バージョン区分 = @バージョン区分 ")        ' REV_002
                .AppendLine(" ORDER BY 営農類型 DESC ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, pChosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))   ' REV_002

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ' REV_002↓
    Private Function GetSeidoUketoriItem(db As DBAccess, pChosakubun As String, versionKbn As String, chosaNen As String) As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("SELECT A.バージョン区分 ")
                .AppendLine("      ,A.調査区分 ")
                .AppendLine("      ,A.項目番号 ")
                .AppendLine("      ,A.項目名 ")
                .AppendLine("      ,A.計算式 ")
                .AppendLine("      ,B.出力項目名 ")
                .AppendLine("FROM (")
                .AppendLine("      SELECT * ")
                .AppendLine("      FROM   個別結果表作成論理 ")
                .AppendLine("      WHERE  調査区分 = @調査区分 ")
                .AppendLine("      AND    バージョン区分 = @バージョン区分 ) A ")
                .AppendLine("LEFT JOIN (")
                .AppendLine("      SELECT * ")
                .AppendLine("      FROM   制度受取金・積立金等項目 ")
                .AppendLine("      WHERE  調査年 = @調査年 ")
                .AppendLine("      AND    調査区分 = @調査区分 ")
                .AppendLine(") B ")
                .AppendLine("ON A.計算式 = '[' + B.項番 + ']' ")
                .AppendLine("WHERE B.出力項目名 IS NOT NULL ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, pChosakubun))
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, versionKbn))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))

            ret = db.GetDataTable(sb.ToString, para)

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function
    ' REV_002↑

    Private Function GetEinouKubunData(db As DBAccess) As String
        Dim table As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Dim ret As String = String.Empty

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                Select Case CommonInfo.Chosakubun
                    Case ComConst.調査区分.営農類型別経営統計_個人
                        .AppendLine("SELECT K000005 As 営農類型 ")
                    Case ComConst.調査区分.営農類型別経営統計_法人
                        .AppendLine("SELECT K000006 As 営農類型 ")
                    Case Else
                        Return ret
                End Select
                .AppendLine(String.Format("FROM   ""{0}"" AS K", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine(" WHERE センサス番号 = @センサス番号")
                .AppendLine(" AND 調査年 = @調査年")
            End With

            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, _CensusNo))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, _ChosaNen))

            table = db.GetDataTable(sb.ToString, para)

            If table.Rows.Count <> 0 Then
                ret = CStr(table.Rows(0).Item("営農類型"))
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    Private Function GetKiboKaisoData(db As DBAccess) As Integer
        Dim table As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Dim ret As Integer = -1

        Try
            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine(String.Format("SELECT ""{0}"" As 規模階層 ", ComConst.個別結果表.規模階層(CommonInfo.Chosakubun)))
                .AppendLine(String.Format("FROM   ""{0}"" AS K", ComConst.個別結果表.テーブル名称(CommonInfo.Chosakubun)(0)))
                .AppendLine(" WHERE センサス番号 = @センサス番号")
                .AppendLine(" AND 調査年 = @調査年")
            End With

            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, _CensusNo))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, _ChosaNen))

            table = db.GetDataTable(sb.ToString, para)

            If table.Rows.Count <> 0 Then
                ret = CInt(table.Rows(0).Item("規模階層"))
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    Private Function InsertKangenRireki(db As DBAccess) As Boolean
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim ret As Boolean

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("INSERT INTO " & TBL_KANGENKANRI)
                .AppendLine("( ")
                .AppendLine(" 調査区分 ")
                .AppendLine(",調査年 ")
                .AppendLine(",センサス番号 ")
                .AppendLine(",還元資料作成日時 ")
                .AppendLine(",更新日付 ")
                .AppendLine(",更新者ID ")
                .AppendLine(") ")
                .AppendLine("VALUES ")
                .AppendLine("( ")
                .AppendLine(" @調査区分")
                .AppendLine(",@調査年")
                .AppendLine(",@センサス番号 ")
                .AppendLine(",GETDATE() ")
                .AppendLine(",GETDATE() ")
                .AppendLine(",@更新者ID")
                .AppendLine(") ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, _ChosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, _CensusNo))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("還元資料作成履歴登録失敗")
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    Private Function UpdateKangenRireki(db As DBAccess) As Boolean
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing
        Dim ret As Boolean

        Try

            sb = New System.Text.StringBuilder
            para = New List(Of DBAccess.Parameter)

            ' SQL文の設定
            With sb
                .AppendLine("UPDATE " & TBL_KANGENKANRI)
                .AppendLine(" SET 還元資料作成日時 = GETDATE()")
                .AppendLine(",更新日付 = GETDATE()")
                .AppendLine(",更新者ID = @更新者ID")
                .AppendLine(" WHERE ")
                .AppendLine("調査区分 = @調査区分")
                .AppendLine("AND 調査年 = @調査年")
                .AppendLine("AND センサス番号 = @センサス番号 ")
            End With

            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, _ChosaNen))
            para.Add(db.CreateParameter("@センサス番号", SqlDbType.VarChar, _CensusNo))
            para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))

            If db.ExecuteNonQuery(sb.ToString, para) = 1 Then
                ret = True
            Else
                Throw New Exception("還元資料作成履歴更新失敗")
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    Private Function GetCheckBoxTextName(ByRef chkBox As CheckBox, ByVal pCheckboxValue As Boolean) As String
        Dim checkBoxName As String = String.Empty

        If (pCheckboxValue = False) Then
            checkBoxName = chkBox.Text
        End If

        If (chkBox.Enabled = True) Then
            If (chkBox.Checked = pCheckboxValue) Then
                checkBoxName = chkBox.Text
            Else
                checkBoxName = String.Empty
            End If
        End If
        Return checkBoxName
    End Function

    Private Sub SetButtonControl(ByVal pEnableFlag As Boolean)
        'Print、ファイル出力ボタンの制御
        Me.btnPrint.Enabled = pEnableFlag
        Me.btnMake.Enabled = pEnableFlag
    End Sub


    Private Sub txtChosaYear1_TextChanged(sender As Object, e As EventArgs) Handles txtChosaYear1.TextChanged
        Call KeinenKobetuControl(1)
    End Sub

    Private Sub txtChosaYear2_TextChanged(sender As Object, e As EventArgs) Handles txtChosaYear2.TextChanged
        Call KeinenKobetuControl(2)
    End Sub

    Private Sub txtChosaYear3_TextChanged(sender As Object, e As EventArgs) Handles txtChosaYear3.TextChanged
        Call KeinenKobetuControl(3)
    End Sub

    Private Sub txtChosaYear4_TextChanged(sender As Object, e As EventArgs) Handles txtChosaYear4.TextChanged
        Call KeinenKobetuControl(4)
    End Sub

    Private Sub txtChosaYear5_TextChanged(sender As Object, e As EventArgs) Handles txtChosaYear5.TextChanged
        Call KeinenKobetuControl(5)
    End Sub

    Private Sub txtPartCenterBg1_TextChanged(sender As Object, e As EventArgs) Handles txtPartCenterBg1.TextChanged
        Call KeinenKobetuControl(1)
    End Sub

    Private Sub txtPartCenterBg2_TextChanged(sender As Object, e As EventArgs) Handles txtPartCenterBg2.TextChanged
        Call KeinenKobetuControl(2)
    End Sub

    Private Sub txtPartCenterBg3_TextChanged(sender As Object, e As EventArgs) Handles txtPartCenterBg3.TextChanged
        Call KeinenKobetuControl(3)
    End Sub

    Private Sub txtPartCenterBg4_TextChanged(sender As Object, e As EventArgs) Handles txtPartCenterBg4.TextChanged
        Call KeinenKobetuControl(4)
    End Sub

    Private Sub txtPartCenterBg5_TextChanged(sender As Object, e As EventArgs) Handles txtPartCenterBg5.TextChanged
        Call KeinenKobetuControl(5)
    End Sub

    Private Sub KeinenKobetuControl(ByVal txtNo As Integer)
        Dim txtBoxs() As System.Windows.Forms.TextBox
        If startFlg = False Then
            txtBoxs = CType(mKeinenKobetuInfo(txtNo - 1), TextBox())

            txtBoxs(INX_FIND).Text = String.Empty
            txtBoxs(INX_YEAR).BackColor = Color.White
            txtBoxs(INX_CENSUS).BackColor = Color.White
        End If

    End Sub

#End Region


End Class