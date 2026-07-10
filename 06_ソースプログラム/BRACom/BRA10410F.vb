'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_000   | 2023.01.13 |大興電子通信        | 要件No.5 新規作成
'//            |            |                    |
'//*************************************************************************************************

Imports System.Text
Imports Microsoft.Office.Interop.Excel
''' <summary>
''' 規模区分別期待標本及び予定経営体数出力画面
''' </summary>
''' <remarks></remarks>
Public Class BRA10410F

    ''' <summary>
    ''' 初期表示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BRA10410F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'センサス実施年設定
            HyohonUtil.SetCombobox(CboCensusNen, HyohonConst.センサス実施年リスト)
            CboCensusNen.SelectedIndex = 0
            '地方農政局等プルダウン設定
            HyohonUtil.SetCombobox(CboKyokuTo, HyohonUtil.Get地方農政局等マスタリスト())
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
    ''' 様式出力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnPrint_Click(sender As Object, e As EventArgs) Handles BtnPrint.Click
        Try
            '入力チェック
            '地方農政局等（必須）
            Dim cbo = CboKyokuTo
            If HyohonUtil.IsEmpty(cbo, (HyohonConst.識別対象外値, HyohonConst.識別対象外値)) Then
                cbo.Focus()
                Message.ShowMsgBox(MessageID.MSG_E_134, MsgBoxStyle.OkOnly)
                Return
            End If
            '経営形態区分（必須）
            cbo = CboKeieiKeitai
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

            '標本リストCSV選択
            Dim fileName = Now.ToString("yyyyMMdd_HHmmss") & "規模区分別期待標本数及び予定経営体数一覧表.xlsx"
            Dim filePath = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName, "Excelブック|*.xlsx")
            If String.IsNullOrEmpty(filePath) Then
                Return
            End If

            'Excel出力
            Cursor.Current = Cursors.WaitCursor
            Using excelOutput = New ExcelOutput(filePath, HyohonUtil.GetComboValue(CboCensusNen),
                                                HyohonUtil.GetComboValue(CboKyokuTo, (HyohonConst.識別対象外値, HyohonConst.識別対象外値)),
                                                HyohonUtil.GetComboValue(CboKeieiKeitai),
                                                HyohonUtil.GetComboValue(CboEinoruikei), HyohonUtil.GetComboValue(CboShufukugyo, HyohonConst.主副業区分リスト(0).Key),
                                                HyohonUtil.GetComboValue(CboAoiro), HyohonUtil.GetComboValue(CboShurakuEino),
                                                HyohonUtil.GetComboValue(CboSeisanhi), HyohonUtil.GetComboValue(CboTahata))
                excelOutput.Execute(Me, overWriteMessageID:=MessageID.MSG_Q_004)
            End Using

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
    ''' Excel出力クラス
    ''' </summary>
    Private Class ExcelOutput
        Inherits ExcelOutputSingleBaseClass

        ''' <summary>
        ''' セルのRANGE
        ''' </summary>
        Private Class CELL_RANGE
            Public Const 地方農政局等 = "A4:B4"
            Public Const 経営形態区分 = "C4:D4"
            Public Const 営農類型区分 = "E4:F4"
            Public Const 主副業区分 = "G4:H4"
            Public Const 青色申告区分 = "I4:J4"
            Public Const 集落営農区分 = "K4:L4"
            Public Const 生産費区分 = "M4:N4"
            Public Const 田畑区分 = "O4:P4"
            Public Const 上限 = "D7"
            Public Const 下限 = "D8"
            Public Const 期待する標本の大きさ = "D9"
            Public Const 選定予定経営体＿名称 = "B11"
            Public Const 選定予定経営体＿件数 = "D11"
            Public Const 行コピー元 = "12:12"
        End Class

        ''' <summary>
        ''' テンプレートの選定予定経営体の行数
        ''' </summary>
        Private Const 選定予定経営体行数 = 10

        ''' <summary>
        ''' センサス実施年
        ''' </summary>
        Private ReadOnly CensusNen As Integer
        ''' <summary>
        ''' 地方農政局等
        ''' </summary>
        Private ReadOnly KyokuTo As (kyoku As Integer, pref As Integer)
        ''' <summary>
        ''' 経営形態区分
        ''' </summary>
        Private ReadOnly KeieiKeitai As Integer
        ''' <summary>
        ''' 営農類型区分
        ''' </summary>
        Private ReadOnly EinoRuikei As Integer
        ''' <summary>
        ''' 主副業区分
        ''' </summary>
        Private ReadOnly Shufukugyo As (選択肢 As Integer, 値リスト As List(Of Integer))
        ''' <summary>
        ''' 青色申告区分
        ''' </summary>
        Private ReadOnly Aoiro As Integer
        ''' <summary>
        ''' 集落営農区分
        ''' </summary>
        Private ReadOnly ShurakuEino As Integer
        ''' <summary>
        ''' 生産費区分
        ''' </summary>
        Private ReadOnly Seisanhi As Integer
        ''' <summary>
        ''' 田畑区分
        ''' </summary>
        Private ReadOnly Tahata As Integer

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="outPath"></param>
        ''' <param name="censusNen"></param>
        ''' <param name="kyokuTo"></param>
        ''' <param name="keieiKeitai"></param>
        ''' <param name="einoRuikei"></param>
        ''' <param name="shufukugyo"></param>
        ''' <param name="aoiro"></param>
        ''' <param name="shurakuEino"></param>
        ''' <param name="seisanhi"></param>
        ''' <param name="tahata"></param>
        Sub New(outPath As String, censusNen As Integer, kyokuTo As (Integer, Integer), keieiKeitai As Integer, einoRuikei As Integer, shufukugyo As (選択肢 As Integer, 値リスト As List(Of Integer)), aoiro As Integer, shurakuEino As Integer, seisanhi As Integer, tahata As Integer)
            MyBase.New("規模区分別期待する標本の大きさ及び予定経営体数一覧表.xlsx", True, False, "規模区分別期待する標本の大きさ及び予定経営体数一覧表", outPath, True)
            Me.CensusNen = censusNen
            Me.KyokuTo = kyokuTo
            Me.KeieiKeitai = keieiKeitai
            Me.EinoRuikei = einoRuikei
            Me.Shufukugyo = shufukugyo
            Me.Aoiro = aoiro
            Me.ShurakuEino = shurakuEino
            Me.Seisanhi = seisanhi
            Me.Tahata = tahata
        End Sub

        ''' <summary>
        ''' 帳票編集
        ''' </summary>
        ''' <param name="xlSheets"></param>
        Protected Overrides Sub ReportEdit(xlSheets As Sheets)
            Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '帳票データを検索
                '抽出条件
                Dim id = HyohonUtil.Get標本リスト設定ID(db, KeieiKeitai, EinoRuikei, Shufukugyo.選択肢, Aoiro, ShurakuEino, Seisanhi, Tahata)
                Dim dtJoken = HyohonUtil.Select標本リスト設定＿抽出条件(db, id)
                '期待値
                Dim dtKitai = HyohonUtil.Select標本リスト設定＿期待値(db, id, KyokuTo.kyoku)
                '選定予定経営体＿名称
                Dim senteiNames = Get選定予定経営体＿名称(db, KyokuTo)
                '選定予定経営体＿I～XIII
                Dim dtKensu = Select選定予定経営体＿I＿XIII(db, dtJoken)

                'シートに出力
                Dim sheet = DirectCast(xlSheets(1), Worksheet)
                '指標部
                SetValue(sheet, CELL_RANGE.地方農政局等, HyohonUtil.Getリスト名称(HyohonUtil.Get地方農政局等マスタリスト(), KyokuTo))
                SetValue(sheet, CELL_RANGE.経営形態区分, HyohonUtil.Getリスト名称(HyohonUtil.Get経営形態区分リスト(), KeieiKeitai))
                SetValue(sheet, CELL_RANGE.営農類型区分, HyohonUtil.Getリスト名称(HyohonConst.営農類型区分リスト, EinoRuikei))
                SetValue(sheet, CELL_RANGE.主副業区分, HyohonUtil.Getリスト名称(HyohonConst.主副業区分リスト, Shufukugyo))
                SetValue(sheet, CELL_RANGE.青色申告区分, HyohonUtil.Getリスト名称(HyohonConst.青色申告区分リスト, Aoiro))
                SetValue(sheet, CELL_RANGE.集落営農区分, HyohonUtil.Getリスト名称(HyohonConst.集落営農区分リスト, ShurakuEino))
                If Seisanhi <> HyohonConst.識別対象外値 Then
                    SetValue(sheet, CELL_RANGE.生産費区分, HyohonUtil.Getリスト名称(HyohonUtil.Get生産費区分リスト(DirectCast(KeieiKeitai, HyohonConst.経営形態区分)), Seisanhi) & "生産費")
                End If
                SetValue(sheet, CELL_RANGE.田畑区分, HyohonUtil.Getリスト名称(HyohonConst.田畑区分リスト, Tahata))
                '抽出条件
                For Each row As DataRow In dtJoken.Rows()
                    Dim kibo = CInt(row("規模番号"))
                    Dim val = If(row("上限"), "").ToString()
                    SetValue(sheet, CELL_RANGE.上限, val, 0, kibo - 1)
                    val = If(row("下限"), "").ToString()
                    SetValue(sheet, CELL_RANGE.下限, val, 0, kibo - 1)
                Next
                '期待値
                For Each row As DataRow In dtKitai.Rows
                    Dim kibo = CInt(row("規模番号"))
                    Dim val = row("標本サイズ").ToString()
                    If val = "0" Then
                        val = ""
                    End If
                    SetValue(sheet, CELL_RANGE.期待する標本の大きさ, val, 0, kibo - 1)
                Next
                '選定予定経営体＿名称
                If senteiNames.Count > 選定予定経営体行数 Then
                    '行追加
                    Dim rng As Range = Nothing
                    Try
                        rng = sheet.Range(CELL_RANGE.行コピー元)
                        For i = 1 To senteiNames.Count - 選定予定経営体行数
                            rng.Copy()
                            rng.Insert()
                        Next
                    Finally
                        ReleaseComObject(rng)
                    End Try
                End If
                For i = 0 To UBound(senteiNames)
                    SetValue(sheet, CELL_RANGE.選定予定経営体＿名称, senteiNames(i), i, 0)
                Next
                '選定予定経営体＿I～XIII
                For Each row As DataRow In dtKensu.Rows
                    Dim rowOffset = senteiNames.Select(Function(x, index) (x, index)).Where(Function(x) x.x = row("名称").ToString()).FirstOrDefault().index
                    For i = 1 To HyohonConst.規模階層区分リスト.Count - 1
                        Dim val = CInt(row(String.Format("規模番号{0}", i)))
                        SetValue(sheet, CELL_RANGE.選定予定経営体＿件数, If(val = 0, "", val.ToString()), rowOffset, i - 1)
                    Next
                Next
            End Using
        End Sub

        ''' <summary>
        ''' 選定予定経営体の名称を取得
        ''' </summary>
        ''' <param name="db"></param>
        ''' <param name="kyokuTo"></param>
        ''' <returns></returns>
        Private Shared Function Get選定予定経営体＿名称(db As DBAccess, kyokuTo As (kyoku As Integer, pref As Integer)) As String()
            Dim sbSql = New StringBuilder()
            Dim para = New List(Of DBAccess.Parameter)
            If kyokuTo.kyoku = HyohonConst.農政局＿全国 OrElse kyokuTo.pref = HyohonConst.都道府県＿全国＿農政局 Then
                sbSql.Append("SELECT 名称 FROM 地方農政局等マスタ")
                If kyokuTo.kyoku = HyohonConst.農政局＿全国 Then
                    '全農政局
                    sbSql.Append(" WHERE 農政局 <> @農政局")
                    para.Add(db.CreateParameter("@農政局", SqlDbType.Int, HyohonConst.農政局＿全国))
                    sbSql.Append(" AND 都道府県 = @都道府県")
                    para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, HyohonConst.都道府県＿全国＿農政局))
                Else
                    '管轄する全都道府県
                    sbSql.Append(" WHERE 農政局 = @農政局")
                    para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kyokuTo.kyoku))
                    sbSql.Append(" AND 都道府県 <> @都道府県")
                    para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kyokuTo.pref))
                End If
                sbSql.Append(" ORDER BY 都道府県")
            Else
                '都道府県 or 拠点名
                sbSql.Append("SELECT ISNULL(KM.拠点名, CM.名称) AS 名称")
                sbSql.Append(" FROM 地方農政局等マスタ CM")
                sbSql.Append(" LEFT JOIN 標本リスト＿拠点マスタ KM")
                sbSql.Append(" ON KM.都道府県 = CM.都道府県")
                sbSql.Append(" AND KM.市区町村 <> @市区町村＿１拠点")
                para.Add(db.CreateParameter("@市区町村＿１拠点", SqlDbType.Int, HyohonConst.市区町村＿１拠点))
                sbSql.Append(" WHERE CM.農政局 = @農政局")
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, kyokuTo.kyoku))
                sbSql.Append(" AND CM.都道府県 = @都道府県")
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, kyokuTo.pref))
                sbSql.Append(" GROUP BY ISNULL(KM.拠点名, CM.名称), CM.都道府県")
                sbSql.Append(" ORDER BY MIN(KM.実査設置拠点), CM.都道府県")
            End If
            Return db.GetDataTable(sbSql.ToString(), para).Rows().Cast(Of DataRow).Select(Function(x) x("名称").ToString()).ToArray()
        End Function

        ''' <summary>
        ''' 選定予定経営体＿I～XIII検索
        ''' </summary>
        ''' <param name="db"></param>
        ''' <param name="dtJoken"></param>
        ''' <returns></returns>
        Private Function Select選定予定経営体＿I＿XIII(db As DBAccess, dtJoken As Data.DataTable) As Data.DataTable
            '標本区分列番号
            Dim hyohonKbnNo = HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.標本区分)
            '比較列番号
            Dim hikakuNo = HyohonUtil.Get標本リスト比較列番号(CensusNen, DirectCast(KeieiKeitai, HyohonConst.経営形態区分),
                                                    DirectCast(EinoRuikei, HyohonConst.営農類型区分),
                                                    DirectCast(ShurakuEino, HyohonConst.集落営農区分),
                                                    DirectCast(Seisanhi, HyohonConst.生産費区分),
                                                    DirectCast(Tahata, HyohonConst.田畑区分))

            Dim sbSql = New StringBuilder()
            Dim para = New List(Of DBAccess.Parameter)
            sbSql.Append("SELECT")
            Dim lower As Decimal
            Dim upper As Decimal
            For i = 0 To dtJoken.Rows.Count - 1
                If Not Decimal.TryParse(If(dtJoken.Rows(i)("下限"), "").ToString(), lower) Then
                    lower = -HyohonConst.標本リスト設定＿抽出上限値
                End If
                If Not Decimal.TryParse(If(dtJoken.Rows(i)("上限"), "").ToString(), upper) Then
                    upper = If(lower <= 0, lower, HyohonConst.標本リスト設定＿抽出上限値)
                End If
                '下限以上～上限未満
                sbSql.Append(String.Format(" SUM(IIF({0} <= NO{1} AND NO{1} < {2}, 1, 0)) AS 規模番号{3},", lower, hikakuNo, upper, i + 1))
            Next
            If KyokuTo.kyoku = HyohonConst.農政局＿全国 OrElse KyokuTo.pref = HyohonConst.都道府県＿全国＿農政局 Then
                '全農政局 or 管轄する全都道府県
                sbSql.Append(" CM.名称")
                sbSql.Append(String.Format(" FROM 標本リスト{0} HL", CensusNen))
                sbSql.Append(" JOIN 地方農政局等マスタ CM")
                sbSql.Append(" ON CM.農政局 = HL.農政局")
                If KyokuTo.kyoku = HyohonConst.農政局＿全国 Then
                    '全農政局
                    sbSql.Append(" WHERE CM.農政局 <> @農政局")
                    para.Add(db.CreateParameter("@農政局", SqlDbType.Int, HyohonConst.農政局＿全国))
                    sbSql.Append(" AND CM.都道府県 = @都道府県")
                    para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, HyohonConst.都道府県＿全国＿農政局))
                Else
                    '管轄する全都道府県
                    sbSql.Append(" AND CM.都道府県 = HL.都道府県")
                    sbSql.Append(" WHERE CM.農政局 = @農政局")
                    para.Add(db.CreateParameter("@農政局", SqlDbType.Int, KyokuTo.kyoku))
                    sbSql.Append(" AND CM.都道府県 <> @都道府県")
                    para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, HyohonConst.都道府県＿全国＿農政局))
                End If
            Else
                '都道府県 or 拠点名
                sbSql.Append(" IIF(KM.市区町村 = @市区町村＿１拠点, CM.名称, KM.拠点名) AS 名称")
                para.Add(db.CreateParameter("@市区町村＿１拠点", SqlDbType.Int, HyohonConst.市区町村＿１拠点))
                sbSql.Append(String.Format(" FROM 標本リスト{0} HL", CensusNen))
                sbSql.Append(" JOIN 地方農政局等マスタ CM")
                sbSql.Append(" ON CM.農政局 = HL.農政局")
                sbSql.Append(" AND CM.都道府県 = HL.都道府県")
                sbSql.Append(" JOIN 標本リスト＿拠点マスタ KM")
                sbSql.Append(" ON KM.都道府県 = HL.都道府県")
                sbSql.Append(" AND (KM.市区町村 = HL.NO4 OR KM.市区町村 = @市区町村＿１拠点)")
                sbSql.Append(" WHERE HL.農政局 = @農政局")
                para.Add(db.CreateParameter("@農政局", SqlDbType.Int, KyokuTo.kyoku))
                sbSql.Append(" AND HL.都道府県 = @都道府県")
                para.Add(db.CreateParameter("@都道府県", SqlDbType.Int, KyokuTo.pref))
            End If
            sbSql.Append(" AND HL.経営形態区分 = @経営形態区分")
            para.Add(db.CreateParameter("@経営形態区分", SqlDbType.Int, KeieiKeitai))
            sbSql.Append(" AND HL.営農類型区分 = @営農類型区分")
            para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, EinoRuikei))
            sbSql.Append(" AND HL.生産費区分 = @生産費区分")
            para.Add(db.CreateParameter("@生産費区分", SqlDbType.Int, Seisanhi))
            If Tahata <> HyohonConst.識別対象外値 Then
                sbSql.Append(" AND HL.田畑区分 = @田畑区分")
                para.Add(db.CreateParameter("@田畑区分", SqlDbType.Int, Tahata))
                If Seisanhi = HyohonConst.生産費区分.小麦 Then
                    '小麦
                    sbSql.Append(String.Format(" AND HL.NO{0} {1} HL.NO{2}",
                        HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.小麦生産費＿田規模),
                        If(Tahata = HyohonConst.田畑区分.田, ">=", "<="),
                        HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.小麦生産費＿畑規模)))
                Else
                    '大豆
                    sbSql.Append(String.Format(" AND HL.NO{0} {1} HL.NO{2}",
                        HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.大豆生産費＿田規模),
                        If(Tahata = HyohonConst.田畑区分.田, ">=", "<="),
                        HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.大豆生産費＿畑規模)))
                End If
            End If
            If Shufukugyo.選択肢 <> HyohonConst.識別対象外値 Then
                sbSql.Append(String.Format(" AND HL.NO{0} IN (", HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.主副業区分)))
                For i = 0 To Shufukugyo.値リスト.Count - 1
                    If i >= 1 Then
                        sbSql.Append(",")
                    End If
                    sbSql.Append(String.Format("@主副業区分{0}", i))
                    para.Add(db.CreateParameter(String.Format("@主副業区分{0}", i), SqlDbType.Int, Shufukugyo.値リスト(i)))
                Next
                sbSql.Append(")")
            End If
            If Aoiro <> HyohonConst.識別対象外値 Then
                sbSql.Append(String.Format(" AND HL.NO{0} = @青色申告区分", HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.青色申告区分)))
                para.Add(db.CreateParameter("@青色申告区分", SqlDbType.Int, Aoiro))
            End If
            If ShurakuEino <> HyohonConst.識別対象外値 Then
                sbSql.Append(String.Format(" AND HL.NO{0} = @集落営農区分", HyohonUtil.Get標本リスト列番号(CensusNen, HyohonConst.標本リスト列番号キー.集落営農区分)))
                para.Add(db.CreateParameter("@集落営農区分", SqlDbType.Int, ShurakuEino))
            End If
            '工程による条件
            Select Case CommonInfo.Koutei
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    '農政局工程
                    sbSql.Append(" AND HL.農政局 = @利用者の農政局")
                    para.Add(db.CreateParameter("@利用者の農政局", SqlDbType.Int, CommonInfo.Kyoku))
                Case CommonInfo.KouteiKubun.Code.Center
                    '実査設置拠点工程
                    sbSql.Append(" AND HL.農政局 = @利用者の農政局")
                    para.Add(db.CreateParameter("@利用者の農政局", SqlDbType.Int, CommonInfo.Kyoku))
                    sbSql.Append(" AND HL.都道府県 = @利用者の都道府県")
                    para.Add(db.CreateParameter("@利用者の都道府県", SqlDbType.Int, If(CommonInfo.Jimusyo = "51", "1", CommonInfo.Jimusyo)))
                    sbSql.Append(" AND HL.実査設置拠点 = @利用者の実査設置拠点")
                    para.Add(db.CreateParameter("@利用者の実査設置拠点", SqlDbType.Int, CommonInfo.Center))
            End Select
            '標本区分による条件
            sbSql.Append(String.Format(" AND HL.NO{0} = @標本区分", hyohonKbnNo))
            para.Add(db.CreateParameter("@標本区分", SqlDbType.Int, HyohonConst.標本区分.現標本))
            '集約
            sbSql.Append(" GROUP BY")
            If KyokuTo.kyoku = HyohonConst.農政局＿全国 OrElse KyokuTo.pref = HyohonConst.都道府県＿全国＿農政局 Then
                '全農政局 or 管轄する全都道府県
                sbSql.Append(" CM.名称")
            Else
                '都道府県 or 拠点名
                sbSql.Append(" IIF(KM.市区町村 = @市区町村＿１拠点, CM.名称, KM.拠点名)")
            End If

            Return db.GetDataTable(sbSql.ToString(), para)
        End Function

        ''' <summary>
        ''' 指定のRangeに値をセット
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="range"></param>
        ''' <param name="value"></param>
        ''' <param name="rowOffset"></param>
        ''' <param name="colOffset"></param>
        Private Sub SetValue(sheet As Worksheet, range As String, value As String, Optional rowOffset As Integer = 0, Optional colOffset As Integer = 0)
            Dim rng As Range = Nothing
            Try
                rng = sheet.Range(range)
                If rowOffset = 0 AndAlso colOffset = 0 Then
                    rng.Value = value
                    Return
                End If
                Dim rngOffset As Range = Nothing
                Try
                    rngOffset = rng.Offset(rowOffset, colOffset)
                    rngOffset.Value = value
                Finally
                    ReleaseComObject(rngOffset)
                End Try
            Finally
                ReleaseComObject(rng)
            End Try
        End Sub
    End Class
End Class
