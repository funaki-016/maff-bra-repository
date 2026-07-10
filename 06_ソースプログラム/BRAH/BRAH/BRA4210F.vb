'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_000   | 2023.10.30 |大興電子通信        | 要件No.9 新規作成
'//  REV_001   | 2024.03.08 |大興電子通信        | 受入テスト不具合対応（連絡票No.211）
'//            |            |                    |
'//*************************************************************************************************
Imports System.Reflection
Imports System.Text
Imports Maff.BRA.Extensions.CellEx
Imports Maff.BRA.Extensions.DataRowEx

''' <summary>
''' 集計倍率管理画面
''' </summary>
''' <remarks></remarks>
Public Class BRA4210F

    ''' <summary>
    ''' 一覧列番号
    ''' </summary>
    Private Enum DGV_COL
        営農類型 = 0
        地域
        規模区分
        母集団
        集計対象数
        営農集計倍率
        非表示＿集落営農区分
        非表示＿地域区分
    End Enum

    ''' <summary>
    ''' 個人か否か
    ''' </summary>
    Private IsKojin As Boolean

    ''' <summary>
    ''' 表示調査年
    ''' </summary>
    Private ShowChosaNen As String = ""
    ''' <summary>
    ''' 表示営農類型区分
    ''' </summary>
    Private ShowEinoruikei As String = ""
    ''' <summary>
    ''' 表示集落営農区分
    ''' </summary>
    Private ShowShurakueino As String = ""
    ''' <summary>
    ''' 選択地域
    ''' </summary>
    Private SelectChiiki As String = ""

    ''' <summary>
    ''' 保存計算日時
    ''' </summary>
    Private SaveCalcDateTime As String = ""

    ''' <summary>
    ''' 地域区分を強制変更
    ''' </summary>
    Private IsForceChange As Boolean = False

    ''' <summary>
    ''' 初期表示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA4210F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '個人か否か
            IsKojin = (CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人)

            'DGVちらつき防止
            dgvList.GetType.InvokeMember("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, Nothing, dgvList, New Object() {True})
            '表示位置
            dgvList.Columns(DGV_COL.地域).CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvList.Columns(DGV_COL.母集団).CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvList.Columns(DGV_COL.集計対象数).CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvList.Columns(DGV_COL.営農集計倍率).CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            '調査年プルダウン設定
            Dim ds = New List(Of KeyValuePair(Of String, String))
            ds.Add(New KeyValuePair(Of String, String)("", ""))
            Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                For Each row As DataRow In DAOKobetsuKekkahyo.GetChosaNen(db, Nothing, Nothing, Nothing, ComConst.欠測値補完.無).Rows
                    Dim nen = row.GetString("調査年")
                    If nen.CompareTo("2022") >= 0 Then
                        ds.Add(New KeyValuePair(Of String, String)(nen, nen))
                    End If
                Next
            End Using
            SetCombo(CboChosaNen, ds)

            '営農類型プルダウン設定(営農類型-集落営農)
            ds = New List(Of KeyValuePair(Of String, String))
            ds.Add(New KeyValuePair(Of String, String)("-", ""))
            For Each kv In ComConst.営農類型区分.リスト
                If Not IsKojin AndAlso kv.Key = ComConst.営農類型区分.水田作 Then
                    ds.Add(New KeyValuePair(Of String, String)(kv.Key & "-1", kv.Value & "（集落営農）"))
                    ds.Add(New KeyValuePair(Of String, String)(kv.Key & "-2", kv.Value & "（集落営農以外）"))
                Else
                    ds.Add(New KeyValuePair(Of String, String)(kv.Key & "-0", kv.Value))
                End If
            Next
            SetCombo(CboEinoruikei, ds)

            '地域プルダウン設定
            ds = New List(Of KeyValuePair(Of String, String))
            ds.Add(New KeyValuePair(Of String, String)("", ""))
            ds.Add(New KeyValuePair(Of String, String)("1", "全国"))
            ds.Add(New KeyValuePair(Of String, String)("2", "北海道・都府県"))
            IsForceChange = True
            SetCombo(CboChiiki, ds)
            IsForceChange = False

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' コンボボックスを設定
    ''' </summary>
    ''' <param name="cbo"></param>
    ''' <param name="ds"></param>
    Private Shared Sub SetCombo(cbo As ComboBox, ds As List(Of KeyValuePair(Of String, String)))
        cbo.ValueMember = "Key"
        cbo.DisplayMember = "Value"
        cbo.DataSource = ds
    End Sub

    ''' <summary>
    ''' コンボボックスの値を取得
    ''' </summary>
    ''' <param name="cbo"></param>
    ''' <returns></returns>
    Private Shared Function GetComboValue(cbo As ComboBox) As String
        Return If(cbo.SelectedValue, "").ToString
    End Function

    ''' <summary>
    ''' 調査年プルダウン選択変更
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CboChosaNen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboChosaNen.SelectedIndexChanged
        Try
            Cursor.Current = Cursors.WaitCursor

            Dim nen = GetComboValue(CboChosaNen)
            If nen = ShowChosaNen Then
                Return
            End If

            '未保存確認
            If lblCalcDateTime.Text <> SaveCalcDateTime Then
                If Message.ShowMsgBox(MessageID.MSG_Q_065, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    CboChosaNen.SelectedValue = ShowChosaNen
                    Return
                End If
            End If

            '初期化
            Initialize()
            IsForceChange = True
            CboEinoruikei.SelectedIndex = -1
            CboChiiki.SelectedIndex = -1
            IsForceChange = False

            '表示調査年を退避
            ShowChosaNen = nen

            '営農類型プルダウンの制御
            If String.IsNullOrEmpty(ShowChosaNen) Then
                CboEinoruikei.Enabled = False
                CboChiiki.Enabled = False
                Return
            End If
            CboEinoruikei.Enabled = True

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 初期化
    ''' </summary>
    Private Sub Initialize()
        '変数クリア
        ShowEinoruikei = ""
        ShowShurakueino = ""
        SelectChiiki = ""
        SaveCalcDateTime = ""
        '画面クリア
        lblCalcDateTime.Text = ""
        dgvList.Rows.Clear()
        dgvList.ClearSelection()
        '計算、適用、保存ボタンを非活性化
        BtnCalc.Enabled = False
        BtnApply.Enabled = False
        BtnSave.Enabled = False
    End Sub

    ''' <summary>
    ''' 営農類型プルダウン選択変更
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CboEinoruikei_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboEinoruikei.SelectedIndexChanged
        Try
            Cursor.Current = Cursors.WaitCursor

            Dim einos = GetComboValue(CboEinoruikei).Split("-"c)
            If einos.Length <> 2 Then
                Return
            End If
            If einos(0) = ShowEinoruikei AndAlso einos(1) = ShowShurakueino Then
                Return
            End If

            '未保存確認
            If lblCalcDateTime.Text <> SaveCalcDateTime Then
                If Message.ShowMsgBox(MessageID.MSG_Q_065, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    CboEinoruikei.SelectedValue = ShowEinoruikei & "-" & ShowShurakueino
                    Return
                End If
            End If

            '初期化
            Initialize()

            '表示した区分を退避
            ShowEinoruikei = einos(0)
            ShowShurakueino = einos(1)

            '地域プルダウンの制御
            If String.IsNullOrEmpty(ShowEinoruikei) Then
                IsForceChange = True
                CboChiiki.SelectedIndex = -1
                CboChiiki.Enabled = False
                IsForceChange = False
                Return
            End If
            CboChiiki.Enabled = True

            Dim dt As DataTable
            Dim hasCalcData = False
            Dim calcDateTime = ""
            '集計倍率管理テーブル検索
            dt = SelectEinoBairitsu(ShowChosaNen, ShowEinoruikei, ShowShurakueino)
            If dt.Rows.Count > 0 Then
                hasCalcData = True
                calcDateTime = dt.Rows(0).GetString("計算日時")
                SelectChiiki = dt.Rows(0).GetString("選択地域")
            Else
                '集計倍率管理テーブルがない場合、集計対象数検索
                dt = If(IsKojin, SelectEinoKojinShukei(ShowChosaNen, ShowEinoruikei), SelectEinoHojinShukei(ShowChosaNen, ShowEinoruikei, ShowShurakueino))
            End If

            '一覧に表示
            For Each row As DataRow In dt.Rows
                Dim values = New List(Of String)
                values.Add(row.GetString("営農類型区分"))
                Dim chiiki = row.GetString("地域区分")
                values.Add(Get地域名称(chiiki))
                values.Add(row.GetString("規模区分"))
                If hasCalcData Then
                    values.Add(IntegerDisp(row.GetInteger("母集団")))
                    values.Add(IntegerDisp(row.GetInteger("集計対象数")))
                    values.Add(DecimalDisp(row.GetDecimal("営農集計倍率")))
                Else
                    values.Add("")
                    values.Add("")
                    values.Add("")
                End If
                values.Add(row.GetString("集落営農区分"))
                values.Add(chiiki)
                AddDgvRow(values.ToArray)
            Next

            '地域プルダウンの選択
            Dim before = GetComboValue(CboChiiki)
            IsForceChange = True
            CboChiiki.SelectedValue = SelectChiiki
            If before = SelectChiiki Then
                CboChiiki_SelectedIndexChanged(CboChiiki, Nothing)
            End If
            IsForceChange = False

            '計算日時を表示
            lblCalcDateTime.Text = calcDateTime
            SaveCalcDateTime = lblCalcDateTime.Text

            '計算、適用、保存ボタンを活性化
            BtnCalc.Enabled = True
            BtnApply.Enabled = True
            BtnSave.Enabled = True

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    '営農集計倍率管理検索SQL
    Private Const SQL_SELECT_EINO_BAIRITSU =
        "SELECT" _
        & " 営農類型区分," _
        & " 集落営農区分," _
        & " 地域区分," _
        & " 規模区分," _
        & " 母集団," _
        & " 集計対象数," _
        & " 営農集計倍率," _
        & " FORMAT(計算日時, 'yyyy/MM/dd HH:mm') AS 計算日時," _
        & " 選択地域" _
        & " FROM 営農集計倍率管理" _
        & " WHERE 調査区分 = @調査区分" _
        & " AND 調査年 = @調査年" _
        & " AND 営農類型区分 = @営農類型区分" _
        & " AND 集落営農区分 = @集落営農区分" _
        & " ORDER BY 3,4"
    ''' <summary>
    ''' 営農集計倍率管理検索
    ''' </summary>
    ''' <param name="chosanen"></param>
    ''' <param name="einoruikei"></param>
    ''' <param name="shurakueino"></param>
    ''' <returns></returns>
    Private Shared Function SelectEinoBairitsu(chosanen As String, einoruikei As String, shurakueino As String) As DataTable
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Dim para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
            para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoruikei))
            para.Add(db.CreateParameter("@集落営農区分", SqlDbType.Int, shurakueino))
            Return db.GetDataTable(SQL_SELECT_EINO_BAIRITSU, para)
        End Using
    End Function

    '営農個人集計対象数集計SQL
    Private Const SQL_SELECT_EINO_KOJIN_SHUKEI =
        "SELECT" _
        & " BASE.営農類型区分," _
        & " 0 AS 集落営農区分," _
        & " BASE.地域区分," _
        & " BASE.規模区分," _
        & " SUM(CASE" _
        & " WHEN K.調査年 IS NOT NULL AND BASE.地域区分 = 0  THEN 1" _
        & " WHEN K.調査年 IS NOT NULL AND BASE.地域区分 = 1  AND K.K000002 = 1 THEN 1" _
        & " WHEN K.調査年 IS NOT NULL AND BASE.地域区分 = 2  AND K.K000002 <> 1 THEN 1" _
        & " ELSE 0 END) AS 集計対象数" _
        & " FROM" _
        & " (SELECT @営農類型区分 AS 営農類型区分, W1.地域区分, W2.規模区分" _
        & " FROM" _
        & " (SELECT 0 AS 地域区分 UNION ALL SELECT 1 UNION ALL SELECT 2) W1" _
        & " CROSS JOIN " _
        & " (SELECT TOP(20) COUNT(*) OVER(ORDER BY OBJECT_ID) AS 規模区分 FROM SYS.OBJECTS) W2) BASE" _
        & " LEFT JOIN 個別結果表＿農業経営＿営農類型＿個人１ K" _
        & " ON K.調査年 = @調査年" _
        & " AND K.K000011 = 1" _
        & " AND K.K000027 = 0" _
        & " AND K.K000005 = BASE.営農類型区分" _
        & " AND K.K000006 = BASE.規模区分" _
        & " GROUP BY" _
        & " BASE.営農類型区分," _
        & " BASE.地域区分," _
        & " BASE.規模区分" _
        & " ORDER BY 3,4"
    ''' <summary>
    ''' 営農個人集計対象数集計
    ''' </summary>
    ''' <param name="chosanen"></param>
    ''' <param name="einoruikei"></param>
    ''' <returns></returns>
    Private Shared Function SelectEinoKojinShukei(chosanen As String, einoruikei As String) As DataTable
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Dim para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
            para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoruikei))
            Return db.GetDataTable(SQL_SELECT_EINO_KOJIN_SHUKEI, para)
        End Using
    End Function

    '営農法人集計対象数集計SQL
    Private Const SQL_SELECT_EINO_HOJIN_SHUKEI =
        "SELECT" _
        & " BASE.営農類型区分," _
        & " BASE.集落営農区分," _
        & " BASE.地域区分," _
        & " BASE.規模区分," _
        & " SUM(CASE" _
        & " WHEN K.調査年 IS NOT NULL AND BASE.地域区分 = 0  THEN 1" _
        & " WHEN K.調査年 IS NOT NULL AND BASE.地域区分 = 1  AND K.K000002 = 1 THEN 1" _
        & " WHEN K.調査年 IS NOT NULL AND BASE.地域区分 = 2  AND K.K000002 <> 1 THEN 1" _
        & " ELSE 0 END) AS 集計対象数" _
        & " FROM" _
        & " (SELECT @営農類型区分 AS 営農類型区分, @集落営農区分 AS 集落営農区分, W1.地域区分, W2.規模区分" _
        & " FROM" _
        & " (SELECT 0 AS 地域区分 UNION ALL SELECT 1 UNION ALL SELECT 2) W1" _
        & " CROSS JOIN" _
        & " (SELECT TOP(20) COUNT(*) OVER(ORDER BY OBJECT_ID) AS 規模区分 FROM SYS.OBJECTS) W2) BASE" _
        & " LEFT JOIN 個別結果表＿農業経営＿営農類型＿法人１ K" _
        & " ON K.調査年 = @調査年" _
        & " AND K.K000026 = 1" _
        & " AND K.K000016 = 0" _
        & " AND K.K000006 = BASE.営農類型区分" _
        & " AND K.K000007 = BASE.規模区分" _
        & " AND (BASE.集落営農区分 = 0" _
        & "  OR  BASE.集落営農区分 = 1 AND K.K000005 = 6" _
        & "  OR  BASE.集落営農区分 = 2 AND K.K000005 <> 6)" _
        & " GROUP BY" _
        & " BASE.営農類型区分," _
        & " BASE.集落営農区分," _
        & " BASE.地域区分," _
        & " BASE.規模区分" _
        & " ORDER BY 3,4"
    ''' <summary>
    ''' 営農法人集計対象数集計
    ''' </summary>
    ''' <param name="chosanen"></param>
    ''' <param name="einoruikei"></param>
    ''' <param name="shurakueino"></param>
    ''' <returns></returns>
    Private Shared Function SelectEinoHojinShukei(chosanen As String, einoruikei As String, shurakueino As String) As DataTable
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Dim para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
            para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoruikei))
            para.Add(db.CreateParameter("@集落営農区分", SqlDbType.Int, shurakueino))
            Return db.GetDataTable(SQL_SELECT_EINO_HOJIN_SHUKEI, para)
        End Using
    End Function

    ''' <summary>
    ''' 地域名称取得
    ''' </summary>
    ''' <param name="chiiki"></param>
    ''' <returns></returns>
    Private Shared Function Get地域名称(chiiki As String) As String
        Select Case chiiki
            Case "0"
                Return "全国"
            Case "1"
                Return "北海道"
            Case "2"
                Return "都府県"
        End Select
        Throw New Exception("想定外の地域区分が指定されました。")
    End Function

    ''' <summary>
    ''' 一覧に行を追加
    ''' </summary>
    ''' <param name="values"></param>
    ''' <returns></returns>
    Private Function AddDgvRow(values As String()) As DataGridViewRow
        Dim row = New DataGridViewRow
        row.CreateCells(dgvList)
        row.SetValues(values)
        row.Height = 20
        dgvList.Rows.Add(row)
        Return row
    End Function

    ''' <summary>
    ''' 整数の表示
    ''' </summary>
    ''' <param name="val"></param>
    ''' <returns></returns>
    Private Function IntegerDisp(val As Integer) As String
        If val = 0 Then
            Return ""
        End If
        Return val.ToString()
    End Function

    ''' <summary>
    ''' 小数の表示
    ''' </summary>
    ''' <param name="val"></param>
    ''' <returns></returns>
    Private Function DecimalDisp(val As Decimal) As String
        If val = 0D Then
            Return ""
        End If
        Return val.ToString("0.00")
    End Function

    ''' <summary>
    ''' 地域プルダウン選択変更
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CboChiiki_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboChiiki.SelectedIndexChanged
        Try
            Cursor.Current = Cursors.WaitCursor

            Dim chiiki = GetComboValue(CboChiiki)
            If Not IsForceChange AndAlso chiiki = SelectChiiki Then
                Return
            End If

            '実行確認
            If Not IsForceChange AndAlso Not String.IsNullOrEmpty(SelectChiiki) Then
                If Message.ShowMsgBox(MessageID.MSG_Q_066, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    CboChiiki.SelectedValue = SelectChiiki
                    Return
                End If
            End If

            '選択地域を退避
            SelectChiiki = chiiki

            '計算日時をクリア
            lblCalcDateTime.Text = ""
            SaveCalcDateTime = ""

            If String.IsNullOrEmpty(SelectChiiki) Then
                '空選択（全てをクリアし、非活性化）
                dgvList.Rows.Cast(Of DataGridViewRow).ToList() _
                    .ForEach(Sub(x)
                                 With x.Cells(DGV_COL.母集団)
                                     .Value = ""
                                     .ReadOnly = True
                                     .Style.BackColor = Color.WhiteSmoke
                                 End With
                                 x.Cells(DGV_COL.営農集計倍率).Value = ""
                             End Sub)
            ElseIf SelectChiiki = "1" Then
                '全国（全国を活性化、北海道・都府県をクリアし、非活性化）
                dgvList.Rows.Cast(Of DataGridViewRow).ToList() _
                    .ForEach(Sub(x)
                                 With x.Cells(DGV_COL.母集団)
                                     If x.Cells(DGV_COL.非表示＿地域区分).GetString = "0" Then
                                         .ReadOnly = False
                                         .Style.BackColor = Color.White
                                     Else
                                         .Value = ""
                                         .ReadOnly = True
                                         .Style.BackColor = Color.WhiteSmoke
                                         x.Cells(DGV_COL.営農集計倍率).Value = ""
                                     End If
                                 End With
                             End Sub)
            Else
                '北海道・都府県（北海道・都府県を活性化、全国をクリアし、非活性化）
                dgvList.Rows.Cast(Of DataGridViewRow).ToList() _
                    .ForEach(Sub(x)
                                 With x.Cells(DGV_COL.母集団)
                                     If x.Cells(DGV_COL.非表示＿地域区分).GetString <> "0" Then
                                         .ReadOnly = False
                                         .Style.BackColor = Color.White
                                     Else
                                         .Value = ""
                                         .ReadOnly = True
                                         .Style.BackColor = Color.WhiteSmoke
                                         x.Cells(DGV_COL.営農集計倍率).Value = ""
                                     End If
                                 End With
                             End Sub)
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 集計倍率チェックメッセージヘッダ
    ''' </summary>
    Private Const CheckMessageHeader = "センサス番号" & vbTab & "営農類型" & vbTab & "規模区分" & vbTab & "集計対象区分" & vbTab & "集計除外区分" & vbTab & "集計倍率" & vbTab & "計算結果"
    ''' <summary>
    ''' 営農集計倍率計算ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnCalc_Click(sender As Object, e As EventArgs) Handles BtnCalc.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            '地域選択チェック
            If String.IsNullOrEmpty(SelectChiiki) Then
                Message.ShowMsgBox(MessageID.MSG_E_156, MsgBoxStyle.OkOnly)
                CboChiiki.Focus()
                Return
            End If

            '実行確認
            If Message.ShowMsgBox(MessageID.MSG_Q_067, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If

            '集計対象数検索
            Dim dt = If(IsKojin, SelectEinoKojinShukei(ShowChosaNen, ShowEinoruikei), SelectEinoHojinShukei(ShowChosaNen, ShowEinoruikei, ShowShurakueino))
            '集計対象数のセットと営農集計倍率、案分値のクリア
            For i = 0 To dt.Rows.Count - 1
                With dgvList.Rows(i)
                    .Cells(DGV_COL.集計対象数).Value = IntegerDisp(dt.Rows(i).GetInteger("集計対象数"))
                    .Cells(DGV_COL.営農集計倍率).Value = ""
                End With
            Next

            '集計倍率の算出
            dgvList.Rows.Cast(Of DataGridViewRow) _
                .Where(Function(x) x.Cells(DGV_COL.母集団).GetInteger <> 0 AndAlso x.Cells(DGV_COL.集計対象数).GetInteger <> 0).ToList _
                .ForEach(Sub(x)
                             'REV_001↓偶数丸めになっていたため、修正
                             'Dim bairitsu = Math.Round(CDec(x.Cells(DGV_COL.母集団).GetInteger) / x.Cells(DGV_COL.集計対象数).GetInteger, 2)
                             Dim bairitsu = ComUtil.Round(CDec(x.Cells(DGV_COL.母集団).GetInteger) / x.Cells(DGV_COL.集計対象数).GetInteger, 2)
                             'REV_001↑
                             x.Cells(DGV_COL.営農集計倍率).Value = DecimalDisp(bairitsu)
                         End Sub)

            '集計倍率が計算結果と異なる経営体と集計対象区分=0または集計除外区分=1かつ営農集計倍率<>0の経営体を抽出
            Dim msg1 = New StringBuilder()
            Dim msg2 = New StringBuilder()
            Dim msg3 = New StringBuilder()
            dgvList.Rows.Cast(Of DataGridViewRow).Where(Function(x) Not x.Cells(DGV_COL.母集団).ReadOnly).OrderBy(Function(x) x.Index).ToList _
                .ForEach(Sub(x)
                             If IsKojin Then
                                 '欠測値補完無
                                 For Each row As DataRow In SelectEinoKojinDiff(ShowChosaNen, ShowEinoruikei, x.Cells(DGV_COL.規模区分).GetString, x.Cells(DGV_COL.非表示＿地域区分).GetString, x.Cells(DGV_COL.営農集計倍率).GetDecimal, False).Rows
                                     msg1.Append(row("センサス番号")).Append(vbTab)
                                     msg1.Append(ShowEinoruikei).Append(vbTab)
                                     msg1.Append(x.Cells(DGV_COL.規模区分).GetInteger).Append(vbTab)
                                     msg1.Append(ComUtil.GetBlankOrString(row("集計対象区分"))).Append(vbTab)
                                     msg1.Append(ComUtil.GetBlankOrString(row("集計除外区分"))).Append(vbTab)
                                     msg1.Append(ComUtil.GetBlankOrString(row("集計倍率"))).Append(vbTab)
                                     msg1.Append(x.Cells(DGV_COL.営農集計倍率).GetDecimal).Append(vbTab)
                                     msg1.Append(vbCrLf)
                                 Next
                                 '欠測値補完有
                                 For Each row As DataRow In SelectEinoKojinDiff(ShowChosaNen, ShowEinoruikei, x.Cells(DGV_COL.規模区分).GetString, x.Cells(DGV_COL.非表示＿地域区分).GetString, x.Cells(DGV_COL.営農集計倍率).GetDecimal, True).Rows
                                     msg2.Append(row("センサス番号")).Append(vbTab)
                                     msg2.Append(ShowEinoruikei).Append(vbTab)
                                     msg2.Append(x.Cells(DGV_COL.規模区分).GetInteger).Append(vbTab)
                                     msg2.Append(ComUtil.GetBlankOrString(row("集計対象区分"))).Append(vbTab)
                                     msg2.Append(ComUtil.GetBlankOrString(row("集計除外区分"))).Append(vbTab)
                                     msg2.Append(ComUtil.GetBlankOrString(row("集計倍率"))).Append(vbTab)
                                     msg2.Append(x.Cells(DGV_COL.営農集計倍率).GetDecimal).Append(vbTab)
                                     msg2.Append(vbCrLf)
                                 Next
                             Else
                                 '法人
                                 For Each row As DataRow In SelectEinoHojinDiff(ShowChosaNen, ShowEinoruikei, ShowShurakueino, x.Cells(DGV_COL.規模区分).GetString, x.Cells(DGV_COL.非表示＿地域区分).GetString, x.Cells(DGV_COL.営農集計倍率).GetDecimal).Rows
                                     msg3.Append(row("センサス番号")).Append(vbTab)
                                     msg3.Append(ShowEinoruikei).Append(vbTab)
                                     msg3.Append(x.Cells(DGV_COL.規模区分).GetInteger).Append(vbTab)
                                     msg3.Append(ComUtil.GetBlankOrString(row("集計対象区分"))).Append(vbTab)
                                     msg3.Append(ComUtil.GetBlankOrString(row("集計除外区分"))).Append(vbTab)
                                     msg3.Append(ComUtil.GetBlankOrString(row("集計倍率"))).Append(vbTab)
                                     msg3.Append(x.Cells(DGV_COL.営農集計倍率).GetDecimal).Append(vbTab)
                                     msg3.Append(vbCrLf)
                                 Next
                             End If
                         End Sub)
            'メッセージを集約
            Dim msg = New StringBuilder()
            If msg1.Length > 0 Then
                msg.Append("欠測値補完無").Append(vbCrLf)
                msg.Append(CheckMessageHeader).Append(vbCrLf)
                msg.Append(msg1)
            End If
            If msg2.Length > 0 Then
                msg.Append("欠測値補完有").Append(vbCrLf)
                msg.Append(CheckMessageHeader).Append(vbCrLf)
                msg.Append(msg2)
            End If
            If msg3.Length > 0 Then
                msg.Append(CheckMessageHeader).Append(vbCrLf)
                msg.Append(msg3)
            End If
            If msg.Length > 0 Then
                Message.ShowMsgForm(Me, MessageID.MSG_E_157, {msg.ToString})
            End If

            '計算日時を表示
            lblCalcDateTime.Text = Now.ToString("yyyy/MM/dd HH:mm")
            '保存計算日時をクリア
            SaveCalcDateTime = ""

            '完了メッセージ
            Message.ShowMsgBox(MessageID.MSG_I_054, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    '営農個人集計倍率差分検索SQL
    Private Const SQL_SELECT_EINO_KOJIN_DIFF =
        "SELECT" _
        & " センサス番号," _
        & " K000011 AS 集計対象区分," _
        & " K000027 AS 集計除外区分," _
        & " K000012 AS 集計倍率" _
        & " FROM 個別結果表＿農業経営＿営農類型＿個人１{0}" _
        & " WHERE 調査年 = @調査年" _
        & " AND K000011 = 1" _
        & " AND K000027 = 0" _
        & " AND K000005 = @営農類型区分" _
        & " AND K000006 = @規模区分" _
        & " {1}" _
        & " AND ISNULL(K000012,0) <> @計算結果" _
        & " AND @計算結果 <> 0" _
        & " UNION ALL" _
        & " SELECT" _
        & " センサス番号," _
        & " K000011," _
        & " K000027," _
        & " K000012" _
        & " FROM 個別結果表＿農業経営＿営農類型＿個人１{0}" _
        & " WHERE 調査年 = @調査年" _
        & " AND (ISNULL(K000011,0) = 0 OR ISNULL(K000027,1) = 1)" _
        & " AND K000005 = @営農類型区分" _
        & " AND K000006 = @規模区分" _
        & " {1}" _
        & " AND ISNULL(K000012,1) <> 0" _
        & " ORDER BY 1"
    ''' <summary>
    ''' 営農個人集計倍率差分検索
    ''' </summary>
    ''' <param name="chosanen"></param>
    ''' <param name="einoruikei"></param>
    ''' <param name="kibo"></param>
    ''' <param name="chiiki"></param>
    ''' <param name="bairitsu"></param>
    ''' <param name="isKessoku"></param>
    ''' <returns></returns>
    Private Shared Function SelectEinoKojinDiff(chosanen As String, einoruikei As String, kibo As String, chiiki As String, bairitsu As Decimal, isKessoku As Boolean) As DataTable
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Dim sql = String.Format(SQL_SELECT_EINO_KOJIN_DIFF,
                                    If(isKessoku, "＿集計用", ""),
                                    If(chiiki = "0", "", If(chiiki = "1", "AND K000002 = 1", "AND K000002 <> 1")))
            Dim para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
            para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoruikei))
            para.Add(db.CreateParameter("@規模区分", SqlDbType.Int, kibo))
            para.Add(db.CreateParameter("@計算結果", SqlDbType.Decimal, bairitsu))
            Return db.GetDataTable(sql, para)
        End Using
    End Function

    '営農法人集計倍率差分検索SQL
    Private Const SQL_SELECT_EINO_HOJIN_DIFF =
        "SELECT" _
        & " センサス番号," _
        & " K000026 AS 集計対象区分," _
        & " K000016 AS 集計除外区分," _
        & " K000015 AS 集計倍率" _
        & " FROM 個別結果表＿農業経営＿営農類型＿法人１" _
        & " WHERE 調査年 = @調査年" _
        & " AND K000026 = 1" _
        & " AND K000016 = 0" _
        & " AND K000006 = @営農類型区分" _
        & " {0}" _
        & " AND K000007 = @規模区分" _
        & " {1}" _
        & " AND ISNULL(K000015,0) <> @計算結果" _
        & " AND @計算結果 <> 0" _
        & " UNION ALL" _
        & " SELECT" _
        & " センサス番号," _
        & " K000026," _
        & " K000016," _
        & " K000015" _
        & " FROM 個別結果表＿農業経営＿営農類型＿法人１" _
        & " WHERE 調査年 = @調査年" _
        & " AND (ISNULL(K000026,0) = 0 OR ISNULL(K000016,1) = 1)" _
        & " AND K000006 = @営農類型区分" _
        & " {0}" _
        & " AND K000007 = @規模区分" _
        & " {1}" _
        & " AND ISNULL(K000015,1) <> 0" _
        & " ORDER BY 1"
    ''' <summary>
    ''' 営農法人集計倍率差分検索
    ''' </summary>
    ''' <param name="chosanen"></param>
    ''' <param name="einoruikei"></param>
    ''' <param name="shurakueino"></param>
    ''' <param name="kibo"></param>
    ''' <param name="chiiki"></param>
    ''' <param name="bairitsu"></param>
    ''' <returns></returns>
    Private Shared Function SelectEinoHojinDiff(chosanen As String, einoruikei As String, shurakueino As String, kibo As String, chiiki As String, bairitsu As Decimal) As DataTable
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Dim sql = String.Format(SQL_SELECT_EINO_HOJIN_DIFF,
                                    If(shurakueino = "0", "", If(shurakueino = "1", "AND K000005 = 6", "AND K000005 <> 6")),
                                    If(chiiki = "0", "", If(chiiki = "1", "AND K000002 = 1", "AND K000002 <> 1")))
            Dim para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
            para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoruikei))
            para.Add(db.CreateParameter("@規模区分", SqlDbType.Int, kibo))
            para.Add(db.CreateParameter("@計算結果", SqlDbType.Decimal, bairitsu))
            Return db.GetDataTable(sql, para)
        End Using
    End Function

    ''' <summary>
    ''' 営農集計倍率適用ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnApply_Click(sender As Object, e As EventArgs) Handles BtnApply.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            '入力チェック
            '計算済チェック
            If String.IsNullOrEmpty(lblCalcDateTime.Text) Then
                Message.ShowMsgBox(MessageID.MSG_E_158, MsgBoxStyle.OkOnly)
                BtnCalc.Focus()
                Return
            End If

            '営農集計倍率桁数チェック
            If Not CheckBairitsu() Then
                Return
            End If

            '実行確認
            If Message.ShowMsgBox(MessageID.MSG_Q_068, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If

            Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try
                    db.BeginTrans()

                    '集計倍率が計算結果と異なる経営体の営農集計倍率を計算結果で更新
                    dgvList.Rows.Cast(Of DataGridViewRow).Where(Function(x) x.Cells(DGV_COL.営農集計倍率).GetDecimal > 0).OrderBy(Function(x) x.Index).ToList _
                        .ForEach(Sub(x)
                                     If IsKojin Then
                                         '個人欠測値補完無
                                         UpdateEinoKojinBairitsu(db, ShowChosaNen, ShowEinoruikei, x.Cells(DGV_COL.規模区分).GetString, x.Cells(DGV_COL.非表示＿地域区分).GetString, x.Cells(DGV_COL.営農集計倍率).GetDecimal, False)
                                         '個人欠測値補完有
                                         UpdateEinoKojinBairitsu(db, ShowChosaNen, ShowEinoruikei, x.Cells(DGV_COL.規模区分).GetString, x.Cells(DGV_COL.非表示＿地域区分).GetString, x.Cells(DGV_COL.営農集計倍率).GetDecimal, True)
                                     Else
                                         '法人
                                         UpdateEinoHojinBairitsu(db, ShowChosaNen, ShowEinoruikei, ShowShurakueino, x.Cells(DGV_COL.規模区分).GetString, x.Cells(DGV_COL.非表示＿地域区分).GetString, x.Cells(DGV_COL.営農集計倍率).GetDecimal)
                                     End If
                                 End Sub)

                    '集計対象区分=0または集計除外区分=1かつ営農集計倍率≠0の経営体の営農集計倍率を0で更新
                    '個人欠測値補完無
                    UpdateEinoKojinBairitsu0(db, ShowChosaNen, ShowEinoruikei, False)
                    '個人欠測値補完有
                    UpdateEinoKojinBairitsu0(db, ShowChosaNen, ShowEinoruikei, True)
                    '法人
                    UpdateEinoHojinBairitsu0(db, ShowChosaNen, ShowEinoruikei, ShowShurakueino)

                    db.CommitTrans()
                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using

            '完了メッセージ
            Message.ShowMsgBox(MessageID.MSG_I_055, MsgBoxStyle.OkOnly)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 集計倍率の桁数チェック
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <returns></returns>
    Private Function CheckBairitsu() As Boolean
        '整数部が5桁を超える場合、エラー
        Dim errorRow = dgvList.Rows.Cast(Of DataGridViewRow).Where(Function(x) CInt(x.Cells(DGV_COL.営農集計倍率).GetDecimal).ToString.Length > 5).FirstOrDefault
        If errorRow Is Nothing Then
            Return True
        End If
        Message.ShowMsgBox(MessageID.MSG_E_159, MsgBoxStyle.OkOnly)
        dgvList.CurrentCell = dgvList.Rows(errorRow.Index).Cells(DGV_COL.母集団)
        dgvList.BeginEdit(True)
        Return False
    End Function

    '営農個人営農集計倍率更新SQL
    Private Const SQL_UPDATE_EINO_KOJIN_BAIRITSU =
        "UPDATE 個別結果表＿農業経営＿営農類型＿個人１{0} SET" _
        & " K000012 = @計算結果," _
        & " 更新日付 = SYSDATETIME()," _
        & " 更新者ID = @更新者ID" _
        & " WHERE 調査年 = @調査年" _
        & " AND K000011 = 1" _
        & " AND K000027 = 0" _
        & " AND K000005 = @営農類型区分" _
        & " AND K000006 = @規模区分" _
        & " {1}" _
        & " AND ISNULL(K000012,0) <> @計算結果"
    ''' <summary>
    ''' 営農個人営農集計倍率更新
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosanen"></param>
    ''' <param name="einoruikei"></param>
    ''' <param name="kibo"></param>
    ''' <param name="chiiki"></param>
    ''' <param name="bairitsu"></param>
    ''' <param name="isKessoku"></param>
    ''' <returns></returns>
    Private Shared Function UpdateEinoKojinBairitsu(db As DBAccess, chosanen As String, einoruikei As String, kibo As String, chiiki As String, bairitsu As Decimal, isKessoku As Boolean) As Integer
        Dim sql = String.Format(SQL_UPDATE_EINO_KOJIN_BAIRITSU,
                                    If(isKessoku, "＿集計用", ""),
                                    If(chiiki = "0", "", If(chiiki = "1", "AND K000002 = 1", "AND K000002 <> 1")))
        Dim para = New List(Of DBAccess.Parameter)
        para.Add(db.CreateParameter("@計算結果", SqlDbType.Decimal, bairitsu))
        para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))
        para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
        para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoruikei))
        para.Add(db.CreateParameter("@規模区分", SqlDbType.Int, kibo))
        Return db.ExecuteNonQuery(sql, para)
    End Function

    '営農法人営農集計倍率更新SQL
    Private Const SQL_UPDATE_EINO_HOJIN_BAIRITSU =
        "UPDATE 個別結果表＿農業経営＿営農類型＿法人１ SET" _
        & " K000015 = @計算結果," _
        & " 更新日付 = SYSDATETIME()," _
        & " 更新者ID = @更新者ID" _
        & " WHERE 調査年 = @調査年" _
        & " AND K000026 = 1" _
        & " AND K000016 = 0" _
        & " AND K000006 = @営農類型区分" _
        & " {0}" _
        & " AND K000007 = @規模区分" _
        & " {1}" _
        & " AND ISNULL(K000015,0) <> @計算結果"

    ''' <summary>
    ''' 営農法人営農集計倍率更新
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosanen"></param>
    ''' <param name="einoruikei"></param>
    ''' <param name="shurakueino"></param>
    ''' <param name="kibo"></param>
    ''' <param name="chiiki"></param>
    ''' <param name="bairitsu"></param>
    ''' <returns></returns>
    Private Shared Function UpdateEinoHojinBairitsu(db As DBAccess, chosanen As String, einoruikei As String, shurakueino As String, kibo As String, chiiki As String, bairitsu As Decimal) As Integer
        Dim sql = String.Format(SQL_UPDATE_EINO_HOJIN_BAIRITSU,
                                If(shurakueino = "0", "", If(shurakueino = "1", "AND K000005 = 6", "AND K000005 <> 6")),
                                If(chiiki = "0", "", If(chiiki = "1", "AND K000002 = 1", "AND K000002 <> 1")))
        Dim para = New List(Of DBAccess.Parameter)
        para.Add(db.CreateParameter("@計算結果", SqlDbType.Decimal, bairitsu))
        para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))
        para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
        para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoruikei))
        para.Add(db.CreateParameter("@規模区分", SqlDbType.Int, kibo))
        Return db.ExecuteNonQuery(sql, para)
    End Function

    '営農個人営農集計倍率0更新SQL
    Private Const SQL_UPDATE_EINO_KOJIN_BAIRITSU_0 =
        "UPDATE 個別結果表＿農業経営＿営農類型＿個人１{0} SET" _
        & " K000012 = 0," _
        & " 更新日付 = SYSDATETIME()," _
        & " 更新者ID = @更新者ID" _
        & " WHERE 調査年 = @調査年" _
        & " AND (ISNULL(K000011,0) = 0 OR ISNULL(K000027,1) = 1)" _
        & " AND K000005 = @営農類型区分" _
        & " AND ISNULL(K000012,1) <> 0"
    ''' <summary>
    ''' 営農個人営農集計倍率0更新
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosanen"></param>
    ''' <param name="einoruikei"></param>
    ''' <param name="isKessoku"></param>
    ''' <returns></returns>
    Private Shared Function UpdateEinoKojinBairitsu0(db As DBAccess, chosanen As String, einoruikei As String, isKessoku As Boolean) As Integer
        Dim sql = String.Format(SQL_UPDATE_EINO_KOJIN_BAIRITSU_0, If(isKessoku, "＿集計用", ""))
        Dim para = New List(Of DBAccess.Parameter)
        para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))
        para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
        para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoruikei))
        Return db.ExecuteNonQuery(sql, para)
    End Function

    '営農法人営農集計倍率0更新SQL
    Private Const SQL_UPDATE_EINO_HOJIN_BAIRITSU_0 =
        "UPDATE 個別結果表＿農業経営＿営農類型＿法人１ SET" _
        & " K000015 = 0," _
        & " 更新日付 = SYSDATETIME()," _
        & " 更新者ID = @更新者ID" _
        & " WHERE 調査年 = @調査年" _
        & " AND (ISNULL(K000026,0) = 0 OR ISNULL(K000016,1) = 1)" _
        & " AND K000006 = @営農類型区分" _
        & " {0}" _
        & " AND ISNULL(K000015,1) <> 0"
    ''' <summary>
    ''' 営農法人営農集計倍率0更新
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosanen"></param>
    ''' <param name="einoruikei"></param>
    ''' <param name="shurakueino"></param>
    ''' <returns></returns>
    Private Shared Function UpdateEinoHojinBairitsu0(db As DBAccess, chosanen As String, einoruikei As String, shurakueino As String) As Integer
        Dim sql = String.Format(SQL_UPDATE_EINO_HOJIN_BAIRITSU_0, If(shurakueino = "0", "", If(shurakueino = "1", "AND K000005 = 6", "AND K000005 <> 6")))
        Dim para = New List(Of DBAccess.Parameter)
        para.Add(db.CreateParameter("@更新者ID", SqlDbType.VarChar, CommonInfo.UserId))
        para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
        para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoruikei))
        Return db.ExecuteNonQuery(sql, para)
    End Function

    ''' <summary>
    ''' 保存ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click

        Try
            Cursor.Current = Cursors.WaitCursor

            '入力チェック
            '営農集計倍率桁数チェック
            If Not CheckBairitsu() Then
                Return
            End If

            '実行確認
            If Message.ShowMsgBox(MessageID.MSG_Q_001, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If

            '保存処理
            Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                Try
                    db.BeginTrans()

                    '既存削除
                    DeleteEinoBairitsu(db, ShowChosaNen, ShowEinoruikei, ShowShurakueino)

                    '登録
                    dgvList.Rows.Cast(Of DataGridViewRow).ToList() _
                        .ForEach(Sub(x)
                                     InsertEinoBairitsu(db, ShowChosaNen, ShowEinoruikei, ShowShurakueino,
                                                        x.Cells(DGV_COL.非表示＿地域区分).GetString,
                                                        x.Cells(DGV_COL.規模区分).GetString,
                                                        x.Cells(DGV_COL.母集団).GetInteger,
                                                        x.Cells(DGV_COL.集計対象数).GetInteger,
                                                        x.Cells(DGV_COL.営農集計倍率).GetDecimal,
                                                        lblCalcDateTime.Text,
                                                        SelectChiiki)
                                 End Sub)

                    db.CommitTrans()
                Catch ex As Exception
                    db.RollBackTrans()
                    Throw ex
                End Try
            End Using

            '保存計算日時を更新
            SaveCalcDateTime = lblCalcDateTime.Text

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

    '営農集計倍率管理削除SQL
    Private Const SQL_DELETE_EINO_BAIRITSU =
        "DELETE FROM 営農集計倍率管理" _
        & " WHERE 調査区分 = @調査区分" _
        & " AND 調査年 = @調査年" _
        & " AND 営農類型区分 = @営農類型区分" _
        & " AND 集落営農区分 = @集落営農区分"
    ''' <summary>
    ''' 営農集計倍率管理削除
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosanen"></param>
    ''' <param name="einoruikei"></param>
    ''' <param name="shurakueino"></param>
    ''' <returns></returns>
    Private Shared Function DeleteEinoBairitsu(db As DBAccess, chosanen As String, einoruikei As String, shurakueino As String) As Integer
        Dim para = New List(Of DBAccess.Parameter)
        para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
        para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
        para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoruikei))
        para.Add(db.CreateParameter("@集落営農区分", SqlDbType.Int, shurakueino))
        Return db.ExecuteNonQuery(SQL_DELETE_EINO_BAIRITSU, para)
    End Function

    '営農集計倍率管理登録SQL
    Private Const SQL_INSERT_EINO_BAIRITSU =
        "INSERT INTO 営農集計倍率管理 VALUES(" _
        & " @調査区分," _
        & " @調査年," _
        & " @営農類型区分," _
        & " @集落営農区分," _
        & " @地域区分," _
        & " @規模区分," _
        & " @母集団," _
        & " @集計対象数," _
        & " @営農集計倍率," _
        & " @計算日時," _
        & " @選択地域" _
        & ")"
    ''' <summary>
    ''' 営農集計倍率管理登録
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="chosanen"></param>
    ''' <param name="einoruikei"></param>
    ''' <param name="shurakueino"></param>
    ''' <param name="chiiki"></param>
    ''' <param name="kibo"></param>
    ''' <param name="boshudan"></param>
    ''' <param name="shukei"></param>
    ''' <param name="bairitsu"></param>
    ''' <param name="calcDateTime"></param>
    ''' <param name="selectChiiki"></param>
    ''' <returns></returns>
    Private Shared Function InsertEinoBairitsu(db As DBAccess, chosanen As String, einoruikei As String, shurakueino As String, chiiki As String, kibo As String, boshudan As Integer, shukei As Integer, bairitsu As Decimal, calcDateTime As String, selectChiiki As String) As Integer
        Dim para = New List(Of DBAccess.Parameter)
        para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, CommonInfo.Chosakubun))
        para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
        para.Add(db.CreateParameter("@営農類型区分", SqlDbType.Int, einoruikei))
        para.Add(db.CreateParameter("@集落営農区分", SqlDbType.Int, shurakueino))
        para.Add(db.CreateParameter("@地域区分", SqlDbType.Int, chiiki))
        para.Add(db.CreateParameter("@規模区分", SqlDbType.Int, kibo))
        para.Add(db.CreateParameter("@母集団", SqlDbType.Decimal, boshudan))
        para.Add(db.CreateParameter("@集計対象数", SqlDbType.Decimal, shukei))
        para.Add(db.CreateParameter("@営農集計倍率", SqlDbType.Decimal, bairitsu))
        Dim dt As DateTime
        If DateTime.TryParse(calcDateTime, dt) Then
            para.Add(db.CreateParameter("@計算日時", SqlDbType.DateTime, dt))
        Else
            para.Add(db.CreateParameter("@計算日時", SqlDbType.DateTime, Nothing))
        End If
        para.Add(db.CreateParameter("@選択地域", SqlDbType.Int, selectChiiki))
        Return db.ExecuteNonQuery(SQL_INSERT_EINO_BAIRITSU, para)
    End Function

    ''' <summary>
    ''' 戻るボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Overrides Sub BtnReturn_Click(sender As Object, e As EventArgs)
        If lblCalcDateTime.Text <> SaveCalcDateTime Then
            If Message.ShowMsgBox(MessageID.MSG_Q_065, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            End If
        End If
        MyBase.BtnReturn_Click(sender, e)
    End Sub

    ''' <summary>
    ''' 一覧キーダウン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DgvList_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvList.KeyDown
        Try
            '入力可能な母集団列のみ選択されている場合のみ処理
            Dim left = dgvList.SelectedCells.Cast(Of DataGridViewCell).Select(Function(x) x.ColumnIndex).Min
            Dim right = dgvList.SelectedCells.Cast(Of DataGridViewCell).Select(Function(x) x.ColumnIndex).Max
            If left <> DGV_COL.母集団 OrElse right <> DGV_COL.母集団 Then
                Return
            End If

            'Deleteキー、BackSpaceキー
            If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
                '読取専用は除外
                dgvList.SelectedCells.Cast(Of DataGridViewCell).Where(Function(x) Not x.ReadOnly).ToList.ForEach(Sub(x) x.Value = "")
                Return
            End If

            'クリップボードからのペースト
            If e.Control AndAlso e.KeyCode = Keys.V Then
                'ペースト開始行
                Dim top = dgvList.SelectedCells.Cast(Of DataGridViewCell).Select(Function(x) x.RowIndex).Min
                If dgvList.Rows(top).Cells(DGV_COL.母集団).ReadOnly Then
                    '読取専用の場合は何もしない
                    Return
                End If
                Dim texts = Clipboard.GetText?.Split({vbLf}, StringSplitOptions.None)
                If texts Is Nothing Then
                    Return
                End If
                For i = 0 To texts.Length - 1
                    If top + i >= dgvList.Rows.Count Then
                        '一覧の最終行まででペースト終了
                        Exit For
                    End If
                    With dgvList.Rows(top + i).Cells(DGV_COL.母集団)
                        If .ReadOnly Then
                            '読取専用になったらペースト終了
                            Exit For
                        End If
                        Dim val As Integer
                        If Not Integer.TryParse(texts(i).Replace(vbTab, "").Trim, val) OrElse val <= 0 Then
                            val = 0
                        End If
                        .Value = IntegerDisp(val)
                    End With
                Next
                Return
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
            If e.ColumnIndex <> DGV_COL.母集団 Then
                '母集団列以外は何もしない
                Return
            End If

            Dim cell = DirectCast(dgvList(e.ColumnIndex, e.RowIndex), DataGridViewTextBoxCell)
            If cell.ReadOnly Then
                '読取専用は何もしない
                Return
            End If
            'Nothingは空文字にする
            If cell.Value Is Nothing Then
                cell.Value = ""
            End If

            '0または正の整数以外の場合、クリア
            Dim i As Integer
            If Not Integer.TryParse(cell.GetString, i) OrElse i <= 0 Then
                cell.Value = ""
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
End Class

