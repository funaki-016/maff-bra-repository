Imports System.Text.RegularExpressions

''' <summary>
''' 労賃単価算出画面
''' </summary>
''' <remarks></remarks>
Public Class BRA8220F

    ''' <summary>生産費</summary>
    Private _seisanhi As String
    ''' <summary>調査年</summary>
    Private _chosaNen As String
    ''' <summary>事務所</summary>
    Private _jimusho As String

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="seisanhi"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="jimusho"></param>
    ''' <remarks></remarks>
    Public Sub New(seisanhi As String, chosaNen As String, jimusho As String)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

        _seisanhi = seisanhi
        _chosaNen = chosaNen
        _jimusho = jimusho

    End Sub

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA8220F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            txtSeisanhi.Text = ComConst.生産費区分.リスト(_seisanhi)

            txtChosaNen.Text = _chosaNen

            txtTodofuken.Text = MasterDao.GetJimusyoName(_jimusho)

            'データを設定する
            Me.SetData()

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_001, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                'シートデータ取得
                Dim dc As Dictionary(Of String, String) = Me.GetFormData()
                
                'エラーチェック１
                Dim msgId As String = String.Empty
                If Not Me.CheckError1(dc, msgId) Then
                    'エラーメッセージ
                    Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                'エラーチェック２
                Dim details As New List(Of String)
                If Not Me.CheckError2(dc, details) Then
                    'エラーメッセージ
                    Message.ShowMsgForm(Me, MessageID.MSG_E_010, {String.Join(vbCrLf, details)})
                    Exit Sub
                End If

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    Try

                        db.BeginTrans()

                        '労賃単価削除
                        DAOOther.DeleteRouchinTanka(db, _seisanhi, _chosaNen, _jimusho)

                        '労賃単価追加
                        DAOOther.InsertRouchinTanka(db, _seisanhi, _chosaNen, _jimusho, dc)

                        db.CommitTrans()

                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try
                End Using

                '完了メッセージ
                Message.ShowMsgBox(MessageID.MSG_I_001, MsgBoxStyle.OkOnly)
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
    ''' データを設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetData()

        Dim dtThis As DataTable
        Dim dtLast As DataTable
        Dim dtRouchin As DataTable

        '対象期間取得（本年）
        Dim this As Dictionary(Of String, String) = ComUtil.RouchinTanka.GetTargetTerm(_seisanhi, _chosaNen)
        '対象期間取得（前年）
        Dim last As Dictionary(Of String, String) = ComUtil.RouchinTanka.GetTargetTerm(_seisanhi, (Integer.Parse(_chosaNen) - 1).ToString)

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '毎月勤労統計取得（本年）
            dtThis = DAOOther.GetMaitsukiKinrouToukei(db, _jimusho, this(ComConst.労賃単価.TERM_LOWER), this(ComConst.労賃単価.TERM_UPPER))
            '毎月勤労統計取得（前年）
            dtLast = DAOOther.GetMaitsukiKinrouToukei(db, _jimusho, last(ComConst.労賃単価.TERM_LOWER), last(ComConst.労賃単価.TERM_UPPER))
            '労賃単価取得
            dtRouchin = DAOOther.GetRouchinTanka(db, _seisanhi, _chosaNen, _jimusho)
        End Using

        'フォームデータ設定
        Me.SetFormData(dtThis, dtLast, dtRouchin)
    End Sub

    ''' <summary>
    ''' Decimal型変換
    ''' </summary>
    ''' <param name="val"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConvDecimal(val As Object) As Decimal
        Return If(IsDBNull(val), 0D, Convert.ToDecimal(val))
    End Function

    ''' <summary>
    ''' フォームデータ設定
    ''' </summary>
    ''' <param name="this"></param>
    ''' <param name="last"></param>
    ''' <param name="rouchin"></param>
    ''' <remarks></remarks>
    Private Sub SetFormData(this As DataTable, last As DataTable, rouchin As DataTable)
        Dim numerator As Decimal
        Dim denominator As Decimal

        '優先順位1
        '５～29人／本年値／男（本年）
        numerator = Aggregate dr In this Into _
                      Sum(Me.ConvDecimal(dr("男＿製造業＿現金給与総額")) * Me.ConvDecimal(dr("男＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿建設業＿現金給与総額")) * Me.ConvDecimal(dr("男＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿運輸業郵便業＿現金給与総額")) * Me.ConvDecimal(dr("男＿運輸業郵便業＿常用労働者数")))
        denominator = Aggregate dr In this Into _
                      Sum(Me.ConvDecimal(dr("男＿製造業＿労働時間")) * Me.ConvDecimal(dr("男＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿建設業＿労働時間")) * Me.ConvDecimal(dr("男＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿運輸業郵便業＿労働時間")) * Me.ConvDecimal(dr("男＿運輸業郵便業＿常用労働者数")))

        Dim M_5_29_Hon As Decimal
        If denominator = 0 Then
            M_5_29_Hon = 0
        Else
            M_5_29_Hon = numerator / denominator
        End If
        Me.txtM_5_29_Hon.Text = M_5_29_Hon.ToString("0")

        '５～29人／本年値／男（前年）
        numerator = Aggregate dr In last Into _
                      Sum(Me.ConvDecimal(dr("男＿製造業＿現金給与総額")) * Me.ConvDecimal(dr("男＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿建設業＿現金給与総額")) * Me.ConvDecimal(dr("男＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿運輸業郵便業＿現金給与総額")) * Me.ConvDecimal(dr("男＿運輸業郵便業＿常用労働者数")))
        denominator = Aggregate dr In last Into _
                      Sum(Me.ConvDecimal(dr("男＿製造業＿労働時間")) * Me.ConvDecimal(dr("男＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿建設業＿労働時間")) * Me.ConvDecimal(dr("男＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿運輸業郵便業＿労働時間")) * Me.ConvDecimal(dr("男＿運輸業郵便業＿常用労働者数")))

        Dim M_5_29_Zen As Decimal
        If denominator = 0 Then
            M_5_29_Zen = 0
        Else
            M_5_29_Zen = numerator / denominator
        End If
        Me.txtM_5_29_Zen.Text = M_5_29_Zen.ToString("0")

        '５～29人／本年値／女（本年）
        numerator = Aggregate dr In this Into _
                      Sum(Me.ConvDecimal(dr("女＿製造業＿現金給与総額")) * Me.ConvDecimal(dr("女＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿建設業＿現金給与総額")) * Me.ConvDecimal(dr("女＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿運輸業郵便業＿現金給与総額")) * Me.ConvDecimal(dr("女＿運輸業郵便業＿常用労働者数")))
        denominator = Aggregate dr In this Into _
                      Sum(Me.ConvDecimal(dr("女＿製造業＿労働時間")) * Me.ConvDecimal(dr("女＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿建設業＿労働時間")) * Me.ConvDecimal(dr("女＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿運輸業郵便業＿労働時間")) * Me.ConvDecimal(dr("女＿運輸業郵便業＿常用労働者数")))

        Dim W_5_29_Hon As Decimal
        If denominator = 0 Then
            W_5_29_Hon = 0
        Else
            W_5_29_Hon = numerator / denominator
        End If
        Me.txtW_5_29_Hon.Text = W_5_29_Hon.ToString("0")

        '５～29人／本年値／女（前年）
        numerator = Aggregate dr In last Into _
                      Sum(Me.ConvDecimal(dr("女＿製造業＿現金給与総額")) * Me.ConvDecimal(dr("女＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿建設業＿現金給与総額")) * Me.ConvDecimal(dr("女＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿運輸業郵便業＿現金給与総額")) * Me.ConvDecimal(dr("女＿運輸業郵便業＿常用労働者数")))
        denominator = Aggregate dr In last Into _
                      Sum(Me.ConvDecimal(dr("女＿製造業＿労働時間")) * Me.ConvDecimal(dr("女＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿建設業＿労働時間")) * Me.ConvDecimal(dr("女＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿運輸業郵便業＿労働時間")) * Me.ConvDecimal(dr("女＿運輸業郵便業＿常用労働者数")))

        Dim W_5_29_Zen As Decimal
        If denominator = 0 Then
            W_5_29_Zen = 0
        Else
            W_5_29_Zen = numerator / denominator
        End If
        Me.txtW_5_29_Zen.Text = W_5_29_Zen.ToString("0")

        '５～29人／本年値／男女平均（本年）
        numerator = Aggregate dr In this Into _
                      Sum(Me.ConvDecimal(dr("男＿製造業＿現金給与総額")) * Me.ConvDecimal(dr("男＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿建設業＿現金給与総額")) * Me.ConvDecimal(dr("男＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿運輸業郵便業＿現金給与総額")) * Me.ConvDecimal(dr("男＿運輸業郵便業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿製造業＿現金給与総額")) * Me.ConvDecimal(dr("女＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿建設業＿現金給与総額")) * Me.ConvDecimal(dr("女＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿運輸業郵便業＿現金給与総額")) * Me.ConvDecimal(dr("女＿運輸業郵便業＿常用労働者数")))
        denominator = Aggregate dr In this Into _
                      Sum(Me.ConvDecimal(dr("男＿製造業＿労働時間")) * Me.ConvDecimal(dr("男＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿建設業＿労働時間")) * Me.ConvDecimal(dr("男＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿運輸業郵便業＿労働時間")) * Me.ConvDecimal(dr("男＿運輸業郵便業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿製造業＿労働時間")) * Me.ConvDecimal(dr("女＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿建設業＿労働時間")) * Me.ConvDecimal(dr("女＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿運輸業郵便業＿労働時間")) * Me.ConvDecimal(dr("女＿運輸業郵便業＿常用労働者数")))

        Dim Avg_5_29_Hon As Decimal
        If denominator = 0 Then
            Avg_5_29_Hon = 0
        Else
            Avg_5_29_Hon = numerator / denominator
        End If
        Me.txtAvg_5_29_Hon.Text = Avg_5_29_Hon.ToString("0")

        '５～29人／本年値／男女平均（前年）
        numerator = Aggregate dr In last Into _
                      Sum(Me.ConvDecimal(dr("男＿製造業＿現金給与総額")) * Me.ConvDecimal(dr("男＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿建設業＿現金給与総額")) * Me.ConvDecimal(dr("男＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿運輸業郵便業＿現金給与総額")) * Me.ConvDecimal(dr("男＿運輸業郵便業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿製造業＿現金給与総額")) * Me.ConvDecimal(dr("女＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿建設業＿現金給与総額")) * Me.ConvDecimal(dr("女＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿運輸業郵便業＿現金給与総額")) * Me.ConvDecimal(dr("女＿運輸業郵便業＿常用労働者数")))
        denominator = Aggregate dr In last Into _
                      Sum(Me.ConvDecimal(dr("男＿製造業＿労働時間")) * Me.ConvDecimal(dr("男＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿建設業＿労働時間")) * Me.ConvDecimal(dr("男＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("男＿運輸業郵便業＿労働時間")) * Me.ConvDecimal(dr("男＿運輸業郵便業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿製造業＿労働時間")) * Me.ConvDecimal(dr("女＿製造業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿建設業＿労働時間")) * Me.ConvDecimal(dr("女＿建設業＿常用労働者数")) _
                        + Me.ConvDecimal(dr("女＿運輸業郵便業＿労働時間")) * Me.ConvDecimal(dr("女＿運輸業郵便業＿常用労働者数")))

        Dim Avg_5_29_Zen As Decimal
        If denominator = 0 Then
            Avg_5_29_Zen = 0
        Else
            Avg_5_29_Zen = numerator / denominator
        End If
        Me.txtAvg_5_29_Zen.Text = Avg_5_29_Zen.ToString("0")

        '通勤手当割合／男女平均
        Dim Avg_Tsukin As Decimal = If(rouchin.Rows.Count > 0, Decimal.Parse(rouchin.Rows(0)("男女平均＿通勤手当割合").ToString), 1.7D)
        Me.txtAvg_Tsukin.Text = Avg_Tsukin.ToString("0.0")

        '30人以上／産業計対比／男女平均
        numerator = Aggregate dr In this Into
                      Sum(Me.ConvDecimal(dr("男女平均＿全産業計＿賃金対比")))
        denominator = Aggregate dr In this Into _
                      Sum(Me.ConvDecimal(dr("男女平均＿全産業計＿労働時間対比")))

        Dim Avg_30_Sangyo As Decimal
        If denominator = 0 Then
            Avg_30_Sangyo = 0
        Else
            Avg_30_Sangyo = numerator * 100 / denominator
        End If
        Me.txtAvg_30_Sangyo.Text = Avg_30_Sangyo.ToString("0.0")

        '優先順位2
        '５～29人／対前年比／男
        Dim M_5_29_Hi As Decimal
        If M_5_29_Zen = 0 Then
            M_5_29_Hi = 0
        Else
            M_5_29_Hi = Decimal.Parse(Me.txtM_5_29_Hon.Text) * 100 / Decimal.Parse(Me.txtM_5_29_Zen.Text)
        End If
        Me.txtM_5_29_Hi.Text = M_5_29_Hi.ToString("0.0")

        '５～29人／対前年比／女   
        Dim W_5_29_Hi As Decimal
        If W_5_29_Zen = 0 Then
            W_5_29_Hi = 0
        Else
            W_5_29_Hi = Decimal.Parse(Me.txtW_5_29_Hon.Text) * 100 / Decimal.Parse(Me.txtW_5_29_Zen.Text)
        End If
        Me.txtW_5_29_Hi.Text = W_5_29_Hi.ToString("0.0")

        '５～29人／対前年比／男女平均   
        Dim Avg_5_29_Hi As Decimal
        If Avg_5_29_Zen = 0 Then
            Avg_5_29_Hi = 0
        Else
            Avg_5_29_Hi = Decimal.Parse(Me.txtAvg_5_29_Hon.Text) * 100 / Decimal.Parse(Me.txtAvg_5_29_Zen.Text)
        End If
        Me.txtAvg_5_29_Hi.Text = Avg_5_29_Hi.ToString("0.0")

        '採用単価／本年値／男女平均   
        Dim Avg_Saiyo_Hon As Decimal = If(rouchin.Rows.Count > 0, Decimal.Parse(rouchin.Rows(0)("男女平均＿採用単価＿本年値").ToString), Avg_5_29_Hon)
        Me.txtAvg_Saiyo_Hon.Text = Avg_Saiyo_Hon.ToString("0")

        '採用単価／前年値／男女平均
        Dim Avg_Saiyo_Zen As Decimal = If(rouchin.Rows.Count > 0, Decimal.Parse(rouchin.Rows(0)("男女平均＿採用単価＿前年値").ToString), Avg_5_29_Zen)
        Me.txtAvg_Saiyo_Zen.Text = Avg_Saiyo_Zen.ToString("0")

        '30人以上／対比単価／男女平均
        Dim Avg_30_Taihi As Decimal = Decimal.Parse(Me.txtAvg_Saiyo_Zen.Text) * Decimal.Parse(Me.txtAvg_30_Sangyo.Text) / 100
        Me.txtAvg_30_Taihi.Text = Avg_30_Taihi.ToString("0")

        '優先順位3
        '採用単価／対前年比／男女平均
        Dim Avg_Saiyo_Hi As Decimal
        If Avg_Saiyo_Zen = 0 Then
            Avg_Saiyo_Hi = 0
        Else
            Avg_Saiyo_Hi = Decimal.Parse(Me.txtAvg_Saiyo_Hon.Text) * 100 / Decimal.Parse(Me.txtAvg_Saiyo_Zen.Text)
        End If
        Me.txtAvg_Saiyo_Hi.Text = Avg_Saiyo_Hi.ToString("0.0")

        '評価単価（65歳未満）／男女平均
        Dim Avg_Hyouka As Decimal = Decimal.Parse(Me.txtAvg_Saiyo_Hon.Text) * (100 - Decimal.Parse(Me.txtAvg_Tsukin.Text)) / 100
        Me.txtAvg_Hyouka.Text = Avg_Hyouka.ToString("0")
    End Sub

    ''' <summary>
    ''' フォームデータ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetFormData() As Dictionary(Of String, String)
        Dim ret As New Dictionary(Of String, String)

        ret("男＿５～２９人＿本年値") = txtM_5_29_Hon.Text
        ret("男＿５～２９人＿前年値") = txtM_5_29_Zen.Text
        ret("男＿５～２９人＿対前年比") = txtM_5_29_Hi.Text
        ret("女＿５～２９人＿本年値") = txtW_5_29_Hon.Text
        ret("女＿５～２９人＿前年値") = txtW_5_29_Zen.Text
        ret("女＿５～２９人＿対前年比") = txtW_5_29_Hi.Text
        ret("男女平均＿５～２９人＿本年値") = txtAvg_5_29_Hon.Text
        ret("男女平均＿５～２９人＿前年値") = txtAvg_5_29_Zen.Text
        ret("男女平均＿５～２９人＿対前年比") = txtAvg_5_29_Hi.Text
        ret("男女平均＿採用単価＿本年値") = txtAvg_Saiyo_Hon.Text
        ret("男女平均＿採用単価＿前年値") = txtAvg_Saiyo_Zen.Text
        ret("男女平均＿採用単価＿対前年比") = txtAvg_Saiyo_Hi.Text
        ret("男女平均＿通勤手当割合") = txtAvg_Tsukin.Text
        ret("男女平均＿評価単価") = txtAvg_Hyouka.Text
        ret("男女平均＿３０人以上＿産業計対比") = txtAvg_30_Sangyo.Text
        ret("男女平均＿３０人以上＿対比単価") = txtAvg_30_Taihi.Text

        Return ret
    End Function

    ''' <summary>
    ''' エラーチェック１
    ''' </summary>
    ''' <param name="dc"></param>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError1(dc As Dictionary(Of String, String), ByRef msgId As String) As Boolean
        Dim ret As Boolean = False
        Dim val As Decimal = 0
        '未入力チェック
        If dc("男女平均＿採用単価＿本年値").Equals(String.Empty) _
            Or dc("男女平均＿採用単価＿前年値").Equals(String.Empty) _
            Or dc("男女平均＿通勤手当割合").Equals(String.Empty) Then

            msgId = MessageID.MSG_E_031
            Return ret

        Else
            '数値以外チェック
            If Not (Decimal.TryParse(dc("男女平均＿採用単価＿本年値"), val) _
                And Decimal.TryParse(dc("男女平均＿採用単価＿前年値"), val) _
                And Decimal.TryParse(dc("男女平均＿通勤手当割合"), val)) Then

                msgId = MessageID.MSG_E_032
                Return ret

            End If

        End If

        ret = True

        Return ret
    End Function

    ''' <summary>
    ''' エラーチェック２
    ''' </summary>
    ''' <param name="dc"></param>
    ''' <param name="details"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError2(dc As Dictionary(Of String, String), ByRef details As List(Of String)) As Boolean
        Dim ret As Boolean = True

        Const max As Integer = ComConst.ERR_MESSAGE_MAX

        Dim cnt As Integer = 0

        '桁数チェック
        Dim msg As String = "{0}件目：{1}の桁数がデータベースの桁数を超えています。"

        For Each kv As KeyValuePair(Of String, String) In dc
            Dim pattern As String
            If kv.Key.Equals("男女平均＿通勤手当割合") Then
                pattern = "^-?[0-9]{1,10}(\.[0-9]{1})?$"
            ElseIf kv.Key.Equals("男＿５～２９人＿対前年比") _
                Or kv.Key.Equals("女＿５～２９人＿対前年比") _
                Or kv.Key.Equals("男女平均＿５～２９人＿対前年比") _
                Or kv.Key.Equals("男女平均＿採用単価＿対前年比") _
                Or kv.Key.Equals("男女平均＿３０人以上＿産業計対比") Then
                pattern = "^-?[0-9]{1,3}(\.[0-9]{1})?$"
            Else
                pattern = "^-?[0-9]{1,10}$"
            End If
            If Not Regex.IsMatch(kv.Value, pattern) Then
                cnt = cnt + 1
                details.Add(String.Format(msg, cnt.ToString.PadLeft(2), kv.Key))
                ret = False
                If cnt = max Then Exit For
            End If
        Next

        Return ret
    End Function

    Private Sub txtAvg_Saiyo_Hon_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtAvg_Saiyo_Hon.Validating
        Dim val As Decimal
        If Decimal.TryParse(Me.txtAvg_Saiyo_Hon.Text, val) And val <> 0 Then
            Dim val1 As Decimal

            '採用単価／対前年比／男女平均
            If Decimal.TryParse(Me.txtAvg_Saiyo_Zen.Text, val1) And val1 <> 0 Then
                txtAvg_Saiyo_Hi.Text = (val * 100 / val1).ToString("0.0")
            Else
                txtAvg_Saiyo_Hi.Text = 0D.ToString("0.0")
            End If

            '評価単価（65歳未満）／男女平均
            If Decimal.TryParse(Me.txtAvg_Tsukin.Text, val1) Then
                txtAvg_Hyouka.Text = (val * (100 - val1) / 100).ToString("0")
            Else
                txtAvg_Hyouka.Text = 0D.ToString("0")
            End If
        Else
            '採用単価／対前年比／男女平均
            txtAvg_Saiyo_Hi.Text = 0D.ToString("0.0")
            '評価単価（65歳未満）／男女平均
            txtAvg_Hyouka.Text = 0D.ToString("0")
        End If
    End Sub

    Private Sub txtAvg_Saiyo_Zen_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtAvg_Saiyo_Zen.Validating
        Dim val As Decimal
        If Decimal.TryParse(Me.txtAvg_Saiyo_Zen.Text, val) And val <> 0 Then
            Dim val1 As Decimal

            '採用単価／対前年比／男女平均
            If Decimal.TryParse(Me.txtAvg_Saiyo_Hon.Text, val1) Then
                txtAvg_Saiyo_Hi.Text = (val1 * 100 / val).ToString("0.0")
            Else
                txtAvg_Saiyo_Hi.Text = 0D.ToString("0.0")
            End If

            '30人以上／対比単価／男女平均
            If Decimal.TryParse(Me.txtAvg_30_Sangyo.Text, val1) Then
                txtAvg_30_Taihi.Text = (val * val1 / 100).ToString("0")
            Else
                txtAvg_30_Taihi.Text = 0D.ToString("0")
            End If
        Else
            '採用単価／対前年比／男女平均
            txtAvg_Saiyo_Hi.Text = 0D.ToString("0.0")
            '30人以上／対比単価／男女平均
            txtAvg_30_Taihi.Text = 0D.ToString("0")
        End If
    End Sub

    Private Sub txtAvg_Tsukin_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtAvg_Tsukin.Validating
        Dim val As Decimal
        If Decimal.TryParse(Me.txtAvg_Tsukin.Text, val) Then
            Dim val1 As Decimal

            '評価単価（65歳未満）／男女平均
            If Decimal.TryParse(Me.txtAvg_Saiyo_Hon.Text, val1) Then
                txtAvg_Hyouka.Text = (val1 * (100 - val) / 100).ToString("0")
            Else
                txtAvg_Hyouka.Text = 0D.ToString("0")
            End If
        Else
            '評価単価（65歳未満）／男女平均
            txtAvg_Hyouka.Text = 0D.ToString("0")
        End If
    End Sub
End Class
