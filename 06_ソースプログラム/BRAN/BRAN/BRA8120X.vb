Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop.Forms
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Text.RegularExpressions

''' <summary>
''' 毎勤データ入力・修正（EXCEL）
''' </summary>
''' <remarks></remarks>

Public Class BRA8120X
    Inherits ExcelInputBaseClass

    ''' <summary>エラーチェック種別</summary>
    Public Enum enmエラーチェック種別
        基本 = 1
        範囲
        追加
    End Enum

    ''' <summary>保存ボタン</summary>
    Private WithEvents btnSaveClose As CommandButton
    ''' <summary>戻るボタン</summary>
    Private WithEvents btnNoSaveClose As CommandButton

    ''' <summary>事務所</summary>
    Private _jimusho As String

    ''' <summary>Excelユーザーフォームハンドル</summary>
    Private _formHwnd As Win32WindowWrapper

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef frm As ExcelInputBaseForm, jimusho As String)
        MyBase.New(frm, System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.毎月勤労統計.入力用ファイル名称), True)
        Try
            _jimusho = jimusho

            'データを設定する
            Me.SetData()

            'Excel画面を表示する
            Me.ShowExcel()

            '処理待ち画面を閉じる
            Me.CloseWaitForm()
        Catch ex As Exception
            'Excel画面を閉じる
            Me.CloseExcel()
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Excel画面を表示する
    ''' </summary>
    ''' <param name="pFilePath"></param>
    ''' <remarks></remarks>
    Private Sub ShowExcel()
        Dim uf As UserForm = Nothing

        Try

            'ボタンの設定
            uf = ComUtilStrictOff.GetExcelForm(xlBook)
            btnSaveClose = ComUtilStrictOff.GetExcelBtnSaveClose(uf)
            btnNoSaveClose = ComUtilStrictOff.GetExcelBtnNoSaveClose(uf)

            'ユーザーフォームを表示する
            ComUtilStrictOff.ShowExcelForm(uf)

            'Excelユーザーフォームハンドル設定
            _formHwnd = New Win32WindowWrapper(ComUtilStrictOff.GetFormHwnd(uf))

            '画面更新有効
            xlApp.ScreenUpdating = True

        Catch ex As Exception
            Throw
        Finally
            ReleaseComObject(uf)
        End Try
    End Sub

    ''' <summary>
    ''' Excel画面を閉じる
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub CloseExcel()

        Try
            'ボタンの解放
            ReleaseComObject(btnSaveClose)
            ReleaseComObject(btnNoSaveClose)

            'Excel画面を閉じる
            MyBase.CloseExcel()

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 保存ボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnSaveClose_Click() Handles btnSaveClose.Click
        Try
            '確認メッセージ
            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_001, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                'Excel前処理
                Me.BeforeExcel()

                'シートデータ取得
                Dim lstDc As List(Of Dictionary(Of String, String)) = Me.GetSheetData()
                
                'エラーチェック
                Dim details As New List(Of String)
                If Not Me.CheckError(lstDc, details) Then
                    'エラーメッセージ
                    Message.ShowMsgForm(_formHwnd, MessageID.MSG_E_024, {String.Join(vbCrLf, details)})
                    Exit Sub
                End If

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    Try

                        db.BeginTrans()

                        '毎月勤労統計削除
                        DAOOther.DeleteMaitsukiKinrouToukei(db, _jimusho)

                        For Each dc As Dictionary(Of String, String) In lstDc
                            '毎月勤労統計追加
                            DAOOther.InsertMaitsukiKinrouToukei(db, _jimusho, dc)
                        Next

                        db.CommitTrans()

                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try
                End Using

                '完了メッセージ
                Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_001, MsgBoxStyle.OkOnly)
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            'Excel後処理
            Me.AfterExcel()

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 戻るボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnNoSaveClose_Click() Handles btnNoSaveClose.Click
        Try

            'Excel画面を閉じる
            Me.CloseExcel()

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' データを設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetData()
        Try
            xlApp.Interactive = False

            'Excel前処理
            Me.BeforeExcel()

            Dim dt As DataTable

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '毎月勤労統計取得
                dt = DAOOther.GetMaitsukiKinrouToukei(db, _jimusho)
            End Using

            'シートデータ設定
            ComUtil.MaitsukiKinrouToukei.SetSheetData(dt, _jimusho, xlSheets, CType(Me, ComObjectProcess))
        Catch ex As Exception
            Throw
        Finally
            'Excel後処理
            Me.AfterExcel()

            xlApp.Interactive = True
        End Try
    End Sub

    ''' <summary>
    ''' シートデータ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSheetData() As List(Of Dictionary(Of String, String))
        Dim ret As New List(Of Dictionary(Of String, String))

        Dim xlSheet As Excel.Worksheet = Nothing

        Try
            'シートの設定
            xlSheet = DirectCast(xlSheets.Item(ComConst.毎月勤労統計.出力用ファイル名称.SheetName), Excel.Worksheet)

            Dim rng1 As Excel.Range = Nothing
            Dim rng2 As Excel.Range = Nothing
            Dim rng3 As Excel.Range = Nothing
            Dim rngArr As Excel.Range = Nothing

            Try

                Dim arrData(,) As Object

                rng1 = xlSheet.Range(ComConst.毎月勤労統計.出力用ファイル名称.Col.First & ComConst.毎月勤労統計.出力用ファイル名称.Row.First)
                If Not rng1.Value Is Nothing Then
                    Dim last As Integer

                    rng2 = xlSheet.Range(ComConst.毎月勤労統計.出力用ファイル名称.Col.First & ComConst.毎月勤労統計.出力用ファイル名称.Row.First + 1)
                    If Not rng2.Value Is Nothing Then
                        rng3 = rng1.End(Excel.XlDirection.xlDown)
                        last = rng3.Row
                    Else
                        last = rng1.Row
                    End If

                    rngArr = xlSheet.Range(ComConst.毎月勤労統計.出力用ファイル名称.Col.First & ComConst.毎月勤労統計.出力用ファイル名称.Row.First & ":" _
                                        & ComConst.毎月勤労統計.出力用ファイル名称.Col.Last & last)

                    arrData = DirectCast(rngArr.Formula, Object(,))

                    For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                        Dim dc As New Dictionary(Of String, String)
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.毎月勤労統計.出力用ファイル名称.Field
                            dc(kv.Value) = arrData(i, kv.Key).ToString
                        Next
                        ret.Add(dc)
                    Next
                End If

            Finally
                ReleaseComObject(rng1)
                ReleaseComObject(rng2)
                ReleaseComObject(rng3)
                ReleaseComObject(rngArr)
            End Try
        Finally
            ReleaseComObject(xlSheet)
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="lstDc"></param>
    ''' <param name="details"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String)) As Boolean
        Dim ret As Boolean = True

        Const max As Integer = ComConst.ERR_MESSAGE_MAX

        Dim row As Integer = 0
        Dim cnt As Integer = 0

        Dim msg As String() = {"" _
                             , "{0}件目：{1}行目　月が入力されていません。" _
                             , "{0}件目：{1}行目　「{2}」に誤りがあります。" _
                             , "{0}件目：{1}行目　「{2}」に誤りがあります。" _
                             , "{0}件目：{1}行目　年月が重複しています。" _
                             , "{0}件目：{1}行目　「現金給与総額」、「労働時間」、「常用労働者数」に不足があります。全て入力するか、全て未入力かに修正してください。" _
                             , "{0}件目：{1}行目　「賃金対比」、「労働時間」が入力されていません。" _
        }

        For Each dc As Dictionary(Of String, String) In lstDc
            row = row + 1

            '月が入力されているか（年・月以外に入力がある）
            If dc("月").Equals(String.Empty) Then
                cnt = cnt + 1
                details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), row.ToString.PadLeft(3)))
                ret = False
                If cnt = max Then Exit For
            End If

            '入力された全ての項目が数値であるか。
            For Each kv As KeyValuePair(Of String, String) In dc
                If Not kv.Value.Equals(String.Empty) Then
                    Dim val As Decimal
                    If Not Decimal.TryParse(kv.Value, val) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), row.ToString.PadLeft(3), kv.Key))
                        ret = False
                        If cnt = max Then Exit For
                    End If
                End If
            Next

            '入力された全ての項目がデータベースの桁数を超えていないか。
            For Each kv As KeyValuePair(Of String, String) In dc
                If Not kv.Value.Equals(String.Empty) Then
                    Dim val As Decimal
                    If Decimal.TryParse(kv.Value, val) Then
                        Dim pattern As String
                        If kv.Key.Equals("年") Then
                            pattern = "^[0-9]{1,4}$"
                        ElseIf kv.Key.Equals("月") Then
                            pattern = "^[0-9]{1,2}$"
                        ElseIf kv.Key.Equals("男＿製造業＿現金給与総額") _
                            Or kv.Key.Equals("男＿建設業＿現金給与総額") _
                            Or kv.Key.Equals("男＿運輸業郵便業＿現金給与総額") _
                            Or kv.Key.Equals("女＿製造業＿現金給与総額") _
                            Or kv.Key.Equals("女＿建設業＿現金給与総額") _
                            Or kv.Key.Equals("女＿運輸業郵便業＿現金給与総額") Then
                            pattern = "^-?[0-9]{1,10}$"
                        ElseIf kv.Key.Equals("男＿製造業＿労働時間") _
                            Or kv.Key.Equals("男＿建設業＿労働時間") _
                            Or kv.Key.Equals("男＿運輸業郵便業＿労働時間") _
                            Or kv.Key.Equals("女＿製造業＿労働時間") _
                            Or kv.Key.Equals("女＿建設業＿労働時間") _
                            Or kv.Key.Equals("女＿運輸業郵便業＿労働時間") Then
                            pattern = "^[0-9]{1,6}(\.[0-9]{1})?$"
                        ElseIf kv.Key.Equals("男＿製造業＿常用労働者数") _
                            Or kv.Key.Equals("男＿建設業＿常用労働者数") _
                            Or kv.Key.Equals("男＿運輸業郵便業＿常用労働者数") _
                            Or kv.Key.Equals("女＿製造業＿常用労働者数") _
                            Or kv.Key.Equals("女＿建設業＿常用労働者数") _
                            Or kv.Key.Equals("女＿運輸業郵便業＿常用労働者数") Then
                            pattern = "^-?[0-9]{1,8}$"
                        Else
                            pattern = "^[0-9]{1,3}(\.[0-9]{1})?$"
                        End If
                        If Not Regex.IsMatch(kv.Value, pattern) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), row.ToString.PadLeft(3), kv.Key))
                            ret = False
                            If cnt = max Then Exit For
                        End If
                    End If

                End If
            Next

            '年・月に重複がないか。
            Dim query = From dct In lstDc Where dct("年").ToString = dc("年").ToString And dct("月").ToString = dc("月").ToString
            If query.Count > 1 Then
                cnt = cnt + 1
                details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), row.ToString.PadLeft(3)))
                ret = False
                If cnt = max Then Exit For
            End If

            '年月の入力がある場合、各産業別の「現金給与総額」、「労働時間」、「常用労働者数」が揃って入力されているか。（又は全て空欄）
            If dc("男＿製造業＿現金給与総額").Equals(String.Empty) Or dc("男＿製造業＿労働時間").Equals(String.Empty) Or dc("男＿製造業＿常用労働者数").Equals(String.Empty) Then
                If Not (dc("男＿製造業＿現金給与総額").Equals(String.Empty) And dc("男＿製造業＿労働時間").Equals(String.Empty) And dc("男＿製造業＿常用労働者数").Equals(String.Empty)) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(3)))
                    ret = False
                    If cnt = max Then Exit For
                End If
            End If

            If dc("男＿建設業＿現金給与総額").Equals(String.Empty) Or dc("男＿建設業＿労働時間").Equals(String.Empty) Or dc("男＿建設業＿常用労働者数").Equals(String.Empty) Then
                If Not (dc("男＿建設業＿現金給与総額").Equals(String.Empty) And dc("男＿建設業＿労働時間").Equals(String.Empty) And dc("男＿建設業＿常用労働者数").Equals(String.Empty)) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(3)))
                    ret = False
                    If cnt = max Then Exit For
                End If
            End If

            If dc("男＿運輸業郵便業＿現金給与総額").Equals(String.Empty) Or dc("男＿運輸業郵便業＿労働時間").Equals(String.Empty) Or dc("男＿運輸業郵便業＿常用労働者数").Equals(String.Empty) Then
                If Not (dc("男＿運輸業郵便業＿現金給与総額").Equals(String.Empty) And dc("男＿運輸業郵便業＿労働時間").Equals(String.Empty) And dc("男＿運輸業郵便業＿常用労働者数").Equals(String.Empty)) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(3)))
                    ret = False
                    If cnt = max Then Exit For
                End If
            End If

            If dc("女＿製造業＿現金給与総額").Equals(String.Empty) Or dc("女＿製造業＿労働時間").Equals(String.Empty) Or dc("女＿製造業＿常用労働者数").Equals(String.Empty) Then
                If Not (dc("女＿製造業＿現金給与総額").Equals(String.Empty) And dc("女＿製造業＿労働時間").Equals(String.Empty) And dc("女＿製造業＿常用労働者数").Equals(String.Empty)) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(3)))
                    ret = False
                    If cnt = max Then Exit For
                End If
            End If

            If dc("女＿建設業＿現金給与総額").Equals(String.Empty) Or dc("女＿建設業＿労働時間").Equals(String.Empty) Or dc("女＿建設業＿常用労働者数").Equals(String.Empty) Then
                If Not (dc("女＿建設業＿現金給与総額").Equals(String.Empty) And dc("女＿建設業＿労働時間").Equals(String.Empty) And dc("女＿建設業＿常用労働者数").Equals(String.Empty)) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(3)))
                    ret = False
                    If cnt = max Then Exit For
                End If
            End If

            If dc("女＿運輸業郵便業＿現金給与総額").Equals(String.Empty) Or dc("女＿運輸業郵便業＿労働時間").Equals(String.Empty) Or dc("女＿運輸業郵便業＿常用労働者数").Equals(String.Empty) Then
                If Not (dc("女＿運輸業郵便業＿現金給与総額").Equals(String.Empty) And dc("女＿運輸業郵便業＿労働時間").Equals(String.Empty) And dc("女＿運輸業郵便業＿常用労働者数").Equals(String.Empty)) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(3)))
                    ret = False
                    If cnt = max Then Exit For
                End If
            End If

            '年月の入力がある場合、全産業計（30人～）の「賃金対比」、「労働時間」共に入力されているか。
            If dc("男女平均＿全産業計＿賃金対比").Equals(String.Empty) Or dc("男女平均＿全産業計＿労働時間対比").Equals(String.Empty) Then
                cnt = cnt + 1
                details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), row.ToString.PadLeft(3)))
                ret = False
                If cnt = max Then Exit For
            End If
        Next

        Return ret
    End Function
End Class
