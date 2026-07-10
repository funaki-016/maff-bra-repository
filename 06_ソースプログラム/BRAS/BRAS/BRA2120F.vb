''' <summary>
''' 補正異動情報一覧画面
''' </summary>
''' <remarks></remarks>
Public Class BRA2120F

    ''' <summary>
    ''' 補正異動情報データクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class HoseiIdouInfo
        ''' <summary></summary>
        Public CensusNo As String
        ''' <summary></summary>
        Public KotaiNo As String
        ''' <summary></summary>
        Public OutIdoDate As String
        ''' <summary></summary>
        Public OutIdoFlag As String
        ''' <summary></summary>
        Public OutFarmCode As String
        ''' <summary></summary>
        Public InIdoDate As String
        ''' <summary></summary>
        Public InIdoFlag As String
        ''' <summary></summary>
        Public InFarmCode As String
        ''' <summary></summary>
        Public IsDelete As Boolean
    End Class


#Region "変数一覧"
    '補正リスト
    Private hoseiList As List(Of BRA2120F.HoseiIdouInfo)

#End Region

#Region "【処理詳細仕様 1】初期表示"
    Public Sub New(info As List(Of HoseiIdouInfo))
        InitializeComponent()

        hoseiList = info
    End Sub

    ''' <summary>
    ''' 画面起動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA2120F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            '一覧表示
            Me.ShowList(hoseiList)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

#End Region

#Region "【処理詳細仕様 4】「削除」ボタンクリック"

    ''' <summary>
    ''' 削除ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            'エラーチェック
            Dim msgId As String = String.Empty
            If Not Me.CheckError(msgId) Then
                'エラーメッセージ
                Message.ShowMsgBox(msgId, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_044, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            Else
                ' 削除フラグ変更
                Me.DeleteFlag()
                ' 一覧表示
                Me.ShowList(hoseiList)
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region

#Region "【処理詳細仕様 5】「戻る」ボタンクリック"

    ''' <summary>
    ''' 戻るボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnReturnhosei_Click(sender As Object, e As EventArgs) Handles btnReturnhosei.Click
        Try
            '確認メッセージ
            If Message.ShowMsgBox(MessageID.MSG_Q_047, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return
            Else
                Me.Close()
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub
#End Region

#Region "処理全般"

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
    ''' 一覧表示
    ''' </summary>
    ''' <param name="hoseiList"></param>
    ''' <remarks></remarks>
    Private Sub ShowList(hoseiList As List(Of BRA2120F.HoseiIdouInfo))
        dgvList.Rows.Clear()

        For row = 0 To hoseiList.Count - 1
            If (hoseiList(row).IsDelete = False) Then
                dgvList.Rows.Add()
                Dim i As Integer = dgvList.Rows.Count - 1
                dgvList.Rows(i).Cells(1).Value = hoseiList(row).CensusNo
                dgvList.Rows(i).Cells(2).Value = hoseiList(row).KotaiNo
                dgvList.Rows(i).Cells(3).Value = GetFormatedDate(hoseiList(row).OutIdoDate)
                dgvList.Rows(i).Cells(4).Value = ComConst.牛トレサデータ.異動名称(Integer.Parse(hoseiList(row).OutIdoFlag))
                dgvList.Rows(i).Cells(5).Value = hoseiList(row).OutFarmCode
                dgvList.Rows(i).Cells(6).Value = GetFormatedDate(hoseiList(row).InIdoDate)
                dgvList.Rows(i).Cells(7).Value = ComConst.牛トレサデータ.異動名称(Integer.Parse(hoseiList(row).InIdoFlag))
                dgvList.Rows(i).Cells(8).Value = hoseiList(row).InFarmCode
                dgvList.Rows(i).Cells(9).Value = row
            End If
        Next
    End Sub

    ''' <summary>
    ''' 日付を表示形式で返す
    ''' </summary>
    ''' <param name="dateString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetFormatedDate(dateString As String) As String
        If dateString.Length > 8 Then
            Return ""
        End If

        Return String.Format("{0}/{1}/{2}", dateString.Substring(0, 4), dateString.Substring(4, 2), dateString.Substring(6, 2))
    End Function

    ''' <summary>
    ''' 削除フラグ変更
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DeleteFlag()
        For row = 0 To dgvList.RowCount - 1
            If Convert.ToBoolean(dgvList.Rows(row).Cells(0).Value) Then
                For i = 0 To hoseiList.Count - 1
                    If CBool(i = Integer.Parse(dgvList.Rows(row).Cells(9).Value.ToString)) Then
                        hoseiList(i).IsDelete = True
                        Exit For
                    End If
                Next
            End If
        Next
    End Sub

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Function CheckError(ByRef msgId As String) As Boolean
        Dim ret As Boolean = False

        For i = 0 To dgvList.Rows.Count - 1
            If Convert.ToBoolean(dgvList.Rows(i).Cells(0).Value) Then
                ret = True
                Exit For
            End If
        Next

        If Not ret Then
            msgId = MessageID.MSG_E_003
        End If

        Return ret

    End Function

#End Region

End Class
