'//*************************************************************************************************
'//  修正履歴
'// ------------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前           |                修  正  内  容
'// -----------+------------+------------------------+----------------------------------------------
'//  REV_001   | 2023.11.27 |Daiko                   | 要件No.20 新規作成
'//            |            |                        |
'//*************************************************************************************************

''' <summary>
''' シート選択画面
''' </summary>
''' <remarks></remarks>
Public Class BRA1420F

    ''' <summary>調査年</summary>
    Private _chosaNen As String
    Private _cb As CheckBox()

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(chosaNen As String)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

        _chosaNen = chosaNen
        _cb = {chk01, chk02, chk03, chk04, chk05, chk06, chk07, chk08, chk09, chk10, chk11, chk12, chk13, chk14, chk15, chk16, chk17, chk18, chk19, chk20, chk21, chk22, chk23, chk24, chk25, chk26, chk27, chk28, chk29, chk30}
    End Sub

    Private Sub BRA1420F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Dim i = 0
            For Each kv As KeyValuePair(Of String, String) In ComConst.調査票.シートデータ範囲(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun)))
                _cb(i).Text = kv.Key
                _cb(i).Checked = True
                i += 1
            Next

            For j As Integer = i To _cb.Length - 1
                _cb(j).Visible = False
            Next

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnAllSelect_Click(sender As Object, e As EventArgs) Handles btnAllSelect.Click
        Try
            For Each cb In _cb
                If cb.Visible Then
                    cb.Checked = True
                End If
            Next
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnAllCancel_Click(sender As Object, e As EventArgs) Handles btnAllCancel.Click
        Try
            For Each cb In _cb
                If cb.Visible Then
                    cb.Checked = False
                End If
            Next
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btnDecision_Click(sender As Object, e As EventArgs) Handles btnDecision.Click
        Try
            Dim sheetList As New List(Of String)
            For Each cb In _cb
                If cb.Checked Then
                    sheetList.Add(cb.Text)
                End If
            Next

            '地域設定
            DirectCast(Owner, BRA1410F).SetSheet(sheetList)

            Me.Close()
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

End Class
