Imports System.Windows.Forms
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2020.10.26 |TSP)                | フェーズ3 要件No.2修正　
'//            |            |                    |
'//*************************************************************************************************

Public Class Message

    ''' <summary>メッセージ内の改行文字</summary>
    Public Const STR_CRLF As String = "<BR>"
    ''' <summary>絶対エラーが複数存在する際のメッセージ</summary>
    Public Const STR_MULTI_ERR_MSG As String = "他の行にも不正エラーが存在します。"

    ''' <summary>メッセージ識別文字</summary>
    Public Const STR_QUESTION As String = "Q"
    Public Const STR_INFOMATION As String = "I"
    Public Const STR_WARNING As String = "W"
    Public Const STR_ERROR As String = "E"

    ''' <summary>メッセージタイトル</summary>
    Public Const CAP_QUESTION As String = "確認"
    Public Const CAP_INFOMATION As String = "通知"
    Public Const CAP_WARNING As String = "警告"
    Public Const CAP_ERROR As String = "エラー"

    ''' <summary>メッセージテーブル</summary>
    Public Shared messageTable As DataTable

    ''' <summary>
    ''' 初期化処理
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub Initialize(ByVal messageFilePath As String)

        Dim ds As DataSet = Nothing

        Try

            messageTable = New DataTable

            messageTable.Columns.Add("key", Type.GetType("System.String"))
            messageTable.Columns.Add("value", Type.GetType("System.String"))

            Using reader As New System.IO.StreamReader(messageFilePath, System.Text.Encoding.Default)
                While (reader.Peek() >= 0)
                    Dim buf As String = reader.ReadLine()
                    Dim item() As String = buf.Split(","c)
                    If item.Length <> 2 OrElse item(0) = String.Empty OrElse item(1) = String.Empty Then
                        Continue While
                    End If
                    '<BR>を改行コードに置換
                    item(1) = item(1).Replace(STR_CRLF, vbCrLf)
                    messageTable.Rows.Add(item)
                End While
            End Using

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' メッセージボックスの呼び出しを行う
    ''' </summary>
    ''' <param name="messageId">メッセージID</param>
    ''' <param name="style">スタイル</param>
    ''' <param name="defBtn">デフォルトボタン</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function ShowMsgBox(ByVal messageId As String, _
                                                ByVal style As Microsoft.VisualBasic.MsgBoxStyle, _
                                                Optional ByVal defBtn As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1) As Microsoft.VisualBasic.MsgBoxResult

        Dim intReturn As Integer = -1

        Try

            intReturn = ShowMessageBox(messageId, style, defBtn)

            If intReturn = -1 Then
                MsgBox("メッセージの表示に失敗しました。", MsgBoxStyle.OkOnly, "エラー")
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return CType(intReturn, Microsoft.VisualBasic.MsgBoxResult)

    End Function

    ''' <summary>
    ''' メッセージボックスの呼び出しを行う
    ''' </summary>
    ''' <param name="msgId">メッセージID</param>
    ''' <param name="msgPara">メッセージ引数</param>
    ''' <param name="style">スタイル</param>
    ''' <param name="defBtn">デフォルトボタン</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function ShowMsgBox(ByVal msgId As String, _
                                                ByVal msgPara() As String, _
                                                ByVal style As Microsoft.VisualBasic.MsgBoxStyle, _
                                                Optional ByVal defBtn As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1) As Microsoft.VisualBasic.MsgBoxResult

        Dim intReturn As Integer = -1

        Try

            intReturn = ShowMessageBox(msgId, msgPara, style, defBtn)

            If intReturn = -1 Then
                MsgBox("メッセージの表示に失敗しました。", MsgBoxStyle.OkOnly, "エラー")
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return CType(intReturn, Microsoft.VisualBasic.MsgBoxResult)

    End Function

    ''' <summary>
    ''' メッセージボックスの呼び出しを行う
    ''' </summary>
    ''' <param name="hWnd">ウィンドウハンドル</param>
    ''' <param name="msgId">メッセージID</param>
    ''' <param name="style">スタイル</param>
    ''' <param name="defBtn">デフォルトボタン</param>
    ''' <remarks></remarks>
    Public Overloads Shared Function ShowMsgBox(ByVal hWnd As System.Windows.Forms.IWin32Window, _
                                                ByVal msgId As String, _
                                                ByVal style As Microsoft.VisualBasic.MsgBoxStyle, _
                                                Optional ByVal defBtn As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1) As Microsoft.VisualBasic.MsgBoxResult

        Dim intReturn As Integer = -1

        Try

            intReturn = ShowMessageBox(hWnd, msgId, style, defBtn)

            If intReturn = -1 Then
                MsgBox("メッセージの表示に失敗しました。", MsgBoxStyle.OkOnly, "エラー")
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return CType(intReturn, Microsoft.VisualBasic.MsgBoxResult)

    End Function

    ''' <summary>
    ''' メッセージボックスの呼び出しを行う
    ''' </summary>
    ''' <param name="hWnd">ウィンドウハンドル</param>
    ''' <param name="msgId">メッセージID</param>
    ''' <param name="msgPara">メッセージ引数</param>
    ''' <param name="style">スタイル</param>
    ''' <param name="defBtn">デフォルトボタン</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function ShowMsgBox(ByVal hWnd As System.Windows.Forms.IWin32Window, _
                                                ByVal msgId As String, _
                                                ByVal msgPara() As String, _
                                                ByVal style As Microsoft.VisualBasic.MsgBoxStyle, _
                                                Optional ByVal defBtn As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1) As Microsoft.VisualBasic.MsgBoxResult

        Dim intReturn As Integer = -1

        Try

            intReturn = ShowMessageBox(hWnd, msgId, msgPara, style, defBtn)

            If intReturn = -1 Then
                MsgBox("メッセージの表示に失敗しました。", MsgBoxStyle.OkOnly, "エラー")
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return CType(intReturn, Microsoft.VisualBasic.MsgBoxResult)

    End Function

    ''' <summary>
    ''' メッセージフォームの呼び出しを行う
    ''' </summary>
    ''' <param name="owner"></param>
    ''' <param name="label"></param>
    ''' <param name="msgId"></param>
    ''' <param name="msgPara"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowMsgForm(owner As System.Windows.Forms.IWin32Window, msgId As String, msgPara() As String, Optional label As String = Nothing)
        Dim frm As New MessageForm()

        Dim message As String
        Dim caption As String
        Dim icon As MessageBoxIcon

        Try
            message = GetMessage(msgId)
            '引数を置換
            message = MyStringFormat(message, msgPara)
            caption = GetMsgCaption(msgId)
            icon = GetMsgIcon(msgId)

            frm.ShowMessage(owner, caption, label, message, icon)
        Catch ex As Exception
            frm.Close()
            Throw
        End Try
    End Sub
    '--- REV.001 ADD START
    ''' <summary>
    ''' メッセージフォーム(はい/いいえボタン)の呼び出しを行う
    ''' </summary>
    ''' <param name="owner"></param>
    ''' <param name="msgId"></param>
    ''' <param name="msgPara"></param>
    ''' <param name="label"></param>
    ''' <remarks></remarks>
    Public Shared Function ShowMsgFormYesNo(owner As System.Windows.Forms.IWin32Window, msgId As String, msgPara() As String, Optional label As String = Nothing) As Windows.Forms.DialogResult
        Dim frm As New MessageFormYesNo()

        Dim message As String
        Dim caption As String
        Dim icon As MessageBoxIcon

        Dim ret As DialogResult

        Try
            message = GetMessage(msgId)
            '引数を置換
            message = MyStringFormat(message, msgPara)
            caption = GetMsgCaption(msgId)
            icon = GetMsgIcon(msgId)

            ret = frm.ShowMessage(owner, caption, label, message, icon)
            Return ret
        Catch ex As Exception
            frm.Close()
            Throw
        End Try
    End Function
    '--- REV.001 ADD END
    ''' <summary>
    ''' 画面表示用メッセージを取得する
    ''' </summary>
    ''' <param name="msgId"></param>
    ''' <param name="msgPara"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetMessage(ByVal msgId As String, _
                               ByVal msgPara() As String) As String

        Dim message As String

        Try

            message = GetMessage(msgId).ToString
            '引数を置換
            message = MyStringFormat(message, msgPara)

        Catch ex As Exception
            Throw ex
        End Try

        Return message
    End Function

    ''' <summary>
    ''' メッセージボックスを表示する
    ''' </summary>
    ''' <param name="msgId">メッセージコード</param>
    ''' <param name="type">スタイル</param>
    ''' <param name="defBtn">デフォルトボタン</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Shared Function ShowMessageBox(ByVal msgId As String, _
                                                     ByVal type As Integer, _
                                                     ByVal defBtn As MessageBoxDefaultButton) As Integer

        Dim intReturn As Integer
        Dim message As String
        Dim caption As String
        Dim icon As MessageBoxIcon

        Try

            message = GetMessage(msgId).ToString
            caption = GetMsgCaption(msgId)
            icon = GetMsgIcon(msgId)

            intReturn = MessageBox.Show(message, caption, CType(type, MessageBoxButtons), icon, defBtn)

        Catch ex As Exception
            Throw ex
        End Try

        Return CType(intReturn, Microsoft.VisualBasic.MsgBoxResult)

    End Function

    ''' <summary>
    ''' メッセージボックスを表示する
    ''' </summary>
    ''' <param name="msgId">メッセージコード</param>
    ''' <param name="msgPara">メッセージ引数</param>
    ''' <param name="type">スタイル</param>
    ''' <param name="defBtn">デフォルトボタン</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Shared Function ShowMessageBox(ByVal msgId As String, _
                                                     ByVal msgPara() As String, _
                                                     ByVal type As Integer, _
                                                     ByVal defBtn As MessageBoxDefaultButton) As Integer

        Dim intReturn As Integer
        Dim message As String
        Dim caption As String
        Dim icon As MessageBoxIcon

        Try

            message = GetMessage(msgId).ToString
            '引数を置換
            message = MyStringFormat(message, msgPara)
            caption = GetMsgCaption(msgId)
            icon = GetMsgIcon(msgId)

            intReturn = MessageBox.Show(message, caption, CType(type, MessageBoxButtons), icon, defBtn)

        Catch ex As Exception
            Throw ex
        End Try

        Return CType(intReturn, Microsoft.VisualBasic.MsgBoxResult)

    End Function

    ''' <summary>
    ''' メッセージボックスを表示する
    ''' </summary>
    ''' <param name="hWnd">ウィンドウハンドル</param>
    ''' <param name="msgId">メッセージコード</param>
    ''' <param name="type">スタイル</param>
    ''' <param name="defBtn">デフォルトボタン</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Shared Function ShowMessageBox(ByVal hWnd As System.Windows.Forms.IWin32Window, _
                                                     ByVal msgId As String, _
                                                     ByVal type As Integer, _
                                                     ByVal defBtn As MessageBoxDefaultButton) As Integer

        Dim intReturn As Integer
        Dim message As String
        Dim caption As String
        Dim icon As MessageBoxIcon

        Try

            message = GetMessage(msgId)
            caption = GetMsgCaption(msgId)
            icon = GetMsgIcon(msgId)

            intReturn = MessageBox.Show(hWnd, message, caption, CType(type, MessageBoxButtons), icon, defBtn)

        Catch ex As Exception
            Throw ex
        End Try

        Return CType(intReturn, Microsoft.VisualBasic.MsgBoxResult)

    End Function

    ''' <summary>
    ''' メッセージボックスを表示する
    ''' </summary>
    ''' <param name="hWnd">ウィンドウハンドル</param>
    ''' <param name="msgId">メッセージコード</param>
    ''' <param name="msgPara">メッセージ引数</param>
    ''' <param name="type">スタイル</param>
    ''' <param name="defBtn">デフォルトボタン</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Shared Function ShowMessageBox(ByVal hWnd As System.Windows.Forms.IWin32Window, _
                                                     ByVal msgId As String, _
                                                     ByVal msgPara() As String, _
                                                     ByVal type As Integer, _
                                                     ByVal defBtn As MessageBoxDefaultButton) As Integer

        Dim intReturn As Integer
        Dim message As String
        Dim caption As String
        Dim icon As MessageBoxIcon

        Try

            message = GetMessage(msgId).ToString
            '引数を置換
            message = MyStringFormat(message, msgPara)
            caption = GetMsgCaption(msgId)
            icon = GetMsgIcon(msgId)

            intReturn = MessageBox.Show(hWnd, message, caption, CType(type, MessageBoxButtons), icon, defBtn)

        Catch ex As Exception
            Throw ex
        End Try

        Return CType(intReturn, Microsoft.VisualBasic.MsgBoxResult)

    End Function

    ''' <summary>
    ''' 文字列の書式項目を文字列形式に置換する
    ''' </summary>
    ''' <param name="format"></param>
    ''' <param name="para"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function MyStringFormat(ByVal format As String, ByVal para() As String) As String

        Dim ret As String = Nothing
        Dim formatParaCount As Integer

        Try
            formatParaCount = UBound(Split(format, "{"))
            '書式文字列の引数指定の数と引数の数を比較
            If formatParaCount > para.Count Then
                '書式文字列の引数指定の数が多い場合、空文字を足す
                ReDim Preserve para(formatParaCount - 1)
            End If
            ret = String.Format(format, para)
        Catch ex As Exception
            Throw
        End Try

        Return ret

    End Function

    ''' <summary>
    ''' メッセージを取得する
    ''' </summary>
    ''' <param name="msgId">メッセージコード</param>
    ''' <returns>メッセージ As String</returns>
    ''' <remarks></remarks>
    Private Shared Function GetMessage(ByVal msgId As String) As String

        Dim strReturn As String = String.Empty
        Dim row() As DataRow

        Try

            row = messageTable.Select(String.Format("{0}='{1}'", "key", msgId))
            If row.Count > 0 Then
                strReturn = row(0).Item(1).ToString
            Else
                strReturn = "[" & msgid & "]に該当するメッセージが登録されていません。"
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return strReturn

    End Function

    ''' <summary>
    ''' メッセージタイトルを取得する
    ''' </summary>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetMsgCaption(ByVal msgId As String) As String

        Dim caption As String = String.Empty

        Try

            Select Case Mid(msgId, 4, 1)
                Case STR_QUESTION
                    caption = CAP_QUESTION
                Case STR_INFOMATION
                    caption = CAP_INFOMATION
                Case STR_WARNING
                    caption = CAP_WARNING
                Case Else
                    caption = CAP_ERROR
            End Select

        Catch ex As Exception
            Throw ex
        End Try

        Return caption

    End Function

    ''' <summary>
    ''' メッセージアイコンを取得する
    ''' </summary>
    ''' <param name="msgId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetMsgIcon(ByVal msgId As String) As MessageBoxIcon

        Dim icon As MessageBoxIcon = MessageBoxIcon.None

        Try

            Select Case Mid(msgId, 4, 1)
                Case STR_QUESTION
                    icon = MessageBoxIcon.Question
                Case STR_INFOMATION
                    icon = MessageBoxIcon.Information
                Case STR_WARNING
                    icon = MessageBoxIcon.Warning
                Case Else
                    icon = MessageBoxIcon.Error
            End Select

        Catch ex As Exception
            Throw ex
        End Try

        Return icon

    End Function

End Class

