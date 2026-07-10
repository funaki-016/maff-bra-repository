'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_000   | 2023.10.30 |大興電子通信        | 要件No.1 新規作成
'//            |            |                    |
'//*************************************************************************************************
Imports System.Text
Imports Microsoft.VisualBasic.FileIO

''' <summary>
''' 米在庫調査票データCSV取込画面
''' </summary>
''' <remarks></remarks>
Public Class BRA11020F

    ''' <summary>
    ''' 米在庫調査票データ
    ''' </summary>
    ''' <returns></returns>
    Public Property KomezaikoData As New List(Of String())

    ''' <summary>
    ''' 取込ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnInput_Click(sender As Object, e As EventArgs) Handles BtnInput.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            'ファイル選択
            Dim filePath = ComUtil.GetFilePath(Of OpenFileDialog)(Me, IniFileInfo.ExcelInPath, filter:="CSVファイル|*.csv")
            If String.IsNullOrEmpty(filePath) Then
                Return
            End If

            'ファイル読込
            Using parser = New TextFieldParser(filePath, Encoding.GetEncoding("shift-jis"))
                parser.TextFieldType = FileIO.FieldType.Delimited
                parser.SetDelimiters(",")
                parser.HasFieldsEnclosedInQuotes = True
                parser.TrimWhiteSpace = True
                While Not parser.EndOfData
                    KomezaikoData.Add(parser.ReadFields())
                End While
            End Using
            If KomezaikoData.Count = 0 Then
                Message.ShowMsgBox(MessageID.MSG_E_141, MsgBoxStyle.OkOnly)
                Return
            End If

            '返り値にOKを設定し、自画面を閉じる
            DialogResult = DialogResult.OK

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

End Class
