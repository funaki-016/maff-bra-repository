'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_000   | 2023.10.30 |大興電子通信        | 要件No.1 新規作成
'//  REV_001   | 2024.03.21 |大興電子通信        | 連絡票No.218対応
'//            |            |                    |
'//*************************************************************************************************
Imports System.IO
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Text
Imports Maff.BRA.Extensions.CellEx
Imports Maff.BRA.Extensions.DataRowEx
Imports Maff.BRA.BRA110.StringBuilderEx
Imports Microsoft.VisualBasic.FileIO

''' <summary>
''' 米在庫リンケージCSV出力画面
''' </summary>
''' <remarks></remarks>
Public Class BRA11010F

    ''' <summary>
    ''' 米在庫リンケージCSV最大列数
    ''' </summary>
    Private Const MAX_COLUMN = 2000

    ''' <summary>
    ''' 一覧FillWight
    ''' </summary>
    Private Shared ReadOnly FILL_WEIGHT As Integer = CInt(Math.Truncate(65535 / MAX_COLUMN))

    ''' <summary>
    ''' 一覧行番号
    ''' </summary>
    Private Enum DGV_ROW
        出力列 = 0
        調査
        項目名称
        個人調査票項番
        法人調査票項番
        米在庫調査票列
    End Enum

    ''' <summary>
    ''' コード変換（個人）<br />
    ''' REV_001：提示資料にミスがあったため、以下のコードを修正<br />
    ''' Q00000301 13：露地にんじん → 13：露地にんじん作<br />
    ''' </summary>
    Private Shared ReadOnly KOJIN_CONV As New Dictionary(Of String, Dictionary(Of String, String)) From {
        {
            "Q00000201", New Dictionary(Of String, String) From
            {
                {"1：水田作", "1"},
                {"2：畑作", "2"},
                {"3：露地野菜作", "3"},
                {"4：施設野菜作", "4"},
                {"5：果樹作", "5"},
                {"6：露地花き作", "6"},
                {"7：施設花き作", "7"},
                {"8：酪農", "8"},
                {"9：繁殖牛", "9"},
                {"10：肥育牛", "10"},
                {"11：養豚", "11"},
                {"12：採卵養鶏", "12"},
                {"13：ブロイラー養鶏", "13"},
                {"14：その他", "14"}
            }
        },
        {
            "Q00000301", New Dictionary(Of String, String) From
            {
                {"1：かんしょ作", "1"},
                {"2：ばれいしょ作", "2"},
                {"3：露地きゅうり作", "3"},
                {"4：露地大玉トマト作", "4"},
                {"5：露地なす作", "5"},
                {"6：露地キャベツ作", "6"},
                {"7：露地ほうれんそう作", "7"},
                {"8：露地たまねぎ作", "8"},
                {"9：露地レタス作", "9"},
                {"10：露地はくさい作", "10"},
                {"11：露地白ねぎ作", "11"},
                {"12：露地だいこん作", "12"},
                {"13：露地にんじん作", "13"},
                {"14：施設きゅうり作", "14"},
                {"15：施設大玉トマト作", "15"},
                {"16：施設ミニトマト作", "16"},
                {"17：施設なす作", "17"},
                {"18：りんご作", "18"},
                {"19：露地温州みかん作", "19"},
                {"20：施設温州みかん作", "20"},
                {"21：露地ぶどう作", "21"},
                {"22：施設ぶどう作", "22"},
                {"23：日本なし作", "23"},
                {"24：もも作", "24"},
                {"25：かき作", "25"},
                {"26：うめ作", "26"},
                {"27：おうとう作", "27"},
                {"28：キウイフルーツ作", "28"},
                {"29：すもも作", "29"},
                {"30：施設ばら作", "30"},
                {"31：茶作", "31"}
            }
        },
        {
            "Q00000501", New Dictionary(Of String, String) From
            {
                {"基本調査", "1"},
                {"詳細調査", "2"}
            }
        },
        {
            "Q00000601", New Dictionary(Of String, String) From
            {
                {"1", "1"},
                {"2", "2"},
                {"3", "3"},
                {"4", "4"},
                {"5", "5"}
            }
        },
        {
            "Q01010101", New Dictionary(Of String, String) From
            {
                {"0：いいえ", "0"},
                {"1：はい", "1"}
            }
        },
        {
            "Q01010102", New Dictionary(Of String, String) From
            {
                {"0：いいえ", "0"},
                {"1：はい", "1"}
            }
        },
        {
            "Q01010103", New Dictionary(Of String, String) From
            {
                {"0：いいえ", "0"},
                {"1：はい", "1"}
            }
        },
        {
            "Q01010151", New Dictionary(Of String, String) From
            {
                {"1：男", "1"},
                {"2：女", "2"}
            }
        },
        {
            "Q01010153", New Dictionary(Of String, String) From
            {
                {"0：いいえ", "0"},
                {"1：はい", "1"}
            }
        },
        {
            "Q01010154", New Dictionary(Of String, String) From
            {
                {"0：いいえ", "0"},
                {"1：はい", "1"}
            }
        },
        {
            "Q02010301", New Dictionary(Of String, String) From
            {
                {"0：いいえ", "0"},
                {"1：はい", "1"}
            }
        },
        {
            "Q03020301", New Dictionary(Of String, String) From
            {
                {"1", "1"}
            }
        },
        {
            "Q08011302", New Dictionary(Of String, String) From
            {
                {"1：ａ", "1"},
                {"2：㎡", "2"}
            }
        },
        {
            "Q10010101", New Dictionary(Of String, String) From
            {
                {"1：男", "1"},
                {"2：女", "2"}
            }
        },
        {
            "Q10010102", New Dictionary(Of String, String) From
            {
                {"1：男", "1"},
                {"2：女", "2"}
            }
        },
        {
            "Q10010302", New Dictionary(Of String, String) From
            {
                {"1：家族", "1"},
                {"2：７ヶ月未満の雇用者", "2"},
                {"3：７ヶ月以上の雇用者", "3"},
                {"4：臨時雇用者", "4"}
            }
        }
    }

    ''' <summary>
    ''' コード変換（法人）
    ''' </summary>
    Private Shared ReadOnly HOJIN_CONV As New Dictionary(Of String, Dictionary(Of String, String)) From {
        {
            "Q00000201", New Dictionary(Of String, String) From
            {
                {"1：水田作", "1"},
                {"2：畑作", "2"},
                {"3：露地野菜作", "3"},
                {"4：施設野菜作", "4"},
                {"5：果樹作", "5"},
                {"6：露地花き作", "6"},
                {"7：施設花き作", "7"},
                {"8：酪農", "8"},
                {"9：繁殖牛", "9"},
                {"10：肥育牛", "10"},
                {"11：養豚", "11"},
                {"12：採卵養鶏", "12"},
                {"13：ブロイラー養鶏", "13"},
                {"14：その他", "14"}
            }
        },
        {
            "Q01010101", New Dictionary(Of String, String) From
            {
                {"0：いいえ", "0"},
                {"1：はい", "1"}
            }
        },
        {
            "Q01010102", New Dictionary(Of String, String) From
            {
                {"0：いいえ", "0"},
                {"1：はい", "1"}
            }
        },
        {
            "Q01010104", New Dictionary(Of String, String) From
            {
                {"0：いいえ", "0"},
                {"1：はい", "1"}
            }
        },
        {
            "Q01010106", New Dictionary(Of String, String) From
            {
                {"1：男", "1"},
                {"2：女", "2"}
            }
        },
        {
            "Q01010108", New Dictionary(Of String, String) From
            {
                {"0：いいえ", "0"},
                {"1：はい", "1"}
            }
        },
        {
            "Q05010151", New Dictionary(Of String, String) From
            {
                {"いいえ", "0"},
                {"はい", "1"}
            }
        },
        {
            "Q05010152", New Dictionary(Of String, String) From
            {
                {"いいえ", "0"},
                {"はい", "1"}
            }
        },
        {
            "Q05010153", New Dictionary(Of String, String) From
            {
                {"いいえ", "0"},
                {"はい", "1"}
            }
        },
        {
            "Q05030305", New Dictionary(Of String, String) From
            {
                {"いいえ", "0"},
                {"はい", "1"}
            }
        }
    }

    ''' <summary>
    ''' 初期表示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BRA11010F_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'DGVちらつき防止
            dgvList.GetType.InvokeMember("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, Nothing, dgvList, New Object() {True})

            '画面初期化
            Initialize()

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 初期化
    ''' </summary>
    Private Sub Initialize()
        '画面クリア
        TxtFilename.Text = ""
        TxtChosanen.Text = ""
        dgvList.Rows.Clear()
        dgvList.Columns.Clear()
        '画面非活性化
        TxtFilename.Enabled = False
        TxtChosanen.Enabled = False
        BtnOutput.Enabled = False
        BtnCsvOutput.Enabled = False
    End Sub

    ''' <summary>
    ''' ファイル識別名称非アクティブ
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TxtFilename_Leave(sender As Object, e As EventArgs) Handles TxtFilename.Leave
        Try
            '半角を全角に変更
            TxtFilename.Text = StrConv(TxtFilename.Text, VbStrConv.Wide) _
                .Replace("\", "￥") _
                .Replace("'", "’") _
                .Replace("""", ChrW(&HFF02))

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 設定入力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnInput_Click(sender As Object, e As EventArgs) Handles BtnInput.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            '設定ファイル選択
            Dim filePath = ComUtil.GetFilePath(Of OpenFileDialog)(Me, IniFileInfo.ExcelInPath, filter:="CSVファイル|*.csv")
            If String.IsNullOrEmpty(filePath) Then
                Return
            End If

            '画面初期化
            Initialize()

            '設定ファイル読込
            Dim settings = New List(Of String())
            Using parser = New TextFieldParser(filePath, Encoding.GetEncoding("shift-jis"))
                parser.TextFieldType = FileIO.FieldType.Delimited
                parser.SetDelimiters(",")
                parser.HasFieldsEnclosedInQuotes = True
                parser.TrimWhiteSpace = True
                While Not parser.EndOfData
                    settings.Add(parser.ReadFields())
                End While
            End Using
            If settings.Count = 0 Then
                Message.ShowMsgBox(MessageID.MSG_E_141, MsgBoxStyle.OkOnly)
                Return
            End If
            '最大列数
            If settings.Count > MAX_COLUMN + 1 Then
                Message.ShowMsgBox(MessageID.MSG_E_142, {MAX_COLUMN.ToString}, MsgBoxStyle.OkOnly)
                Return
            End If

            '１行目の処理
            If settings(0).Length < 2 Then
                Message.ShowMsgBox(MessageID.MSG_E_143, {"1"}, MsgBoxStyle.OkOnly)
                Return
            End If

            '２行目以降の処理
            Dim lineH(settings.Count) As String
            Dim line0(settings.Count) As String
            Dim line1(settings.Count) As String
            Dim line2(settings.Count) As String
            Dim line3(settings.Count) As String
            Dim line4(settings.Count) As String
            lineH(0) = "出力列"
            line0(0) = "調査(0:固定値、1:営農、2:米在庫)"
            line1(0) = "項目名称"
            line2(0) = "個人調査票項番"
            line3(0) = "法人調査票項番"
            line4(0) = "米在庫調査票列"
            For i = 1 To settings.Count - 1
                Dim setting = settings(i)
                If setting.Length <> 5 Then
                    Message.ShowMsgBox(MessageID.MSG_E_143, {(i + 1).ToString}, MsgBoxStyle.OkOnly)
                    Return
                End If
                '出力列
                lineH(i) = i.ToString
                '調査
                line0(i) = setting(0)
                '項目名称
                line1(i) = setting(1)
                '個人調査票項番
                line2(i) = setting(2)
                '法人調査票項番
                line3(i) = setting(3)
                '米在庫調査票列
                line4(i) = setting(4)
            Next

            'ファイル識別名称
            TxtFilename.Text = settings(0)(0)
            TxtFilename_Leave(TxtFilename, New EventArgs())
            '調査年
            TxtChosanen.Text = settings(0)(1)

            '一覧に列追加
            For i = 0 To settings.Count - 1
                Dim col = New DataGridViewTextBoxColumn
                col.Width = 100
                col.CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                dgvList.Columns.Add(col)
                If i = 0 Then
                    '先頭列は読取専用、列固定
                    col.Width = 200
                    col.ReadOnly = True
                    col.Frozen = True
                End If
            Next
            '一覧に行追加
            AddDgvRow(lineH).ReadOnly = True
            AddDgvRow(line0)
            AddDgvRow(line1)
            AddDgvRow(line2)
            AddDgvRow(line3)
            AddDgvRow(line4)

            '一覧設定
            dgvList.ClearSelection()
            '２列目以降は調査の値によって活性／非活性制御
            For i = 1 To dgvList.Columns.Count - 1
                Dim cell = dgvList.Rows(DGV_ROW.調査).Cells(i)
                ChosaControl(cell)
            Next

            '画面活性化
            TxtFilename.Enabled = True
            TxtChosanen.Enabled = True
            BtnOutput.Enabled = True
            BtnCsvOutput.Enabled = True

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

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
    ''' 設定出力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnOutput_Click(sender As Object, e As EventArgs) Handles BtnOutput.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            '共通入力チェック
            If Not CommonCheck() Then
                Return
            End If

            '調査年取得
            Dim chosaNen = CInt(TxtChosanen.Text)

            'ファイルパス取得
            Dim fileName = String.Format("設定_{0}_{1}_{2}.csv", TxtFilename.Text, chosaNen, Now.ToString("yyyyMMdd"))
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName:=fileName, filter:="CSVファイル|*.csv")
            If String.IsNullOrEmpty(filePath) Then
                Return
            End If

            'CSV出力
            Using sw As New StreamWriter(filePath, False, Encoding.GetEncoding("shift_jis"))
                '１行目出力
                sw.WriteLine(String.Join(",", {TxtFilename.Text, chosaNen, "", "", ""}))
                '２行目以降出力
                For i = 1 To dgvList.Columns.Count - 1
                    If String.IsNullOrEmpty(dgvList.Rows(DGV_ROW.調査).Cells(i).GetString) Then
                        '調査が空の列は出力しない
                        Continue For
                    End If
                    Dim row = New List(Of Object)
                    For j = 1 To dgvList.Rows.Count - 1
                        row.Add(dgvList.Rows(j).Cells(i).GetString)
                    Next
                    sw.WriteLine(String.Join(",", row))
                Next
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
    ''' 共通入力チェック
    ''' </summary>
    ''' <returns></returns>
    Private Function CommonCheck() As Boolean

        'ファイル識別名称
        '必須
        If String.IsNullOrEmpty(TxtFilename.Text) Then
            Message.ShowMsgBox(MessageID.MSG_E_144, MsgBoxStyle.OkOnly)
            TxtFilename.Focus()
            Return False
        End If

        '調査年
        '必須
        If String.IsNullOrEmpty(TxtChosanen.Text) Then
            Message.ShowMsgBox(MessageID.MSG_E_145, MsgBoxStyle.OkOnly)
            TxtChosanen.Focus()
            Return False
        End If
        '数字
        Dim chosaNen As Integer
        If Not Integer.TryParse(TxtChosanen.Text, chosaNen) Then
            Message.ShowMsgBox(MessageID.MSG_E_146, MsgBoxStyle.OkOnly)
            TxtChosanen.Focus()
            Return False
        End If
        '2022以上
        If chosaNen < 2022 Then
            Message.ShowMsgBox(MessageID.MSG_E_147, MsgBoxStyle.OkOnly)
            TxtChosanen.Focus()
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' CSV出力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnCsvOutput_Click(sender As Object, e As EventArgs) Handles BtnCsvOutput.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            '共通入力チェック
            If Not CommonCheck() Then
                Return
            End If

            '調査年取得
            Dim chosaNen = CInt(TxtChosanen.Text)

            '一覧のチェック
            Dim errorCell = CheckDgvList(chosaNen)
            If errorCell IsNot Nothing Then
                dgvList.CurrentCell = errorCell
                If Not errorCell.ReadOnly Then
                    dgvList.BeginEdit(True)
                End If
                Return
            End If

            '調査に2:米在庫が指定されている場合
            Dim komezaikoData As List(Of String()) = Nothing
            If dgvList.Rows(DGV_ROW.調査).Cells.Cast(Of DataGridViewCell).Where(Function(x) x.GetString = "2").FirstOrDefault IsNot Nothing Then
                '米在庫調査票データCSV取込
                Using frm = New BRA11020F
                    If frm.ShowDialog(Me) = DialogResult.OK Then
                        komezaikoData = frm.KomezaikoData
                    Else
                        Return
                    End If
                End Using
            End If

            'ファイルパス取得
            Dim fileName = String.Format("米在庫リンケージ_{0}_{1}_{2}.csv", TxtFilename.Text, chosaNen, Now.ToString("yyyyMMdd"))
            Dim filePath As String = ComUtil.GetFilePath(Of SaveFileDialog)(Me, IniFileInfo.ExcelOutPath, fileName:=fileName, filter:="CSVファイル|*.csv")
            If String.IsNullOrEmpty(filePath) Then
                Return
            End If

            '営農個人データ検索
            Dim dtKojin = SelectEino("調査票＿農業経営＿営農類型＿個人", chosaNen)
            '営農個人可変データ検索
            Dim dtKahen = SelectEinoKojinKahen(chosaNen)
            '可変データを加工
            Dim dcCensus = New Dictionary(Of String, Dictionary(Of String, String()))
            For Each gCensus In dtKahen.Rows.Cast(Of DataRow).GroupBy(Function(x) x.GetString("センサス番号"))
                Dim dcItem = New Dictionary(Of String, String())
                For Each gItem In gCensus.GroupBy(Function(x) x.GetString("項目番号"))
                    Dim meisai() As String = Nothing
                    Dim isFirst = True
                    For Each dr In gItem
                        Dim no = dr.GetInteger("明細番号")
                        If isFirst Then
                            isFirst = False
                            ReDim meisai(no + 1)
                        End If
                        meisai(no) = dr.GetString("値")
                    Next
                    dcItem.Add(gItem.Key, meisai)
                Next
                dcCensus.Add(gCensus.Key, dcItem)
            Next

            '営農法人データ検索
            Dim dtHojin = SelectEino("調査票＿農業経営＿営農類型＿法人", chosaNen)

            If dtKojin.Rows.Count = 0 AndAlso dtHojin.Rows.Count = 0 Then
                'データなしエラー
                Message.ShowMsgBox(MessageID.MSG_E_155, MsgBoxStyle.OkOnly)
                TxtChosanen.Focus()
                Return
            End If

            '進捗ダイアログ
            Dim progressDialog As New ProgressDialog
            progressDialog.Maximum = dtKojin.Rows.Count + dtHojin.Rows.Count
            progressDialog.Show(Me)
            Dim ret = False
            Try
                '出力設定を取得
                Dim settings = New List(Of (chosa As String, title As String, kojin As String, kojinConv As Boolean, hojin As String, hojinConv As Boolean, kome As Integer))
                For i = 1 To dgvList.Columns.Count - 1
                    Dim chosa = dgvList.Rows(DGV_ROW.調査).Cells(i).GetString
                    If String.IsNullOrEmpty(chosa) Then
                        '調査が空の列は出力しない
                        Continue For
                    End If
                    Dim kojin = dgvList.Rows(DGV_ROW.個人調査票項番).Cells(i).GetString
                    Dim hojin = dgvList.Rows(DGV_ROW.法人調査票項番).Cells(i).GetString
                    Dim kome As Integer
                    If Not Integer.TryParse(dgvList.Rows(DGV_ROW.米在庫調査票列).Cells(i).GetString, kome) Then
                        kome = Integer.MaxValue
                    End If
                    settings.Add((dgvList.Rows(DGV_ROW.調査).Cells(i).GetString,
                        dgvList.Rows(DGV_ROW.項目名称).Cells(i).GetString,
                        kojin,
                        KOJIN_CONV.ContainsKey(kojin.Split("_"c)(0)),
                        hojin,
                        HOJIN_CONV.ContainsKey(hojin),
                        kome))
                Next

                'CSV出力
                Using sw As New StreamWriter(filePath, False, Encoding.GetEncoding("shift_jis"))

                    'ヘッダの出力
                    Dim header = New StringBuilder
                    settings.Select(Function(x) x.title).ToList.ForEach(Sub(x) header.AddCol(x))
                    sw.WriteLine(header.ToCsvLine)

                    '営農個人の出力
                    For Each kojin As DataRow In dtKojin.Rows
                        Dim censusNo = kojin.GetString("センサス番号")
                        Dim dcItem As Dictionary(Of String, String()) = Nothing
                        dcCensus.TryGetValue(censusNo, dcItem)
                        Dim komezaiko = komezaikoData?.Where(Function(x) x(0).Replace("a", "") = censusNo).FirstOrDefault
                        Dim row = New StringBuilder
                        For Each setting In settings
                            If setting.chosa = "0" Then
                                '固定値
                                row.AddCol(setting.kojin)
                            ElseIf setting.chosa = "1" Then
                                '営農
                                If String.IsNullOrEmpty(setting.kojin) Then
                                    '指定なし
                                    row.AddEmptyCol()
                                ElseIf setting.kojin.Contains("_") Then
                                    '可変
                                    If dcItem IsNot Nothing Then
                                        '項目番号と明細番号に分割
                                        Dim koban = Split(setting.kojin, "_")
                                        If dcItem.ContainsKey(koban(0)) AndAlso dcItem(koban(0)).Length > CInt(koban(1)) Then
                                            Dim val = If(dcItem(koban(0))(CInt(koban(1))), "")
                                            If setting.kojinConv Then
                                                If KOJIN_CONV(koban(0)).ContainsKey(val) Then
                                                    val = KOJIN_CONV(koban(0))(val)
                                                End If
                                            End If
                                            row.AddCol(val)
                                        Else
                                            row.AddEmptyCol()
                                        End If
                                    Else
                                        row.AddEmptyCol()
                                    End If
                                Else
                                    '可変以外
                                    Dim val = kojin.GetString(setting.kojin)
                                    If setting.kojinConv Then
                                        If KOJIN_CONV(setting.kojin).ContainsKey(val) Then
                                            val = KOJIN_CONV(setting.kojin)(val)
                                        End If
                                    End If
                                    row.AddCol(val)
                                End If
                            Else
                                '米在庫
                                If komezaiko IsNot Nothing AndAlso komezaiko.Length > setting.kome - 1 Then
                                    row.AddCol(komezaiko(setting.kome - 1))
                                Else
                                    row.AddEmptyCol()
                                End If
                            End If
                        Next
                        sw.WriteLine(row.ToCsvLine)

                        '進捗を進める
                        progressDialog.AddValue = 1
                    Next

                    '営農法人の出力
                    For Each hojin As DataRow In dtHojin.Rows
                        Dim censusNo = hojin.GetString("センサス番号")
                        Dim komezaiko = komezaikoData?.Where(Function(x) x(0).Replace("a", "") = censusNo).FirstOrDefault
                        Dim row = New StringBuilder
                        For Each setting In settings
                            If setting.chosa = "0" Then
                                '固定値
                                row.AddCol(setting.hojin)
                            ElseIf setting.chosa = "1" Then
                                '営農
                                If String.IsNullOrEmpty(setting.hojin) Then
                                    '指定なし
                                    row.AddEmptyCol()
                                Else
                                    '指定あり
                                    Dim val = hojin.GetString(setting.hojin)
                                    If setting.hojinConv Then
                                        If HOJIN_CONV(setting.hojin).ContainsKey(val) Then
                                            val = HOJIN_CONV(setting.hojin)(val)
                                        End If
                                    End If
                                    row.AddCol(val)
                                End If
                            Else
                                '米在庫
                                If komezaiko IsNot Nothing AndAlso komezaiko.Length > setting.kome - 1 Then
                                    row.AddCol(komezaiko(setting.kome - 1))
                                Else
                                    row.AddEmptyCol()
                                End If
                            End If
                        Next
                        sw.WriteLine(row.ToCsvLine)

                        '進捗を進める
                        progressDialog.AddValue = 1
                    Next

                End Using

                '正常終了
                ret = True

            Catch ioe As IOException
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ioe)
                progressDialog.ShowMsgBox(MessageID.MSG_E_058, MsgBoxStyle.OkOnly)
            Finally
                If progressDialog IsNot Nothing Then
                    '進捗ダイアログを閉じる
                    progressDialog.endDispose()
                    progressDialog = Nothing
                End If
            End Try

            '完了メッセージ
            If ret Then
                '米在庫調査票CSVのみに存在するデータをカウント
                Dim komeOnlyCount = 0
                If komezaikoData IsNot Nothing Then
                    Dim kojinCensus = dtKojin.Rows.Cast(Of DataRow).Select(Function(x) x.GetString("センサス番号")).ToList
                    Dim hojinCensus = dtHojin.Rows.Cast(Of DataRow).Select(Function(x) x.GetString("センサス番号")).ToList
                    komeOnlyCount = komezaikoData.Select(Function(x) x(0).Replace("a", "")).Where(Function(y) IsNumeric(y)).Where(Function(z) Not kojinCensus.Contains(z) AndAlso Not hojinCensus.Contains(z)).Count
                End If
                If komeOnlyCount = 0 Then
                    Message.ShowMsgBox(MessageID.MSG_I_002, MsgBoxStyle.OkOnly)
                Else
                    Message.ShowMsgBox(MessageID.MSG_W_001, {komeOnlyCount.ToString("#,0")}, MsgBoxStyle.OkOnly)
                End If
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
    ''' 一覧のチェック
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    Private Function CheckDgvList(chosaNen As Integer) As DataGridViewCell
        '調査
        '1:営農がない
        If dgvList.Rows(DGV_ROW.調査).Cells.Cast(Of DataGridViewCell).Where(Function(x) x.GetString = "1").FirstOrDefault Is Nothing Then
            Message.ShowMsgBox(MessageID.MSG_E_148, MsgBoxStyle.OkOnly)
            Return dgvList.Rows(DGV_ROW.調査).Cells(0)
        End If

        '項目名称
        '活性の場合、必須
        Dim cell = dgvList.Rows(DGV_ROW.項目名称).Cells.Cast(Of DataGridViewCell).Where(Function(x) Not x.ReadOnly AndAlso String.IsNullOrEmpty(x.GetString)).FirstOrDefault
        If cell IsNot Nothing Then
            Message.ShowMsgBox(MessageID.MSG_E_149, MsgBoxStyle.OkOnly)
            Return cell
        End If

        '調査票項番
        '調査が1:営農の場合、個人／法人いずれか必須
        For Each i In dgvList.Rows(DGV_ROW.調査).Cells.Cast(Of DataGridViewCell).Where(Function(x) x.GetString = "1").Select(Function(x) x.ColumnIndex).ToList
            If String.IsNullOrEmpty(dgvList.Rows(DGV_ROW.個人調査票項番).Cells(i).GetString) _
                AndAlso String.IsNullOrEmpty(dgvList.Rows(DGV_ROW.法人調査票項番).Cells(i).GetString) Then
                Message.ShowMsgBox(MessageID.MSG_E_150, MsgBoxStyle.OkOnly)
                Return dgvList.Rows(DGV_ROW.個人調査票項番).Cells(i)
            End If
        Next

        '個人調査票項番
        '存在チェック
        Dim kojinItemNos = GetItemNoList(True, chosaNen)
        For Each i In dgvList.Rows(DGV_ROW.調査).Cells.Cast(Of DataGridViewCell).Where(Function(x) x.GetString = "1").Select(Function(x) x.ColumnIndex).ToList
            cell = dgvList.Rows(DGV_ROW.個人調査票項番).Cells(i)
            Dim itemNo = cell.GetString
            If String.IsNullOrEmpty(itemNo) Then
                Continue For
            End If
            If IsKahenItemNo(itemNo) Then
                '可変の項番のみ取得
                itemNo = itemNo.Split("_"c)(0)
            End If
            If Not kojinItemNos.Contains(itemNo) Then
                Message.ShowMsgBox(MessageID.MSG_E_151, MsgBoxStyle.OkOnly)
                Return cell
            End If
        Next

        '明細番号数字3桁チェック
        For Each cell In dgvList.Rows(DGV_ROW.個人調査票項番).Cells.Cast(Of DataGridViewCell).Where(Function(x) IsKahenItemNo(x.GetString)).ToList
            Dim meisaiNo = cell.GetString.Split("_"c)(1)
            Dim i As Integer
            If Not Integer.TryParse(meisaiNo, i) Then
                i = 0
            End If
            If meisaiNo.Length <> 3 OrElse i < 1 OrElse i > 999 Then
                Message.ShowMsgBox(MessageID.MSG_E_152, MsgBoxStyle.OkOnly)
                Return cell
            End If
        Next

        '法人調査票項番
        '存在チェック
        Dim hojinItemNos = GetItemNoList(False, chosaNen)
        For Each i In dgvList.Rows(DGV_ROW.調査).Cells.Cast(Of DataGridViewCell).Where(Function(x) x.GetString = "1").Select(Function(x) x.ColumnIndex).ToList
            cell = dgvList.Rows(DGV_ROW.法人調査票項番).Cells(i)
            Dim itemNo = cell.GetString
            If String.IsNullOrEmpty(itemNo) Then
                Continue For
            End If
            If Not hojinItemNos.Contains(itemNo) Then
                Message.ShowMsgBox(MessageID.MSG_E_153, MsgBoxStyle.OkOnly)
                Return cell
            End If
        Next

        '米在庫調査票列
        '活性の場合、必須
        cell = dgvList.Rows(DGV_ROW.米在庫調査票列).Cells.Cast(Of DataGridViewCell).Where(Function(x) Not x.ReadOnly).Where(Function(x) String.IsNullOrEmpty(x.GetString)).FirstOrDefault
        If cell IsNot Nothing Then
            Message.ShowMsgBox(MessageID.MSG_E_154, MsgBoxStyle.OkOnly)
            Return cell
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' 調査票項番（テーブル列名）検索SQL
    ''' </summary>
    Private Const SQL_SELECT_ITEM_NO =
        "SELECT" _
        & " NAME" _
        & " FROM SYS.COLUMNS" _
        & " WHERE OBJECT_ID IN" _
        & " (SELECT OBJECT_ID" _
        & " FROM SYS.OBJECTS" _
        & " WHERE NAME = {0})" _
        & " UNION" _
        & " SELECT " _
        & " 項目番号" _
        & " FROM 調査票項目マスタ" _
        & " WHERE バージョン区分 = @バージョン区分" _
        & " AND 調査区分 = @調査区分" _
        & " AND 可変区分 = 1"
    ''' <summary>
    ''' 調査票項番（テーブル列名）リスト取得
    ''' </summary>
    ''' <param name="isKojin"></param>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    Private Shared Function GetItemNoList(isKojin As Boolean, chosaNen As Integer) As List(Of String)
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Dim sql = String.Format(SQL_SELECT_ITEM_NO, If(isKojin, "'調査票＿農業経営＿営農類型＿個人'", "'調査票＿農業経営＿営農類型＿法人'"))
            Dim para = New List(Of DBAccess.Parameter)
            Dim chosakubun = If(isKojin, ComConst.調査区分.営農類型別経営統計_個人, ComConst.調査区分.営農類型別経営統計_法人)
            para.Add(db.CreateParameter("@バージョン区分", SqlDbType.Int, ComUtil.getVersionKubun(chosaNen.ToString, chosakubun)))
            para.Add(db.CreateParameter("@調査区分", SqlDbType.Int, chosakubun))
            Return db.GetDataTable(sql, para).Rows.Cast(Of DataRow).Select(Function(x) x.GetString("NAME")).ToList
        End Using
    End Function

    ''' <summary>
    ''' 調査票項番が可変か否かを判定
    ''' </summary>
    ''' <param name="val"></param>
    ''' <returns></returns>
    Private Shared Function IsKahenItemNo(val As String) As Boolean
        If val Is Nothing Then
            Return False
        End If
        Return val.StartsWith("Q") AndAlso val.Contains("_")
    End Function

    '営農データ検索SQL
    Private Const SQL_SELECT_EINO = "SELECT * FROM {0} WHERE 調査年 = @調査年 ORDER BY センサス番号"
    ''' <summary>
    ''' 営農データ検索
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="chosaNen"></param>
    ''' <returns></returns>
    Private Shared Function SelectEino(tableName As String, chosaNen As Integer) As DataTable
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Dim sql = String.Format(SQL_SELECT_EINO, tableName)
            Dim para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosaNen))
            Return db.GetDataTable(sql, para)
        End Using
    End Function

    '営農個人可変データ検索SQL
    Private Const SQL_SELECT_EINO_KOJIN_KAHEN =
        "SELECT" _
        & " センサス番号," _
        & " 項目番号," _
        & " 明細番号," _
        & " ISNULL(値,'') AS 値" _
        & " FROM 調査票＿農業経営＿営農類型＿個人＿可変" _
        & " WHERE 調査年 = @調査年" _
        & " ORDER BY センサス番号,項目番号,明細番号 DESC"
    ''' <summary>
    ''' 営農個人可変データ検索
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="chosanen"></param>
    ''' <returns></returns>
    Private Shared Function SelectEinoKojinKahen(chosanen As Integer) As DataTable
        Using db = New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            Dim para = New List(Of DBAccess.Parameter)
            para.Add(db.CreateParameter("@調査年", SqlDbType.Int, chosanen))
            Return db.GetDataTable(SQL_SELECT_EINO_KOJIN_KAHEN, para)
        End Using
    End Function

    ''' <summary>
    ''' 列追加
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DgvList_ColumnAdded(sender As Object, e As DataGridViewColumnEventArgs) Handles dgvList.ColumnAdded
        Try
            e.Column.FillWeight = FILL_WEIGHT
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' セル描画
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DgvList_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles dgvList.CellPainting
        Try
            If e.RowIndex < 0 Then
                Return
            End If
            Dim cell = dgvList(e.ColumnIndex, e.RowIndex)
            If e.RowIndex <= 1 AndAlso e.ColumnIndex > 0 Then
                '１行目、２行目の２列目以降は中央寄せ
                cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If
            If cell.ReadOnly Then
                '読取専用セルを薄いグレーにする
                cell.Style.BackColor = Color.WhiteSmoke
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' セル編集開始
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DgvList_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvList.CellBeginEdit
        Try
            Dim cell = DirectCast(dgvList(e.ColumnIndex, e.RowIndex), DataGridViewTextBoxCell)
            '最大入力可能桁数を設定する
            Select Case e.RowIndex
                Case DGV_ROW.調査
                    '1桁
                    cell.MaxInputLength = 1
                    dgvList.ImeMode = ImeMode.Disable
                Case DGV_ROW.項目名称
                    '100桁
                    cell.MaxInputLength = 100
                    dgvList.ImeMode = ImeMode.NoControl
                Case DGV_ROW.個人調査票項番
                    Dim chosa = dgvList.Rows(DGV_ROW.調査).Cells(e.ColumnIndex).GetString
                    '固定値は100桁、項番は13桁(Qnnnnnnnn_mmm)
                    cell.MaxInputLength = If(chosa = "0", 100, 13)
                    dgvList.ImeMode = ImeMode.NoControl
                Case DGV_ROW.法人調査票項番
                    Dim chosa = dgvList.Rows(DGV_ROW.調査).Cells(e.ColumnIndex).GetString
                    '固定値は100桁、項番は9桁(Qnnnnnnnn)
                    cell.MaxInputLength = If(chosa = "0", 100, 9)
                    dgvList.ImeMode = ImeMode.NoControl
                Case DGV_ROW.米在庫調査票列
                    '2桁
                    cell.MaxInputLength = 2
                    dgvList.ImeMode = ImeMode.Disable
            End Select

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
            Dim cell = DirectCast(dgvList(e.ColumnIndex, e.RowIndex), DataGridViewTextBoxCell)
            If cell.ReadOnly Then
                '読取専用は何もしない
                Return
            End If
            'Nothingは空文字にする
            If cell.Value Is Nothing Then
                cell.Value = ""
            End If

            '半角カンマを全角に置換する
            cell.Value = cell.GetString.Replace(",", "，")

            '調査の場合、入力値によって活性／非活性を制御する
            If e.RowIndex = 1 Then
                ChosaControl(cell)
            End If

            '米在庫調査票列は、1～99以外はクリア
            If e.RowIndex = 5 Then
                Dim i As Integer
                If Not Integer.TryParse(cell.GetString, i) Then
                    i = 0
                End If
                If i < 1 OrElse i > 99 Then
                    cell.Value = ""
                End If
            End If

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 調査の入力値による制御
    ''' </summary>
    ''' <param name="cell"></param>
    Private Sub ChosaControl(cell As DataGridViewCell)
        Dim chosa = cell.GetString
        '調査が、0,1,2以外はクリア
        If chosa <> "0" AndAlso chosa <> "1" AndAlso chosa <> "2" Then
            cell.Value = ""
            chosa = ""
        End If
        '項目名称行
        With dgvList.Rows(DGV_ROW.項目名称).Cells(cell.ColumnIndex)
            .ReadOnly = (chosa = "")
            If .ReadOnly Then
                .Value = ""
            End If
            .Style.BackColor = If(.ReadOnly, Color.WhiteSmoke, Color.White)
        End With
        '個人調査票項番行
        With dgvList.Rows(DGV_ROW.個人調査票項番).Cells(cell.ColumnIndex)
            .ReadOnly = (chosa = "" OrElse chosa = "2")
            If .ReadOnly Then
                .Value = ""
            End If
            .Style.BackColor = If(.ReadOnly, Color.WhiteSmoke, Color.White)
        End With
        '法人調査票項番行
        With dgvList.Rows(DGV_ROW.法人調査票項番).Cells(cell.ColumnIndex)
            .ReadOnly = (chosa = "" OrElse chosa = "2")
            If .ReadOnly Then
                .Value = ""
            End If
            .Style.BackColor = If(.ReadOnly, Color.WhiteSmoke, Color.White)
        End With
        '米在庫調査票列行
        With dgvList.Rows(DGV_ROW.米在庫調査票列).Cells(cell.ColumnIndex)
            .ReadOnly = (chosa = "" OrElse chosa = "0" OrElse chosa = "1")
            If .ReadOnly Then
                .Value = ""
            End If
            .Style.BackColor = If(.ReadOnly, Color.WhiteSmoke, Color.White)
        End With
    End Sub

End Class

Namespace BRA110

    ''' <summary>
    ''' StringBuilder拡張
    ''' </summary>
    Public Module StringBuilderEx
        ''' <summary>
        ''' 値から改行を削除し、ダブルクォーテーションで囲って追加
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <param name="value"></param>
        <Extension>
        Public Sub AddCol(sb As StringBuilder, value As String)
            sb.Append("""")
            sb.Append(If(value, "").Replace(vbCr, "").Replace(vbLf, ""))
            sb.Append(""",")
        End Sub
        ''' <summary>
        ''' 空列を追加
        ''' </summary>
        ''' <param name="sb"></param>
        <Extension>
        Public Sub AddEmptyCol(sb As StringBuilder)
            sb.Append(""""",")
        End Sub
        ''' <summary>
        ''' 文字列化し、最後のカンマを除去
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <returns></returns>
        <Extension>
        Public Function ToCsvLine(sb As StringBuilder) As String
            Return sb.ToString.TrimEnd(","c)
        End Function

    End Module

End Namespace